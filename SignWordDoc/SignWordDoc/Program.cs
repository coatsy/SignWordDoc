using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SignWordDoc.Services;
using SignWordDoc.Data;
using System.Security.Cryptography.X509Certificates;
using System.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.IO.Packaging;
using System.Xml;

namespace SignWordDoc
{
    class Program
    {

        // this is the name of the content controls in the template document
        const string CONTENT_CONTROL_NAME_POLICY_NUMBER = "contentPolicyNumber";
        const string CONTENT_CONTROL_NAME_INSURED = "contentInsured";
        const string CONTENT_CONTROL_NAME_START_DATE = "contentStartDate";
        const string CONTENT_CONTROL_NAME_END_DATE = "contentEndDate";
        const string CONTENT_CONTROL_NAME_SUM_INSURED = "contentSumInsured";
        const string CONTENT_CONTROL_NAME_PREMIUM = "contentPremium";
        const string CONTENT_CONTROL_NAME_SPECIAL_CONDITIONS = "contentSpecialConditions";
        const string CONTENT_CONTROL_NAME_EXCLUSIONS = "contentExclusions";

        static readonly string RT_OfficeDocument =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        static readonly string OfficeObjectID = "idOfficeObject";
        static readonly string SignatureID = "idPackageSignature";
        static readonly string ManifestHashAlgorithm = "http://www.w3.org/2000/09/xmldsig#sha1";

        static void Main(string[] args)
        {
            // check that five arguments were passed
            if (args.Length != 5)
            {
                Console.WriteLine("Usage: SignWordDoc.exe <templateFileName> <policyId> <outputFileName> <certfile> <certPassword>");
                Console.ReadKey();
                return;
            }

            // check that the first argument is a file that exists
            if (! File.Exists(args[0]))
            {
                Console.WriteLine($"TemplateFileName \"{args[0]}\" is invalid or does not exist");
                Console.ReadKey();
                return;
            }

            // check that the fourth argument is a file that exists
            if (!File.Exists(args[3]))
            {
                Console.WriteLine($"CertificateFileName \"{args[3]}\" is invalid or does not exist");
                Console.ReadKey();
                return;
            }

            // create an instance of the DataService
            // in a production system, you'd use some kind of DI pattern
            // but we'll just create the mock one here
            IDataService dataService = new MockDataService();

            // check that the customerId is valid
            if (! dataService.IsValidPolicyId(args[1]))
            {
                Console.WriteLine($"PolicyId \"{args[1]}\" is invalid or does not exist");
                Console.ReadKey();
                return;
            }

            // check that the file is a valid word document and has the correct placeholders
            try
            {
                using (WordprocessingDocument templateDocument = WordprocessingDocument.Open(args[0], false))
                {

                    // check that the template document has placeholders for the appropriate information
                    if (!TemplateHasPlaceholders(templateDocument))
                    {
                        Console.WriteLine($"TemplateFile \"{args[0]}\" does not contain the appropriate placeholders for merging");
                        Console.ReadKey();
                        return;
                    }
                }
            }
            catch (OpenXmlPackageException)
            {
                Console.WriteLine($"TemplateFile \"{args[0]}\" is not a valid OpenXML Document");
                Console.ReadKey();
                return;
            }

            // Replace the placeholders with the appropriate data
            var outputStream = (MemoryStream)InsertPolicyData(args[0], dataService.GetPolicy(args[1]));

            // now sign the stream
            SignPackage(outputStream, args[3], args[4]);

            // write the resulting file top the file system
            using (FileStream fs = new FileStream(args[2], FileMode.Create))
            {
                outputStream.Seek(0, SeekOrigin.Begin);
                outputStream.CopyTo(fs);
            }
        }

        private static void SignPackage(MemoryStream outputStream, string certPath, string certPassword)
        {

            // this from Richard diZerega
            // https://github.com/richdizz/microsoft-graph-app-only/blob/master/RichdizzReady/Program.cs
            var certfile = System.IO.File.OpenRead(certPath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);

            var certificate = new X509Certificate2(
                certificateBytes,
                certPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet); //switches are important to work in webjob

            var package = WordprocessingDocument.Open(outputStream, true).Package;

            // This from Wouter's sample code from 10 years ago
            List<Uri> partsToSign = new List<Uri>();
            List<PackageRelationshipSelector> relationshipsToSign =
                new List<PackageRelationshipSelector>();
            List<Uri> finishedItems = new List<Uri>();
            foreach (PackageRelationship relationship in
                package.GetRelationshipsByType(RT_OfficeDocument))
            {
                AddSignableItems(relationship,
                    partsToSign, relationshipsToSign);
            }
            PackageDigitalSignatureManager mgr =
                new PackageDigitalSignatureManager(package);
            mgr.CertificateOption = CertificateEmbeddingOption.InSignaturePart;

            string signatureID = SignatureID;
            string manifestHashAlgorithm = ManifestHashAlgorithm;
            DataObject officeObject = CreateOfficeObject(signatureID, manifestHashAlgorithm);
            Reference officeObjectReference = new Reference("#" + OfficeObjectID);
            mgr.Sign(partsToSign, certificate,
                relationshipsToSign, signatureID,
                new DataObject[] { officeObject },
                new Reference[] { officeObjectReference });



        }

        static void AddSignableItems(
    PackageRelationship relationship,
    List<Uri> partsToSign,
    List<PackageRelationshipSelector> relationshipsToSign)
        {
            PackageRelationshipSelector selector =
                new PackageRelationshipSelector(
                    relationship.SourceUri,
                    PackageRelationshipSelectorType.Id,
                    relationship.Id);
            relationshipsToSign.Add(selector);
            if (relationship.TargetMode == TargetMode.Internal)
            {
                PackagePart part = relationship.Package.GetPart(
                    PackUriHelper.ResolvePartUri(
                        relationship.SourceUri, relationship.TargetUri));
                if (partsToSign.Contains(part.Uri) == false)
                {
                    partsToSign.Add(part.Uri);
                    foreach (PackageRelationship childRelationship in
                        part.GetRelationships())
                    {
                        AddSignableItems(childRelationship,
                            partsToSign, relationshipsToSign);
                    }
                }
            }
        }

        static DataObject CreateOfficeObject(
    string signatureID, string manifestHashAlgorithm)
        {
            XmlDocument document = new XmlDocument();
            document.LoadXml(String.Format(Properties.Resources.OfficeObject,
                signatureID, manifestHashAlgorithm));
            DataObject officeObject = new DataObject();
            // do not change the order of the following two lines
            officeObject.LoadXml(document.DocumentElement); // resets ID
            officeObject.Id = OfficeObjectID; // required ID, do not change
            return officeObject;
        }


        private static bool TemplateHasPlaceholders(WordprocessingDocument templateDocument)
        {
            // for now, just return true
            return true;
        }

        private static Stream InsertPolicyData(string templateDocument, Policy policy)
        {
            MemoryStream ms = new MemoryStream();
            GetDocStreamFromTemplate(ms, templateDocument);
            using (WordprocessingDocument output = GetDocFromTemplate(ms))
            {
                // Policy Number
                ReplaceContentControl(output, CONTENT_CONTROL_NAME_POLICY_NUMBER, policy.PolicyId);
                // Start Date
                ReplaceContentControl(output, CONTENT_CONTROL_NAME_START_DATE, policy.StartDate.ToLongDateString());
                // Start Date
                ReplaceContentControl(output, CONTENT_CONTROL_NAME_END_DATE, policy.EndDate.ToLongDateString());
                // Details of Insured
                ReplaceContentControlOpenXML(output, CONTENT_CONTROL_NAME_INSURED, GenerateInsuredDetails(policy.Insured));
                // Sum Insured
                ReplaceContentControl(output, CONTENT_CONTROL_NAME_SUM_INSURED, policy.SumInsured.ToString("c"));
                // Premium
                ReplaceContentControl(output, CONTENT_CONTROL_NAME_PREMIUM, policy.Premium.ToString("c"));
                // Special Conditions
                ReplaceContentControlOpenXML(output, CONTENT_CONTROL_NAME_SPECIAL_CONDITIONS, new OpenXmlElement[] { new Paragraph(new Run(new Text(policy.SpecialConditions))) });
                // Exclusions
                ReplaceContentControlOpenXML(output, CONTENT_CONTROL_NAME_EXCLUSIONS, new OpenXmlElement[] { new Paragraph(new Run(new Text(policy.Exclusions))) });
            }
            return ms;
        }

        private static OpenXmlElement[] GenerateInsuredDetails(List<Customer> insured)
        {
            List<OpenXmlElement> paras = new List<OpenXmlElement>();
            int custno = 0;
            foreach (var customer in insured)
            {
                custno++;
                // Customer Number Heading
                var paraHead = new Paragraph() ;
                var paraHeadProps = new ParagraphProperties();
                var paraHeadStyleId = new ParagraphStyleId() { Val = "Heading2" };
                paraHeadProps.Append(paraHeadStyleId);
                paraHead.Append(paraHeadProps);
                paraHead.Append(new Run(new Text($"Customer {custno}")));
                paras.Add(paraHead);

                // Sub Headings for each property

                // Full Name Heading
                var paraNameHead = new Paragraph();

                var paraNameHeadProps = new ParagraphProperties();
                var paraNameHeadStyleId = new ParagraphStyleId() { Val = "Heading3" };

                paraNameHeadProps.Append(paraNameHeadStyleId);

                paraNameHead.Append(paraNameHeadProps);
                paraNameHead.Append(new Run(new Text("Full Legal Name")));
                paras.Add(paraNameHead);

                // Full Name text
                paras.Add(new Paragraph(new Run(new Text(customer.LegalName))));

                // Date Of Birth Heading
                var paraDoBHead = new Paragraph();

                var paraDoBHeadProps = new ParagraphProperties();
                var paraDoBHeadStyleId = new ParagraphStyleId() { Val = "Heading3" };

                paraDoBHeadProps.Append(paraDoBHeadStyleId);

                paraDoBHead.Append(paraDoBHeadProps);
                paraDoBHead.Append(new Run(new Text("Date of Birth")));
                paras.Add(paraDoBHead);

                // DoB text
                paras.Add(new Paragraph(new Run(new Text(customer.DoB.ToLongDateString()))));

                // Address Heading
                var paraAddressHead = new Paragraph();

                var paraAddressHeadProps = new ParagraphProperties();
                var paraAddressHeadStyleId = new ParagraphStyleId() { Val = "Heading3" };

                paraAddressHeadProps.Append(paraAddressHeadStyleId);

                paraAddressHead.Append(paraAddressHeadProps);
                paraAddressHead.Append(new Run(new Text("Address")));

                paras.Add(paraAddressHead);

                // Address text
                var paraAddress = new Paragraph();
                int addressLine = 0;
                foreach (var line in customer.Address.Split('\n'))
                {
                    if (addressLine > 0)
                    {
                        paraAddress.Append(new Break());
                    }
                    paraAddress.Append(new Run(new Text(line)));
                    addressLine++;
                }

                paras.Add(paraAddress);

                // Pre-existing Conditions Heading
                var paraConditionsHead = new Paragraph();

                var paraConditionsHeadProps = new ParagraphProperties();
                var paraConditionsHeadStyleId = new ParagraphStyleId() { Val = "Heading3" };

                paraConditionsHeadProps.Append(paraConditionsHeadStyleId);

                paraConditionsHead.Append(paraConditionsHeadProps);
                paraConditionsHead.Append(new Run(new Text("Pre-Existing Conditions")));
                paras.Add(paraConditionsHead);

                // Pre-existing Conditions

                var paraConditions = new Paragraph();
                int conditionsLine = 0;
                foreach (var line in customer.PreExistingConditions.Split('\n'))
                {
                    if (conditionsLine > 0)
                    {
                        paraConditions.Append(new Break());
                    }
                    paraConditions.Append(new Run(new Text(line)));
                    conditionsLine++;
                }
                paras.Add(paraConditions);
            }


            return paras.ToArray();
        }

        private static void GetDocStreamFromTemplate(MemoryStream ms, string templateDocument)
        {
            using (var source = new FileStream(templateDocument, FileMode.Open, FileAccess.Read))
            {
                source.Seek(0, SeekOrigin.Begin);
                source.CopyTo(ms);
            }
        }

        private static WordprocessingDocument GetDocFromTemplate(Stream stream)
        {
            return WordprocessingDocument.Open(stream, true);
        }



        /// <summary>
        /// Finds a content control and replaces it with text in-line (using a Run(Text()))
        /// </summary>
        /// <param name="doc">Document to Search</param>
        /// <param name="contentContolName">Name of the content control to replace</param>
        /// <param name="newText">Text to insert</param>
        private static void ReplaceContentControl(WordprocessingDocument doc, string contentContolName, string newText)
        {
            var control = doc.MainDocumentPart.Document.Descendants<SdtRun>().Where(r => r.SdtProperties.GetFirstChild<Tag>().Val.Value == contentContolName).FirstOrDefault();
            if (control != null)
            {
                // if we've found the content control, get a refernce to its parent, then remove the control
                var parent = control.Parent;
                control.Remove();

                // now add the text as a child of the parent (replacing the content control)
                parent.AppendChild<Run>(new Run(new Text(newText)));

            }

        }

        /// <summary>
        /// Finds a content control and replaces it with the OpenXML Part passed in
        /// Can be used for Paragraphs, Tables, etc
        /// </summary>
        /// <param name="doc">Document to Search</param>
        /// <param name="contentControlName">Name of the content control to replace</param>
        /// <param name="content">Array of OpenXML elements to insert</param>
        private static void ReplaceContentControlOpenXML(WordprocessingDocument doc, string contentControlName, OpenXmlElement[] content)
        {
            var control = doc.MainDocumentPart.Document.Descendants<SdtBlock>().Where(b => b.SdtProperties.GetFirstChild<Tag>().Val.Value == contentControlName).FirstOrDefault();
            if (control != null)
            {
                foreach (var element in content)
                {
                    control.InsertBeforeSelf(element);
                }
                control.Remove();
                //var parent = control.Parent;
                //var newContent = content.Length == 1 ? content[0] : new Paragraph(content);
                //parent.ReplaceChild<SdtBlock>(newContent, control);
            }
        }
    }

}
