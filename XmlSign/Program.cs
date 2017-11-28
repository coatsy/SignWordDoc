namespace DocSign
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Packaging;
    using System.Security;
    using System.Security.Cryptography;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Cryptography.Xml;
    using System.Text;
    using System.Xml;

    class Program
    {
        static readonly string RT_OfficeDocument =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        static readonly string OfficeObjectID = "idOfficeObject";
        static readonly string SignatureID = "idPackageSignature";
        static readonly string ManifestHashAlgorithm = "http://www.w3.org/2000/09/xmldsig#sha1";

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
            }
            else
            {
                X509Certificate2 certificate = GetCertificate();
                if (certificate != null)
                {
                    foreach (string path in args)
                    {
                        if (File.Exists(path))
                        {
                            SignFile(path, certificate);
                        }
                        else
                        {
                            Console.WriteLine("File not found: {0}", path);
                        }
                    }
                }
            }
        }

        static void PrintUsage()
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine("Office Open XML Digital Signing Utility");
            message.AppendLine("Copyright (C) Wouter van Vugt, all rights reserved");
            message.AppendLine();
            message.AppendLine("docsign.exe - ");
            message.AppendLine("\tUtility to digitally sign a series of Office Open XML documents");
            message.AppendLine("\tusing a user level X509 certificate.");
            message.AppendLine();
            message.AppendLine("docsign.exe <path> <path> ...");
            message.AppendLine();
            message.AppendLine("\t- Options -");
            message.AppendLine();
            message.AppendLine("<path> -");
            message.AppendLine("\ta path to an Office Open XML document");
            message.AppendLine();
            Console.WriteLine(message.ToString());
        }

        static X509Certificate2 GetCertificate()
        {
            X509Store certStore = new X509Store(StoreLocation.CurrentUser);
            certStore.Open(OpenFlags.ReadOnly);
            X509Certificate2Collection certs =
                X509Certificate2UI.SelectFromCollection(
                    certStore.Certificates,
                    "Select a certificate",
                    "Please select a certificate",
                    X509SelectionFlag.SingleSelection);
            return certs.Count > 0 ? certs[0] : null;
        }

        static void SignFile(string path, X509Certificate2 certificate)
        {
            try
            {
                using (Package package = Package.Open(path))
                {
                    SignPackage(package, certificate);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error signing {0}: {1}", path, ex.Message);
            }
        }

        static void SignPackage(Package package, X509Certificate2 certificate)
        {
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
    }
}