# Sign Word Doc
Sample code for generating a word document from a template and a data source, then digitally signing it.

## Content
[SignWordDoc](https://github.com/coatsy/SignWordDoc/tree/master/SignWordDoc) is the working code

[XmlSign](https://github.com/coatsy/SignWordDoc/tree/master/XmlSign) is a copy of code Wouter van Vugt posted 10 years or so ago and now seems to be unavailable (thanks way-back-machine)

## Instructions

### Note
To run this sample, you'll need an X509 certificate. I've included one in the project that I generated (using the [batch file in this folder](https://github.com/coatsy/SignWordDoc/blob/master/CreateCoatsyDocSign.cmd)). You may want to generate or even obtain your own. There's a good explanation of [using the MakeCert tool](https://blog.jayway.com/2014/09/03/creating-self-signed-certificates-with-makecert-exe-for-development/#comments) from [Elizabeth Andrews](https://blog.jayway.com/author/elizabethandrews/) which I used to craft the cmd file.

## Architecture
The idea of the sample is to take a data source (in this case, mocked) and generate a merged document from a template document, then sign it using a certificate which becomes invalid if the document is altered in any way. The Open Packaging Convention (of which the Word docx format is an instance) allows for signing all or parts of a document. In this case we're signing all of the document.

Here's the general process the sample uses:

![Architecture Diagram](/images/Architecture.png)

* After some basic parameter checking, the template document is merged with the data for this instance to create the document.
* The Contents of the document are collected and a hash calculated with using the X509 certificate. This hash (and some other info) is stored in the document - this is the digital signature.
* The signed file is written out to disk.

## More Details

### Merging the data with the template document
There are a number of approaches I've used in the past to do programmatic merging of data with an existing template. This time I've gone with one based on Content Controls.

The template document has the basic structure of the output required and wherever data needs to be inserted, I've added a content control with a well known name.

The procedure for adding a Content Control is quite straight-forward:

1. Ensure the [Developer Tab is enabled](https://support.office.com/en-us/article/Show-the-Developer-tab-E1192344-5E56-4D45-931B-E5FD9BEA2D45) in Word

2. With the cursor positiond where you want the data to be inserted, click the Plain Text Content Control button on the Developer Tab:

![Insert Content Control](/images/InsertContentControl1.png)

3. This will insert a content control at the cursor. With the content control selected, click the Properties button:

![Open Content Control Properties](/images/InsertContentControl2.png)

4. This will pop up the properties dialog for the control. Give the control a unique Tag - we'll be using that string to find the control programmatically:

![Content Control Properties Dialog](/images/InsertContentControl3.png)

The code to insert the data programmatically is also pretty easy. Using a Linq query, we can get a reference to the content control:
```csharp
var control = doc.MainDocumentPart.Document.Descendants<SdtRun>().Where(r => r.SdtProperties.GetFirstChild<Tag>().Val.Value == contentContolName).FirstOrDefault();
```
Next we find the parent of that control and remove the control, adding a `Run` of our own containing the `Text` we're adding:
```csharp
var parent = control.Parent;
control.Remove();

// now add the text as a child of the parent (replacing the content control)
parent.AppendChild<Run>(new Run(new Text(newText)));
```
Doing this with whole paragraphs is a bit more involved, because there's not an easy way to add a bunch of elements in the right place in the document. The approach I took was to find the content control, and add the element before the control, then remove the control:
```csharp
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
    }
}
```

### Signing the merged document
To sign the document, you need a certificate and you need to use the Open Packaging Convention's facility for storing a hash based on the certificate. Wouter van Vugt did a great post about this 10 years ago and it still works great. I've linked to all the details in the [XmlSign/readme.md](/XmlSign/readme.md).

I've also included a certificate I generated and used for testing, as well as a batch file you can modify to generate your own. There's a good explanation of [using the MakeCert tool](https://blog.jayway.com/2014/09/03/creating-self-signed-certificates-with-makecert-exe-for-development/#comments) from [Elizabeth Andrews](https://blog.jayway.com/author/elizabethandrews/) which I used to craft the cmd file.

### Untrusted Certificates are Untrusted
Of course, if you generate your own certificate, Word has no way to validate it. When you open the signed document, you'll see it's signed, but that the cert isn't validated:
![The certificate is untrusted and couldn't be validated](/images/UntrustedCert1.png)
![The certificate is untrusted and couldn't be validated](/images/UntrustedCert2.png)
![The certificate is untrusted and couldn't be validated](/images/UntrustedCert3.png)

If you Click the link to trust the user's identity, then all will be well:

![The certificate is trusted and could be validated](/images/TrustedCert1.png)
![The certificate is trusted and could be validated](/images/TrustedCert2.png)
![The certificate is trusted and could be validated](/images/TrustedCert3.png)
