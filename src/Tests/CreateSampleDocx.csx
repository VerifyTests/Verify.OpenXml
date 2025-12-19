#r "nuget: DocumentFormat.OpenXml, 3.0.0"
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using (var doc = WordprocessingDocument.Create("sample.docx", WordprocessingDocumentType.Document))
{
    doc.PackageProperties.Title = "Sample Document";
    doc.PackageProperties.Creator = "Test Author";
    doc.PackageProperties.Subject = "Test Subject";
    
    var mainPart = doc.AddMainDocumentPart();
    mainPart.Document = new Document(
        new Body(
            new Paragraph(
                new Run(
                    new Text("Hello World! This is a sample Word document.")
                )
            ),
            new Paragraph(
                new Run(
                    new Text("This is the second paragraph with some more text.")
                )
            )
        )
    );
    mainPart.Document.Save();
}
Console.WriteLine("Created sample.docx");
