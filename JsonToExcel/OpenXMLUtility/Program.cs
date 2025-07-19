// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml.Packaging;
using OpenXMLUtility;

//OpenXMLUtility.OpenXmlOperations.ReplaceTextWithSAX(@"/Users/anandneema/Downloads/Anand Neema.docx","Professional","Anand");

string path = @"/Users/anandneema/Downloads/Anand Neema.docx";
WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, true);

MainDocumentPart mainPart = wordDoc.MainDocumentPart;
OpenXmlOperations.ReplaceTags(mainPart, "Professional", "Anand");

Console.WriteLine("Hello, World!");
