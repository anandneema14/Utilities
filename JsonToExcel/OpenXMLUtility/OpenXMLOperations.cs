using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OpenXMLUtility;

public static class OpenXmlOperations
{
    public static void ReplaceTextWithSAX(string path, string textToReplace, string replacementText)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (OpenXmlReader reader = OpenXmlPartReader.Create(mainPart))
                using (OpenXmlWriter writer = OpenXmlPartWriter.Create(memoryStream))
                {
                    writer.WriteStartDocument();
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Text))
                        {
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);
                                string text = reader.GetText().Replace(textToReplace, replacementText);
                                writer.WriteString(text);
                            }
                            else
                            {
                                writer.WriteEndElement();
                            }
                        }
                        else
                        {
                            if (reader.IsStartElement)
                                writer.WriteStartElement(reader);
                            else if (reader.IsEndElement)
                                writer.WriteEndElement();
                        }
                    }
                }

                memoryStream.Position = 0;
                mainPart.FeedData(memoryStream);
                Console.WriteLine(mainPart.Document.Body.InnerText);
            }
        }
    }
    
    // Example method to replace a tag and retain formatting
    public static void ReplaceTags(MainDocumentPart mainPart, string tagName, string tagValue)
    {
        // Find all SdtBlock content controls with the given tagName
        IEnumerable<SdtBlock> tagFields = mainPart.Document.Body.Descendants<SdtBlock>()
            .Where(r => r.SdtProperties.GetFirstChild<Tag>().Val == tagName);

        foreach (var field in tagFields)
        {
            // Get paragraph properties (if any)
            ParagraphProperties paraProps = field.Descendants<ParagraphProperties>().FirstOrDefault();

            // Create a new paragraph
            var newParagraph = new Paragraph();
            if (paraProps != null)
            {
                // Clone and assign the paragraph properties to preserve formatting
                newParagraph.AppendChild(paraProps.CloneNode(true));
            }

            // Get run properties (style such as font, color)
            RunProperties runProp = field.SdtProperties.GetFirstChild<RunProperties>();

            // Create a run and assign run properties
            var newRun = new Run();
            if (runProp != null)
            {
                newRun.Append(runProp.CloneNode(true));
            }
            
            // Set the text
            var newText = new Text(tagValue);
            newRun.Append(newText);
            newParagraph.Append(newRun);

            // Insert the new paragraph just before the field and remove the original field
            field.Parent.InsertBefore(newParagraph, field);
            field.Remove();
        }
        Console.WriteLine(mainPart.Document.Body.InnerText);
    }
}