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
}