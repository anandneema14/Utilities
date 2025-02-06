using System.Text.Json;

namespace JsonToExcel;

public class JsonHelper
{
    public static string GetJsonContentByName(string filePath, string propertyName, string propertyValue)
    {
        string jsonString = File.ReadAllText(filePath);
        using (JsonDocument document = JsonDocument.Parse(jsonString))
        {
            foreach (var element in document.RootElement.EnumerateArray())
            {
                try
                {
                    if (element.TryGetProperty(propertyName, out JsonElement propertyElement) &&
                        propertyElement.GetString() == propertyValue)
                    {
                        return element.GetRawText();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            throw new Exception($"Property {propertyName} not found in file {filePath}");
        }
    }
}