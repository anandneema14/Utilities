using Aspose.Cells;

namespace JsonToExcel;

public class ExcelHelper
{
    private const string propertyName = "Name";
    private const string jsonFilePath = "JsonToExcel.json";

    public static void ReadExcelFile(string filePath)
    {
        using (var workbook = new Workbook(filePath))
        {
            var worksheet = workbook.Worksheets[0]; //Read the first sheet
            var cells = worksheet.Cells;

            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {
                    Console.Write($"{cells[i, j].Value}\t");
                }
                Console.WriteLine();
            }
        }
    }

    public static void WriteExcelFile(string filePath, List<List<string>> data)
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        for (int i = 0; i < data.Count; i++)
        {
            for (int j = 0; j < data[i].Count; j++)
            {
                worksheet.Cells[i, j].Value = data[i][j];
            }
        }
        workbook.Save(filePath);
    }

    public static void WriteToSpecificCell(string filePath, List<string> data)
    {
        int rowIndex = 1;
        int columnIndex = 3;
        var workbook = new Workbook(filePath);
        var worksheet = workbook.Worksheets[0];

        foreach (var dataItem in data)
        {
            string jsonContent = JsonHelper.GetJsonContentByName(jsonFilePath, propertyName, dataItem);
            worksheet.Cells[rowIndex, columnIndex].Value = jsonContent;
            rowIndex++;
        }
        workbook.Save(filePath);
    }

    public static List<string> GetListFromExcel(string filePath)
    {
        var workbook = new Workbook(filePath);
        var worksheet = workbook.Worksheets[0];
        var cells = worksheet.Cells;
        List<string> data = new List<string>();

        for (int i = 0; i < cells.MaxDataRow; i++)  //skip the data row
        {
            data.Add(cells[i, 2].Value.ToString());
        }
        return data;
    }
}