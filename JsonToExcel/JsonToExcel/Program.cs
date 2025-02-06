
using JsonToExcel;

string excelFilePath = @"..\..\JsonToExcel\JsonToExcel.xlsx";

try
{
    List<string> lstString = ExcelHelper.GetListFromExcel(excelFilePath);
    ExcelHelper.WriteToSpecificCell(excelFilePath, lstString);
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}

