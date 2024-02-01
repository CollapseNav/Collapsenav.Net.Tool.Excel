using OfficeOpenXml;

namespace Collapsenav.Net.Tool.Excel;

public static class EPPlusExt
{
    public static IExcelReader GetExcelReader(this ExcelWorksheet sheet)
    {
        return ExcelTool.GetExcelReader(sheet, ExcelType.EPPlus);
    }
    public static IExcelCellReader GetCellReader(this ExcelWorksheet sheet)
    {
        return ExcelTool.GetCellReader(sheet, ExcelType.EPPlus);
    }
}