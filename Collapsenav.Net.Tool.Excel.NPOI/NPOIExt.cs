using NPOI.SS.UserModel;

namespace Collapsenav.Net.Tool.Excel;
public static class NPOIExt
{
    public static IExcelReader GetExcelReader(this ISheet sheet)
    {
        return ExcelTool.GetExcelReader(sheet, ExcelType.NPOI);
    }
    public static IExcelCellReader GetCellReader(this ISheet sheet)
    {
        return ExcelTool.GetCellReader(sheet, ExcelType.NPOI);
    }
}