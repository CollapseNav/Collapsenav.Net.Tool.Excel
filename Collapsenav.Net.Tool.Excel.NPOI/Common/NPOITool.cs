using System.Data;
using NPOI.SS.UserModel;

namespace Collapsenav.Net.Tool.Excel;

public partial class NPOITool
{
    /// <summary>
    /// 获取 NPOI中使用 的 Workbook
    /// </summary>
    public static IWorkbook NPOIWorkbook(string path)
    {
        using var notCloseStream = new NPOINotCloseStream(path);
        return notCloseStream.GetWorkBook();
    }
    /// <summary>
    /// 获取 NPOI中使用 的 Workbook
    /// </summary>
    public static IWorkbook NPOIWorkbook(Stream stream)
    {
        using var notCloseStream = stream is NPOINotCloseStream nstream ? nstream : new NPOINotCloseStream(stream);
        return notCloseStream.GetWorkBook();
    }
    /// <summary>
    /// 获取 NPOI中使用 的 Sheet
    /// </summary>
    public static ISheet NPOISheet(string path, string? sheetname = null)
    {
        var workbook = NPOIWorkbook(path);
        return sheetname.IsNull() ? workbook.GetSheetAt(ExcelTool.NPOIZero) : workbook.GetSheet(sheetname);
    }
    /// <summary>
    /// 获取 NPOI中使用 的 Sheet
    /// </summary>
    public static ISheet NPOISheet(Stream stream, string? sheetname = null)
    {
        var workbook = NPOIWorkbook(stream);
        if (workbook.NumberOfSheets == 0)
            return workbook.CreateSheet("sheet1");
        return sheetname.IsNull() ? workbook.GetSheetAt(ExcelTool.NPOIZero) : workbook.GetSheet(sheetname);
    }
    /// <summary>
    /// 获取 NPOI中使用 的 Sheet
    /// </summary>
    public static ISheet NPOISheet(string path, int sheetindex)
    {
        var workbook = NPOIWorkbook(path);
        return workbook.GetSheetAt(sheetindex);
    }
    /// <summary>
    /// 获取 NPOI中使用 的 Sheet
    /// </summary>
    public static ISheet NPOISheet(Stream stream, int sheetindex)
    {
        var workbook = NPOIWorkbook(stream);
        return workbook.GetSheetAt(sheetindex);
    }

    /// <summary>
    /// 获取表格header(仅限简单的单行表头)
    /// </summary>
    /// <param name="sheet">工作簿</param>
    /// <param name="range"></param>
    public static IEnumerable<string> ExcelHeader(ISheet sheet, SimpleRange? range = null)
    {
        IRow row;
        if (range == null)
            row = sheet.GetRow(ExcelTool.NPOIZero);
        else
            row = sheet.GetRow(ExcelTool.NPOIZero + range.Row);
        var header = row?.Cells?.Select(item => item.ToString()?.Trim() ?? string.Empty);
        return header ?? Enumerable.Empty<string>();
    }
    /// <summary>
    /// 获取表格header和对应的index
    /// </summary>
    public static IDictionary<string, int> HeadersWithIndex(ISheet sheet, SimpleRange? range = null)
    {
        IRow row;
        if (range == null)
            row = sheet.GetRow(ExcelTool.NPOIZero);
        else
            row = sheet.GetRow(ExcelTool.NPOIZero + range.Row);
        var headers = row?.Cells?
        .Where(item => item.ToString().NotNull())?
        .ToDictionary(item => item.ToString()?.Trim() ?? DateTime.Now.ToTimestamp().ToString(), item => item.ColumnIndex);
        return headers ?? new Dictionary<string, int>();
    }
}