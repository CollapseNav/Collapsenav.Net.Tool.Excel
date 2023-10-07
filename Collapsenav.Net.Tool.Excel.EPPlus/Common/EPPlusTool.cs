using System.Data;
using OfficeOpenXml;

namespace Collapsenav.Net.Tool.Excel;

public partial class EPPlusTool
{
    /// <summary>
    /// 获取 EPPlus中使用 的 ExcelPackage
    /// </summary>
    public static ExcelPackage EPPlusPackage(Stream stream)
    {
        return new(stream);
    }
    public static ExcelPackage EPPlusPackage(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"{path} not exist");
        using var fs = path.OpenReadShareStream();
        try
        {
            var pack = new ExcelPackage(fs);
            return pack;
        }
        catch (Exception)
        {
            throw;
        }
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Workbook
    /// </summary>
    public static ExcelWorkbook EPPlusWorkbook(string path)
    {
        return EPPlusPackage(path).Workbook;
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Workbook
    /// </summary>
    public static ExcelWorkbook EPPlusWorkbook(Stream stream)
    {
        return EPPlusPackage(stream).Workbook;
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheets
    /// </summary>
    public static ExcelWorksheets EPPlusSheets(string path)
    {
        var sheets = EPPlusWorkbook(path).Worksheets;
        return sheets;
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheets
    /// </summary>
    public static ExcelWorksheets EPPlusSheets(Stream stream)
    {
        var sheets = EPPlusWorkbook(stream).Worksheets;
        return sheets;
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheet
    /// </summary>
    public static ExcelWorksheet EPPlusSheet(string path, string? sheetname = null)
    {
        var sheets = EPPlusSheets(path);
        if (sheets.Count == 0)
            return sheets.Add("sheet1");
        return sheetname.IsNull() ? sheets[ExcelTool.EPPlusZero] : sheets[sheetname];
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheet
    /// </summary>
    public static ExcelWorksheet EPPlusSheet(Stream stream, string? sheetname = null)
    {
        var sheets = EPPlusSheets(stream);
        if (sheets.Count == 0)
            return sheets.Add("sheet1");
        return sheetname.IsNull() ? sheets[ExcelTool.EPPlusZero] : sheets[sheetname];
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheet
    /// </summary>
    public static ExcelWorksheet EPPlusSheet(string path, int sheetindex)
    {
        var sheets = EPPlusSheets(path);
        return sheets[sheetindex];
    }
    /// <summary>
    /// 获取 EPPlus中使用 的 Sheet
    /// </summary>
    public static ExcelWorksheet EPPlusSheet(Stream stream, int sheetindex)
    {
        var sheets = EPPlusSheets(stream);
        return sheets[sheetindex];
    }
    /// <summary>
    /// 将表格数据转换为T类型的集合
    /// </summary>
    public static async Task<IEnumerable<T>> ExcelToEntityAsync<T>(IExcelReader reader, ReadConfig<T>? config = null)
    {
        config ??= ReadConfig<T>.GenDefaultConfig();
        return await config.ToEntityAsync(reader);
    }
    /// <summary>
    /// 获取表格header(仅限简单的单行表头)
    /// </summary>
    /// <param name="sheet">工作簿</param>
    /// <param name="range"></param>
    public static IEnumerable<string> ExcelHeader(ExcelWorksheet sheet, SimpleRange? range = null)
    {
        IEnumerable<ExcelRangeBase>? data = null;
        if (range == null)
            data = sheet.Cells[ExcelTool.EPPlusZero, ExcelTool.EPPlusZero, ExcelTool.EPPlusZero, sheet.Dimension.Columns];
        else
        {
            data = GetExcelRangeBaseByRange(sheet, range);
        }
        return data.Select(item => item.Value?.ToString()?.Trim() ?? string.Empty).ToList() ?? Enumerable.Empty<string>();
    }
    /// <summary>
    /// 获取表格header和对应的index
    /// </summary>
    public static IDictionary<string, int> HeadersWithIndex(ExcelWorksheet sheet, SimpleRange? range = null)
    {
        IEnumerable<ExcelRangeBase>? data = null;
        if (range == null)
            data = sheet.Cells[ExcelTool.EPPlusZero, ExcelTool.EPPlusZero, ExcelTool.EPPlusZero, sheet.Dimension.Columns];
        else
        {
            if (range.SelectRow == null)
                data = sheet.Cells[ExcelTool.EPPlusZero + range.Row, ExcelTool.EPPlusZero + range.Col, ExcelTool.EPPlusZero + range.Row, sheet.Dimension.Columns + range.Col];
            else
            {
                data = GetExcelRangeBaseByRange(sheet, range);
            }
        }
        var headers = data.ToDictionary(item => item.Value?.ToString()?.Trim() ?? DateTime.Now.ToTimestamp().ToString(), item => item.End.Column - ExcelTool.EPPlusZero);
        return headers;
    }
    private static IEnumerable<ExcelRangeBase> GetExcelRangeBaseByRange(ExcelWorksheet sheet, SimpleRange range)
    {
        if (range.SelectRow == null)
            return sheet.Cells[ExcelTool.EPPlusZero + range.Row, ExcelTool.EPPlusZero + range.Col, ExcelTool.EPPlusZero + range.Row, sheet.Dimension.Columns + range.Col];
        else
        {
            for (var i = ExcelTool.EPPlusZero; i < sheet.Dimension.Columns; i++)
            {
                var data = sheet.Cells[i, ExcelTool.EPPlusZero + range.Col, i, sheet.Dimension.Columns + range.Col];
                if (data.NotEmpty() && range.SelectRow(data.Select(item => item.Value.ToString())!))
                {
                    range.SkipRow(i - ExcelTool.EPPlusZero);
                    return data;
                }
            }
            return Enumerable.Empty<ExcelRangeBase>();
        }
    }
}