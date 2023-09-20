namespace Collapsenav.Net.Tool.Excel;
public static class ExcelExt
{
    /// <summary>
    /// 是否 xls 文件
    /// </summary>
    public static bool IsXls(this string filepath)
    {
        return ExcelTool.IsXls(filepath);
    }
    /// <summary>
    /// 是否 xlsx 文件
    /// </summary>
    public static bool IsXlsx(this string filepath)
    {
        return ExcelTool.IsXlsx(filepath);
    }
}
