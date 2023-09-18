namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 导出注解
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
public sealed class ExcelExportAttribute : ExcelAttribute
{
    public ExcelExportAttribute(string excelField) : base(excelField)
    {
    }
}