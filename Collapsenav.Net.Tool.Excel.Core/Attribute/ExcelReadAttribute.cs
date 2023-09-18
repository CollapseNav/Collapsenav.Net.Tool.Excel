namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 导入注解
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
public sealed class ExcelReadAttribute : ExcelAttribute
{
    public ExcelReadAttribute(string excelField) : base(excelField)
    {
    }
}