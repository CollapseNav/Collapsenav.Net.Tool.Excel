namespace Collapsenav.Net.Tool.Excel;
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = true)]
public class ExcelAttribute : Attribute
{
    /// <summary>
    /// 表头项
    /// </summary>
    readonly string excelField;
    public ExcelAttribute()
    {
        excelField = string.Empty;
    }
    public ExcelAttribute(string excelField) : this()
    {
        this.excelField = excelField;
    }
    public string ExcelField { get => excelField; }
}