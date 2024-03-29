namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 尝试使用 IExcelRead 统一 NPOI , EPPlus , MiniExcel 的调用
/// </summary>
public interface IExcelReader : IExcelReader<string> { }
/// <summary>
/// 尝试使用 IExcelRead 统一 NPOI , EPPlus , MiniExcel 的调用
/// </summary>
public interface IExcelReader<T> : IExcelContainer<T>, IExcelHeader
{
    /// <summary>
    /// 重新初始化Header
    /// </summary>
    void InitHeader(SimpleRange range);
}