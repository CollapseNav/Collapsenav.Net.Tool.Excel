namespace Collapsenav.Net.Tool.Excel;

/// <summary>
/// 单元格
/// </summary>
public interface IReadCell
{
    /// <summary>
    /// 行数
    /// </summary>
    int Row { get; }
    /// <summary>
    /// 列数
    /// </summary>
    int Col { get; }
    /// <summary>
    /// 纯字符串格式的value
    /// </summary>
    string StringValue { get; }
    /// <summary>
    /// 值类型
    /// </summary>
    Type? ValueType { get; }
    /// <summary>
    /// 单元格的值
    /// </summary>
    object? Value { get; set; }
    /// <summary>
    /// 从其他单元格读取值
    /// </summary>
    void CopyCellFrom(IReadCell cell);
}
/// <summary>
/// 单元格
/// </summary>
public interface IReadCell<T> : IReadCell
{
    /// <summary>
    /// 适配不同实现的单元格
    /// </summary>
    T? Cell { get; set; }
}
