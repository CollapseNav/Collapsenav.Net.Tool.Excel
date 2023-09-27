namespace Collapsenav.Net.Tool.Excel;

/// <summary>
/// 配置excel读取时的范围
/// </summary>
public class SimpleRange
{
    public int Row { get; private set; }
    public int Col { get; private set; }
    public SimpleRange()
    {
        (Row, Col) = (0, 0);
    }
    /// <summary>
    /// 跳过行
    /// </summary>
    public void SkipRow(int row)
    {
        Row = row;
    }
    /// <summary>
    /// 跳过列
    /// </summary>
    public void SkipCol(int col)
    {
        Col = col;
    }
}