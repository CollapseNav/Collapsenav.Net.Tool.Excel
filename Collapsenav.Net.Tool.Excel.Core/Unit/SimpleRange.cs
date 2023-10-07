namespace Collapsenav.Net.Tool.Excel;

/// <summary>
/// 配置excel读取时的范围
/// </summary>
public class SimpleRange
{
    public int Row { get; private set; }
    public int? EndRow { get; private set; }
    public int Col { get; private set; }
    public int? EndCol { get; private set; }
    public Func<IEnumerable<object>, bool>? StartFrom { get; set; }
    public Func<IEnumerable<object>, bool>? StopAt { get; set; }
    public SimpleRange()
    {
        (Row, Col) = (0, 0);
    }
    public bool IsDefault()
    {
        return Row == 0 && Col == 0 && StartFrom == null && StopAt == null;
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
    public void EndRowAt(int row)
    {
        EndRow = row;
    }
    public void EndColAt(int col)
    {
        EndCol = col;
    }
}