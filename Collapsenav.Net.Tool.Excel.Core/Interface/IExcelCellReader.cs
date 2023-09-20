namespace Collapsenav.Net.Tool.Excel;

public interface IExcelCellReader : IExcelReader<IReadCell>
{
    /// <summary>
    /// excel文件流
    /// </summary>
    Stream? ExcelStream { get; }
    /// <summary>
    /// 原地保存, 不创建新的文件
    /// </summary>
    void Save(bool autofit = true);
    /// <summary>
    /// 保存到指定的流
    /// </summary>
    /// <param name="stream">指定的流</param>
    /// <param name="autofit">是否自动适配宽度</param>
    void SaveTo(Stream stream, bool autofit = true);
    /// <summary>
    /// 保存到指定文件路径
    /// </summary>
    /// <param name="path">指定的文件路径</param>
    /// <param name="autofit">是否自动适配宽度</param>
    void SaveTo(string path, bool autofit = true);
    /// <summary>
    /// 获取对应excel流
    /// </summary>
    Stream GetStream();
#if NET6_0_OR_GREATER && NETCOREAPP
    public static IExcelCellReader GetCellReader(object sheet)
    {
        return CellReaderSelector.GetCellReader(sheet);
    }
    public static IExcelCellReader GetCellReader(Stream stream, ExcelType? excelType = null)
    {
        return CellReaderSelector.GetCellReader(stream, excelType.ToString());
    }
    public static IExcelCellReader GetCellReader(string path, ExcelType? excelType = null)
    {
        using var fs = path.OpenReadWriteShareStream();
        return GetCellReader(fs, excelType);
    }
#endif
}
