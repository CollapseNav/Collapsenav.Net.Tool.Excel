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
}
