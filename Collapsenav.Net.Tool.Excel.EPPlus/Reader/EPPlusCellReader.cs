using System.Collections;
using OfficeOpenXml;
namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 使用EPPlus获取excel的单元格
/// </summary>
public class EPPlusCellReader : IExcelCellReader
{
    public int Zero => ExcelTool.EPPlusZero;
    public ExcelWorksheet _sheet;
    protected ExcelPackage _pack;
    public Stream? ExcelStream { get; protected set; }
    protected IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected int rowCount;
    protected ISheetCellReader? SheetReader;
    public EPPlusCellReader(ISheetCellReader sheetReader, string? sheetName = null) : this(sheetReader.SheetStream, sheetName)
    {
        SheetReader = sheetReader;
    }
    public EPPlusCellReader()
    {
        _pack = new ExcelPackage();
        _sheet = _pack.Workbook.Worksheets.Add("sheet1");
        HeaderList = Enumerable.Empty<string>();
        HeaderIndex = new Dictionary<string, int>();
        rowCount = 0;
    }
    public EPPlusCellReader(string path) : this(path.OpenCreateReadWriteShareStream())
    {
    }
    public EPPlusCellReader(Stream stream, string? sheetName = null) : this(EPPlusTool.EPPlusSheet(stream, sheetName))
    {
        ExcelStream = stream;
    }
    public EPPlusCellReader(ExcelWorksheet sheet)
    {
        _sheet = sheet;
        _pack ??= new ExcelPackage();
        if (_pack.Workbook.Worksheets.Count == 0)
            _sheet = _pack.Workbook.Worksheets.Add("sheet1", sheet);

        if (sheet.Dimension != null)
        {
            rowCount = sheet.Dimension?.Rows ?? 0;
            HeaderIndex = EPPlusTool.HeadersWithIndex(sheet);
            HeaderList = HeaderIndex.Select(item => item.Key).ToList();
        }
        else
        {
            HeaderIndex = new Dictionary<string, int>();
            HeaderList = Enumerable.Empty<string>();
        }
    }
    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers { get => HeaderList; }
    public IDictionary<string, int> HeadersWithIndex { get => HeaderIndex; }
    public IEnumerable<IReadCell> this[string field]
    {
        get
        {
            for (var i = Zero; i < rowCount + Zero; i++)
                yield return new EPPlusCell(_sheet.Cells[i, HeaderIndex[field] + Zero]);
        }
    }
    public IEnumerable<IReadCell> this[int row] => _sheet.Cells[row + Zero, Zero, row + Zero, Zero + Headers.Count()].Select(item => new EPPlusCell(item));
    public IReadCell this[int row, int col] => new EPPlusCell(_sheet.Cells[row + Zero, col + Zero]);
    public IReadCell this[string field, int row] => new EPPlusCell(_sheet.Cells[row + Zero, HeaderIndex[field] + Zero]);
    public void Dispose()
    {
        ExcelStream?.Dispose();
        _pack?.Dispose();
    }
    public void AutoSize()
    {
        if (HeadersWithIndex.NotEmpty())
            foreach (var col in HeadersWithIndex)
                _sheet.Column(col.Value + Zero).AutoFit();
    }
    public void Save(bool autofit = true)
    {
        if (ExcelStream == null)
            throw new NullReferenceException();
        SaveTo(ExcelStream, autofit);
    }
    /// <summary>
    /// 保存到流
    /// </summary>
    public void SaveTo(Stream stream, bool autofit = true)
    {
        if (autofit)
            AutoSize();
        stream.SeekToOrigin();
        stream.Clear();
        _pack.SaveAs(stream);
        stream.SeekToOrigin();
    }
    /// <summary>
    /// 保存到文件
    /// </summary>
    public void SaveTo(string path, bool autofit = true)
    {
        using var fs = path.OpenCreateReadWriteShareStream();
        SaveTo(fs, autofit);
    }
    /// <summary>
    /// 获取流
    /// </summary>
    public Stream GetStream()
    {
        ExcelStream ??= new MemoryStream();
        SaveTo(ExcelStream);
        return ExcelStream;
    }
    public IEnumerator<IEnumerable<IReadCell>> GetEnumerator()
    {
        for (var row = 0; row < rowCount; row++)
            yield return this[row];
    }
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
