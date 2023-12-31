using System.Collections;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
namespace Collapsenav.Net.Tool.Excel;

/// <summary>
/// 使用NPOI获取excel的单元格
/// </summary>
public class NPOICellReader : IExcelCellReader
{
    public int Zero => ExcelTool.NPOIZero;
    public ISheet _sheet;
    protected IWorkbook _workbook;
    public Stream? ExcelStream { get; protected set; }
    protected NPOINotCloseStream? notCloseStream;
    public IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected int rowCount;
    protected ISheetCellReader? SheetReader;
    public NPOICellReader(ISheetCellReader sheetReader, string? sheetName = null) : this(sheetReader.SheetStream, sheetName)
    {
        SheetReader = sheetReader;
    }
    public NPOICellReader()
    {
        _workbook = new SXSSFWorkbook();
        _sheet = _workbook.CreateSheet("sheet1");
        rowCount = 0;
        HeaderIndex = new Dictionary<string, int>();
        HeaderList = Enumerable.Empty<string>();
    }
    public NPOICellReader(string path) : this(path.OpenCreateReadWriteShareStream())
    {
    }

    public NPOICellReader(Stream stream, string? sheetName = null) : this(NPOITool.NPOISheet(stream, sheetName))
    {
        stream.SeekToOrigin();
        ExcelStream = stream;
        notCloseStream ??= new NPOINotCloseStream(stream);
    }
    public NPOICellReader(ISheet sheet)
    {
        _sheet = sheet;
        _workbook = sheet.Workbook;

        rowCount = sheet.LastRowNum + 1;
        HeaderIndex = NPOITool.HeadersWithIndex(sheet);
        HeaderList = HeaderIndex.Select(item => item.Key).ToList() ?? Enumerable.Empty<string>();
    }
    public void InitHeader(SimpleRange range)
    {
        if (range.IsDefault())
            return;
        HeaderIndex = NPOITool.HeadersWithIndex(_sheet, range);
        HeaderList = HeaderIndex.Select(item => item.Key).ToList() ?? Enumerable.Empty<string>();
    }
    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers { get => HeaderList; }
    public IDictionary<string, int> HeadersWithIndex { get => HeaderIndex; }
    public IEnumerable<IReadCell> this[string field]
    {
        get
        {
            for (var i = Zero; i < rowCount + Zero; i++)
                yield return new NPOICell(GetCell(GetRow(i), HeaderIndex[field] + Zero));
        }
    }
    public IEnumerable<IReadCell> this[int row] => GetRow(row + Zero).Cells.Select(item => new NPOICell(item));
    public IReadCell this[int row, int col]
    {
        get
        {
            return new NPOICell(GetCell(GetRow(row), col));
        }
    }
    public IReadCell this[string field, int row] => new NPOICell(GetCell(GetRow(row), HeaderIndex[field]));

    public void Dispose()
    {
        ExcelStream?.Dispose();
        notCloseStream?.Dispose();
        _workbook?.Close();
    }
    public void AutoSize()
    {
        if (HeadersWithIndex.NotEmpty())
            foreach (var col in HeadersWithIndex)
                _sheet.AutoSizeColumn(col.Value);
    }
    private IRow GetRow(int row)
    {
        var excelRow = _sheet.GetRow(row);
        excelRow ??= _sheet.CreateRow(row);
        return excelRow;
    }
    private ICell GetCell(IRow row, int col)
    {
        var cell = row.GetCell(col, MissingCellPolicy.RETURN_NULL_AND_BLANK);
        cell ??= row.CreateCell(col);
        return cell;
    }

    public void Save(bool autofit = true)
    {
        if (ExcelStream == null)
            return;
        ExcelStream.Clear();
        notCloseStream ??= new NPOINotCloseStream();
        SaveTo(notCloseStream, autofit);
        notCloseStream.CopyTo(ExcelStream);
    }
    /// <summary>
    /// 保存到流
    /// </summary>
    public void SaveTo(Stream stream, bool autofit = true)
    {
        if (autofit)
            AutoSize();
        stream.Clear();
        using var fs = new NPOINotCloseStream();
        fs.SeekToOrigin();
        _sheet.Workbook.Write(fs, true);
        fs.SeekToOrigin();
        fs.CopyTo(stream);
        stream.SeekToOrigin();
    }

    /// <summary>
    /// 保存到文件
    /// </summary>
    public void SaveTo(string path, bool autofit = true)
    {
        using var fs = path.OpenCreateReadWriteShareStream();
        SaveTo(fs, autofit);
        fs.Dispose();
    }

    /// <summary>
    /// 获取流
    /// </summary>
    public Stream GetStream()
    {
        ExcelStream ??= new MemoryStream();
        notCloseStream ??= new NPOINotCloseStream();
        SaveTo(notCloseStream);
        notCloseStream.CopyTo(ExcelStream);
        notCloseStream.SeekToOrigin();
        ExcelStream.SeekToOrigin();
        return notCloseStream;
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