using System.Collections;
using MiniExcelLibs;
namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 使用MiniExcel获取excel的单元格
/// </summary>
/// <remarks>
/// 由于提供了将修改应用到传入的流或者文件中的功能<br/>
/// 所以会长期保持文件为打开的状态
/// </remarks>
public class MiniCellReader : IExcelCellReader
{
    public int Zero => ExcelTool.MiniZero;
    public List<IDictionary<string, object>> _sheet;
    protected IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected List<MiniRow> _rows;
    public Stream? ExcelStream { get; protected set; }
    protected int rowCount;
    protected int colCount;
    protected ISheetCellReader? SheetReader;
    public MiniCellReader(ISheetCellReader sheetReader, string? sheetName = null) : this(sheetReader.SheetStream, sheetName)
    {
        SheetReader = sheetReader;
    }
    public MiniCellReader(string path) : this(path.OpenCreateReadWriteShareStream()) { }
    public MiniCellReader(Stream? stream = null, string? sheetName = null)
    {
        _rows = new List<MiniRow>(1000);
        // 如果是一个非空的文件流
        if (stream != null)
        {
            ExcelStream = stream;

            // 初始化 _sheet
            IEnumerable<IDictionary<string, object>>? temp;
            if (sheetName.NotEmpty())
                temp = ExcelStream.Query(sheetName: sheetName) as IEnumerable<IDictionary<string, object>>;
            else
                temp = ExcelStream.Query() as IEnumerable<IDictionary<string, object>>;
            _sheet = temp?.ToList() ?? new List<IDictionary<string, object>>();

            foreach (var (data, index) in _sheet.SelectWithIndex())
                _rows.Add(new MiniRow(data, index));

            rowCount = _sheet.Count;
            colCount = _sheet.First().Count;

            var sheetFirst = _sheet.First();
            HeaderList = sheetFirst.Select(item => item.Value?.ToString() ?? string.Empty).ToList();
            HeaderIndex = sheetFirst.Select((item, index) => (item.Value, index)).ToDictionary(item => item.Value?.ToString() ?? item.index.ToString(), item => item.index);
        }
        else
        {
            ExcelStream = new MemoryStream();
            _sheet = new List<IDictionary<string, object>>(1000);
            rowCount = 0;
            HeaderList = Enumerable.Empty<string>();
            HeaderIndex = new Dictionary<string, int>();
        }
    }


    public void InitHeader(SimpleRange range)
    {
        if (range.IsDefault())
            return;
        HeaderIndex = MiniTool.HeadersWithIndex(_sheet, range);
        HeaderList = HeaderIndex.Select(item => item.Key);
    }

    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers => HeaderList;
    public IDictionary<string, int> HeadersWithIndex => HeaderIndex;
    public IEnumerable<IReadCell> this[string field]
    {
        get
        {
            var data = HeadersWithIndex;
            int index = data[field] + Zero;
            for (var row = Zero; row < rowCount + Zero; row++)
                yield return GetMiniRow(row)[index];
        }
    }

    public IEnumerable<IReadCell> this[int row]
    {
        get
        {
            return GetMiniRow(row + Zero).Cells;
        }
    }
    public IReadCell this[int row, int col]
    {
        get
        {
            colCount = colCount <= col ? col : colCount;
            return GetMiniRow(row)[col];
        }
    }
    public IReadCell this[string field, int row]
    {
        get
        {
            var data = HeadersWithIndex;
            return GetMiniRow(row)[data[field] + Zero];
        }
    }

    private MiniRow GetMiniRow(int row)
    {
        if (row < rowCount)
            return _rows[row];
        KeyValuePair<string, object>[] kvs = Enumerable.Range(0, colCount + 1).Select(item => new KeyValuePair<string, object>(MiniCell.GetSCol(item), "")).ToArray();
        for (var rowNum = _rows.Count; rowNum <= row; rowNum++)
        {
#if NETSTANDARD2_0
            IDictionary<string, object> newRow = new Dictionary<string, object>();
            foreach (var kv in kvs)
                newRow.Add(kv);
#else
            IDictionary<string, object> newRow = new Dictionary<string, object>(kvs);
#endif
            _sheet.Add(newRow);
            _rows.Add(new MiniRow(newRow, rowNum));
        }
        rowCount = _rows.Count;
        return _rows.Last();
    }

    public void Dispose()
    {
        ExcelStream?.Dispose();
    }

    public void Save(bool autofit = true)
    {
        SaveTo(ExcelStream);
    }

    /// <summary>
    /// 保存到流
    /// </summary>
    public void SaveTo(Stream? stream, bool autofit = true)
    {
        if (stream == null)
            throw new NullReferenceException();
        stream.SeekToOrigin();
        stream.Clear();
        stream.SaveAs(_sheet, printHeader: false);
        stream.SeekToOrigin();
    }
    /// <summary>
    /// 保存到文件
    /// </summary>
    public void SaveTo(string path, bool autofit = true)
    {
        using var fs = path.OpenCreateReadWriteShareStream();
        SaveTo(fs);
        fs.Dispose();
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
        for (var row = 0; row < rowCount + Zero; row++)
            yield return this[row];
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
