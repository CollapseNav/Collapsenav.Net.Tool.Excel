using System.Diagnostics;
using System.Linq.Expressions;
using MiniExcelLibs;
namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 使用MiniExcel获取excel的单元格
/// </summary>
public class MiniCellReader : IExcelCellReader
{
    public int Zero => ExcelTool.MiniZero;
    protected List<IDictionary<string, object>> _sheet;
    protected List<MiniRow> _rows;
    protected Stream _stream;
    protected int rowCount;
    protected int colCount;
    public MiniCellReader()
    {
        Init();
    }
    public MiniCellReader(string path)
    {
        var fs = path.OpenCreateReadWriteShareStream();
        Init(fs);
    }
    public MiniCellReader(Stream stream)
    {
        Init(stream);
    }
    private void Init(Stream stream)
    {
        try
        {
            _ = stream.Query().First();
            _stream = stream;
            _sheet = _stream.Query() as List<IDictionary<string, object>>;
            _rows = new List<MiniRow>();
            rowCount = _sheet.Count;
            colCount = _sheet.First().Count;
        }
        catch
        {
            Init();
            _stream = stream;
        }
    }
    private void Init()
    {
        _stream = new MemoryStream();
        _sheet = new List<IDictionary<string, object>>(10000);
        _rows = new List<MiniRow>(10000);
        rowCount = 0;
    }


    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers
    {
        get
        {
            var sheetFirst = _sheet.First();
            return sheetFirst.Select(item => item.Value?.ToString() ?? string.Empty);
        }
    }
    public IDictionary<string, int> HeadersWithIndex
    {
        get
        {
            var sheetFirst = _sheet.First();
            return sheetFirst.Select((item, index) => (item.Value, index)).ToDictionary(item => item.Value?.ToString() ?? item.index.ToString(), item => item.index);
        }
    }
    public IEnumerable<IReadCell> this[string field]
    {
        get
        {
            var data = HeadersWithIndex;
            int index = data[field] + Zero;
            for (var row = Zero; row < rowCount + Zero; row++)
                yield return _rows[row].Cells[index];
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
        for (var rowNum = _rows.Count; rowNum <= row; rowNum++)
        {
            IDictionary<string, object> newRow = new Dictionary<string, object>();
            _sheet.Add(newRow);
            _rows.Add(new MiniRow(newRow, rowNum));
        }
        rowCount = _rows.Count;
        return _rows.Last();
    }

    public void Dispose()
    {
        _stream.Dispose();
    }

    public void Save(bool autofit = true)
    {
        SaveTo(_stream);
    }

    /// <summary>
    /// 保存到流
    /// </summary>
    public void SaveTo(Stream stream, bool autofit = true)
    {
        Console.WriteLine("save");
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
        SaveTo(_stream);
        return _stream;
    }

    public IEnumerator<IEnumerable<IReadCell>> GetEnumerator()
    {
        for (var row = 0; row < rowCount + Zero; row++)
            yield return this[row];
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
