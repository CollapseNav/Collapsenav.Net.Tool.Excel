using MiniExcelLibs;

namespace Collapsenav.Net.Tool.Excel;

public class MiniExcelReader : IExcelReader
{
    public int Zero => ExcelTool.MiniZero;
    protected Stream SheetStream;
    private readonly Stream toDispose;
    protected IEnumerable<dynamic> sheet;
    protected IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected int rowCount;
    protected ISheetReader<IExcelReader>? SheetReader;
    public MiniExcelReader(ISheetReader<IExcelReader> sheetReader, string sheetName) : this(sheetReader.SheetStream, sheetName)
    {
        SheetReader = sheetReader;
    }
    /// <summary>
    /// 读取之后会复制一份, 并且关闭原来的文件流, 解除文件占用
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    public MiniExcelReader(string path) : this(path.OpenReadWriteShareStream())
    {
        // 通过文件读取的时候需要将原来的文件释放
        toDispose.Dispose();
    }

    /// <summary>
    /// 会复制传入的文件流, 不会关闭传入的流
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="sheetName"></param>
    public MiniExcelReader(Stream stream, string? sheetName = null)
    {
        SheetStream = new MemoryStream();
        stream.SeekAndCopyTo(SheetStream);
        // 记录传入的流，该构造函数不会关闭原来的流
        toDispose = stream;

        if (sheetName.NotEmpty())
            sheet = SheetStream.Query(sheetName: sheetName!).ToList();
        else
            sheet = SheetStream.Query().ToList();
        rowCount = sheet.Count();
        var sheetFirst = (sheet.FirstOrDefault() as IEnumerable<KeyValuePair<string, object>>) ?? Enumerable.Empty<KeyValuePair<string, object>>();
        HeaderList = sheetFirst.Select(item => item.Value?.ToString() ?? string.Empty);
        HeaderIndex = sheetFirst.Select((item, index) => (item.Value, index)).ToDictionary(item => item.Value?.ToString() ?? item.index.ToString(), item => item.index);
    }
    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers { get => HeaderList; }
    public IDictionary<string, int> HeadersWithIndex { get => HeaderIndex; }
    public IEnumerable<string> this[string field]
    {
        get
        {
            for (var i = Zero; i < rowCount + Zero; i++)
            {
                if (sheet.ElementAt(i) is not IEnumerable<KeyValuePair<string, object>> row)
                    yield return string.Empty;
                else
                    yield return row.ElementAt(HeaderIndex[field] + Zero).Value?.ToString() ?? string.Empty;
            }
        }
    }

    public IEnumerable<string> this[int row]
    {
        get
        {
            if (sheet.ElementAt(row) is not IEnumerable<KeyValuePair<string, object>> rowData)
                return Enumerable.Empty<string>();
            else
                return rowData.Select(item => item.Value?.ToString() ?? string.Empty);
        }
    }
    public string this[int row, int col]
    {
        get
        {
            if (sheet.ElementAt(row) is not IEnumerable<KeyValuePair<string, object>> rowData)
                return string.Empty;
            else
                return rowData.ElementAt(col + Zero).Value?.ToString() ?? string.Empty;
        }
    }
    public string this[string field, int row]
    {
        get
        {
            if (sheet.ElementAt(row) is not IEnumerable<KeyValuePair<string, object>> rowData)
                return string.Empty;
            else
                return rowData.ElementAt(HeaderIndex[field] + Zero).Value?.ToString() ?? string.Empty;
        }
    }

    public void Dispose()
    {
        SheetStream.Dispose();
    }

    public IEnumerator<IEnumerable<string>> GetEnumerator()
    {
        for (var row = 0; row < rowCount; row++)
        {
            if (sheet.ElementAt(row) is not IEnumerable<KeyValuePair<string, object>> rowData)
                yield return Enumerable.Empty<string>();
            else
                yield return rowData.Select(item => item.Value?.ToString() ?? string.Empty);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
