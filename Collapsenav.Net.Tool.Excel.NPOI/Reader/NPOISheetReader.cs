namespace Collapsenav.Net.Tool.Excel;

public class NPOISheetReader : ISheetReader<IExcelReader>
{
    public IExcelReader this[int index] => Readers.ElementAt(index);

    public IExcelReader this[string sheetName]
    {
        get
        {
            if (Sheets.ContainsKey(sheetName))
                return Sheets[sheetName];
            throw new Exception($"不存在名称为 {sheetName} 的工作簿");
        }
    }

    public Stream SheetStream { get; private set; }
    private readonly Stream toDispose;
    public IEnumerable<IExcelReader> Readers { get; private set; }
    public IDictionary<string, IExcelReader> Sheets { get; private set; }

    public NPOISheetReader(string path) : this(path.OpenReadShareStream())
    {
        toDispose.Dispose();
    }

    public NPOISheetReader(Stream stream)
    {
        SheetStream = new MemoryStream();
        stream.SeekAndCopyTo(SheetStream);

        toDispose = stream;

        var workBook = NPOITool.NPOIWorkbook(SheetStream);
        List<string> sheetNames = new();
        Sheets = new Dictionary<string, IExcelReader>();
        foreach (var sheet in workBook)
            Sheets.Add(sheet.SheetName, new NPOIExcelReader(sheet));
        Readers = Sheets.Select(item => item.Value).ToList();
    }
}