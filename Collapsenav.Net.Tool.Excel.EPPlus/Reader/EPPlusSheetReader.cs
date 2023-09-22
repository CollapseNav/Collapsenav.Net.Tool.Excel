namespace Collapsenav.Net.Tool.Excel;

public class EPPlusSheetReader : ISheetReader<IExcelReader>
{
    public IExcelReader this[int index] => Readers.ElementAt(index);

    public IExcelReader this[string sheetName]
    {
        get
        {
            if (Sheets.ContainsKey(sheetName))
                return Sheets[sheetName];
            throw new Exception();
        }
    }

    public Stream SheetStream { get; private set; }
    private readonly Stream toDispose;

    public IEnumerable<IExcelReader> Readers { get; private set; }

    public IDictionary<string, IExcelReader> Sheets { get; private set; }

    public EPPlusSheetReader(string path) : this(path.OpenReadShareStream())
    {
        toDispose.Dispose();
    }

    public EPPlusSheetReader(Stream stream)
    {
        SheetStream = new MemoryStream();
        stream.SeekAndCopyTo(SheetStream);
        toDispose = stream;
        var workSheets = EPPlusTool.EPPlusSheets(SheetStream);
        var sheetNames = workSheets.Select(item => item.Name).ToList();
        Sheets = new Dictionary<string, IExcelReader>();
        sheetNames.ToDictionary(item => item, item => new EPPlusExcelReader(workSheets[item])).ForEach(item => Sheets.Add(item.Key, item.Value));
        Readers = Sheets.Select(item => item.Value).ToList();
    }
}