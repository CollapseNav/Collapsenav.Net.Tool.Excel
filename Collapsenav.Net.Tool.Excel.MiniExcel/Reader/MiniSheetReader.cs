using MiniExcelLibs;

namespace Collapsenav.Net.Tool.Excel;

public class MiniSheetReader : ISheetReader<IExcelReader>
{
    public IEnumerable<IExcelReader> Readers { get; protected set; }
    public Stream SheetStream { get; protected set; }
    private readonly Stream toDispose;
    public IDictionary<string, IExcelReader> Sheets { get; protected set; }
    public IExcelReader this[string sheetName]
    {
        get
        {
            if (Sheets.ContainsKey(sheetName))
                return Sheets[sheetName];
            throw new Exception($"不存在名称为 {sheetName} 的工作簿");
        }
    }

    public IExcelReader this[int index] => Readers.ElementAt(index);
    public MiniSheetReader(string path) : this(path.OpenReadShareStream())
    {
        toDispose.Dispose();
    }
    public MiniSheetReader(Stream stream)
    {
        // copy 传入的流
        SheetStream = new MemoryStream();
        stream.SeekAndCopyTo(SheetStream);

        toDispose = stream;

        var sheetNames = MiniExcel.GetSheetNames(SheetStream);
        Sheets = new Dictionary<string, IExcelReader>();
        sheetNames.ToDictionary(item => item, item => new MiniExcelReader(this, item)).ForEach(item => Sheets.Add(item.Key, item.Value));
        Readers = Sheets.Select(item => item.Value).ToList();
    }
}