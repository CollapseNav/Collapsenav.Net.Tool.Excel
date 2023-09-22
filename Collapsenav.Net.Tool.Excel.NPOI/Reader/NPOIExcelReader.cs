using System.Collections;
using NPOI.SS.UserModel;
namespace Collapsenav.Net.Tool.Excel;

public class NPOIExcelReader : IExcelReader
{
    public int Zero => ExcelTool.NPOIZero;
    protected ISheet sheet;
    public IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected int rowCount;
    public int RowCount { get => rowCount; }
    public IEnumerable<string> Headers { get => HeaderList; }
    public IDictionary<string, int> HeadersWithIndex { get => HeaderIndex; }
    private readonly Stream? toDispose;
    public NPOIExcelReader(string path) : this(path.OpenReadShareStream())
    {
        toDispose!.Dispose();
    }
    public NPOIExcelReader(Stream stream, string? sheetName = null) : this(NPOITool.NPOISheet(stream, sheetName))
    {
        toDispose = stream;
    }
    public NPOIExcelReader(ISheet sheet)
    {
        this.sheet = sheet;
        rowCount = sheet.LastRowNum + 1;
        HeaderIndex = NPOITool.HeadersWithIndex(sheet);
        HeaderList = HeaderIndex.Select(item => item.Key).ToList();
    }

    public IEnumerable<string> this[string field]
    {
        get
        {
            for (var i = Zero; i < rowCount + Zero; i++)
                yield return sheet.GetRow(i).GetCell(HeaderIndex[field] + Zero).ToString() ?? string.Empty;
        }
    }
    public IEnumerable<string> this[int row] => sheet.GetRow(row + Zero).Select(item => item.ToString() ?? string.Empty);
    public string this[int row, int col] => sheet.GetRow(row).GetCell(col).ToString() ?? string.Empty;
    public string this[string field, int row] => sheet.GetRow(row).GetCell(HeaderIndex[field]).ToString() ?? string.Empty;

    public void Dispose()
    {
        sheet.Workbook.Close();
        // toDispose?.Dispose();
    }

    public IEnumerator<IEnumerable<string>> GetEnumerator()
    {
        for (var row = 0; row < rowCount; row++)
            yield return this[row];
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}