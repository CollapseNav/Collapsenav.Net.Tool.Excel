using System.Collections;
using OfficeOpenXml;

namespace Collapsenav.Net.Tool.Excel;

public class EPPlusExcelReader : IExcelReader
{
    private object[,] sheetData;
    public int Zero => ExcelTool.EPPlusZero;
    protected ExcelWorksheet sheet;
    protected IDictionary<string, int> HeaderIndex;
    protected IEnumerable<string> HeaderList;
    protected int rowCount;
    private readonly Stream? toDispose;
    public EPPlusExcelReader(string path) : this(path.OpenReadShareStream())
    {
        toDispose!.Dispose();
    }
    public EPPlusExcelReader(Stream stream, string? sheetName = null) : this(EPPlusTool.EPPlusSheet(stream, sheetName))
    {
        toDispose = stream;
    }
    public EPPlusExcelReader(ExcelWorksheet sheet)
    {
        this.sheet = sheet;

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
        sheetData = (sheet.Cells.Value as object[,]) ?? new object[0, 0];
    }
    public void InitHeader(SimpleRange range)
    {
        HeaderIndex = EPPlusTool.HeadersWithIndex(sheet, range);
        HeaderList = HeaderIndex.Select(item => item.Key).ToList();
    }
    public IEnumerable<string> this[string field]
    {
        get
        {
            for (var i = Zero; i < rowCount + Zero; i++)
            {
                yield return sheet.Cells[i, HeaderIndex[field] + Zero].Value.ToString() ?? string.Empty;
            }
        }
    }
    public IEnumerable<string> this[int row]
    {
        get
        {
            List<string> _data = new(HeaderList.Count());
            foreach (var h in HeadersWithIndex)
                _data.Add(sheetData[row, h.Value]?.ToString() ?? string.Empty);
            return _data;
        }
    }
    public string this[int row, int col] => sheetData[row, col].ToString() ?? string.Empty;
    public string this[string field, int row] => sheetData[row, HeaderIndex[field]].ToString() ?? string.Empty;
    public IEnumerable<string> Headers => HeaderList;
    public IDictionary<string, int> HeadersWithIndex => HeaderIndex;
    public int RowCount => rowCount;
    public void Dispose()
    {
        sheet.Workbook.Dispose();
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