namespace Collapsenav.Net.Tool.Excel;

public class MiniTool
{
    public static IDictionary<string, int> HeadersWithIndex(IEnumerable<dynamic> sheet, SimpleRange? range = null)
    {
        IEnumerable<KeyValuePair<string, object>>? sheetFirst = null;
        if (range == null)
        {
            sheetFirst = (sheet.FirstOrDefault() as IEnumerable<KeyValuePair<string, object>>) ?? Enumerable.Empty<KeyValuePair<string, object>>();
        }
        else if (range.StartFrom == null)
        {
            sheetFirst = (sheet.Skip(range.Row).FirstOrDefault() as IEnumerable<KeyValuePair<string, object>>) ?? Enumerable.Empty<KeyValuePair<string, object>>();
        }
        else
        {
            for (var i = ExcelTool.MiniZero; i < sheet.Count(); i++)
            {
                sheetFirst = sheet.Skip(i).FirstOrDefault() as IEnumerable<KeyValuePair<string, object>>;
                if (sheetFirst.NotEmpty() && range.StartFrom(sheetFirst!.Select(item => item.Value?.ToString() ?? string.Empty)))
                {
                    range.SkipRow(i - ExcelTool.MiniZero);
                    break;
                }
            }
            sheetFirst ??= Enumerable.Empty<KeyValuePair<string, object>>();
        }
        return sheetFirst.Select((item, index) => (item.Value, index)).ToDictionary(item => item.Value?.ToString() ?? item.index.ToString(), item => item.index);
    }
}