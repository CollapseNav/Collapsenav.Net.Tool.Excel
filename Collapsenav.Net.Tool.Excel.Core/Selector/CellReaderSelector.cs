namespace Collapsenav.Net.Tool.Excel;

public class CellReaderSelector
{
    private static Dictionary<ExcelType, Func<Stream, IExcelCellReader>> StreamSelectorDict = null;
    private static Dictionary<ExcelType, Func<object, IExcelCellReader>> ObjSelectorDict = null;
    public static void Add(ExcelType excelType, Func<object, IExcelCellReader> func)
    {
        ObjSelectorDict ??= new();
        ObjSelectorDict.AddOrUpdate(excelType, func);
    }
    public static void Add(ExcelType excelType, Func<Stream, IExcelCellReader> func)
    {
        StreamSelectorDict ??= new();
        StreamSelectorDict.AddOrUpdate(excelType, func);
    }
    public static IExcelCellReader GetCellReader(object obj, ExcelType? excelType = null)
    {
        if (obj == null || ObjSelectorDict.IsEmpty()) return null;
        if (excelType.HasValue && !ObjSelectorDict.ContainsKey(excelType.Value))
            return null;
        if (excelType != null)
            return ObjSelectorDict[excelType.Value](obj);
        foreach (var kv in ObjSelectorDict)
        {
            var reader = kv.Value(obj);
            if (reader != null)
                return reader;
        }
        return null;
    }
    public static IExcelCellReader GetCellReader(Stream stream, ExcelType? excelType = null)
    {
        if (stream == null || StreamSelectorDict.IsEmpty()) return null;
        excelType ??= DefaultExcelType(stream);
        if (!StreamSelectorDict.ContainsKey(excelType.Value))
            return null;
        IExcelCellReader reader = StreamSelectorDict[excelType.Value](stream);
        return reader;
    }

    public static ExcelType? DefaultExcelType(Stream stream)
    {
        if (ObjSelectorDict.Count == 1)
            return ObjSelectorDict.First().Key;
        return stream.Length switch
        {
            >= 5 * 1024 * 1024 => ObjSelectorDict.ContainsKey(ExcelType.MiniExcel) && StreamSelectorDict.ContainsKey(ExcelType.MiniExcel) ? ExcelType.MiniExcel : null,
            <= 5 * 1024 * 1024 => ObjSelectorDict.ElementAt(new Random().Next() % ObjSelectorDict.Count).Key,
        };
    }
}