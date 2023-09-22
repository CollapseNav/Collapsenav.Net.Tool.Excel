namespace Collapsenav.Net.Tool.Excel;
public static class Init
{
    public static void InitTypeSelector()
    {
        Selector.AddTypeSelector(ExcelType.MiniExcel,
            obj => obj is Stream,
            stream => stream.Length > 5 * 1024 ? 99 : 50
        );
        Selector.AddCellSelector(ExcelType.MiniExcel,
            obj => (obj is Stream stream) ? new MiniCellReader(stream) : throw new Exception(),
            stream => new MiniCellReader(stream)
        );
        Selector.AddExcelSelector(ExcelType.MiniExcel,
            obj => (obj is Stream stream) ? new MiniExcelReader(stream) : throw new Exception(),
            stream => new MiniExcelReader(stream)
        );
    }
}