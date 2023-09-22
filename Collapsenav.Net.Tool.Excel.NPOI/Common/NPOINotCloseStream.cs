using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Collapsenav.Net.Tool.Excel;
/// <summary>
/// 专门为NPOI做的不关闭流
/// </summary>
/// <remarks>
/// 一般的npoi会自动关闭流，影响到后面的使用
/// </remarks>
public class NPOINotCloseStream : MemoryStream
{
    public bool? IsXlsx { get; set; }
    public NPOINotCloseStream() { }
    public NPOINotCloseStream(Stream stream)
    {
        stream.SeekAndCopyTo(this);
        stream.SeekToOrigin();
        this.SeekToOrigin();
        if (stream is FileStream fs)
        {
            if (fs.Name.IsXlsx())
                IsXlsx = true;
        }
    }
    public NPOINotCloseStream(string path)
    {
        using var fs = path.OpenCreateReadWriteShareStream();
        fs.CopyTo(this);
        fs.Dispose();
        this.SeekToOrigin();
        if (path.IsXlsx())
            IsXlsx = true;
    }
    public override void Close() { }
    public IWorkbook GetWorkBook()
    {
        IWorkbook workbook;
        if (IsXlsx.HasValue)
        {
            workbook = IsXlsx == true ? XSSF() : HSSF();
        }
        else
        {
            try
            {
                workbook = XSSF();
            }
            catch
            {
                Seek(0, SeekOrigin.Begin);
                workbook = HSSF();
            }
        }
        return workbook;

        HSSFWorkbook HSSF()
        {
            return Length > 0 ? new HSSFWorkbook(this) : new HSSFWorkbook();
        }

        IWorkbook XSSF()
        {
            return Length > 0 ? new XSSFWorkbook(this) : new XSSFWorkbook();
        }
    }
}