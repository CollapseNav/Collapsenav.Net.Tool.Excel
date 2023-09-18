using Collapsenav.Net.Tool.Excel;
// var config = ReadConfig<CylinderData>.GenConfigBySummary();
// var data = config.ToEntity("./test.xlsx");
var data = Enumerable.Empty<CylinderData>();
var dd = data.GetType();
var d = dd.GenericTypeArguments.First();
Console.WriteLine();

public class CylinderData
{
    /// <summary>
    /// 序号
    /// </summary>
    public string? Index { get; set; }
    /// <summary>
    /// 充装单位
    /// </summary>
    public string? FillingUnit { get; set; }
    /// <summary>
    /// 制造单位
    /// </summary>
    public string? ManufacturingUnit { get; set; }
    /// <summary>
    /// 出厂编号
    /// </summary>
    public string? FactoryNo { get; set; }
    /// <summary>
    /// 制造年月
    /// </summary>
    public string? MakeMonth { get; set; }
    /// <summary>
    /// 气瓶种类
    /// </summary>
    public string? CylinderType { get; set; }
    /// <summary>
    /// 气瓶规格
    /// </summary>
    public string? Specifications { get; set; }
    /// <summary>
    /// 充装介质
    /// </summary>
    public string? Medium { get; set; }
    /// <summary>
    /// 气瓶状态
    /// </summary>
    public string? Status { get; set; }
    /// <summary>
    /// 气瓶条码
    /// </summary>
    public string? Barcode { get; set; }
    /// <summary>
    /// 上报单位
    /// </summary>
    public string? ReportingUnit { get; set; }
    /// <summary>
    /// 修改单位
    /// </summary>
    public string? ModifyUnit { get; set; }
    /// <summary>
    /// 末次修改时间
    /// </summary>
    public DateTime? LastModificationTime { get; set; }
    /// <summary>
    /// 自有编号
    /// </summary>
    public string? SelfNumber { get; set; }
    /// <summary>
    /// 公称工作压力
    /// </summary>
    public decimal? Pressure { get; set; }
    /// <summary>
    /// 容积
    /// </summary>
    public decimal? Volume { get; set; }
    /// <summary>
    /// 设计壁厚
    /// </summary>
    public decimal? Thickness { get; set; }
    /// <summary>
    /// 末次检验年月
    /// </summary>
    public DateTime? LastInspection { get; set; }
    /// <summary>
    /// 下次检验年月
    /// </summary>
    public DateTime? NextInspection { get; set; }
    /// <summary>
    /// 使用登记代码
    /// </summary>
    public string? RegistrationCode { get; set; }
    /// <summary>
    /// 报废年月
    /// </summary>
    public DateTime? ScrapDate { get; set; }
    /// <summary>
    /// 皮重
    /// </summary>
    public decimal? TareWeight { get; set; }
    /// <summary>
    /// 上报时间
    /// </summary>
    public DateTime? ReportTime { get; set; }
}


