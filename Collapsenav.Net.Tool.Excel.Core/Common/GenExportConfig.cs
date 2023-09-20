namespace Collapsenav.Net.Tool.Excel;
public partial class ExportConfig<T>
{
    /// <summary>
    /// 根据 T 生成默认的 Config
    /// </summary>
    /// <remarks>
    /// 如果设置了注解,就根据注解生成配置<br/>
    /// 否则就直接根据属性名称生成
    /// </remarks>
    public static ExportConfig<T> GenDefaultConfig(IEnumerable<T>? data = null)
    {
        var type = data.NotEmpty() ? data!.First()!.GetType() : typeof(T);
        // 根据 T 中设置的 ExcelExportAttribute 创建导入配置
        if (type.AttrValues<ExcelExportAttribute>().NotEmpty())
            return GenConfigByAttribute(data);
        // 直接根据属性名称创建导入配置
        return GenConfigByProps(data);
    }
    /// <summary>
    /// 根据 T 中设置的 ExcelExportAttribute 创建导出配置
    /// </summary>
    public static ExportConfig<T> GenConfigByAttribute(IEnumerable<T>? data = null)
    {
        return new ExportConfig<T>(ExcelConfig<T, BaseCellOption<T>>.GenConfigByAttribute<ExcelExportAttribute>()) { Data = data ?? Enumerable.Empty<T>() };
    }
    /// <summary>
    /// 直接根据属性名称创建导出配置
    /// </summary>
    public static ExportConfig<T> GenConfigByProps(IEnumerable<T>? data = null)
    {
        return new ExportConfig<T>(ExcelConfig<T, BaseCellOption<T>>.GenConfigByProps()) { Data = data ?? Enumerable.Empty<T>() };
    }
    /// <summary>
    /// 根据注释生成对应导出配置
    /// </summary>
    /// <remarks>
    /// 需要预先生成注释的xml文档<br/>
    /// 也可以将项目的 GenerateDocumentationFile 设为 True 自动生成xml文档
    /// </remarks>
    public static ExportConfig<T> GenConfigBySummary(IEnumerable<T>? data = null)
    {
        return new ExportConfig<T>(ExcelConfig<T, BaseCellOption<T>>.GenConfigBySummary()) { Data = data ?? Enumerable.Empty<T>() };
    }
}