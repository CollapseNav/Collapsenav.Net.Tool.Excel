using System.Reflection;

namespace Collapsenav.Net.Tool.Excel;

/// <summary>
/// 表格配置
/// </summary>
public class ExcelConfig<T, CellOption> : IExcelConfig<T, CellOption> where CellOption : ICellOption, new()
{
    public ExcelConfig()
    {
        FieldOption = new List<CellOption>();
        DtoType = typeof(T);
        Range = new SimpleRange();
    }
    public ExcelConfig(IEnumerable<(string, string)>? kvs) : this()
    {
        InitFieldOption(kvs);
    }
    public ExcelConfig(IEnumerable<StringCellOption>? options) : this()
    {
        if (options == null || options.IsEmpty())
            FieldOption = Enumerable.Empty<CellOption>();
        else
        {
            FieldOption = options.Select(item => new CellOption
            {
                ExcelField = item.FieldName,
                PropName = item.PropName,
            }).ToList();
        }
    }
    public Type DtoType { get; protected set; }
    public IEnumerable<CellOption> FieldOption { get; set; }
    public SimpleRange Range { get; private set; }
    public IEnumerable<T> Data
    {
        get => data ?? Enumerable.Empty<T>(); set
        {
            data = value;
            if (data.NotEmpty())
                DtoType = data!.GetType().GenericTypeArguments.First();
            else
                DtoType = typeof(T);
        }
    }
    private IEnumerable<T>? data;
    public virtual IEnumerable<string> Headers { get => FieldOption.Select(item => item.ExcelField ?? string.Empty); }
    protected IDictionary<string, int>? headerIndex;
    public virtual IDictionary<string, int> HeadersWithIndex
    {
        get
        {
            headerIndex ??= Headers.SelectWithIndex().ToDictionary(item => item.value, item => item.index);
            return headerIndex;
        }
        protected set => headerIndex = value;
    }
    /// <summary>
    /// 通过元组初始化配置
    /// </summary>
    /// <param name="kvs"></param>
    public virtual void InitFieldOption(IEnumerable<(string Key, string Value)>? kvs)
    {
        FieldOption = new List<CellOption>();
        if (kvs.NotEmpty())
        {
            foreach (var (Key, Value) in kvs!)
                Add(Key, Value);
        }
    }
    /// <summary>
    /// 根据给出的表头筛选options
    /// </summary>
    public void FilterOptionByHeaders(IEnumerable<string>? headers)
    {
        FieldOption = FilterOptionByHeaders(FieldOption, headers).ToList();
    }
    /// <summary>
    /// 根据给出的表头筛选options
    /// </summary>
    public static IEnumerable<CellOption> FilterOptionByHeaders(IEnumerable<CellOption>? options, IEnumerable<string>? headers)
    {
        if (options == null || headers == null)
            return Enumerable.Empty<CellOption>();
        return options.Where(item => headers.Where(head => head == item.ExcelField).Any());
    }
    public virtual IExcelConfig<T, CellOption> Add(CellOption option)
    {
        FieldOption = FieldOption.Append(option);
        return this;
    }
    public virtual IExcelConfig<T, CellOption> AddIf(bool check, CellOption option)
    {
        return check ? Add(option) : this;
    }
    public virtual IExcelConfig<T, CellOption> Add(string field, PropertyInfo prop)
    {
        FieldOption = FieldOption.Append(new CellOption { ExcelField = field, Prop = prop });
        return this;
    }
    public virtual IExcelConfig<T, CellOption> AddIf(bool check, string field, PropertyInfo prop)
    {
        return check ? Add(field, prop) : this;
    }
    public virtual IExcelConfig<T, CellOption> Add(string field, string propName)
    {
        FieldOption = FieldOption.Append(new CellOption { ExcelField = field, PropName = propName });
        return this;
    }
    public virtual IExcelConfig<T, CellOption> AddIf(bool check, string field, string propName)
    {
        return check ? Add(field, propName) : this;
    }
    public virtual IExcelConfig<T, CellOption> Remove(string field)
    {
        FieldOption = FieldOption.Where(item => item.ExcelField != field).ToList();
        return this;
    }

    /// <summary>
    /// 直接根据属性名称创建配置
    /// </summary>
    public static IExcelConfig<T, CellOption> GenConfigByProps(IExcelConfig<T, CellOption>? config = null)
    {
        config ??= new ExcelConfig<T, CellOption>();
        config.ClearFieldOption();
        foreach (var prop in config.DtoType.Props().Where(i => i.PropertyType.IsBaseType()))
            config.Add(prop.Name, prop);
        return config;
    }

    /// <summary>
    /// 根据 T 中设置的 ExcelAttribute 创建配置
    /// </summary>
    public static IExcelConfig<T, CellOption> GenConfigByAttribute<Attr>(IExcelConfig<T, CellOption>? config = null) where Attr : ExcelAttribute
    {
        config ??= new ExcelConfig<T, CellOption>();
        config.ClearFieldOption();
        var attrData = config.DtoType.AttrValue<Attr>();
        foreach (var prop in attrData)
        {
            if (prop.Value != null && prop.Value.ExcelField.NotEmpty())
                config.Add(prop.Value!.ExcelField, prop.Key);
        }
        return config;
    }

    /// <summary>
    /// 通过注释生成配置
    /// </summary>
    public static IExcelConfig<T, CellOption> GenConfigBySummary(Type? type = null, IExcelConfig<T, CellOption>? config = null)
    {
        config ??= new ExcelConfig<T, CellOption>();
        config.ClearFieldOption();
        type ??= config.DtoType;
        // 获取所有注释node
        var nodes = XmlExt.GetXmlDocuments().GetSummaryNodes().Where(item => item.FullName.Contains(type.FullName ?? string.Empty)).ToArray();
        var kvs = type!.Props().Where(i => i.PropertyType.IsBaseType()).Select(i => i.Name)?.Select(propName =>
        {
            var node = nodes.FirstOrDefault(item => item.FullName.Split('.').Last() == propName);
            if (node is null)
                return new KeyValuePair<string, PropertyInfo?>(propName, type.GetProperty(propName));
            return new KeyValuePair<string, PropertyInfo?>(node.Summary ?? string.Empty, type.GetProperty(propName));
        }).ToArray();
        if (kvs.NotEmpty())
        {
            foreach (var node in kvs!)
                config.Add(new CellOption { ExcelField = node.Key, Prop = node.Value });
        }
        return config;
    }

    public virtual void ClearData()
    {
        Data = new List<T>();
    }

    public void ClearFieldOption()
    {
        FieldOption = new List<CellOption>();
    }

    public virtual IExcelConfig<T, CellOption> SkipRow(int row)
    {
        Range.SkipRow(row);
        return this;
    }

    public virtual IExcelConfig<T, CellOption> StartFrom(Func<IEnumerable<object>, bool> selectRow)
    {
        Range.StartFrom = selectRow;
        return this;
    }

    public virtual IExcelConfig<T, CellOption> StopAt(Func<IEnumerable<object>, bool> selectRow)
    {
        Range.StopAt = selectRow;
        return this;
    }
}
