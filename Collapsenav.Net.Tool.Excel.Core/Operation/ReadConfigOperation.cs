namespace Collapsenav.Net.Tool.Excel;

public partial class ReadConfig<T>
{
    /// <summary>
    /// 将表格数据转换为T类型的集合(更快)
    /// </summary>
    public virtual IEnumerable<T> ToEntity<E>(IExcelReader<E> sheet)
    {
        sheet.InitHeader(Range);
        var header = sheet.HeadersWithIndex;
        var rowCount = sheet.RowCount;
        if (Range.StopAt == null)
        {
            foreach (var index in Enumerable.Range(Range.Row + 1, (Range.EndRow ?? rowCount) - 1 - Range.Row))
            {
                var dataRow = sheet[index].ToList();
                // 根据对应传入的设置 为obj赋值
                if (dataRow.NotEmpty())
                {
                    var obj = Activator.CreateInstance<T>();
                    foreach (var option in FieldOption.Where(o => o.ExcelField.IsEmpty() || header.ContainsKey(o.ExcelField)))
                    {
                        if (option.ExcelField.NotNull())
                        {
                            var value = dataRow[header[option.ExcelField]];
                            option.Prop!.SetValue(obj, option.Action == null ? value : option.Action(value?.ToString() ?? string.Empty));
                        }
                        else
                            option.Prop!.SetValue(obj, option.Action == null ? null : option.Action(string.Empty));
                    }
                    Init?.Invoke(obj);
                    yield return obj;
                }
            }
        }
        else
        {
            foreach (var index in Enumerable.Range(Range.Row + 1, rowCount - 1 - Range.Row))
            {
                var dataRow = sheet[index].ToList();
                // 根据对应传入的设置 为obj赋值
                if (dataRow.NotEmpty() && !Range.StopAt((dataRow as IEnumerable<object>)!))
                {
                    var obj = Activator.CreateInstance<T>();
                    foreach (var option in FieldOption)
                    {
                        if (option.ExcelField.NotNull())
                        {
                            var value = dataRow[header[option.ExcelField]];
                            option.Prop!.SetValue(obj, option.Action == null ? value : option.Action(value?.ToString() ?? string.Empty));
                        }
                        else
                            option.Prop!.SetValue(obj, option.Action == null ? null : option.Action(string.Empty));
                    }
                    Init?.Invoke(obj);
                    yield return obj;
                }
                else
                    break;
            }
        }
    }
    /// <summary>
    /// 将表格数据转换为T类型的集合
    /// </summary>
    public virtual async Task<IEnumerable<T>> ToEntityAsync(IExcelReader sheet)
    {
        return await Task.Factory.StartNew(() => ToEntity(sheet).ToList());
        // return await Task.Factory.StartNew(() => IsShuffle ? ToEntity(sheet).Shuffle() : ToEntity(sheet));
    }


    /// <summary>
    /// 转换到实体
    /// </summary>
    public async Task<IEnumerable<T>> ToEntityAsync(ExcelType? excelType = null)
    {
        return await ToEntityAsync(ExcelStream, excelType);
    }
    /// <summary>
    /// 转换到实体
    /// </summary>
    public async Task<IEnumerable<T>> ToEntityAsync(string path, ExcelType? excelType = null)
    {
        using var fs = path.OpenReadShareStream();
        return await ToEntityAsync(fs, excelType);
    }
    /// <summary>
    /// 转换到实体
    /// </summary>
    public virtual async Task<IEnumerable<T>> ToEntityAsync(Stream stream, ExcelType? excelType = null)
    {
        using var reader = ExcelTool.GetExcelReader(stream, excelType);
        return await ToEntityAsync(reader);
    }

    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public IEnumerable<T> ToEntity(ExcelType? excelType = null)
    {
        return ToEntity(ExcelStream, excelType);
    }
    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public IEnumerable<T> ToEntity(string path, ExcelType? excelType = null)
    {
        using var fs = path.OpenReadShareStream();
        return ToEntity(fs, excelType);
    }
    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public virtual IEnumerable<T> ToEntity(Stream stream, ExcelType? excelType = null)
    {
        using var reader = ExcelTool.GetExcelReader(stream, excelType);
        return ToEntity(reader);
    }

    /// <summary>
    /// 转换到实体
    /// </summary>
    public static async Task<IEnumerable<T>> ExcelToEntityAsync(string path)
    {
        return await ExcelTool.ExcelToEntityAsync<T>(path);
    }
    /// <summary>
    /// 转换到实体
    /// </summary>
    public static async Task<IEnumerable<T>> ExcelToEntityAsync(Stream stream)
    {
        return await ExcelTool.ExcelToEntityAsync<T>(stream);
    }
    /// <summary>
    /// 转换到实体
    /// </summary>
    public static async Task<IEnumerable<T>> ExcelToEntityAsync(IExcelReader reader)
    {
        return await ExcelTool.ExcelToEntityAsync<T>(reader);
    }
    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public static IEnumerable<T> ExcelToEntity(string path)
    {
        return ExcelTool.ExcelToEntity<T>(path);
    }
    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public static IEnumerable<T> ExcelToEntity(Stream stream)
    {
        return ExcelTool.ExcelToEntity<T>(stream);
    }
    /// <summary>
    /// 转换到实体(更快)
    /// </summary>
    public static IEnumerable<T> ExcelToEntity(IExcelReader reader)
    {
        return ExcelTool.ExcelToEntity<T>(reader);
    }
}
