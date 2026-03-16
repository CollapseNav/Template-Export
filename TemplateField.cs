using Collapsenav.Net.Tool;
using NPOI.XWPF.UserModel;
/// <summary>
/// 模板字段配置
/// </summary>
public class TemplateField
{
    public TemplateConfig? Config { get; set; }
    public TemplateField(string field, TemplateConfig? config = null)
    {
        Field = field;
        Config = config;
    }
    /// <summary>
    /// 模板字段
    /// </summary>
    public string Field { get; set; }
    /// <summary>
    /// 模板
    /// </summary>
    public string Template
    {
        get
        {
            if (Config == null)
                return field;
            if (Type == TemplateType.Normal)
                return Config.NormalPrefix + Field + Config.Suffix;
            else if (Type == TemplateType.Row)
                return Config.RowPrefix + Field + Config.Suffix;
            else if (Type == TemplateType.Col)
                return Config.ColPrefix + Field + Config.Suffix;
            else
                return field;
        }
    }
    /// <summary>
    /// 模板类型
    /// </summary>
    public TemplateType Type { get; set; } = TemplateType.Normal;
    public XWPFParagraph DocPara { get; set; }
    public XWPFTable? DocTable { get; set; }
    /// <summary>
    /// 是否列表
    /// </summary>
    public bool IsList { get; set; } = false;
    /// <summary>
    /// 表格行位置
    /// </summary>
    public int Row { get; set; } = -1;
    /// <summary>
    /// 表格列位置
    /// </summary>
    public int Col { get; set; } = -1;
    /// <summary>
    /// 获取模板字段
    /// </summary>
    /// <param name="para"></param>
    /// <param name="config"></param>
    /// <returns></returns>
    public static List<TemplateField> GetFields(XWPFParagraph para, TemplateConfig config)
    {
        List<TemplateField> fields = new List<TemplateField>();
        // 后缀没匹配直接返回空集合
        if (!para.Text.Contains(config.Suffix))
            return fields;
        // 匹配到行模板
        if (para.Text.Contains(config.RowPrefix))
        {
            var fs = GetFieldstring(para, config, TemplateType.Row)
                .Select(item => new TemplateField(item, config) { Type = TemplateType.Row, IsList = true, DocPara = para }).ToList();
            fields.AddRange(fs);
        }// 匹配到列模板
        else if (para.Text.Contains(config.ColPrefix))
        {
            var fs = GetFieldstring(para, config, TemplateType.Col)
                .Select(item => new TemplateField(item, config) { Type = TemplateType.Col, IsList = true, DocPara = para }).ToList();
            fields.AddRange(fs);
        }// 匹配到普通模板
        else if (para.Text.Contains(config.NormalPrefix))
        {
            var fs = GetFieldstring(para, config, TemplateType.Normal)
                .Select(item => new TemplateField(item, config) { DocPara = para }).ToList();
            fields.AddRange(fs);
        }
        return fields;
    }

    public static List<TemplateField> GetFields(XWPFDocument doc, TemplateConfig config)
    {
        List<TemplateField> fields = new List<TemplateField>();
        foreach (var para in doc.Paragraphs)
        {
            fields.AddRange(GetFields(para, config));
        }
        foreach (var table in doc.Tables)
        {
            fields.AddRange(GetFields(table, config));
        }
        return fields;
    }

    /// <summary>
    /// 根据表格获取字段
    /// </summary>
    /// <param name="table"></param>
    /// <param name="config"></param>
    /// <returns></returns>
    public static List<TemplateField> GetFields(XWPFTable table, TemplateConfig config)
    {
        List<TemplateField> fields = new List<TemplateField>();
        var rows = table.Rows;
        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var cells = row.GetTableCells();
            for (var j = 0; j < cells.Count; j++)
            {
                var cell = cells[j];
                foreach (var para in cell.Paragraphs)
                {
                    var temp = GetFields(para, config);
                    temp.ForEach(item =>
                    {
                        item.Row = i;
                        item.Col = j;
                        item.DocTable = table;
                    });
                    fields.AddRange(temp);
                }
            }
        }
        return fields;
    }
    /// <summary>
    /// 匹配多个模板字段
    /// </summary>
    /// <param name="para"></param>
    /// <param name="config"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public static IEnumerable<string> GetFieldstring(XWPFParagraph para, TemplateConfig config, TemplateType type)
    {
        var text = para.Text;
        string[]? fields = null;
        switch (type)
        {
            case TemplateType.Normal:
                fields = text.Split(config.NormalPrefix);
                break;
            case TemplateType.Row:
                fields = text.Split(config.RowPrefix);
                break;
            case TemplateType.Col:
                fields = text.Split(config.ColPrefix);
                break;
        }
        if (fields.NotEmpty())
            return fields.Skip(1).Select(field => field.Substring(0, field.IndexOf(config.Suffix))).ToArray();
        else
            return Enumerable.Empty<string>();
    }
}

public enum TemplateType
{
    Normal, Row, Col
}