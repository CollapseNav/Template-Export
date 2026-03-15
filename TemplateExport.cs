using Collapsenav.Net.Tool;
using NPOI.XWPF.UserModel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
public enum TemplateType
{
    Normal, Row, Col
}
public static class TemplateExport
{
    public static TemplateConfig Config { get; set; } = new TemplateConfig();
    /// <summary>
    /// word模板导出
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="templatePath"></param>
    /// <param name="outputPath"></param>
    /// <param name="data"></param>
    public static void ExportWordByTemplate<T>(string templatePath, string outputPath, T data)
    {
        using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
        {
            XWPFDocument doc = new XWPFDocument(fs);
            foreach (var para in doc.Paragraphs)
            {
                if (para.Text.Contains(Config.Suffix) && para.Text.Contains(Config.NormalPrefix))
                {
                    var fields = GetTemplateFields(para, TemplateType.Normal);
                    foreach (var field in fields)
                    {
                        ReplaceKeyInParagraph(para, data, field, TemplateType.Normal);
                    }
                }
            }
            foreach (var table in doc.Tables)
            {
                // 处理表格动态行的情况
                if (table.Text.Contains(Config.Suffix) && table.Text.Contains(Config.RowPrefix))
                {
                    var fields = GetFields(table, TemplateType.Row);
                    var listField = fields.First().Field.Split('.')[0];
                    List<object> listData = (List<object>)data.GetValue(listField);
                    foreach (var field in fields)
                    {
                        var dataIndex = 0;
                        foreach (var obj in listData)
                        {
                            // 遍历一次数据，则行数+1
                            var cell = table.GetRow(field.Row + dataIndex++).GetCell(field.Col);
                            foreach (var para in cell.Paragraphs)
                            {
                                ReplaceKeyInParagraph(para, obj, field.Field.Replace(listField + ".", ""), TemplateType.Row);
                            }
                        }
                    }
                }
                // 处理表格动态列的情况
                if (table.Text.Contains(Config.Suffix) && table.Text.Contains(Config.ColPrefix))
                {
                    var fields = GetFields(table, TemplateType.Col);
                    var listField = fields.First().Field.Split('.')[0];
                    List<object> listData = (List<object>)data.GetValue(listField);
                    foreach (var field in fields)
                    {
                        var dataIndex = 0;
                        foreach (var obj in listData)
                        {
                            // 遍历一次数据，则列数+1
                            var cell = table.GetRow(field.Row).GetCell(field.Col + dataIndex++);
                            foreach (var para in cell.Paragraphs)
                            {
                                ReplaceKeyInParagraph(para, obj, field.Field.Replace(listField + ".", ""), TemplateType.Row);
                            }
                        }
                    }
                }
                // 处理一般的固定单元格模板
                if (table.Text.Contains(Config.Suffix) && table.Text.Contains(Config.NormalPrefix))
                {
                    var fields = GetFields(table, TemplateType.Normal);
                    foreach (var field in fields)
                    {
                        var cell = table.GetRow(field.Row).GetCell(field.Col);
                        foreach (var para in cell.Paragraphs)
                        {
                            ReplaceKeyInParagraph(para, data, field.Field, TemplateType.Normal);
                        }
                    }
                }
            }
            using (FileStream outFs = new FileStream(outputPath, FileMode.Create))
            {
                doc.Write(outFs);
            }
        }
    }

    /// <summary>
    /// 获取模板的字段
    /// </summary>
    /// <param name="para"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    private static string GetTemplateField(XWPFParagraph para, TemplateType type)
    {
        string field = string.Empty;
        var end = para.Text.IndexOf(Config.Suffix);
        var start = 0;
        switch (type)
        {
            case TemplateType.Normal:
                start = para.Text.IndexOf(Config.NormalPrefix);
                field = para.Text.Substring(start + Config.NormalPrefix.Length, end - start - Config.NormalPrefix.Length);
                break;
            case TemplateType.Row:
                start = para.Text.IndexOf(Config.RowPrefix);
                field = para.Text.Substring(start + Config.RowPrefix.Length, end - start - Config.RowPrefix.Length);
                break;
            case TemplateType.Col:
                start = para.Text.IndexOf(Config.ColPrefix);
                field = para.Text.Substring(start + Config.ColPrefix.Length, end - start - Config.ColPrefix.Length);
                break;
        }
        return field;
    }
    /// <summary>
    /// 获取模板的字段
    /// </summary>
    /// <param name="para"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    private static List<string> GetTemplateFields(XWPFParagraph para, TemplateType type)
    {
        var text = para.Text;
        List<string> fields = new List<string>();
        switch (type)
        {
            case TemplateType.Normal:
                fields = text.Split(Config.NormalPrefix).ToList();
                break;
            case TemplateType.Row:
                fields = text.Split(Config.RowPrefix).ToList();
                break;
            case TemplateType.Col:
                fields = text.Split(Config.ColPrefix).ToList();
                break;
        }
        fields = fields.Skip(1).Select(field => field.Substring(0, field.IndexOf(Config.Suffix))).ToList();
        return fields;
    }

    /// <summary>
    /// 获取所有模板以及模板所在的行列信息
    /// </summary>
    /// <param name="table"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    private static List<TableField> GetFields(XWPFTable table, TemplateType type)
    {
        List<TableField> fields = new List<TableField>();
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
                    if (!para.Text.Contains(Config.Suffix))
                        continue;
                    switch (type)
                    {
                        case TemplateType.Normal:
                            if (para.Text.Contains(Config.NormalPrefix) && !para.Text.Contains(Config.RowPrefix) && !para.Text.Contains(Config.ColPrefix))
                            {
                                var fs = GetTemplateFields(para, TemplateType.Normal).Select(item => new TableField() { Field = item }).ToList();
                                fs.ForEach(field =>
                                {
                                    field.Row = i;
                                    field.Col = j;
                                });
                                fields.AddRange(fs);
                            }
                            break;
                        case TemplateType.Row:
                            if (para.Text.Contains(Config.RowPrefix))
                            {
                                var fs = GetTemplateFields(para, TemplateType.Row).Select(item => new TableField() { Field = item }).ToList();
                                fs.ForEach(field =>
                                {
                                    field.Row = i;
                                    field.Col = j;
                                });
                                fields.AddRange(fs);
                            }
                            break;
                        case TemplateType.Col:
                            if (para.Text.Contains(Config.ColPrefix))
                            {
                                var fs = GetTemplateFields(para, TemplateType.Col).Select(item => new TableField() { Field = item }).ToList();
                                fs.ForEach(field =>
                                {
                                    field.Row = i;
                                    field.Col = j;
                                });
                                fields.AddRange(fs);
                            }
                            break;
                    }
                }
            }
        }
        return fields;
    }

    /// <summary>
    /// 模板替换
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="para"></param>
    /// <param name="data"></param>
    /// <param name="field"></param>
    /// <param name="type"></param>
    private static void ReplaceKeyInParagraph<T>(XWPFParagraph para, T data, string field, TemplateType type)
    {
        var temp = string.Empty;
        var start = para.Text.IndexOf(Config.NormalPrefix);
        var end = para.Text.IndexOf(Config.Suffix);
        if (start >= 0)
            temp = para.Text.Substring(start, end - start) + Config.Suffix;
        var tempRun = para.Runs.Count > 0 ? para.Runs.FirstOrDefault(p => p.Text.Contains(temp)) : null;
        var value = data.GetValue(field);
        string text;
        if (temp.IsEmpty() && type != TemplateType.Normal)
        {
            text = value.ToString();
        }
        else if (temp.IsEmpty() && type == TemplateType.Normal)
        {
            return;
        }
        else
        {
            if (tempRun == null)
                text = para.Text.Replace(temp, value.ToString());
            else
                text = tempRun.Text.Replace(temp, value.ToString());
        }

        if (tempRun != null)
        {
            tempRun.SetText(text);
        }
        else
        {
            tempRun = para.Runs.Count > 0 ? para.Runs[0] : null;
            for (int i = para.Runs.Count - 1; i >= 0; i--)
            {
                para.RemoveRun(i);
            }
            var newRun = para.CreateRun();
            newRun.SetText(text);
            if (tempRun != null)
            {
                newRun.FontFamily = tempRun.FontFamily;
                newRun.FontSize = tempRun.FontSize;
                newRun.IsBold = tempRun.IsBold;
                newRun.IsItalic = tempRun.IsItalic;
                newRun.SetColor(tempRun.GetColor());
                newRun.Underline = tempRun.Underline;
            }
        }
    }
    /// <summary>
    /// 合并多个pdf
    /// </summary>
    /// <param name="paths"></param>
    /// <param name="outPath"></param>
    public static void MergePdf(IEnumerable<string> paths, string outPath)
    {
        using (PdfDocument targetDoc = new PdfDocument())
        {
            foreach (string file in paths)
            {
                // 以导入模式打开源文档
                using (PdfDocument sourceDoc = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                {
                    // 遍历并将每一页加入目标文档
                    foreach (PdfPage page in sourceDoc.Pages)
                    {
                        targetDoc.AddPage(page);
                    }
                }
            }
            targetDoc.Save(outPath);
        }
    }
    /// <summary>
    /// word转为pdf
    /// </summary>
    /// <param name="docpath"></param>
    /// <param name="pdfpath"></param>
    public static void DocToPdf(string docpath, string pdfpath)
    {
        Spire.Doc.Document document = new Spire.Doc.Document(docpath);
        document.SaveToFile(pdfpath, Spire.Doc.FileFormat.PDF);
    }

    /// <summary>
    /// excel转为pdf
    /// </summary>
    /// <param name="docpath"></param>
    /// <param name="pdfpath"></param>
    public static void ExcelToPdf(string excelpath, string pdfpath)
    {
        Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
        workbook.LoadFromFile(excelpath);
        workbook.SaveToFile(pdfpath, Spire.Xls.FileFormat.PDF);
    }
}