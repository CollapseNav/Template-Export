using Collapsenav.Net.Tool;
using NPOI.XWPF.UserModel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
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
            // 先拿到所有模板
            var docFields = TemplateField.GetFields(doc, Config);
            // 第一步先处理段落中的模板
            var paraFields = docFields.Where(item => item.IsList == false).ToList();
            foreach (var para in doc.Paragraphs)
            {
                var fields = paraFields.Where(item => item.DocPara == para);
                foreach (var field in fields)
                {
                    ReplaceKeyInParagraph(para, data, field.Field, TemplateType.Normal);
                }
            }
            // 然后处理表格中的模板
            var tableFields = docFields.Where(item => item.IsList == true).ToList();
            // 考虑可能会有多个列表的情况
            var tableFieldGroup = tableFields.GroupBy(item => item.Field.Split('.')[0]).ToList();
            foreach (var table in doc.Tables)
            {
                // 多个列表数据
                foreach (var fieldGroup in tableFieldGroup)
                {
                    // 先获取列表数据
                    var listData = (IEnumerable<object>?)data.GetValue(fieldGroup.Key);
                    if (listData == null)
                        continue;
                    // 然后遍历模板
                    var fields = fieldGroup.Where(item => item.DocTable == table).ToList();
                    foreach (var field in fields)
                    {
                        var dataIndex = 0;
                        foreach (var obj in listData)
                        {
                            // 遍历一次数据，根据模板类型，获取单元格
                            // 如果是行模式，则列数不变，行数递增
                            // 如果是列模式，则行数不变，列数递增
                            var cell = table.GetRow(field.Row + (field.Type == TemplateType.Row ? dataIndex++ : 0))
                                            .GetCell(field.Col + (field.Type == TemplateType.Col ? dataIndex++ : 0));
                            foreach (var para in cell.Paragraphs)
                            {
                                ReplaceKeyInParagraph(para, obj, field.Field.Replace(fieldGroup.Key + ".", ""), field.Type);
                            }
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