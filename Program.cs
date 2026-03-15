var data = new
{
    Title = "Readme",
    Name = "Template-Export",
    Description = "Excel,Word 的模板导出",
    RowData = new List<object>
    {
        new {Field1="李世民",Field2=0,Field3=DateTime.Now},
        new {Field1="郝大通",Field2=1,Field3=DateTime.Now},
        new {Field1="哑梢公",Field2=2,Field3=DateTime.Now},
    }
};
string docxPath = "./Demo.docx";
string exportPath = "./Export.docx";
string pdfPath = "./Export.pdf";
string mergePath = "./Merge.pdf";
TemplateExport.ExportWordByTemplate(docxPath, exportPath, data);
TemplateExport.DocToPdf(exportPath, pdfPath);
TemplateExport.MergePdf(new List<string> { pdfPath, pdfPath }, mergePath);