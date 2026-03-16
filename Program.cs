var data = new
{
    Title = "Readme",
    Name = "Template-Export",
    Description = "Excel,Word 的模板导出",
    RowData = new object[]
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
// 根据模板导出
TemplateExport.ExportWordByTemplate(docxPath, exportPath, data);
// 将导出的文件转为PDF
TemplateExport.DocToPdf(exportPath, pdfPath);
// 合并PDF文件
TemplateExport.MergePdf(new List<string> { pdfPath, pdfPath }, mergePath);