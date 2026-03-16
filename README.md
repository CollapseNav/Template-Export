# Template-Export

通过NPOI实现简单的模板导出功能

默认`{{}}`用于标记

`{{r-**}}`标记以行的形式展开

`{{c-**}}`标记以列的形式展开

都可以通过 `TemplateConfig` 进行自定义, 三种定义最好不要完全相同

`TemplateExport.ExportWordByTemplate` 依据模板导出Word文档

`TemplateExport.DocToPdf` 转换Word文档为PDF文件, 通过 **FreeSpire** 实现

`TemplateExport.MergePdf` 合并多个PDF文件, 通过 **PdfSharp** 实现

## 示例

```csharp
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
```
