using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

namespace Telerik.Documents.Flow.MailMergeUtil.Tests
{
    [TestClass]
    public class TelerikDocumentFlowMailMergeTests
    {
        private TelerikDocumentFlowMailMerge _documentProcessor = null;

        [TestInitialize]
        public void Setup()
        {
            _documentProcessor = new TelerikDocumentFlowMailMerge();
        }

        [TestMethod]
        public void MergeDocumentsTest()
        {
            var target = new RadFlowDocument();
            var source = new RadFlowDocument();
            Paragraph targetParagraph = source.Sections.AddSection().Blocks.AddParagraph();
            targetParagraph.Inlines.AddRun("DOCUMENT_PROCESSOR");
            targetParagraph.Inlines.AddRun("Telerik");

            Paragraph sourceParagraph = source.Sections.AddSection().Blocks.AddParagraph();
            sourceParagraph.Inlines.AddRun("TEST");
            sourceParagraph.Inlines.AddRun("Merge Documents");
            // target will contain merged content and styles. 
            target.Merge(source);

            var runs = target.EnumerateChildrenOfType<Run>().ToList();
            Assert.IsTrue(runs.Count == 4);
        }

        [TestMethod]
        public async Task MailMergeTextTest()
        {
            (byte[] Template, DataSet Model) = GetDocumentHavingJustPlaceholders();

            byte[] result = await _documentProcessor.MailMerge(Template, Model);

            RadFlowDocument document = ToDocument(result);
            var runs = document.EnumerateChildrenOfType<Run>().ToList();

            Assert.AreEqual(runs[0].ToString().Trim(), Model.Tables["Table1"].Rows[0]["Name"].ToString());
            Assert.AreEqual(runs[1].ToString().Trim(), Model.Tables["Table1"].Rows[0]["Address"].ToString());
        }

        [TestMethod]
        public async Task MailMergeImageTest()
        {
            (byte[] Template, DataSet Model) = GetDocumentHavingJustImages();

            byte[] result = await _documentProcessor.MailMerge(Template, Model);

            RadFlowDocument document = ToDocument(result);
            ImageInline image = document.EnumerateChildrenOfType<ImageInline>().FirstOrDefault();

            Assert.IsNotNull(image);
        }

        [TestMethod]
        public async Task MailMergeMultipleImageTest()
        {
            (byte[] Template, DataSet Model) = GetDocumentWithTwoImages();
            Model.Tables["MergeImagesTable"].Columns.Add("FieldName", typeof(string));
            Model.Tables["MergeImagesTable"].Columns.Add("FieldData", typeof(byte[]));
            Model.Tables["MergeImagesTable"].Rows.Add("ClientSignatureImage", File.ReadAllBytes(@"DocumentProcessorsTest\Telerik\logo.png"));
            
            byte[] result = await _documentProcessor.MailMerge(Template, Model);

            RadFlowDocument document = ToDocument(result);
            Assert.AreEqual(2, document.EnumerateChildrenOfType<ImageInline>().Count());
        }

        [TestMethod]
        public async Task MailMergeTableWith1RowNColumnsTest()
        {
            (byte[] Template, DataSet Model) = GetDocumentHavingTableWith1RowNColumns();

            byte[] result = await _documentProcessor.MailMerge(Template, Model);

            RadFlowDocument document = ToDocument(result);
            var runs = document.EnumerateChildrenOfType<Run>().ToList();
            Table generatedTable = document.EnumerateChildrenOfType<Table>().FirstOrDefault();

            Assert.IsNotNull(generatedTable);
            Assert.AreEqual(generatedTable.Rows.Count, Model.Tables["Table1"].Rows.Count);
            Assert.AreEqual(generatedTable.Rows[0].Cells.Count, Model.Tables["Table1"].Columns.Count);
        }

        [TestMethod]
        public async Task MailMergeTableWithNRows1ColumnsTest()
        {
            (byte[] Template, DataSet Model) = GetDocumentHavingTableWithNRows1ColumnsTest();

            byte[] result = await _documentProcessor.MailMerge(Template, Model);

            RadFlowDocument document = ToDocument(result);
            var runs = document.EnumerateChildrenOfType<Run>().ToList();
            Table generatedTable = document.EnumerateChildrenOfType<Table>().FirstOrDefault();

            Assert.IsNotNull(generatedTable);
        }

        [TestMethod]
        public async Task MailMergeWithHeaderTest()
        {
            int width = 100;
            int height = 100;
            var header = new RadFlowDocument();
            var editorHeader = new RadFlowDocumentEditor(header);
            editorHeader.InsertText($"[Image({width};{height}):ClientLogo]");
            editorHeader.InsertText("[ClientId]");
            editorHeader.InsertText("[ClientName]");
            editorHeader.InsertText("[ClientAddress]");

            var template = new RadFlowDocument();
            var editorTemplate = new RadFlowDocumentEditor(template);
            editorTemplate.InsertText("[Name]");
            editorTemplate.InsertText("[Address]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1':[
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                }
                              ]
                            }";

            var infoData = new Dictionary<string, object>();

            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);

            byte[] mergedDocument = _documentProcessor.AddTemplate(ToBytes(template), ToBytes(header));
            // target will contain merged content and styles. 
            byte[] result = await _documentProcessor.MailMerge(mergedDocument, dataSet);

            RadFlowDocument document = ToDocument(result);
            var runs = document.EnumerateChildrenOfType<Run>().ToList();
            var images = document.EnumerateChildrenOfType<ImageInline>().ToList();

            Assert.IsNotNull(runs);
            Assert.IsNotNull(images);
        }

        private RadFlowDocument ToDocument(byte[] data)
        {
            var provider = new DocxFormatProvider();
            return provider.Import(data);
        }

        private byte[] ToBytes(RadFlowDocument doc)
        {
            var provider = new DocxFormatProvider();
            return provider.Export(doc);
        }

        private (byte[] Template, DataSet Model) GetDocumentHavingJustPlaceholders()
        {
            var template = new RadFlowDocument();
            Paragraph sourceParagraph = template.Sections.AddSection().Blocks.AddParagraph();
            sourceParagraph.Inlines.AddRun("[Name]");
            sourceParagraph.Inlines.AddRun("[Address]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1': [
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                }
                              ]
                            }";

            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);

            return (Template: ToBytes(template), Model: dataSet);
        }

        private (byte[] Template, DataSet Model) GetDocumentHavingJustImages()
        {
            var template = new RadFlowDocument();
            Paragraph sourceParagraph = template.Sections.AddSection().Blocks.AddParagraph();
            sourceParagraph.Inlines.AddRun($"[Image(100;100):ClientLogo]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1': [
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                }
                              ]
                            }";
            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);
            return (Template: ToBytes(template), Model: dataSet);
        }

        private (byte[] Template, DataSet Model) GetDocumentWithTwoImages()
        {
            var template = new RadFlowDocument();
            Paragraph sourceParagraph = template.Sections.AddSection().Blocks.AddParagraph();
            sourceParagraph.Inlines.AddRun($"[Image(100;100):ClientLogo]");
            sourceParagraph.Inlines.AddRun($"[Image(100;100):ClientSignatureImage]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1': [
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                }
                              ],
                              'MergeImagesTable': []
                            }";
            // FieldData is set by the test method
            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);
            return (Template: ToBytes(template), Model: dataSet);
        }

        private (byte[] Template, DataSet Model) GetDocumentHavingTableWith1RowNColumns()
        {
            var template = new RadFlowDocument();
            var editor = new RadFlowDocumentEditor(template);
            editor.InsertText("[TableStart:Table1]");
            Table table = editor.InsertTable();
            TableRow row = table.Rows.AddTableRow();
            row.Cells.AddTableCell().Blocks.AddParagraph().Inlines.AddRun("[Name]");
            row.Cells.AddTableCell().Blocks.AddParagraph().Inlines.AddRun("[Address]");
            editor.InsertText("[TableEnd:Table1]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1': [
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                },
                                {
                                  'Name': 1,
                                  'Address': 'item 1'
                                }
                              ]
                            }";

            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);
            return (Template: ToBytes(template), Model: dataSet);
        }

        private (byte[] Template, DataSet Model) GetDocumentHavingTableWithNRows1ColumnsTest()
        {
            var template = new RadFlowDocument();
            var editor = new RadFlowDocumentEditor(template);
            editor.InsertText("[TableStart:Table1]");
            Table table = editor.InsertTable();
            TableRow row1 = table.Rows.AddTableRow();
            row1.Cells.AddTableCell().Blocks.AddParagraph().Inlines.AddRun("[Name]");
            TableRow row2 = table.Rows.AddTableRow();
            row2.Cells.AddTableCell().Blocks.AddParagraph().Inlines.AddRun("[Address]");
            editor.InsertText("[TableEnd:Table1]");

            string json = @"{
                            'HeaderData': [
                                {
                                  'ClientId': 80,
                                  'ClientName': 'OE324',
                                  'ClientAddress': 'USA'
                                }
                              ],
                              'Table1': [
                                {
                                  'Name': 0,
                                  'Address': 'item 0'
                                }
                              ]
                            }";
            DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(json);
            return (Template: ToBytes(template), Model: dataSet);
        }
    }
}
