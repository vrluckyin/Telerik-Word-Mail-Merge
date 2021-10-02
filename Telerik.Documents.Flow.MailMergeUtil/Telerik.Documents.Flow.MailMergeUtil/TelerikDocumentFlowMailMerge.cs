using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Documents.Flow.MailMergeUtil.Tokenizer;
using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
using Telerik.Windows.Documents.Flow.FormatProviders.Pdf;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;

namespace Telerik.Documents.Flow.MailMergeUtil
{
    public class TelerikDocumentFlowMailMerge
    {
        public TelerikDocumentFlowMailMerge()
        {
        }

        //Dataset will be converted as format => <<placeholder>><<index>> => Val0 Val1. 0 and 1 are row indexes
        public async Task<byte[]> MailMerge(byte[] templateBytes, DataSet data, DocumentFormatType documentFormatType = DocumentFormatType.DOCX)
        {
            RadFlowDocument template = ToDocument(templateBytes);
            var mergeFieldProcesses = new List<MergeFieldTokenizerBase>() { new MergeFieldSquareTokenizerProcess(), new MergeFieldTriangleTokenizerProcess() };
            DynamicDataObject templateData = data.ToTemplateData();
            foreach (MergeFieldTokenizerBase mergeFieldProcess in mergeFieldProcesses)
            {
                var runs = template.EnumerateChildrenOfType<Run>().ToList();
                mergeFieldProcess.MailMerge(runs, templateData);
            }

            RadFlowDocument doc = template.MailMerge(new List<DynamicDataObject>() { templateData });

            return ToTargetDocument(doc, documentFormatType);
        }

        private byte[] ToTargetDocument(RadFlowDocument doc, DocumentFormatType documentFormatType)
        {
            switch (documentFormatType)
            {
                case DocumentFormatType.PDF:
                    var pdfProvider = new PdfFormatProvider();
                    using (var output = new MemoryStream())
                    {
                        pdfProvider.Export(doc, output);
                        return output.ToArray();
                    }
                case DocumentFormatType.DOCX:
                    var docxProvider = new DocxFormatProvider();
                    using (var output = new MemoryStream())
                    {
                        docxProvider.Export(doc, output);
                        return output.ToArray();
                    }
                default:
                    return ToBytes(doc);
            }
        }


        public byte[] AddTemplate(byte[] document, byte[] template)
        {
            if (template == null || template.Length == 0)
            {
                // match legacy behavior
                return ToBytes(new RadFlowDocument());
            }

            RadFlowDocument targetDocument = ToDocument(document);
            var editor = new RadFlowDocumentEditor(targetDocument);
            RadFlowDocument sourceDocument = ToDocument(template);
            var importer = new DocumentElementImporter(targetDocument, sourceDocument, ConflictingStylesResolutionMode.UseTargetStyle);
            var inlineBases = sourceDocument.EnumerateChildrenOfType<InlineBase>().ToList();
            foreach (InlineBase run in inlineBases)
            {
                if (run.Parent.Parent is TableCell)
                {
                    continue;
                }
                InlineBase importedRun = importer.Import(run);
                editor.InsertInline(importedRun);
            }

            var tables = sourceDocument.EnumerateChildrenOfType<Table>().ToList();
            foreach (Table table in tables)
            {
                Table insertedTable = editor.InsertTable();
                foreach (TableRow row in table.Rows)
                {
                    TableRow importedRow = importer.Import(row);
                    insertedTable.Rows.Add(importedRow);
                }
            }

            return ToBytes(targetDocument);
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
    }
}
