using System;
using System.Data;
using System.Linq;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;

namespace Telerik.Documents.Flow.MailMergeUtil.TokenProcessors
{
    /// <summary>
    /// Token Syntax => [html:MergeFieldName]
    /// </summary>
    public class HtmlTokenProcessor : IMergeTokenProcessor
    {
        public bool Process(PlaceholderTokenGroup token, DynamicDataObject data)
        {
            PlaceholderToken placeholder = token.Placeholders.Where(w => !w.Text.Equals(token.StartIdentifier) && !w.Text.Equals(token.EndIdentifier)).FirstOrDefault();
            if (placeholder == null)
            {
                return true;
            }
            var radFlowDocumentEditor = new RadFlowDocumentEditor(placeholder.Parent.Document);

            InlineBase currentNode = placeholder.Parent;
            radFlowDocumentEditor.MoveToInlineStart(currentNode);

            object value = data.Get(token.Placeholder.Trim());
            foreach (string dataKey in token.Placeholder.Trim().Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries))
            {
                value = data.Get(dataKey);
                if (value != null)
                {
                    break;
                }
            }

            if (value != null)
            {
                Telerik.Windows.Documents.Flow.FormatProviders.Html.HtmlFormatProvider provider = new Telerik.Windows.Documents.Flow.FormatProviders.Html.HtmlFormatProvider();
                RadFlowDocument document = provider.Import(value.ToString());

                radFlowDocumentEditor.InsertDocument(document, new InsertDocumentOptions() { ConflictingStylesResolutionMode = ConflictingStylesResolutionMode.UseTargetStyle, InsertLastParagraphMarker = false });
                Run replacement = placeholder.Placeholder.Clone();
                replacement.Text = "";
                currentNode = radFlowDocumentEditor.InsertInline(replacement);
                radFlowDocumentEditor.MoveToInlineEnd(currentNode);
            }

            foreach (PlaceholderToken text in token.TextTokens)
            {
                radFlowDocumentEditor.MoveToInlineStart(text.Parent);
                radFlowDocumentEditor.InsertInline(text.Placeholder);
            }
            return value != null;
        }
    }
}
