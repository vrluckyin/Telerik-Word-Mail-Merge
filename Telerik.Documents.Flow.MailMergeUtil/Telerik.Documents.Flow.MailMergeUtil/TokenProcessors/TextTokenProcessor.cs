using System.Data;
using System.Linq;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;

namespace Telerik.Documents.Flow.MailMergeUtil.TokenProcessors
{
    public class TextTokenProcessor : IMergeTokenProcessor
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
            if (value != null)
            {
                string replacementText = value.ToString();
                Run replacement = placeholder.Placeholder.Clone();
                replacement.Text = replacementText;
                currentNode = radFlowDocumentEditor.InsertInline(replacement);
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
