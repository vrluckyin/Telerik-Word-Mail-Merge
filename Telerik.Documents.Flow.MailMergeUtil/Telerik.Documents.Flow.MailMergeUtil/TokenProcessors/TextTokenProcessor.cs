using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Documents.Media;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

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
