using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Flow.MailMergeUtil.Model;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

namespace Telerik.Documents.Flow.MailMergeUtil.TokenProcessors
{
    public class ImageTokenProcessor : IMergeTokenProcessor
    {
        public bool Process(PlaceholderTokenGroup token, DynamicDataObject data)
        {
            PlaceholderToken start = token.Placeholders.First();
            var radFlowDocumentEditor = new RadFlowDocumentEditor(start.Parent.Document);
            radFlowDocumentEditor.MoveToInlineStart(start.Parent);
            object value = data.Get(token.Placeholder);
            foreach (string dataKey in token.Placeholder.Split(new char[] { '(', ')', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                value = data.Get(dataKey);
                if (value != null)
                {
                    break;
                }
            }
            //Image placeholder will be Image(200;100):ClientLogo
            //Regex that extraces 200X100 from placeholder
            //WxH format
            Match imgSize = new Regex(@"\(.*?\)").Match(token.Placeholder);
            var sizeParts = imgSize.Value.Split(new char[] { '(', ')', ';', ':', ',' }, StringSplitOptions.RemoveEmptyEntries).ToList().Select(s => { if (Int32.TryParse(s, out int r)) { return r; } return -1; }).Where(w => w > 0).ToList();
            var sizes = Enumerable.Range(0, 2).Select(s => s < sizeParts.Count ? sizeParts[s] : 100).ToList();

            byte[] imageArray = value as byte[];
            string imgExt = "jpg";

            if (value is string imageUrl && Uri.IsWellFormedUriString(imageUrl, UriKind.Absolute))
            {
                imageArray = GetStreamFromUrl(Convert.ToString(value));
            }
            else if (value is string base64Image && TryParseBase64Image(base64Image, out byte[] imageBytes))
            {
                imageArray = imageBytes;
            }

            if (imageArray != null)
            {
                using (Stream stream = new MemoryStream(imageArray))
                {
                    ImageInline img = radFlowDocumentEditor.InsertImageInline(stream, imgExt, new Size(sizes[0], sizes[1]));
                }
            }
            foreach (PlaceholderToken text in token.TextTokens)
            {
                radFlowDocumentEditor.MoveToInlineStart(text.Parent);
                radFlowDocumentEditor.InsertInline(text.Placeholder);
            }
            return true;
        }

        private bool TryParseBase64Image(string base64input, out byte[] imageBytes)
        {
            try
            {
                imageBytes = Convert.FromBase64String(base64input);
                return true;
            }
            catch
            {
                imageBytes = null;
                return false;
            }
        }

        private byte[] GetStreamFromUrl(string url)
        {
            byte[] imageData = null;

            using (var wc = new System.Net.WebClient())
            {
                imageData = wc.DownloadData(url);
            }

            return imageData;
        }
    }
}
