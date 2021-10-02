using System.Text;

namespace Telerik.Documents.Flow.MailMergeUtil
{
    public class TemplatePlaceHolder
    {
        public string PlaceHolderKey { get; set; }
        public byte[] PlaceHolderContent { get; set; }
        public bool IsTemplate { get; set; }

        public string PlaceHolderContentAsString { get { return UTF8Encoding.UTF8.GetString(PlaceHolderContent); } }
    }
}