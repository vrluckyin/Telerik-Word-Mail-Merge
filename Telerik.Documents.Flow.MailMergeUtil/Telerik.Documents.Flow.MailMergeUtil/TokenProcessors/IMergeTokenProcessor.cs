using Telerik.Documents.Flow.MailMergeUtil.Model;

namespace Telerik.Documents.Flow.MailMergeUtil.TokenProcessors
{
    public interface IMergeTokenProcessor
    {
        bool Process(PlaceholderTokenGroup token, DynamicDataObject data);
    }
}
