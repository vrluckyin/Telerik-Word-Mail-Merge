using System.Dynamic;

namespace Telerik.Documents.Flow.MailMergeUtil.Model
{
    public class DynamicObjectSetMemberBinder : SetMemberBinder
    {
        public DynamicObjectSetMemberBinder(string name, bool ignoreCase)
            : base(name, ignoreCase)
        {
        }

        public override DynamicMetaObject FallbackSetMember(DynamicMetaObject target, DynamicMetaObject value, DynamicMetaObject errorSuggestion)
        {
            return target;
        }
    }

    public class DynamicObjectGetMemberBinder : GetMemberBinder
    {
        public DynamicObjectGetMemberBinder(string name, bool ignoreCase)
            : base(name, ignoreCase)
        {
        }

        public override DynamicMetaObject FallbackGetMember(DynamicMetaObject target, DynamicMetaObject errorSuggestion)
        {
            return target;
        }
    }
}
