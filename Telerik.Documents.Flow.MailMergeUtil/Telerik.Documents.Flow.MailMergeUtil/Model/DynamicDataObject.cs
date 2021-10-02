using System.Collections.Generic;
using System.Dynamic;

namespace Telerik.Documents.Flow.MailMergeUtil.Model
{
    public class DynamicDataObject : DynamicObject
    {
        private readonly Dictionary<string, object> dictionary = new Dictionary<string, object>();

        public Dictionary<string, object> DataDictionary
        {
            get
            {
                return dictionary;
            }
        }

        public int Count
        {
            get
            {
                return dictionary.Count;
            }
        }

        public IEnumerable<string> GetColumnNames()
        {
            return dictionary.Keys;
        }

        public void Set(string propertyName, object value)
        {
            TrySetMember(new DynamicObjectSetMemberBinder(propertyName, false), value);
        }

        public object Get(string propertyName)
        {
            TryGetMember(new DynamicObjectGetMemberBinder(propertyName, false), out object result);
            return result;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            string propertyName = binder.Name;
            return dictionary.TryGetValue(propertyName, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            string propertyName = binder.Name;
            dictionary[propertyName] = value;
            return true;
        }
    }
}
