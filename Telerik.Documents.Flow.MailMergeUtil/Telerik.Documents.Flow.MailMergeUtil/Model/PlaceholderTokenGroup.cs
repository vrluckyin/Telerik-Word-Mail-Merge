using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Telerik.Documents.Flow.MailMergeUtil.Model
{
    //Holds token and its adjacent text  tokens
    //For this phrase => "[City], [State] [Zip]"
    //[City] token will have following properties
    //StartIdentifier = [
    //EndIdentifier = ]
    //Placeholders = [, City, ]
    //TextTokens = ,
    public class PlaceholderTokenGroup
    {
        public string StartIdentifier { get; set; }
        public string EndIdentifier { get; set; }
        public List<PlaceholderToken> Placeholders { get; set; }
        public List<PlaceholderToken> TextTokens { get; set; }
        public string Placeholder
        {
            get
            {
                return String.Join("", Placeholders.Where(w => !w.Text.Equals(StartIdentifier) && !w.Text.Equals(EndIdentifier)).Select(s => s.ToString()));
            }
        }
        public PlaceholderTokenType PlaceholderType { get; set; }
        public bool IsCleanupRequired { get; set; }
        public override string ToString()
        {
            return $"{String.Join("", TextTokens == null ? Placeholders : Placeholders.Union(TextTokens))}";
        }
    }
}
