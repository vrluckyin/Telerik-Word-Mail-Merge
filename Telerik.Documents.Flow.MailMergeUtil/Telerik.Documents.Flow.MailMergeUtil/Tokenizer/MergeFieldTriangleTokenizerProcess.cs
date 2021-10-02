using System.Collections.Generic;

namespace Telerik.Documents.Flow.MailMergeUtil.Tokenizer
{
    public class MergeFieldTriangleTokenizerProcess : MergeFieldTokenizerBase
    {
        public const string START_IDENTIFIER_TRIANGLE = "<<";
        public const string END_IDENTIFIER_TRIANGLE = ">>";

        public MergeFieldTriangleTokenizerProcess() : base(START_IDENTIFIER_TRIANGLE, END_IDENTIFIER_TRIANGLE)
        {

        }

        public override string DataKey(List<string> tokens)
        {
            return tokens[tokens.Count - 1];
        }
    }
}
