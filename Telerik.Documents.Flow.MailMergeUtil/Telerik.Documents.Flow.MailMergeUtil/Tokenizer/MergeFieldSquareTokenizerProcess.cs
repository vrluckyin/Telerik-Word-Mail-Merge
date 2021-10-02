using System.Collections.Generic;

namespace Telerik.Documents.Flow.MailMergeUtil.Tokenizer
{
    public class MergeFieldSquareTokenizerProcess : MergeFieldTokenizerBase
    {
        public const string START_IDENTIFIER_SQUARE = "[";
        public const string END_IDENTIFIER_SQUARE = "]";

        public MergeFieldSquareTokenizerProcess() : base(START_IDENTIFIER_SQUARE, END_IDENTIFIER_SQUARE)
        {

        }

        public override string DataKey(List<string> tokens)
        {
            return tokens[tokens.Count - 1];
        }
    }
}
