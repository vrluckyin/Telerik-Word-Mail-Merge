using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Media;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

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
