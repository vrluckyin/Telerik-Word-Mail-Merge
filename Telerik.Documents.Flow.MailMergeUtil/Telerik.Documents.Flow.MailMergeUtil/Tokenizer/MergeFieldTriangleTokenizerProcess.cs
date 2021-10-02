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
