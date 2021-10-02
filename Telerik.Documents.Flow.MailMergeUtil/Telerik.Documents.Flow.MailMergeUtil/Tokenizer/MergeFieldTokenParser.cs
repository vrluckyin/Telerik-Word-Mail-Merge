using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Telerik.Documents.Flow.MailMergeUtil.Model;

namespace Telerik.Documents.Flow.MailMergeUtil.Tokenizer
{
    /*
       1) Collect tokens - [], <<>>
       2) Identify token type - Text, Image or Table
       3) For text, just replace placeholder with actual data
       4) For image, add an image based on given source - byte[] or image url. Set width/height if given. Default size is 100x100
       5) For table, repeat steps 1 to 4 as table is container and may have multiple place holders.
       6) At last apply, telerik merge

       Telerik limitations:
       * Image does not support - [Image(W;H):PropertyName]
       * Table does not support that generate multiple records.
           [TableStart:TableName]
               [Name], [Address]
           [TableEnd:TableName]
       * Telerik does not support - merging partial documents like we have common header


        */
    public class MergeFieldTokenParser
    {
        private readonly string _startIdentifier = null;
        private readonly string _endIdentifier = null;
        public PlaceholderTokenGroup Data { get; set; }

        public MergeFieldTokenParser(string startIdentifier, string endIdentifier)
        {
            _startIdentifier = startIdentifier;
            _endIdentifier = endIdentifier;

            Data = new PlaceholderTokenGroup();
            Data.StartIdentifier = _startIdentifier;
            Data.EndIdentifier = _endIdentifier;
        }

        private PlaceholderToken FindStart(List<PlaceholderToken> runs, int index)
        {
            bool currentElementHasIdentifier = false;
            for (; index >= 0 && currentElementHasIdentifier == false; index--)
            {
                currentElementHasIdentifier = runs[index].Text.Contains(_startIdentifier);
            }
            return runs[index + 1];
        }

        private PlaceholderToken FindEnd(List<PlaceholderToken> runs, int index)
        {
            bool currentElementHasIdentifier = false;
            for (; index < runs.Count && currentElementHasIdentifier == false; index++)
            {
                currentElementHasIdentifier = runs[index].Text.Contains(_endIdentifier);
            }
            return runs[index - 1];
        }

        private PlaceholderToken FindEndTillStartIdentifierFound(List<PlaceholderToken> runs, int index)
        {
            PlaceholderToken endRun = null;
            for (; index < runs.Count; index++)
            {
                if (runs[index].Text.IndexOf(_startIdentifier) >= 0)
                {
                    endRun = runs[index - 1]; //return Run before next StartIdentifier that will be processed on next token
                    break;
                }
            }
            return endRun ?? runs.Last();
        }

        public void Parse(List<PlaceholderToken> runs, int index)
        {
            PlaceholderToken startRun = FindStart(runs, index);
            int placeholderStartIndex = runs.IndexOf(startRun);

            PlaceholderToken endRun = FindEnd(runs, placeholderStartIndex);
            int placeholderEndIndex = runs.IndexOf(endRun);

            //Telerik sometime returns place holder value in this way
            //Date Sent: [DateSent] and Contribution Date is [ContributionDate]
            //Runs will be
            //Date Sent: [
            //D
            //ateSent]
            //[ContributionDate] <==uptoNextStartIdenfierRun
            PlaceholderToken uptoNextStartIdenfierRun = FindEndTillStartIdentifierFound(runs, placeholderEndIndex);
            int uptoNextStartIdenfierRunIndex = runs.IndexOf(uptoNextStartIdenfierRun);

            string startRunText = startRun.ToString();
            string endRunText = Convert.ToString(endRun);

            Data.Placeholders = runs.Where((w, i) => i >= placeholderStartIndex && i <= placeholderEndIndex).ToList();
            Data.TextTokens = runs.Where((w, i) => i > placeholderEndIndex && i <= uptoNextStartIdenfierRunIndex).ToList();
            Data.PlaceholderType = IdentifyTokenType(Data.Placeholder);
        }

        private PlaceholderTokenType IdentifyTokenType(string placeholder)
        {
            //currently, there are only three types of token image and text.
            //table token will be separately after parsing all token as table may have multiple placeholders of image and other merge fields
            if (placeholder.ToLower().Contains("image("))
            {
                return PlaceholderTokenType.Image;
            }
            else if (placeholder.ToLower().StartsWith("html:"))
            {
                return PlaceholderTokenType.Html;
            }
            return PlaceholderTokenType.Text;
        }

        public override string ToString()
        {
            return Data.ToString();
        }
    }
}
