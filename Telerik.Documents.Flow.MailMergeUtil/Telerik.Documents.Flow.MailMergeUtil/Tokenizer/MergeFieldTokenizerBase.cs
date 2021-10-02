using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Documents.Flow.MailMergeUtil.TokenProcessors;
using Telerik.Documents.Media;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

namespace Telerik.Documents.Flow.MailMergeUtil.Tokenizer
{
    /*
       Base class for all custom token filter
        */
    public abstract class MergeFieldTokenizerBase
    {
        private readonly string _startIdentifier = null;
        private readonly string _endIdentifier = null;
        private readonly Dictionary<PlaceholderTokenType, IMergeTokenProcessor> _replacer = new Dictionary<PlaceholderTokenType, IMergeTokenProcessor>()
        {
            {   PlaceholderTokenType.Image, new ImageTokenProcessor() },
            {   PlaceholderTokenType.Html, new HtmlTokenProcessor() },
            {   PlaceholderTokenType.Text, new TextTokenProcessor() },
            {   PlaceholderTokenType.Table, new TableTokenProcessor() },
        };

        public MergeFieldTokenizerBase(string startIdentifier, string endIdentifier)
        {
            _startIdentifier = startIdentifier;
            _endIdentifier = endIdentifier;
        }

        public abstract string DataKey(List<string> tokens);

        public virtual List<MergeFieldTokenParser> IdenfityTokens(List<Run> identifiedRuns)
        {
            //List<Run> cleanupTokens = ReplaceTable(identifiedRuns);
            List<PlaceholderToken> runs = SplitTokens(identifiedRuns);

            //Sometime template may have this field : Date: [BirthDate]
            //So Telerik generates 3 Run objects, 1) "Date: [", 2) "BirthDate", 3) "]"
            //that's why we checked start and ends with start idenfier
            var tokens = runs.Select((s, i) => new { Item = s, Index = i }).Where((w, i) => w.Item.Placeholder.Text.Contains(_startIdentifier)).ToList();

            //for each starting token, find tokens that falls between start and end identifier
            var mergeFields = tokens.Select((s, i) =>
            {
                var mergeField = new MergeFieldTokenParser(_startIdentifier, _endIdentifier);
                mergeField.Parse(runs, s.Index);
                return mergeField;
            }).ToList();

            //identify tables after parsing all tokens - image and text placeholders
            mergeFields = IdentifyTables(mergeFields);
            return mergeFields;
        }

        //Token may be detected as per below
        //[BirthDate] => will be converted into 3 separate tokens
        //===> [, BirthDate, ]
        //][  => will be converted into 2 separate tokens
        //===> ], [
        //] testing this is the [Name => will be converted into 4 separate tokens
        //===> ], testing this is the, [, Name
        public List<PlaceholderToken> SplitTokens(List<Run> runs)
        {
            var result = new List<PlaceholderToken>();
            foreach (Run run in runs)
            {
                //No Trim because need to maintain space
                if (run.Text.Equals(_startIdentifier) || run.Text.Equals(_endIdentifier))
                {
                    result.Add(new PlaceholderToken(run, run.Text));
                }
                else
                {
                    //consider this kind of token => 
                    //1) [City], [State
                    //2) ] This is testing [Country]
                    string runText = run.Text;
                    string startIdentifierReplacer = "~`";
                    string endIdentifierReplacer = "`~";
                    runText = runText.Replace(_startIdentifier, $"{_startIdentifier}{startIdentifierReplacer}");
                    runText = runText.Replace(_endIdentifier, $"{_endIdentifier}{endIdentifierReplacer}");
                    //consider this kind of token =>
                    //1) //[~`City`~], [~`State
                    //2) `~] This is testing [~`Country`~]
                    var splitTokenRuns = runText.Split(new char[] { _startIdentifier[0], _endIdentifier[0] }, StringSplitOptions.RemoveEmptyEntries).Select((s, i) =>
                    {
                        string splitRunText = s.Replace(startIdentifierReplacer, _startIdentifier).Replace(endIdentifierReplacer, _endIdentifier);
                        return splitRunText;
                    }).ToList();

                    //To maintain the sequence of identifier at what index it will be added
                    //it should not be just adding post fix and prefix of star/end token => [ + token + ]
                    //For 2a) it will be wrongly evaluated
                    //1a)[City],
                    //1b)[State
                    //2a)] This is testing [Country
                    //2b)]
                    foreach (string splitToken in splitTokenRuns)
                    {
                        var newRuns = splitToken.Split(new char[] { _startIdentifier[0], _endIdentifier[0] }, StringSplitOptions.RemoveEmptyEntries).Select((s, i) =>
                        {
                            return new { Index = splitToken.IndexOf(s), TextNode = new PlaceholderToken(run, s) };
                        }).ToDictionary(k => k.Index, v => v.TextNode);

                        int startIdentifierIndex = splitToken.IndexOf(_startIdentifier);
                        int endIdentifierIndex = splitToken.IndexOf(_endIdentifier);
                        PlaceholderToken startIdentifierRun = startIdentifierIndex >= 0 ? new PlaceholderToken(run, _startIdentifier) : null;
                        PlaceholderToken endIdentifierRun = endIdentifierIndex >= 0 ? new PlaceholderToken(run, _endIdentifier) : null;

                        if (startIdentifierIndex >= 0)
                        {
                            newRuns.Add(startIdentifierIndex, startIdentifierRun);
                        }
                        if (endIdentifierIndex >= 0)
                        {
                            newRuns.Add(endIdentifierIndex, endIdentifierRun);
                        }

                        result.AddRange(newRuns.OrderBy(o => o.Key).Select(s => s.Value).ToList());
                    }
                }
            }
            return result;
        }

        //currently support table that generate multiple rows for [TableStart:<<tableName>>] ....table content ... [TableEnd:<<tableName>>]
        public virtual List<MergeFieldTokenParser> IdentifyTables(List<MergeFieldTokenParser> tokens)
        {
            var result = new List<MergeFieldTokenParser>();
            for (int i = 0; i < tokens.Count;)
            {
                MergeFieldTokenParser token = tokens[i];
                //contains TableStart:Funds ......... TableEnd:Funds
                if (token.Data.Placeholder.ToLower().StartsWith("tablestart:"))
                {
                    string tableEndTokenText = token.Data.Placeholder.ToLower().Replace("tablestart:", "tableend:");
                    var table = new MergeFieldTokenParser("tablestart:", "tableend:");
                    table.Data.Placeholders = new List<PlaceholderToken>();
                    table.Data.PlaceholderType = PlaceholderTokenType.Table;
                    int startIndex = tokens.IndexOf(token);
                    MergeFieldTokenParser tableEndToken = tokens.Skip(startIndex).FirstOrDefault(w => w.Data.Placeholder.ToLower().StartsWith(tableEndTokenText));
                    if (tableEndToken != null)
                    {
                        int endIndex = tokens.IndexOf(tableEndToken);
                        var tableTokens = tokens.Skip(startIndex).Take(endIndex - startIndex + 1).ToList();
                        var tt = tableTokens.SelectMany(s => s.Data.Placeholders).ToList();
                        table.Data.Placeholders.AddRange(tt);
                        result.Add(table);
                        i += tableTokens.Count;
                    }
                    else
                    {
                        i++;
                        continue;
                    }

                }
                else if (token.Data.Placeholder.ToLower().StartsWith("formatphone"))
                {
                    token.Data.StartIdentifier = "formatphone";
                    result.Add(token);
                    i++;
                }
                else
                {
                    result.Add(token);
                    i++;
                }
            }
            return result;
        }

        public virtual void MailMerge(List<Run> runs, DynamicDataObject data)
        {
            List<MergeFieldTokenParser> mergeFields = IdenfityTokens(runs);
            foreach (MergeFieldTokenParser mergeField in mergeFields)
            {
                mergeField.Data.IsCleanupRequired = _replacer[mergeField.Data.PlaceholderType].Process(mergeField.Data, data);
            }
            Cleanup(mergeFields);
        }

        //removes placeholders
        //We can use replace as well to intact styles.
        private void Cleanup(MergeFieldTokenParser mergeField)
        {
            if (!mergeField.Data.IsCleanupRequired)
            {
                return;
            }
            PlaceholderTokenGroup token = mergeField.Data;
            for (int i = 0; i < token.Placeholders.Count; i++)
            {
                token.Placeholders[i].Cleanup();
            }
            for (int i = 0; token.TextTokens != null && i < token.TextTokens.Count; i++)
            {
                token.TextTokens[i].Cleanup();
            }
        }

        private void Cleanup(List<MergeFieldTokenParser> mergeFields)
        {
            foreach (MergeFieldTokenParser mergeField in mergeFields)
            {
                Cleanup(mergeField);
            }
        }
    }
}
