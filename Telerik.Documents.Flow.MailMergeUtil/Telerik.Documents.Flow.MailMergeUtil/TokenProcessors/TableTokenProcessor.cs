using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Flow.MailMergeUtil.Model;
using Telerik.Documents.Flow.MailMergeUtil.Tokenizer;
using Telerik.Documents.Media;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

namespace Telerik.Documents.Flow.MailMergeUtil.TokenProcessors
{
    public class TableTokenProcessor : IMergeTokenProcessor
    {
        private readonly string _tableStartToken = "[";
        private readonly string _tableEndToken = "]";

        private string GetTableName(string tableNameToken)
        {
            string tableName = tableNameToken?.Split(new char[] { _tableStartToken[0], ':', _tableEndToken[0] }, StringSplitOptions.RemoveEmptyEntries).Skip(1).First();
            return tableName.Trim();
        }

        private int GetTableCount(string tableName, DynamicDataObject data)
        {
            object count = data.Get($"{tableName}:Count");
            return count != null ? (int)count : 0;
        }

        public bool Process(PlaceholderTokenGroup token, DynamicDataObject data)
        {
            PlaceholderToken start = token.Placeholders.First();
            string tableName = GetTableName(token.Placeholder);
            int dataItemCount = GetTableCount(tableName, data);
            //TableStart:TableName, TableEnd:TableName and at least one placeholder
            if (dataItemCount <= 0 || token.Placeholders.Count <= 2)
            {
                //return false; // false => placeholder table will be as it is if there is no data.
                return true; //true => for removing table if there is nothing to replace
            }

            var tableTokens = token.Placeholders.SkipWhile((ri) =>
            {
                return !ri.Text.Contains(_tableStartToken); //skip all tokens till table start definition - [TableStart:Name]
            }).SkipWhile((ri) =>
            {
                return !ri.Text.Contains(_tableEndToken); //skip all tokens till first field found - [TableStart:Name] [Item]
            }).ToList();
            //a table may have multiple placeholders, to process table need to find out actual table object from any child object
            Table table = null;
            foreach (PlaceholderToken placeholder in tableTokens)
            {
                DocumentElementBase elementBase = placeholder.Parent.Parent;
                table = elementBase as Table;
                //TableStart/TableEnd token may be part of parent table. In this case it identifies parent table and process it instead of actual table
                if (placeholder.Text.Equals(_tableStartToken) || placeholder.Text.Equals(_tableEndToken))
                {
                    continue;
                }
                //for current element, loop thru each paraenet of a child to find out actual table object
                while (elementBase != null && table == null)
                {
                    elementBase = elementBase.Parent;
                    table = elementBase as Table;
                }
                if (table != null)
                {
                    break;
                }
            };

            if (table == null)
            {
                return false;
            }

            //Based on given item count, cloning rows
            //if there are 2 items passed in DataTable, then create 2 rows
            //a table may have multiple rows with placeholder
            //Example 1 - single row, multiple columns
            //[Name], [Address]

            //Example 2 - multiple rows, single/multiple columns
            //[Name]
            //[Address]
            int templatedTableRows = table.Rows.Count;

            /* Removing other columns for "No Data Message"
             *  Table may have multiple columns
             *  "No Data Message" should have one row and one column so that message should not be wrapped into single cell
             *  Remove other cells if there are more than 1 cell in the table
            */
            if (dataItemCount == 1)
            {
                //Find the key having NODATAMSG###
                //If key is avaialble then remove all other cells except first one
                //Write message after removing NODATAMSG###
                var noDataMsgCell = data.DataDictionary.Where(a => a.Key.StartsWith($"{tableName}:") && Convert.ToString(a.Value).StartsWith("NODATAMSG###")).Select(k => k.Key).FirstOrDefault();
                if (noDataMsgCell != null)
                {
                    data.DataDictionary[noDataMsgCell] = Convert.ToString(data.DataDictionary[noDataMsgCell]).Replace("NODATAMSG###", String.Empty);
                    while (table.Rows[0].Cells.Count > 1)
                    {
                        //Removing first cell always re-arranges indexes that's why removing first cell only
                        table.Rows[0].Cells.Remove(table.Rows[0].Cells[1]);
                    }
                }
            }

            //first row is always skipped as it contains placeholders/merge fields
            for (int i = 1; i < dataItemCount; i++)
            {
                for (int r = 0; r < templatedTableRows; r++)
                {
                    table.Rows.Add(table.Rows[r].Clone());
                }
            }

            IEnumerable<string> columns = data.GetColumnNames();
            for (int i = 0; i < dataItemCount; i++)
            {
                //data has following keys for list /datatable
                //TableName:0Name, TableName:0Address
                //TableName:1Name, TableName:1Address
                //based on current index, it gets all columns from data objects
                //and replace cell values of a row
                string tableRowKey = $"{tableName}:{i}";
                var rowData = new DynamicDataObject();
                //Replace will be tricky
                //Contributions:0 will replace these both data => Contributions:0 and Contributions:1FundContributions:0
                var rowKeys = columns.Where(w => w.StartsWith(tableRowKey)).Select(s => new { RowKey = (s.IndexOf(tableRowKey) > -1 ? s.Remove(s.IndexOf(tableRowKey), tableRowKey.Length) : s), DataKey = s }).ToList();
                foreach (var key in rowKeys)
                {
                    rowData.Set(key.RowKey, data.Get(key.DataKey));
                }


                for (int r = 0; r < templatedTableRows; r++)
                {
                    int rowIndex = (templatedTableRows * i) + r;
                    TableRow row = table.Rows[rowIndex];
                    var runs = row.EnumerateChildrenOfType<Run>().ToList();
                    var mergeFieldProcesses = new List<MergeFieldTokenizerBase>() { new MergeFieldSquareTokenizerProcess(), new MergeFieldTriangleTokenizerProcess() };
                    foreach (MergeFieldTokenizerBase mergeFieldProcess in mergeFieldProcesses)
                    {
                        mergeFieldProcess.MailMerge(runs, rowData);
                    }
                }
            }
            return true;
        }

    }
}
