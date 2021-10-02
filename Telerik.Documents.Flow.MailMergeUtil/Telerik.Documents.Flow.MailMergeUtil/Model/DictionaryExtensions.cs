using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Telerik.Documents.Flow.MailMergeUtil.Model
{
    public static class DictionaryExtensions
    {
        public static DynamicDataObject ToTemplateData(this DataSet data)
        {
            Dictionary<string, object> dictData = data.AsMergeData();
            Dictionary<string, object> parentChildRelationsData = ParentChildAdjustment(data);
            dictData = dictData.Merge(parentChildRelationsData);
            DynamicDataObject templateData = dictData.ToTemplateData();
            return templateData;
        }


        private static Dictionary<string, object> Merge(this Dictionary<string, object> dictA, Dictionary<string, object> dictB)
        {
            return dictA.Keys.Union(dictB.Keys).ToDictionary(k => k, k => dictA.ContainsKey(k) ? dictA[k] : dictB[k]);
        }

        private static DynamicDataObject ToTemplateData(this Dictionary<string, object> dictionary)
        {
            var data = new DynamicDataObject();
            foreach (KeyValuePair<string, object> kvp in dictionary)
            {
                data.Set(kvp.Key, kvp.Value);
            }
            return data;
        }

        private static Dictionary<string, object> ParentChildAdjustment(DataSet data)
        {
            var result = new Dictionary<string, object>();
            if (data.Tables.Contains("ParentChildRelationsTable"))
            {
                DataTable parentChildRelationsTable = data.Tables["ParentChildRelationsTable"];
                if (parentChildRelationsTable.Rows.Count > 0)
                {
                    string parentTableName = Convert.ToString(parentChildRelationsTable.Rows[0]["ParentTable"]);
                    string childTableName = Convert.ToString(parentChildRelationsTable.Rows[0]["ChildTable"]);
                    string parentColumnName = Convert.ToString(parentChildRelationsTable.Rows[0]["ParentTableColumn"]);
                    string childColumnName = Convert.ToString(parentChildRelationsTable.Rows[0]["ChildTableColumn"]);
                    if (data.Tables.Contains(parentTableName) && data.Tables.Contains(childTableName))
                    {
                        DataTable parentTable = data.Tables[parentTableName];
                        DataTable childTable = data.Tables[childTableName];
                        if (parentTable.Rows.Count > 0 && childTable.Rows.Count > 0)
                        {
                            for (int pr = 0; pr < parentTable.Rows.Count; pr++)
                            {
                                DataRow parentRow = parentTable.Rows[pr];
                                object parentColumnVal = parentRow[parentColumnName];
                                if (parentColumnVal != null)
                                {
                                    var childRows = childTable.AsEnumerable().Where(childRow => childRow.Field<object>(childColumnName).Equals(parentColumnVal)).ToList();
                                    result.Add($"{parentTableName}:{pr}{childTableName}:Count", childRows.Count);
                                    for (int cr = 0; cr < childRows.Count; cr++)
                                    {
                                        DataRow childRow = childRows[cr];
                                        foreach (DataColumn currentChildColumn in childTable.Columns)
                                        {
                                            //Parent child format key
                                            result.Add($"{parentTableName}:{pr}{childTableName}:{cr}{currentChildColumn.ColumnName}", childRow[currentChildColumn.ColumnName]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        private static Dictionary<string, object> AsMergeDataTable(this DataTable dt, string tableName = "")
        {
            string regionColumnName = "TableName";
            if (dt.Rows.Count > 0 && dt.Columns.Contains(regionColumnName))
            {
                string tableNameFromDb = Convert.ToString(dt.Rows[0][regionColumnName]);
                tableName = String.IsNullOrWhiteSpace(tableNameFromDb) ? tableName : tableNameFromDb;
            }
            if (!String.IsNullOrWhiteSpace(tableName))
            {
                tableName += ":";
            }

            int rowCount = dt.Rows.Count;
            var dictionaries = Enumerable.Range(0, dt.Rows.Count).Select((i) => AsMergeDataRow(dt.Rows[i], rowCount == 1, tableName, i.ToString())).ToList();

            var result = dictionaries.SelectMany(dict => dict)
                         .ToLookup(pair => pair.Key, pair => pair.Value)
                         .ToDictionary(group => group.Key, group => group.First());

            result.Add($"{tableName}Count", rowCount);
            return result;


            Dictionary<string, object> AsMergeDataRow(DataRow dr, bool hasOnlyOneRow, string tableName = "", string rowIndex = "")
            {
                if (String.IsNullOrWhiteSpace(tableName))
                {
                    rowIndex = ""; //there is no region, no table
                }

                var result = Enumerable.Range(0, dr.Table.Columns.Count).Select((i) => new { Key = $"{tableName}{rowIndex}{dr.Table.Columns[i]}", Value = dr[dr.Table.Columns[i]] }).ToDictionary(k => k.Key, v => v.Value);
                //If table has only one row, merge fields are simple merge fields
                //If table has more than one row, merge fields are converted as array fields like [Name] => [Name[0]], [Name[1]] etc
                if (hasOnlyOneRow)
                {
                    var zeroIndexData = Enumerable.Range(0, dr.Table.Columns.Count).Select((i) => new { Key = $"{dr.Table.Columns[i]}", Value = dr[dr.Table.Columns[i]] }).ToDictionary(k => k.Key, v => v.Value);
                    result = result.Merge(zeroIndexData);
                }

                return result;
            }
        }

        private static Dictionary<string, object> AsMergeData(this DataSet ds)
        {
            UpdateDataSetTableNames(ds);
            //first table is used for header and footer and that is not part of region
            var dictionaries = Enumerable.Range(0, ds.Tables.Count).Select((i) => ds.Tables[i].AsMergeDataTable(ds.Tables[i].TableName)).ToList();

            var result = dictionaries.SelectMany(dict => dict)
                         .ToLookup(pair => pair.Key, pair => pair.Value)
                         .ToDictionary(group => group.Key, group => group.First());

            return result;
        }

        private static void UpdateDataSetTableNames(DataSet ds)
        {
            if (ds.Tables.Count > 1)
            {
                for (int i = 1; i < ds.Tables.Count; i++)
                {
                    if (ds.Tables[i].Columns.Contains("TableName") && ds.Tables[i].Rows.Count > 0)
                    {
                        ds.Tables[i].TableName = UpdateDataSetTableNames(ds.Tables[i]);
                    }

                }
            }
        }

        private static string UpdateDataSetTableNames(DataTable dt)
        {
            string tableName = dt.Columns.Contains("TableName") && dt.Rows.Count > 0 ? dt.Rows[0]["TableName"].ToString() : String.Empty;
            return tableName;
        }
    }
}
