using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace DatabaseToDataDictionary
{
    class Program
    {
        static SqlConnection conn;
        static SqlCommand cmd;
        static SqlDataAdapter adp;

        static void Main(string[] args)
        {
            string connStr = @"Server=.\SQLEXPRESS;
                               Database = ServisTakip;
                               Trusted_Connection = True;";
            conn = new SqlConnection(connStr);
            cmd = new SqlCommand();
            cmd.Connection = conn;
            adp = new SqlDataAdapter();
            adp.SelectCommand = cmd;
            conn.Open();

            string fileName = "DataDictionary.docx";
            var doc = DocX.Create(fileName);

            List<KeyValuePair<string, List<TableDetail>>> listTableNameAndColumns = GetTableNames();
            for (int i = 0; i < listTableNameAndColumns.Count; i++)
            {
                KeyValuePair<string, List<TableDetail>> keyValuePair = listTableNameAndColumns[i];

                WriteDocx(doc, keyValuePair.Key, keyValuePair.Value);
            }
            doc.Save();
            Process.Start("WINWORD.EXE", fileName);
        }

        static void WriteDocx(DocX doc, string key, List<TableDetail> value)
        {
            Table t = doc.AddTable(value.Count + 1, 3);
            t.Alignment = Alignment.center;
            t.Design = TableDesign.ColorfulList;
            t.Rows[0].Cells[0].Paragraphs.First().Append(key);
            t.Rows[0].Cells[1].Paragraphs.First().Append("DataType");
            t.Rows[0].Cells[2].Paragraphs.First().Append("Explanation");
            for (int i = 0; i < value.Count; i++)
            {
                TableDetail tableDetail = value[i];
                t.Rows[i + 1].Cells[0].Paragraphs.First().Append(tableDetail.ColumnName);
                t.Rows[i + 1].Cells[1].Paragraphs.First().Append(tableDetail.Type);
                t.Rows[i + 1].Cells[2].Paragraphs.First().Append("");
            }
            doc.InsertTable(t);
            Paragraph paragraph = doc.InsertParagraph("\n");
        }

        static List<KeyValuePair<string, List<TableDetail>>> GetTableNames()
        {
            List<KeyValuePair<string, List<TableDetail>>> list = new List<KeyValuePair<string, List<TableDetail>>>();
            string sql = "Select name From sys.Tables";
            cmd.CommandText = sql;
            DataTable dt = new DataTable();
            adp.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string tableName = dt.Rows[i][0].ToString();

                List<TableDetail> tblDetail = GetTable(tableName);
                KeyValuePair<string, List<TableDetail>> keyValuePair = new KeyValuePair<string, List<TableDetail>>(tableName, tblDetail);
                list.Add(keyValuePair);
            }
            return list;
        }

        static List<TableDetail> GetTable(string tableName)
        {
            string sql = "select column_Name, data_type, character_maximum_Length from INFORMATION_SCHEMA.COLUMNS where table_name=@TableName";
            cmd.Parameters.Clear();
            cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = tableName;
            cmd.CommandText = sql;
            DataTable dt = new DataTable();
            adp.Fill(dt);
            List<TableDetail> listTable = new List<TableDetail>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string columnName = dt.Rows[i][0].ToString();
                string dataType = dt.Rows[i][1].ToString();
                if (dataType == "varchar") dataType += "(" + dt.Rows[i][2].ToString() + ")";
                listTable.Add(new TableDetail { ColumnName = columnName, Type = dataType });
            }
            return listTable;
        }
    }

    class TableDetail
    {
        public string ColumnName { get; set; }

        public string Type { get; set; }
    }
}
