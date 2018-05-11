using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace ExportJournal
{
    internal class MdbWriter
    {
        internal string MdbName { get; set; }



        private string ConnectionString
        {
            get
            {
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + MdbName + ";Persist Security Info=False;";
            }
        }

        internal List<string> GetTableNames()
        {
            try
            {
                List<string> tableNames = new List<string>();

                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();

                    var schema = myConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (var row in schema.Rows.OfType<DataRow>())
                    {
                        string tableName = row.ItemArray[2].ToString();
                        tableNames.Add(tableName);
                    }
                    return tableNames;
                }

            }
            catch (Exception ex)
            {

                throw new InvalidOperationException(ex.Message);
            }
        }

        internal void DeleteJournal()
        {
            try
            {
                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();

                    // Execute Queries
                    OleDbCommand cmd = myConnection.CreateCommand();
                    cmd.CommandText = "Drop Table Journal";

                    cmd.ExecuteNonQuery();

                }

            }
            catch (Exception ex)
            {

                throw new InvalidOperationException(ex.Message);
            }
        }
        internal void CreateJournal()
        {
            try
            {
                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();

                    // Execute Queries
                    OleDbCommand cmd = myConnection.CreateCommand();
                    cmd.CommandText = "Create table Journal (Betreff CHAR, Beginntam DATETIME, Dauer INT, Text1 MEMO, Kategorien CHAR )";

                    cmd.ExecuteNonQuery();

                }

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }

        }


        internal void InsertRows(List<JournalEntry> list)
        {
            try
            {
                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();


                    foreach (var je in list)
                    {


                        // Execute Queries
                        OleDbCommand cmd = myConnection.CreateCommand();
                        cmd.CommandText = "INSERT INTO Journal " + "([Betreff], [Beginntam],  [Dauer], [Text1], [Kategorien]) " + "VALUES(@Betreff, @Beginntam, @Dauer, @Text1, @Kategorien)";

                        // add named parameters
                        cmd.Parameters.AddRange(new OleDbParameter[]
                           {
                               new OleDbParameter("@Betreff", je.Betreff??Convert.DBNull),
                               new OleDbParameter("@Beginntam", GetDateWithoutMilliseconds(je.BeginntAm)),
                               new OleDbParameter("@Dauer", je.Dauer),
                               new OleDbParameter("@Text1", je.Text??Convert.DBNull),
                               new OleDbParameter("@Kategorien", je.Kategorien??Convert.DBNull)
                           });




                        cmd.ExecuteNonQuery();
                    }

                }

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }
        }

        private DateTime GetDateWithoutMilliseconds(DateTime d)
        {
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }
    }
}
