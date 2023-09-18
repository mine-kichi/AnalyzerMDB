using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnalyzerMBD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<string> GetColumns(string connectionString, string tableName)
        {
            List<string> columns = new List<string>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Get the schema of the table
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                                          new object[] { null, null, tableName, null });

                foreach (DataRow row in schemaTable.Rows)
                {
                    columns.Add(row["COLUMN_NAME"].ToString());
                }
            }

            return columns;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\work\programming\github\Tools\AnalyzerMBD\Database1.mdb;";

            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                foreach (DataRow row in schemaTable.Rows)
                {
                    Console.WriteLine(row["TABLE_NAME"].ToString());
                }
            }

            string query = "SELECT ID FROM [table];";
            List<string> columnNames = GetColumns(connString, "table");

            foreach (string columnName in columnNames)
            {
                Console.WriteLine(columnName);
            }

            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            // カラムの名前で値を取得します
                            string value = reader["ID"].ToString();
                            Console.WriteLine("ID:" + value);

                            // Do something with the value
                        }
                    }
                }
            }


            /*            OleDbConnection conn = new OleDbConnection();
                        OleDbCommand comm = new OleDbCommand();

                        conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\work\programming\github\Tools\AnalyzerMBD\Database1.mdb"; // MDB名など
            //            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\work\programming\github\Tools\AnalyzerMBD\Database1.mdb"; // MDB名など

                        // 接続します。
                        conn.Open();

                        // SELECT文を設定します。
                        comm.CommandText = @"SELECT count(*) FROM table;";
                        comm.Connection = conn;
                        OleDbDataReader reader = comm.ExecuteReader();

                        // 結果を表示します。
                        while (reader.Read())
                        {
                            int id = (int)reader.GetValue(0);
                            string accountName = (string)reader.GetValue(1);
                            int accountNumber = (int)reader.GetValue(2);

            //                Console.WriteLine("ID:" + id + " AccountName:" + accountName + " AccountNumber:" + accountNumber);
                        }
                        // 接続を解除します。
                        conn.Close();
            */


            /*            //MDBに接続
                        System.Data.OleDb.OleDbConnection Con = new System.Data.OleDb.OleDbConnection();
                        Con.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\work\programming\github\Tools\AnalyzerMBD\Database1.mdb"; // MDB名など;
                        Con.Open();
                        //SQL文を実行
                        System.Data.OleDb.OleDbCommand command = Con.CreateCommand();
                        //実行するSQLクエリーを指定
                        command.CommandText = "SELECT * FROM [table]";
                        //結果が返ってくるまで待機する秒数
                        command.CommandTimeout = 30;
                        //指定した SQL コマンドを実行して SqlDataReader を構築する
                        System.Data.OleDb.OleDbDataReader reader = command.ExecuteReader();
                        command.Dispose();
                        //読み込んだデータを表示
                        while (reader.Read())
                        {
                            Console.WriteLine(reader[0].ToString() + "," + reader[1].ToString());
                            Console.WriteLine(reader["Id"].ToString() + "," + reader["Name"].ToString());
                        }
                        reader.Close();
                        //接続を閉じる
                        Con.Close();*/
        }
    }
}
