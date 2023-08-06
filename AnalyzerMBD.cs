using System;
using System.Data.OleDb;
using System.Windows.Forms;

public partial class MainForm : Form
{
    public MainForm()
    {
        InitializeComponent();
    }

    private void btnLoad_Click(object sender, EventArgs e)
    {
        string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\path\to\your\database.mdb;";

        using (OleDbConnection connection = new OleDbConnection(connString))
        {
            connection.Open();
            string query = "SELECT * FROM YourTableName";

            using (OleDbCommand command = new OleDbCommand(query, connection))
            {
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        // カラムの名前で値を取得します
                        string value = reader["ColumnName"].ToString();

                        // Do something with the value
                    }
                }
            }
        }
    }
}
