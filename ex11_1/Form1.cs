using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;

namespace ex11_1
{
    public partial class frmExport : Form
    {
        SqlConnection booksConnection;
        SqlCommand publishersCommand;
        SqlDataAdapter publishersAdapter;
        DataTable publishersTable;
        public frmExport()
        {
            InitializeComponent();
        }

        private void frmExport_Load(object sender, EventArgs e)
        {
            string path = Path.GetFullPath("SQLBooksDB.mdf");
            // connect to books database
            booksConnection = new
            //    SqlConnection("Data Source=.\\SQLEXPRESS; AttachDBFilename=" + Application.StartupPath + "\\SQLBooksDB.mdf;" +
            //"Integrated Security=True; Connect Timeout=30; User Instance=True");
            SqlConnection("Data Source=.\\SQLEXPRESS; AttachDBFilename=" + path + ";" +
            "Integrated Security=True; Connect Timeout=30; User Instance=True;");
            booksConnection.Open();
            // establish command object
            publishersCommand = new SqlCommand("Select * From Publishers", booksConnection);
            // establish data adapter/data table
            publishersAdapter = new SqlDataAdapter();
            publishersAdapter.SelectCommand = publishersCommand;
            publishersTable = new DataTable();
            publishersAdapter.Fill(publishersTable);
        }

        private void frmExport_FormClosing(object sender, FormClosingEventArgs e)
        {
            // close the connection KWSalesConnection.Close(); // dispose of theobjects
            booksConnection.Dispose();
            publishersCommand.Dispose(); 
            publishersAdapter.Dispose();
            publishersTable.Dispose();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            StreamWriter outputFile; 
            outputFile = new StreamWriter(Application.StartupPath + "\\publishers.csv"); 
            // write headers
            for (int n = 0; n < publishersTable.Columns.Count; n++)
            {
                outputFile.Write(publishersTable.Columns[n]); 
                if (n < publishersTable.Columns.Count - 1) 
                    outputFile.Write(",");
            }
            outputFile.Write(outputFile.NewLine); 
            // write all fields
            foreach(DataRow myRow in publishersTable.Rows) 
            {
                for (int n = 0; n < publishersTable.Columns.Count; n++)
                {
                    if (myRow[n] != null) outputFile.Write(myRow[n].ToString());
                    if (n < publishersTable.Columns.Count - 1)
                        outputFile.Write(",");
                }
                outputFile.Write(outputFile.NewLine);
            }
            outputFile.Close();
        }
    }
}
