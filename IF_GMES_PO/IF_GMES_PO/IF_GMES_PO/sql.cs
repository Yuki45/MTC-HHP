using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Data.Sql;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace IF_GMES_PO
{
    public partial class sql : Form
    {
        public SqlConnection con;
        public SqlDataAdapter dAdapter;
        public DataSet dSet;
        public string conString = @"Server=107.102.47.105;Database=PROD;user id=sa;password=seinadminhhp";
        public sql()
        {
            InitializeComponent();
        }

        private static string GetConnectionString()
        {
            String connString = "Password=dnsdud20vw;User ID=gmes20vw;Data Source=107.102.102.22:1521/SEINMESRW;Persist Security Info=True";
            return connString;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = GetConnectionString();
            DataTable hasil = new DataTable();
            using (OracleConnection connection = new OracleConnection())
            {
                connection.ConnectionString = connectionString;

                try
                {
                    connection.Open();

                    string sql = textBox1.Text;
                    OracleCommand command = new OracleCommand(sql, connection);

                    OracleDataAdapter reader = new OracleDataAdapter(command);

                    reader.Fill(hasil);
                    connection.Close();

                    dgView.DataSource = hasil;
                }
                catch
                {
                    MessageBox.Show("test NG");
                }

            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void exportExcelToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            int rowsTotal = 0;
            int colsTotal = 0;
            int I = 0;
            int j = 0;
            int iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();

            try
            {
                Excel.Workbook excelBook = xlApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = dgView.RowCount - 1;
                colsTotal = dgView.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = dgView.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = dgView.Rows[I].Cells[j].Value;
                    }
                }
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //RELEASE ALLOACTED RESOURCES
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }
    }
}
