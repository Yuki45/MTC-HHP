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
    public partial class Form1 : Form
    {
        public SqlConnection con;
        public SqlDataAdapter dAdapter;
        public DataSet dSet;
        public string conString = @"Server=107.102.47.105;Database=PROD;user id=sa;password=seinadminhhp";
        public Form1()
        {
            InitializeComponent();
        }

        private static string GetConnectionString()
        {
            String connString =  "Password=dnsdud20vw;User ID=gmes20vw;Data Source=107.102.102.22:1521/SEINMESRW;Persist Security Info=True";
            return connString;
        }

        private static void ConnectingToOracle()
        {
            string connectionString = GetConnectionString();
            using (OracleConnection connection = new OracleConnection())
            {
                connection.ConnectionString = connectionString;
                
                try
                {
                    connection.Open();

                 OracleCommand command = connection.CreateCommand();
                 string sql = "SELECT DISTINCT OWNER, OBJECT_NAME   FROM DBA_OBJECTS  WHERE OBJECT_TYPE = 'TABLE'    AND OWNER = '[some other schema]'";
                 command.CommandText = sql;

                 OracleDataReader reader = command.ExecuteReader();
                 while (reader.Read())
                 {
                     string myField = (string)reader["OBJECT_NAME"];
                    // Console.WriteLine(myField);
                 }
                    MessageBox.Show("test ok");

                    connection.Close();
                }catch
                {
                    MessageBox.Show("test NG");
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ConnectingToOracle();
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now.AddDays(2);
            button1_Click(0, null);
            timer1.Enabled = true;
            timer1.Start();
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

                    string sql = "Select  B.Plan_Ymd As Work_Date,LINE_CODE,  (SELECT SUBSTR(line_nm,1,25) FROM tbm_md_line WHERE fct_code = b.fct_code AND line_code = b.line_code   AND use_yn ='Y'     AND del_yn = 'N') Line,    B.Po_No AS PO,    B.Model_Code As Model,   B.Plan_Qty, NVL(B.Work_Prior,'1') As Priority,   B.Plan_Start_Dt As Start_Datetime,  b.plan_comp_dt AS End_datetime From Tbp_Pm_Po_Ymd_Plan B Where B.Proc_Code = '7160'  And B.Plan_Ymd >= '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "'  and B.Plan_Ymd <= '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' And PLANT_CODE='P529'  ORDER BY B.Plan_Ymd, B.line_code, B.Po_No ";
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
            foreach (DataRow r in hasil.Rows)
            {
                IF_GMES(r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(), r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString());                
            }
        }
        
        private bool IF_GMES(string workdate,string line_code, string line, string po, string model, string qty, string priority, string timestart, string endtime)
        {
            /*-------------------------------------------------
             * work date, line, po, model, plan qty, priority, start time , end time
            EXEC SP_Create_PO_GMES_IF '20120124','SI0A','I201','A01','1','010037921881','SMT-C5050/DIR',500,'0001','20120124080000','20120124090000'
            [dbo].[SP_Create_PO_GMES_IF](@vi_WORK_DATE varchar(10), 
											  @vi_SECT varchar(10), 
											  @vi_SHOP_CODE varchar(10), 
											  @vi_LINE varchar(10),
											  @vi_CHANGE_GROUP char(1),											  
											  @vi_PO_NUMBER varchar(20),
											  @vi_MODEL_NUMBER varchar(25),
											  @vi_PLAN_QTY numeric(6),
											  @vi_PRIORITY varchar(4),
											  @vi_START_DATETIME varchar(20),
											  @vi_END_DATETIME varchar(20)) 

            -------------------------------------------------*/
            bool hasil = false;
            try
            {
                con = new SqlConnection(conString);
                con.Open();
                string sql = "EXEC SP_Create_PO_GMES_IF '" + workdate + "','','" + line_code + "','" + line + "','','" + po + "','" + model + "'," + qty + ",'" + priority + "','" + timestart + "','" + endtime + "'";
                   
                SqlCommand command = new SqlCommand(sql, con);
                command.ExecuteScalar();
                hasil = true;
            }
            catch
            {
                 hasil = false;
            }
            con.Close();
            return hasil;
           
        }

        private void button2_Click(object sender, EventArgs e)
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            //Application.Exit();
        }
    }
}
