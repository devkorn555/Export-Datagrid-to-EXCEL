using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ExpEx
{
    public partial class Form1 : Form
    {

        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook wBook;
        Microsoft.Office.Interop.Excel.Worksheet wSheet;

        SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-MPTIJRG\SQLEXPRESS;Initial Catalog=C0FFEE;User ID=sa;PWD=12345");
        DataTable dt;
        SqlDataAdapter da;

        public Form1()
        {
            InitializeComponent();
        }


        private void Export()
        {
            excel = new Microsoft.Office.Interop.Excel.Application();

            wBook = excel.Workbooks.Add(Type.Missing);
            wSheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.ActiveSheet;
            wSheet.Name = "Test";

            int exheadCol = 1;
            foreach (DataColumn col in dt.Columns)
            {
                excel.Cells[1, exheadCol] = col.ColumnName.ToString();
                exheadCol += 1;
            }


            int exStartCol = 1;
            int exStartRow = 2;
            int coldt = 0;
            foreach (DataColumn col in dt.Columns)
            {
                foreach (DataRow row in dt.Rows)
                {
                    excel.Cells[exStartRow, exStartCol] = row.ItemArray[coldt];
                    exStartRow += 1;
                }

                exStartRow = 2;
                coldt += 1;
                exStartCol += 1;
            }

            string pathFile = Application.StartupPath + @"\testFile.xlsx";

            if (File.Exists(pathFile))
            {
                File.Delete(pathFile);
                wBook.SaveAs(pathFile);
                wBook.Close();
                excel.Quit();

            }
            else
            {
                wBook.SaveAs(pathFile);
                wBook.Close();
                excel.Quit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Export();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            da = new SqlDataAdapter("SELECT * FROM TBCate", con);
            dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
    }
}
