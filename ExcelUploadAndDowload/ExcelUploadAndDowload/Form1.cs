using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelUploadAndDowload
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Excel.Application excelApp;
        Excel.Workbook workbook;
        Excel._Worksheet worksheet;
        Excel.Range range;

        private void btnGetData_Click(object sender, EventArgs e)
        {
            //OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
                //excelApp = new Excel.Application();
                //workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                //worksheet = workbook.Sheets[1];
                //range = worksheet.UsedRange;

                // Create DataTable to hold the data
                System.Data.DataTable dt = new System.Data.DataTable();

                // Iterate through Excel rows and columns
                for (int i = 1; i <= range.Rows.Count; i++)
                {
                    DataRow row = dt.NewRow();
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        // Add header row to DataTable
                        if (i == 1)
                        {
                            dt.Columns.Add(range.Cells[i, j].Value.ToString());
                        }
                        else
                        {
                            // Add data rows to DataTable
                            row[j - 1] = range.Cells[i, j].Value;
                        }
                    }
                    if (i != 1) dt.Rows.Add(row);
                }

                // Bind DataTable to DataGridView
                dataGridView1.DataSource = dt;

                // Close Excel objects
                workbook.Close();
                excelApp.Quit();

            //}
        }

        private void btnUploadData_Click(object sender, EventArgs e)
        {
            string c = "Data source = KRISHNA\\sqlexpress; Initial catalog = SagarDB; Integrated security = true";
            SqlConnection connection = new SqlConnection(c);
            connection.Open();
            for (int i = 0; i < dataGridView1.Rows.Count;i++)
            {
                string val1 = dataGridView1.Rows[i].Cells[0].Value.ToString() ?? "";
                string val2 = dataGridView1.Rows[i].Cells[1].Value.ToString() ?? "";
                SqlCommand cmd = new SqlCommand("Insert into ExcelData1(Question,Progress) values('" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "','" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "');", connection);
                cmd.ExecuteNonQuery();
            }
            connection.Close();
            MessageBox.Show("Successfull");

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                worksheet = workbook.Sheets[1];
                range = worksheet.UsedRange;
            }
        }

        

        private void btnUpload_Click(object sender, EventArgs e)
        {
            string c = "Data source = KRISHNA\\sqlexpress; Initial catalog = SagarDB; Integrated security = true";
            SqlConnection connection = new SqlConnection(c);
            connection.Open();

            SqlCommand cmd = new SqlCommand("Insert into ExcelData1(Question,Progress) values("+1+");", connection);
            cmd.ExecuteNonQuery();

        }
    }
}
