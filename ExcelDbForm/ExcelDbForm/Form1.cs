using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ExcelDbForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            string filePath = "D:\\DotNetWorking\\EmployeeData.xlsx";

            if (System.IO.File.Exists(filePath))
            {
                System.Diagnostics.Process.Start(filePath);
            }
            else
            {
                MessageBox.Show("File doesn't exists");
            }
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {


            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                Excel._Worksheet worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

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


            }




        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = "c:\\",
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Initialize Excel application
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                Excel._Worksheet worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

                // Establish SQL Server connection
                string c = "Data source = KRISHNA\\sqlexpress; Initial catalog = SagarDB; Integrated security = true";
                using (SqlConnection connection = new SqlConnection(c))
                {
                    connection.Open();

                    // Create SqlCommand and SqlBulkCopy objects
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = connection;

                        // Assuming the first row contains column headers
                        for (int i = 2; i <= range.Rows.Count; i++)
                        {
                            cmd.Parameters.Clear();
                            // Generate INSERT statement dynamically based on Excel data
                            string query = "INSERT INTO tbl_empl (empl_id, empl_name, desg_id, dept_id, salary) VALUES (@Column1, @Column2, @Column3, @Column4, @Column5)";
                            cmd.CommandText = query;

                            // Set parameters based on Excel data
                            cmd.Parameters.AddWithValue("@Column1", range.Cells[i, 1].Value);
                            cmd.Parameters.AddWithValue("@Column2", range.Cells[i, 2].Value);
                            cmd.Parameters.AddWithValue("@Column3", range.Cells[i, 3].Value);
                            cmd.Parameters.AddWithValue("@Column4", range.Cells[i, 4].Value);
                            cmd.Parameters.AddWithValue("@Column5", range.Cells[i, 5].Value);
                            // Add more parameters as needed for your table columns

                            // Execute INSERT query
                            cmd.ExecuteNonQuery();
                        }

                        MessageBox.Show("Data transferred successfully to database!");
                    }
                }

                // Clean up Excel objects
                workbook.Close();
                excelApp.Quit();
            }
        }
    }
}
