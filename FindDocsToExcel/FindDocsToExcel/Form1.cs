using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace FindDocsToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string folderPath;
        ArrayList arrayList = new ArrayList();
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                folderPath = dialog.SelectedPath;
                textBox1.Text = folderPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                MessageBox.Show("Please select a folder first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Sheets[1];
            worksheet.Cells[1, 1] = "Doc Files";

            int rowIndex = 2; // Start from row 2 for data

            DirectoryInfo place = new DirectoryInfo(folderPath);
            foreach (FileInfo file in place.GetFiles("*.docx", SearchOption.AllDirectories))
            {
                worksheet.Cells[rowIndex, 1] = file.Name;
                rowIndex++;
            }
            excelApp.Visible = true;
        }
    }
}
