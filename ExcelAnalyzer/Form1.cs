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

namespace ExcelAnalyzer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Create a dialog box that can open Excel files.
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Title = "Find an Excel Spreadsheet";
            fbd.Filter = "Microsoft Excel Worksheet|*.xlsx";

            // Display the dialog box, then determine if the OK button was clicked.
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Create an object to represent the entire Microsoft Excel Application.
                // Provide access to the workbook through the filename obtained earlier.
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook book = excelApp.Workbooks.Open(fbd.FileName);

                Excel.Worksheet sheet = excelApp.ActiveSheet as Excel.Worksheet;

                // Count the number of rows and columns in the spreadsheet.
                int totalRows = sheet.UsedRange.Rows.Count;
                int totalColumns = sheet.UsedRange.Columns.Count;

                // Create data table for importing spreadsheet data.
                System.Data.DataTable table = new System.Data.DataTable();                

                //Lazy loops to populate table with spreadsheet data.
                for (int i = 0; i < totalColumns; i++)
                {
                    table.Columns.Add((string)sheet.Cells[1, i + 1].Value2, typeof(string));
                }
                               
                for (int i = 0; i < totalRows - 1; i++)
                {
                    table.Rows.Add();
                    for (int j = 0; j < totalColumns; j++)
                    {
                        table.Rows[i][j] = sheet.Cells[j + 1][i + 2].Value2;
                    }
                }

                // Add table into the data grid.
                dataGridView1.DataSource = table;

                //Close the Excel Workbook.
                book.Close(true, Type.Missing, Type.Missing);


            } else { Application.Exit(); }

            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                MessageBox.Show("YAHAHA! You Found Me!");

                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];

                if (!richTextBox1.Text.Contains(row.Cells["First Name"].Value.ToString()))
                {
                    richTextBox1.Text += row.Cells["First Name"].Value.ToString() + "\n";
                }

            }
        }
    }
}
