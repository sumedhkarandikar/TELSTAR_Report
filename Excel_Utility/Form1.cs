using ExcelDataReader;
using Microsoft.Reporting.WebForms;
using Microsoft.Reporting.WinForms;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Excel_Utility
{
    public partial class Form1 : Form
    {
        private string selectedFileName;
       
        public Form1()
        {
            InitializeComponent();
            File_Name.Visible = false;
            Folder_Name.Visible = false;
        }

        public void Input_btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Title = "Select a file";
            openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";

            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFileName = openFileDialog.FileName;
                File_Name.Text = selectedFileName;
                File_Name.Visible = true;
            }
        }

        private void Output_btn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            folderBrowserDialog.Description = "Select a folder to save the file";

            
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFolderPath = folderBrowserDialog.SelectedPath;
                Folder_Name.Text = selectedFolderPath;
                Folder_Name.Visible = true;

            }
        }
        private void textBox1_Validating_1(object sender, CancelEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string input = textBox.Text.Trim();

            if (!IsValidDate(input))
            {
                MessageBox.Show("Invalid date format. Please enter a date in the format YYYYMMDD.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox.Focus();
                e.Cancel = true;
            }
        }
        private bool IsValidDate(string input)
        {
            // Check if the input string has 8 characters
            if (input.Length != 8)
                return false;

            // Check if all characters are digits
            if (!input.All(char.IsDigit))
                return false;

            // Parse the input string as a date
            if (!DateTime.TryParseExact(input, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out _))
                return false;

            return true;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            File_Name.Text = "";
            Folder_Name.Text = "";
            textBox1.Text = "";
            Success_txt.Text = "";
            Error_txt.Text = "";
        }
       
        private void Process_Click(object sender, EventArgs e)
        {

            string filePath = selectedFileName;


            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                
                DataSet1 dataSet = new DataSet1();
                int sheetIndex = 0; // Index of the worksheet to read, 0 for the first sheet
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                DataTable dataTable = new DataTable("DataTable1");

                // Load columns
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // Load rows
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    ExcelRange row = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow newRow = dataTable.Rows.Add();
                    foreach (var cell in row)
                    {

                        newRow[cell.Start.Column - 1] = cell.Text;
                    }
                }
                

                dataSet.Tables.Add(dataTable);
                if (dataSet.Tables.Count > 0)
                {
                    dataGridView1.DataSource = dataSet.Tables[1];
                }

                //string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", filePath);
                //string query = string.Format("SELECT * FROM [{0}$]", worksheet.Name);

                //DataSet data = new DataSet();
                //using (OleDbConnection con = new OleDbConnection(connectionString))
                //{
                //    con.Open();
                //    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                //    adapter.Fill(data);
                //}



                //////////////////////////////////////////////
                ///

                 reportViewer2 = new ReportViewer();
                reportViewer2.ProcessingMode = ProcessingMode.Local;
                reportViewer2.LocalReport.ReportPath = @"C:\\Users\\prathamesh_bhuvad\\Desktop\\VASP SOLUTIONS\\Excel_Utility\\Excel_Utility\\Excel_Utility\\rptJob.rdlc"; // Path to your RDLC report file

                // Add your DataSet to the report
                reportViewer2.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dataSet.Tables[1])); // 'DataSet1' is the name of the dataset in your report

                // Display the report
                reportViewer2.Dock = DockStyle.Fill;
                this.Controls.Add(reportViewer2);
                reportViewer2.RefreshReport();








            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            
        }

        private void reportViewer2_Load(object sender, EventArgs e)
        {
           
        }
    }
}
