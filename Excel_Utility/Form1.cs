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
            //DataSet1 dataSet = new DataSet1();
            //DataTable dataTable = new DataTable("DataTable1");

            //using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            //{
            //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you want to read from the first worksheet

            //    int rowCount = worksheet.Dimension.Rows;
            //    int colCount = worksheet.Dimension.Columns;

            //    // Add columns to DataTable
            //    for (int col = 1; col <= colCount; col++)
            //    {
            //        // Set column names from Excel file headers
            //        string columnName = worksheet.Cells[1, col].Value?.ToString() ?? "Column" + col;
            //        //dataTable.Columns.Add(sheet.Cells[1, col].Value?.ToString() ?? $"Column{col}", typeof(string));
            //         foreach (DataColumn column in dataTable.Columns)
            //    {
            //        column.ColumnName = column.ColumnName.Replace(" ", ""); // Remove spaces
            //    }
            //        dataTable.Columns.Add(columnName);
            //    }
               

            //    // Add rows to DataTable (start from row 2 to skip headers)
            //    for (int row = 2; row <= rowCount; row++)
            //    {
            //        DataRow dataRow = dataTable.Rows.Add();
            //        for (int col = 1; col <= colCount; col++)
            //        {
            //            dataRow[col - 1] = worksheet.Cells[row, col].Value;
            //        }
            //    }
            //}



            //DataTable rs = new DataTable("DataTable1");
            DataSet1 dataSet = new DataSet1();

            using (var odConnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+filePath+";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                odConnection.Open();

                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = odConnection;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT [Work Date],[PR Employee Number],[Employee],[Job],[Cost Code Description],[Pay Type],[Hours] [Work Performed Comments] FROM [Source Data$]";
                    using (OleDbDataAdapter oleda = new OleDbDataAdapter(cmd))
                    { 
                        oleda.Fill(dataSet);
                    }
                }
                odConnection.Close();
            }

            foreach (DataTable dataTable in dataSet.Tables)
            {
                foreach (DataColumn column in dataTable.Columns)
                {
                    column.ColumnName = column.ColumnName.Replace(" ", ""); // Remove spaces
                }
               
            }


            //==============================================================
            
            if (dataSet.Tables.Count > 0)
            {
                dataGridView1.DataSource = dataSet.Tables[1];
            }
            reportViewer1 = new ReportViewer();
            reportViewer1.ProcessingMode = ProcessingMode.Local;
            reportViewer1.LocalReport.ReportPath = @"C:\\Users\\prathamesh_bhuvad\\Desktop\\VASP SOLUTIONS\\Excel_Utility\\Excel_Utility\\Excel_Utility\\rptJob.rdlc"; // Path to your RDLC report file

            // Add your DataSet to the report
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dataSet.Tables[1])); // 'DataSet1' is the name of the dataset in your report

            // Display the report
            reportViewer1.Dock = DockStyle.Fill;
            this.Controls.Add(reportViewer1);
            reportViewer1.RefreshReport();


        }
        private void ExportReportToPdf(ReportViewer reportViewer, string outputPath)
        {
            try
            {
                // Set processing mode to Local
                reportViewer.ProcessingMode = ProcessingMode.Local;

                // Render the report to PDF format
                Warning[] warnings;
                string[] streamIds;
                string mimeType;
                string encoding;
                string fileNameExtension;

                byte[] pdfBytes = reportViewer.LocalReport.Render(
                    "PDF", null, out mimeType, out encoding, out fileNameExtension,
                    out streamIds, out warnings);

                // Save the rendered PDF content to a file
                File.WriteAllBytes(outputPath, pdfBytes);

                MessageBox.Show("Report exported to PDF successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while exporting the report to PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
