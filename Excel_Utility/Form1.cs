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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;




namespace Excel_Utility
{
    public partial class Form1 : Form
    {
        private string selectedFileName;
        private string Job_value;
        private string selectedFolderPath;

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
                selectedFolderPath = folderBrowserDialog.SelectedPath;
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
            if (File_Name.Text == "" && Folder_Name.Text == "" && textBox1.Text == "")
            {
                MessageBox.Show("Please select the 'Input File','Output Folder' and 'Week Ending Date' for processing the data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (Folder_Name.Text == "" && textBox1.Text == "")
            {
                MessageBox.Show("Please select the 'Output Folder' and 'Week Ending Date' for processing the data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox1.Text == "")
            {
                MessageBox.Show("Please select the 'Week Ending Date'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {

                string filePath = selectedFileName;

                ////DataTable rs = new DataTable("DataTable1");
                DataSet1 dataSet = new DataSet1();

                using (var odConnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=NO;';"))
                {
                    odConnection.Open();

                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = odConnection;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT [F1],[F2],[F3],[F4],[F6],[F7],[F8],[F9],[F11] FROM [Source Data$A2:Z]";
                        using (OleDbDataAdapter oleda = new OleDbDataAdapter(cmd))
                        {
                            oleda.Fill(dataSet);
                        }
                    }
                    odConnection.Close();
                }

                MapHeaders(dataSet);

                //==============================================================




                string columnName = "Job";

                // Create a HashSet to store unique string values
                HashSet<string> uniqueValues = new HashSet<string>();

                // Iterate over each DataTable in the DataSet
                foreach (DataTable dataTable in dataSet.Tables)
                {
                    // Get the index of the specified column
                    int columnIndex = dataTable.Columns.IndexOf(columnName);

                    // Check if the column exists in the DataTable
                    if (columnIndex != -1)
                    {
                        // Iterate over each row in the DataTable
                        foreach (DataRow row in dataTable.Rows)
                        {
                            // Get the value of the specified column for the current row
                            string value = row[columnIndex].ToString();

                            // Add the value to the HashSet if it's not already present
                            if (!uniqueValues.Contains(value))
                            {
                                uniqueValues.Add(value);
                                string pattern = @"^\d{2}-\d{5}$";
                                if (Regex.IsMatch(value, pattern))
                                {
                                    // Use LINQ to DataSet to filter rows based on a string condition
                                    var query = from DataRow ro in dataTable.Rows
                                                where ro.Field<string>("Job") == value
                                                select ro;

                                    // Create a new DataTable to store filtered data
                                    DataTable filteredDataTable = query.Any() ? query.CopyToDataTable() : dataTable.Clone();

                                   // dataGridView1.DataSource = filteredDataTable;

                                    DataTable Newtable = new DataTable("DataTable2");
                                    Newtable = filteredDataTable.Copy();
                                    ReportParameter Jobno = new ReportParameter("strJobNo", value);
                                    Job_value = value;

                                    Success_txt.AppendText("Job no " + value + " processed successfully." + Environment.NewLine);

                                    reportViewer1 = new ReportViewer();
                                    reportViewer1.ProcessingMode = ProcessingMode.Local;
                                    reportViewer1.LocalReport.ReportPath = @"C:\\Users\\prathamesh_bhuvad\\Desktop\\VASP SOLUTIONS\\Excel_Utility\\Excel_Utility\\Excel_Utility\\rptJob.rdlc"; // Path to your RDLC report file
                                    this.reportViewer1.LocalReport.SetParameters(Jobno);
                                    // Add your DataSet to the report
                                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", Newtable)); // 'DataSet1' is the name of the dataset in your report

                                    // Display the report
                                   // reportViewer1.Dock = DockStyle.Fill;
                                    //this.Controls.Add(reportViewer1);
                                    //reportViewer1.RefreshReport();

                                    ExportReportToPdf(reportViewer1, selectedFolderPath);



                                }
                                else if (value == "")
                                {
                                    Error_txt.AppendText("Empty Job no detected." + Environment.NewLine);

                                }
                                else if(!Regex.IsMatch(value, pattern))
                                {
                                    Error_txt.AppendText("Job no: " + value + " is not in specified format." + Environment.NewLine);

                                }

                            }
                        }
                    }

                }

                MessageBox.Show("Report exported to PDF successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


                //==============================================================
            }

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

                string outputFileName = $"{Job_value+"_"+textBox1.Text}.pdf";

                // Combine the output directory and file name
                string outputFilePath = Path.Combine(outputPath, outputFileName);

                // Save the PDF to the specified output file path
                File.WriteAllBytes(outputFilePath, pdfBytes);

                //// Save the rendered PDF content to a file
                //File.WriteAllBytes(outputPath, pdfBytes);
                
               // MessageBox.Show("Report exported to PDF successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while exporting the report to PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.reportViewer1.RefreshReport();
        }
        private void MapHeaders(DataSet dataSet)
        {
            // Define mapping from old column names to new column names
            Dictionary<string, string> headerMapping = new Dictionary<string, string>
        {
            {"F1", "WorkDate"},
            {"F2", "PREmployeeNumber"},
            {"F3", "Employee"}, 
            { "F4", "Job" },
            { "F6", "CostCode"},
            { "F7", "CostCodeDescription"},
            { "F8", "PayType"},
            { "F9", "Hours"},
            { "F11", "WorkPerformedComments"},
            
        };

            // Iterate through the tables and columns and map headers
            foreach (DataTable table in dataSet.Tables)
            {
                foreach (DataColumn column in table.Columns)
                {
                    if (headerMapping.ContainsKey(column.ColumnName))
                    {
                        column.ColumnName = headerMapping[column.ColumnName];
                    }
                    // You can add an else clause here if you want to handle columns without a mapping
                }
            }
        }

       
    }
}
