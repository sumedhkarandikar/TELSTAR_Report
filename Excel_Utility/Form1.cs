using Microsoft.Reporting.WebForms;
using Microsoft.Reporting.WinForms;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;
using static WPFViewModelExternalControls.PivotDataGrid;




namespace Excel_Utility
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        private string selectedFileName;
        private string Job_value;
        private string selectedFolderPath;
        private string[] ColumnHead;

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
           // FileInfo templateFile = new FileInfo(selectedFileName);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            

            try
            {
                Process.Enabled = false;
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

                    DataSet1 dataSet = new DataSet1();
                    DataTable dt = new DataTable("DataTable2");

                    string[] rowData;

                    // Load Excel file
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                        int[] rowNumber = { 15, 4 ,6};
                        int columnCount = worksheet.Dimension.End.Column;
                        rowData = new string[columnCount * 3];

                        // Read data from the row into the array
                        for (int col = 1; col <= columnCount; col++)
                        {
                            rowData[col - 1] = worksheet.Cells[rowNumber[0], col].Value?.ToString();
                            rowData[col + worksheet.Dimension.End.Column - 1] = worksheet.Cells[rowNumber[1], col].Value?.ToString();
                            rowData[col + 2*worksheet.Dimension.End.Column - 1] = worksheet.Cells[rowNumber[2], col].Value?.ToString();
                            
                        }

                        string[] modifiedArray = new string[rowData.Length];
                        for (int i = 0; i < rowData.Length; i++)
                        {
                            if (rowData[i] != null)
                            {
                                // Remove square brackets from the string
                                modifiedArray[i] = rowData[i].Replace("[", "").Replace("]", "");
                            }

                        }

                        string[] lastWordsArray = new string[modifiedArray.Length];

                        for (int i = 0; i < modifiedArray.Length; i++)
                        {
                            if (modifiedArray[i] != null)
                            {
                                string[] words = modifiedArray[i].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                                if (words.Length > 0)
                                {
                                    // Get the last word
                                    lastWordsArray[i] = words[words.Length - 1];
                                }
                                else
                                {
                                    lastWordsArray[i] = string.Empty;
                                }
                            }
                            else
                            {
                                lastWordsArray[i] = string.Empty;
                            }
                        }
                        ColumnHead = lastWordsArray.Where(item => !string.IsNullOrEmpty(item)).ToArray();



                        DataTable dataTable = new DataTable();
                        string[] columnNames = ColumnHead;
                        foreach (string colu in columnNames)
                        {
                            dataTable.Columns.Add(colu);
                        }

                        ExcelWorksheet worksheet2 = package.Workbook.Worksheets[0];
                        var startCell = worksheet2.Cells["A1"];
                        var endCell = worksheet2.Dimension.End;

                        for (int row = startCell.Start.Row; row <= endCell.Row; row++)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int col = startCell.Start.Column; col <= endCell.Column; col++)
                            {
                                string colu = ExcelColumnToName(col);
                                if (columnNames.Contains(colu))
                                {
                                    dataRow[colu] = worksheet2.Cells[row, col].Value?.ToString();
                                }
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                        
                        dataSet.Tables.Add(dataTable);


                        dt.Merge(dataTable);
                        ChangeColumnHeaders(dt, dt.Columns[0].ColumnName, "WorkDate");
                        ChangeColumnHeaders(dt, dt.Columns[1].ColumnName, "Employee");
                        ChangeColumnHeaders(dt, dt.Columns[2].ColumnName, "PREmployeeNumber");
                        ChangeColumnHeaders(dt, dt.Columns[3].ColumnName, "CostCodeDescription");
                        ChangeColumnHeaders(dt, dt.Columns[4].ColumnName, "PayType");
                        ChangeColumnHeaders(dt, dt.Columns[5].ColumnName, "Hours");
                        ChangeColumnHeaders(dt, dt.Columns[6].ColumnName, "WorkPerformedComments");
                        ChangeColumnHeaders(dt, dt.Columns[7].ColumnName, "Job");
                        ChangeColumnHeaders(dt, dt.Columns[8].ColumnName, "WO");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string dateTimeString = dt.Rows[i]["WorkDate"].ToString();
                            DateTime dateTimeValue;
                            if (DateTime.TryParse(dateTimeString, out dateTimeValue))
                            {
                                dt.Rows[i]["WorkDate"] = dateTimeValue.ToString("yyyy-MM-dd"); 
                            }
                            
                        }


                        Console.WriteLine(dt);
                    }

                    string columnName = dt.Columns[7].ColumnName;
                    HashSet<string> uniqueValues = new HashSet<string>();

                    // Get the index of the specified column
                    int columnIndex = dt.Columns.IndexOf(columnName);


                    if (columnIndex != -1)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string value = row[columnIndex].ToString();
                            if (!uniqueValues.Contains(value))
                            {
                                uniqueValues.Add(value);
                                string pattern = @"^\d{2}-\d{5}$";
                                if (Regex.IsMatch(value, pattern))
                                {
                                    var query = from DataRow ro in dt.Rows
                                                where ro.Field<string>(columnName) == value
                                                select ro;

                                    DataTable filteredDataTable = query.Any() ? query.CopyToDataTable() : dt.Clone();

                                    DataTable Newtable = new DataTable("DataTable2");
                                    Newtable = filteredDataTable.Copy();
                                    ReportParameter Jobno = new ReportParameter("strJobNo", value);
                                    Job_value = value;

                                    string targetColumnName = dt.Columns[8].ColumnName;
                                    string result = null;
                                    foreach (DataRow Wrow in dt.Rows)
                                    {
                                        if (Convert.ToString(Wrow[columnName]) == value)
                                        {
                                            result = Convert.ToString(row[targetColumnName]);
                                            break;
                                        }
                                    }


                                    FileInfo file = new FileInfo(selectedFileName);
                                    if (!file.Exists)
                                    {
                                        throw new FileNotFoundException("Existing Excel file not found.");
                                    }
                                    int startRow = 15;
                                    using (ExcelPackage package = new ExcelPackage(file))
                                    {
                                        
                                        ExcelWorksheet sourceSheet = package.Workbook.Worksheets["Service Field Report"];
                                        if (sourceSheet == null)
                                        {
                                            throw new InvalidOperationException($"Worksheet not found in the existing Excel file.");
                                        }

                                        ExcelWorksheet newSheet = package.Workbook.Worksheets.Add(value, sourceSheet);
                                        //package.Save();

                                        ExcelWorksheet worksheet = package.Workbook.Worksheets[value]; 
                                        if (worksheet == null)
                                        {
                                            throw new InvalidOperationException("Worksheet not found in the existing Excel file.");
                                        }

                                        // Copy DataTable column headers to Excel from the specified row
                                        int columnCount = filteredDataTable.Columns.Count;
                                        for (int col = 1; col <= columnCount; col++)
                                        {
                                            worksheet.Cells[startRow, col].Value = filteredDataTable.Columns[col - 1].ColumnName;
                                        }

                                        worksheet.Cells["J4"].Value = value;
                                        worksheet.Cells["J6"].Value = result;
                                        
                                        // Copy DataTable data to Excel from the specified row
                                        int rowCount = filteredDataTable.Rows.Count;
                                        
                                        for (int Rrow = 0; Rrow < rowCount; Rrow++)
                                        {
                                           

                                            object cellValue1 = filteredDataTable.Rows[Rrow][0];
                                            worksheet.Cells[startRow + Rrow, 1].Value = cellValue1.ToString();
                                            ExcelRange cell1 = worksheet.Cells[startRow + Rrow, 1];
                                            cell1.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell1.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                            object cellValue2 = filteredDataTable.Rows[Rrow][1];
                                            worksheet.Cells[startRow + Rrow, 2].Value = cellValue2.ToString();
                                            ExcelRange cell2 = worksheet.Cells[startRow + Rrow, 2];
                                            cell2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                            object cellValue3 = filteredDataTable.Rows[Rrow][2];
                                            worksheet.Cells[startRow + Rrow, 3].Value = cellValue3.ToString();
                                            ExcelRange cell3 = worksheet.Cells[startRow + Rrow, 3];
                                            cell3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                            object cellValue4 = filteredDataTable.Rows[Rrow][3];
                                            worksheet.Cells[startRow + Rrow, 4].Value = cellValue4.ToString();
                                            ExcelRange cell4 = worksheet.Cells[startRow + Rrow, 4];
                                            cell4.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell4.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell4.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell4.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                            object cellValue5 = filteredDataTable.Rows[Rrow][4];
                                            worksheet.Cells[startRow + Rrow, 5].Value = cellValue5.ToString();
                                            ExcelRange cell5 = worksheet.Cells[startRow + Rrow, 5];
                                            cell5.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell5.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell5.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell5.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                            // Get the value from the 6th column of the filtered DataTable
                                            object cellValue6 = filteredDataTable.Rows[Rrow][5];
                                                worksheet.Cells[startRow + Rrow, 6].Value = cellValue6.ToString();
                                                ExcelRange cell6 = worksheet.Cells[startRow + Rrow, 6];
                                                cell6.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                                cell6.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                                cell6.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                                cell6.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                                            object cellValue7 = filteredDataTable.Rows[Rrow][6];    
                                            worksheet.Cells[startRow + Rrow, 7].Value = cellValue7.ToString();
                                            worksheet.Cells[startRow + Rrow, 7, startRow + Rrow, 13].Merge = true;
                                            ExcelRange cell7 = worksheet.Cells[startRow + Rrow, 7, startRow + Rrow, 13];
                                            cell7.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            cell7.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                            cell7.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                            cell7.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                                        }

                                        package.Save();
                                    }
                                

                                    ReportParameter WOname = new ReportParameter("strWONo", result);

                                    reportViewer1 = new ReportViewer();
                                    reportViewer1.ProcessingMode = ProcessingMode.Local;
                                    string executableDirectory = Application.StartupPath;
                                    string projectDirectory = Directory.GetParent(Directory.GetParent(executableDirectory).FullName).FullName;
                                    //string projectDirectory =executableDirectory;

                                    string reportFolderPath = Path.Combine(projectDirectory, "Report");
                                    string reportFileName = "rptJob.rdlc";
                                    string reportPath = Path.Combine(reportFolderPath, reportFileName);
                                    reportViewer1.LocalReport.ReportPath = reportPath; // Path of RDLC report file
                                    this.reportViewer1.LocalReport.SetParameters(Jobno);
                                    this.reportViewer1.LocalReport.SetParameters(WOname);
                                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", Newtable)); 

                                    ExportReportToPdf(reportViewer1, selectedFolderPath);

                                    Success_txt.AppendText("Job no " + value + " processed successfully." + Environment.NewLine);

                                }
                                else if (value == "")
                                {
                                    Error_txt.AppendText("Empty Job no detected." + Environment.NewLine);

                                }
                                else if (!Regex.IsMatch(value, pattern))
                                {
                                    if (value != "Job")
                                    {
                                        Error_txt.AppendText("Job no: " + value + " is not in specified format." + Environment.NewLine);
                                    }
                                }

                            }
                        }
                    }

                    MessageBox.Show("Report exported to PDF successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch(Exception ex)
            {
                Error_txt.AppendText("Could not process input file." + Environment.NewLine);

                MessageBox.Show($"An error occurred while exporting the report to PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            Process.Enabled = true;

        }


        private void ExportReportToPdf(ReportViewer reportViewer, string outputPath)
       
            {
             try
             {
                
                reportViewer.ProcessingMode = ProcessingMode.Local;

                // Render the report to PDF format
                byte[] renderedBytes;
                string[] streams;
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


                //renderedBytes = reportViewer.LocalReport.Render(
                //"Excel", null, out mimeType, out encoding, out fileNameExtension, out streams, out warnings);

                //// Save the rendered report to a file
                //SaveFileDialog saveFileDialog = new SaveFileDialog();
                //saveFileDialog.Filter = "Excel Files (*.xls)|*.xls";
                //saveFileDialog.FileName = "C:\\Users\\prathamesh_bhuvad\\Desktop\\VASP SOLUTIONS\\output\\Report.xlsx"; // You can set a default file name
                //if (saveFileDialog.ShowDialog() == DialogResult.OK)
                //{
                //    using (FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                //    {
                //        fs.Write(renderedBytes, 0, renderedBytes.Length);
                //        fs.Close();
                //    }
                //}





               






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
       private string ExcelColumnToName(int column)
        {
            int dividend = column;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        
        static void ChangeColumnHeaders(DataTable dataTable, string oldHeader, string newHeader)
        {
            if (dataTable.Columns.Contains(oldHeader))
            {
                dataTable.Columns[oldHeader].ColumnName = newHeader;
            }
        }

    }
}
