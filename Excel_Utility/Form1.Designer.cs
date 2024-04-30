namespace Excel_Utility
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Input_btn = new System.Windows.Forms.Button();
            this.Output_btn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Process = new System.Windows.Forms.Button();
            this.Clear = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            this.Success_txt = new System.Windows.Forms.TextBox();
            this.Error_txt = new System.Windows.Forms.TextBox();
            this.File_Name = new System.Windows.Forms.Label();
            this.Folder_Name = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // Input_btn
            // 
            this.Input_btn.Location = new System.Drawing.Point(53, 44);
            this.Input_btn.Name = "Input_btn";
            this.Input_btn.Size = new System.Drawing.Size(155, 23);
            this.Input_btn.TabIndex = 0;
            this.Input_btn.Text = "Select Input Excel File";
            this.Input_btn.UseVisualStyleBackColor = true;
            this.Input_btn.Click += new System.EventHandler(this.Input_btn_Click);
            // 
            // Output_btn
            // 
            this.Output_btn.Location = new System.Drawing.Point(53, 85);
            this.Output_btn.Name = "Output_btn";
            this.Output_btn.Size = new System.Drawing.Size(155, 23);
            this.Output_btn.TabIndex = 1;
            this.Output_btn.Text = "Select Output Folder";
            this.Output_btn.UseVisualStyleBackColor = true;
            this.Output_btn.Click += new System.EventHandler(this.Output_btn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 157);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Week Ending Date (YYYYMMDD)";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(252, 150);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(203, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Validating += new System.ComponentModel.CancelEventHandler(this.textBox1_Validating_1);
            // 
            // Process
            // 
            this.Process.Location = new System.Drawing.Point(212, 197);
            this.Process.Name = "Process";
            this.Process.Size = new System.Drawing.Size(115, 23);
            this.Process.TabIndex = 4;
            this.Process.Text = "Process";
            this.Process.UseVisualStyleBackColor = true;
            this.Process.Click += new System.EventHandler(this.Process_Click);
            // 
            // Clear
            // 
            this.Clear.Location = new System.Drawing.Point(195, 496);
            this.Clear.Name = "Clear";
            this.Clear.Size = new System.Drawing.Size(75, 23);
            this.Clear.TabIndex = 5;
            this.Clear.Text = "Clear";
            this.Clear.UseVisualStyleBackColor = true;
            this.Clear.Click += new System.EventHandler(this.Clear_Click);
            // 
            // Exit
            // 
            this.Exit.Location = new System.Drawing.Point(314, 496);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(75, 23);
            this.Exit.TabIndex = 6;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // Success_txt
            // 
            this.Success_txt.Location = new System.Drawing.Point(63, 260);
            this.Success_txt.Multiline = true;
            this.Success_txt.Name = "Success_txt";
            this.Success_txt.Size = new System.Drawing.Size(223, 80);
            this.Success_txt.TabIndex = 7;
            // 
            // Error_txt
            // 
            this.Error_txt.Location = new System.Drawing.Point(314, 260);
            this.Error_txt.Multiline = true;
            this.Error_txt.Name = "Error_txt";
            this.Error_txt.Size = new System.Drawing.Size(223, 32);
            this.Error_txt.TabIndex = 8;
            // 
            // File_Name
            // 
            this.File_Name.AutoSize = true;
            this.File_Name.Location = new System.Drawing.Point(354, 90);
            this.File_Name.Name = "File_Name";
            this.File_Name.Size = new System.Drawing.Size(35, 13);
            this.File_Name.TabIndex = 9;
            this.File_Name.Text = "label2";
            // 
            // Folder_Name
            // 
            this.Folder_Name.AutoSize = true;
            this.Folder_Name.Location = new System.Drawing.Point(252, 90);
            this.Folder_Name.Name = "Folder_Name";
            this.Folder_Name.Size = new System.Drawing.Size(35, 13);
            this.Folder_Name.TabIndex = 10;
            this.Folder_Name.Text = "label2";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(95, 241);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(161, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Successfully Processed Job Nos";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(389, 244);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "Errors / Issues";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(314, 313);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(240, 150);
            this.dataGridView1.TabIndex = 13;
            // 
            // reportViewer1
            // 
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "Excel_Utility.rptJob.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(619, 33);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(396, 246);
            this.reportViewer1.TabIndex = 14;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1108, 546);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Folder_Name);
            this.Controls.Add(this.File_Name);
            this.Controls.Add(this.Error_txt);
            this.Controls.Add(this.Success_txt);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.Clear);
            this.Controls.Add(this.Process);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Output_btn);
            this.Controls.Add(this.Input_btn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Input_btn;
        private System.Windows.Forms.Button Output_btn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Process;
        private System.Windows.Forms.Button Clear;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.TextBox Success_txt;
        private System.Windows.Forms.TextBox Error_txt;
        private System.Windows.Forms.Label File_Name;
        private System.Windows.Forms.Label Folder_Name;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
    }
}

