using ExcelTrayLibrary;

namespace WindowsFormsApp1
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
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.excelHelper1 = new ExcelTrayLibrary.ExcelHelper();
            this.excelHelper2 = new ExcelTrayLibrary.ExcelHelper();
            this.excelHelper3 = new ExcelTrayLibrary.ExcelHelper();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(250, 114);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(12, 12);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(157, 134);
            this.listBox1.TabIndex = 1;
            // 
            // excelHelper1
            // 
            this.excelHelper1.FileName = "C:\\Data\\Customers.xlsx";
            this.excelHelper1.Headers = ExcelTrayLibrary.UseHeader.NO;
            this.excelHelper1.Mode = ExcelTrayLibrary.MixType.AsText;
            this.excelHelper1.Provider = ExcelTrayLibrary.ExcelProvider.XLSX;
            // 
            // excelHelper2
            // 
            this.excelHelper2.FileName = "C:\\Data\\Book1.xlsx";
            this.excelHelper2.Headers = ExcelTrayLibrary.UseHeader.YES;
            this.excelHelper2.Mode = ExcelTrayLibrary.MixType.AsText;
            this.excelHelper2.Provider = ExcelTrayLibrary.ExcelProvider.XLSX;
            // 
            // excelHelper3
            // 
            this.excelHelper3.FileName = null;
            this.excelHelper3.Headers = ExcelTrayLibrary.UseHeader.YES;
            this.excelHelper3.Mode = ExcelTrayLibrary.MixType.AsText;
            this.excelHelper3.Provider = ExcelTrayLibrary.ExcelProvider.XLS;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private ExcelHelper excelHelper1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private ExcelHelper excelHelper2;
        private ExcelHelper excelHelper3;
    }
}

