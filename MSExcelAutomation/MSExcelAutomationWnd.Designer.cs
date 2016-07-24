namespace MSExcelAutomation
{
    partial class MSExcelAutomationWnd
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MSExcelAutomationWnd));
            this.automateExcelSpreadsheet = new System.Windows.Forms.Button();
            this.closeMainWnd = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // automateExcelSpreadsheet
            // 
            this.automateExcelSpreadsheet.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.automateExcelSpreadsheet.Cursor = System.Windows.Forms.Cursors.Hand;
            this.automateExcelSpreadsheet.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.automateExcelSpreadsheet.Location = new System.Drawing.Point(300, 318);
            this.automateExcelSpreadsheet.Margin = new System.Windows.Forms.Padding(4);
            this.automateExcelSpreadsheet.Name = "automateExcelSpreadsheet";
            this.automateExcelSpreadsheet.Size = new System.Drawing.Size(179, 32);
            this.automateExcelSpreadsheet.TabIndex = 0;
            this.automateExcelSpreadsheet.Text = "Automate Excel";
            this.automateExcelSpreadsheet.UseVisualStyleBackColor = true;
            this.automateExcelSpreadsheet.Click += new System.EventHandler(this.automateExcelSpreadsheet_Click);
            // 
            // closeMainWnd
            // 
            this.closeMainWnd.BackColor = System.Drawing.Color.Transparent;
            this.closeMainWnd.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("closeMainWnd.BackgroundImage")));
            this.closeMainWnd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.closeMainWnd.Cursor = System.Windows.Forms.Cursors.Hand;
            this.closeMainWnd.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.closeMainWnd.Location = new System.Drawing.Point(716, -2);
            this.closeMainWnd.Name = "closeMainWnd";
            this.closeMainWnd.Size = new System.Drawing.Size(57, 44);
            this.closeMainWnd.TabIndex = 1;
            this.closeMainWnd.UseVisualStyleBackColor = false;
            this.closeMainWnd.Click += new System.EventHandler(this.closeMainWnd_Click);
            // 
            // MSExcelAutomationWnd
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(772, 411);
            this.Controls.Add(this.closeMainWnd);
            this.Controls.Add(this.automateExcelSpreadsheet);
            this.Font = new System.Drawing.Font("Verdana", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MSExcelAutomationWnd";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MSExcelAutomation";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button automateExcelSpreadsheet;
        private System.Windows.Forms.Button closeMainWnd;
    }
}

