namespace GoogleDocsIntegration.Forms
{
    partial class GoogleDocsHelperControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.lblActiveWks = new System.Windows.Forms.Label();
            this.btnRetrieveData = new System.Windows.Forms.Button();
            this.btnPushData = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(137, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Active Google Spreadsheet";
            // 
            // lblActiveWks
            // 
            this.lblActiveWks.AutoSize = true;
            this.lblActiveWks.Location = new System.Drawing.Point(7, 38);
            this.lblActiveWks.Name = "lblActiveWks";
            this.lblActiveWks.Size = new System.Drawing.Size(10, 13);
            this.lblActiveWks.TabIndex = 1;
            this.lblActiveWks.Text = "-";
            // 
            // btnRetrieveData
            // 
            this.btnRetrieveData.Location = new System.Drawing.Point(10, 76);
            this.btnRetrieveData.Name = "btnRetrieveData";
            this.btnRetrieveData.Size = new System.Drawing.Size(107, 23);
            this.btnRetrieveData.TabIndex = 2;
            this.btnRetrieveData.Text = "Retrieve Data";
            this.btnRetrieveData.UseVisualStyleBackColor = true;
            this.btnRetrieveData.Click += new System.EventHandler(this.btnRetrieveData_Click);
            // 
            // btnPushData
            // 
            this.btnPushData.Location = new System.Drawing.Point(10, 105);
            this.btnPushData.Name = "btnPushData";
            this.btnPushData.Size = new System.Drawing.Size(107, 23);
            this.btnPushData.TabIndex = 2;
            this.btnPushData.Text = "Push Data";
            this.btnPushData.UseVisualStyleBackColor = true;
            // 
            // GoogleDocsHelperControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnPushData);
            this.Controls.Add(this.btnRetrieveData);
            this.Controls.Add(this.lblActiveWks);
            this.Controls.Add(this.label1);
            this.Name = "GoogleDocsHelperControl";
            this.Size = new System.Drawing.Size(311, 530);
            this.Load += new System.EventHandler(this.GoogleDocsHelperControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblActiveWks;
        private System.Windows.Forms.Button btnRetrieveData;
        private System.Windows.Forms.Button btnPushData;
    }
}
