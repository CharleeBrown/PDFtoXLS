namespace PDFtoXLS
{
    partial class mainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.excelBtn = new System.Windows.Forms.Button();
            this.pullBtn = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // excelBtn
            // 
            this.excelBtn.Location = new System.Drawing.Point(58, 91);
            this.excelBtn.Name = "excelBtn";
            this.excelBtn.Size = new System.Drawing.Size(104, 36);
            this.excelBtn.TabIndex = 0;
            this.excelBtn.Text = "PDF to Excel";
            this.excelBtn.UseVisualStyleBackColor = true;
            this.excelBtn.Click += new System.EventHandler(this.button1_Click);
            // 
            // pullBtn
            // 
            this.pullBtn.Location = new System.Drawing.Point(58, 36);
            this.pullBtn.Name = "pullBtn";
            this.pullBtn.Size = new System.Drawing.Size(104, 30);
            this.pullBtn.TabIndex = 2;
            this.pullBtn.Text = "Get Files";
            this.pullBtn.UseVisualStyleBackColor = true;
            this.pullBtn.Click += new System.EventHandler(this.button2_Click);
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(190, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(457, 317);
            this.listView1.TabIndex = 4;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSkyBlue;
            this.ClientSize = new System.Drawing.Size(659, 363);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.pullBtn);
            this.Controls.Add(this.excelBtn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "mainForm";
            this.Text = "PDFtoXLS";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button excelBtn;
        private System.Windows.Forms.Button pullBtn;
        private System.Windows.Forms.ListView listView1;
    }
}

