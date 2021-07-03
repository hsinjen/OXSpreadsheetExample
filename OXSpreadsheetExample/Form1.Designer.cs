
namespace OXSpreadsheetExample
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btn新增列 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn新增列
            // 
            this.btn新增列.Location = new System.Drawing.Point(12, 12);
            this.btn新增列.Name = "btn新增列";
            this.btn新增列.Size = new System.Drawing.Size(150, 46);
            this.btn新增列.TabIndex = 0;
            this.btn新增列.Text = "新增列";
            this.btn新增列.UseVisualStyleBackColor = true;
            this.btn新增列.Click += new System.EventHandler(this.btn新增列_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 30F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn新增列);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn新增列;
    }
}

