
namespace CovidTestOutlookCollector
{
    partial class CovidTestOutlookCollector
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
            this.Auswertung = new System.Windows.Forms.Button();
            this.LogOutput = new System.Windows.Forms.RichTextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // Auswertung
            // 
            this.Auswertung.Location = new System.Drawing.Point(12, 12);
            this.Auswertung.Name = "Auswertung";
            this.Auswertung.Size = new System.Drawing.Size(776, 46);
            this.Auswertung.TabIndex = 0;
            this.Auswertung.Text = "Auswertung starten";
            this.Auswertung.UseVisualStyleBackColor = true;
            this.Auswertung.Click += new System.EventHandler(this.Auswertung_Click);
            // 
            // LogOutput
            // 
            this.LogOutput.Location = new System.Drawing.Point(12, 64);
            this.LogOutput.Name = "LogOutput";
            this.LogOutput.ReadOnly = true;
            this.LogOutput.Size = new System.Drawing.Size(776, 374);
            this.LogOutput.TabIndex = 1;
            this.LogOutput.Text = "";
            // 
            // CovidTestOutlookCollector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.LogOutput);
            this.Controls.Add(this.Auswertung);
            this.Name = "CovidTestOutlookCollector";
            this.Text = "CovidTestOutlookCollector";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Auswertung;
        private System.Windows.Forms.RichTextBox LogOutput;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

