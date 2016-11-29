namespace SwTEst2
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
            this.Macro1_but = new System.Windows.Forms.Button();
            this.Macro2_but = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Macro1_but
            // 
            this.Macro1_but.Location = new System.Drawing.Point(12, 12);
            this.Macro1_but.Name = "Macro1_but";
            this.Macro1_but.Size = new System.Drawing.Size(338, 82);
            this.Macro1_but.TabIndex = 0;
            this.Macro1_but.Text = "Macro1 Сопряжение";
            this.Macro1_but.UseVisualStyleBackColor = true;
            this.Macro1_but.Click += new System.EventHandler(this.Macro1_but_Click);
            // 
            // Macro2_but
            // 
            this.Macro2_but.Location = new System.Drawing.Point(12, 100);
            this.Macro2_but.Name = "Macro2_but";
            this.Macro2_but.Size = new System.Drawing.Size(338, 82);
            this.Macro2_but.TabIndex = 1;
            this.Macro2_but.Text = "Macro2 Корпус";
            this.Macro2_but.UseVisualStyleBackColor = true;
            this.Macro2_but.Click += new System.EventHandler(this.Macro2_but_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(362, 195);
            this.Controls.Add(this.Macro2_but);
            this.Controls.Add(this.Macro1_but);
            this.Name = "Form1";
            this.Text = "Solid Works Курсовая работа";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Macro1_but;
        private System.Windows.Forms.Button Macro2_but;
    }
}

