namespace Massen_Mailer
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.StartenButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.CSVPfadTextBox = new System.Windows.Forms.TextBox();
            this.CSVPfadSuchenButton = new System.Windows.Forms.Button();
            this.StatusLabel = new System.Windows.Forms.Label();
            this.Ladebalken = new System.Windows.Forms.ProgressBar();
            this.TestEmailAdresseTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TestButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // StartenButton
            // 
            this.StartenButton.Enabled = false;
            this.StartenButton.Location = new System.Drawing.Point(436, 104);
            this.StartenButton.Name = "StartenButton";
            this.StartenButton.Size = new System.Drawing.Size(88, 23);
            this.StartenButton.TabIndex = 0;
            this.StartenButton.Text = "Alle Versenden";
            this.StartenButton.UseVisualStyleBackColor = true;
            this.StartenButton.Click += new System.EventHandler(this.StartenButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Steuerdatei";
            // 
            // CSVPfadTextBox
            // 
            this.CSVPfadTextBox.BackColor = System.Drawing.Color.White;
            this.CSVPfadTextBox.Location = new System.Drawing.Point(90, 21);
            this.CSVPfadTextBox.Name = "CSVPfadTextBox";
            this.CSVPfadTextBox.ReadOnly = true;
            this.CSVPfadTextBox.Size = new System.Drawing.Size(340, 20);
            this.CSVPfadTextBox.TabIndex = 2;
            // 
            // CSVPfadSuchenButton
            // 
            this.CSVPfadSuchenButton.Location = new System.Drawing.Point(436, 19);
            this.CSVPfadSuchenButton.Name = "CSVPfadSuchenButton";
            this.CSVPfadSuchenButton.Size = new System.Drawing.Size(88, 23);
            this.CSVPfadSuchenButton.TabIndex = 3;
            this.CSVPfadSuchenButton.Text = "Durchsuchen";
            this.CSVPfadSuchenButton.UseVisualStyleBackColor = true;
            this.CSVPfadSuchenButton.Click += new System.EventHandler(this.CSVPfadSuchenButton_Click);
            // 
            // StatusLabel
            // 
            this.StatusLabel.AutoSize = true;
            this.StatusLabel.Location = new System.Drawing.Point(23, 117);
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(168, 13);
            this.StatusLabel.TabIndex = 4;
            this.StatusLabel.Text = "Status: 0/0 versendet ETA: 0m 0s";
            // 
            // Ladebalken
            // 
            this.Ladebalken.Location = new System.Drawing.Point(26, 133);
            this.Ladebalken.Name = "Ladebalken";
            this.Ladebalken.Size = new System.Drawing.Size(498, 23);
            this.Ladebalken.TabIndex = 5;
            // 
            // TestEmailAdresseTextBox
            // 
            this.TestEmailAdresseTextBox.Location = new System.Drawing.Point(126, 47);
            this.TestEmailAdresseTextBox.Name = "TestEmailAdresseTextBox";
            this.TestEmailAdresseTextBox.Size = new System.Drawing.Size(398, 20);
            this.TestEmailAdresseTextBox.TabIndex = 7;
            this.TestEmailAdresseTextBox.TextChanged += new System.EventHandler(this.TestEmailAdresseTextBox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Test Email-Adresse";
            // 
            // TestButton
            // 
            this.TestButton.Enabled = false;
            this.TestButton.Location = new System.Drawing.Point(336, 104);
            this.TestButton.Name = "TestButton";
            this.TestButton.Size = new System.Drawing.Size(94, 23);
            this.TestButton.TabIndex = 8;
            this.TestButton.Text = "Test versenden";
            this.TestButton.UseVisualStyleBackColor = true;
            this.TestButton.Click += new System.EventHandler(this.TestButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(554, 178);
            this.Controls.Add(this.TestButton);
            this.Controls.Add(this.TestEmailAdresseTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Ladebalken);
            this.Controls.Add(this.StatusLabel);
            this.Controls.Add(this.CSVPfadSuchenButton);
            this.Controls.Add(this.CSVPfadTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartenButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Massen Mailer";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button StartenButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox CSVPfadTextBox;
        private System.Windows.Forms.Button CSVPfadSuchenButton;
        private System.Windows.Forms.Label StatusLabel;
        private System.Windows.Forms.ProgressBar Ladebalken;
        private System.Windows.Forms.TextBox TestEmailAdresseTextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button TestButton;
    }
}

