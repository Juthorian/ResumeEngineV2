namespace ResumeEngineV2
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txtBoxKeyword = new System.Windows.Forms.TextBox();
            this.lblEnterKeyword = new System.Windows.Forms.Label();
            this.btnKeywordSubmit = new System.Windows.Forms.Button();
            this.lblResults = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.lblLogin = new System.Windows.Forms.Label();
            this.textBoxUsername = new System.Windows.Forms.TextBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.btnLoginSubmit = new System.Windows.Forms.Button();
            this.btnLogout = new System.Windows.Forms.Button();
            this.lblUsername = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.resultsView = new System.Windows.Forms.DataGridView();
            this.FileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Percent = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblAddTextBox = new System.Windows.Forms.Label();
            this.cmbWeight = new System.Windows.Forms.ComboBox();
            this.txtBoxExperience = new System.Windows.Forms.TextBox();
            this.lblExperience = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.resultsView)).BeginInit();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(402, 48);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(436, 321);
            this.progressBar1.TabIndex = 0;
            // 
            // txtBoxKeyword
            // 
            this.txtBoxKeyword.Location = new System.Drawing.Point(30, 61);
            this.txtBoxKeyword.Name = "txtBoxKeyword";
            this.txtBoxKeyword.Size = new System.Drawing.Size(180, 20);
            this.txtBoxKeyword.TabIndex = 2;
            // 
            // lblEnterKeyword
            // 
            this.lblEnterKeyword.AutoSize = true;
            this.lblEnterKeyword.Location = new System.Drawing.Point(61, 30);
            this.lblEnterKeyword.Name = "lblEnterKeyword";
            this.lblEnterKeyword.Size = new System.Drawing.Size(121, 13);
            this.lblEnterKeyword.TabIndex = 0;
            this.lblEnterKeyword.Text = "Please enter a keyword:";
            // 
            // btnKeywordSubmit
            // 
            this.btnKeywordSubmit.Location = new System.Drawing.Point(302, 61);
            this.btnKeywordSubmit.Name = "btnKeywordSubmit";
            this.btnKeywordSubmit.Size = new System.Drawing.Size(75, 23);
            this.btnKeywordSubmit.TabIndex = 6;
            this.btnKeywordSubmit.Text = "Submit";
            this.btnKeywordSubmit.UseVisualStyleBackColor = true;
            this.btnKeywordSubmit.Click += new System.EventHandler(this.btnKeywordSubmit_Click);
            // 
            // lblResults
            // 
            this.lblResults.AutoSize = true;
            this.lblResults.Location = new System.Drawing.Point(399, 18);
            this.lblResults.Name = "lblResults";
            this.lblResults.Size = new System.Drawing.Size(45, 13);
            this.lblResults.TabIndex = 12;
            this.lblResults.Text = "Results:";
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWoker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_Completed);
            // 
            // lblLogin
            // 
            this.lblLogin.AutoSize = true;
            this.lblLogin.Location = new System.Drawing.Point(27, 119);
            this.lblLogin.Name = "lblLogin";
            this.lblLogin.Size = new System.Drawing.Size(33, 13);
            this.lblLogin.TabIndex = 0;
            this.lblLogin.Text = "Login";
            this.lblLogin.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // textBoxUsername
            // 
            this.textBoxUsername.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.textBoxUsername.Location = new System.Drawing.Point(30, 146);
            this.textBoxUsername.Name = "textBoxUsername";
            this.textBoxUsername.Size = new System.Drawing.Size(180, 20);
            this.textBoxUsername.TabIndex = 3;
            this.textBoxUsername.Text = "Example: jbraham@aecon.com";
            this.textBoxUsername.Enter += new System.EventHandler(this.textBoxUsername_Enter);
            this.textBoxUsername.Leave += new System.EventHandler(this.textBoxUsername_Leave);
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(30, 195);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(180, 20);
            this.textBoxPassword.TabIndex = 5;
            // 
            // btnLoginSubmit
            // 
            this.btnLoginSubmit.Location = new System.Drawing.Point(135, 114);
            this.btnLoginSubmit.Name = "btnLoginSubmit";
            this.btnLoginSubmit.Size = new System.Drawing.Size(75, 23);
            this.btnLoginSubmit.TabIndex = 1;
            this.btnLoginSubmit.Text = "Submit";
            this.btnLoginSubmit.UseVisualStyleBackColor = true;
            this.btnLoginSubmit.Click += new System.EventHandler(this.btnLoginSubmit_Click);
            // 
            // btnLogout
            // 
            this.btnLogout.Location = new System.Drawing.Point(12, 346);
            this.btnLogout.Name = "btnLogout";
            this.btnLogout.Size = new System.Drawing.Size(75, 23);
            this.btnLogout.TabIndex = 9;
            this.btnLogout.Text = "Logout";
            this.btnLogout.UseVisualStyleBackColor = true;
            this.btnLogout.Click += new System.EventHandler(this.btnLogout_Click);
            // 
            // lblUsername
            // 
            this.lblUsername.AutoSize = true;
            this.lblUsername.Location = new System.Drawing.Point(9, 149);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(15, 13);
            this.lblUsername.TabIndex = 2;
            this.lblUsername.Text = "U";
            this.lblUsername.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(9, 198);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(14, 13);
            this.lblPassword.TabIndex = 4;
            this.lblPassword.Text = "P";
            this.lblPassword.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // resultsView
            // 
            this.resultsView.AllowUserToAddRows = false;
            this.resultsView.AllowUserToDeleteRows = false;
            this.resultsView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.resultsView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FileName,
            this.Percent});
            this.resultsView.Location = new System.Drawing.Point(402, 48);
            this.resultsView.Name = "resultsView";
            this.resultsView.ReadOnly = true;
            this.resultsView.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.resultsView.RowHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.resultsView.RowHeadersWidth = 71;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            this.resultsView.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.resultsView.Size = new System.Drawing.Size(436, 321);
            this.resultsView.TabIndex = 12;
            this.resultsView.TabStop = false;
            this.resultsView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.resultsView_CellClick);
            // 
            // FileName
            // 
            this.FileName.HeaderText = "File Name";
            this.FileName.Name = "FileName";
            this.FileName.ReadOnly = true;
            this.FileName.Width = 260;
            // 
            // Percent
            // 
            this.Percent.HeaderText = "Percent Match";
            this.Percent.Name = "Percent";
            this.Percent.ReadOnly = true;
            this.Percent.Width = 85;
            // 
            // lblAddTextBox
            // 
            this.lblAddTextBox.AutoSize = true;
            this.lblAddTextBox.Location = new System.Drawing.Point(9, 61);
            this.lblAddTextBox.Name = "lblAddTextBox";
            this.lblAddTextBox.Size = new System.Drawing.Size(13, 13);
            this.lblAddTextBox.TabIndex = 1;
            this.lblAddTextBox.Text = "+";
            this.lblAddTextBox.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblAddTextBox.Click += new System.EventHandler(this.lblAddTextBox_Click);
            // 
            // cmbWeight
            // 
            this.cmbWeight.DisplayMember = "1";
            this.cmbWeight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWeight.Enabled = false;
            this.cmbWeight.FormattingEnabled = true;
            this.cmbWeight.ItemHeight = 13;
            this.cmbWeight.Items.AddRange(new object[] {
            "90%",
            "80%",
            "70%",
            "60%",
            "50%",
            "40%",
            "30%",
            "20%",
            "10%",
            "100%"});
            this.cmbWeight.Location = new System.Drawing.Point(220, 60);
            this.cmbWeight.Name = "cmbWeight";
            this.cmbWeight.Size = new System.Drawing.Size(58, 21);
            this.cmbWeight.TabIndex = 5;
            this.cmbWeight.ValueMember = "1";
            // 
            // txtBoxExperience
            // 
            this.txtBoxExperience.Location = new System.Drawing.Point(220, 91);
            this.txtBoxExperience.Name = "txtBoxExperience";
            this.txtBoxExperience.Size = new System.Drawing.Size(37, 20);
            this.txtBoxExperience.TabIndex = 7;
            this.txtBoxExperience.Text = "0";
            // 
            // lblExperience
            // 
            this.lblExperience.AutoSize = true;
            this.lblExperience.Location = new System.Drawing.Point(263, 94);
            this.lblExperience.Name = "lblExperience";
            this.lblExperience.Size = new System.Drawing.Size(122, 13);
            this.lblExperience.TabIndex = 8;
            this.lblExperience.Text = "Min Years of Experience";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(875, 387);
            this.Controls.Add(this.lblExperience);
            this.Controls.Add(this.txtBoxExperience);
            this.Controls.Add(this.cmbWeight);
            this.Controls.Add(this.lblAddTextBox);
            this.Controls.Add(this.resultsView);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUsername);
            this.Controls.Add(this.btnLogout);
            this.Controls.Add(this.btnLoginSubmit);
            this.Controls.Add(this.textBoxPassword);
            this.Controls.Add(this.textBoxUsername);
            this.Controls.Add(this.lblLogin);
            this.Controls.Add(this.lblResults);
            this.Controls.Add(this.btnKeywordSubmit);
            this.Controls.Add(this.lblEnterKeyword);
            this.Controls.Add(this.txtBoxKeyword);
            this.Controls.Add(this.progressBar1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Resume Search Engine";
            ((System.ComponentModel.ISupportInitialize)(this.resultsView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox txtBoxKeyword;
        private System.Windows.Forms.Label lblEnterKeyword;
        private System.Windows.Forms.Button btnKeywordSubmit;
        private System.Windows.Forms.Label lblResults;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label lblLogin;
        private System.Windows.Forms.TextBox textBoxUsername;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Button btnLoginSubmit;
        private System.Windows.Forms.Button btnLogout;
        private System.Windows.Forms.Label lblUsername;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.DataGridView resultsView;
        private System.Windows.Forms.DataGridViewTextBoxColumn FileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Percent;
        private System.Windows.Forms.Label lblAddTextBox;
        private System.Windows.Forms.ComboBox cmbWeight;
        private System.Windows.Forms.TextBox txtBoxExperience;
        private System.Windows.Forms.Label lblExperience;
    }
}

