namespace KR_DailyReport
{
    partial class FormMain
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
            cmbStartTimeStamp = new ComboBox();
            lblParamName = new Label();
            lblValue = new Label();
            btnExcelOutput = new Button();
            lblStatus = new Label();
            dgvDailyData = new DataGridView();
            dtpDate = new DateTimePicker();
            btnGetData = new Button();
            groupBox1 = new GroupBox();
            ((System.ComponentModel.ISupportInitialize)dgvDailyData).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // cmbStartTimeStamp
            // 
            cmbStartTimeStamp.FormattingEnabled = true;
            cmbStartTimeStamp.Location = new Point(84, 53);
            cmbStartTimeStamp.Margin = new Padding(2);
            cmbStartTimeStamp.Name = "cmbStartTimeStamp";
            cmbStartTimeStamp.Size = new Size(117, 23);
            cmbStartTimeStamp.TabIndex = 1;
            cmbStartTimeStamp.SelectedIndexChanged += cmbItemName1_SelectedIndexChanged;
            // 
            // lblParamName
            // 
            lblParamName.AutoSize = true;
            lblParamName.Location = new Point(17, 31);
            lblParamName.Margin = new Padding(2, 0, 2, 0);
            lblParamName.Name = "lblParamName";
            lblParamName.Size = new Size(55, 15);
            lblParamName.TabIndex = 2;
            lblParamName.Text = "指定日付";
            // 
            // lblValue
            // 
            lblValue.AutoSize = true;
            lblValue.Location = new Point(17, 56);
            lblValue.Margin = new Padding(2, 0, 2, 0);
            lblValue.Name = "lblValue";
            lblValue.Size = new Size(55, 15);
            lblValue.TabIndex = 2;
            lblValue.Text = "開始時刻";
            // 
            // btnExcelOutput
            // 
            btnExcelOutput.Location = new Point(364, 52);
            btnExcelOutput.Name = "btnExcelOutput";
            btnExcelOutput.Size = new Size(95, 23);
            btnExcelOutput.TabIndex = 6;
            btnExcelOutput.Text = "EXCELへ出力";
            btnExcelOutput.UseVisualStyleBackColor = true;
            btnExcelOutput.Click += btnExcelOutput_Click;
            // 
            // lblStatus
            // 
            lblStatus.Location = new Point(21, 756);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(175, 13);
            lblStatus.TabIndex = 8;
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // dgvDailyData
            // 
            dgvDailyData.AllowUserToAddRows = false;
            dgvDailyData.AllowUserToDeleteRows = false;
            dgvDailyData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDailyData.Location = new Point(21, 116);
            dgvDailyData.Name = "dgvDailyData";
            dgvDailyData.ReadOnly = true;
            dgvDailyData.RowTemplate.Height = 25;
            dgvDailyData.Size = new Size(480, 626);
            dgvDailyData.TabIndex = 9;
            // 
            // dtpDate
            // 
            dtpDate.Format = DateTimePickerFormat.Short;
            dtpDate.Location = new Point(84, 25);
            dtpDate.Name = "dtpDate";
            dtpDate.Size = new Size(117, 23);
            dtpDate.TabIndex = 10;
            dtpDate.ValueChanged += dtpDate_ValueChanged;
            // 
            // btnGetData
            // 
            btnGetData.Location = new Point(218, 53);
            btnGetData.Name = "btnGetData";
            btnGetData.Size = new Size(95, 23);
            btnGetData.TabIndex = 6;
            btnGetData.Text = "データ抽出";
            btnGetData.UseVisualStyleBackColor = true;
            btnGetData.Click += btnGetData_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblParamName);
            groupBox1.Controls.Add(lblValue);
            groupBox1.Controls.Add(dtpDate);
            groupBox1.Controls.Add(btnExcelOutput);
            groupBox1.Controls.Add(cmbStartTimeStamp);
            groupBox1.Controls.Add(btnGetData);
            groupBox1.Location = new Point(21, 16);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(484, 88);
            groupBox1.TabIndex = 11;
            groupBox1.TabStop = false;
            groupBox1.Text = "抽出条件";
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(525, 781);
            Controls.Add(groupBox1);
            Controls.Add(dgvDailyData);
            Controls.Add(lblStatus);
            Margin = new Padding(2);
            Name = "FormMain";
            Text = "日報出力";
            Activated += FormMain_Activated;
            FormClosing += FormMain_FormClosing;
            Load += FormMain_Load;
            ((System.ComponentModel.ISupportInitialize)dgvDailyData).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private ComboBox cmbStartTimeStamp;
        private Label lblParamName;
        private Label lblValue;
        private Button btnExcelOutput;
        private Label lblStatus;
        private DataGridView dgvDailyData;
        private DateTimePicker dtpDate;
        private Button btnGetData;
        private GroupBox groupBox1;
    }
}