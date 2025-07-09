namespace ExcelSheetAutoPrinter
{
	partial class frmMain
	{
		/// <summary>
		/// 필수 디자이너 변수입니다.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 사용 중인 모든 리소스를 정리합니다.
		/// </summary>
		/// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form 디자이너에서 생성한 코드

		/// <summary>
		/// 디자이너 지원에 필요한 메서드입니다. 
		/// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnLoadExcel = new System.Windows.Forms.Button();
			this.txtSrcFilePath = new System.Windows.Forms.TextBox();
			this.btnFileSelect = new System.Windows.Forms.Button();
			this.lblSrcFile = new System.Windows.Forms.Label();
			this.txtDestFilePath = new System.Windows.Forms.TextBox();
			this.lblDestFile = new System.Windows.Forms.Label();
			this.gvExcel = new System.Windows.Forms.DataGridView();
			this.btnScheduleStart = new System.Windows.Forms.Button();
			this.btnScheduleStop = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.gvExcel)).BeginInit();
			this.SuspendLayout();
			// 
			// btnLoadExcel
			// 
			this.btnLoadExcel.Location = new System.Drawing.Point(16, 93);
			this.btnLoadExcel.Name = "btnLoadExcel";
			this.btnLoadExcel.Size = new System.Drawing.Size(163, 23);
			this.btnLoadExcel.TabIndex = 0;
			this.btnLoadExcel.Text = "엑셀불러오기, PDF저장";
			this.btnLoadExcel.UseVisualStyleBackColor = true;
			this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
			// 
			// txtSrcFilePath
			// 
			this.txtSrcFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtSrcFilePath.Location = new System.Drawing.Point(114, 39);
			this.txtSrcFilePath.Name = "txtSrcFilePath";
			this.txtSrcFilePath.Size = new System.Drawing.Size(942, 21);
			this.txtSrcFilePath.TabIndex = 1;
			// 
			// btnFileSelect
			// 
			this.btnFileSelect.Location = new System.Drawing.Point(14, 10);
			this.btnFileSelect.Name = "btnFileSelect";
			this.btnFileSelect.Size = new System.Drawing.Size(165, 23);
			this.btnFileSelect.TabIndex = 2;
			this.btnFileSelect.Text = "파일 선택";
			this.btnFileSelect.UseVisualStyleBackColor = true;
			this.btnFileSelect.Click += new System.EventHandler(this.btnFileSelect_Click);
			// 
			// lblSrcFile
			// 
			this.lblSrcFile.AutoSize = true;
			this.lblSrcFile.Location = new System.Drawing.Point(11, 42);
			this.lblSrcFile.Name = "lblSrcFile";
			this.lblSrcFile.Size = new System.Drawing.Size(97, 12);
			this.lblSrcFile.TabIndex = 3;
			this.lblSrcFile.Text = "엑셀 파일 경로 : ";
			// 
			// txtDestFilePath
			// 
			this.txtDestFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtDestFilePath.Location = new System.Drawing.Point(114, 66);
			this.txtDestFilePath.Name = "txtDestFilePath";
			this.txtDestFilePath.Size = new System.Drawing.Size(942, 21);
			this.txtDestFilePath.TabIndex = 4;
			// 
			// lblDestFile
			// 
			this.lblDestFile.AutoSize = true;
			this.lblDestFile.Location = new System.Drawing.Point(11, 69);
			this.lblDestFile.Name = "lblDestFile";
			this.lblDestFile.Size = new System.Drawing.Size(96, 12);
			this.lblDestFile.TabIndex = 5;
			this.lblDestFile.Text = "PDF 파일 경로 : ";
			// 
			// gvExcel
			// 
			this.gvExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.gvExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.gvExcel.Location = new System.Drawing.Point(16, 122);
			this.gvExcel.Name = "gvExcel";
			this.gvExcel.RowTemplate.Height = 23;
			this.gvExcel.Size = new System.Drawing.Size(1040, 356);
			this.gvExcel.TabIndex = 6;
			// 
			// btnScheduleStart
			// 
			this.btnScheduleStart.Location = new System.Drawing.Point(14, 484);
			this.btnScheduleStart.Name = "btnScheduleStart";
			this.btnScheduleStart.Size = new System.Drawing.Size(75, 23);
			this.btnScheduleStart.TabIndex = 7;
			this.btnScheduleStart.Text = "시작";
			this.btnScheduleStart.UseVisualStyleBackColor = true;
			this.btnScheduleStart.Click += new System.EventHandler(this.btnScheduleStart_Click);
			// 
			// btnScheduleStop
			// 
			this.btnScheduleStop.Enabled = false;
			this.btnScheduleStop.Location = new System.Drawing.Point(95, 484);
			this.btnScheduleStop.Name = "btnScheduleStop";
			this.btnScheduleStop.Size = new System.Drawing.Size(75, 23);
			this.btnScheduleStop.TabIndex = 8;
			this.btnScheduleStop.Text = "중지";
			this.btnScheduleStop.UseVisualStyleBackColor = true;
			this.btnScheduleStop.Click += new System.EventHandler(this.btnScheduleStop_Click);
			// 
			// frmMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1070, 762);
			this.Controls.Add(this.btnScheduleStop);
			this.Controls.Add(this.btnScheduleStart);
			this.Controls.Add(this.gvExcel);
			this.Controls.Add(this.lblDestFile);
			this.Controls.Add(this.txtDestFilePath);
			this.Controls.Add(this.lblSrcFile);
			this.Controls.Add(this.btnFileSelect);
			this.Controls.Add(this.txtSrcFilePath);
			this.Controls.Add(this.btnLoadExcel);
			this.Name = "frmMain";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Excel Sheet Auto Printer";
			((System.ComponentModel.ISupportInitialize)(this.gvExcel)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnLoadExcel;
		private System.Windows.Forms.TextBox txtSrcFilePath;
		private System.Windows.Forms.Button btnFileSelect;
		private System.Windows.Forms.Label lblSrcFile;
		private System.Windows.Forms.TextBox txtDestFilePath;
		private System.Windows.Forms.Label lblDestFile;
		private System.Windows.Forms.DataGridView gvExcel;
		private System.Windows.Forms.Button btnScheduleStart;
		private System.Windows.Forms.Button btnScheduleStop;
	}
}

