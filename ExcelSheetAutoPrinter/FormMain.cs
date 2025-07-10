using BaiqiSoft.HtmlEditorControl;
using log4net;
using Quartz;
using Quartz.Impl;
using Quartz.Logging;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetAutoPrinter
{
	public partial class frmMain : Form
	{
        private static readonly ILog logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		private string pdfFileFullPath = string.Empty; // PDF 파일 전체 경로
		private string pdfFileName = string.Empty; // PDF 파일 이름

		PrintDocument printDoc;		// winform PrintDocument
        MstHtmlEditor htmlEditor;	// HtmlEditor (NuGet 패키지 추가 : BaiqiSoft.WinFormsHtmlEditor.NET4)
        HtmlToImage hToi;			// Html을 Image로 Convert (NuGet 패키지 추가 : Select.HtmlToPdf)
            
		StdSchedulerFactory factory = null;
		IScheduler scheduler = null;
        List<JobInfo> jobList = null;

		public frmMain()
		{
			InitializeComponent();

			InitForm();
		}

		private void InitForm()
		{
			//CheckForIllegalCrossThreadCalls = false;

			printDoc = new PrintDocument();
            printDoc.PrintPage += PrintDoc_PrintPage;
 
            // Html Editor을 Panel에 Panel에 Add (화면에 뿌려 줌)
            htmlEditor = new MstHtmlEditor();
            htmlEditor.Dock = DockStyle.Fill; // 화면 사이즈에 맞춰서 크기가 변경 될 수 있게 부모 컨테이너에 Docking 시킴
            //this.Controls.Add(htmlEditor);
 
            hToi = new HtmlToImage();
		}

		/// <summary>
        /// Print Page Set Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            Image img = hToi.ConvertHtmlString(htmlEditor.BodyHTML);  // Html을 Image로 변환
 
            Graphics g = e.Graphics;
            g.DrawImage(img, 10, 10);   // x, y 시작점
 
            /*
            // Draw String에 대한 구문
            PointF drawPoint = new PointF(100, 100);            
            // 2중 using 문 사용.
            using (Font font = new Font("Lucida Console", 30))
            using (SolidBrush drawBrush = new SolidBrush(Color.Black))
            {
                g.DrawString("Hello,\n printer", font, drawBrush, drawPoint);
            }
            */
        }

		private void btnLoadExcel_Click(object sender, EventArgs e)
		{
			LoadExcelFile(txtSrcFilePath.Text);
		}

		private void LoadExcelFile(string text)
		{
			Excel.Application excelApp = null;
			Excel.Workbook workBook = null;
			Excel.Worksheet workSheet = null;
			string path = txtSrcFilePath.Text;

            try
			{
				excelApp = new Excel.Application();														// 엑셀 어플리케이션 생성
                workBook = excelApp.Workbooks.Open(path);												// 워크북 열기
                workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet;							// 엑셀 첫번째 워크시트 가져오기

				Excel.Range range = workSheet.UsedRange;												// 사용중인 셀 범위를 가져오기

				for (int columnNo = 1; columnNo <= range.Columns.Count; columnNo++)						// 가져온 열 만큼 반복
				{
					string strColumnName = (string)(range.Cells[1, columnNo] as Excel.Range).Value2;	// 첫번째 행의 셀 값 가져오기

					gvExcel.Columns.Add(strColumnName, strColumnName);									// 데이터 그리드뷰에 열 추가
				}

				for (int rowNo = 2; rowNo <= range.Rows.Count; rowNo++)									// 가져온 행 만큼 반복
                {
					DataGridViewRow row = new DataGridViewRow();										// 데이터 그리드뷰 행 생성
					for (int columnNo = 1; columnNo <= range.Columns.Count; columnNo++)					// 가져온 열 만큼 반복
                    {
                        string str = (string)(range.Cells[rowNo, columnNo] as Excel.Range).Value2;		// 셀 데이터 가져옴

						Excel.Range cell = range.Cells[rowNo, columnNo] as Excel.Range;					// 셀 객체 가져오기
						DataGridViewTextBoxCell cellControl = new DataGridViewTextBoxCell();			// 데이터 그리드뷰 텍스트 박스 셀 생성
						cellControl.Value = (cell.Value == null) ? "" : cell.Value.ToString();			// 셀 값 설정
						row.Cells.Add(cellControl);														// 행에 셀 추가
					}

					gvExcel.Rows.Add(row); // 데이터 그리드뷰에 행 추가
                }

				workSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;					// 페이지 방향을 가로로 설정
				workSheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Path.Combine(Path.GetDirectoryName(path), pdfFileName)); // 워크시트 PDF로 저장

				workBook.Close(true);   // 워크북 닫기
                excelApp.Quit();        // 엑셀 어플리케이션 종료

				logger.Info("Excel file loaded and PDF created successfully.");
			}
			catch (Exception ex)
			{
				logger.Error(ex);
			}
			finally
			{
				ReleaseObject(workSheet);
                ReleaseObject(workBook);
                ReleaseObject(excelApp);
			}
		}

		/// <summary>
        /// 액셀 객체 해제 메소드
        /// </summary>
        /// <param name="obj"></param>
        private void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);  // 액셀 객체 해제
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();                       // 가비지 수집
            }
        }

		private void btnFileSelect_Click(object sender, EventArgs e)
		{
			try
			{
				OpenFileDialog ofd = new OpenFileDialog();

				ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb|All Files|*.*";	// 엑셀 파일 필터
				ofd.ShowDialog();														// 파일 선택 대화창 표시

				if (ofd.FileName != string.Empty)										// 파일이 선택되었으면
				{
					txtSrcFilePath.Text = ofd.FileName;									// 텍스트 박스에 파일 경로 표시

					pdfFileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf";				// PDF 파일 이름 설정
					pdfFileFullPath = Path.Combine(Path.GetDirectoryName(ofd.FileName), pdfFileName);	// PDF 파일 전체 경로 설정

					txtDestFilePath.Text = Path.Combine(Path.GetDirectoryName(ofd.FileName), pdfFileName); // PDF 파일 경로 설정
				}

				logger.Info("Selected file: " + txtSrcFilePath.Text); // 선택한 파일 경로 로그 출력
			}
			catch (Exception ex)
			{
				logger.Error(ex);
			}
		}

		private async void btnScheduleStart_Click(object sender, EventArgs e)
		{
			await Task.Run(() => ScheduleStart());	// 스케줄 시작 비동기 실행
		}

		private async void btnScheduleStop_Click(object sender, EventArgs e)
		{
			await Task.Run(() => ScheduleStop());	// 스케줄 중지 비동기 실행
		}

		private async void ScheduleStart()
		{
			if (this.btnScheduleStart.InvokeRequired)	// UI 스레드가 아닌 경우 Invoke 호출
			{
				this.btnScheduleStart.Invoke(new Action(delegate()
				{
					this.btnScheduleStart.Enabled = false;	// 스케줄 시작 버튼 비활성화
				}
				));
			}
			else
			{
				this.btnScheduleStart.Enabled = false;  // 스케줄 시작 버튼 비활성화
			}

			if (this.btnScheduleStop.InvokeRequired)	// UI 스레드가 아닌 경우 Invoke 호출
			{
				this.btnScheduleStop.Invoke(new Action(delegate()
				{
					this.btnScheduleStop.Enabled = true;	// 스케줄 중지 버튼 활성화
				}
				));
			}
			else
			{
				this.btnScheduleStop.Enabled = true;   // 스케줄 중지 버튼 활성화
			}

			// 스케줄 시작 로직 (예: Quartz 스케줄러 사용 등)
			//LogProvider.SetCurrentLogProvider(new ConsoleLogProvider());

            factory = new StdSchedulerFactory();
			scheduler = await factory.GetScheduler();

            // Job 목록 생성
			jobList = new List<JobInfo>();

            jobList.Add(new JobInfo() { Key = "1", CronExpression = "0/5 * * * * ?", StartTime = DateTime.Now, EndTime = DateTime.Now.AddSeconds(30) });
            //jobList.Add(new JobInfo() { Key = "2", CronExpression = "0/10 * * * * ?", StartTime = DateTime.Now, EndTime = DateTime.Now.AddSeconds(30) });
            //jobList.Add(new JobInfo() { Key = "3", CronExpression = "0/15 * * * * ?", StartTime = DateTime.Now, EndTime = DateTime.Now.AddSeconds(30) });

            foreach (var job in jobList)
            {
                // Job 정의
                IJobDetail jobdetail = JobBuilder.Create<TestJob>()
                             .WithIdentity(job.Key)
                             .Build();

                // Job 주기 정의
                ITrigger trigger = TriggerBuilder.Create()
                                    .WithIdentity($"{job.Key}_trigger")
                                    .StartNow()
                                    .WithCronSchedule(job.CronExpression)
                                    .Build();

                // Scheduler 에 Job 추가
                await scheduler.ScheduleJob(jobdetail, trigger);
            }

            // Scheduler 시작
            await scheduler.Start();
		}

		private async void ScheduleStop()
		{
			if (this.btnScheduleStart.InvokeRequired)	// UI 스레드가 아닌 경우 Invoke 호출
			{
				this.btnScheduleStart.Invoke(new Action(delegate()
				{
					this.btnScheduleStart.Enabled = true;	// 스케줄 시작 버튼 활성화
				}
				));
			}
			else
			{
				this.btnScheduleStart.Enabled = true;  // 스케줄 시작 버튼 활성화
			}

			if (this.btnScheduleStop.InvokeRequired)	// UI 스레드가 아닌 경우 Invoke 호출
			{
				this.btnScheduleStop.Invoke(new Action(delegate()
				{
					this.btnScheduleStop.Enabled = false;   // 스케줄 중지 버튼 비활성화
				}
				));
			}
			else
			{
				this.btnScheduleStop.Enabled = false;   // 스케줄 중지 버튼 비활성화
			}

			// Scheduler 시작
			// https://cwkcw.tistory.com/450
            await scheduler.Shutdown();		// 스케줄러 종료
		}

		private void btnPrint_Click(object sender, EventArgs e)
		{
			// https://cwkcw.tistory.com/645
			try
			{
				logger.Info("Print Start");

				PrintDialog printDlg = new PrintDialog();
 
            		if (printDlg.ShowDialog() == DialogResult.OK)
            		{
                		printDoc.PrinterSettings = printDlg.PrinterSettings;
                		printDoc.Print();
            		}

				logger.Info("Print End");
			}
			catch (Exception ex)
			{
				logger.Error(ex);
			}
		}

		// bind this method to its TextChanged event handler:
		// richTextBox.TextChanged += richTextBox_TextChanged;
		private void txtLog_TextChanged(object sender, EventArgs e)
		{
			// set the current caret position to the end
			txtLog.SelectionStart = txtLog.Text.Length;

			// scroll it automatically
			txtLog.ScrollToCaret();
		}
	}
}
