using ExcelDataReader;
using log4net;
using System;
using System.Data;
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

		private string pdfFileName = string.Empty; // PDF 파일 이름

		public frmMain()
		{
			InitializeComponent();

			InitForm();
		}

		private void InitForm()
		{
			//CheckForIllegalCrossThreadCalls = false;
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

					pdfFileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"; // PDF 파일 이름 설정

					txtDestFilePath.Text = Path.Combine(Path.GetDirectoryName(ofd.FileName), pdfFileName); // PDF 파일 경로 설정
				}
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

		private void ScheduleStart()
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
		}

		private void ScheduleStop()
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
		}
	}
}
