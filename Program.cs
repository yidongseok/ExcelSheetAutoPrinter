using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetAutoPrinter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== Excel ?쒗듃 PDF 蹂??& 異쒕젰 ?쒖옉 ===");

            string excelFilePath = @"C:\寃쎈줈\?뚯씪.xlsx";
            string sheetName = "?쒗듃?대쫫";
            string pdfOutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output", "ExportedSheet.pdf");
            string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output", "print_log.txt");

            Log(logFile, $"[{DateTime.Now}] PDF 蹂???쒖옉");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;

                workbook = excelApp.Workbooks.Open(excelFilePath);

                Excel.Worksheet worksheet = workbook.Sheets[sheetName] as Excel.Worksheet;

                if (worksheet == null)
                {
                    Log(logFile, $"[?ㅻ쪟] ?쒗듃 '{sheetName}'??瑜? 李얠쓣 ???놁뒿?덈떎.");
                    return;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(pdfOutputPath));

                worksheet.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    pdfOutputPath
                );

                Log(logFile, $"[{DateTime.Now}] PDF ????꾨즺: {pdfOutputPath}");

                Process printProcess = new Process();
                printProcess.StartInfo = new ProcessStartInfo()
                {
                    FileName = pdfOutputPath,
                    Verb = "print",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                printProcess.Start();

                Log(logFile, $"[{DateTime.Now}] PDF 異쒕젰 紐낅졊 ?꾩넚");
            }
            catch (Exception ex)
            {
                Log(logFile, $"[?ㅻ쪟] {ex.Message}");
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                Log(logFile, $"[{DateTime.Now}] Excel 醫낅즺");
            }
        }

        static void Log(string logFile, string message)
        {
            Console.WriteLine(message);
            File.AppendAllText(logFile, message + Environment.NewLine);
        }
    }
}
