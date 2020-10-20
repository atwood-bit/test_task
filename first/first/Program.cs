using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace first
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        static void Main(string[] args)
        {
            Word.Application word = new Word.Application();
            var exePath = AppDomain.CurrentDomain.BaseDirectory;
            object path = Path.Combine(exePath, "Docs\\testDoc.rtf");
            Word.Document doc = word.Documents.Open(ref path);
            string pathExcel = Path.Combine(exePath, "Docs\\testDoc.xlsx");
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Process appProcess = GetExcelProcess(excelapp);
            int CurrCol = 1;
            for (int i = 1; i <= doc.Tables[1].Rows.Count; i++)
            {
                for (int j = 1; j < doc.Tables[1].Columns.Count; j++)
                {
                    string s = doc.Tables[1].Cell(i, j).Range.Text;
                    s = s.Replace("\r\a", string.Empty);
                    switch (s)
                    {
                        case "Регистрационный номер сделки":
                        case "Номер договора":
                        case "Счет контрагента":
                        case "Адрес контрагента":
                        case "Наименование договора":
                            try
                            {
                                worksheet.Rows[1].Columns[CurrCol] = s;
                                string value = doc.Tables[1].Cell(i, j + 1).Range.Text;
                                value = value.Replace("\r\a", string.Empty);
                                worksheet.Rows[2].Columns[CurrCol] = value;
                                worksheet.Columns.EntireColumn.AutoFit();
                                worksheet.Cells.NumberFormat = "#";
                                ++CurrCol;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            break;
                    }
                }
            }
            word.Quit();
            workbook.SaveAs(pathExcel);
            excelapp.Visible = true;
            Console.WriteLine("Файл Excel сохранен.");
            Console.ReadLine();
            excelapp.Application.Quit();
            appProcess.Kill();
        }
        static Process GetExcelProcess(Excel.Application excelApp)
        {
            GetWindowThreadProcessId(excelApp.Hwnd, out int id);
            return Process.GetProcessById(id);
        }
    }
}
