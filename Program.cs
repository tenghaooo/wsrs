using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System;

namespace wsrs
{
    class Program
    {
       static void Main(string[] args)
        {

            
            var excelApp = new Excel.Application();
            excelApp.Visible = false;

            // 開啟驗證結果Excel表
            Excel.Workbook excelBook = excelApp.Workbooks.Open("D:\\TemplateFiles\\sample.xlsx");

            // 設定CaseInfo
            Excel.Worksheet caseInfoSheet = excelBook.Worksheets["caseinfo"];
            CaseInfo caseinfo = new CaseInfo();
            caseinfo = getCaseInfo(caseInfoSheet);

            excelBook.Close();
            excelApp.Quit();

            Console.WriteLine(caseinfo.author);
            Console.ReadLine();
        }

        static CaseInfo getCaseInfo(Excel.Worksheet sheet)
        {
            var caseinfo = new CaseInfo();
            caseinfo.userName = sheet.Cells[2, "A"].Value.ToString();
            caseinfo.projectName = sheet.Cells[2, "B"].Value.ToString();
            caseinfo.reportName = sheet.Cells[2, "C"].Value.ToString();
            caseinfo.period = sheet.Cells[2, "D"].Value.ToString();
            caseinfo.reportCode = sheet.Cells[2, "E"].Value.ToString();
            caseinfo.author = sheet.Cells[2, "F"].Value.ToString();
            caseinfo.tool = sheet.Cells[2, "G"].Value.ToString();
            caseinfo.year = sheet.Cells[2, "H"].Value.ToString();
            caseinfo.startDate = sheet.Cells[2, "I"].Value.ToString();
            caseinfo.endDate = sheet.Cells[2, "J"].Value.ToString();
            caseinfo.level = sheet.Cells[2, "K"].Value.ToString();
            return caseinfo;
        }
    }
}
