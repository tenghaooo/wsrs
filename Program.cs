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
            var wordApp = new Word.Application();

            excelApp.Visible = false;
            wordApp.Visible = false;

            // open result excel table
            Excel.Workbook excelBook = excelApp.Workbooks.Open("D:\\TemplateFiles\\sample.xlsx");
            // open report template
            Word.Document report = wordApp.Documents.Open("D:\\TemplateFiles\\sample.docx");

            // set case info
            Excel.Worksheet caseInfoSheet = excelBook.Worksheets["caseinfo"];
            CaseInfo caseinfo = new CaseInfo();
            caseinfo = getCaseInfo(caseInfoSheet);

            // write caseinfo to report
            report.Content.Find.Execute("p_userName", false, false, false, false, false, true, 1, false, caseinfo.userName, 2, false, false, false, false);
            report.Content.Find.Execute("p_projectName", false, false, false, false, false, true, 1, false, caseinfo.projectName, 2, false, false, false, false);
            report.Content.Find.Execute("p_reportName", false, false, false, false, false, true, 1, false, caseinfo.reportName, 2, false, false, false, false);
            report.Content.Find.Execute("p_period", false, false, false, false, false, true, 1, false, caseinfo.period, 2, false, false, false, false);
            report.Content.Find.Execute("p_reportCode", false, false, false, false, false, true, 1, false, caseinfo.reportCode, 2, false, false, false, false);
            report.Content.Find.Execute("p_author", false, false, false, false, false, true, 1, false, caseinfo.author, 2, false, false, false, false);
            report.Content.Find.Execute("p_year", false, false, false, false, false, true, 1, false, caseinfo.year, 2, false, false, false, false);
            report.Content.Find.Execute("p_startDate", false, false, false, false, false, true, 1, false, caseinfo.startDate, 2, false, false, false, false);
            report.Content.Find.Execute("p_endDate", false, false, false, false, false, true, 1, false, caseinfo.endDate, 2, false, false, false, false);
            report.Content.Find.Execute("p_tool", false, false, false, false, false, true, 1, false, caseinfo.tool, 2, false, false, false, false);
            // write caseinfo to header
            foreach (Word.Section section in report.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Find.Execute("p_userName", false, false, false, false, false, true, 1, false, caseinfo.userName, 2, false, false, false, false);
                headerRange.Find.Execute("p_projectName", false, false, false, false, false, true, 1, false, caseinfo.projectName, 2, false, false, false, false);
                headerRange.Find.Execute("p_reportName", false, false, false, false, false, true, 1, false, caseinfo.reportName, 2, false, false, false, false);
                headerRange.Find.Execute("p_period", false, false, false, false, false, true, 1, false, caseinfo.period, 2, false, false, false, false);
            }




            report.SaveAs2("D:\\TemplateFiles\\test.docx");


            excelBook.Close();
            excelApp.Quit();
            report.Close();
            wordApp.Quit();

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
