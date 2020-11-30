using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;

namespace wsrs
{
    class Program
    {
       static void Main(string[] args)
        {
            Console.WriteLine("[L] Start running...");

            var excelApp = new Excel.Application();
            var wordApp = new Word.Application();

            excelApp.Visible = false;
            wordApp.Visible = false;

            Console.WriteLine("[L] Openning result excel");
            // open result excel table
            Excel.Workbook excelBook = excelApp.Workbooks.Open("D:\\TemplateFiles\\sample.xlsx");
            

            List<Unit> Units = new List<Unit>();

            Console.WriteLine("[L] Loading units and sites");
            // setup units and sites
            Excel.Worksheet targetSheet = excelBook.Worksheets["targets"];
            int i = 2;
            int sumOfSites = 0;
            // loop targets in sheet
            while (targetSheet.Cells[i, "A"].Value != null)
            {
                string currentUnit = targetSheet.Cells[i, "A"].Value.ToString();
                string currentSiteUrl = targetSheet.Cells[i, "B"].Value.ToString();
                string currentSiteName = targetSheet.Cells[i, "C"].Value.ToString();
                // first unit and first site
                if (Units.Count == 0)
                {
                    Unit newUnit = new Unit();
                    newUnit.name = currentUnit;
                    Site newSite = new Site();
                    newSite.url = currentSiteUrl;
                    newSite.name = currentSiteName;
                    newUnit.sites.Add(newSite);
                    Units.Add(newUnit);
                }
                else
                {
                    // check if current unit already exist 
                    bool exist = false;
                    int j = 0;
                    for (; j < Units.Count; j++)
                    {
                        if (currentUnit == Units[j].name)
                        {
                            exist = true;
                            break;
                        }
                    }
                    
                    // if unit already exist, just new a site
                    if (exist)
                    {
                        Site newSite = new Site();
                        newSite.url = currentSiteUrl;
                        newSite.name = currentSiteName;
                        Units[j].sites.Add(newSite);
                    }
                    // if unit not exist, new a unit and site
                    else
                    {
                        Unit newUnit = new Unit();
                        newUnit.name = currentUnit;
                        Site newSite = new Site();
                        newSite.url = currentSiteUrl;
                        newSite.name = currentSiteName;
                        newUnit.sites.Add(newSite);
                        Units.Add(newUnit);
                    }
                }
                i++;
                sumOfSites++;
            }

            Console.WriteLine("[L] There are total " + sumOfSites.ToString() + " sites and " + Units.Count.ToString() + " units");

            Console.WriteLine("[L] Loading vulns");
            // setup vulns
            Excel.Worksheet resultSheet = excelBook.Worksheets["result"];
            int x = 2;
            while (resultSheet.Cells[x, "A"].Value != null)
            {
                string currentUrl = resultSheet.Cells[x, "A"].Value.ToString();
                string currentSiteName = resultSheet.Cells[x, "B"].Value.ToString();
                string currentVulnName = resultSheet.Cells[x, "C"].Value.ToString();
                string currentVulnLevel = resultSheet.Cells[x, "D"].Value.ToString();
                string currentVulnUrl = resultSheet.Cells[x, "E"].Value.ToString();

                // add vuln to units sites
                Vuln newVuln = new Vuln();
                newVuln.name = currentVulnName;
                newVuln.level = currentVulnLevel;
                newVuln.vulnUrl = currentVulnUrl;
                // find current vuln in which unit and site
                for (int y = 0; y < Units.Count; y++)
                {
                    bool found = false;
                    for (int z = 0; z < Units[y].sites.Count; z++)
                    {
                        if (Units[y].sites[z].name == currentSiteName)
                        {
                            Units[y].sites[z].vulns.Add(newVuln);
                            Units[y].sites[z].numOfVulns[currentVulnLevel]++;
                            found = true;
                            break;
                        }
                    }
                    if (found)
                        break;
                }
                x++;
            }

            Console.WriteLine("[L] Loading case info");
            // set case info
            Excel.Worksheet caseInfoSheet = excelBook.Worksheets["caseinfo"];
            CaseInfo caseinfo = new CaseInfo();
            caseinfo = getCaseInfo(caseInfoSheet);


            Console.WriteLine("[L] Creating unit report");
            /*
             * Big Loop For Units, Create Unit Report
             * */

            for (int U = 0; U < Units.Count; U++)
            {

                Console.WriteLine("[L] Creating report " + (U + 1).ToString() + "/" + Units.Count.ToString());
                // open report template
                Word.Document report = wordApp.Documents.Open("D:\\TemplateFiles\\sample.docx");
                string reportPath = "D:\\Reports\\";
                string reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "-" + Units[U].name + "_" + caseinfo.period + ".docx";

                Console.WriteLine("    Writting case info to report");
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

                Console.WriteLine("    Writting case info to header");
                // write caseinfo to header
                foreach (Word.Section section in report.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Find.Execute("p_userName", false, false, false, false, false, true, 1, false, caseinfo.userName, 2, false, false, false, false);
                    headerRange.Find.Execute("p_projectName", false, false, false, false, false, true, 1, false, caseinfo.projectName, 2, false, false, false, false);
                    headerRange.Find.Execute("p_reportName", false, false, false, false, false, true, 1, false, caseinfo.reportName, 2, false, false, false, false);
                    headerRange.Find.Execute("p_period", false, false, false, false, false, true, 1, false, caseinfo.period, 2, false, false, false, false);
                }

                // write table one


                Console.WriteLine("    Saving report");
                report.SaveAs2(reportPath + reportName);
                report.Close();

                Console.WriteLine("    Done");
            }
           

            excelBook.Close();
            excelApp.Quit();
            
            wordApp.Quit();

            Console.WriteLine("[L] Finish!!!");
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
