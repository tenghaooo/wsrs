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

            printBanner();
            
            var excelApp = new Excel.Application();
            var wordApp = new Word.Application();
            excelApp.Visible = false;
            wordApp.Visible = false;

            List<Unit> Units = new List<Unit>();
            CaseInfo caseinfo = new CaseInfo();

            Console.WriteLine("[L] Start running...");

            // open result excel table
            Console.WriteLine("[L] Openning result excel");
            Excel.Workbook excelBook = excelApp.Workbooks.Open("D:\\TemplateFiles\\sample.xlsx");
            
            // setup units and sites
            Console.WriteLine("[L] Loading units and sites");
            Excel.Worksheet targetSheet = excelBook.Worksheets["targets"];
            setUnitsAndSites(targetSheet, ref Units);

            // setup vulns
            Console.WriteLine("[L] Loading vulns");
            Excel.Worksheet resultSheet = excelBook.Worksheets["result"];
            setVulns(resultSheet, ref Units);

            // set case info
            Console.WriteLine("[L] Loading case info");
            Excel.Worksheet caseInfoSheet = excelBook.Worksheets["caseinfo"];
            setCaseInfo(caseInfoSheet, ref caseinfo);

            /*
             * Big Loop For Units, Create Unit Report
             * */
            Console.WriteLine("[L] Creating unit report");
            for (int U = 0; U < Units.Count; U++)
            {
                Console.WriteLine("[L] Creating report " + (U + 1).ToString() + "/" + Units.Count.ToString());
                // open report template
                Word.Document report = wordApp.Documents.Open("D:\\TemplateFiles\\sample.docx");
                string reportPath = "D:\\Reports\\";
                string reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "_" + caseinfo.period + ".docx"; ;
                if (Units[U].name != "000")
                    reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "-" + Units[U].name + "_" + caseinfo.period + ".docx";
               
                // write caseinfo to report and header
                Console.WriteLine("        Writting caseinfo to report & header");
                writeCaseInfoToReport(ref report, caseinfo, Units[U]);

                // write table one
                Console.WriteLine("        Writting table 1");
                writeTableOneToReport(ref report, caseinfo, Units[U]);
                
                // write table two
                Console.WriteLine("        Writting table 2");
                writeTableTwoToReport(ref report, caseinfo, Units[U]);

                // save report
                Console.WriteLine("        Saving report");
                report.SaveAs2(reportPath + reportName);
                report.Close();
            
                Console.WriteLine("        Done.");
            }
            
            excelBook.Close();
            excelApp.Quit();
            
            wordApp.Quit();

            Console.WriteLine("[L] Finish!!!");
            Console.ReadLine();
        }

        static void writeTableTwoToReport(ref Word.Document report, CaseInfo caseinfo, Unit unit)
        {
            Word.Table tableTwo = report.Tables[4];
            int row = 2;
            for (int i = 0; i < unit.sites.Count; i++)
            {
                for (int j = 0; j < unit.sites[i].vulns.Count; j++)
                {
                    tableTwo.Cell(row, 1).Range.Text = unit.sites[i].url;
                    tableTwo.Cell(row, 2).Range.Text = unit.sites[i].name;
                    tableTwo.Cell(row, 3).Range.Text = unit.sites[i].vulns[j].name;
                    tableTwo.Cell(row, 4).Range.Text = unit.sites[i].vulns[j].level;
                    tableTwo.Cell(row, 5).Range.Text = unit.sites[i].vulns[j].vulnUrl;
                    tableTwo.Rows.Add();
                    row++;
                }
            }
            tableTwo.Rows.Last.Delete();
        }

        static void writeTableOneToReport(ref Word.Document report, CaseInfo caseinfo, Unit unit)
        {
            Word.Table tableOne = report.Tables[3];
            int indexOfSites = 0;
            int row = 3;
            do
            {
                if (indexOfSites > 0)
                {
                    tableOne.Rows.Add();
                    row++;
                }
                tableOne.Cell(row, 1).Range.Text = (indexOfSites + 1).ToString();
                tableOne.Cell(row, 2).Range.Text = unit.sites[indexOfSites].url;
                tableOne.Cell(row, 3).Range.Text = unit.sites[indexOfSites].name;
                tableOne.Cell(row, 4).Range.Text = unit.sites[indexOfSites].numOfLevelVulns["Critical"].ToString();
                tableOne.Cell(row, 5).Range.Text = unit.sites[indexOfSites].numOfLevelVulns["High"].ToString();
                tableOne.Cell(row, 6).Range.Text = unit.sites[indexOfSites].numOfLevelVulns["Medium"].ToString();
                tableOne.Cell(row, 7).Range.Text = unit.sites[indexOfSites].numOfLevelVulns["Low"].ToString();

                indexOfSites++;
            } while (indexOfSites < unit.sites.Count);

            // merge cells
            tableOne.Cell(1, 1).Merge(tableOne.Cell(2, 1));
            tableOne.Cell(1, 2).Merge(tableOne.Cell(2, 2));
            tableOne.Cell(1, 3).Merge(tableOne.Cell(2, 3));

            tableOne.Cell(1, 1).Range.Text = "序號";
            tableOne.Cell(1, 2).Range.Text = "URL/IP";
            tableOne.Cell(1, 3).Range.Text = "網站名稱";

            // delete columns according to verify level and merge
            if (caseinfo.level == "Critical" || caseinfo.level == "critical")
            {
                tableOne.Columns[7].Delete();
                tableOne.Columns[6].Delete();
                tableOne.Columns[5].Delete();
            }
            else if (caseinfo.level == "High" || caseinfo.level == "high")
            {
                tableOne.Columns[7].Delete();
                tableOne.Columns[6].Delete();
                tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
            }
            else if (caseinfo.level == "Medium" || caseinfo.level == "medium")
            {
                tableOne.Columns[7].Delete();
                tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
            }
            else if (caseinfo.level == "Low" || caseinfo.level == "low")
            {
                tableOne.Cell(1, 6).Merge(tableOne.Cell(1, 7));
                tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
            }
            tableOne.Cell(1, 4).Range.Text = "弱點數量";
        }

        static void writeCaseInfoToReport(ref Word.Document report, CaseInfo caseinfo, Unit unit)
        {
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
            report.Content.Find.Execute("p_numOfSites", false, false, false, false, false, true, 1, false, unit.sites.Count.ToString(), 2, false, false, false, false);
            if (unit.name != "000")
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "-" + unit.name, 2, false, false, false, false);
            else
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "", 2, false, false, false, false);


            string levelString = "";
            if (caseinfo.tool == "WebInspect" || caseinfo.tool == "webinspect" || caseinfo.tool == "Webinspect")
            {
                switch(caseinfo.level)
                {
                    case "Critical":
                    case "critical":
                        levelString = "嚴重風險（Critical）";
                        break;
                    case "High":
                    case "high":
                        levelString = "嚴重 / 高風險（Critical / High）";
                        break;
                    case "Medium":
                    case "medium":
                        levelString = "嚴重 / 高 / 中風險（Critical / High / Medium）";
                        break;
                    case "Low":
                    case "low":
                        levelString = "嚴重 / 高 / 中 / 低風險（Critical / High / Medium / Low）";
                        break;
                }
            }
            else
            {
                switch (caseinfo.level)
                {
                    case "Critical":
                    case "critical":
                        levelString = "嚴重風險（Critical）";
                        break;
                    case "High":
                    case "high":
                        levelString = "高風險（Critical）";
                        break;
                    case "Medium":
                    case "medium":
                        levelString = "高 / 中風險（High / Medium）";
                        break;
                    case "Low":
                    case "low":
                        levelString = "高 / 中 / 低風險（High / Medium / Low）";
                        break;
                }
            }
            report.Content.Find.Execute("p_level", false, false, false, false, false, true, 1, false, levelString, 2, false, false, false, false);
            
            foreach (Word.Section section in report.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Find.Execute("p_userName", false, false, false, false, false, true, 1, false, caseinfo.userName, 2, false, false, false, false);
                headerRange.Find.Execute("p_projectName", false, false, false, false, false, true, 1, false, caseinfo.projectName, 2, false, false, false, false);
                headerRange.Find.Execute("p_reportName", false, false, false, false, false, true, 1, false, caseinfo.reportName, 2, false, false, false, false);
                headerRange.Find.Execute("p_period", false, false, false, false, false, true, 1, false, caseinfo.period, 2, false, false, false, false);
                if (unit.name != "000")
                    headerRange.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "-" + unit.name, 2, false, false, false, false);
                else
                    headerRange.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "", 2, false, false, false, false);
            }
        }

        static void setCaseInfo(Excel.Worksheet sheet, ref CaseInfo caseinfo)
        {
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
        }

        static void printBanner()
        {
            // print banner
            Console.WriteLine(@" ___       __   ________  ________  ________       ");
            Console.WriteLine(@"|\  \     |\  \|\   ____\|\   __  \|\   ____\      ");
            Console.WriteLine(@"\ \  \    \ \  \ \  \___|\ \  \|\  \ \  \___|_     ");
            Console.WriteLine(@" \ \  \  __\ \  \ \_____  \ \   _  _\ \_____  \    ");
            Console.WriteLine(@"  \ \  \|\__\_\  \|____|\  \ \  \\  \\|____|\  \   ");
            Console.WriteLine(@"   \ \____________\____\_\  \ \__\\ _\ ____\_\  \  ");
            Console.WriteLine(@"    \|____________|\_________\|__|\|__|\_________\ ");
            Console.WriteLine(@"                  \|_________|        \|_________| ");
            Console.WriteLine(@"                                                   ");
            Console.WriteLine(@"                                                   ");
            Console.WriteLine(@"                   v0.1 by tenghaooo               ");
            Console.WriteLine(@"                                                   ");
            Console.WriteLine(@"===================================================");
        }

        static void setVulns(Excel.Worksheet resultSheet, ref List<Unit> Units)
        {
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
                            Units[y].sites[z].numOfLevelVulns[currentVulnLevel]++;
                            found = true;
                            break;
                        }
                    }
                    if (found)
                        break;
                }
                x++;
            }
        }

        static void setUnitsAndSites(Excel.Worksheet targetSheet, ref List<Unit> Units)
        {
            int i = 2;
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
            }
        }
    }
}
