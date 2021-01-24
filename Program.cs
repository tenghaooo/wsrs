using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

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
            Excel.Workbook excelBook = null;
            Word.Document vulnDes = null;
            Word.Document vulnSolu = null;
            Word.Document vulnCheck = null;

            try
            {
                
                printBanner();

                List<Unit> Units = new List<Unit>();
                CaseInfo caseinfo = new CaseInfo();

                string resultExcel = @"D:\TemplateFiles\sample.xlsx";

                Console.WriteLine("[L] Start running...");

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write(@"[I] Please input result excel full path (ex. D:\TemplateFiles\sample.xlsx): ");
                Console.ResetColor();

                resultExcel = Console.ReadLine();

                // open template docx
                vulnDes = wordApp.Documents.Open(@"D:\TemplateFiles\vulndes.docx", ReadOnly: true);
                vulnSolu = wordApp.Documents.Open(@"D:\TemplateFiles\vulnsolu.docx", ReadOnly: true);
                vulnCheck = wordApp.Documents.Open(@"D:\TemplateFiles\vulncheck.docx", ReadOnly: true);

                // open result excel table
                Console.WriteLine("[L] Openning result excel");
                excelBook = excelApp.Workbooks.Open(resultExcel, ReadOnly: true);

                // setup units and sites
                Console.WriteLine("[L] Loading units and sites");
                Excel.Worksheet targetSheet = excelBook.Worksheets["targets"];
                setUnitsAndSites(targetSheet, ref Units);

                int sum = 0;
                for (int i = 0; i < Units.Count; i++)
                {
                    sum += Units[i].sites.Count;
                }
                Console.WriteLine("[L] Total " + Units.Count + " units and " + sum + " sites");

                // setup caseinfo
                Console.WriteLine("[L] Loading case info");
                Excel.Worksheet caseInfoSheet = excelBook.Worksheets["caseinfo"];
                setCaseInfo(caseInfoSheet, ref caseinfo);

                // setup vulns
                Console.WriteLine("[L] Loading vulns");
                Excel.Worksheet resultSheet = excelBook.Worksheets["result"];
                if (caseinfo.secondScan == "初" || caseinfo.secondScan == "初掃")
                {
                    setVulns(resultSheet, ref Units);
                }
                else if (caseinfo.secondScan == "複" || caseinfo.secondScan == "複掃")
                {
                    setVulns2(resultSheet, ref Units);
                }
                

                if (caseinfo.secondScan == "初" || caseinfo.secondScan == "初掃")
                {
                    /*
                    * Big Loop For Units, Create Unit Report, First Scan
                    * */
                    Console.WriteLine("[L] Creating unit report, First Scan");
                    for (int U = 0; U < Units.Count; U++)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[L] Creating report " + (U + 1).ToString() + "/" + Units.Count.ToString() + ", Unit: " + Units[U].name);
                        Console.ResetColor();

                        int vulnNum = getVulnNumOfUnit(Units[U]);
                        List<string> vulnsName = getVulnsNameOfUnit(Units[U]);
                        bool haveVuln = (vulnNum == 0) ? false : true;
                        Word.Document report = null;

                        string reportPath = "D:\\Reports\\";
                        string reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "_" + caseinfo.period + ".docx"; ;
                        if (Units[U].name != "000")
                            reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "-" + Units[U].name + "_" + caseinfo.period + ".docx";

                        try
                        {
                            // open report template
                            if (!haveVuln)
                            {
                                report = wordApp.Documents.Open(@"D:\TemplateFiles\no_vuln_sample.docx");
                            }
                            else
                            {
                                report = wordApp.Documents.Open(@"D:\TemplateFiles\sample.docx");
                            }

                            if (haveVuln)
                            {
                                // write report create date
                                Console.WriteLine("    [L] Writting create date");
                                writeCreateDateToReport(ref report);

                                // write caseinfo to report and header
                                Console.WriteLine("    [L] Writting caseinfo to report & header");
                                writeCaseInfoToReport(ref report, caseinfo, Units[U]);

                                // write table one
                                Console.WriteLine("    [L] Writting table 1");
                                writeTableOneToReport(ref report, caseinfo, Units[U]);

                                // write table two
                                Console.WriteLine("    [L] Writting table 2");
                                writeTableTwoToReport(ref report, caseinfo, Units[U]);

                                // write vuln description
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns descriptions");
                                Console.ResetColor();
                                writeVulnDesToReport(ref report, ref vulnDes, caseinfo, Units[U], vulnsName);

                                // write vuln check
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns check");
                                Console.ResetColor();
                                writeVulnCheckToReport(ref report, ref vulnCheck, caseinfo, Units[U], vulnsName);

                                // write vuln solution
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns solutions");
                                Console.ResetColor();
                                writeVulnSoluToReport(ref report, ref vulnSolu, caseinfo, Units[U], vulnsName);

                            }
                            else if (!haveVuln)
                            {
                                // write report create date
                                Console.WriteLine("    [L] Writting create date");
                                writeCreateDateToReport(ref report);

                                // write caseinfo to report and header
                                Console.WriteLine("    [L] Writting caseinfo to report & header");
                                writeCaseInfoToReport(ref report, caseinfo, Units[U]);

                                // write table one
                                Console.WriteLine("    [L] Writting table 1");
                                writeTableOneToReport(ref report, caseinfo, Units[U]);

                                Console.WriteLine("    [L] This is a 0 vulns report");

                            }

                            // update content
                            Console.WriteLine("    [L] Updating content");
                            foreach (Word.TableOfContents tableOfContents in report.TablesOfContents)
                            {
                                tableOfContents.Update();
                            }
                            foreach (Word.TableOfFigures tableOfFigures in report.TablesOfFigures)
                            {
                                tableOfFigures.Update();
                            }
                            foreach (Word.Range storyRange in report.StoryRanges)
                            {
                                storyRange.Fields.Update();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[L] Failed!!!");
                            Console.WriteLine(ex.Message);
                            Console.ResetColor();
                        }
                        finally
                        {
                            // save report
                            Console.WriteLine("    [L] Saving report");
                            if (report != null)
                            {
                                report.SaveAs2(reportPath + reportName);
                                report.Close();
                            }

                        }
                        Console.WriteLine("    [L] Done.");
                    }
                }
                else if (caseinfo.secondScan == "複" || caseinfo.secondScan == "複掃")
                {
                    /*
                    * Big Loop For Units, Create Unit Report, Second Scan
                    * */
                    Console.WriteLine("[L] Creating unit report, Second Scan");
                    for (int U = 0; U < Units.Count; U++)
                    { 
                        if (!needToCreateReport(Units[U]))
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("[L] Skip this report because there are no vulns at first scan, Unit: "
                                + (U + 1).ToString() + "/" + Units.Count.ToString() + " " + Units[U].name);
                            Console.ResetColor();
                            continue;
                        }

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[L] Creating report " + (U + 1).ToString() + "/" + Units.Count.ToString() + ", Unit: " + Units[U].name);
                        Console.ResetColor();

                        int vulnNum = getVulnNumOfUnit2(Units[U]);
                        List<string> vulnsName = getVulnsNameOfUnit2(Units[U]);
                        bool haveVuln = (vulnNum == 0) ? false : true;
                        Word.Document report = null;

                        string reportPath = "D:\\Reports\\";
                        string reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "_" + caseinfo.period + ".docx"; ;
                        if (Units[U].name != "000")
                            reportName = "H07" + caseinfo.reportCode + "_" + caseinfo.period + "." + caseinfo.userName + caseinfo.reportName + "-" + Units[U].name + "_" + caseinfo.period + ".docx";

                        try
                        {
                            // open report template
                            if (!haveVuln)
                            {
                                report = wordApp.Documents.Open(@"D:\TemplateFiles\no_vuln_sample2.docx");
                            }
                            else
                            {
                                report = wordApp.Documents.Open(@"D:\TemplateFiles\sample2.docx");
                            }

                            if (haveVuln)
                            {
                                // write report create date
                                Console.WriteLine("    [L] Writting create date");
                                writeCreateDateToReport(ref report);

                                // write caseinfo to report and header
                                Console.WriteLine("    [L] Writting caseinfo to report & header");
                                writeCaseInfoToReport2(ref report, caseinfo, Units[U]);

                                // write table one
                                Console.WriteLine("    [L] Writting table 1");
                                writeTableOneToReport2(ref report, caseinfo, Units[U]);

                                // write table two
                                Console.WriteLine("    [L] Writting table 2");
                                writeTableTwoToReport2(ref report, caseinfo, Units[U]);

                                // write vuln description
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns descriptions");
                                Console.ResetColor();
                                writeVulnDesToReport(ref report, ref vulnDes, caseinfo, Units[U], vulnsName);

                                // write vuln check
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns check");
                                Console.ResetColor();
                                writeVulnCheckToReport(ref report, ref vulnCheck, caseinfo, Units[U], vulnsName);

                                // write vuln solution
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("    [L] Writting vulns solutions");
                                Console.ResetColor();
                                writeVulnSoluToReport(ref report, ref vulnSolu, caseinfo, Units[U], vulnsName);

                            }
                            else if (!haveVuln)
                            {
                                // write report create date
                                Console.WriteLine("    [L] Writting create date");
                                writeCreateDateToReport(ref report);

                                // write caseinfo to report and header
                                Console.WriteLine("    [L] Writting caseinfo to report & header");
                                writeCaseInfoToReport2(ref report, caseinfo, Units[U]);

                                // write table one
                                Console.WriteLine("    [L] Writting table 1");
                                writeTableOneToReport2(ref report, caseinfo, Units[U]);

                                // write table two
                                Console.WriteLine("    [L] Writting table 2");
                                writeTableTwoToReport2(ref report, caseinfo, Units[U]);

                                Console.WriteLine("    [L] This is a 0 vulns report");

                            }

                            // update content
                            Console.WriteLine("    [L] Updating content");
                            foreach (Word.TableOfContents tableOfContents in report.TablesOfContents)
                            {
                                tableOfContents.Update();
                            }
                            foreach (Word.TableOfFigures tableOfFigures in report.TablesOfFigures)
                            {
                                tableOfFigures.Update();
                            }
                            foreach (Word.Range storyRange in report.StoryRanges)
                            {
                                storyRange.Fields.Update();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[L] Failed!!!");
                            Console.WriteLine(ex.Message);
                            Console.ResetColor();
                        }
                        finally
                        {
                            // save report
                            Console.WriteLine("    [L] Saving report");
                            if (report != null)
                            {
                                report.SaveAs2(reportPath + reportName);
                                report.Close();
                            }

                        }
                        Console.WriteLine("    [L] Done.");
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("    [E] caseinfo.secondScan wrong value");
                    Console.ResetColor();
                }
                
                
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[L] Failed!!!");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
            finally
            {
                if (vulnDes != null)
                    vulnDes.Close();
                if (vulnSolu != null)
                    vulnSolu.Close();
                if (vulnCheck != null)
                    vulnCheck.Close();
                if (excelBook != null)
                    excelBook.Close();
                if (excelApp != null)
                    excelApp.Quit();
                if (wordApp != null)
                    wordApp.Quit();
            }
            
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("[L] Finish. Press Enter to quit!!!");
            Console.ResetColor();
            Console.ReadLine();
        }

        /*  IF needToCreateReport is True, it means there are vulns in this unit at first scan
         *  IF needToCreateReport is False, it means there are no vulns in this unit at first scan, so dont need to create report at second scan
         */
        static bool needToCreateReport(Unit unit)
        {
            bool result = true;
            int sum = 0;
            for (int i = 0; i < unit.sites.Count; i++)
            {
                sum += unit.sites[i].numOfLevelVulnsByLevel["Critical"];
                sum += unit.sites[i].numOfLevelVulnsByLevel["High"];
                sum += unit.sites[i].numOfLevelVulnsByLevel["Medium"];
                sum += unit.sites[i].numOfLevelVulnsByLevel["Low"];
            }

            if (sum == 0)
                result = false;

            return result;
        }
        static void writeVulnCheckToReport(ref Word.Document report, ref Word.Document vulnCheck, CaseInfo caseinfo, Unit unit, List<string> vulnsName)
        {
            Word.Range srcRange = vulnCheck.Content;
            Word.Range desRange = report.Content;

            // Dictionary<vulnName, Dictionary<siteName, vulnUrl>>
            Dictionary<string, Dictionary<string, string>> vulnSiteAndVulnUrl = getVulnSiteAndVulnUrl(unit);
            int vulnCount = 0;

            for (int i = 0; i < vulnsName.Count; i++, vulnCount++)
            {

                Console.WriteLine("        [L] " + (i + 1).ToString() + "/" + vulnsName.Count.ToString() + ": " + vulnsName[i]);
                
                // find desRange Start
                foreach (Word.Paragraph p in report.Paragraphs)
                {
                    if (p.Range.Text == "安全強化建議（修補方式）\r")
                    {
                        desRange.Start = p.Previous().Range.End;
                        desRange.End = p.Previous().Range.End;
                        break;
                    }
                }

                /*
                // if this is a new vuln then break a new page
                if (vulnCount != 0)
                {
                    desRange.InsertBreak(Word.WdBreakType.wdPageBreak);
                    // find desRange Start
                    foreach (Word.Paragraph p in report.Paragraphs)
                    {
                        if (p.Range.Text == "安全強化建議（修補方式）\r")
                        {
                            desRange.Start = p.Previous().Range.End;
                            desRange.End = p.Previous().Range.End;
                            break;
                        }
                    }
                }
                */
                Word.Paragraph temp;
                bool found = false;

                // set default srcRange
                foreach (Word.Paragraph p in vulnCheck.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == "p_noVuln弱點驗證不存在\r")
                    {
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;
                        break;
                    }
                }

                // find real srcRange and copy paste
                foreach (Word.Paragraph p in vulnCheck.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == vulnsName[i] + "\r")
                    {
                        // range of vuln title
                        srcRange.Start = temp.Range.Start;
                        srcRange.End = temp.Range.End;
                        found = true;

                        // paste vuln title
                        srcRange.Copy();
                        Thread.Sleep(100);
                        desRange.PasteSpecial(DataType: Word.WdPasteOptions.wdMatchDestinationFormatting);
                        Thread.Sleep(100);

                        // range of vuln check content
                        temp = temp.Next();
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;

                        // paste sites in this vuln
                        for (int j = 0; j < vulnSiteAndVulnUrl[vulnsName[i]].Count; j++)
                        {
                            Console.WriteLine("            [L] site " + (j + 1).ToString() + "/" + vulnSiteAndVulnUrl[vulnsName[i]].Count.ToString() + ": " + vulnSiteAndVulnUrl[vulnsName[i]].ElementAt(j).Key);
                            // set desRange
                            foreach (Word.Paragraph ptemp in report.Paragraphs)
                            {
                                if (ptemp.Range.Text == "安全強化建議（修補方式）\r")
                                {
                                    desRange.Start = ptemp.Previous().Range.End;
                                    desRange.End = ptemp.Previous().Range.End;
                                    break;
                                }
                            }
                            // paste vuln check content
                            srcRange.Copy();
                            Thread.Sleep(100);
                            desRange.PasteSpecial(DataType: Word.WdPasteOptions.wdMatchDestinationFormatting);
                            Thread.Sleep(100);

                            // replace p_vulnSiteName and p_vulnUrl
                            report.Content.Find.Execute("p_vulnSiteName", false, false, false, false, false, true, 1, false, vulnSiteAndVulnUrl[vulnsName[i]].ElementAt(j).Key, 2, false, false, false, false);
                            report.Content.Find.Execute("p_vulnUrl", false, false, false, false, false, true, 1, false, vulnSiteAndVulnUrl[vulnsName[i]].ElementAt(j).Value, 2, false, false, false, false);

                            // set desRange
                            foreach (Word.Paragraph ptemp in report.Paragraphs)
                            {
                                if (ptemp.Range.Text == "安全強化建議（修補方式）\r")
                                {
                                    desRange.Start = ptemp.Previous().Range.End;
                                    desRange.End = ptemp.Previous().Range.End;
                                    break;
                                }
                            }
                            desRange.InsertBreak(Word.WdBreakType.wdPageBreak);

                        }

                        break;
                    }
                }
                if (!found)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("        [E] This vuln isn't exist in vulnCheck doc: " + vulnsName[i]);
                    Console.ResetColor();
                    srcRange.Copy();
                    Thread.Sleep(100);
                    desRange.PasteSpecial(DataType: Word.WdPasteOptions.wdMatchDestinationFormatting);
                    Thread.Sleep(100);
                    report.Content.Find.Execute("p_noVuln", false, false, false, false, false, true, 1, false, vulnsName[i], 2, false, false, false, false);
                }
                
            }

            // insert break
            foreach (Word.Paragraph p in report.Paragraphs)
            {
                if (p.Range.Text == "弱點手動檢核\r")
                {
                    Word.Range temp = report.Content;
                    temp.Start = p.Previous().Range.End;
                    temp.End = p.Previous().Range.End;
                    temp.InsertBreak(Word.WdBreakType.wdPageBreak);
                    break;
                }
            }
        }

        static Dictionary<string, Dictionary<string, string>> getVulnSiteAndVulnUrl(Unit unit)
        {
            Dictionary<string, Dictionary<string, string>> result = new Dictionary<string, Dictionary<string, string>>();

            for (int i = 0; i < unit.sites.Count; i++)
            {
                for (int j = 0; j < unit.sites[i].vulns.Count; j++)
                {

                    // contain vuln name
                    if (result.ContainsKey(unit.sites[i].vulns[j].name))
                    {
                        // contain vuln name and contain site name
                        if (result[unit.sites[i].vulns[j].name].ContainsKey(unit.sites[i].name))
                        {
                            // nothing to do
                        }
                        // contain vuln name but not contain site name
                        else
                        {
                            result[unit.sites[i].vulns[j].name].Add(unit.sites[i].name, unit.sites[i].vulns[j].vulnUrl);
                        }
                    }
                    // not contain vuln name
                    else
                    {
                        Dictionary<string, string> newVulnDic = new Dictionary<string, string>();
                        newVulnDic.Add(unit.sites[i].name, unit.sites[i].vulns[j].vulnUrl);
                        result.Add(unit.sites[i].vulns[j].name, newVulnDic);
                    }
                }
            }

            return result;
        }

        static void writeVulnSoluToReport(ref Word.Document report, ref Word.Document vulnSolu, CaseInfo caseinfo, Unit unit, List<string> vulnsName)
        {
            Word.Range srcRange = vulnSolu.Content;
            Word.Range desRange = report.Content;

            for (int i = 0; i < vulnsName.Count; i++)
            {
                Console.WriteLine("        [L] " + (i + 1).ToString() + "/" + vulnsName.Count.ToString() + ": " + vulnsName[i]);
                // find desRange Start
                desRange.Start = report.Content.End;
                desRange.End = report.Content.End;
                
                Word.Paragraph temp;
                bool found = false;

                // set default srcRange
                foreach (Word.Paragraph p in vulnSolu.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == "p_noVuln弱點修補建議不存在\r")
                    {
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;
                        break;
                    }
                }

                // find real srcRange
                foreach (Word.Paragraph p in vulnSolu.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == vulnsName[i] + "\r")
                    {
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("        [E] This vuln isn't exist in vulnSolu doc: " + vulnsName[i]);
                    Console.ResetColor();
                }
                srcRange.Copy();
                Thread.Sleep(100);
                desRange.PasteSpecial(DataType: Word.WdPasteOptions.wdMatchDestinationFormatting);
                Thread.Sleep(100);
                report.Content.Find.Execute("p_noVuln", false, false, false, false, false, true, 1, false, vulnsName[i], 2, false, false, false, false);
            }
        }

        static List<string> getVulnsNameOfUnit(Unit unit)
        {
            List<string> vulnsName = new List<string>();

            for (int i = 0; i < unit.sites.Count; i++)
            {
                for (int j = 0; j < unit.sites[i].vulns.Count; j++)
                {
                    if (!vulnsName.Contains(unit.sites[i].vulns[j].name))
                    {
                        vulnsName.Add(unit.sites[i].vulns[j].name);
                    }
                }
            }
            return vulnsName;
        }

        static List<string> getVulnsNameOfUnit2(Unit unit)
        {
            List<string> vulnsName = new List<string>();

            for (int i = 0; i < unit.sites.Count; i++)
            {
                for (int j = 0; j < unit.sites[i].vulns.Count; j++)
                {
                    if (!vulnsName.Contains(unit.sites[i].vulns[j].name))
                    {
                        if (unit.sites[i].vulns[j].vulnUrl2 != "此弱點已不存在")
                            vulnsName.Add(unit.sites[i].vulns[j].name);
                    }
                }
            }
            return vulnsName;
        }

        static void writeVulnDesToReport(ref Word.Document report, ref Word.Document vulnDes, CaseInfo caseinfo, Unit unit, List<string> vulnsName)
        {

            Word.Range srcRange = vulnDes.Content;
            Word.Range desRange = report.Content;

            for (int i = 0; i < vulnsName.Count; i++)
            {
                Console.WriteLine("        [L] " + (i + 1).ToString() + "/" + vulnsName.Count.ToString() + ": " + vulnsName[i]);
                // find desRange Start
                foreach (Word.Paragraph p in report.Paragraphs)
                {
                    if (p.Range.Text == "弱點手動檢核\r")
                    {
                        desRange.Start = p.Previous().Range.End;
                        desRange.End = p.Previous().Range.End;
                        break;
                    }
                }

                Word.Paragraph temp;
                bool found = false;

                // set default srcRange
                foreach (Word.Paragraph p in vulnDes.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == "p_noVuln弱點說明不存在\r")
                    {
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;
                        break;
                    }
                }

                // find real srcRange
                foreach (Word.Paragraph p in vulnDes.Paragraphs)
                {
                    temp = p;
                    if (temp.Range.Text == vulnsName[i] + "\r")
                    {
                        srcRange.Start = temp.Range.Start;
                        while (temp.Next().Range.Text != "endofparagraph\r")
                        {
                            temp = temp.Next();
                        }
                        srcRange.End = temp.Range.End;
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("        [E] This vuln isn't exist in vulnDes doc: " + vulnsName[i]);
                    Console.ResetColor();
                }
                srcRange.Copy();
                Thread.Sleep(100);
                desRange.PasteSpecial(DataType: Word.WdPasteOptions.wdMatchDestinationFormatting);
                Thread.Sleep(100);
                report.Content.Find.Execute("p_noVuln", false, false, false, false, false, true, 1, false, vulnsName[i], 2, false, false, false, false);
            }
        }

        
        static int getVulnNumOfSiteFirstScan(Site site)
        {
            int vulnNum = 0;

            vulnNum += site.numOfLevelVulnsByLevel["Critical"];
            vulnNum += site.numOfLevelVulnsByLevel["High"];
            vulnNum += site.numOfLevelVulnsByLevel["Medium"];
            vulnNum += site.numOfLevelVulnsByLevel["Low"];

            return vulnNum;
        }

        static int getVulnNumOfUnit(Unit unit)
        {
            int vulnNum = 0;

            for (int i = 0; i < unit.sites.Count; i++)
            {
                vulnNum += unit.sites[i].vulns.Count;
            }

            return vulnNum;
        }

        static int getVulnNumOfUnit2(Unit unit)
        {
            int vulnNum = 0;

            for (int i = 0; i < unit.sites.Count; i++)
            {
                vulnNum += unit.sites[i].numOfLevelVulnsByLevel2["Critical"];
                vulnNum += unit.sites[i].numOfLevelVulnsByLevel2["High"];
                vulnNum += unit.sites[i].numOfLevelVulnsByLevel2["Medium"];
                vulnNum += unit.sites[i].numOfLevelVulnsByLevel2["Low"];
            }

            return vulnNum;
        }

        static void writeCreateDateToReport(ref Word.Document report)
        {
            String sDate = DateTime.Now.ToString();
            DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));
            String yy = (datevalue.Year - 1911).ToString();
            String mn = datevalue.Month.ToString();
            String dy = datevalue.Day.ToString();
            report.Content.Find.Execute("p_rYear", false, false, false, false, false, true, 1, false, yy, 2, false, false, false, false);
            report.Content.Find.Execute("p_rMonth", false, false, false, false, false, true, 1, false, mn, 2, false, false, false, false);
            report.Content.Find.Execute("p_rDay", false, false, false, false, false, true, 1, false, dy, 2, false, false, false, false);
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
            
            for (int i = 2; i <= tableTwo.Rows.Count; i++)
            {
                tableTwo.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        static void writeTableTwoToReport2(ref Word.Document report, CaseInfo caseinfo, Unit unit)
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
                    tableTwo.Cell(row, 6).Range.Text = unit.sites[i].vulns[j].vulnUrl2;
                    tableTwo.Rows.Add();
                    row++;
                }
            }
            tableTwo.Rows.Last.Delete();

            for (int i = 2; i <= tableTwo.Rows.Count; i++)
            {
                tableTwo.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableTwo.Cell(i, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        static void writeTableOneToReport2(ref Word.Document report, CaseInfo caseinfo, Unit unit)
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
                tableOne.Cell(row, 4).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Critical"].ToString();
                tableOne.Cell(row, 5).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["High"].ToString();
                tableOne.Cell(row, 6).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Medium"].ToString();
                tableOne.Cell(row, 7).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Low"].ToString();
                if (getVulnNumOfSiteFirstScan(unit.sites[indexOfSites]) != 0)
                {
                    tableOne.Cell(row, 8).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel2["Critical"].ToString();
                    tableOne.Cell(row, 9).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel2["High"].ToString();
                    tableOne.Cell(row, 10).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel2["Medium"].ToString();
                    tableOne.Cell(row, 11).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel2["Low"].ToString();
                }
                else
                {
                    tableOne.Cell(row, 8).Range.Text = "N/A";
                    tableOne.Cell(row, 9).Range.Text = "N/A";
                    tableOne.Cell(row, 10).Range.Text = "N/A";
                    tableOne.Cell(row, 11).Range.Text = "N/A";
                }
                

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
                tableOne.Columns[11].Delete();
                tableOne.Columns[10].Delete();
                tableOne.Columns[9].Delete();
                tableOne.Columns[5].Delete();
                tableOne.Columns[5].Delete();
                tableOne.Columns[5].Delete();
                tableOne.Columns[4].Width = 90;
                tableOne.Columns[5].Width = 90;
            }
            else if (caseinfo.level == "High" || caseinfo.level == "high")
            {
                tableOne.Columns[11].Delete();
                tableOne.Columns[10].Delete();
                tableOne.Columns[6].Delete();
                tableOne.Columns[6].Delete();
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                {
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }
                else
                {
                    tableOne.Columns[6].Delete();
                    tableOne.Columns[4].Delete();
                    tableOne.Columns[4].Width = 90;
                    tableOne.Columns[5].Width = 90;
                }
            }
            else if (caseinfo.level == "Medium" || caseinfo.level == "medium")
            {
                tableOne.Columns[11].Delete();
                tableOne.Columns[7].Delete();
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                {
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }
                else
                {
                    tableOne.Columns[7].Delete();
                    tableOne.Columns[4].Delete();
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }

            }
            else if (caseinfo.level == "Low" || caseinfo.level == "low")
            {
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                {
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }
                else
                {
                    tableOne.Columns[8].Delete();
                    tableOne.Columns[4].Delete();
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }

            }
            tableOne.Cell(1, 4).Range.Text = "初掃弱點數量";
            tableOne.Cell(1, 5).Range.Text = "複掃弱點數量";

            for (int i = 3; i <= tableOne.Rows.Count; i++)
            {
                tableOne.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableOne.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }

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
                tableOne.Cell(row, 4).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Critical"].ToString();
                tableOne.Cell(row, 5).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["High"].ToString();
                tableOne.Cell(row, 6).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Medium"].ToString();
                tableOne.Cell(row, 7).Range.Text = unit.sites[indexOfSites].numOfLevelVulnsByLevel["Low"].ToString();

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
                tableOne.Columns[4].Width = 60;
            }
            else if (caseinfo.level == "High" || caseinfo.level == "high")
            {
                tableOne.Columns[7].Delete();
                tableOne.Columns[6].Delete();
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                else
                {
                    tableOne.Columns[4].Delete();
                    tableOne.Columns[4].Width = 60;
                }
                    
                
            }
            else if (caseinfo.level == "Medium" || caseinfo.level == "medium")
            {
                tableOne.Columns[7].Delete();
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                {
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                }
                else
                {
                    tableOne.Columns[4].Delete();
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                }
                
            }
            else if (caseinfo.level == "Low" || caseinfo.level == "low")
            {
                if (caseinfo.tool == "WebInspect" || caseinfo.tool == "Webinspect" || caseinfo.tool == "webinspect")
                {
                    tableOne.Cell(1, 6).Merge(tableOne.Cell(1, 7));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                    tableOne.Cell(1, 4).Merge(tableOne.Cell(1, 5));
                }
                else
                {
                    tableOne.Columns[4].Delete();
                    tableOne.Cell(1, 6).Merge(tableOne.Cell(1, 7));
                    tableOne.Cell(1, 5).Merge(tableOne.Cell(1, 6));
                }
                
            }
            tableOne.Cell(1, 4).Range.Text = "弱點數量";

            for (int i = 3; i <= tableOne.Rows.Count; i++)
            {
                tableOne.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableOne.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }

        }

        static int getNumOfSitesOfSecondScan(Unit unit)
        {
            int numOfSites = 0;
            int numOfZeroVulnSiteInFirstScan = 0;

            for (int i = 0; i < unit.sites.Count; i++)
            {
                if (getVulnNumOfSiteFirstScan(unit.sites[i]) == 0)
                {
                    numOfZeroVulnSiteInFirstScan++;
                }
            }

            numOfSites = unit.sites.Count - numOfZeroVulnSiteInFirstScan;

            return numOfSites;
        }

        static void writeCaseInfoToReport2(ref Word.Document report, CaseInfo caseinfo, Unit unit)
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
            report.Content.Find.Execute("p_numOfSites", false, false, false, false, false, true, 1, false, getNumOfSitesOfSecondScan(unit).ToString(), 2, false, false, false, false);

            // write unit name
            if (unit.name != "000")
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "-" + unit.name, 2, false, false, false, false);
            else
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "", 2, false, false, false, false);

            // write level string
            string levelString = "";
            if (caseinfo.tool == "WebInspect" || caseinfo.tool == "webinspect" || caseinfo.tool == "Webinspect")
            {
                switch (caseinfo.level)
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
                        levelString = "高風險（High）";
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

            // write info in header
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

            // write no vuln string
            string noVuln = "";
            if (caseinfo.level == "Critical" || caseinfo.level == "critical")
            {
                noVuln = "嚴重";
            }
            else if (caseinfo.level == "High" || caseinfo.level == "high")
            {
                noVuln = "高";
            }
            else if (caseinfo.level == "Medium" || caseinfo.level == "medium")
            {
                noVuln = "中";
            }
            else if (caseinfo.level == "Low" || caseinfo.level == "low")
            {
                noVuln = "低";
            }
            report.Content.Find.Execute("p_noVuln", false, false, false, false, false, true, 1, false, noVuln, 2, false, false, false, false);
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

            // write unit name
            if (unit.name != "000")
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "-" + unit.name, 2, false, false, false, false);
            else
                report.Content.Find.Execute("p_unitName", false, false, false, false, false, true, 1, false, "", 2, false, false, false, false);

            // write level string
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
                        levelString = "高風險（High）";
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

            // write info in header
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

            // write no vuln string
            string noVuln = "";
            if (caseinfo.level == "Critical" || caseinfo.level == "critical")
            {
                noVuln = "嚴重";
            }
            else if (caseinfo.level == "High" || caseinfo.level == "high")
            {
                noVuln = "高";
            }
            else if (caseinfo.level == "Medium" || caseinfo.level == "medium")
            {
                noVuln = "中";
            }
            else if (caseinfo.level == "Low" || caseinfo.level == "low")
            {
                noVuln = "低";
            }
            report.Content.Find.Execute("p_noVuln", false, false, false, false, false, true, 1, false, noVuln, 2, false, false, false, false);
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
            caseinfo.secondScan = sheet.Cells[2, "L"].Value.ToString();
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
            Console.WriteLine(@"                   v0.8 by tenghaooo               ");
            Console.WriteLine(@"                                                   ");
            Console.WriteLine(@"===================================================");
        }

        static void setVulns2(Excel.Worksheet resultSheet, ref List<Unit> Units)
        {
            int x = 2;
            while (resultSheet.Cells[x, "A"].Value != null)
            {
                string currentUrl = resultSheet.Cells[x, "A"].Value.ToString();
                string currentSiteName = resultSheet.Cells[x, "B"].Value.ToString();
                string currentVulnName = resultSheet.Cells[x, "C"].Value.ToString();
                string currentVulnLevel = resultSheet.Cells[x, "D"].Value.ToString();
                string currentVulnUrl = resultSheet.Cells[x, "E"].Value.ToString();
                string currentVulnUrl2 = resultSheet.Cells[x, "F"].Value.ToString();

                // add vuln to units sites
                Vuln newVuln = new Vuln();
                newVuln.name = currentVulnName;
                newVuln.level = currentVulnLevel;
                newVuln.vulnUrl = currentVulnUrl;
                newVuln.vulnUrl2 = currentVulnUrl2;

                // find current vuln in which unit and site
                for (int y = 0; y < Units.Count; y++)
                {
                    bool found = false;
                    for (int z = 0; z < Units[y].sites.Count; z++)
                    {
                        if (Units[y].sites[z].name == currentSiteName)
                        {
                            Units[y].sites[z].vulns.Add(newVuln);
                            Units[y].sites[z].numOfLevelVulnsByLevel[currentVulnLevel]++;
                            if (currentVulnUrl2 != "此弱點已不存在")
                            {
                                Units[y].sites[z].numOfLevelVulnsByLevel2[currentVulnLevel]++;
                            }
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
                                Units[y].sites[z].numOfLevelVulnsByLevel[currentVulnLevel]++;
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
