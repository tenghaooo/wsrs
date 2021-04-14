using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wsrs
{
    class CaseInfo
    {
        public string userName { get; set; }
        public string projectName { get; set; }
        public string reportName { get; set; }
        public string period { get; set; }
        public string reportCode { get; set; }
        public string author { get; set; }
        public string tool { get; set; }
        public string year { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string level { get; set; }
        public string secondScan { get; set; }

        public string testType { get; set; }
        public CaseInfo() { }


        public CaseInfo(string userName, string projectName, string reportName, string period, string reportCode, string author, string tool, string year, string startDate, string endDate, string level, string secondScan, string testType)
        {
            this.userName = userName;
            this.projectName = projectName;
            this.reportName = reportName;
            this.period = period;
            this.reportCode = reportCode;
            this.author = author;
            this.tool = tool;
            this.year = year;
            this.startDate = startDate;
            this.endDate = endDate;
            this.level = level;
            this.secondScan = secondScan;
            this.testType = testType;
        }
    }
}
