using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wsrs
{
    class Site
    {
        public string url { get; set; } 
        public string name { get; set; }
        public List<Vuln> vulns { get; set; }

        public Dictionary<string, int> numOfLevelVulnsByLevel { get; set; }
        public Dictionary<string, int> numOfLevelVulnsByLevel2 { get; set; }

        public Site() {
            vulns = new List<Vuln>();
            numOfLevelVulnsByLevel = new Dictionary<string, int>();
            numOfLevelVulnsByLevel.Add("Critical", 0);
            numOfLevelVulnsByLevel.Add("High", 0);
            numOfLevelVulnsByLevel.Add("Medium", 0);
            numOfLevelVulnsByLevel.Add("Low", 0);

            numOfLevelVulnsByLevel2 = new Dictionary<string, int>();
            numOfLevelVulnsByLevel2.Add("Critical", 0);
            numOfLevelVulnsByLevel2.Add("High", 0);
            numOfLevelVulnsByLevel2.Add("Medium", 0);
            numOfLevelVulnsByLevel2.Add("Low", 0);
        }

        public Site(string url, string name, List<Vuln> vulns, Dictionary<string, int> numOfVulns)
        {
            this.url = url;
            this.name = name;
            this.vulns = vulns;
            this.numOfLevelVulnsByLevel = numOfVulns;
        }
        
    }
}
