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

        public Dictionary<string, int> numOfLevelVulns { get; set; }

        public Site() {
            vulns = new List<Vuln>();
            numOfLevelVulns = new Dictionary<string, int>();
            numOfLevelVulns.Add("Critical", 0);
            numOfLevelVulns.Add("High", 0);
            numOfLevelVulns.Add("Medium", 0);
            numOfLevelVulns.Add("Low", 0);
        }

        public Site(string url, string name, List<Vuln> vulns, Dictionary<string, int> numOfVulns)
        {
            this.url = url;
            this.name = name;
            this.vulns = vulns;
            this.numOfLevelVulns = numOfVulns;
        }
        
    }
}
