using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wsrs
{
    class Vuln
    {
        public string name { set; get; }
        public string level { set; get; }
        public string vulnUrl { set; get; }
        public string vulnUrl2 { set; get; }

        public Vuln() { }
        public Vuln(string name, string level, string vulnUrl, string vulnUrl2)
        {
            this.name = name;
            this.level = level;
            this.vulnUrl = vulnUrl;
            this.vulnUrl2 = vulnUrl2;
        }
    }
}
