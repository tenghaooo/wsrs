﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wsrs
{
    class Site
    {
        public string url;
        public string name;
        public ArrayList vulns;

        public Site() {
            vulns = new ArrayList();
        }

        public Site(string url, string name, ArrayList vulns)
        {
            this.url = url;
            this.name = name;
            this.vulns = vulns;
        }
        
    }
}
