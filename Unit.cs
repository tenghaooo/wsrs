using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wsrs
{
    class Unit
    {
        public string name { get; set; }
        public List<Site> sites { get; set; }

        public Unit()
        {
            sites = new List<Site>();
        }

        public Unit(string name, List<Site> sites)
        {
            this.name = name;
            this.sites = sites;
        }
    }
}
