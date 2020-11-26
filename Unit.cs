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
        public ArrayList sites { get; set; }

        public Unit()
        {
            sites = new ArrayList();
        }

        public Unit(string name, ArrayList sites)
        {
            this.name = name;
            this.sites = sites;
        }
        
    }
}
