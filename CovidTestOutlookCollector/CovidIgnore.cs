using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CovidTestOutlookCollector
{
    class CovidIgnore
    {
        public List<string> files { get; set; }
        public List<string> ids { get; set; }

        public CovidIgnore()
        {
            files = new List<string>();
            ids = new List<string>();
        }
    }
}
