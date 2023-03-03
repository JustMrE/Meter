using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meter
{
    public class Formula
    {
        public Dictionary<string, List<ForTags>> formulas { get; set; }

        public Formula()
        {
            formulas = new Dictionary<string, List<ForTags>>();
        }
    }
}
