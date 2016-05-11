using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer.src
{
    public abstract class ConverterState
    {
        public string src { get; set; }
        public string dest { get; set; }
    }
}
