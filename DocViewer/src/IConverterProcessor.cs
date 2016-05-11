using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer
{
    public interface IConverterProcessor
    {
        System.Collections.Generic.Dictionary<short, Action<string, string>> _dic { get; }
        void Converter(short policy, string src, string dest, out string msg);
        void ExcelConvert(string src, string dest);
        void PowerConvert(string src, string dest);
        void WordConvert(string src, string dest);
    }
}
