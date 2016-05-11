using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer
{
    public class ConverterProcessor : IConverterProcessor
    {
        private Converter converter;
        //委托方式处理文档转换
        public Dictionary<short, Action<string, string>> _dic
        {
            get
            {
                return new Dictionary<short, Action<string, string>>() 
                {
                    {0, (string src, string dest)=>WordConvert(src,dest)},
                    {1, (string src, string dest)=>ExcelConvert(src,dest)},
                    {2, (string src, string dest)=>PowerConvert(src,dest)},
                    {3, (string src, string dest)=>PdfConvert(src,dest)}
                };
            }
        }

        public void WordConvert(string src, string dest)
        {
            string msg;
            converter = new WordConverter();
            converter.Convert(src, dest, out msg);
        }

        public void ExcelConvert(string src, string dest)
        {
            string msg;
            converter = new ExcelConverter();
            converter.Convert(src, dest, out msg);
        }
        public void PowerConvert(string src, string dest)
        {
            string msg;
            converter = new PowerPointConverter();
            converter.Convert(src, dest, out msg);
        }

        public void PdfConvert(string src, string dest)
        {
            string msg;
            converter = new PowerPointConverter();
            converter.Convert(src, dest, out msg);
        }
        public void Converter(short policy, string src, string dest, out string msg)
        {
            msg = null;
            if (_dic.ContainsKey(policy))
            {
                _dic[policy].Invoke(src, dest);
            }
        }
    }
}
