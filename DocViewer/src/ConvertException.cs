using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocViewer
{
   public class ConvertException : Exception
    {
        public ConvertException(String message)
            : base(message)
        {
        }
    }
}
