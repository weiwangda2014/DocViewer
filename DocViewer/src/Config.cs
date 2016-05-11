
namespace DocViewer
{
    /// <summary>
    /// 配置表
    /// </summary>
    public class Config
    {
        /// <summary>
        /// pdf2swf.exe路径
        /// </summary>
        public  string PDF2SWF_PATH = null;
        /// <summary>
        /// pdf2swf.exe转换超时限制
        /// </summary>
        public  int PDF2SWF_TimeOut = 0;
        public Config()
        {
            PDF2SWF_PATH = System.Configuration.ConfigurationManager.AppSettings["PDF2SWFSrc"];
            string time_out = System.Configuration.ConfigurationManager.AppSettings["PDF2SWFTimeOut"];
            PDF2SWF_TimeOut = int.Parse(time_out);

        }

    }
}
