using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer
{
    public class SwfConverter : Converter
    {
        public override void Convert(string src, string dest, out string msg)
        {
            msg = null;
            if (string.IsNullOrEmpty(src))
            {
                msg = "源文件不能为空";
                return;
            }
            var contentType = MimeTypes.GetContentType(src);
            if (contentType != "application/pdf")
            {
                msg = "源文件应为pdf文件";
                return;
            }
            if (string.IsNullOrEmpty(dest))
            {
                msg = "目标路径不能为空";
                return;
            }
            if (!File.Exists(src))
            {
                msg = "源文件不存在";
                return;
            }

            if (File.Exists(dest))
            {
                msg = "目标文件已存在";
                return;
            }

            try
            {
                //将pdf文档转成temp.swf文件
                string cmd = String.Format("\"{0}\" -o \"{1}\" -t -s flashversion=9"
                    //,Config.PDF2SWF_PATH
                     , src.ToString()
                     , dest.ToString());
                string tmsg = null;
                RunShell(cmd, out tmsg);
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
            }
        }
        /// <summary>
        /// 运行命令
        /// </summary>
        /// <param name="strShellCommand">命令字符串</param>
        private static void RunShell(string ShellCommand, out string msg)
        {
            msg = null;
            Config con = new Config();

            using (Process process = new Process())
            {
                process.StartInfo.FileName = con.PDF2SWF_PATH;
                process.StartInfo.Arguments = ShellCommand;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;

                StringBuilder output = new StringBuilder();
                StringBuilder error = new StringBuilder();

                using (System.Threading.AutoResetEvent outputWaitHandle = new System.Threading.AutoResetEvent(false))
                using (System.Threading.AutoResetEvent errorWaitHandle = new System.Threading.AutoResetEvent(false))
                {
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (e.Data == null)
                        {
                            outputWaitHandle.Set();
                        }
                        else
                        {
                            output.AppendLine(e.Data);
                        }
                    };
                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (e.Data == null)
                        {
                            errorWaitHandle.Set();
                        }
                        else
                        {
                            error.AppendLine(e.Data);
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    if (process.WaitForExit(con.PDF2SWF_TimeOut) &&
                        outputWaitHandle.WaitOne(con.PDF2SWF_TimeOut) &&
                        errorWaitHandle.WaitOne(con.PDF2SWF_TimeOut))
                    {
                        msg = "pdf转换swf成功";
                    }
                    else
                    {
                        msg = "pdf转换swf出现延时";
                    }
                }
            }
        }
    }
}
