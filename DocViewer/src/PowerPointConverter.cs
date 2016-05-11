using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer
{
    public class PowerPointConverter : Converter
    {

        private Microsoft.Office.Interop.PowerPoint.Application app;
        private Microsoft.Office.Interop.PowerPoint.Presentations presentations;
        private Microsoft.Office.Interop.PowerPoint.Presentation presentation;

        public override void Convert(string src, string dest, out string msg)
        {
            msg = null;
            try
            {
                if (string.IsNullOrEmpty(src))
                {
                    msg = "源文件不能为空";
                    return;
                }
                var contentType = MimeTypes.GetContentType(src);
                if (contentType != "application/mspowerpoint")
                {
                    msg = "源文件应为powerpoint文件";
                    return;
                }
                if ( string.IsNullOrEmpty(dest))
                {
                    msg = "目标路径不能为空";
                    return;
                }
                if (!File.Exists(src))
                {
                    msg = "文件不存在";
                    return;

                }

                if (IsPasswordProtected(src))
                {
                    msg = "存在密码";
                    return;
                }

                app = new Microsoft.Office.Interop.PowerPoint.Application();
                presentations = app.Presentations;
                presentation = presentations.Open(src.ToString(), Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                presentation.ExportAsFixedFormat(dest.ToString(), Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                // presentation.ExportAsFixedFormat(outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState.msoFalse, PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PowerPoint.PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null, PowerPoint.PpPrintRangeType.ppPrintAll, "", false, false, false, false, false, null);
            }
            catch (Exception e)
            {
                release();
                throw new ConvertException(e.Message);
            }
            release();
        }

        private void release()
        {
            if (presentation != null)
            {
                try
                {
                    presentation.Close();
                    releaseCOMObject(presentation);
                }
                catch (Exception e)
                {

                }
            }

            if (app != null)
            {
                try
                {
                    app.Quit();
                    releaseCOMObject(app);
                }
                catch (Exception e)
                {

                }
            }
        }
    }
}
