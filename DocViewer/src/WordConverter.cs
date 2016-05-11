using System;
using System.IO;

namespace DocViewer
{
    public class WordConverter : Converter
    {
        private Microsoft.Office.Interop.Word.Application app;
        private Microsoft.Office.Interop.Word.Documents docs;
        private Microsoft.Office.Interop.Word.Document doc;
        object paraMissing = Type.Missing;
        public override void Convert(string src, string dest, out string msg)
        {
            msg = null;
            Object nothing = System.Reflection.Missing.Value;
            try
            {
                if (string.IsNullOrEmpty(src))
                {
                    msg = "源文件不能为空";
                    return;
                }
                var contentType = MimeTypes.GetContentType(src);
                if (contentType != "application/msword")
                {
                    msg = "源文件应为word文件";
                    return;
                }
                if (string.IsNullOrEmpty(dest))
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

                app = new Microsoft.Office.Interop.Word.Application();
                docs = app.Documents;
                doc = docs.Open(src, false, true, false, nothing, nothing, true, nothing, nothing, nothing, nothing, false, false, nothing, true, nothing);
                doc.ExportAsFixedFormat(dest, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 1, 1, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, false, false, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, false, false, false, nothing);

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
            if (doc != null)
            {
                try
                {
                    ((Microsoft.Office.Interop.Word._Document)doc).Close(false);
                    releaseCOMObject(doc);
                }
                catch (Exception e)
                {

                }
            }

            if (docs != null)
            {
                try
                {
                    docs.Close(false);
                    releaseCOMObject(docs);
                }
                catch (Exception e)
                {

                }
            }

            if (app != null)
            {
                try
                {

                    ((Microsoft.Office.Interop.Word._Application)app).Quit();
                    releaseCOMObject(app);
                }
                catch (Exception e)
                {

                }
            }

        }

    }
}
