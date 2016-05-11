using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocViewer
{
    public class ExcelConverter : Converter
    {
        private Microsoft.Office.Interop.Excel.Application app;
        private Microsoft.Office.Interop.Excel.Workbooks books;
        private Microsoft.Office.Interop.Excel.Workbook book;

        public override void Convert(string src, string dest, out string msg)
        {
            msg = null;
            Object nothing = Type.Missing;
            try
            {
                if (string.IsNullOrEmpty(src))
                {
                    msg = "源文件不能为空";
                    return;
                }
                var contentType = MimeTypes.GetContentType(src);
                if (contentType != "application/excel")
                {
                    msg = "源文件应为excel文件";
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

                app = new Microsoft.Office.Interop.Excel.Application();
                books = app.Workbooks;
                book = books.Open(src, false, true, nothing, nothing, nothing, true, nothing, nothing, false, false, nothing, false, nothing, false);

                bool hasContent = false;
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                {
                    Microsoft.Office.Interop.Excel.Range range = sheet.UsedRange;
                    if (range != null)
                    {
                        Microsoft.Office.Interop.Excel.Range found = range.Cells.Find("*", nothing, nothing, nothing, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, nothing, nothing, nothing);
                        if (found != null)
                        {
                            hasContent = true;
                        }
                        releaseCOMObject(found);
                        releaseCOMObject(range);
                    }
                }

                if (!hasContent)
                {
                    msg = "此文件无内容";
                }
                book.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, dest, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityMinimum, false, false, nothing, nothing, false, nothing);
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
            if (book != null)
            {
                try
                {
                    book.Close(false);
                    releaseCOMObject(book);
                }
                catch (Exception e)
                {

                }
            }

            if (books != null)
            {
                try
                {
                    books.Close();
                    releaseCOMObject(books);
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
