using DocViewer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormsApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tryConvert(Converter converter, String inputFile, String outputFile)
        {
            try
            {
                string msg = null;
                converter.Convert(inputFile, outputFile, out msg);
                MessageBox.Show(msg);
            }
            catch (ConvertException err)
            {
                MessageBox.Show(err.Message + "\n\n" + err.StackTrace);
            }
        }

        private void btnWordClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String src = op.FileName;
            var contentType = MimeTypes.GetContentType(src);
            if (contentType != "application/msword")
            {
                string msg = "源文件应为word文件";
                MessageBox.Show(msg);
                return;
            }
            String outputFile = String.Concat(src, ".pdf");
            Converter converter = new WordConverter();
            tryConvert(converter, src, outputFile);
        }

        private void btnExcelClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String src = op.FileName;
            var contentType = MimeTypes.GetContentType(src);
            if (contentType != "application/excel")
            {
                string msg = "源文件应为excel文件";
                MessageBox.Show(msg);
                return;
            }
            String outputFile = String.Concat(src, ".pdf");
            Converter converter = new ExcelConverter();
            tryConvert(converter, src, outputFile);
        }

        private void btnPptClick(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.ShowDialog();
            String src = op.FileName;
            var contentType = MimeTypes.GetContentType(src);
            if (contentType != "application/mspowerpoint")
            {
                string msg = "源文件应为powerpoint文件";
                MessageBox.Show(msg);
                return;
            }
            String outputFile = String.Concat(src, ".pdf");
            Converter converter = new PowerPointConverter();
            tryConvert(converter, src, outputFile);
        }
    }
}
