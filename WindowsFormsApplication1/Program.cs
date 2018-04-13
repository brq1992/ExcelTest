using NPOI.XWPF.UserModel;
using System;
using System.IO;
using System.Reflection;

namespace WindowsFormsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());

            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();
            using (FileStream sw = File.Create("blank.docx"))
            {
                doc.Write(sw);
            }
        }
    }
}
