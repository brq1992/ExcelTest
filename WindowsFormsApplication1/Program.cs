
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

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
            //System.Windows.Forms.Application.EnableVisualStyles();
            //System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            //System.Windows.Forms.Application.Run(new Form1());

            GenerateMainWorkBook();


            WorkBookManager manager = new WorkBookManager();
        }

        private static void GenerateMainWorkBook()
        {
            IWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet("MainSheet");

            FileStream sw = File.Create(@"e:/MainWorkBook.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
