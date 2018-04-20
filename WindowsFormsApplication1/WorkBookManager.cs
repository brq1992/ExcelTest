using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using IFont = NPOI.SS.UserModel.IFont;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;

namespace WindowsFormsApplication1
{

//255,255,204   10
//0,255,0  52
//255,192,0  50
//247,150,70  57
//28,250,227 49
//255,0,0 48
//153,153,255 20
//128,0,128 55
//150,150,150 14
//146,208,80 51
//128,100,162 13
//255,204,153 11
//155,187,89  15
//255,255,255 40
//204,255,204 61
//255,128,128 22
//255,153,204 45
//255,255,0 47
//79,129,189 43
//75,172,198 42
//204,255,255 41 
//153,204,255 44
    class WorkBookManager
    {
        /// <summary>
        /// 主表
        /// </summary>
        private XSSFWorkbook mainWorkbook;

        private Dictionary<string, short> dictionary; 
        void Onece()
        {
            

            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //fill background
            ICellStyle style1 = workbook.CreateCellStyle();
            style1.FillForegroundColor = IndexedColors.Blue.Index;
            style1.FillPattern = FillPattern.SolidForeground;
            style1.FillBackgroundColor = IndexedColors.Pink.Index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            FileStream sw = File.Create(@"e:/test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
        public WorkBookManager()
        {

            dictionary = new Dictionary<string, short>
            {
                {"255,255,204", 10},
                {"0,255,0", 52},
                {"255,192,0", 50},
                {"247,150,70", 57},
                {"28,250,227", 49},
                {"255,0,0", 48},
                {"153,153,255", 20},
                {"128,0,128", 55},
                {"150,150,150", 14},
                {"146,208,80", 51},
                {"128,100,162", 13},
                {"255,204,153", 11},
                {"155,187,89", 15},
                {"255,255,255", 40},
                {"204,255,204", 61},
                {"255,128,128", 22},
                {"255,153,204", 45},
                {"255,255,0", 47},
                {"79,129,189", 43},
                {"75,172,198", 42},
                {"204,255,255", 41},
                {"153,204,255", 44},
                {"204,153,255",43}
            };

            GenerateMainWorkBook();
            XSSFSheet mainSheet = (XSSFSheet) mainWorkbook.GetSheet("MainSheet");

            GetSalaryList(mainSheet);

            //FileStream atdFileStream = new FileStream(@"e:/上海嵩恒网络科技股份有限公司_考勤报表_20180401-20180412.xlsx", FileMode.Open);
            //IWorkbook attendanceWorkbook = new XSSFWorkbook(atdFileStream);
            //atdFileStream.Close();
            //ISheet atdSheet = attendanceWorkbook.GetSheet("打卡时间");
            ////for (int i = 2; i < atdSheet.LastRowNum; i++)
            ////{
            ////    IRow sRow = atdSheet.GetRow(i);

            ////    IRow tRow = mainSheet.GetRow(i-2);
            ////    if (sRow != null && tRow != null)
            ////    {
            ////        for (int j = 0; j < sRow.LastCellNum; j++)
            ////        {
            ////            ICell tCell = tRow.CreateCell(j+2);
            ////            ICell sCell = sRow.GetCell(j);
            ////            if (tCell == null || sCell == null)
            ////            {
            ////                break;
            ////            }
            ////            tCell.SetCellValue(sCell.ToString());
            ////            tCell.SetCellFormula("VLOOKUP()");
            ////        }
            ////    }
            ////}

            //for (int i = 0; i < mainSheet.LastRowNum; i++)
            //{
            //    IRow tRow = mainSheet.GetRow(i);
            //    if (tRow == null)
            //    {
            //        continue;
            //    }
            //    ICell tCell = tRow.GetCell(0);
            //    if (tCell == null)
            //    {
            //        continue;
            //    }
            //    for (int j = 0; j < atdSheet.LastRowNum; j++)
            //    {
            //        IRow sRow = atdSheet.GetRow(j);
            //        if (sRow == null)
            //        {
            //            continue;
            //        }
            //        ICell sCell = sRow.GetCell(0);
            //        if (sCell == null)
            //        {
            //            continue;
            //        }
            //        if (tCell.ToString().Equals(sCell.ToString()))
            //        {
            //            for (int k = 1; k < sRow.LastCellNum; k++)
            //            {
            //                ICell tWCell = tRow.CreateCell(k+2);
            //                ICell sGCell = sRow.GetCell(k);
            //                if (tWCell == null || sGCell == null)
            //                {
            //                    break;
            //                }
            //                tWCell.SetCellValue(sGCell.ToString());
            //                CopyCellStyle(mainWorkbook, attendanceWorkbook, tCell, sCell);
            //            }
            //            break;
            //        }
            //    }
            //}








            FinishedMainWorkBook();
        }


        public static void CopyCellStyle(ICellStyle fromStyle, ICellStyle toStyle)
        {
            //toStyle.Alignment = fromStyle.Alignment;
            //toStyle.VerticalAlignment = fromStyle.VerticalAlignment;
            //toStyle.WrapText = fromStyle.WrapText;
            //////边框和边框颜色
            ///// 
            ///// 
            //toStyle.BorderBottom = fromStyle.BorderBottom;
            //toStyle.BorderLeft = fromStyle.BorderLeft;
            //toStyle.BorderRight = fromStyle.BorderRight;
            //toStyle.BorderTop = fromStyle.BorderTop;


            ////toStyle.TopBorderColor = fromStyle.TopBorderColor;
            ////toStyle.BottomBorderColor = fromStyle.BottomBorderColor;
            ////toStyle.RightBorderColor = fromStyle.RightBorderColor;
            ////toStyle.LeftBorderColor = fromStyle.LeftBorderColor;
            //////背景和前景
            //toStyle.FillBackgroundColor = GetColor(fromStyle.FillBackgroundColor); ;// fromStyle.FillBackgroundColor;
            //toStyle.FillForegroundColor = GetColor(fromStyle.FillForegroundColor);//fromStyle.FillForegroundColor;
            //toStyle.FillPattern = fromStyle.FillPattern;
            toStyle.FillPattern = FillPattern.SolidForeground;


            //toStyle.DataFormat = fromStyle.DataFormat;
            ////toStyle.Hidden=fromStyle.Hidden;
            //toStyle.IsHidden = fromStyle.IsHidden;
            //toStyle.Indention = fromStyle.Indention;//首行缩进
            //toStyle.IsLocked = fromStyle.IsLocked;
            //toStyle.Rotation = fromStyle.Rotation;//旋转
        }

        private static short GetColor(short fillForegroundColor)
        {
            short color;
            switch (fillForegroundColor)
            {
                case HSSFColor.Blue.Index:
                {
                    color = HSSFColor.Blue.Index;
                    break;
                }
                case HSSFColor.Green.Index:
                {
                    color = HSSFColor.Green.Index;
                    break;
                }
                case HSSFColor.Aqua.Index:
                {
                    color = HSSFColor.Aqua.Index;
                    break;
                }
                case HSSFColor.Black.Index:
                {
                    color = HSSFColor.Black.Index;
                    break;
                }
                case HSSFColor.BlueGrey.Index:
                {
                    color = HSSFColor.BlueGrey.Index;
                    break;
                }
                case HSSFColor.BrightGreen.Index:
                {
                    color = HSSFColor.BrightGreen.Index;
                    break;
                }
                case HSSFColor.Brown.Index:
                {
                    color = HSSFColor.Brown.Index;
                    break;
                }
                case HSSFColor.Coral.Index:
                {
                    color = HSSFColor.Coral.Index;
                    break;
                }
                case HSSFColor.CornflowerBlue.Index:
                {
                    color = HSSFColor.CornflowerBlue.Index;
                    break;
                }
                case HSSFColor.DarkBlue.Index:
                {
                    color = HSSFColor.DarkBlue.Index;
                    break;
                }
                case HSSFColor.DarkGreen.Index:
                {
                    color = HSSFColor.DarkGreen.Index;
                    break;
                }
                case HSSFColor.DarkRed.Index:
                {
                    color = HSSFColor.DarkRed.Index;
                    break;
                }
                case HSSFColor.DarkTeal.Index:
                {
                    color = HSSFColor.DarkTeal.Index;
                    break;
                }
                case HSSFColor.DarkYellow.Index:
                {
                    color = HSSFColor.DarkYellow.Index;
                    break;
                }
                case HSSFColor.Gold.Index:
                {
                    color = HSSFColor.Gold.Index;
                    break;
                }
                case HSSFColor.Grey25Percent.Index:
                {
                    color = HSSFColor.Grey25Percent.Index;
                    break;
                }
                case HSSFColor.Grey40Percent.Index:
                {
                    color = HSSFColor.Grey40Percent.Index;
                    break;
                }
                case HSSFColor.Grey50Percent.Index:
                {
                    color = HSSFColor.Grey50Percent.Index;
                    break;
                }
                case HSSFColor.Grey80Percent.Index:
                {
                    color = HSSFColor.Grey80Percent.Index;
                    break;
                }
                case HSSFColor.Yellow.Index:
                {
                    color = HSSFColor.Yellow.Index;
                    break;
                }
                case HSSFColor.White.Index:
                {
                    color = HSSFColor.White.Index;
                    break;
                }
                case HSSFColor.Violet.Index:
                {
                    color = HSSFColor.Violet.Index;
                    break;
                }
                case HSSFColor.Turquoise.Index:
                {
                    color = HSSFColor.Turquoise.Index;
                    break;
                }
                case HSSFColor.Teal.Index:
                {
                    color = HSSFColor.Teal.Index;
                    break;
                }
                case HSSFColor.Tan.Index:
                {
                    color = HSSFColor.Tan.Index;
                    break;
                }
                case HSSFColor.SkyBlue.Index:
                {
                    color = HSSFColor.SkyBlue.Index;
                    break;
                }
                case HSSFColor.SeaGreen.Index:
                {
                    color = HSSFColor.SeaGreen.Index;
                    break;
                }
                case HSSFColor.RoyalBlue.Index:
                {
                    color = HSSFColor.RoyalBlue.Index;
                    break;
                }
                case HSSFColor.Rose.Index:
                {
                    color = HSSFColor.Rose.Index;
                    break;
                }
                case HSSFColor.Red.Index:
                {
                    color = HSSFColor.Red.Index;
                    break;
                }
                default :
                {
                    Console.WriteLine("default: " + fillForegroundColor);
                    color = HSSFColor.Automatic.Index;
                    break;
                }
            }
            return color;
        }
        private void GetSalaryList(ISheet mainSheet)
        {
            FileStream salaryFileStream = new FileStream(@"e:/2018年3月发薪名单.xlsx", FileMode.Open);
            XSSFWorkbook salaryWorkbook = new XSSFWorkbook(salaryFileStream);
            XSSFSheet salarySheet = (XSSFSheet) salaryWorkbook.GetSheet("Sheet1");
            for (int i = 0; i < salarySheet.PhysicalNumberOfRows; i++)
            {
                XSSFRow tRow = (XSSFRow) mainSheet.CreateRow(i);
                XSSFRow sRow = (XSSFRow) salarySheet.GetRow(i);
                if (sRow != null && tRow != null)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        XSSFCell sCell = (XSSFCell) sRow.GetCell(j);
                        if (sCell == null)
                        {
                            break;
                        }
                        string cellValue = sCell.ToString();
                        XSSFCell tCell = (XSSFCell) tRow.CreateCell(j);
                        //CopyCellStyle(mainWorkbook, salaryWorkbook, tCell, sCell);

                        XSSFCellStyle style = (XSSFCellStyle) sCell.CellStyle;
                        XSSFCellStyle style1 = (XSSFCellStyle) mainWorkbook.CreateCellStyle();
                        XSSFColor color = null;
                        if (style.FillForegroundColorColor != null)
                        {
                            byte[] pa = style.FillForegroundColorColor.RGB;
                            string key = pa[0] + "," + pa[1] + "," + pa[2];
                            if (dictionary.ContainsKey(key))
                            {
                                style1.FillForegroundColor = dictionary[key];
                            }
                            else
                            {
                                Console.WriteLine("找不到该颜色!" + key);
                                style1.FillForegroundColor = HSSFColor.Automatic.Index;
                            }
                        }
                        else
                        {
                            Console.WriteLine("找不到该颜色!");
                            style1.FillForegroundColor = HSSFColor.Automatic.Index;
                        }
                        //byte[] pa1 = style.FillBackgroundColorColor.RGB;
                        //style1.FillForegroundColor = 20;//GetColor(sCell.CellStyle.FillForegroundColor);
                        style1.FillPattern = sCell.CellStyle.FillPattern;
                        //style1.FillBackgroundColor = 20;//GetColor(sCell.CellStyle.FillForegroundColor);
                        tCell.CellStyle = style1;
                        tCell.SetCellValue(cellValue);
                    }
                }
            }
            salaryFileStream.Close();
            salaryWorkbook.Close();
        }

        private void CopyCellStyle(IWorkbook toBook, IWorkbook fromBook, ICell toCell, ICell fromCell)
        {
            ICellStyle style = toBook.CreateCellStyle();
            IFont toFont = toBook.CreateFont();
            CopyCellStyle(fromCell.CellStyle, style);
            IFont fromFont = fromCell.CellStyle.GetFont(fromBook);
            CopyFont(toFont, fromFont);
            //style.SetFont(toFont);
            toCell.CellStyle = style;
        }

        private static void CopyFont(IFont toFont, IFont fromFont)
        {
            toFont.Color = fromFont.Color;
            toFont.FontHeightInPoints = fromFont.FontHeightInPoints;
            toFont.IsBold = fromFont.IsBold;
            toFont.IsItalic = fromFont.IsItalic;
            toFont.Underline = fromFont.Underline;
            //toFont.Boldweight = fromFont.Boldweight;
            //toFont.Charset = fromFont.Charset;
            //toFont.FontHeight = fromFont.FontHeight;
            toFont.FontName = fromFont.FontName;
            //toFont.IsStrikeout = fromFont.IsStrikeout;
            //toFont.TypeOffset = fromFont.TypeOffset;
        }

        private void FinishedMainWorkBook()
        {
            FileStream sw = File.Create(@"e:/MainWorkBook.xlsx");
            mainWorkbook.Write(sw);
            sw.Close();
            mainWorkbook.Close();
        }


        private void GenerateMainWorkBook()
        {
            mainWorkbook = new XSSFWorkbook();
            mainWorkbook.CreateSheet("MainSheet");
        }

        private void GetAllColor()
        {
            FileStream salaryFileStream = new FileStream(@"e:/考勤201803李静V1.xlsx", FileMode.Open);
            XSSFWorkbook salaryWorkbook = new XSSFWorkbook(salaryFileStream);
            XSSFSheet salarySheet = (XSSFSheet) salaryWorkbook.GetSheet("汇总表");
            Dictionary<string, string> dic = new Dictionary<string, string>();
            for (int i = 0; i < salarySheet.PhysicalNumberOfRows; i++)
            {
                IRow sRow = salarySheet.GetRow(i);
                if (sRow != null)
                {
                    for (int j = 0; j < sRow.LastCellNum; j++)
                    {
                        ICell sCell = sRow.GetCell(j);
                        if (sCell == null)
                        {
                            continue;
                        }
                        ICellStyle style = sCell.CellStyle;
                        if (style.FillForegroundColorColor != null)
                        {
                            byte[] pa = style.FillForegroundColorColor.RGB;
                            string rgb = pa[0] + "," + pa[1] + "," + pa[2];
                            if (!dic.ContainsKey(rgb))
                            {
                                dic.Add(rgb, sCell.ToString());
                            }
                            //Console.WriteLine("sCell.value: " + sCell.ToString() + pa[0] + " " + pa[1] + " " + pa[2]);
                        }
                    }
                }
            }
            salaryFileStream.Close();
            salaryWorkbook.Close();

            foreach (KeyValuePair<string, string> item in dic)
            {
                Console.WriteLine("sCell.value: "+item.Value + item.Key.ToString());
            }
        }

        ///   
        /// IRow Copy Command  
        ///  
        /// Description:  Inserts a existing row into a new row, will automatically push down  
        ///               any existing rows.  Copy is done cell by cell and supports, and the  
        ///               command tries to copy all properties available (style, merged cells, values, etc...)  
        ///   
        //private void CopyRow(IWorkbook workbook, ISheet worksheet, int sourceRowNum, int destinationRowNum)
        //{
        //    // Get the source / new row  
        //    IRow newRow = worksheet.GetRow(destinationRowNum);
        //    IRow sourceRow = worksheet.GetRow(sourceRowNum);

        //    // If the row exist in destination, push down all rows by 1 else create a new row  
        //    if (newRow != null)
        //    {
        //        worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1);
        //    }
        //    else
        //    {
        //        newRow = worksheet.CreateRow(destinationRowNum);
        //    }

        //    // Loop through source columns to add to new row  
        //    for (int i = 0; i < sourceRow.LastCellNum; i++)
        //    {
        //        // Grab a copy of the old/new cell  
        //        ICell oldCell = sourceRow.GetCell(i);
        //        ICell newCell = newRow.CreateCell(i);

        //        // If the old cell is null jump to next cell  
        //        if (oldCell == null)
        //        {
        //            newCell = null;
        //            continue;
        //        }

        //        // Copy style from old cell and apply to new cell  
        //        ICellStyle newCellStyle = workbook.CreateCellStyle();
        //        newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
        //        newCell.CellStyle = newCellStyle;

        //        // If there is a cell comment, copy  
        //        if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

        //        // If there is a cell hyperlink, copy  
        //        if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

        //        // Set the cell data type  
        //        newCell.SetCellType(oldCell.CellType);

        //        // Set the cell data value  
        //        switch (oldCell.CellType)
        //        {
        //            case ICellType.BLANK:
        //                newCell.SetCellValue(oldCell.StringCellValue);
        //                break;
        //            case ICellType.BOOLEAN:
        //                newCell.SetCellValue(oldCell.BooleanCellValue);
        //                break;
        //            case ICellType.ERROR:
        //                newCell.SetCellErrorValue(oldCell.ErrorCellValue);
        //                break;
        //            case ICellType.FORMULA:
        //                newCell.SetCellFormula(oldCell.CellFormula);
        //                break;
        //            case ICellType.NUMERIC:
        //                newCell.SetCellValue(oldCell.NumericCellValue);
        //                break;
        //            case ICellType.STRING:
        //                newCell.SetCellValue(oldCell.RichStringCellValue);
        //                break;
        //            case ICellType.Unknown:
        //                newCell.SetCellValue(oldCell.StringCellValue);
        //                break;
        //        }
        //    }

        //    // If there are are any merged regions in the source row, copy to new row  
        //    for (int i = 0; i < worksheet.NumMergedRegions; i++)
        //    {
        //        CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);
        //        if (cellRangeAddress.FirstRow == sourceRow.RowNum)
        //        {
        //            CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
        //                                                                        (newRow.RowNum +
        //                                                                         (cellRangeAddress.FirstRow -
        //                                                                          cellRangeAddress.LastRow)),
        //                                                                        cellRangeAddress.FirstColumn,
        //                                                                        cellRangeAddress.LastColumn);
        //            worksheet.AddMergedRegion(newCellRangeAddress);
        //        }
        //    }

        //}  
    }
}
