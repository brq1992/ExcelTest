using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace WindowsFormsApplication1
{


    class WorkBookManager
    {
        /// <summary>
        /// 主表
        /// </summary>
        private IWorkbook mainWorkbook;

        public WorkBookManager()
        {
            GenerateMainWorkBook();
            ISheet mainSheet = mainWorkbook.GetSheet("MainSheet");

            GetSalaryList(mainSheet);

            FileStream atdFileStream = new FileStream(@"e:/上海嵩恒网络科技股份有限公司_考勤报表_20180401-20180412.xlsx", FileMode.Open);
            IWorkbook attendanceWorkbook = new XSSFWorkbook(atdFileStream);
            atdFileStream.Close();
            ISheet atdSheet = attendanceWorkbook.GetSheet("打卡时间");
            //for (int i = 2; i < atdSheet.LastRowNum; i++)
            //{
            //    IRow sRow = atdSheet.GetRow(i);

            //    IRow tRow = mainSheet.GetRow(i-2);
            //    if (sRow != null && tRow != null)
            //    {
            //        for (int j = 0; j < sRow.LastCellNum; j++)
            //        {
            //            ICell tCell = tRow.CreateCell(j+2);
            //            ICell sCell = sRow.GetCell(j);
            //            if (tCell == null || sCell == null)
            //            {
            //                break;
            //            }
            //            tCell.SetCellValue(sCell.ToString());
            //            tCell.SetCellFormula("VLOOKUP()");
            //        }
            //    }
            //}

            for (int i = 0; i < mainSheet.LastRowNum; i++)
            {
                IRow tRow = mainSheet.GetRow(i);
                if (tRow == null)
                {
                    continue;
                }
                ICell tCell = tRow.GetCell(0);
                if (tCell == null)
                {
                    continue;
                }
                for (int j = 0; j < atdSheet.LastRowNum; j++)
                {
                    IRow sRow = atdSheet.GetRow(j);
                    if (sRow == null)
                    {
                        continue;
                    }
                    ICell sCell = sRow.GetCell(0);
                    if (sCell == null)
                    {
                        continue;
                    }
                    if (tCell.ToString().Equals(sCell.ToString()))
                    {
                        for (int k = 1; k < sRow.LastCellNum; k++)
                        {
                            ICell tWCell = tRow.CreateCell(k+2);
                            ICell sGCell = sRow.GetCell(k);
                            if (tWCell == null || sGCell == null)
                            {
                                break;
                            }
                            tWCell.SetCellValue(sGCell.ToString());
                        }
                        break;
                    }
                }
            }








            FinishedMainWorkBook();
        }

        private void GetSalaryList(ISheet mainSheet)
        {
            FileStream salaryFileStream = new FileStream(@"e:/2018年3月发薪名单.xlsx", FileMode.Open);
            IWorkbook salaryWorkbook = new XSSFWorkbook(salaryFileStream);
            ISheet salarySheet = salaryWorkbook.GetSheet("Sheet1");
            IRow sRow;
            IRow tRow;
            ICell tCell;
            ICell sCell;
            for (int i = 0; i < salarySheet.PhysicalNumberOfRows; i++)
            {
                tRow = mainSheet.CreateRow(i);
                sRow = salarySheet.GetRow(i);
                if (sRow != null && tRow != null)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        sCell = sRow.GetCell(j);
                        if (sCell == null)
                        {
                            break;
                        }
                        string cellValue = sCell.ToString();
                        // Console.WriteLine(cellValue);  
                        tCell = tRow.CreateCell(j);
                        tCell.SetCellValue(cellValue);
                    }
                }
            }
            salaryFileStream.Close();
            salaryWorkbook.Close();
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
