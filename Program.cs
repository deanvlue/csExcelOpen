using System;
using System.IO;
using NPOI;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using NPOI.XWPF.UserModel;

namespace ExcelOpen
{
    class Program
    {

        private static string GetNewFileName(string filename)
        {
            string dirName = Path.GetDirectoryName(filename);
            string fileExtension = Path.GetExtension(filename);
            string newfileName = Path.GetFileNameWithoutExtension(filename) + "_mod" + fileExtension;

            return Path.Join(dirName, newfileName);
            //return dirName + fileName + "_mod" + fileExtension;

        }
        static void Main(string[] args)
        {

            if (args.Length == 0)
            {
                Console.WriteLine("Uso: siebelreport.exe [archivo_excel.xls]");
                System.Environment.Exit(1);
            }
            string _fileName;
            _fileName = args[0];

            if (!_fileName.Contains("xlsx"))
            {
                Console.WriteLine("Uso: siebelreport.exe [archivo_excel.xls]");
                Console.WriteLine("El archivo debe ser un archivo de Excel");
                System.Environment.Exit(1);
            }

            //Console.WriteLine(_fileName);
            //Console.ReadLine();
            string newFile = GetNewFileName(_fileName);

            // determine filetype first
            XSSFWorkbook hssfwb;
            using (FileStream excelFile = new FileStream(_fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(excelFile);
                ISheet reporte = hssfwb.GetSheetAt(0);


                using (FileStream newExcelFile = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook newWorkBook = new XSSFWorkbook();
                    //newWorkBook = new XSSFWorkbook(newExcelFile);
                    ISheet nuevoReporte = newWorkBook.CreateSheet("REPORTE");
                    IRow headerRow = nuevoReporte.CreateRow(0);

                    // Creates Header Row
                    headerRow.CreateCell(0).SetCellValue("DIA");
                    headerRow.CreateCell(1).SetCellValue("ACUMULACIONES");
                    headerRow.CreateCell(2).SetCellValue("REDENCIONES");
                    headerRow.CreateCell(3).SetCellValue("ID_TIENDA");


                    for (int row = 1; row <= reporte.LastRowNum; row++)
                    {
                        if (reporte.GetRow(row) != null) // row is when the row only conatains empty cells
                        {
                            //string fechaEvento = reporte.GetRow(row).GetCell(0).StringCellValue + "/2019";
                            string[] fecha = reporte.GetRow(row).GetCell(0).StringCellValue.Split("/");

                            uint idTienda;
                            idTienda = Convert.ToUInt32(reporte.GetRow(row).GetCell(1).StringCellValue);

                            //reporte.GetRow(row).GetCell(2).GetType();

                            double acumulaciones = reporte.GetRow(row).GetCell(2).NumericCellValue;
                            double redenciones = reporte.GetRow(row).GetCell(3).NumericCellValue;
                            //Console.WriteLine("{0},{1},{2},{3}", fechaEvento, acumulaciones, redenciones, idTienda);

                            IRow rowNewReport = nuevoReporte.CreateRow(row);
                            //rowNewReport.CreateCell(0).SetCellValue(fechaEvento);
                            var dateCell = rowNewReport.CreateCell(0);
                                dateCell.SetCellType(CellType.Formula);
                            dateCell.CellFormula = string.Format("DATE(2019,{1},{0})", fecha[0], fecha[1]);
                                
                            rowNewReport.CreateCell(1).SetCellValue(acumulaciones);
                            rowNewReport.CreateCell(2).SetCellValue(redenciones);
                            rowNewReport.CreateCell(3).SetCellValue(idTienda);
                        }

                    }
                    Console.WriteLine("Escribiendo archivo...");
                    newWorkBook.Write(newExcelFile);
                }
            }



            /*
                * Validate input is an excel file
                * get the filename by spliting the argument in filename and extension

                * open an get the index 0 sheet
                * add 
             */



            //Console.WriteLine("Hello World!");
            /* var newFile = @"newbook.core.xlsx";

            using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet1 = workbook.CreateSheet("Sheet1");

                sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
                var rowIndex = 0;
                IRow row = sheet1.CreateRow(rowIndex);
                row.Height = 30 * 80;
                row.CreateCell(0).SetCellValue("This is sparta");
                sheet1.AutoSizeColumn(0);
                rowIndex++;

                var sheet2 = workbook.CreateSheet("HOJA2");

                var style1 = workbook.CreateCellStyle();
                style1.FillForegroundColor = HSSFColor.Blue.Index2;
                style1.FillPattern = FillPattern.SolidForeground;

                var style2 = workbook.CreateCellStyle();
                style2.FillForegroundColor = HSSFColor.Blue.Index2;
                style2.FillPattern = FillPattern.SolidForeground;

                var cell2 = sheet2.CreateRow(1).CreateCell(0);
                cell2.CellStyle = style1;
                cell2.SetCellValue(0);

                cell2 = sheet2.CreateRow(1).CreateCell(0);
                cell2.CellStyle = style2;
                cell2.SetCellValue(1);

                workbook.Write(fs);

            }
        }

            static private string JoinString(string[] array)
            {
                string result = string.Join(",", array);
                return result;
            }
        }*/

        }
    }
}