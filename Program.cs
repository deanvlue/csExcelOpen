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
            string fileName = Path.GetFileName(filename);
            string fileExtension = Path.GetExtension(filename);

            return dirName + fileName + "_mod" + fileExtension;

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
            //string newFile = GetNewFileName(_fileName);


            HSSFWorkbook hssfwb;
            using (FileStream excelFile = new FileStream(_fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(excelFile);
            }

            ISheet repote = hssfwb.GetSheetAt(0);
            for (int row = 0; row ) 







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