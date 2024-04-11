using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

namespace ExcelFileReaderPOC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // load an existing workbook
            IWorkbook wb = new XSSFWorkbook("C:\\Users\\f.ribeiro.martins\\Desktop\\MassaTeste.xlsx");

            // get the first worksheet
            ISheet ws = wb.GetSheetAt(0);

            for (int i = 1; i < ws.LastRowNum; i++)
            {
                // get the first row
                IRow row = ws.GetRow(i);

                for (int j = 0; j < row.LastCellNum; j++)
                {
                    // get the first cell
                    ICell cell = row.GetCell(j);
                    // get the cell value
                    var cellValue = GetCellValue(cell);

                    Console.Write(cellValue + " | ");
                }
                Console.WriteLine();
            }
            Console.ReadLine();
        }

        private static object GetCellValue(ICell cell)
        {
            object cellValue;

            if (cell == null)
            {
                return null;
            }

            if (cell.CellType == CellType.Numeric)
                if (DateUtil.IsCellDateFormatted(cell))
                    cellValue = cell.DateCellValue;
                else
                    cellValue = cell.NumericCellValue;
            else if (cell.CellType == CellType.Boolean)
                cellValue = cell.BooleanCellValue;
            else
                cellValue = cell.StringCellValue;

            return cellValue;
        }
    }
}
