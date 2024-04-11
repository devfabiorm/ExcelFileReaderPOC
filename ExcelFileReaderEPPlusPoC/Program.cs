using OfficeOpenXml;

namespace ExcelFileReaderEPPlusPoC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var existingFile = new FileInfo("C:\\Users\\f.ribeiro.martins\\Desktop\\MassaTeste.xlsx");

            using var package = new ExcelPackage(existingFile);
            //Get the first worksheet in the workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    Console.Write(worksheet.Cells[row, col]?.Value + " | ");
                }

                Console.WriteLine();
            }
            Console.ReadLine();
        }
    }
}
