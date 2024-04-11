using ExcelDataReader;
using Microsoft.VisualBasic;

namespace ExcelFileReaderEDRPoC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            const string filePath = "C:\\Users\\f.ribeiro.martins\\Desktop\\MassaTeste.xlsx";

            //Config to exclude header
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true,
                }
            };

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            // Auto-detect format, supports:
            //  - Binary Excel files (2.0-2003 format; *.xls)
            //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataSet = reader.AsDataSet(conf);
            var dataTable = dataSet.Tables[0];

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    Console.Write(dataTable.Rows[i][j] + " | ");
                }
                Console.WriteLine();
            }

            Console.ReadLine();
        }
    }
}
