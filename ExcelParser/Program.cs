using ExcelDataReader;
using System;
using System.IO;

namespace ExcelParser
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Instead, pull from args.
            var filePath = "C:\\Users\\shawn\\code\\excelparser\\ExcelParser\\ExcelParser\\data\\test.xlsx";
            
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    var i = 1;
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                            Console.Write("Row: " + i);        
                            if (i == 1)
                            {
                                // handle headers
                                Console.Write(" " + reader.GetString(0));
                                Console.Write(" " + reader.GetString(1));
                            } else
                            {
                                // data row
                                Console.Write(" " + reader.GetDouble(0));
                                Console.Write(" " + reader.GetDateTime(1));
                            }
                            Console.WriteLine();
                            i++;
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                }
            }
        }
    }
}
