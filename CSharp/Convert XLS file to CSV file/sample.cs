using System;
using System.IO;
using System.Collections;

namespace Sample
{
    class Test
    {
        static void Main(string[] args)
        {
			// Before starting, we recommend to get a free key:
            // https://sautinsoft.com/start-for-free/
            
            // Apply the key here:
			// UseOffice.SetLicense("...");
			
            // Convert XLS file to CSV file.
            // If you need more information about UseOffice .Net email us at:
            // support@sautinsoft.com.

            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            string inpFile = Path.GetFullPath(@"..\..\..\..\..\..\TestFiles\example.xls");
            string outFile = Path.GetFullPath("Result.csv");

            // Prepare UseOffice .Net, loads MS Excel in memory
            int ret = u.InitExcel();

            // Return values:
            // 0 - Loading successfully
            // 1 - Can't load MS Excel library in memory 

            if (ret == 1)
            {
                Console.WriteLine("Error! Can't load MS Excel library in memory");
                return;
            }

            // Perform the conversion.
            ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.XLS_to_CSV);

            // Release MS Excel from memory
            u.CloseExcel();

            // 0 - Converting successfully
            // 1 - Can't open input file. Check that you are using full local path to input file, URL and relative path are not supported
            // 2 - Can't create output file. Please check that you have permissions to write by this path or probably this path already used by another application
            // 3 - Converting failed, please contact with our Support Team
            // 4 - MS Office isn't installed. The component requires that any of these versions of MS Office should be installed: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.
            if (ret == 0)
            {
                // Open the result.
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
            }
            else
                Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.");
        }

    }
}