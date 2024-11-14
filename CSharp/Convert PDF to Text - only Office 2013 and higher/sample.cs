using System;
using System.IO;
using System.Collections;
using SautinSoft;

namespace Sample
{
    class Test
    {
        static void Main(string[] args)
        {
			// Before starting, we recommend to get a free 100-day key:
            // https://sautinsoft.com/start-for-free/
            
            // Apply the key here:
			// UseOffice.SetLicense("...");
			
            // Convert PDF file to Text file. Works only in Office 2013 and higher.

            // If you are looking for solution without MS Office
            // Please take a look at our PDF Focus .Net: https://www.sautinsoft.com/products/pdf-focus/index.php

            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            string inpFile = Path.GetFullPath(@"..\..\..\..\..\..\TestFiles\example.pdf");
            string outFile = Path.GetFullPath("Result.txt");

            // Prepare UseOffice .Net, loads MS Word in memory
            if (u.InitWord() != 0)
            {
                Console.WriteLine("Error: Can't load MS Word in memory!");
                Console.WriteLine("Please contact SautinSoft's support Team: support@sautinsoft.com.");
                Console.ReadLine();
            }

            // Check MS Office version
            if (u.OfficeVersion >= UseOffice.eOfficeVersion.Office2013)
            {
                // Converting ...
                int result = u.ConvertFile(inpFile, outFile, UseOffice.eDirection.PDF_to_TEXT);

                if (result == 0)
                {
                    Console.WriteLine("Converting successfully!");
                    // Open the result.
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

                }
                else
                    Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.");
            }
            else
            {
                Console.WriteLine("To convert PDF documents, please install MS Office 2013 or higher.");
            }
            u.CloseOffice();
        }
    }
}