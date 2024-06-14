using System;
using System.IO;
using System.Collections;

namespace Sample
{
    class Test
    {
        static void Main(string[] args)
        {
            // Convert DOCX to PDF in memory.
            // If you need more information about UseOffice .Net email us at:
            // support@sautinsoft.com.

            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            // We need files to read data from it and demostrate the result of conversion.
            string inpFile = Path.GetFullPath(@"..\..\..\..\..\..\TestFiles\example.docx");
            string outFile = Path.GetFullPath("Result.pdf");

            // Prepare UseOffice .Net, loads MS Word in memory
            int ret = u.InitWord();

            // Return values:
            // 0 - Loading successfully
            // 1 - Can't load MS Word library in memory 
            if (ret == 1)
            {
                Console.WriteLine("Error! Can't load MS Word library in memory");
                return;
            }

            // Perform the conversion.
            byte[] docxBytes = File.ReadAllBytes(inpFile);
            byte[] pdfBytes = null;

            // If you are making the conversion on a server, please specify this temporary 
            // directory and set read/write permissions on it.
            // You may set any path.
            u.TemporaryDirectory = Path.GetTempPath();
            pdfBytes = u.ConvertBytes(docxBytes, SautinSoft.UseOffice.eDirection.DOCX_to_PDF);

            // Release MS Word from memory
            u.CloseWord();

            // 0 - Converting successfully
            // 1 - Can't open input file. Check that you are using full local path to input file, URL and relative path are not supported
            // 2 - Can't create output file. Please check that you have permissions to write by this path or probably this path already used by another application
            // 3 - Converting failed, please contact with our Support Team
            // 4 - MS Office isn't installed. The component requires that any of these versions of MS Office should be installed: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.
            if (pdfBytes != null)
            {
                // Open the result.
                File.WriteAllBytes(outFile, pdfBytes);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
            }
            else
                Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.");
        }

    }
}