using System;
using System.IO;
using System.Collections.Generic;

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
			
            // Convert Word Excel PowerPoint documents to PDF format.
            // If you need more information about UseOffice .Net email us at:
            // support@sautinsoft.com.
            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            // The directory which contains Word, Excel, PowerPoint files: *.doc, *.docx, *.rtf, *.txt, *.xls, *.xlsx, *.csv, *.ppt, *.pptx
            string directoryWithFiles = Path.GetFullPath(@"..\..\..\..\..\..\TestFiles\");

            //Prepare UseOffice .Net, loads MS Word, Excel, PowerPoint into memory
            int ret = u.InitOffice();

            // Return values:
            // 0 - Loading successfully
            // 1 - Can't load MS Excel (Word and PowePoint are loaded successfully)
            // 10 - Can't load MS Word (Excel and PowerPoint are loaded successfully)
            // 11 - Can't load MS Word and Excel (PowerPoint loaded successfully)
            // 100 - Can't load MS PowerPoint (Excel and Word are loaded successfully)
            // 101 - Can't load MS Excel and PowerPoint (Word loaded successfully)
            // 110 - Can't load PowerPoint and Word (Excel loaded successfully)
            // 111 - Can't load MS Office

            if (ret == 111)
                return;

            string[] filters = null;

            switch (ret)
            {
                case 0: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx", "*.csv", "*.ppt", "*.pptx" }; break;
                case 1: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.ppt", "*.pptx" }; break;
                case 10: filters = new string[] { "*.xls", "*.xlsx", "*.csv", "*.ppt", "*.pptx" }; break;
                case 11: filters = new string[] { "*.ppt", "*.pptx" }; break;
                case 100: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx", "*.csv" }; break;
                case 101: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt" }; break;
                case 110: filters = new string[] { "*.xls", "*.xlsx", "*.csv" }; break;
                default: return;
            }

            // Convert all documents (Word, Excel, PorwerPoint) to PDF.

            // 1. Get list of MS Office files from directory
            List<string> inpFiles = new List<string>();

            foreach (string filter in filters)
            {
                inpFiles.AddRange(Directory.GetFiles(directoryWithFiles, filter));
            }

            // 2. Convert all documents to PDF.
            string ext = "";
            string outFilePath = "";
            DirectoryInfo outDir = new DirectoryInfo(Directory.GetCurrentDirectory()).CreateSubdirectory("Results");

            for (int i = 0; i < inpFiles.Count; i++)
            {
                SautinSoft.UseOffice.eDirection direction = SautinSoft.UseOffice.eDirection.DOC_to_PDF;
                ext = Path.GetExtension((string)inpFiles[i]).ToLower();

                // doc and docx
                if (ext.IndexOf("doc") > 0)
                    direction = SautinSoft.UseOffice.eDirection.DOC_to_PDF;
                else if (ext.IndexOf("rtf") > 0)
                    direction = SautinSoft.UseOffice.eDirection.RTF_to_PDF;
                else if (ext.IndexOf("txt") > 0)
                    direction = SautinSoft.UseOffice.eDirection.TEXT_to_PDF;
                // xls and xlsx
                else if (ext.IndexOf("xls") > 0)
                    direction = SautinSoft.UseOffice.eDirection.XLS_to_PDF;
                else if (ext.IndexOf("csv") > 0)
                    direction = SautinSoft.UseOffice.eDirection.XLS_to_PDF;
                // ppt and pptx
                else if (ext.IndexOf("ppt") > 0)
                    direction = SautinSoft.UseOffice.eDirection.PPT_to_PDF;

                // Save the result into the current directory
                string outFileName = (Path.GetExtension(inpFiles[i]) + "topdf.pdf").TrimStart('.');
                outFilePath = Path.Combine(outDir.FullName, outFileName);

                u.ConvertFile((string)inpFiles[i], outFilePath, direction);

                Console.WriteLine($"{i + 1} of {inpFiles.Count}...");
            }
            Console.WriteLine("Done!");

            u.CloseOffice();

            // Open the folder (current directory) with the results.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outDir.FullName) { UseShellExecute = true });

        }
    }
}
