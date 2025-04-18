﻿using System;
using System.IO;
using System.Collections.Generic;

namespace Sample
{
    class Test
    {
        static void Main(string[] args)
        {
            // Convert Word Excel PowerPoint documents to PDF format.
            // If you need more information about UseOffice .Net email us at:
            // support@sautinsoft.com

            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            // The directory which contains Word, Excel, PowerPoint files: *.doc, *.docx, *.rtf, *.txt, *.xls, *.xlsx, *.ppt, *.pptx
            string directoryWithFiles = Path.GetFullPath(@"..\..\..\Files");

            // Prepare UseOffice .Net, loads MS Word, Excel, PowerPoint into memory
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
                case 0: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx", "*.ppt", "*.pptx" }; break;
                case 1: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.ppt", "*.pptx" }; break;
                case 10: filters = new string[] { "*.xls", "*.xlsx", "*.ppt", "*.pptx" }; break;
                case 11: filters = new string[] { "*.ppt", "*.pptx" }; break;
                case 100: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx" }; break;
                case 101: filters = new string[] { "*.doc", "*.docx", "*.rtf", "*.txt" }; break;
                case 110: filters = new string[] { "*.xls", "*.xlsx" }; break;
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
            string ext;
            string outFilePath;
            string outDir = Path.GetFullPath(@"..\..\..\Results");
            Directory.CreateDirectory(outDir);

            for (int i = 0; i < inpFiles.Count; i++)
            {
                SautinSoft.UseOffice.eDirection direction = SautinSoft.UseOffice.eDirection.DOC_to_PDF;
                ext = Path.GetExtension(inpFiles[i]).ToLower();

                // doc and docx
                if (ext == ".doc")
                    direction = SautinSoft.UseOffice.eDirection.DOC_to_PDF;
                if (ext == ".docx")
                    direction = SautinSoft.UseOffice.eDirection.DOCX_to_PDF;
                else if (ext == ".rtf")
                    direction = SautinSoft.UseOffice.eDirection.RTF_to_PDF;
                else if (ext == ".txt")
                    direction = SautinSoft.UseOffice.eDirection.TEXT_to_PDF;

                // xls and xlsx
                else if (ext == ".xls")
                    direction = SautinSoft.UseOffice.eDirection.XLS_to_PDF;
                else if (ext == ".xlsx")
                    direction = SautinSoft.UseOffice.eDirection.XLSX_to_PDF;

                // ppt and pptx
                else if (ext == ".ppt")
                    direction = SautinSoft.UseOffice.eDirection.PPT_to_PDF;
                else if (ext == ".pptx")
                    direction = SautinSoft.UseOffice.eDirection.PPTX_to_PDF;

                // Save the result into the current directory
                string outFileName = (ext + "topdf.pdf").TrimStart('.');
                outFilePath = Path.Combine(outDir, outFileName);

                int conversion = u.ConvertFile(inpFiles[i], outFilePath, direction);

                Console.WriteLine($"{i + 1} of {inpFiles.Count}...");
            }
            Console.WriteLine("Done!");

            u.CloseOffice();

            // Open the folder (current directory) with the results.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outDir) { UseShellExecute = true });
        }
    }
}

