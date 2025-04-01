Imports System
Imports System.IO
Imports System.Collections

Namespace Sample
    Friend Class Test
        Shared Sub Main(ByVal args() As String)
		
			' Before starting, we recommend to get a free key:
            ' https://sautinsoft.com/start-for-free/
            
            ' Apply the key here:
			' UseOffice.SetLicense("...");
            ' Convert DOCX to PDF in memory.
            ' If you need more information about UseOffice .Net email us at:
            ' support@sautinsoft.com.

            Dim u As New SautinSoft.UseOffice()

            ' We need files to read data from it and demostrate the result of conversion.
            Dim inpFile As String = Path.GetFullPath("..\..\..\..\..\..\TestFiles\example.docx")
            Dim outFile As String = Path.GetFullPath("Result.pdf")

            ' Prepare UseOffice .Net, loads MS Word in memory
            Dim ret As Integer = u.InitWord()

            ' Return values:
            ' 0 - Loading successfully
            ' 1 - Can't load MS Word library in memory 
            If ret = 1 Then
                Console.WriteLine("Error! Can't load MS Word library in memory")
                Return
            End If

            ' Perform the conversion.
            Dim docxBytes() As Byte = File.ReadAllBytes(inpFile)
            Dim pdfBytes() As Byte = Nothing

            ' If you are making the conversion on a server, please specify this temporary 
            ' directory and set read/write permissions on it.
            ' You may set any path.
            u.TemporaryDirectory = Path.GetTempPath()
            pdfBytes = u.ConvertBytes(docxBytes, SautinSoft.UseOffice.eDirection.DOCX_to_PDF)

            ' Release MS Word from memory
            u.CloseWord()

            ' 0 - Converting successfully
            ' 1 - Can't open input file. Check that you are using full local path to input file, URL and relative path are not supported
            ' 2 - Can't create output file. Please check that you have permissions to write by this path or probably this path already used by another application
            ' 3 - Converting failed, please contact with our Support Team
            ' 4 - MS Office isn't installed. The component requires that any of these versions of MS Office should be installed: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.
            If pdfBytes IsNot Nothing Then
                ' Open the result.
                File.WriteAllBytes(outFile, pdfBytes)
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
            Else
                Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.")
            End If
        End Sub

    End Class
End Namespace
