Imports System
Imports System.IO
Imports System.Collections
Imports SautinSoft

Namespace Sample
    Friend Class Test
        Shared Sub Main(ByVal args() As String)
            ' Convert PDF file to Text file. Works only in Office 2013 and higher.

            ' If you are looking for solution without MS Office
            ' Please take a look at our PDF Focus .Net: https://www.sautinsoft.com/products/pdf-focus/index.php

            Dim u As New SautinSoft.UseOffice()

            Dim inpFile As String = Path.GetFullPath("..\..\..\..\TestFiles\example.pdf")
            Dim outFile As String = Path.GetFullPath("Result.txt")

            ' Prepare UseOffice .Net, loads MS Word in memory
            If u.InitWord() <> 0 Then
                Console.WriteLine("Error: Can't load MS Word in memory!")
                Console.WriteLine("Please contact SautinSoft's support Team: support@sautinsoft.com.")
                Console.ReadLine()
            End If

            ' Check MS Office version
            If u.OfficeVersion >= UseOffice.eOfficeVersion.Office2013 Then
                ' Converting ...
                Dim result As Integer = u.ConvertFile(inpFile, outFile, UseOffice.eDirection.PDF_to_TEXT)

                If result = 0 Then
                    Console.WriteLine("Converting successfully!")
                    ' Open the result.
                    System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})

                Else
                    Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.")
                End If
            Else
                Console.WriteLine("To convert PDF documents, please install MS Office 2013 or higher.")
            End If
            u.CloseOffice()
        End Sub
    End Class
End Namespace
