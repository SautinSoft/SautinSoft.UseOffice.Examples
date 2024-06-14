Imports System
Imports System.IO
Imports System.Collections.Generic

Namespace Sample
    Friend Class Test

        Shared Sub Main(ByVal args() As String)
            ' Convert Word Excel PowerPoint documents to PDF format.
            ' If you need more information about UseOffice .Net email us at:
            ' support@sautinsoft.com.
            Dim u As New SautinSoft.UseOffice()

            ' The directory which contains Word, Excel, PowerPoint files: *.doc, *.docx, *.rtf, *.txt, *.xls, *.xlsx, *.csv, *.ppt, *.pptx
            Dim directoryWithFiles As String = Path.GetFullPath("..\..\..\..\..\..\TestFiles\")

            'Prepare UseOffice .Net, loads MS Word, Excel, PowerPoint into memory
            Dim ret As Integer = u.InitOffice()

            ' Return values:
            ' 0 - Loading successfully
            ' 1 - Can't load MS Excel (Word and PowePoint are loaded successfully)
            ' 10 - Can't load MS Word (Excel and PowerPoint are loaded successfully)
            ' 11 - Can't load MS Word and Excel (PowerPoint loaded successfully)
            ' 100 - Can't load MS PowerPoint (Excel and Word are loaded successfully)
            ' 101 - Can't load MS Excel and PowerPoint (Word loaded successfully)
            ' 110 - Can't load PowerPoint and Word (Excel loaded successfully)
            ' 111 - Can't load MS Office

            If ret = 111 Then
                Return
            End If

            Dim filters() As String = Nothing

            Select Case ret
                Case 0
                    filters = New String() {"*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx", "*.csv", "*.ppt", "*.pptx"}
                Case 1
                    filters = New String() {"*.doc", "*.docx", "*.rtf", "*.txt", "*.ppt", "*.pptx"}
                Case 10
                    filters = New String() {"*.xls", "*.xlsx", "*.csv", "*.ppt", "*.pptx"}
                Case 11
                    filters = New String() {"*.ppt", "*.pptx"}
                Case 100
                    filters = New String() {"*.doc", "*.docx", "*.rtf", "*.txt", "*.xls", "*.xlsx", "*.csv"}
                Case 101
                    filters = New String() {"*.doc", "*.docx", "*.rtf", "*.txt"}
                Case 110
                    filters = New String() {"*.xls", "*.xlsx", "*.csv"}
                Case Else
                    Return
            End Select

            ' Convert all documents (Word, Excel, PorwerPoint) to PDF.

            ' 1. Get list of MS Office files from directory
            Dim inpFiles As New List(Of String)()

            For Each filter As String In filters
                inpFiles.AddRange(Directory.GetFiles(directoryWithFiles, filter))
            Next filter

            ' 2. Convert all documents to PDF.
            Dim ext As String = ""
            Dim outFilePath As String = ""
            Dim outDir As DirectoryInfo = (New DirectoryInfo(Directory.GetCurrentDirectory())).CreateSubdirectory("Results")

            For i As Integer = 0 To inpFiles.Count - 1
                Dim direction As SautinSoft.UseOffice.eDirection = SautinSoft.UseOffice.eDirection.DOC_to_PDF
                ext = Path.GetExtension(CStr(inpFiles(i))).ToLower()

                ' doc and docx
                If ext.IndexOf("doc") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.DOC_to_PDF
                ElseIf ext.IndexOf("rtf") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.RTF_to_PDF
                ElseIf ext.IndexOf("txt") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.TEXT_to_PDF
                    ' xls and xlsx
                ElseIf ext.IndexOf("xls") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.XLS_to_PDF
                ElseIf ext.IndexOf("csv") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.XLS_to_PDF
                    ' ppt and pptx
                ElseIf ext.IndexOf("ppt") > 0 Then
                    direction = SautinSoft.UseOffice.eDirection.PPT_to_PDF
                End If

                ' Save the result into the current directory
                Dim outFileName As String = (Path.GetExtension(inpFiles(i)) & "topdf.pdf").TrimStart("."c)
                outFilePath = Path.Combine(outDir.FullName, outFileName)

                u.ConvertFile(CStr(inpFiles(i)), outFilePath, direction)

                Console.WriteLine($"{i + 1} of {inpFiles.Count}...")
            Next i
            Console.WriteLine("Done!")

            u.CloseOffice()

            ' Open the folder (current directory) with the results.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outDir.FullName) With {.UseShellExecute = True})

        End Sub
    End Class
End Namespace
