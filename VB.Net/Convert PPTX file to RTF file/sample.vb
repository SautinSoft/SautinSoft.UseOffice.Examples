Imports System
Imports System.IO
Imports System.Collections

Namespace Sample
    Friend Class Test
        Shared Sub Main(ByVal args() As String)
            ' Convert PPTX file to RTF file
            ' If you need more information about UseOffice .Net email us at:
            ' support@sautinsoft.com.
            Dim u As New SautinSoft.UseOffice()

            Dim inpFile As String = Path.GetFullPath("..\..\..\..\TestFiles\example.pptx")
            Dim outFile As String = Path.GetFullPath("Result.rtf")

            ' Prepare UseOffice .Net, loads MS PowerPoint in memory
            Dim ret As Integer = u.InitPowerPoint()

            ' Return values:
            ' 0 - Loading successfully
            ' 1 - Can't load MS PowerPoint library in memory
            If ret = 1 Then
                Console.WriteLine("Error! Can't load MS PowerPoint library in memory")
                Return
            End If

            ' Perform the conversion.
            ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.PPTX_to_RTF)

            ' Release MS PowerPoint from memory
            u.ClosePowerPoint()

            ' 0 - Converting successfully
            ' 1 - Can't open input file. Check that you are using full local path to input file, URL and relative path are not supported
            ' 2 - Can't create output file. Please check that you have permissions to write by this path or probably this path already used by another application
            ' 3 - Converting failed, please contact with our Support Team
            ' 4 - MS Office isn't installed. The component requires that any of these versions of MS Office should be installed: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.
            If ret = 0 Then
                Console.WriteLine("Converting successfully!")
                ' Open the result.
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})

            Else
                Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.")
            End If
        End Sub
    End Class
End Namespace
