
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.IO
Imports System.Drawing

Partial Public Class _Default
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
		If Not IsPostBack Then
			convDir.DataSource = System.Enum.GetNames(GetType(SautinSoft.UseOffice.eDirection))
			convDir.DataBind()
		End If
		resultMessage.Text = ""
		fileMessage.Text = ""
    End Sub
    Protected Sub convDir_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        If uploadedDocument.PostedFile.FileName.Length = 0 Then
            resultMessage.Text = "Please select an input document at first!"
            Return
        End If


    End Sub
    Protected Sub convert_Click(ByVal sender As Object, ByVal e As EventArgs)
		If uploadedDocument.PostedFile.FileName.Length = 0 Then
			resultMessage.Text = "Please select an input document at first!"
			Return
		End If

		' 1. Prepare a directory for converting and storing temporary files
		' Remove all files in this directory
		Dim workDirectory As String = "/converted/"
		Dim workPath As String = Server.MapPath(".") & workDirectory
        Dim allFiles() As String = Directory.GetFiles(workPath, "*.*")
        For Each file As String In allFiles
            System.IO.File.Delete(file)
        Next file


		' 2. Save a document from FileUpload control to a temporary file: "Hour-Min-Sec-MSec.xxx", e.g. "9-34-12-807.doc".
		' Because UseOffice .Net can convert only documents as files
		Dim fileName As String = String.Format("{0: h-m-s-f}", Date.Now)
		Dim tempFilePath As String = Path.Combine(workPath, fileName & Path.GetExtension(uploadedDocument.PostedFile.FileName))
		File.WriteAllBytes(tempFilePath, uploadedDocument.FileBytes)


		resultMessage.Text = "Converting ..."


		' 3. Get a converting direction which user has just choosed
		' Get a file extension for a resulting file.
		Dim ext As String = ".rtf"
		'get extension
		Dim pos As Integer = convDir.Text.LastIndexOf("_"c)
		pos += 1
		ext = (convDir.Text.Substring(pos, convDir.Text.Length - pos)).ToLower()


		If ext.CompareTo("text") = 0 Then
			ext = "txt"
		End If
		ext = "." & ext

		' 4. Create a path for a resulted file.
		Dim resultPath As String = Path.Combine(workPath, Path.ChangeExtension(uploadedDocument.FileName, ext))

		' 5. Launch UseOffice .Net and start converting
		Dim u As New SautinSoft.UseOffice()

		'only Word										   
		If convDir.Text = "DOC_to_RTF" OrElse convDir.Text = "DOCX_to_RTF" OrElse convDir.Text = "DOC_to_HTML" OrElse convDir.Text = "DOCX_to_HTML" OrElse convDir.Text = "DOC_to_Text" OrElse convDir.Text = "HTML_to_DOC" OrElse convDir.Text = "HTML_to_RTF" OrElse convDir.Text = "HTML_to_Text" OrElse convDir.Text = "RTF_to_Text" OrElse convDir.Text = "RTF_to_HTML" OrElse convDir.Text = "RTF_to_DOC" OrElse convDir.Text = "DOC_to_PDF" OrElse convDir.Text = "DOCX_to_PDF" OrElse convDir.Text = "RTF_to_PDF" Then
			u.InitWord()
		End If
		'only Excel
		If convDir.Text = "XLS_to_HTML" OrElse convDir.Text = "XLS_to_XML" OrElse convDir.Text = "XLS_to_CSV" OrElse convDir.Text = "XLS_to_Text" Then
			u.InitExcel()
		End If
		'Word + Excel
		If convDir.Text = "XLS_to_RTF" OrElse convDir.Text = "XLS_to_PDF" OrElse convDir.Text = "XLSX_to_RTF" OrElse convDir.Text = "XLSX_to_PDF" Then
			u.InitWord()
			u.InitExcel()
		End If
		'only PowerPoint
		If convDir.Text = "PPT_to_PDF" OrElse convDir.Text = "PPT_to_HTML" OrElse convDir.Text = "PPT_to_RTF" OrElse convDir.Text = "PPTX_to_PDF" Then
			u.InitPowerPoint()
		End If
		Dim convDirection As SautinSoft.UseOffice.eDirection = CType(System.Enum.Parse(GetType(SautinSoft.UseOffice.eDirection), convDir.SelectedValue), SautinSoft.UseOffice.eDirection)

		If File.Exists(resultPath) Then
			File.Delete(resultPath)
		End If


		' 6. Convert a temporary file to a desired result
		Dim result As Integer = u.ConvertFile(tempFilePath, resultPath, convDirection)

		Select Case result
				' 7. Show a resulted file as a link
			Case 0
				resultMessage.Text = "Converting successfully!"
				Dim href As String = Request.UrlReferrer.AbsoluteUri
				href = href.Remove(href.LastIndexOf("/"))
				href &= workDirectory & Path.GetFileName(resultPath)
                fileMessage.NavigateUrl = href
                fileMessage.Target = "_blank"
				fileMessage.Text = Path.GetFileName(resultPath)

			Case 1
				resultMessage.Text = "Can't open input file."
			Case 2
				resultMessage.Text = "Can't create output file."
			Case 3
				resultMessage.Text = "Converting error!"
			Case Else
		End Select


		'u.KillProcesses("WINWORD");
		'u.KillProcesses("EXCEL");
		'u.KillProcesses("POWERPNT");
		u.CloseOffice()

		'Remove temporary file
        File.Delete(tempFilePath)
    End Sub
End Class