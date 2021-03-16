using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Drawing;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            convDir.DataSource = Enum.GetNames(typeof(SautinSoft.UseOffice.eDirection));
            convDir.DataBind();
        }
        resultMessage.Text = "";
        fileMessage.Text = "";
    }
    protected void convDir_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (uploadedDocument.PostedFile.FileName.Length == 0)
        {
            resultMessage.Text = "Please select an input document at first!";
            return;
        }


    }

    protected void convert_Click(object sender, EventArgs e)
    {
        if (uploadedDocument.PostedFile.FileName.Length == 0)
        {
            resultMessage.Text = "Please select an input document at first!";
            return;
        }

        // 1. Prepare a directory for converting and storing temporary files
        // Remove all files in this directory
        string workDirectory = @"/converted/";
        string workPath = Server.MapPath(".") + workDirectory;

        string[] allFiles = Directory.GetFiles(workPath, "*.*");
        foreach (string file in allFiles)
            File.Delete(file);
        


        // 2. Save a document from FileUpload control to a temporary file: "Hour-Min-Sec-MSec.xxx", e.g. "9-34-12-807.doc".
        // Because UseOffice .Net can convert only documents as files
        string fileName = String.Format("{0: h-m-s-f}", DateTime.Now);
        string tempFilePath = Path.Combine(workPath, fileName + Path.GetExtension(uploadedDocument.PostedFile.FileName));
        File.WriteAllBytes(tempFilePath, uploadedDocument.FileBytes);


        resultMessage.Text = "Converting ...";


        // 3. Get a converting direction which user has just choosed
        // Get a file extension for a resulting file.
        string ext = ".rtf";
        //get extension
        int pos = convDir.Text.LastIndexOf('_');
        pos++;
        ext = (convDir.Text.Substring(pos, convDir.Text.Length - pos)).ToLower();


        if (ext.CompareTo("text") == 0)
        {
            ext = "txt";
        }
        ext = "." + ext;

        // 4. Create a path for a resulted file.
        string resultPath = Path.Combine(workPath, Path.ChangeExtension(uploadedDocument.FileName, ext));
        
        // 5. Launch UseOffice .Net and start converting
        SautinSoft.UseOffice u = new SautinSoft.UseOffice();

        //only Word										   
        if (convDir.Text == "DOC_to_RTF" || convDir.Text == "DOCX_to_RTF" || convDir.Text == "DOC_to_HTML" || convDir.Text == "DOCX_to_HTML" || convDir.Text == "DOC_to_Text" || convDir.Text == "HTML_to_DOC" || convDir.Text == "HTML_to_RTF" || convDir.Text == "HTML_to_Text" || convDir.Text == "RTF_to_Text" || convDir.Text == "RTF_to_HTML" || convDir.Text == "RTF_to_DOC" || convDir.Text == "DOC_to_PDF" || convDir.Text == "DOCX_to_PDF" || convDir.Text == "RTF_to_PDF")
        {
            u.InitWord();
        }
        //only Excel
        if (convDir.Text == "XLS_to_HTML" || convDir.Text == "XLS_to_XML" || convDir.Text == "XLS_to_CSV" || convDir.Text == "XLS_to_Text")
        {
            u.InitExcel();
        }
        //Word + Excel
        if (convDir.Text == "XLS_to_RTF" || convDir.Text == "XLS_to_PDF" || convDir.Text == "XLSX_to_RTF" || convDir.Text == "XLSX_to_PDF")
        {
            u.InitWord();
            u.InitExcel();
        }
        //only PowerPoint
        if (convDir.Text == "PPT_to_PDF" || convDir.Text == "PPT_to_HTML" || convDir.Text == "PPT_to_RTF" || convDir.Text == "PPTX_to_PDF")
        {
            u.InitPowerPoint();
        }
        SautinSoft.UseOffice.eDirection convDirection = (SautinSoft.UseOffice.eDirection)Enum.Parse(typeof(SautinSoft.UseOffice.eDirection), convDir.SelectedValue);

        if (File.Exists(resultPath))
        {
            File.Delete(resultPath);
        }


        // 6. Convert a temporary file to a desired result
        int result = u.ConvertFile(tempFilePath, resultPath, convDirection);        
        
        switch (result)
        {
                // 7. Show a resulted file as a link
            case 0: resultMessage.Text = "Converting successfully!";
                string href = Request.UrlReferrer.AbsoluteUri;
                href = href.Remove(href.LastIndexOf("/"));                
                href+= workDirectory + Path.GetFileName(resultPath);
                fileMessage.NavigateUrl = href;
                fileMessage.Target = "_blank";
                fileMessage.Text = Path.GetFileName(resultPath);
                break;

            case 1: resultMessage.Text = "Can't open input file."; break;
            case 2: resultMessage.Text = "Can't create output file."; break;
            case 3: resultMessage.Text = "Converting error!"; break;
            default: break;
        }
        

        //u.KillProcesses("WINWORD");
        //u.KillProcesses("EXCEL");
        //u.KillProcesses("POWERPNT");
        u.CloseOffice();

        //Remove temporary file
        File.Delete(tempFilePath);
    }    
}
