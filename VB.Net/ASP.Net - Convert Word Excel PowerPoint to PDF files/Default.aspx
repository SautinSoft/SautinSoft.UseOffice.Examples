<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>UseOffice .Net - You will likely be surprised at the amount of conversion ways!</title>
    <style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;	
    font-size: 20px;
}
        .auto-style1 {
            height: 41px;
        }
-->
    </style>
</head>
<body style="margin-left: 20px;">
<form id="form1" runat="server">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="2" class="auto-style1"><span class="style1"><strong>UseOffice .Net</strong> <span style="font-size: 10pt"> - Requires MS Office installed, any version: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.</span></span></td>
      </tr>
      <tr>
        <td width="247"><img src="box_useoffice_net.jpg" width="247" height="250" alt=""/></td>
        <td>
          <p class="style1">Before deploying this sample at your server, please take a look at:<br />
              <a href="https://www.sautinsoft.com/products/useoffice/examples/installation-on-server.php" target="_blank">How to install UseOffice .Net
          at Windows Servers.</a></p>             
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td></td>
      </tr>
      <tr>
        <td colspan="2" style="height: 168px"><div style="font-size: 12pt; font-family: Verdana"><p><b>Step 1:</b> Select a document (PDF, DOC, DOCX, RTF, HTML, XLS, XLSX, Text, CSV, PPT, PPTX):</p>
            <asp:FileUpload ID="uploadedDocument" runat="server" Width="805px" style="font-size: 12pt; font-family: Verdana" />
            <p><b>Step 2:</b> Specify the direction: &nbsp;
            <asp:DropDownList ID="convDir" runat="server" Font-Names="Verdana" Font-Size="12pt"
                Width="173px" OnSelectedIndexChanged="convDir_SelectedIndexChanged" style="font-size: 12pt; font-family: Verdana">
            </asp:DropDownList></p>            
            <p><b>Step 3: </b>
            <asp:Button ID="convert" runat="server" Font-Names="Verdana" Font-Size="12pt" Text="Convert" OnClick="convert_Click" style="font-size: 12pt; font-family: Verdana; width: 150px; height: 35px;" /></p>
                        <p>
            <asp:HyperLink ID="fileMessage" runat="server">[fileMessage]</asp:HyperLink>
            <br />
            <asp:Label ID="resultMessage" runat="server" Font-Bold="True" Font-Names="Verdana" ForeColor="#FF8000" Width="1024px"></asp:Label></p>            
            
                                              </div></td>
      </tr>
    </table>
    
    <div>        
        <br />
        &nbsp;
                    &nbsp;</div>
    <div align="center"><br />
        <span style="font-size: 9pt; font-family: Verdana"><a href="https://www.sautinsoft.com" target="_blank">Copyright &copy; SautinSoft 2002 - <script>document.write(new Date().getFullYear())</script></a> </span>
    </div>
</form>
</body>
</html>
