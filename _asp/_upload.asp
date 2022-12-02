<!--#INCLUDE FILE="_incheader.asp"-->

<%
'Stores only files with size less than MaxFileSize


Dim DestinationPath
DestinationPath = Server.mapPath("upload")

'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm %><!--#INCLUDE FILE="_incupload.asp"--><% 

Server.ScriptTimeout = 2000
Form.SizeLimit = &HA00000

'{b}Set the upload ID for this form.
'Progress bar window will receive the same ID.
if len(Request.QueryString("UploadID"))>0 then
	Form.UploadID = Request.QueryString("UploadID")'{/b}
end if
'was the Form successfully received?
Const fsCompletted  = 0

If Form.State = fsCompletted Then 'Completted
  'was the Form successfully received?
  if Form.State = 0 then
    'Do something with upload - save, enumerate, ...
    'response.write "<br><b>Upload result: Form was accepted.</b>" 
    'response.write "<br>Number of file fields:" & Form.Files.Count
    'response.write "<br>Request total bytes: " & Request.TotalBytes
	Form.Files.Save DestinationPath 
	response.write "<br>File was saved." 'to " & DestinationPath & " folder."
  End If
ElseIf Form.State > 10 then
  Const fsSizeLimit = &HD
  Select case Form.State
		case fsSizeLimit: response.write  "<br><Font Color=red>Source form size (" & Form.TotalBytes & "B) exceeds form limit (" & Form.SizeLimit & "B)</Font><br>"
		case else response.write "<br><Font Color=red>Some form error.</Font><br>"
  end Select
End If'Form.State = 0 then


'{b}get an unique upload ID for this upload script and progress bar.
Dim UploadID, PostURL
UploadID = Form.NewUploadID

'Send this ID as a UploadID QueryString parameter to this script.
PostURL = Request.ServerVariables("SCRIPT_NAME") & "?UploadID=" & UploadID'{/b}

%>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>GSA_HTW Upload</TITLE>
<STYLE TYPE="text/css">
 <!--TD	{font-family:Arial,Helvetica,sans-serif }TH	{font-family:Arial,Helvetica,sans-serif }TABLE	{font-size:10pt;font-family:Arial,Helvetica,sans-serif }--></STYLE>
<meta name="robots" content="noindex,nofollow">
</HEAD>
<body>
<!--#INCLUDE FILE="_incbodyline.asp"-->

<Div style=width:600>
<TABLE cellSpacing=2 cellPadding=1 width="100%" bgColor=white border=0>
  
  <TR>
    <TD colSpan=2>
<br>
  </TD></TR></TABLE>

<form name="file_upload" method="POST" ENCTYPE="multipart/form-data" OnSubmit="return ProgressBar();" Action="<%=PostURL%>">

<Div ID=files>
   File to Upload: <input type="file" name="File1"><br>
</Div>

<P>Form size limit is <%=Form.SizeLimit \ 1024 %> k </P>

<br>

Description: <input Name=Description1 Size=60><br>
<input Name=SubmitButton Value="Submit »" Type=Submit><br>
</Form>

<SCRIPT>
//Open window with progress bar.
function ProgressBar(){
  var ProgressURL
  ProgressURL = 'progress.asp?UploadID=<%=UploadID%>'

  var v = window.open(ProgressURL,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=350,height=200')
  
  return true;
}
</SCRIPT> 

<Script>
//Expand form with a new File fields if needed.
var nfiles = 3;
function Expand(){
  nfiles++
  var adh = '<BR> File '+nfiles+' : <input type="file" name="File'+nfiles+'">';
  files.insertAdjacentHTML('BeforeEnd',adh);
  return false;
}
</Script>

<HR COLOR=silver Size=1>
</Div>
</BODY>
<!--#INCLUDE FILE="_incfooter.asp"-->
</HTML>