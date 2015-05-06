<%@ Language=VBScript %>
<%

    if request("v_fbm_id")<>"" then
		v_fbm_id=request("v_fbm_id")
	else
		v_fbm_id=null
	end if    
%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="upload.asp" -->
<%

server.ScriptTimeout=3600
'NOTE - YOU MUST HAVE VBSCRIPT v5.0 INSTALLED ON YOUR WEB SERVER
'	   FOR THIS LIBRARY TO FUNCTION CORRECTLY. YOU CAN OBTAIN IT
'	   FREE FROM MICROSOFT WHEN YOU INSTALL INTERNET EXPLORER 5.0
'	   OR LATER.
' Create the FileUploader
Dim Uploader, File
Set Uploader = New FileUploader

' This starts the upload process
Uploader.Upload()
'******************************************
' Use [FileUploader object].Form to access 
' additional form variables submitted with
' the file upload(s). (used below)
'******************************************
' Check if any files were uploaded
If Uploader.Files.Count = 0 Then
	Response.Write "File(s) not uploaded."
Else
	' Loop through the uploaded files
	'File=Uploader.Files.items
	dim count
	count=cint(0)
	For Each File In Uploader.Files.Items
		' Check where the user wants to save the file
		'If Uploader.Form("saveto") = "disk" Then
			' Save the file
			dim vv
			'vv=server.MapPath("/vpd/fileupload") --old
			'vv=server.MapPath("/vid/fileupload")
			'File.SaveToDisk "d:\vpd\images"--old
			File.SaveToDisk "d:\wls\new\moc\docs"
			'binary_path="/vpd/images/" & cstr(File.FileName)--old
			binary_path="/wls/new/moc/docs/" & cstr(File.FileName)  
		'end if		
		' Output the file details to the browser
		'Response.Write "File Uploaded: " & File.FileName & "<br>"
		'Response.Write "Size: " & File.FileSize & " bytes<br>"
		'Response.Write "Type: " & File.ContentType & "<br><br>"
		count=count+1
		
	Next
End If

	'Response.Write uploader.form("") & "<BR>"

	if uploader.form("description")<>"" then
		description=replace(uploader.form("description"),"'","''")
	else
		description=null
	end if    
	if uploader.form("document_type")<>"" then
		document_type=uploader.form("document_type")
	else
		document_type=null
	end if    
	
	if uploader.form("exist_path")<>"" then
		exist_path=uploader.form("exist_path")
	else
		exist_path=null
	end if    
	if session("moc_user_id")<>"" then
		uploaded_by=session("moc_user_id")
	else
		uploaded_by="Unknown"
	end if
	
	Idval=uploader.form("v_document_id")			
	if Idval = "" then
		strSql = "INSERT INTO wls_fbm_attachements(document_id, fbm_id, description,doc_path,uploaded_by,uploaded_date) Values(seq_wls_fbm_attachements.nextval, '"&v_fbm_id&"','"&description&"','"&binary_path&"','"&uploaded_by&"',sysdate"&")"
		v_message = "Document Created Successfully"
	else
		if binary_path="" then
			binary_path=exist_path
		elseif binary_path<>exist_path then
			set fs=CreateObject("Scripting.FileSystemObject")
			if exist_path<>"" then
				v_doc_split=split(exist_path,"/")
				n=0
				for each k in v_doc_split
					if n=5 then
					v_fname=k
					end if
					n=n+1
				next
				v_del_file=server.MapPath("docs") & "\" & v_fname
				'Response.Write v_del_file
				'Response.end
				v_new_file=v_del_file & "bak"
				'response.Write v_new_file
				'Response.End 
				fs.CopyFile v_del_file,v_new_file, True
				fs.DeleteFile v_del_file ,true 
				set fs=nothing
			end if 'exists path empty check
		end if 'if check ends here
		strSql = "Update wls_fbm_attachements set description='"&description&"',doc_path='"&binary_path&"',uploaded_by='"&uploaded_by&"',uploaded_date=sysdate  where document_id=" & Idval & ""
		v_message = "Document Updated Successfully"
	end if	
	'Response.Write strSql
	'Response.end   
	connObj.Execute(strSql)
	connObj.Close
	set connObj=nothing

	'Response.Redirect "docs_maint.asp?v_message="&v_message
%>
<form name=form1 action="" method=post>
<input type=hidden name=v_fbm_id value="<%=v_fbm_id%>">
</form>

    <script language="javascript">
		//v_fbm_id=document.form1.v_fbm_id.value;
		//var v_fbm_id=<%=v_fbm_id%>
		v_page="docs_maint.asp?v_fbm_id=<%=v_fbm_id%>"
		//alert(v_page);
		//v_page="events_maint.asp?v_fbm_id="+v_fbm_id;
		//document.form1.target="projdetail";
		//document.form1.action=v_page;
		//document.form1.submit();
		self.opener.parent.document.form1.action=v_page;
		self.opener.parent.document.form1.submit();
		self.close();		
    </script>		
