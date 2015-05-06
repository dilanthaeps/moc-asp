<!--#include file="common_dbconn.asp"-->
<% 
	set fs=CreateObject("Scripting.FileSystemObject")
    'filepath=server.MapPath ("images")

	if request("v_fbm_id")<>"" then
		v_fbm_id=request("v_fbm_id")
	end if
	dim i
	v_message = ""
	For each i in Request.Form("v_deleteditems")
		strSqlfpath="Select doc_path from wls_fbm_attachements where document_id="&i&""
		set rsObjfpath=connObj.execute(strSqlfpath)
		if not(rsObjfpath.eof or rsObjfpath.bof) then
			rsObjfpath.movefirst
			v_doc_path=rsObjfpath("doc_path")
			v_doc_split=split(v_doc_path,"/")
			n=0
			for each k in v_doc_split
				if n=5 then
				v_fname=k
				end if
				n=n+1
			next
			v_del_file=server.MapPath("docs") & "\" & v_fname
			v_new_file=v_del_file & "bak"
			'response.Write v_new_file
			'Response.End 
			'fs.CopyFile v_del_file,v_new_file, True
			fs.DeleteFile v_del_file ,true 
		end if
 		strSql="DELETE FROM wls_fbm_attachements where document_id ="&i&"" 
		v_message = v_message&"Document Id :<i>"&i&"</i> deleted Successfully !!<br>"
		ConnObj.Execute(strSql)
	next      
	set fs=nothing
	ConnObj.Close
	Set ConnObj=nothing
	Response.Redirect "docs_maint.asp?v_fbm_id="&v_fbm_id&"&v_message="&v_message 
%>

