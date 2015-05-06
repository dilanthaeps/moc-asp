<%@ Language=VBScript %>
<%option explicit%>
<%on error resume next%>
<!--#include file="common_dbconn.asp"-->
<%
dim vcode,sNote,arr,i
arr=split(Request.Body,"~")
vcode = arr(0)
sNote = Replace(arr(1),"'","''")
connObj.execute "Update MOC_VESSEL_NOTES set note='" & sNote & "' where vessel_code='" & vcode & "'",i
if i=0 then
	connObj.execute "Insert into MOC_VESSEL_NOTES values('" & vcode & "','" & sNote & "')"
end if
if err.number=0 then
	Response.Write "Note updated"
else
	Response.Write "Error updating note (" & err.Description & ")"
end if
%>