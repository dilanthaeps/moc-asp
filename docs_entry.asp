	<!--#include file="common_dbconn.asp"-->
<%	
	if request("v_fbm_id")<>"" then
		v_fbm_id=request("v_fbm_id")
	end if
	Dim Idval
	Idval = Request.QueryString("v_document_id")
	if Idval <> "" then
		v_mode="edit"
		v_header="Update Document Details"
		strSql = "select * from wls_fbm_attachements where document_id="&Idval&"" 
		Set rsObj = connObj.Execute(strSql)
	else 
		v_mode="Add"
		v_header="Add New Document"
	end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="moc.css"></link>
<script language="VBScript" runat=server>
	   function SFIELD(fname)
	      if v_mode="edit" then
				rsObj.MoveFirst
	      		Do Until rsObj.EOF
					v_tem = rsObj(cstr(fname))
					rsObj.MoveNext
				Loop
	         SFIELD=v_tem
	         
	      else
	         SFIELD = ""
	      end if
	   End function
</script>
<SCRIPT LANGUAGE="JAVASCRIPT">
	  function validate_fields (thisForm) 
		  {
			if (thisForm.description.value == "")
			{
				alert ("Description value cannot be blank");
				thisForm.description.focus();
				return false;
			}
			/*if (thisForm.file1.value == "")
			{
				alert ("Choose File Name");
				thisForm.file1.focus();
				return false;
			}*/

  		  }
</SCRIPT>
<TITLE>Documents Entry Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<form name=form1  action="uploader.asp?v_fbm_id=<%=request("v_fbm_id")%>"  method=post enctype='multipart/form-data' OnSubmit="return validate_fields(this)">
<TABLE cellSpacing=1 cellPadding=1 width="50%" border=1>
<%if v_mode="edit" then%>
	<input type=hidden name="v_document_id" value="<%=SFIELD("document_id")%>">
<%end if%>  
<input type=hidden name="v_fbm_id" value="<%=v_fbm_id%>">

<TR>
    <TD class=tableheader>Description</TD>
    <TD class=columncolor><input type=text name=description value="<%=SFIELD("description")%>" maxlength=50 size=60></td>
</tr>

<TR>
    <TD class=tableheader>File</TD>
    <TD class=columncolor>
    <INPUT type="file"  name="file1"  value="<%=SFIELD("doc_path")%>" size=60>
</tr>
<tr>
    <%if SFIELD("doc_path")<>"" then%>
       <td class=columncolor><a href="<%=SFIELD("doc_path")%>" target="new">View</a>&nbsp;</td>
       <td class=columncolor>Existing path:<%=SFIELD("doc_path")%></td>
       <input type=hidden name="exist_path" value="<%=SFIELD("doc_path")%>">
    <%end if%>

</TR>
<tr>
  <td class=columncolor colspan=2 align=center>
  <input type=submit value=save name=submit class=cmdbutton>
  <input type=reset value=reset name=reset class=cmdbutton>
</tr>
</TABLE>
  
</form>
		 <% if v_mode="edit" then%>
			<!--#include file="common_footer.asp"-->
         <% end if%>			
</BODY>
</HTML>
