	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	System Parameter Maintenance
	'	Template Path	:	system_parameter_maint.asp
	'	Functionality	:	To show the list of system parameters available
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	21st August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = false
	Dim Idval
	Idval = Request.QueryString("v_sys_para")
	if Idval ="" then
		strSql1 = "Select sys_para_id,para_desc from moc_system_parameters where parent_id ='No Parent'"
	else
		strSql1 = "Select sys_para_id,para_desc from moc_system_parameters where parent_id ='No Parent' and sys_para_id <> '" & Idval&"'"
	end if
	Set rsObj1 = connObj.Execute(strSql1)

	if Idval <> "" then
		v_mode="edit"
		v_header="Update System Parameter Detils"
		strSql = "Select sys_para_id,para_desc,parent_id,remarks, to_char(create_date,'dd/mm/yyyy') create_date,to_char(last_modified_date,'dd/mm/yyyy') last_modified_date,related_asp_pages,sort_order from moc_system_parameters where sys_para_id="
		strSql = strSql & "'" & Idval & "'"
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Add New System Parameter"
	end if
%>


<html>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="VBScript" runat=server>
	   function SFIELD(fname)
	      if v_mode="edit" then
			'if not (rsObj.bof or rsObj.eof) then
				rsObj.MoveFirst
	      		Do Until rsObj.EOF
					v_tem = rsObj(cstr(fname))
					rsObj.MoveNext
				Loop
			'end if
	         SFIELD=v_tem

	      else
	         SFIELD = ""
	      end if
	   End function
</script>
<SCRIPT LANGUAGE="JAVASCRIPT">
	  function validate_fields()
		  {
			if (document.form1.sys_para_id.value == "")
			{
				alert ("Enter System Parameter ID");
				document.form1.sys_para_id.focus();
				return false;
			}
			if (document.form1.para_desc.value == "")
			{
				alert ("Enter System Parameter Name");
				document.form1.para_desc.focus();
				return false;
			}
			if(document.form1.related_asp_pages.value.length>1000)
			{
				alert("Related Pages more than 1000 charectors");
				document.form1.related_asp_pages.focus();
				return false;
			}
			if(document.form1.related_asp_pages.value.length<1)
			{
				alert("Related Pages Value cannotbe blank");
				document.form1.related_asp_pages.focus();
				return false;
			}
			//Category Remarks blank check
			if (document.form1.remarks.value.length > 4000)
			{
				alert ("Remarks more than 4000 chars");
				document.form1.remarks.focus();
				return false;
			}
			//Sort Order check
			if (  document.form1.sort_order.value =="")
			{
				alert ("Please enter sort order number  ");
				document.form1.sort_order.focus();
				return false;
			}
			//Sort Order check
			if ( (isNaN( document.form1.sort_order.value)))
			{
				alert ("Please enter number only");
				document.form1.sort_order.focus();
				return false;
			}
  		  }
</SCRIPT>
<TITLE> System Parameter Entry Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<form name=form1  action=system_parameter_save.asp method=post  OnSubmit="return validate_fields(this)">
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1>

	<TR>
		<TD  colspan=2>
			<b>Note:</b> <%=Mid(SFIELD("remarks"), 1, 100) & " ... <FONT STYLE='font-size:8pt'><B>(Please refer remarks for more info.)</B></FONT>" %>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Parameter ID
		</TD>
		<% if v_mode="edit" then %>
		<TD class=columncolor>
			<INPUT type=hidden size=20  name=sys_para_id value="<%=(SFIELD("sys_para_id"))%>" maxlength=20><%=(SFIELD("sys_para_id"))%><font color=red>*</font>
		</TD>
		<% else %>
		<TD class=columncolor>
			<INPUT type=text size=20  name=sys_para_id value="<%=(SFIELD("sys_para_id"))%>" maxlength=20><font color=red>*</font>
		</TD>
		<% end if %>
	</TR>


	<TR>
		<TD class=tableheader>
			Parent ID
		</TD>
		<TD class=columncolor>
			<SELECT name=parent_id >
			<option value="No Parent">None</option>
			<%
			if not ( rsObj1.bof or rsObj1.eof ) then
				while not rsObj1.eof
					%>

					<option value="<%=rsObj1("sys_para_id")%>"  <% if sfield("parent_id") = rsObj1("sys_para_id") then Response.Write "selected" %> ><%=rsObj1("para_desc")%></option>
					<%
					rsObj1.movenext
				wend
			end if
			%>
			</SELECT>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Parameter Description
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=50  name=para_desc value="<%=(SFIELD("para_desc"))%>"><font color=red>*</font>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Remarks
		</TD>
		<TD class=columncolor>
			<textarea name="remarks" cols="50" rows="5"><%=SFIELD("remarks")%></textarea>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Related Pages
		</TD>
		<TD class=columncolor>
			<textarea name="related_asp_pages" cols="50" rows="5"><%=SFIELD("related_asp_pages")%></textarea><font color=red>*</font>
			<b>Note:</b> Please take adequate care before modifying or deleting information in this field. <br>The related Programming pages are available in the Related Asp Pages Field
		</TD>
	</TR>

	<TR>
		<TD class=tableheader>
			Sort Order
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=4  name=sort_order value="<%=SFIELD("sort_order")%>" maxlength=4><font color=red>*</font>
		</TD>
	</TR>
<%
	if Idval <> "" then
%>
	<TR>
		<TD class=tableheader>
			Create Date
		</TD>
		<TD class=columncolor>
		<%=(SFIELD("create_date"))%>
		 <INPUT type="hidden" name=mode value="edit">
		</TD>
	</TR>
	<TR>
		<TD class=tableheader>
			Modified Date
		</TD>
		<TD class=columncolor>
			<%=(SFIELD("last_modified_date"))%>
		</TD>
	</TR>
<%
	end if
%>
</TABLE>
<font color=red>*</font> Denotes Mandatory Field.<br>
<table align=left width=50%>
<tr><td align=center>
<input type=submit value=save name=submit>
<input type=reset value=reset name=reset>
</td></tr>
</table>
</form>

<%'close record set and connection object
if v_mode="edit" then
  rsObj.Close
  Set rsObj=nothing
end if
rsObj1.close
Set rsObj1=nothing
connObj.Close
set connObj=nothing
%>

</BODY>
</HTML>
