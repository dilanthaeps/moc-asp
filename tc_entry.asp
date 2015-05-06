	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Timer Charterer Entry
	'	Template Path	:	tc_entry.asp
	'	Functionality	:	To allow the entry/edit of the MOC TC details
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd  August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = false
	Dim Idval
	Idval = Request.QueryString("v_time_charterer_ID")
	if Idval <> "" then
		v_mode="edit"
		v_header="Update MOC Time Charterer Detils"
		strSql = "Select time_charterer_id, short_name, full_name, address,telephone,mobile,fax_no,email,pic,remarks, to_char(create_date,'dd/mm/yyyy') create_date,to_char(last_modified_date,'dd/mm/yyyy') last_modified_date, created_by, last_modified_by from moc_time_charterers where time_charterer_id="
		strSql = strSql & "'" & Idval & "'"
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Add New MOC Time Charterer"
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

			if (document.form1.short_name.value == "")
			{
				alert ("Enter Short Name");
				document.form1.short_name.focus();
				return false;
			}
			if(document.form1.full_name.value.length<1)
			{
				alert("Enter Full Name  ");
				document.form1.full_name.focus();
				return false;
			}
			if(document.form1.address.value.length>100)
			{
				alert("Address Maximum 100 Chars only");
				document.form1.address.focus();
				return false;
			}
			//Category Remarks Max char check
			if (document.form1.remarks.value.length > 200)
			{
				alert ("Remarks more than 200 chars");
				document.form1.remarks.focus();
				return false;
			}

  		  }
</SCRIPT>
<TITLE> MOC Time Charterer Entry/Edit Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<form name=form1  action=tc_save.asp method=post  OnSubmit="return validate_fields(this)">
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1>


<% if v_mode="edit" then %>
	<TR>
		<TD class=tableheader>
			Time Charterer ID
		</TD>
		<TD class=columncolor>
			 <%=(SFIELD("time_charterer_id"))%>
		</TD>
	</TR>
<% end if %>



	<TR>
		<TD class=tableheader>
			Short Name
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=20  name=short_name value="<%=(SFIELD("short_name"))%>" maxlength=20><font color=red>*</font>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Full Name
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=50  name=full_name value="<%=(SFIELD("full_name"))%>" maxlength=100><font color=red>*</font>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Address
		</TD>
		<TD class=columncolor>
			<textarea name="address" cols="50" rows="5"><%=SFIELD("address")%></textarea>
		</TD>
	</TR>

	<TR>
		<TD class=tableheader>
			Telephone
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=30  name=telephone value="<%=SFIELD("telephone")%>" maxlength=30>
		</TD>
	</TR>

	<TR>
		<TD class=tableheader>
			Mobile
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=30  name=mobile value="<%=SFIELD("mobile")%>" maxlength=30>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Fax No
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=30  name=fax_no value="<%=SFIELD("fax_no")%>" maxlength=30>
		</TD>
	</TR>


	<TR>
		<TD class=tableheader>
			Email
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=30  name=email value="<%=SFIELD("email")%>" maxlength=255>
		</TD>
	</TR>

	<TR>
		<TD class=tableheader>
			Person Incharge
		</TD>
		<TD class=columncolor>
			<INPUT type=text size=50  name=pic value="<%=SFIELD("pic")%>" maxlength=50>
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
		 <INPUT type="hidden" name=time_charterer_id value="<%=SFIELD("time_charterer_id")%>">
		</TD>
	</TR>
	<TR>
		<TD class=tableheader>
			Create By
		</TD>
		<TD class=columncolor>
		<%=(SFIELD("created_by"))%>	&nbsp;
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
	<TR>
		<TD class=tableheader>
			Last Modified By
		</TD>
		<TD class=columncolor>
		<%=(SFIELD("last_modified_by"))%>&nbsp;
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
connObj.Close
set connObj=nothing
%>

</BODY>
</HTML>
