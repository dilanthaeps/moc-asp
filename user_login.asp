<%	'===========================================================================
	'	Template Name	:	MOC User Login
	'	Template Path	:	user_login.asp
	'	Functionality	:	To show the list of User Login
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css"></link>
<title>Tanker Pacific - Inspection - Login Screen</title>
<script language="Javascript">
	function fncall()
	{
		var w = document.body.clientWidth; 
		winStats='toolbar=no,location=no,directories=no,menubar=no,'
		winStats+='scrollbars=yes'
		if(w<=800){
			winStats+=',left=170,top=90,width=450,height=220'
		} else {
		winStats+=',left=70,top=50,width=850,height=500'
		}
		adWindow=window.open("change_password_entry.asp?","change_password",winStats);     
		adWindow.focus();
	}
</script>
</head>
<body class="bgcolorlogin" onLoad="javascript:document.form1.user_id.focus();">
 
<br>
<table width="100%">
	<tr>
		<td align="left">
			<img SRC="Images/Logo1.gif" ALIGN="Left" ALT="Tanker Pacific Management" WIDTH="155" HEIGHT="81">
		</td>
	</tr>
</table>
<p> &nbsp; </p>
<table width="100%">
	<tr>
		<td class="message">
			<blink><center><% Response.write request("v_message")%></center></blink>
		</td>
	</tr>
</table>
<p> &nbsp; </p>
<form action="login_process.asp?prev_path=<%= Request.QueryString("prev_path") %>" method="post" id="form1" name="form1">
<table align="center" border="1">
<td align="center" bgcolor="Tan">
	<font size="2" style="FONT-FAMILY: 'Arial','verdana';"><strong>Inspection Information System</strong></font>
</td>
</tr>
<tr>
<td class="columncolor">
	<table border="1" align="center">
	<!--<tr>		<td align=center class=columncolor>			<FONT  size=1 style="FONT-FAMILY: 'verdana', Arial;"><b>Existing Users</b></font>			</td>	</tr>-->
	<tr>
		<td align="center" class="columncolor">
			<font size="1" style="FONT-FAMILY: 'verdana', Arial;"><b>Enter Your ID and Password to Sign In	</b></font>	
		</td>
	</tr>
	<tr>
	<td>
		<table border="0" align="center" cellpadding="2" cellspacing="0">
		<tr> <td align="right" nowrap><font size="2" style="FONT-FAMILY: 'verdana', Arial;">User ID:</font></td>
		<td><input name="user_id" size="17" maxlength="32" value></td>
		</tr>
		<tr> <td align="right" nowrap><font size="2" style="FONT-FAMILY: 'verdana', Arial;">Password:</font></td>
		<td><input name="password" type="password" size="17" maxlength="32"></td></tr>
		<tr>
		<td>&nbsp;</td>
		<td><input name="save" type="submit" value="Sign In"></td>
		</tr>
		</table> 
	</td>
	</tr>
	<%
		'strSql="select para_desc from moc_system_parameters where sys_para_id='ADMINISTRATOR_EMAIL_ID'"
		'set rsObj=connObj.execute(strSql)
		'if not (rsObj.eof or rsObj.bof) then
			'rsObj.movefirst
			'admin_mail_id=rsObj("para_value")
		'else
			admin_mail_id="iamakov"
		'end if
	%>
	<tr><td>New Users Please Contact <a href="mailto:<%=admin_mail_id%>?subject=MOC Information System System Access Privilege Request"><b>Administrator</b></a>&nbsp;</td></tr>
	</table>
</td>
</tr>
</table>
</form>
</body>
</html>
