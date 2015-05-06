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
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<title>Manage System Parameters</title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<SCRIPT LANGUAGE="JavaScript">
function addEditRecord(sysParaID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 225)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("system_parameter_entry.asp?v_sys_para=" + sysParaID, "sysParaAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<STYLE TYPE="text/css">
<!--
#dek {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}
//-->
</STYLE>
<body class=bgcolorlogin onLoad="scroll()" STYLE="MARGIN-TOP:0px">
<!--#include file="menu_include.asp"-->
<DIV ID="dek"></DIV>
<SCRIPT TYPE="text/javascript">
<!--
Xoffset=400;    // modify these values to ...
Yoffset=500;    // change the popup position.

var old,skn,iex=(document.all),yyy=1000;

var ns4=document.layers
var ns6=document.getElementById&&!document.all
var ie4=document.all

if (ns4)
skn=document.dek
else if (ns6)
skn=document.getElementById("dek").style
else if (ie4)
skn=document.all.dek.style
if(ns4)document.captureEvents(Event.MOUSEMOVE);
else{
skn.visibility="visible"
skn.display="none"
}
document.onmousemove=get_mouse;

function popup(msg,bak){
var content="<TABLE  WIDTH=400 BORDER=1 BORDERCOLOR=black CELLPADDING=2 CELLSPACING=0 "+
"BGCOLOR="+bak+"><TD class='columncolor1'>"+msg+"</TD></TABLE>";
yyy=Yoffset;
 if(ns4){skn.document.write(content);skn.document.close();skn.visibility="visible"}
 if(ns6){document.getElementById("dek").innerHTML=content;skn.display=''}
 if(ie4){document.all("dek").innerHTML=content;skn.display=''}
}

function get_mouse(e){
var x=(ns4||ns6)?e.pageX:event.x+document.body.scrollLeft;
skn.left=x+Xoffset;
var y=(ns4||ns6)?e.pageY:event.y+document.body.scrollTop;
skn.top=y+yyy;
}

function kill(){
yyy=-1000;
if(ns4){skn.visibility="hidden";}
else if (ns6||ie4)
skn.display="none"
}

//-->
</SCRIPT>
<center>
<h4> System Parameters Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New System Parameter
<%
   'strSql="select sys_para_id,para_desc,remarks,parent_id,to_char(create_date,'dd/Mon/yyyy') create_date,to_char(last_modified_date,'dd/Mon/yyyy') last_modified_date,related_asp_pages  from MOC_system_parameters order by parent_id,sys_para_id"
   strSql="select level,sys_para_id,para_desc, nvl(trim(remarks), ' ') remarks, parent_id,to_char(create_date,'dd/Mon/yyyy') create_date,to_char(last_modified_date,'dd/Mon/yyyy') last_modified_date,related_asp_pages ,sort_order from MOC_system_parameters connect by prior sys_para_id=parent_id start with parent_id='No Parent'"
   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action=system_parameter_delete.asp method=post>
	<table >
		<tr>
			<td class=tableheader>Parameter ID</td>
			<td class=tableheader>Parameter Description</td>
			<td class=tableheader>Remarks</td>
			<td class=tableheader>Parent ID</td>
			<td class=tableheader>Sort Order</td>
			<td class=tableheader>Create Date</td>
			<td class=tableheader>Last Modified Date</td>
			<td class=tableheader>Related Pages</td>
		</tr>
		<% if not(rsObj.eof or rsObj.bof) then
		dim c,cclass,r_count
		c=0
		r_count=0
		while not rsObj.EOF
		if c=0 then
		cclass="columncolor2"
		c=1
		elseif c=1 then
		cclass="columncolor3"
		c=0
		end if
		%>

<%
	If rsObj("remarks") = NULL Or rsObj("remarks") = "" Or IsNull(rsObj("remarks")) Then
		v_remarks = ""
	Else
		v_remarks = Mid(rsObj("remarks"), 1, 30) & " ..."
	End If
%>

		<tr>
			<td class="<%=cclass%>"><% if rsObj("parent_id")="No Parent" then %><% else %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% end if %><INPUT type="checkbox"  name=v_deleteditems  value="<%=rsObj("sys_para_id")%>"><a href="JavaScript:addEditRecord('<%=rsObj("sys_para_id")%>')"><%=replace(rsObj("sys_para_id")," ","&nbsp;")%></a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("para_desc")%>&nbsp;</td>
			<td class="<%=cclass%>" onmouseoverr="javascript:popup('<%=rsObj("remarks") %>','#F1F3C5');" onmouseoutt="Javascript:kill()"><%=v_remarks%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("parent_id")%>&nbsp;</td>
			<td class="<%=cclass%>"><%if rsObj("level")="1" then Response.Write "&nbsp;" else Response.write rsObj("sort_order")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("create_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("last_modified_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("related_asp_pages")%>&nbsp;</td>
	    </tr>
       <%rsObj.MoveNext
		 r_count=r_count+1
         wend
         Response.Write "<tr><td colspan=3 align=left><b>Record Count :</b>"&r_count&"</td></tr>"
         else
         Response.Write "<tr><td colspan=3 align=center><b> No Data Found </b></td></tr>"
         end if
       %>

  </table>
  <p></p>

 <center> <INPUT type="submit" value="Delete" id=submit1 name=submit1> </center>
<!--#include file="common_footer.asp"-->
</form>
</center>
</table>
</body>
</html>