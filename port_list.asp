   <!--#include file="common_dbconn.asp"-->
 <%	'===========================================================================
	'	Template Name	:	MOC PORT List
	'	Template Path	:	port_list.asp
	'	Functionality	:	To choose the port from the port list
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	5th September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<title>PORT List - Tanker Pacific</title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="vbscript">
dim oColl,objSel
dim timerID,sSearch
sub window_onload
	set oColl = document.getElementsByName("portname")
end sub
sub document_onkeypress
	if window.event.keyCode<32 or window.event.keyCode>122 then exit sub
	if timer<>0 then
		if not IsEmpty(objSel) then
			objSel.style.backgroundColor = ""
			window.status = ""
		end if
		clearTimeout timerID
		timerID = setTimeout("ClearTimer",1000)
	end if
	
	sSearch = ucase(sSearch & chr(window.event.keyCode))
	window.status = sSearch
	for i=0 to oColl.length-1
		if sSearch < ucase(oColl(i).innerText) then exit for
	next
	set objSel = oColl(i)
	objSel.scrollIntoView
	objSel.style.backgroundColor = "yellow"
end sub
sub ClearTimer
	sSearch = ""
	objSel.style.backgroundColor = ""
	window.status = ""
	timerID = 0
end sub
</script>
<script language="Javascript">
function fncall(port)
{
	opener.document.form1.inspection_port.value=port;
	self.close();
}
</script>

</head>
<body>
<center>
<h4> PORT List</h4>
<p></p>

<% 

   'strSql="SELECT PORT,COUNTRY,AIRPORT_CODE,PORT_CODE FROM PORT ORDER BY PORT "
   strSql="SELECT PORT_name port,port_country COUNTRY FROM PORTS_LIBRARY ORDER BY trim(upper(PORT_name)) "
   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action= method=post>
	<table>
		<tr>
			<td class=tableheader>Port</td>
			<td class=tableheader>Country</td>
		</tr>	
		<% if not(rsObj.eof or rsObj.bof) then
		dim c,cclass,r_count
		c=0
		r_count=0
		while not rsObj.EOF 
		if c=0 then
		cclass="tabledata"
		c=1
		elseif c=1 then
		cclass="tabledataalt"
		c=0
		end if
		%>
		   
		<tr>
			<td class="<%=cclass%>" id="portname">
				<a href="Javascript:fncall('<%=rsObj("port")%>')"><%=rsObj("port")%>&nbsp;</a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("country")%>&nbsp;</td>
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
  
  
  <input type=BUTTON name=ebut value=Exit OnClick="Javascript:parent.close();">
  <!--#include file="common_footer.asp"-->  
</form>
</center>
</table>    
</body>
</html>