<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
 <link REL="stylesheet" HREF="moc.css">
 <style>

TD
{
	 font-size:10px;
}
.clsVessel
{
	font-size:12px;
	font-weight:bold;
	
}
.clsExpired
{
	color:red;
	font-weight:bold;
}
.clsHighlighted
{
	color:red;
	font-weight:bold;
}
.clsBasisMOC
{
	color:darkblue;
	font-size:10px;
	font-weight:bold;
}
</style>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>
<script language="vbscript">
sub MOC_change()   
	if frm1.cmbselect.value="MOC" then			    
		window.location.href="rpt_observations_MOC.asp"		   
	elseif frm1.cmbselect.value="FLEET" then		 
		window.location.href="rpt_observations_fleet.asp"
end if  		
end sub
Sub window_onbeforeprint
	divTopMenu.style.display = "none"
	Hide.style.display = "none"
	hide1.style.display=""	
End Sub
Sub window_onafterprint
	divTopMenu.style.display = ""
	Hide.style.display = ""	
	hide1.style.display="none"
End Sub
sub RefreshPage()
	dim sUrl
	sd1=frm1.v_insp_from_date.value
	sd2=frm1.v_insp_to_date.value
	if IsDate(sd1) then
		s1 = day(sd1) & " " & monthname(month(sd1),true) & " " & year(sd1)
	else
		MsgBox "Please enter a valid from date",vbInformation,"MOC"
		frm1.v_insp_from_date.focus
		exit sub
	end if
	if IsDate(sd2) then
		s2 = day(sd2) & " " & monthname(month(sd2),true) & " " & year(sd2)
	else
		MsgBox "Please enter a valid to date",vbInformation,"MOC"
		frm1.v_insp_to_date.focus
		exit sub
	end if
	sUrl = "rpt_observations_moc.asp?SDATE1=" & s1 & "&SDATE2=" & s2
	window.location.href = sUrl
end sub
</script>
   </head>
<% 
  dim SDATE1,SDATE2,dt
  SDATE1 = Request.QueryString("SDATE1")
  SDATE2 = Request.QueryString("SDATE2")
  if SDATE1="" then
	dt = cdate("1 " & MonthName(month(now),true) & " " & year(now))
	SDATE2 = FormatDateTimeValues(dt-1,2)
	SDATE1 = FormatDateTimeValues(DateAdd("m",-1,dt),2)
  end if
%>   
<BODY class="bgcolorlogin">
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<div id="hide1" style="display:none">
     <center><h3 style="margin-bottom:0" id="no_hide">Observations by MOC - From 1 Jan 2006<br><h4 style="margin:0;">(Whole Fleet)</h4></center>
</div>     
<form name="frm1">    
       <blockquote><blockquote>
      <table  id="hide"><tr><td style="font-size=9pt"><b>Group by</b>
					 <select id="cmbselect" name="cmbselect" width="100px" onchange="MOC_change()" class=menuHide>
					   <option value="MOC" selected>MOC</option>
					   <option value="FLEET">FLEET</option>
					</select>
			    <td style="width:80px">&nbsp;
			    <td colspan=3><h3 style="margin-bottom:0" id="no_hide">Observations by MOC - From 1 Jan 2006<br><center><h4 style="margin:0;">(Whole Fleet)</h4></center></td>
			 <tr><td>&nbsp;</td> <td style="width:80px">&nbsp;
			     <td nowrap align=center><b>Date from</b><br>
			         <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<%=SDATE1%>" SIZE="12"
						onblur="vbscript:checkDate frm1.v_insp_from_date,'Inspection Date From','frm1'">
						<A HREF="javascript:show_calendar('frm1.v_insp_from_date',frm1.v_insp_from_date.value);">
						<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		         <td>&nbsp;&nbsp;&nbsp;<b>Date to</b><br>
			         <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_to_date" VALUE="<%=SDATE2%>" SIZE="12"
						onblur="vbscript:checkDate frm1.v_insp_to_date,'Inspection Date To','frm1'">
						<A HREF="javascript:show_calendar('frm1.v_insp_to_date',frm1.v_insp_to_date.value);">
						<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
				<td><input type=button id="button1" name="button1" value="refresh" onclick=RefreshPage()>		  
     </table></blockquote></blockquote>   
               
    <table align="center" cellpadding=2 cellspacing=0 >      
      <tr>                
        <table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey width="90" align=center>
               <tr class=tableheader ><td align=center><b>Section</b></td><td align=center><b>Question</b></td>
       <%dim rsShortName,rsdetails,rsname,rschapter
         dim i,name1,str1,sh_name,str2,count2,chapter1,arritem
         dim sh_name1,count1,sect ,moc_id        
         if sDate1="" then
         	sDate1="1 jan 2006"
         	sDate2="31 dec 2006"
         end if
         str1 ="" 
         i=0
          set rsShortName=Server.CreateObject("ADODB.Recordset")
          strSql= "select short_name from moc_master where entry_type='MOC' order by short_name"
          rsShortName.CursorLocation=adUseClient
          rsShortName.Open strSql,connObj,adOpenDynamic,adLockReadOnly
        while not rsShortName.eof %>          
           <td style="writing-mode:tb-rl;filter:flipv fliph"><b><%=rsShortName("short_name")%></b></td>
          <%rsShortName.MoveNext
        wend%></tr>
		<%set rsdetails=Server.CreateObject("ADODB.Recordset")
         strSql= "select distinct section,question_text,chapter,mvq.sort_order" 
         strSql=strSql & " from moc_deficiencies md,moc_viq_questions mvq ,moc_inspection_requests mir,moc_master mm"
         strSql=strSql & " where md.section=mvq.question_number"
         strSql=strSql & " and mir.moc_id=mm.moc_id"
         strSql=strSql & " and entry_type='MOC'"
         strSql=strSql & " and md.request_id=mir.request_id"
         if sDate1<>"" then
			strSql=strSql & " and trunc(mir.inspection_date) between '" & sDate1 & "' and '" & sDate2 & "'"
         end if
         strSql=strSql & " order by mvq.sort_order"
         Set rsdetails=connObj.execute(strSql)
        
        set rschapter=Server.CreateObject("ADODB.Recordset")
        strSql= "select distinct section,chapter,mvq.sort_order"
        strSql=strSql & " from moc_deficiencies md,moc_viq_questions mvq "
        strSql=strSql & " where md.section=mvq.question_number"        
        strSql=strSql & " order by mvq.sort_order"
        rschapter.CursorLocation=aduseclient
        rschapter.open strSql,connObj
                
        set rsname=Server.CreateObject("ADODB.Recordset")
        strSql= " select section,question_text,short_name,count(*) count1,mvq.sort_order,mm.moc_id"
        strSql=strSql & " from moc_deficiencies md,moc_viq_questions mvq ,moc_inspection_requests mir,moc_master mm"
        strSql=strSql & " where md.section=mvq.question_number "
        strSql=strSql & " and mir.moc_id=mm.moc_id"
        strSql=strSql & " and md.request_id=mir.request_id "
        if sDate1<>"" then
			strSql=strSql & " and trunc(mir.inspection_date) between '" & sDate1 & "' and '" & sDate2 & "'"
		end if
        strSql=strSql & " group by short_name,section,question_text,mvq.sort_order,mm.moc_id"
        strSql=strSql & " order by mvq.sort_order,short_name"
        rsname.CursorLocation=aduseclient
        rsname.open strSql,connObj
        
        while not rsdetails.EOF%>
           <%sect=rsdetails("section")
             chapter1=rsdetails("chapter")
             rschapter.Filter="chapter='" & chapter1 & "'"
           if not rschapter.EOF then
             arritem=rschapter.GetRows
           end if             
           if sect=arritem(0,0) then
				  i=split(sect,".")%>				  
				<tr bgcolor=white><td><b><%=i(0)%></b></td>
				        <td colspan="37"><b><%=rsdetails("chapter")%></b></td></tr>
         <%end if%> 
         <tr bgcolor=white>
           <td align="left"><%=sect%></td>
           <td><%=rsdetails("question_text")%></td>           
              <%rsname.Filter="section='" & sect & "'"
			   rsShortName.movefirst
			   while not rsShortName.EOF
					str2=""
					count2="&nbsp;"
					if not rsname.EOF then
						if rsShortName("short_name")=rsname("short_name") then
							str2=rsname("short_name")
							moc_id=rsname("moc_id")
							count2=rsname("count1")
							rsname.MoveNext
						end if
					end if%>
				<td align="center"><a href="rpt_observations.asp?VIQ=<%=sect%>&MOC=<%=moc_id%>&sDate1=<%=sDate1%>&sDate2=<%=sDate2%>"><%=count2%></a></td>
					<%rsShortName.MoveNext
			   wend%> 
          </tr>
          <%rsdetails.movenext
       wend%>  
     </table>
</tr></table>
</form>
</body>
</html>