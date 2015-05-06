<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC TC Assignment Entry/Edit/Listing
	'	Template Path	:	.../moc_tc_asgn.asp
	'	Functionality	:	To view/edit the MOC TC Assignment
	'	stored_proc		:	N/A
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
		Response.Buffer = false
%>

<html>
<head>
<title>MOC - MOC to Time Charterer Assignment </title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
</head>


<% 
v_mode = "edit"  
v_header="Update MOC Time Charterer Assignment"
SQL = " SELECT MOC.MOC_ID , MOC.SHORT_NAME , TC.TIME_CHARTERER_ID , TC.SHORT_NAME TC_NAME"
SQL = SQL &  " FROM   MOC_MASTER MOC , MOC_TIME_CHARTERERS TC"
SQL = SQL &  " where moc.entry_type='MOC'"
SQL = SQL &  " ORDER BY MOC.Short_name,TIME_CHARTERER_ID"
Set rsObj_moc_tc = connObj.Execute(SQL)

strSql = "SELECT   MTA.TC_MOC_ASGN_ID , MTA.MOC_ID , MTA.TIME_CHARTERER_ID FROM MOC_TC_MOC_ASGN MTA"
Set rsObj_asgn = connObj.Execute(strSql)

'strSql="Select count(*) cnt from wls_vw_vessels_new"
'set rsObj_ves_cnt=connObj.execute(strSql)

'while not rsObj_ves_cnt.eof
'v_no_vessels = rsObj_ves_cnt("cnt") '  Number of Vessel per MOC
'rsObj_ves_cnt.movenext
'wend

v_no_cols = 7 ' This variable decides the number of coloumns to be used in the table. This variable can be modified depending upon the requirement.	   

record_found = "No"
While not rsObj_asgn.eof and record_found <> "Yes"
 record_found = "Yes"
 rsObj_asgn.MoveNext
Wend
v_tem=0
v_check=" checked=false"
%>
	   
<script language="VBScript" runat=server>
	function getObjValue(MOC_ID,TIME_CHARTERER_ID)
	if record_found = "Yes" then
	rsObj_asgn.MoveFirst
	end if
		v_tem=0
		v_check=" "
		while not rsObj_asgn.eof and v_tem <> 1
		if (cstr(rsObj_asgn("MOC_ID")) = cstr(MOC_ID) and cstr(rsObj_asgn("TIME_CHARTERER_ID"))=cstr(TIME_CHARTERER_ID)) then
		v_tem=1
		v_check=" checked=true"
		'getObjValue=v_tem
		'exit do
		else
		v_tem=0
		v_check=" "
		end if
		rsObj_asgn.MoveNext
		wend
	getObjValue=v_tem
	end function
</script>
	   
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="VBScript" runat=server>
	   function SFIELD(fname)
	      if v_mode="edit" then
				rsObj_moc_tc.MoveFirst
	      		Do Until rsObj_moc_tc.EOF
					v_tem = rsObj_moc_tc(cstr(fname))
					rsObj_moc_tc.MoveNext
				Loop
	         SFIELD=v_tem
	         
	      else
	         SFIELD = ""
	      end if
	   End function
	   
</script>

<TITLE>Tanker Pacific - MOC - Time Charterer Assignment</TITLE>
</HEAD>
<BODY>
<!--#include file="menu_include.asp"-->
<%Response.write Now() &"<BR>" %>
<h3><%= v_header %></h3>
<font color=red size=+2><%= Request.QueryString("v_message")%></font>
<br> 
<form name=thisform  action=moc_tc_asgn_save.asp method=post >
<table border=0>

<%
v_moc = " "
if not (rsObj_moc_tc.eof or rsObj_moc_tc.bof) then 
	
	while not rsObj_moc_tc.eof
		
		if cstr(v_moc) <> cstr(rsObj_moc_tc("moc_id")) then
			Response.Write chr(13)&chr(13)&"<tr><td class=tableheader colspan=" & v_no_cols & ">"&rsObj_moc_tc("short_name")&"-"&rsObj_moc_tc("moc_id")&"</td></tr><tr>"
			v_moc = rsObj_moc_tc("moc_id")
			v_ctr=1
		end if
		
		Response.Write "<td class=tabledata > <INPUT type='checkbox'   name=v_"&rsObj_moc_tc("moc_id")&"_"&rsObj_moc_tc("TIME_CHARTERER_ID")&" value="&getObjValue(rsObj_moc_tc("moc_id"),rsObj_moc_tc("time_charterer_id"))&" "& v_check&">"&rsObj_moc_tc("TC_NAME")&"&nbsp; </td>"
		if cint(v_ctr) = v_no_cols then
			Response.Write "</tr><tr>"
			v_ctr=0
		end if
		
		
		rsObj_moc_tc.movenext
		
		if not rsObj_moc_tc.eof  then
			
				if cstr(v_moc) <> cstr(rsObj_moc_tc("moc_id"))  then
					
						if cint(v_ctr) < cint(v_no_cols) then
							Response.Write "<td class=tabledata colspan="&cint(v_no_cols)-cint(v_ctr)&">&nbsp;</td>"
						end if
						Response.Write "</tr><tr><td colspan="&v_no_cols &">&nbsp;</td></tr>"
				end if	
			
		else
					
					if cint(v_ctr) < cint(v_no_cols) then
						Response.Write "<td class=tabledata colspan="&cint(v_no_cols)-cint(v_ctr)&">&nbsp;</td>"
					end if
			
			Response.Write "</tr>"
		end if	' end of rsObj_moc_tc
	v_ctr=v_ctr+1
	wend 'rsObj_moc_tc

end if 'rsObj_moc_tc
%>
</table>
<%
rsObj_moc_tc.close
set rsObj_moc_tc = nothing
rsObj_asgn.close
set rsObj_asgn = nothing
connObj.close
set connObj = nothing
%>

<INPUT type="submit" value="Submit" id=submit1 name=submit1>
</FORM>
</BODY>
</HTML>
