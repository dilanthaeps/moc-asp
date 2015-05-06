<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->

<%	'===========================================================================
	'	Template Name	:	Inspection Request Maintenance
	'	Template Path	:	ins_request_maint.asp
	'	Functionality	:	To show the list of requests
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	31st August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = true
	
	dim SDATE1, SDATE2, dt1
	dim status, insp_type_disp, v_Page, v_mess, v_button_disabled
	dim v_filter, v_remarks_filter_condition, v_defi_filter_condition
	dim v_expiry_filter_condition, v_selected, v_ctr
	dim strSqlRemStatus, rsObjRemStatus, strSqlDefStatus
	dim rsObjDefStatus, class_color, v_style_start, v_style_end, v_diff_days


	if request("status")="" or isnull(request("status")) then
		status = "ACTIVE"
	else
		status = request("status")
	end if

	if request("status")="All"  then
		status=""
	end if

	if request("insp_type")="" or request("insp_type")= null  then
		insp_type_disp="MOC"
	else
		insp_type_disp=request("insp_type")
	end if

	if insp_type_disp = "All"  then
		insp_type_disp=""
	end if
	
	SDATE1 = request("from_inspection_date")
	SDATE2 = request("to_inspection_date")
	if SDATE1="" and SDATE2="" then
		dt1 = "1 " & monthname(month(now),true) & " " & year(now)
		dt1 = CDate(dt1)
		dt1 = DateAdd("m",-6,dt1)
		'SDATE1 = FormatDateTimeValues(dt1,1)
		SDATE1 = "1 Jul 2002"
		SDATE2 = FormatDateTimeValues(now + 60,1)
	else
		if SDATE1="" then SDATE1 = SDATE2
		if SDATE2="" then SDATE2 = SDATE1
	end if

	Function IIF(expr, trueValue, falseValue)

		If expr Then
			IIF = trueValue
		Else
			IIF = falseValue
		End If
		
	End Function
	
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "attachment; filename=MOC Inspections-" & Day(Date()) & " " & MonthName(Month(Date()),true) & " " & Year(Date()) & ".xls"
%>

<html>
<head>
<TITLE>Vessel Inspections - Tanker Pacific</TITLE>
<META HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 2000 14:25:27 GMT">
<LINK REL="stylesheet" HREF="http://webserve2/wls/moc/moc.css"></LINK>
<style>
.clsIncident
{
	font-family:webdings;
	font-size:15px;
	color:red;
	cursor:default;
}
</style>
</head>
<body>
<table border=0 width=100%>
<tr>
	<td  align=center colspan=1>
		<p>
		<font size=4 style="font-weight:bold" >Inspection Requests </font>
	</td>
</tr>
</table>
<%
v_filter =""
strSql = "SELECT   MIR.REQUEST_ID , MOC_FN_INS_REMARK_COUNT(request_id) remark_count,expences_in_usd,upper(mi.short_name) short_name,"
strSql = strSql & " moc_fn_ins_active_remark_count(request_id) active_remark_count,MIR.VESSEL_CODE ,"
strSql = strSql & " wls_fn_vessel_name(vessel_code) vessel_name,moc_fn_ins_active_def_count(request_id) active_def_count, "
strSql = strSql & " MOC_FN_DEF_COUNT(request_id) def_count,MIR.MOC_ID , wls_fn_moc_name(moc_id) moc_name, "
strSql = strSql & " MIR.INSP_STATUS , to_char(MIR.INSPECTION_DATE,'DD-Mon-YYYY') inspection_date1, "
strSql = strSql & " MIR.INSPECTION_PORT , to_char(MIR.EXPIRY_DATE,'DD-Mon-YYYY') expiry_date1 , "
strSql = strSql & " nvl(trunc(expiry_date),trunc(sysdate))-trunc(sysdate) diff_days, insp_type,"
strSql = strSql & " substr(moc_fn_basis_sire_short_name(mir.basis_sire), 1, 255) basis_sire_name, "
strSql = strSql & " nvl(expiry_date,sysdate)-sysdate diff_days  "
strSql = strSql & " FROM MOC_INSPECTION_REQUESTS MIR,moc_inspectors mi where mir.inspector_id=mi.inspector_id(+) and 1=1 "

if request("fleet_code")="ACTIVE_VESSELS" then
	strSql=strSql & " and MIR.vessel_code in(select vessel_code from wls_vw_vessels_new where fleet_code in('AFRAMAX','FSO','PRODUCT','SUEZMAX','VLCC'))"
	v_filter="Fleet "
elseif request("fleet_code")<>"" then
	strSql=strSql & " and MIR.vessel_code in(select vessel_code from wls_vw_vessels_new where fleet_code='" & request("fleet_code") &"')"
	v_filter="Fleet "
end if

if request("vessel_code")<>"" then
	strSql=strSql & " and vessel_code='" &request("vessel_code") &"'"
	v_filter="Vessel "
end if

if insp_type_disp="INCIDENT" then
	strSql=strSql & " and insp_type ='" & insp_type_disp & "'"
elseif insp_type_disp<>"" then
	'strSql=strSql & " and insp_type='" & insp_type_disp & "'"
	strSql=strSql & " and moc_fn_moc_type(moc_id)='" & insp_type_disp & "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Inspection Type "
	else
		v_filter=v_filter&" Inspection Type "
	end if
end if

if status<>"" then
	strSql=strSql & " and status='" & status& "'"
	if v_filter <> "" then
	v_filter=v_filter&" , Status "
	else
	v_filter=v_filter&" Status "
	end if
end if

if request("moc_id")<>"" then
	strSql=strSql & " and  moc_id=" &request("moc_id")
	if v_filter <> "" then
	v_filter=v_filter&" , MOC "
	else
	v_filter=v_filter&" MOC "
	end if
end if

if request("insp_status")<>"" then
	strSql=strSql & " and insp_status='" & request("insp_status")& "'"
	if v_filter <> "" then
	v_filter=v_filter&" , Inspection Status "
	else
	v_filter=v_filter&" Inspection Status "
	end if
end if

if request("inspection_port")<>"" then
	strSql=strSql & " and inspection_port='" & request("inspection_port")& "'"
	if v_filter <> "" then
	v_filter=v_filter&" , Inspection Port "
	else
	v_filter=v_filter&" Inspection Port "
	end if
end if
if request("cmbInspector")<>"" then
	strSql=strSql & " and mi.inspector_id='" & request("cmbInspector")& "'"
	if v_filter <> "" then
	v_filter=v_filter&" , Inspector "
	else
	v_filter=v_filter&" Inspector "
	end if
end if

strSql=strSql & " and trunc(inspection_date) between '" & SDATE1 & "' and '" & SDATE2 & "'"
if v_filter <> "" then
	v_filter = v_filter & " , Inspection Date "
else
	v_filter=v_filter & " Inspection Date "
end if

v_remarks_filter_condition = ""
If	Trim(Request("v_remarks_filter")) <> "" Then
	v_remarks_filter_condition = " and moc_fn_ins_active_remark_count(request_id) > 0 "
End If

v_defi_filter_condition = ""
If	Trim(Request("v_defi_filter")) <> "" Then
	v_defi_filter_condition = " and moc_fn_ins_active_def_count(request_id) > 0 "
End If

v_expiry_filter_condition = ""
If	Trim(Request("v_expiry_filter")) <> "" Then
	If Trim(Request("v_expiry_filter")) = "2months" Then
		v_expiry_filter_condition = " and nvl(trunc(expiry_date), trunc(sysdate)) - trunc(sysdate) between 0 and 60 and expiry_date is not null "
	ElseIf Trim(Request("v_expiry_filter")) = "overdue" Then
		v_expiry_filter_condition = " and nvl(trunc(expiry_date), trunc(sysdate)) - trunc(sysdate) < 0 "
	End If
End If

strSql = strSql & v_remarks_filter_condition & v_defi_filter_condition & v_expiry_filter_condition

if v_filter ="" then
	v_filter = "No Filter (All records shown)"
end if

if request("item")<>"" then
	strSql=strSql & " order by " & request("item")& " " & request("order")
else
	strSql=strSql & " order by wls_fn_vessel_name(vessel_code), wls_fn_moc_name(moc_id), INSPECTION_DATE "
end if

'Response.Write strSql
'Response.End
Set rsObj=connObj.execute(strSql)

%>
<TABLE WIDTH=100% BORDER="1" CELLPADDING="0" CELLSPACING="1">
  <tr>
    <td class="tableheader" nowrap>Vessel</td>
    <td class="tableheader" valign=top>MOC</td>
    <td class="tableheader" valign=top>Status</td>
    <td class="tableheader" valign=top>Inspection Date</td>
    <td class="tableheader" valign=top>Inspection Port</td>
    <%if request("cmbInspector")<>"" then %>
      <td class="tableheader"  valign="top">Inspector</td>
     <%end if
   if UserIsAdmin then%>
		<td class="tableheader"  valign="top">Cost</td>
	<%End If%>
    <td class="tableheader" valign=top>Expiry</td>
  </tr>
</TABLE>
<TABLE WIDTH=100% BORDER="1" CELLPADDING="0" CELLSPACING="1">
<%
v_ctr=0
if not (rsObj.bof or rsObj.eof) then
while not rsObj.eof
v_ctr=v_ctr+1
	if (v_ctr mod 2) = 0 then
		class_color="columncolor2"
	else
		class_color="columncolor3"
	end if
%>
  <tr valign=bottom>
    <td class="<%=class_color%>"><%=rsObj("vessel_name") %>&nbsp;</td>
    <td class="<%=class_color%>"><%=rsObj("moc_name") %>&nbsp;</td>
    <td class="<%=class_color%>">
<%
		v_style_start = ""
		v_style_end = ""
		v_diff_days=clng(rsObj("diff_days"))
		
		If v_diff_days < 0 then
			v_style_start = "<FONT SIZE='2' COLOR='red'><B><I>"
			v_style_end = "</I></B></FONT>"
			Response.Write v_style_start & "Expired !" & v_style_end
		Else
			If Isnull(rsObj("basis_sire_name")) Or rsObj("basis_sire_name") = Null Or rsObj("basis_sire_name") = "" Then
				Response.Write v_style_start & replace(rsObj("insp_status"),"_"," ") & v_style_end
			Else
				Response.Write v_style_start & replace(rsObj("insp_status"),"_"," ") & " / <FONT SIZE='1' COLOR='darkblue'><B>" & rsObj("basis_sire_name") & "</B></FONT>" & v_style_end
			End If
		End If
%>
		&nbsp;
    </td>
    <td class="<%=class_color%>"><%=mid(rsObj("inspection_date1"),1,7)&mid(rsObj("inspection_date1"),10,2) %>&nbsp;</td>
    <td class="<%=class_color%>"><%=rsObj("inspection_port") %>&nbsp;</td>
    <%if request("cmbInspector")<>"" then %>
      <td class="<%=class_color%>"><%=rsObj("short_name")%>&nbsp;</td>
     <%end if
     if UserIsAdmin then%> 
		<td class="<%=class_color%>"><%=rsObj("expences_in_usd")%>&nbsp;&nbsp;</td>
	<%End If%>
    <td class="<%=class_color%>">
<%
		v_style_start = ""
		v_style_end = ""
		v_diff_days=clng(rsObj("diff_days"))
		if v_diff_days < 0 then
			'Response.Write ("<IMG border=0 src='Images/expired.gif' alt="  & v_diff_days & "> ")
			v_style_start = "<FONT SIZE='2' COLOR='red'><B><I>"
			v_style_end = "</I></B></FONT>"
		elseif v_diff_days = 0 then
			'Response.Write "&nbsp;"
		elseif v_diff_days <= 60 then
			'Response.Write ("<IMG border=0 src='Images/2month2.gif' alt="&v_diff_days&"> ")
			v_style_start = "<FONT SIZE='2' COLOR='red'><B><I>"
			v_style_end = "</I></B></FONT>"
		end if
%>		
		<%=v_style_start & mid(rsObj("expiry_date1"),1,7)&mid(rsObj("expiry_date1"),10,2) & v_style_end %>&nbsp;
		</td>
  </tr>
<%
rsObj.movenext
wend
else
Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if ' if not (rsObj.bof or rsObj.eof) then
%>
</table>
</body>
</html>
<%
rsObj.close
set rsObj=nothing

%>