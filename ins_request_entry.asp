<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Inspection Request Entry
	'	Template Path	:	ins_request_entry.asp
	'	Functionality	:	To allow the entry/edit of the MOC Inspection Request details
	'	Called By		:	ins_request_maint.asp
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	31st  August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = true
	
	dim v_button_disabled, v_mode, v_header, v_vessel_name_disp, v_item, v_tem, Is_Sire
	dim v_vessel_name, v_vessel_short_name, v_inspection_date_disp, v_inspection_port, v_moc_short_name, v_moc_full_name, v_moc_pic
	dim v_imo_number, v_operation, MOCMailBody, AgentMailBody
	dim v_coordinator_name, v_coordinator_design, v_disabled,v_file_name
	dim strSqlVessName, VessMailBody, strSqlVessMailBody
	dim rsObj_insp_status, rsObj_insp_type, rsObj_status, rsObj_Operation, rsObj_Remark_Status
	dim rsObj_tech_status, rsObj_currency, rsObj_vessels, rsObj_moc, rsObj_agent
	dim rsObj_inspector, rsObj_pic, rsObjVessName, RequestOffMailBody, strSqlRequestOffMailBody
	dim rsObjRequestOffMailBody, rsObjVessMailBody,SQL,rsDocs
	dim v_agent_full_name,v_agent_pic,v_agent_email,v_agent_telephone,v_agent_fax,v_agent_mobile
	dim v_inspector_name, v_inspector_company
	
	v_button_disabled = "DISABLED"
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Then
	if UserIsAdmin then
		v_button_disabled = ""
	End If

	Dim Idval
	Idval = Request.QueryString("v_ins_request_id")
	if Idval <> "0" then
		v_mode="edit"
		v_header="Inspection Details"
		if request("v_read_mode") = "Yes" then
		v_header = v_header & "(Read Only)"
		end if
		strSql = "SELECT "
		strSql = strSql & " A.REQUEST_ID , A.INSP_STATUS , A.VESSEL_CODE , A.MOC_ID , A.INSPECTION_PORT, A.IS_SIRE  "
		strSql = strSql & " , to_char(A.INSPECTION_DATE,'DD-Mon-YYYY') INSPECTION_DATE  , A.OPERATION , A.AGENT_ID , A.INSPECTOR_ID , to_char(A.REQUEST_DATE ,'DD-Mon-YYYY') REQUEST_DATE "
		strSql = strSql & " , A.CONFIRMED_OR_REJECTED , to_char(A.DATE_CONFIRM_REJECT,'DD-Mon-YYYY') DATE_CONFIRM_REJECT , A.inspection_remarks , "
		strSql = strSql & " A.DEFICIENCY_RECD , to_char(A.SIRE_RECD_DATE ,'DD-Mon-YYYY') SIRE_RECD_DATE"
		strSql = strSql & " , to_char(A.DATE_REPLIED_TO_SIRE ,'DD-Mon-YYYY') DATE_REPLIED_TO_SIRE, to_char(A.DATE_ACCEPTED,'DD-Mon-YYYY') DATE_ACCEPTED, to_char(A.EXPIRY_DATE ,'DD-Mon-YYYY') EXPIRY_DATE, A.LOCAL_CURRENCY , A.EXPENCES,a.EXPENCES_IN_USD"
		strSql = strSql & " , A.SUPTD_ATTENDED , A.COMBINED , A.BASIS_SIRE , a.BASIS_SIRE_moc BASIS_SIRE_moc_name,to_char( A.VESSEL_ADVISED_DATE ,'DD-Mon-YYYY') VESSEL_ADVISED_DATE"
		strSql = strSql & " , A.OFFICE_ADVISED_DATE ,to_char( A.AGENT_ADVISED_DATE ,'DD-Mon-YYYY') agent_advised_date, A.DATE_DEFS_TO_VESSEL , "
		strSql = strSql & " to_char(A.DATE_DEFS_REPLIED ,'DD-Mon-YYYY') DATE_DEFS_REPLIED, to_char(A.DATE_SIRE_TO_VESSEL,'DD-Mon-YYYY') DATE_SIRE_TO_VESSEL"
		strSql = strSql & " , A.TECH_ALERTED , to_char(A.TECH_ALERT_DATE,'DD-Mon-YYYY') TECH_ALERT_DATE, A.TECH_DEPT_REPLIED , to_char(A.TECH_REPLY_DATE ,'DD-Mon-YYYY') TECH_REPLY_DATE, A.TECH_DECLINED_REASON "
		strSql = strSql & " , A.TECH_STATUS , A.TECH_PIC ,a.status,a.insp_type "
		strSql = strSql & " , to_char(A.CREATE_DATE,'DD-Mon-YYYY') CREATE_DATE, A.CREATED_BY , to_char(a.LAST_MODIFIED_DATE,'DD-Mon-YYYY') LAST_MODIFIED_DATE , A.LAST_MODIFIED_BY "
		strSql = strSql & " , ocimf_report_number, po_number, detention "
		strSql = strSql & " FROM MOC_INSPECTION_REQUESTS A "
		strSql = strSql & " WHERE A.REQUEST_ID =" & Idval
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Create New Inspection Record "
	end if
	if request("v_readonly")="Yes" then
		v_read_mode="Yes"
	end if
	' Select List -      Inspection Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Status' "
	strSql = strSql & "order by sort_order"
	'Response.Write strSql
	set rsObj_insp_status = connObj.Execute(strSql)
	' Select List -       Type
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Inspection_Type' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_insp_type = connObj.Execute(strSql)

	' Select List -       Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Inspection_Status' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_status = connObj.Execute(strSql)

	' Select List -      Operation
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Operation' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Operation = connObj.Execute(strSql)

	' Select List -      Remark_Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Remark_Status' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Remark_Status = connObj.Execute(strSql)

	' Select List -      Tech_Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Tech_Status' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Tech_Status = connObj.Execute(strSql)

	' Select List -      LOCAL Currency
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Currency' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Currency = connObj.Execute(strSql)

	' Select List -      Vessels
	strSql = "SELECT vessel_code, initcap(vessel_name) vessel_name "
	strSql = strSql & "from wls_vw_vessels_new "
	strSql = strSql & " order by vessel_name "
	'Response.Write strSql
	set rsObj_vessels = connObj.Execute(strSql)

	' Select List -      MOC
	strSql = "SELECT moc_id, entry_type, short_name short_name "
	strSql = strSql & "from moc_master "
	strSql = strSql & "order by short_name "
	'Response.Write strSql
	set rsObj_moc = connObj.Execute(strSql)

	' Select List -      Agents
	'strSql = "Select 0 agent_id, 'Select Agent  ' short_name  from dual "
	'strSql = strSql & " union "
	strSql = "SELECT agent_id, short_name "
	strSql = strSql & "from moc_agents_master "
	strSql = strSql & "order by short_name "
	'Response.Write strSql
	set rsObj_agent = connObj.Execute(strSql)

	' Select List -      Inspector
	strSql = "SELECT inspector_id, short_name "
	strSql = strSql & "from moc_inspectors "
	strSql = strSql & "order by short_name "
	'Response.Write strSql
	set rsObj_inspector = connObj.Execute(strSql)

	' Select List -      Person Incharge
	SQL = " SELECT distinct oum.user_id, oum.user_name, oum.user_name email"
	SQL = SQL &  "   FROM oca_user_master oum, oca_user_user_grp_asgn oug, oca_user_group_master ogm"
	SQL = SQL &  "  WHERE (    (oum.user_id = oug.user_id)"
	SQL = SQL &  "         AND (oug.user_grp_id = ogm.user_grp_id)"
	SQL = SQL &  "         AND (oum.status = 'ACTIVE')"
	SQL = SQL &  "         AND (upper(ogm.user_grp_name) in ('MARINE SUPER','OPS SUPER','TECH SUPER'))"
	SQL = SQL &  "        )"
	SQL = SQL &  "  order by upper(oum.user_name)"
	set rsObj_pic = connObj.Execute(SQL)

	v_disabled = "DISABLED"
	
	If v_mode = "edit" and v_button_disabled="" Then
		v_disabled = ""
	End If

'start of office mail body
	RequestOffMailBody = "== CONFIRMATION REQUIRED PLEASE ==" & vbcrlf & vbcrlf

	strSqlRequestOffMailBody = "select mir.moc_id, mm.short_name moc_short_name, mm.full_name moc_full_name, mm.pic moc_pic, v.vessel_name, v.vessel_short_name, mm.entry_type, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "MOC_FN_VESSEL_IMO_NO(v.vessel_code)imo_number, mir.operation, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "mir.inspection_port, to_char(inspection_date, 'DD-Mon-YYYY') inspection_date_disp, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_sys_para_desc('Coordinator Name', 'Coordinator') coordinator_name, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_sys_para_desc('Coordinator Design', 'Coordinator') coordinator_design, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_full_name(mir.agent_id)agent_full_name, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_pic(mir.agent_id)agent_pic, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_telephone(mir.agent_id)agent_telephone, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_fax(mir.agent_id)agent_fax, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_mobile(mir.agent_id)agent_mobile, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "moc_fn_agent_email(mir.agent_id)agent_email, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "MOC_FN_INSPECTOR_NAME(mir.inspector_id)inspector_name, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "MOC_FN_INSPECTOR_COMPANY(mir.inspector_id)inspector_company "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "from moc_inspection_requests mir, moc_master mm, "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "wls_vw_vessels_new v "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "where mir.moc_id = mm.moc_id "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "and mir.vessel_code = v.vessel_code "
	strSqlRequestOffMailBody = strSqlRequestOffMailBody & "and request_id = " & Idval	
	
'Response.Write strSqlRequestOffMailBody
'Response.End

	Set rsObjRequestOffMailBody = connObj.Execute(strSqlRequestOffMailBody)	
	If rsObjRequestOffMailBody.EOF = False Then
		v_vessel_name = rsObjRequestOffMailBody("vessel_name")
		v_vessel_short_name = rsObjRequestOffMailBody("vessel_short_name")
		v_inspection_date_disp = rsObjRequestOffMailBody("inspection_date_disp")
		v_inspection_port = rsObjRequestOffMailBody("inspection_port")
		v_moc_short_name = rsObjRequestOffMailBody("moc_short_name")
		v_moc_full_name = rsObjRequestOffMailBody("moc_full_name")
		v_moc_pic = rsObjRequestOffMailBody("moc_pic")
		v_imo_number = rsObjRequestOffMailBody("imo_number")
		v_operation = rsObjRequestOffMailBody("operation")
		v_coordinator_name = rsObjRequestOffMailBody("coordinator_name")
		v_coordinator_design = rsObjRequestOffMailBody("coordinator_design")
		v_agent_full_name = rsObjRequestOffMailBody("agent_full_name")
		v_agent_pic = rsObjRequestOffMailBody("agent_pic")
		v_agent_email = rsObjRequestOffMailBody("agent_email")
		v_agent_telephone = rsObjRequestOffMailBody("agent_telephone")
		v_agent_fax = rsObjRequestOffMailBody("agent_fax")
		v_agent_mobile = rsObjRequestOffMailBody("agent_mobile")
		v_inspector_name = rsObjRequestOffMailBody("inspector_name")
		v_inspector_company = rsObjRequestOffMailBody("inspector_company")		
		select case rsObjRequestOffMailBody("entry_type")
			case "MOC":v_file_name = v_moc_short_name
			case "PSC","TMNL","FLAG","TVEL":v_file_name = rsObjRequestOffMailBody("entry_type")
		end select
	End If
			
	rsObjRequestOffMailBody.Close
	Set rsObjRequestOffMailBody = Nothing

	RequestOffMailBody = "== CONFIRMATION REQUIRED PLEASE ==" & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "Attention: XXX " & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "Re: " & v_vessel_name & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "We are planning " & v_moc_short_name & " inspection"
	if v_inspection_port<>"" then
		RequestOffMailBody = RequestOffMailBody &  " at " & v_inspection_port
	end if
	RequestOffMailBody = RequestOffMailBody &  " on/around " & v_inspection_date_disp & "." & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "Please confirm the vessel is ready for this Oil Company Inspection." & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "If you feel, for any reason, that this inspection should not be carried out, please advise reasons." & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "Appreciate your immediate reply." & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody &  "Note: A reply must be received within 2 days of this broadcast." & vbcrlf & vbcrlf
	''RequestOffMailBody = RequestOffMailBody &  "Regards / XXX" & vbcrlf
	'RequestOffMailBody = RequestOffMailBody &  "XXX" & vbcrlf
	''RequestOffMailBody = RequestOffMailBody &  v_coordinator_design & vbcrlf & vbcrlf
	''Changes made by Bikash - inputs from Rina Lim 22-jan-2008
	''
	RequestOffMailBody = RequestOffMailBody &  "Brgds / XXX" & vbCrLf
	RequestOffMailBody = RequestOffMailBody &  "MSQV Dept." & vbCrLf
	
	'RequestOffMailBody = RequestOffMailBody &  "Support Services - Marine" & vbcrlf & vbcrlf
	RequestOffMailBody = RequestOffMailBody & vbcrlf & vbcrlf & "Please click the following link to reply:" & vbcrlf & vbcrlf & "http://webserve2/wls/moc/technical_reply.asp?v_ins_request_id=" & Request("v_ins_request_id")
	RequestOffMailBody = RequestOffMailBody & "&v_vessel_name=" & Replace(v_vessel_name, " ", "%20") & "&v_moc_name=" & Replace(v_moc_short_name, " ", "%20")
'end of office mail body

'start of vessel mail body
	VessMailBody = "Dear Capt. <Surname>," & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "Please be advised that we have arranged for a ship inspection of your vessel by "
	VessMailBody = VessMailBody &  v_moc_short_name & " at " & v_inspection_port & " around " & v_inspection_date_disp & ". "
	VessMailBody = VessMailBody &  "Kindly ensure that the vessel is ready in all aspects for said inspection on arrival "
	VessMailBody = VessMailBody &  "and that the housekeeping onboard is maintained to a high standard. We will keep you informed of the inspection arrangement." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "If you feel, for any reason, that this inspection should not be carried out, please respond immediately with reasons." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "A) In preparation for the Inspector, please neatly have following documents ready in a Ship Inspection Reference file, "
	VessMailBody = VessMailBody &  "and hand the compiled documents to him: -" & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	-  Print a clean copy of the VIQ. It is strongly encouraged that completed sections are retained at respective locations e.g. Chapter 4 on Bridge." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Print a clean copy of  the VPQ." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Copy of crew list." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Prepare a matrix showing experience of all officers as per VIQ 3.10." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Copies of all trading certificates if the originals are required to be landed ashore with agents." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Copy of latest Class Survey Status including any COCs and Memos. Please obtain one from Technical Dept immediately prior inspection." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Please also have Officers' and Crew Personnel Certificates and documents folders ready." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Crew Medical Report." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Last D&A certificates and last alcohol test record." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Expiry dates of all LSA/FFA equipments onboard" & vbCrLf
	VessMailBody = VessMailBody &  "	-  List of IMO/OCIMF publications onboard." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Mooring Equipment Folder including BHC certificate tested to rendering force and Mooring equipment inspection records -Refer VIQ 9.4." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Foam Sample Analysis Report." & vbCrLf
	VessMailBody = VessMailBody &  "	-  All small certificates and calibration records." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Record of examination of all Lifting Gear - Refer VIQ 11.32." & vbCrLf
	VessMailBody = VessMailBody &  "	-  Hours of Work and Rest Record." & vbCrLf
	VessMailBody = VessMailBody &  "	-  All approved Manuals, GA/Capacity/Piping diagram(s) or Plan(s) should be kept handy for inspection." & vbCrLf
	VessMailBody = VessMailBody &  "	-  List of work carried out in the yard with class endorsement." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "B) Prepare vessel as per OCIMF VIQ 2009 edition." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "C) Please ensure that the Officer matrix is carefully completed as per VIQ Chap. 3.10." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "D) In addition to VIQ, please ensure the following: -" & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	1.  Both oil record books ( all entries must be as per MARPOL), hot work permit, enclosed space entry forms are correctly filled in and signed by all concerned." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	2.  Cargo/COW plans, ship/shore safety checklists are prepared, completed correctly and complied with strictly." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	3.  Drills as per company and other statutory requirements are carried out and recorded correctly." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	4.  Passage plan made berth to berth: -" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Including parallel indexing and NO GO AREA (marked on charts)" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Interval for fixing position" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Position fixed by at least two means i.e. Primary radar/visual and Secondary GPS" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Tides, UKC and air draft calculated" & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	5.  All officers and crew are well versed with the operation of: -" & vbCrLf
	VessMailBody = VessMailBody &  "		-  LSA + FFA equipments including donning of SCBA" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Emergency Steering Gear" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Emergency Fire Pump" & vbCrLf
	VessMailBody = VessMailBody &  "		-  Operation of Lifeboat engine" & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	6.  All officers aware of operation and calibration of portable equipment." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	7.  All Navigational Publication and Charts corrected to latest notices to mariners available onboard." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	8.  All loose drums, pipes/plates are securely lashed." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	9.  The Deck Air Compressor if onboard must be securely lashed. The battery must be disconnected and removed from the location ( should be" & vbCrLf 
	VessMailBody = VessMailBody &  "	     kept inside the accommodation block area and not on deck)." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	10. Fire wires are rigged properly as per terminal / ISGOTT and tended to at all times." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	      Please advise if you have onboard any of the following:-" & vbCrLf
	VessMailBody = VessMailBody &  "	      - Pipes on deck or in Engine Room." & vbCrLf
	VessMailBody = VessMailBody &  "	      - Soft patches or clamps on lines." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "E) Upon completion of inspection:-" & vbCrLf
	VessMailBody = VessMailBody &  "	1.  Always discuss each observation with the Inspector and endeavour to insert suitable remarks in the observation sheet." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	    Please feel free to call Capt. Gerard (Direct: +65 6433 5130 / Mob: +65 9339 3928) or Capt. Mishra (Direct: +65 6433 5225 / Mob: +65 9638 3922)" & vbCrLf
	VessMailBody = VessMailBody &  "	    for any assistance in this matter. If the inspector does not allow Master to insert any comments, please highlight the fact to us in your message" & vbCrLf
	VessMailBody = VessMailBody &  "	    and send us your comments in the same message." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	2.  Please ensure that all observations are entered into worklist program under assignor ""MOC"" as NCR's for proper tracking and suitable disposition." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "	3.  Within 3 calendar days of the completed inspection, please revert with the report for the following items of each observation to your" & vbcrlf
	VessMailBody = VessMailBody &  "	    Fleet Superintendent, keeping the ""Inspection Team"" in copy: -" & vbCrLf
	VessMailBody = VessMailBody &  "		a) Reason of occurence" & vbCrLf
	VessMailBody = VessMailBody &  "		b) What has been done to rectify the condition" & vbCrLf
	VessMailBody = VessMailBody &  "		c) What preventive measures will be implemented for a lasting solution" & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "Note:-" & vbCrLf
	VessMailBody = VessMailBody &  "====" & vbCrLf
	VessMailBody = VessMailBody &  "Vessel is requested to keep ""Inspection Team"" updated of vessel's discharge schedule, so we can monitor the inspection closely to avoid any" & vbCrLf
	VessMailBody = VessMailBody &  "miscommunication to arrange inspector's boarding to carry out the said inspection. Kindly note that SIRE inspection carried out during discharge" & vbCrLf
	VessMailBody = VessMailBody &  "operation carries more value, as such we prefer inspections be carried out at 1st discharge operation unless otherwise specified." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "Kindly acknowledge receipt." & vbCrLf & vbCrLf
	VessMailBody = VessMailBody &  "Brgds / XXX" & vbCrLf
	VessMailBody = VessMailBody &  "Inspection Team" & vbCrLf & "MSQV Department" & vbCrLf
	VessMailBody = VessMailBody &  "<Team's E-mail ID: inspections@tanker.com.sg>"
'end of vessel mail body

'start of MOC mailbody
	MOCMailBody = "To 	: " & v_moc_full_name & vbcrlf
	MOCMailBody = MOCMailBody &  "Attn 	: " & v_moc_pic & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "From 	: Tanker Pacific Management (S) Pte Ltd" & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "Dear Sir" & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "We would like to request inspection for service approval as per following details:" & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "Vessel Name 	: " & v_vessel_name & vbcrlf
	MOCMailBody = MOCMailBody &  "IMO Number  	: " & v_imo_number & vbcrlf
	MOCMailBody = MOCMailBody &  "ETA         	: " & v_inspection_date_disp & vbcrlf
	MOCMailBody = MOCMailBody &  "Operation   	: " & v_operation & vbcrlf
	MOCMailBody = MOCMailBody &  "At Port     	: " & v_inspection_port & vbcrlf & vbcrlf
	
	MOCMailBody = MOCMailBody &  "Please note agent's details as follows:" & vbcrlf
	MOCMailBody = MOCMailBody &  "Company	: " & v_agent_full_name & vbcrlf
	MOCMailBody = MOCMailBody &  "Tel		: " & v_agent_telephone & vbcrlf
	MOCMailBody = MOCMailBody &  "Fax		: " & v_agent_fax & vbcrlf
	MOCMailBody = MOCMailBody &  "E-mail		: " & v_agent_email & vbcrlf
	MOCMailBody = MOCMailBody &  "PIC		: " & v_agent_pic & " (Mobile: " & v_agent_mobile & ")" & vbcrlf & vbcrlf
	
	MOCMailBody = MOCMailBody &  "Please contact following personnel for inspection details and confirmation of vessel’s ETA:" & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "		Name			Designation                   	 Tel.Office    	Mobile" & vbcrlf
	
	MOCMailBody = MOCMailBody &  "Primary:     	Ms. XXX					Inspection Coordinator		65 6433 5XXX" & vbcrlf
	MOCMailBody = MOCMailBody &  "Secondary:   	Capt. Gerard D’Souza  	Superintendent (Vetting)	65 6433 5130  	65 9339 3928" & vbcrlf
	MOCMailBody = MOCMailBody &  "Alternative: 	Capt. Prashant Mishra  	General Manager				65 6433 5225  	65 9638 3922" & vbcrlf & vbcrlf
	MOCMailBody = MOCMailBody &  "We would appreciate if you could advise us at the earliest." & vbcrlf & vbcrlf & vbcrlf
	
	'--------------
	'MOCMailBody = MOCMailBody &  "For smooth and prompt response, we would appreciate it if all vetting-related communications"
	'MOCMailBody = MOCMailBody &  "are directed to our vessel inspections group email: mocinspectionteam@tanker.com.sg." & vbcrlf & vbcrlf 
	'MOCMailBody = MOCMailBody &  "<Sender's Undersign>"
	'--------------
	'	Bikash - Inputs from Rina Lim - 22/01/08
	'--------------
'end of MOC mailbody

'Start of Agent mailbody
	AgentMailBody = "To   :	" & v_agent_full_name & vbcrlf
	AgentMailBody = AgentMailBody &  "Attn :	" & v_agent_pic & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Cc   :	Capt. <Surname>, Master of " & v_vessel_name & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Fm  :	Tanker Pacific Management (Singapore) Pte Ltd" & vbcrlf & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Dear " & v_agent_pic & "," & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Kindly be advised that we have arranged for an "
	AgentMailBody = AgentMailBody &  v_moc_short_name & " Oil Company inspection onboard the subject vessel at "
	AgentMailBody = AgentMailBody &  v_inspection_port & " on/around " & v_inspection_date_disp & ". "
	AgentMailBody = AgentMailBody &  "For your guidance, " & v_inspector_name & " of " & v_inspector_company
	AgentMailBody = AgentMailBody &  " will endeavour to carry out this inspection on behalf of " & v_moc_short_name & ". "
	AgentMailBody = AgentMailBody &  "The scheduled inspector will liaise with you directly." & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Once he contacts you, please render all assistance to the inspector for boarding and "
	AgentMailBody = AgentMailBody &  "dis-embarking the vessel. Kindly also obtain the necessary clearances from Port and Terminal "
	AgentMailBody = AgentMailBody &  "authorities for him to board the vessel. Additionally, please also keep the inspector informed "
	AgentMailBody = AgentMailBody &  "about the vessel’s itinerary. Please keep us advised of the arrangements and inspection progress in due course." & vbcrlf & vbcrlf
	AgentMailBody = AgentMailBody &  "Please note that expenses incurred for subject inspection are to Owner's account and to be invoiced in final D/A." & vbcrlf & vbcrlf
	
	AgentMailBody = AgentMailBody &  "PLEASE NOTE: IF THE INSPECTING COMPANY OR THE INSPECTOR INFORMS YOU THAT ANY OTHER PERSON BESIDES THE INSPECTOR HIMSELF WILL BE BOARDING THE VESSEL, PLEASE INFORM THIS OFFICE IMMEDIATELY FOR APPROVAL." & vbcrlf & vbcrlf

	'AgentMailBody = AgentMailBody &  "MASTER, Capt. <Surname>, reading in copy, please prepare your good vessel accordingly for this forth coming inspection." & vbcrlf & vbcrlf

	
	'--------------
	'AgentMailBody = AgentMailBody &  "For smooth and prompt response, we would appreciate it if communications related to subject inspection "
	'AgentMailBody = AgentMailBody &  "are directed to our vessel inspections group email : mocinspectionteam@tanker.com.sg." & vbcrlf & vbcrlf
	'AgentMailBody = AgentMailBody &  "<Sender's Undersign>"
	'--------------
	'	Bikash - Inputs from Rina Lim - 22/01/08
	'--------------
	
'end of Agent mailbody
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<meta HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 1996 14:25:27 GMT">
<link REL="stylesheet" HREF="moc.css"></link>
<style>
.clsFile
{
	font-size:9px;
}
</style>
<script language="VBScript" runat="server">
function SFIELD(fname)
	if v_mode="edit" then
	 	'rsObj.MoveFirst
		'	Do Until rsObj.EOF
	 		v_tem = rsObj(cstr(fname))
	 	'	rsObj.MoveNext
	 	'Loop
	   SFIELD=v_tem
	else
	   SFIELD = ""
	end if
End function

function RO
	if v_read_mode="yes" then
		RO=" readonly "
	else
		RO=""
	end if
end function

function ROB ' Read Only Button
	if request("v_read_mode")="Yes" then
		ROB=" disabled "
		else
		ROB=""
	end if
end function
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub TPMDnDCommonCtrl1_OnFileAdded(index)
	set obj = form1.TPMDnDCommonCtrl1
	dim oTr,oTd,sFile,sHTML,sName,sNewName,sFeatures

	sFile = form1.TPMDnDCommonCtrl1.GetFileNameAt(index)
	sFile = Replace(sFile,"'","&#39;")

	for each tr in tbFiles.rows
		if sFile = tr.cells(1).children(0).GetAttribute("filepath") then
			ShowMessage "File already exists in upload list"
			exit sub
		end if
	next

	

	sName = window.showModalDialog("SelectFileType.htm","","dialogHeight:156px;dialogWidth:165px;resizable:yes;center:yes;status:no;scroll:no;")
	sName = "<%=v_vessel_short_name%>/<%=v_file_name%>/" & sName
	
	sNewName = InputBox("Please specify a name for the file you are uploading" & vbcrlf & vbcrlf & _
			"This will be used to describe the document in all displays", "Fileuploader", sName)
	if sNewName<>"" then sName = sNewName
	
	set oTr = tbFiles.insertRow
	oTr.style.backgroundColor="white"
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	sHTML = "<a href='" & sFile & "' filepath='" & sFile & "' target='moc_document'>" & sName & "</a>"
	sHTML = sHTML & "<input name=txtDocID type=hidden value=''><input name=txtDocName type=hidden value='" & sName & "'><input name=txtDocPath type=hidden value='" & sFile & "'>"
	oTd.innerHTML = sHTML
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	oTd.innerHTML = "<span style='cursor:hand;color:blue' onclick='RemoveFile(me.parentElement.parentElement)'>delete</span>"
End Sub

sub RemoveFile(objTr)
	form1.TPMDnDCommonCtrl1.RemoveFileNameFromList(objTr.cells(1).children(0).getAttribute("filepath"))
	tbFiles.deleteRow(objTr.sectionRowIndex)
end sub

sub ShowMessage(s)
	window.status = s
	setTimeout "ClearMessage",2000
end sub

sub ClearMessage
	window.status=""
end sub

function requestOffMail
	dim sTo,sSubject,sBody,sCC

	if form1.tech_pic.value = "" or form1.tech_pic.value = " " then
		MsgBox "Please choose Technical Person Incharge !",vbInformation,"Send Mail to Technical"
		form1.tech_pic.focus
		exit function
	end if

	sTo = form1.tech_pic.options(form1.tech_pic.selectedIndex).value	
	sCC="<Vsl Group>; Inspection Team"
	sBody = txtOfficeMailBody.value
	sBody = Replace(sBody, "<TECH_PIC>",form1.tech_pic.value)
	sSubject = txtVesselMailSubject.value
	sSubject = Replace(sSubject,vbcrlf,"")
	
	form1.mail.displayMailClient cstr(sTo),cstr(sCC), "", cstr(sSubject), cstr(sBody), ""
	
	form1.tech_alert_date.value = day(now) & " " & MonthName(Month(now),true) & " " & year(now)
	form1.submit
end function

function adviseVesselMail
	dim sSubject, sBody,sCC
	sCC="<Vsl Group>; FPD - HR; Inspection Team"
	sSubject = txtVesselMailSubject.value
	sSubject = Replace(sSubject,vbcrlf,"")
	sBody = txtVesselMailBody.value
	
	form1.mail.displayMailClient "<VSL>",cstr(sCC),"",cstr(sSubject),cstr(sBody),""
	
	form1.vessel_advised_date.value = day(now) & " " & MonthName(Month(now),true) & " " & year(now)
	form1.submit
end function

function adviseAgentMail
	dim sSubject, sBody,sCC
	sSubject = txtAgentMailSubject.value
	sSubject = Replace(sSubject,vbcrlf,"")
	sBody = txtAgentMailBody.value
	sCC="<VSL>; <Vsl Group>; Disbursements; Inspection Team"
	form1.mail.displayMailClient "",cstr(sCC),"",cstr(sSubject),cstr(sBody),""
	
	form1.agent_advised_date.value = day(now) & " " & MonthName(Month(now),true) & " " & year(now)
	form1.submit
end function

function callRequestMOCReport()
	dim sSubject, sBody,sCC
	sCC="<Vsl Ops>; Inspection Team"
	sSubject = txtMOCMailSubject.value
	sSubject = Replace(sSubject,vbcrlf,"")
	sBody = txtMOCMailBody.value

	form1.mail.displayMailClient "",cstr(sCC),"",cstr(sSubject),cstr(sBody),""
	
	form1.request_date.value = day(now) & " " & MonthName(Month(now),true) & " " & year(now)
	form1.submit
end function

Sub insp_type_onpropertychange
	dim index
	if window.event.propertyName<>"selectedIndex" then exit sub
	index = form1.insp_type.selectedIndex - 1
	if index=-1 then index=0
	if divMoc.innerHTML = moc_id(index).outerHTML then exit sub
	divMoc.innerHTML = moc_id(index).outerHTML
	divMoc.children(0).style.display = ""
End Sub

Sub window_onload
	if form1.basis_sire_moc_name.value <>"" then
		td_basis_sire.className = "tabledata"
	end if
End Sub
Sub SetStatus
	form1.status.value = statusOuter.value
End Sub
Sub SetDetention
	form1.detention.value = detentionOuter.value
End Sub

-->
</SCRIPT>
<script language="Javascript" src="js_date.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<script language="JavaScript" src="autocomplete.js"></script>
<script LANGUAGE="JAVASCRIPT">
function validate_fields()
{
	if (form1.insp_status.value.length<2)
	{
		alert ("Enter Inspection Status");
		form1.insp_status.focus();
		return false;
	}
	if(form1.vessel_code.value.length<2)
	{
		alert("Enter Vessel Name  ");
		form1.vessel_code.focus();
		return false;
	}
	if(form1.moc_id.value == "")
	{
		alert("Enter MOC Name  ");
		form1.moc_id.focus();
		return false;
	}
	if (form1.insp_type.value == "")
	{
		alert ("Enter Inspection Type");
		form1.insp_type.focus();
		return false;
	}
	if(form1.inspection_date.value.length<2)
	{
		alert("Enter Inspection Date");
		form1.inspection_date.focus();
		return false;
	}
	if(form1.remarks.value.length>1000)
	{
		alert("Remarks Maximum 1000 Chars only");
		form1.remarks.focus();
		return false;
	}
	if(form1.tech_declined_reason.value.length>1000)
	{
		alert("Technical Decline Reason Maximum 1000 Chars only");
		form1.tech_declined_reason.focus();
		return false;
	}
	if(isNaN(form1.EXPENCES_IN_USD.value))
	{
		alert("Expenses in US$ Numeric only!");
		form1.EXPENCES_IN_USD.focus();
		return false;
	}
	if(td_basis_sire.className == "textareah")
	{
		form1.basis_sire.value=""
		form1.basis_sire_moc_name.value=""
	}
	if(form1.status.value=="")
	{
		form1.status.value = statusOuter.value
	}
	if(form1.detention.value=="")
	{
		form1.detention.value = detentionOuter.value
	}
}
function cClose()
{
	var name= confirm("Are you sure? The changes will be lost!!")
	if (name== true)
	{
		v_val = "ins_request_maint.asp?";
		self.opener.form1.action=v_val;
		self.opener.form1.action
		self.opener.form1.submit();
		self.close();
		return false;
	}
	else
	{
	return false;
	}
}

function port_list()
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
	winStats+=',left=460,top=100,width=300,height=470'
	}else{
	winStats+=',screenX=350,screenY=200,width=575,height=400'
	}
	adWindow=window.open("port_list.asp","port_list",winStats);
	adWindow.focus();
}
function getPONumber()
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes,left=460,top=100,width=300,height=470'

	adWindow=window.open("po_list.asp?REQUESTID=<%=IdVal%>","po_list",winStats);
	adWindow.focus();
	return false;
}
function sire_list(vessel_code,request_id)
{
	if(form1.request_id.value == "0")
	{
		alert("Please save the inspection record before slecting the \"Basis SIRE\" inspection");
		return;
	}
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=460,top=100,width=300,height=470'
	}else{
		winStats+=',screenX=350,screenY=200,width=305,height=400'
	}
	adWindow=window.open("base_sire_list.asp?vessel_code="+vessel_code+"&request_id="+request_id,"base_sire_list",winStats);
	adWindow.focus();
}
function fn_merge_call(questionnaire_id,moc_id)
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
	winStats+=',left=460,top=100,width=500,height=270'
	}else{
	winStats+=',screenX=350,screenY=200,width=575,height=400'
	}
	adWindow=window.open("merge_file_call.asp?questionnaire_id="+questionnaire_id+"&moc_id="+moc_id,"base_sire_list",winStats);
	adWindow.focus();
}
function pick_inspection(v_picked_inspection_id)
{
		if (form1.basis_sire.value != '') {
		winStats='toolbar=no,location=no,directories=no,menubar=no,'
		winStats+='scrollbars=yes'
		if (navigator.appName.indexOf("Microsoft")>=0) {
			winStats+=',left=100,top=10,width=720,height=650'
		}else{
		winStats+=',screenX=350,screenY=200,width=575,height=400'
		}
			adWindow=window.open("ins_request_entry.asp?v_ins_request_id="+v_picked_inspection_id+"&v_read_mode=Yes","ro_inspection_request",winStats);
			adWindow.focus();
		}
}

function addEditMOCRecord(MOCID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 175)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("moc_entry.asp?v_child_opener=yes&v_moc_id=" + MOCID, "MOCAddEdit", winStats);

	windNew.focus();

	return false;
}

function addEditAgentRecord(AgentID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 175)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("agent_entry.asp?v_child_opener=yes&v_agent_id=" + AgentID, "agentAddEdit", winStats);

	windNew.focus();

	return false;
}

function addEditInspectorRecord(inspectorID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 175)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("inspector_entry.asp?v_child_opener=yes&v_inspector_id=" + inspectorID, "inspectorAddEdit", winStats);

	windNew.focus();

	return false;
}

function substStar(inVal)
{
	if (inVal == "") return "*"; else return inVal;
}

function callAdviseAgentReport()
{
	var params;
	var noteString;
	
	noteString = "";
	
	noteString = prompt("Enter Additional MOC", "");

	params = "repcallnew20.asp?rr1=moc_advice_agent.rpt";
	params += "&rp1=" + substStar('<% =Request("v_ins_request_id") %>');
	params += "&rp2=" + substStar(noteString);
	params += "&rp3=" + substStar("");
	params += "&rp4=" + substStar("");
	params += "&rp5=" + substStar("");
	params += "&rp6=" + substStar("");
	params += "&rp7=" + substStar("");
	params += "&rp8=" + substStar("");
	params += "&rp9=" + substStar("");
	params += "&rp10=" + substStar("");
	params += "&rp11=" + substStar("");
	params += "&rp12=" + substStar("");
	params += "&rp13=" + substStar("");
	params += "&rp14=" + substStar("");
	params += "&rp15=" + substStar("");
	params += "&rp16=" + substStar("");
	params += "&rp17=" + substStar("");
	params += "&rp18=" + substStar("");
	params += "&rp19=" + substStar("");
	params += "&rp20=" + substStar("");
	params += "&rs1="
	params += "&rs2="
	params += "&rs3="
	params += "&rs4="
	params += "&rs5="
	params += "&rs6="
	params += "&rs7="

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,resizable=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0) 
	{
		winStats += ',left=0,top=0,width=' + (screen.width - 10) + ',height=' + (screen.height - 30)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=350,height=180'
	}

	windName = String(parseInt(String(Math.random() * 1000)));

	repWind = window.open(params, windName, winStats);
	repWind.focus();

	return false;
}

function callRequestMOCReportOLD()
{
	var params;
	var repName, noteString;
	var portName = new String;
	
	portName = form1.inspection_port.value;
	noteString = "";
	
	if (portName.toUpperCase() == "SINGAPORE")
	{
		repName = "request_to_moc_sing.rpt";
		noteString = prompt("Enter Note", "");
	}
	else
	{		
		repName = "request_to_moc.rpt";
	}

	params = "repcallnew20.asp?rr1=" + repName;
	params += "&rp1=" + substStar('<% =Request("v_ins_request_id") %>');
	params += "&rp2=" + substStar(noteString);
	params += "&rp3=" + substStar("");
	params += "&rp4=" + substStar("");
	params += "&rp5=" + substStar("");
	params += "&rp6=" + substStar("");
	params += "&rp7=" + substStar("");
	params += "&rp8=" + substStar("");
	params += "&rp9=" + substStar("");
	params += "&rp10=" + substStar("");
	params += "&rp11=" + substStar("");
	params += "&rp12=" + substStar("");
	params += "&rp13=" + substStar("");
	params += "&rp14=" + substStar("");
	params += "&rp15=" + substStar("");
	params += "&rp16=" + substStar("");
	params += "&rp17=" + substStar("");
	params += "&rp18=" + substStar("");
	params += "&rp19=" + substStar("");
	params += "&rp20=" + substStar("");
	params += "&rs1="
	params += "&rs2="
	params += "&rs3="
	params += "&rs4="
	params += "&rs5="
	params += "&rs6="
	params += "&rs7="

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,resizable=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0) 
	{
		winStats += ',left=0,top=0,width=' + (screen.width - 10) + ',height=' + (screen.height - 30)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=350,height=180'
	}

	windName = String(parseInt(String(Math.random() * 1000)));

	repWind = window.open(params, windName, winStats);
	repWind.focus();

	return false;
}
function rej_remark()
{
	if (document.form1.insp_status.value=='ACCEPTED BASED SIRE'  || document.form1.insp_status.value=='PENDING BASED SIRE' || document.form1.insp_status.value=='REPLIED BASED SIRE')
  	{
  		document.form1.basis_sire_moc_name.className='textyellowcolor';
  		td_basis_sire.className='tabledata';
  	}
  	else
  	{
  		document.form1.basis_sire_moc_name.className='textareah';
  		td_basis_sire.className='textareah';
  		document.form1.basis_sire.value='';
  		document.form1.basis_sire_moc_name.value='';
  	}
}

</script>
<title>Add/Edit Inspection Request Details</title>
</head>
<body class="bcolor">
<%		
	v_vessel_name_disp = ""
	If SFIELD("vessel_code") <> "" Then
		strSqlVessName = "select vessel_name from wls_vw_vessels_new "
		strSqlVessName = strSqlVessName & "where vessel_code = '" & SFIELD("vessel_code") & "'"

		Set rsObjVessName = connObj.Execute(strSqlVessName)
			
		If rsObjVessName.EOF = False Then
			v_vessel_name_disp = " - " & rsObjVessName("vessel_name")
		End If

		rsObjVessName.Close
		Set rsObjVessName = Nothing

	End If

	v_header = v_header & v_vessel_name_disp
	v_header = v_header & "<br>"
	v_header = v_header & "<a href='ins_request_def_maint.asp?v_ins_request_id=" & Idval  & "' target='moc_deficiencies'>"
	v_header = v_header & "<span style='text-align:center;font-size:9;color:blue'>Deficiencies</span></a>"
	v_header = v_header & "&nbsp;&nbsp;"
	v_header = v_header & "<a href='http://webserve2/vid/create_data_file.asp?vessel_code=" & SFIELD("vessel_code") & "&questionnaire_id=10000380' target='vessel_particulars'>"
	v_header = v_header & "<span style='text-align:center;font-size:9;color:blue'>Vessel Particulars</span></a>"
%>
<div style="spacing:0;padding:0;margin:0;">
<span style="float:right;background-color:khaki;padding:5px;border:1px solid blue">
<span>
<b>Status&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b>
<select name="statusOuter" onchange="SetStatus">
<%
if not(rsObj_status.eof or rsObj_status.bof) then
	while not rsObj_status.eof
	%>
	<option value="<%=rsObj_status("sys_para_id")%>" <%if not(isnull(SFIELD("status"))) then%><%if cstr(rsObj_status("sys_para_id"))=cstr(SFIELD("status")) then%>selected<%end if%><%end if%>><%=rsObj_status("para_desc")%></option>
	<%		rsObj_status.movenext
	wend
end if
%>
</select>
</span>
<br>
<span>
<b>Detention</b>
<select name="detentionOuter" onchange="SetDetention" style="width:expression(statusOuter.offsetWidth);">
  <option value='NO'<%if SFIELD("detention")="NO" then Response.Write " selected"%>>No
  <option value='YES'<%if SFIELD("detention")="YES" then Response.Write " selected"%>>Yes
</select>
</span>
</span>
<span style="font-size:20px;font-weight:bold;color:maroon;padding:5px;"><%=v_header%></span>
</div>

<%'start of hidden moc combos%>
<%rsObj_moc.filter="entry_type='MOC'"%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Select MOC--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter="entry_type='PSC'"%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Select PSC--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter="entry_type='TMNL'"%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Select TMNL--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter="entry_type='TVEL'"%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Select USCG--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter="entry_type='FLAG'"%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Select FLAG--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter=0%>
<select name="moc_id" id="moc_id" STYLE="width:150pt;display:none" onkeypress="control_onkeypress" onblur="control_onblur">
  <option value="<%%>">--Reported to--</option>
  <%
			if not(rsObj_moc.eof or rsObj_moc.bof) then
				while not rsObj_moc.eof
		%>
  <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
  <%		rsObj_moc.movenext
				wend
			end if
		%>
</select>
<%rsObj_moc.filter = 0%>
<%'end of hidden moc combos%>

<textarea name="txtVesselMailBody" class=textareah><%=VessMailBody%></textarea>
<textarea name="txtOfficeMailBody" class=textareah><%=RequestOffMailBody%></textarea>
<textarea name="txtMOCMailBody" class=textareah><%=MOCMailBody%></textarea>
<textarea name="txtAgentMailSubject" class=textareah><%="MT " & v_vessel_name & " - " & v_moc_short_name & " Inspection at " & v_inspection_port & " on/around " & v_inspection_date_disp%></textarea>
<textarea name="txtAgentMailBody" class=textareah><%=AgentMailBody%></textarea>
<textarea name="txtVesselMailSubject" class=textareah><%=v_vessel_name & " - " & v_moc_short_name & " Inspection at " & v_inspection_port & " on/around " & v_inspection_date_disp%></textarea>
<textarea name="txtMOCMailSubject" class=textareah><%="MT " & v_vessel_name & " IMO " & v_imo_number & " - Request for " & v_moc_short_name & " Vetting Inspection at " & v_inspection_port & " around " & v_inspection_date_disp%></textarea>

<form name="form1" style="margin-bottom:0" action="ins_request_save.asp" method="post" onsubmit="javascript:return validate_fields();">
<input type=hidden name=status value="<%=SFIELD("status")%>">
<input type=hidden name=detention value="<%=SFIELD("detention")%>">
<object id="mail" style="LEFT: 0px; TOP: 0px" name="mail" codebase="MailClient.CAB" classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" width="0" height="0" VIEWASTEXT>
	<param NAME="_ExtentX" VALUE="26">
	<param NAME="_ExtentY" VALUE="26">
</object>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
    <tr>
      <td class="tableheader">
        <div align="right">Status</div>
      </td>
      <td class="tabledata">
        <select name="insp_status" onchange="javascript:rej_remark();" STYLE="width:150pt">
          <%
				if not(rsObj_insp_status.eof or rsObj_insp_status.bof) then
					while not rsObj_insp_status.eof
			%>
          <option value="<%=rsObj_insp_status("sys_para_id")%>" <%if not(isnull(SFIELD("insp_status"))) then%><%if (cstr(rsObj_insp_status("sys_para_id"))=cstr(SFIELD("insp_status")) or cstr(rsObj_insp_status("sys_para_id"))="Request_to_be_Sent") then%>selected<%end if%><%end if%>><%=rsObj_insp_status("para_desc")%></option>
          <%		rsObj_insp_status.movenext
					wend
				end if
			%>
        </select>
	  </td>
      <td class="textareah" id="td_basis_sire" align="left" colspan=2>
        <span style="height:20px;padding-top:5px" class=tableheader>Basis SIRE</span>

        <span class=tabledata>
      <% if SFIELD("basis_sire_moc_name") > "" then %>
      <a href="Javascript:pick_inspection(form1.basis_sire.value)">
        <input type="hidden" name="basis_sire" id="basis_sire" value="<%=SFIELD("basis_sire")%>">
        <input type="text" id="basis_sire_moc_name" name="basis_sire_moc_name" value="<%=SFIELD("basis_sire_moc_name")%>" readonly style="font-weight:bold;color:blue;cursor:hand" title="<%=SFIELD("basis_sire_moc_name")%> - Click to view the details of the record">
      </a><a href="Javascript:sire_list('<%=SFIELD("vessel_code")%>','<%=SFIELD("request_id")%>')" title="Select from List">Select</a>
      <% else %>
        <input type="hidden" name="basis_sire" value="<%=SFIELD("basis_sire")%>">
        <input type="text" id="basis_sire_moc_name" name="basis_sire_moc_name" value="<%=SFIELD("basis_sire_moc_name")%>" readonly title="<%=SFIELD("basis_sire_moc_name")%> ">
	  	 <a href="Javascript:sire_list('<%=SFIELD("vessel_code")%>','<%=SFIELD("request_id")%>')" title="Select from List">Select</a>
      <%end if %>
      </span>
       </td>
   </tr>
    <tr>
<%
	If Idval <> "0" Then
%>
      <td align="right" class="tabledata">&nbsp;</td>
      <td class="tabledata">
		  <input type="hidden" id="request_id" name="request_id" value="0<%=SFIELD("REQUEST_ID")%>">
		  <input TYPE="hidden" NAME="vessel_code" VALUE="<%=SFIELD("vessel_code")%>">
	  </td>
<%
	Else
%>
      <td align="right" class="tableheader">Vessel</td>
      <td class="tabledata">
		  <input type="hidden" id="request_id" name="request_id" value="0<%=SFIELD("REQUEST_ID")%>">
          <select name="vessel_code" id="vessel_code" onkeypress="control_onkeypress" onblur="control_onblur">
          <option value="<%%>">--Select Vessel--</option>
          <%
					if not(rsObj_vessels.eof or rsObj_vessels.bof) then
						while not rsObj_vessels.eof
				%>
          <option value="<%=rsObj_vessels("vessel_code")%>" <%if not(isnull(SFIELD("vessel_code"))) then%><%if cstr(rsObj_vessels("vessel_code"))=cstr(SFIELD("vessel_code")) then%>selected<%end if%><%end if%>><%=rsObj_vessels("vessel_name")%></option>
          <%			  rsObj_vessels.movenext
						wend
					end if
				%>
        </select>
        <font color="red">*</font></td>
<%
	End If
%>
      <td rowspan="14" valign="top">               
        <table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td class="tabledata" width="150px">
                <input type="button" style="width:140px" value="Request Technical" <%=v_disabled%> name="v_request_office" OnClick="JavaScript:return requestOffMail();">
            <td class="tabledata">
              <input type="text" name="tech_alert_date" value="<%=SFIELD("tech_alert_date")%>" onblur="vbscript:valid_date tech_alert_date,'Technical Alert Date','form1'">
              <a HREF="javascript:show_calendar('form1.tech_alert_date',form1.tech_alert_date.value);">
              <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
              </a> </td>
          </tr>
          <tr>
            <td class="tableheader">
              <div align="right">Technical Reply</div>
            </td>
            <td class="tabledata">
              <select name="tech_status">
                <option value="<%%>">Select Reply Status</option>
                <%
					if not(rsObj_tech_status.eof or rsObj_tech_status.bof) then
						while not rsObj_tech_status.eof
				%>
                <option value="<%=rsObj_tech_status("sys_para_id")%>" <%if not(isnull(SFIELD("tech_status"))) then%><%if cstr(rsObj_tech_status("sys_para_id"))=cstr(SFIELD("tech_status")) then%>selected<%end if%><%end if%>><%=rsObj_tech_status("para_desc")%></option>
                <%		rsObj_tech_status.movenext
						wend
					end if
				%>
              </select>
            </td>
          </tr>
          <tr>
            <td class="tableheader">
              <div align="right">Date Replied</div>
            </td>
            <td class="tabledata">
              <input type="text" name="tech_reply_date" value="<%=SFIELD("tech_reply_date")%>" onblur="vbscript:valid_date tech_reply_date,'Technical Reply Date','form1'">
              <a HREF="javascript:show_calendar('form1.tech_reply_date',form1.tech_reply_date.value);">
              <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
              </a> </td>
          </tr>
          <tr>
            <td class="tableheader">
              <div align="right">Person In Charge</div>
            </td>
            <td class="tabledata">
              <div align="left">
                <select name="tech_pic" id="tech_pic" onkeypress="control_onkeypress" onblur="control_onblur">
                  <option value="<%%>">--Select PIC--</option>
                  <%
					if not(rsObj_pic.eof or rsObj_pic.bof) then
						while not rsObj_pic.eof
				%>
                  <option TITLE="<%=rsObj_pic("email")%>" value="<%=rsObj_pic("user_id")%>" <%if not(isnull(SFIELD("tech_pic"))) then%><%if cstr(rsObj_pic("user_id"))=cstr(SFIELD("tech_pic")) then%>selected<%end if%><%end if%>><%=rsObj_pic("user_name")%></option>
                  <%		
							rsObj_pic.movenext
						wend
					end if
				%>
                </select>
              </div>
            </td>
          </tr>
          <tr>
            <td colspan=2 class="tableheader">Technical Remarks<br>
              <textarea name="tech_declined_reason" rows="5" style="width:100%"><%=SFIELD("tech_declined_reason")%></textarea>
            </td>
          </tr>
          <tr><td colspan=2 style="font-size:5px;">&nbsp;
          <tr>
            <td class="tabledata">
		      <input type="button" name="request_moc" value="Request MOC" style="width:140px" <%=v_disabled%> onclick="javascript:return callRequestMOCReport();">
            <td class="tabledata" height="14">
			  <input type="text" name="request_date" value="<%=SFIELD("request_date")%>" onblur="vbscript:valid_date request_date,'Request Date','form1'">
			  <a HREF="javascript:show_calendar('form1.request_date',form1.request_date.value);">
			  <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
			  </a> </td>
		  <tr>
		    <td class="tableheader num">Date Confirm/Reject
		    <td class="tabledata">
		      <input type="text" name="date_confirm_reject" value="<%=SFIELD("date_confirm_reject")%>" onblur="vbscript:valid_date date_confirm_reject,'Confirm/Reject Date','form1'">
		      <a HREF="javascript:show_calendar('form1.date_confirm_reject',form1.date_confirm_reject.value);">
		      <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
		      </a> </td>
		  </tr>
		  <tr><td colspan=2 style="font-size:5px;">&nbsp;
		  <tr>
		    <td class="tabledata">
              <input type="button" name="advise_vessel" value="Advise Vessel" style="width:140px" <%=v_disabled%> OnClick="javascript:return adviseVesselMail()">
            <td class="tabledata">
              <input type="text" name="vessel_advised_date" VALUE="<%=SFIELD("vessel_advised_date")%>" onblur="vbscript:valid_date vessel_advised_date,'Vessel Advised Date','form1'">
			  <a HREF="javascript:show_calendar('form1.vessel_advised_date',form1.vessel_advised_date.value);">
			  <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
			  </a>
          <tr><td colspan=2 style="font-size:5px;">&nbsp;
          <tr>
            <td class="tabledata">
              <input type="button" name="advise_agent" value="Advise Agent" style="width:140px" <%=v_disabled%> OnClick="JavaScript:return AdviseAgentMail();">
            <td class="tabledata">
              <input type="text" name="agent_advised_date" VALUE="<%=SFIELD("agent_advised_date")%>" onblur="vbscript:valid_date agent_advised_date,'Agent Advised Date','form1'">
              <a HREF="javascript:show_calendar('form1.agent_advised_date',form1.agent_advised_date.value);">
              <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
              </a></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Type</div>
      </td>
      <td width="31%" class="tabledata">
        <select name="insp_type">
          <option value="<%%>">Select Type</option>
          <%
					if not(rsObj_insp_type.eof or rsObj_insp_type.bof) then
						while not rsObj_insp_type.eof
				%>
          <option value="<%=rsObj_insp_type("sys_para_id")%>" <%if not(isnull(SFIELD("insp_type"))) then%><%if cstr(rsObj_insp_type("sys_para_id"))=cstr(SFIELD("insp_type")) then%>selected<%end if%><%end if%>><%=rsObj_insp_type("para_desc")%></option>
          <%		rsObj_insp_type.movenext
						wend
					end if
				%>
        </select>
        <font color="red">*</font>

	&nbsp;&nbsp;
 	<%
	If IsNull(SFIELD("IS_SIRE")) Then
 	Is_Sire = "N" 
	Else
	Is_Sire = SFIELD("IS_SIRE")
	End If
	%>

	Is Sire ? <Input Type=Radio Name="IS_SIRE" Value="Y" <% if Is_Sire = "Y" Then Response.write " checked " %>>Yes &nbsp;&nbsp;<Input Type=Radio Name="IS_SIRE" Value="N" <% if Is_Sire = "N" Then Response.write " checked " %> >No
 



      </td>
    </tr>
    <tr>
      <td class="tableheader">
        <div align="right">MOC</div>
      </td>
      <td class="tabledata" nowrap>
        <div nowrap>
        <span id=divMoc>
        <select name="moc_id" id="moc_id" STYLE="width:150pt" onkeypress="control_onkeypress" onblur="control_onblur">
          <option value="<%%>">--Select MOC--</option>
          <%
					if not(rsObj_moc.eof or rsObj_moc.bof) then
						while not rsObj_moc.eof
				%>
          <option value="<%=rsObj_moc("moc_id")%>" <%if not(isnull(SFIELD("moc_id"))) then%><%if cstr(rsObj_moc("moc_id"))=cstr(SFIELD("moc_id")) then%>selected<%end if%><%end if%>><%=rsObj_moc("short_name")%></option>
          <%		rsObj_moc.movenext
						wend
					end if
				%>
        </select>
        </span>
        <input TYPE="image" NAME="moc_edit_icon" <%=v_button_disabled%> SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit" OnClick="JavaScript:return addEditMOCRecord(form1.moc_id.value)" WIDTH="11" HEIGHT="16"><font color="red">*</font></td>
        </div>
    </tr>
   <tr>
      <td width="16%" class="tableheader">
        <div align="right">Inspection Port</div>
      </td>
      <td width="31%" class="tabledata">
        <input type="text" name="inspection_port" value="<%=SFIELD("inspection_port")%>" maxlength="50">
        <a href="Javascript:port_list()">Select</a> </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Inspection Date</div>
      </td>
      <td width="31%" class="tabledata">
        <input type="text" name="inspection_date" value="<%=SFIELD("inspection_date")%>" onblur="vbscript:valid_date inspection_date,'Inspection Date','form1'">
        <a HREF="javascript:show_calendar('form1.inspection_date',form1.inspection_date.value);">
        <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
        </a> <font color="red">*</font> </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Operation</div>
      </td>
      <td width="31%" class="tabledata">
        <select name="operation">
          <option value="<%%>">--Select Operation--</option>
          <%
					if not(rsObj_Operation.eof or rsObj_Operation.bof) then
						while not rsObj_Operation.eof
				%>
          <option value="<%=rsObj_Operation("sys_para_id")%>" <%if not(isnull(SFIELD("operation"))) then%><%if cstr(rsObj_Operation("sys_para_id"))=cstr(SFIELD("operation")) then%>selected<%end if%><%end if%>><%=rsObj_Operation("para_desc")%></option>
          <%		rsObj_Operation.movenext
						wend
					end if
				%>
        </select>
      </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Agent</div>
      </td>
      <td width="31%" class="tabledata">
        <select name="agent_id" id="agent_id" STYLE="width:150pt" onkeypress="control_onkeypress" onblur="control_onblur">
          <option value="<%%>">--Select Agent--</option>
          <%
					if not(rsObj_agent.eof or rsObj_agent.bof) then
						while not rsObj_agent.eof
				%>
          <option value="<%=rsObj_agent("agent_id")%>" <%if not(isnull(SFIELD("agent_id"))) then%><%if cstr(rsObj_agent("agent_id"))=cstr(SFIELD("agent_id")) then%>selected<%end if%><%end if%>><%=rsObj_agent("short_name")%></option>
          <%		rsObj_agent.movenext
						wend
					end if
				%>
        </select>
		<input TYPE="image" NAME="agent_edit_icon" <% =v_button_disabled %> SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit" OnClick="JavaScript:return addEditAgentRecord(form1.agent_id.value)" WIDTH="11" HEIGHT="16">
      </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Inspector Name</div>
      </td>
      <td width="31%" class="tabledata">
        <select name="inspector_id" id="inspector_id" STYLE="width:150pt" onkeypress="control_onkeypress" onblur="control_onblur">
          <option value="<%%>">--Select Inspector--</option>
          <%
					if not(rsObj_inspector.eof or rsObj_inspector.bof) then
						while not rsObj_inspector.eof
				%>
          <option value="<%=rsObj_inspector("inspector_id")%>" <%if not(isnull(SFIELD("inspector_id"))) then%><%if cstr(rsObj_inspector("inspector_id"))=cstr(SFIELD("inspector_id")) then%>selected<%end if%><%end if%>><%=rsObj_inspector("short_name")%></option>
          <%		rsObj_inspector.movenext
						wend
					end if
				%>
        </select>
		<input TYPE="image" NAME="inspector_edit_icon" <%=v_button_disabled%> SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit" OnClick="JavaScript:return addEditInspectorRecord(form1.inspector_id.value)" WIDTH="11" HEIGHT="16">
      </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader" height="14">
        <div align="right">Date Report Received</div>
      </td>
      <td width="31%" class="tabledata" height="14">
        <input type="text" name="sire_recd_date" value="<%=SFIELD("sire_recd_date")%>" onblur="vbscript:valid_date sire_recd_date,'SIRE Received Date','form1'">
        <a HREF="javascript:show_calendar('form1.sire_recd_date',form1.sire_recd_date.value);">
        <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
        </a> </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader" height="14">
        <div align="right">OCIMF Report No.</div>
      </td>
      <td width="31%" class="tabledata" height="14">
        <input type="text" name="ocimf_report_number" value="<%=SFIELD("ocimf_report_number")%>" maxlength=50>
	  </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Date Report Replied</div>
      </td>
      <td width="31%" class="tabledata">
        <input type="text" name="date_replied_to_sire" value="<%=SFIELD("date_replied_to_sire")%>" onblur="vbscript:valid_date date_replied_to_sire,'Reply to SIRE Date','form1'">
        <a HREF="javascript:show_calendar('form1.date_replied_to_sire',form1.date_replied_to_sire.value);">
        <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
        </a> </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Date of Acceptance</div>
      </td>
      <td width="31%" class="tabledata">
        <input type="text" name="date_accepted" value="<%=SFIELD("date_accepted")%>" onblur="vbscript:valid_date date_accepted,'Date Accepted','form1'">
        <a HREF="javascript:show_calendar('form1.date_accepted',form1.date_accepted.value);">
        <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
        </a> </td>
    </tr>
    <tr>
      <td width="16%" class="tableheader">
        <div align="right">Expiry Date</div>
      </td>
      <td width="31%" class="tabledata">
        <input type="text" name="expiry_date" value="<%=SFIELD("expiry_date")%>" onblur="vbscript:valid_date expiry_date,'Expiry Date','form1'">
        <a HREF="javascript:show_calendar('form1.expiry_date',form1.expiry_date.value);">
        <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
        </a> </td>
    </tr>
    <tr>
      <td colspan="1" class="tableheader" align="right"> Expenses in US$ </td>
      <td width="31%" class="tabledata">
        <input type="text" id="EXPENCES_IN_USD" name="EXPENCES_IN_USD" style="text-align:right" size=10 value="<%=SFIELD("EXPENCES_IN_USD")%>">
        <span style="height:20px;padding-top:5px" class=tableheader>&nbsp;PO&nbsp;</span>
        <input type="text" id="PO_NUMBER" name="PO_NUMBER" size=10 value="<%=SFIELD("PO_NUMBER")%>">
        <input TYPE="image" NAME="po_number_icon" <%=v_button_disabled%> SRC="Images/click_to_open.gif" TITLE="Click to get PO number" OnClick="JavaScript:return getPONumber()" WIDTH="11" HEIGHT="16">
     </td>
    </tr>
    <tr>
      <td colspan=3>
        <table width=100% border="0" cellspacing=1 cellpadding=0>
          <tr>
            <td align="left" width=50% class="tableheader"><strong>Remarks</strong><br>
              <textarea name="remarks" rows="5" class="textarea" style="width:100%"><%=SFIELD("inspection_remarks")%></textarea>
		</table>
      </td>
    </tr>
<%
'documents
SQL = "Select * from moc_documents where deleted is null and doc_type='INSPECTION' and parent_id=" & Idval & " order by doc_id"
set rsDocs = connObj.execute(SQL)
%>  
    <tr>
      <td colspan=3 style="font-size:12px;font-weight:bold;color:red"><br>NOTE: The documents listed here are for internal reference only.<BR>
      Please do not release any MOC related documents to third parties. 
      Kindly consult Capt Mishra if required.
    <tr>
      <td style="vertical-align:top;">
		<OBJECT id=TPMDnDCommonCtrl1 style="height:90px;width:90%;LEFT: 0px; TOP: 0px; BACKGROUND-COLOR: midnightblue" 
		data=data:application/x-oleobject;base64,EcFFRl5khkOn3XMSXMq6vAAHAADYEwAA2BMAAA== 
		classid=clsid:4645C111-645E-4386-A7DD-73125CCABABC codebase="TPMDnDCommon.dll#version=1,0,0,0" VIEWASTEXT></OBJECT>
	  <td colspan=2 style="vertical-align:top;">
	    <table width=100% border=0 cellspacing=1 cellpadding=0 bgcolor=lightgrey>
	      <tr class="tableheader">
	        <td width=15px class=clsFile>
	        <td class=clsFile>File
	        <td width=35px class=clsFile>
	      </tr>
	      <tr bgcolor=white><td colspan=3 class=clsFile>&nbsp;
	      <%
	      while not rsDocs.eof%>
	      <tr bgcolor=white>
	        <td class=clsFile>
	        <td class=clsFile><a href="<%=MOC_PATH%><%=rsDocs("doc_path")%>" target="moc_document"><%=rsDocs("doc_name")%></a>
				<input name=txtDocID type=hidden value="<%=rsDocs("doc_id")%>">
				<input name=txtDocName type=hidden value="<%=rsDocs("doc_name")%>">
				<input name=txtDocPath type=hidden value="">
	        <td class=clsFile>
	        <%if UserIsAdmin then%>
	        <a href="javascript:void(0)" onclick="javascript:window.open('ins_doc_delete.asp?doc_id=<%=rsDocs("doc_id")%>','mocdocdelete','width=400,height=100,top=150,left=150,location=no');">Delete</a>
	        <%end if%>
	      <%
	        rsDocs.movenext
	      wend%>
	      <tbody id=tbFiles class=clsFile>
	      <!--rows will be added dynamically-->
	      </tbody>
	    </table>
    </tr>
    <tr>
      <td colspan="3" class="tabledata" align="center">
        <input type="submit" value="Save" id="submit1" name="submit1" <%=ROB%> <%=v_button_disabled %>>
        &nbsp;
        <input type="submit" value="Save and Close" name="v_save_close" <%=ROB%> <%=v_button_disabled %>>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <button OnClick="javascript:window.close();">Close without Save</button>
      </td>
    </tr>
    <tr>
      <td colspan="3">&nbsp;
    <% if v_mode="edit" then %>
    <tr>
      <td colspan="2" class="tabledata" align="left"> <strong> Created By :</strong>
        <%=SFIELD("created_by")%>&nbsp;&nbsp;<strong>Create Date: </strong><%=SFIELD("create_date")%>
      </td>
      <td colspan="2" class="tabledata" align="left"> <strong>Last Modified By :</strong>
        <%=SFIELD("last_modified_by")%>&nbsp;&nbsp;<strong>Last Modified Date:</strong>
        <%=SFIELD("last_modified_date")%> </td>
    </tr>
    <% end if %>
  </table>
</form>
<font color="red">*</font> Denotes Mandatory Field.<br>

<%'close record set and connection object
if v_mode="edit" then
  rsObj.Close
  Set rsObj=nothing
end if

rsObj_status.close
set rsObj_status = nothing

rsObj_insp_type.close
set rsObj_insp_type = nothing

rsObj_insp_status.close
set rsObj_insp_status = nothing

rsObj_Operation.close
set rsObj_Operation = nothing

rsObj_Remark_Status.close
set rsObj_Remark_Status = nothing

rsObj_Tech_Status.close
set rsObj_Tech_Status = nothing

rsObj_Currency.close
set rsObj_Currency = nothing

rsObj_vessels.close
set rsObj_vessels = nothing

rsObj_moc.close
set rsObj_moc = nothing

rsObj_agent.close
set rsObj_agent = nothing

rsObj_inspector.close
set rsObj_inspector = nothing

rsObj_pic.close
set rsObj_pic = nothing

connObj.Close
set connObj=nothing
%>
<script LANGUAGE="javascript">
<!--

var v_mess="<%=request("v_message") %>";
if (v_mess!="") {
	alert(v_mess);
}
//-->
</script>
</body>
</html>
