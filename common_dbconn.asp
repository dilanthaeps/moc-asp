
<%	'===========================================================================
	'	Template Name	:	Data Base Connection
	'	Template Path	:	.../common_dbconn.asp
	'	Functionality	:	OLEDB Connection (DSN LESS)
	'	Created By		:	Babu.P, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Dim connObj,rsObj,strSql
	dim USER, UserIsAdmin, UserIsSuper
	dim MOC_PATH,MOC_NOTIFICATION_PATH
	MOC_PATH = "\\fileserver\fileuploader\Moc\"
    MOC_NOTIFICATION_PATH="\\fileserver\fileuploader\Moc\Notification\"
	Set connObj=server.CreateObject("ADODB.Connection")
	connObj.Open "Provider=MSDAORA.1;User ID=appln1; Password=1appln;Data Source=ntoracle_2002;"
	'connObj.Open "Provider=MSDAORA.1;User ID=danaos; Password=danaos;Data Source=ntoracle_2002;"
   
	USER = Request.ServerVariables("LOGON_USER")
	if trim(USER)<>"" then
		USER = right(USER, len(USER) - instr(1,USER,"\"))
		USER = ucase(USER)
	end if
	if USER = "MISHRA" or USER = "KARENKWOK" or USER = "GURDEEPJ" or USER = "GERARDD" or USER = "JITENDERS" or USER="SANKAR" then
		UserIsAdmin = true
	else
		UserIsAdmin = false
	end if
	
	DIM rsOCA,str
	str = "Select upper(a.USER_ID) from OCA_USER_MASTER a,OCA_USER_USER_GRP_ASGN b "
	str = str & " where upper(a.USER_ID)='" & USER & "' and a.USER_ID=b.USER_ID "
	str = str & " and a.STATUS='ACTIVE' "
	str = str & " and USER_GRP_ID in (Select USER_GRP_ID from OCA_USER_GROUP_MASTER where upper(USER_GRP_NAME)='OPS SUPER' or upper(USER_GRP_NAME)='TECH SUPER')"
	set rsOCA=connObj.Execute(str)
	if not rsOCA.eof then
		UserIsSuper=true
	else
		UserIsSuper=false
	end if
	rsOCA.close
	set rsOCA=nothing
	
%>