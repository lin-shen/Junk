LPARAMETER loServer

LOCAL loProcess
PRIVATE Request, Response, Server, Session, Process
STORE .null. TO Request, Response, Server, Session, Process

#INCLUDE WCONNECT.H
*#DEFINE TIMEOUT 14400	&& 4 hours - for debugging
#DEFINE TIMEOUT  900	&& 15 minutes - for normal live operation

#DEFINE SOURCE_EXT	".vfphtml"	&& the extension of the script/template source files (including the .)


*** AccessKey settings ***
#DEFINE ACCESSKEYS_ENABLED		.F.		&& Master On/Off switch for accessKeys.
#DEFINE ACCESSKEY_ERROR_TRIMLEN	40		&& See coProcess.SetAccessKey()


*#DEFINE POLICY_CAT_ROOT_NAME	"- Top Level -"	&& This is the name for the non-existant parent of categories at the root level.
*#DEFINE ORG_ROLE_ROOT_NAME		"- Top Level -"	&& This is the name for the non-existant parent of org-roles at the root level.
#DEFINE SENDMAIL_NEWS					1
#DEFINE SENDMAIL_MESSAGE				2
#DEFINE SENDMAIL_LEAVEREQUEST_NEW		3
#DEFINE SENDMAIL_STAFF					4
#DEFINE SENDMAIL_LEAVEREQUEST_ACCEPTED	5
#DEFINE SENDMAIL_LEAVEREQUEST_DECLINED	6
#DEFINE SENDMAIL_PAYSLIP				7

#DEFINE WEBCODE_PAYROLLUSER_BOUNDARY	1000000

#DEFINE MESSAGE_JOINER					" &nbsp;|&nbsp; "	&& added spaces so this can be a wrap-point in the text
#DEFINE MESSAGE_SPLIT_CHAR				'|'
#DEFINE MESSAGE_TRIM					"&nbsp;"
#DEFINE MESSAGE_VALIDATION_PREFIX		"VALIDATION:"

#DEFINE MAX_WORD_LENGTH		80	&& This is the longest allowed word in leaveRequest comments.

&&NOTE: This next one is copied to the top of staffGroupControl.vfphtml - ensure they are equal in both places.
#DEFINE MY_DETAILS_GROUP	-1				&& Used to represent a dummy group containing just the current user.
#DEFINE MY_DETAILS_LABEL	"<My Details>"	&& What the user sees as the text of the above option.
#DEFINE NO_TEMPLATE			-1				&& Used to represent a dummy group containing just the current user.
#DEFINE MY_TEMPLATE			99999				&& Used to represent a dummy group containing just the current user.
#DEFINE MY_TEMPLATE_LABEL	"<My Template>"	&& What the user sees as the text of the above option.
#DEFINE EVERYONE_OPTION		-1				&& Used to represent a selection of "Everyone in the current group".
#DEFINE EVERYONE_LABEL		"<Everyone>"	&& What the user sees as the text of the above option.

* 17/05/2010  CMGM  MRD 4.2.2.2  Default for "Previous Pays" dropdown list
#DEFINE SELECT_VALUE		0
#DEFINE SELECT_LABEL		"<Select>"
* 26/03/2012  RAJ Default for Templates dropdown list
#DEFINE DEF_LABEL_TEMPLATE	"<My Template>"

*!* 30/11/2009;TTP4784;JCF: Split the allowances up into more cols.  Since we might need to re-juggle the cols later it's easier to do it with defines...
&&NOTE: the following must start at one and go up by one with no gaps!
#DEFINE GS_COL_Timesheet			1
#DEFINE GS_COL_Wage					2
*!* 30/11/2009;TTP4784;JCF: Only split Allowances into these 3 for now... may need more later.
#DEFINE GS_COL_Allowances_Amount	3
#DEFINE GS_COL_Allowances_Rate		4
#DEFINE GS_COL_Allowances_Units		5
#DEFINE GS_COL_SickLeave			6
#DEFINE GS_COL_AnnualLeave			7
#DEFINE GS_COL_ShiftLeave			8
#DEFINE GS_COL_OtherLeave			9
#DEFINE GS_COL_LongServiceLeave		10
#DEFINE GS_COL_LieuTime				11
#DEFINE GS_COL_RDO					12
#DEFINE GS_COL_BereavementLeave		13
#DEFINE GS_COL_PublicHoliday		14
#DEFINE GS_COL_AltLeaveAccrued		15
#DEFINE GS_COL_AltLeavePaid			16
#DEFINE GS_COL_DaysPaid				17
#DEFINE GS_COL_RelevantDaysPaid		18
#DEFINE GS_COL_UnpaidLeave          19
&&NOTE: Must be the same value as the last one above..
#DEFINE GS_COL_MAX					19

#DEFINE CHANGED_SLIPNAME	1
#DEFINE CHANGED_ADDRESS		2
#DEFINE CHANGED_SUBURB		3
#DEFINE CHANGED_CITY		4
#DEFINE CHANGED_PHONE		5
#DEFINE CHANGED_EMAIL		6
#DEFINE CHANGED_POSTCODE	7
#DEFINE CHANGED_PASSWORD	8
#DEFINE CHANGED_MOBILE		9
#DEFINE CHANGED_ADDRESS2	10
#DEFINE CHANGED_UDF_L1		11
#DEFINE CHANGED_UDF_L2		12
#DEFINE CHANGED_UDF_D1		13
#DEFINE CHANGED_UDF_D2		14
#DEFINE CHANGED_UDF_C1		15
#DEFINE CHANGED_UDF_C2		16
#DEFINE CHANGED_UDF_C3		17
#DEFINE CHANGED_UDF_N1		18
#DEFINE CHANGED_UDF_M1		19
#DEFINE CHANGED_BANK		20
#DEFINE SMTP_SERVER         "smtp.office365.com:587"
#DEFINE SMTP_USER           "mystaffinfoadmin@myob.com"
#DEFINE SMTP_USER_PASSWORD  "N)pHac4Ty&y<D@vG?:2gU{>Tz"
#DEFINE SMTP_SENDER_NAME    "MyStaffInfo Admin"
#DEFINE SMTP_SENDER_EMAIL   "MyStaffInfoAdmin@myob.com"

*!* 17/03/2011  CMGM  MSI 2011.02  TTP6637  Added nonce
#DEFINE NONCE_DURATION		3600		&& 1 hour (60 mins x 60 secs)
#DEFINE MAX_MESSAGE_SIZE    30000
loProcess = CREATE("coProcess", loServer)

IF VARTYPE(loProcess) != 'O'
	WAIT WINDOW NOWAIT "Unable to create Process object..."
	RETURN .F.
ENDIF

loProcess.Process()

RETURN

*================================================================================*
DEFINE CLASS coProcess AS WWC_Process
*================================================================================*
	cResponseClass	= [WWC_PAGERESPONSE]
	Security		= null
	AppSettings		= null
	AppName			= "MyStaffInfo"
	Licence			= 0
	Employee		= 0
	
	prevLoginDT     = ""

	cDataPath		= ""
	cHtmlPagePath	= ""
	cError			= ""
	cInfo			= ""
	cErrorField		= ""
	cExt			= ""	&& used in security classes
	

	oAccessKeyList	= null
	oAccessItemList	= null
	oAccessHiddenList = null
	lIgnoreAudit	= .F.	&& allow individual process hits to not be audited
	
	timesheet_PayType = 0		&& 1- template, 2 - batch
	
	IsPassReset = "N"
	

	*################################################################################*
	
#DEFINE TOC_Process_

	FUNCTION Process()
		PRIVATE Response, Request, Process, Server, Factory, AppSettings, Security

		This.oRequest.cQueryString = This.SanitiseString(This.oRequest.cQueryString, 1)

		Response	= This.oResponse
		Request		= This.oRequest
		Process		= This
		Server		= This.oServer
		Factory 	= NEWOBJECT("BusinessFactory", "Factory.prg")

		* This would make every page "uncacheable", but as we are running HTTPS, is it needed?
		* it is needed as per security audit particularly for People using public computers
		
		Response.AddForceReload()
		
        **use [http://127.0.0.1/mystaffinfo/companypage.si] as url to test anything locally
        **and use lbIsLocalHost variable to add any exceptions for development environment
         
		LOCAL lbIsLocalHost, lbLocalSecured
		IF AT("127.0.0.1",Request.GetCurrentUrl())>0 
 		   lbIsLocalHost = .T.
 		   IF Request.IsLinkSecure("443")
 		      lbLocalSecured = .T.
 		   ELSE
 		      lbLocalSecured = .F.
 		   ENDIF
 		ELSE
 		   lbIsLocalHost = .F.
 		   lbLocalSecured = .F.
		ENDIF
		* Paths:
		This.cHtmlPagePath	= Server.oConfig.oCoProcess.cHtmlPagePath
		This.cDataPath		= Server.oConfig.oCoProcess.cDataPath

		* Session management:
		This.oSession			= CREATEOBJECT("wwSession")
		Session					= This.oSession
		Session.nSessionTimeout = TIMEOUT

		* Retrieve the Session Cookie:
		lcCookie = Request.GetCookie("mystaffinfo")

		* 23/06/2011  CMGM  2011.03  Stratsec APP-09  Add HttpOnly flag
		* 23/06/2011  CMGM  2011.03  Stratsec APP-10  Set secure flag: send cookies only when going through secure channels
		lcDefaultPath = [/;]
		lcHttpOnly = [HttpOnly;]
		lcSecure = IIF(lbIsLocalHost AND !lbLocalSecured, [], [Secure;])
		lcPath = lcDefaultPath + lcHttpOnly + lcSecure

		* Check if we have a valid Session:
		IF !Session.IsValidSession(lcCookie)
			Session.oRequest = Request
			* We have to create the Session:
			lcCookie = Session.NewSession()

			* 23/06/2011  CMGM  2011.03  Stratsec APP-09, APP-10  Now set the Path properties; Fix above is a work-around -> we are appending other properties to the Path 
			* Response.AddCookie("mystaffinfo", lcCookie)
			Response.AddCookie("mystaffinfo", lcCookie, lcPath)
		ENDIF

		SYS(1104)
		UNLOCK all
		FLUSH force
		CLOSE DATABASES ALL

		This.oAccessKeyList = CREATEOBJECT("COLLECTION")
		This.oAccessItemList = CREATEOBJECT("COLLECTION")
		This.oAccessHiddenList = CREATEOBJECT("COLLECTION")

		* Check if user is logged in:
		Security = Factory.GetSecurityObject()
		
		IF !Security.Login()
			** Login wants to do a Redirect, so we better not do anything else..!
			This.LoginPage()
			RETURN
		ENDIF

		* If we are logged in on this request (form was posted or cookie was already there), we continue here...
		* Reset state again for This user:
		This.Licence	= Session.GetSessionVar("licence")
		This.Employee	= VAL(Session.GetSessionVar("employee"))
		This.prevLoginDT = Session.GetSessionVar("prevLoginDT")
		This.IsPassReset = Session.GetSessionVar("isPassReset")

		This.cExt = "html"

		Factory.cRootPath = ADDBS(This.cDataPath)
		Factory.cDataPath = This.CompanyDataPath()

		USE ADDBS(This.cDataPath) + "globalAppSettings.dbf" ALIAS "globalAppSettings"

		This.AppSettings = Factory.GetApplicationSettingsObject()
		AppSettings = This.AppSettings

		AppSettings.CompanyId = This.Licence

		* CM 02/02/2006 internal portal logging:
		IF FILE(This.CompanyDataPath() + "portalLogging.mem")
			IF !FILE(This.CompanyDataPath() + "siteLog.dbf")
				CREATE TABLE (This.CompanyDataPath() + "siteLog.dbf") FREE (pksitelog Integer AUTOINC, employee Integer, url Character(100), dt T)
				USE IN SELECT("siteLog")
			ENDIF
			IF This.SelectData(This.Licence, "siteLog")
				INSERT INTO siteLog (employee, url, dt) VALUES (This.Employee, This.oRequest.GetLogicalPath(), DATETIME())
			ENDIF
		ENDIF

		* EVENT MANAGEMENT HERE
		* These 3 record events are for synchronsing (uploads):
		BINDEVENT(This, "AddRecord", This, "OnAddRecord", 2)
		BINDEVENT(This, "ChangeRecord", This, "OnChangeRecord", 2)
		BINDEVENT(This, "DeleteRecord", This, "OnDeleteRecord", 2)

		DODEFAULT()
		RETURN .T.
	ENDFUNC

	*################################################################################*
#DEFINE TOC_Pages_

	*> +define: Pages
	* Expand the login page or the getPassword page.
	PROCEDURE LoginPage(tlGetPassword)
		PRIVATE plGetPasswordPage
		LOCAL lcDispPage

		* 08/03/2011  CMGM  MSI 2011.02  TTP6636  Create a new session for every time the login page is requested (end the session here)
		lcCookie = Request.GetCookie("mystaffinfo")
		lcDispPage = Session.GetSessionVar("DispSecureLoginPage")
		IF Session.IsValidSession(lcCookie)
			Session.EndSession()
		ENDIF

		plGetPasswordPage = (Request.QueryString("getPasswd") == "1") OR tlGetPassword

		* This page is not subject to the master layout so we go directly to its template:
		IF lcDispPage="Y"
			Session.SetSessionVar("dispSecureLoginPage", "Y")
			Response.ExpandScript(This.CompanyHtmlPath() + "securelogin" + SOURCE_EXT , Server.nScriptMode)
		ELSE
			Session.SetSessionVar("dispSecureLoginPage", "N")
			Response.ExpandScript(This.CompanyHtmlPath() + "login" + SOURCE_EXT , Server.nScriptMode)
		ENDIF
	ENDPROC

	*================================================================================*

	PROCEDURE AdminPage()
		PRIVATE poPage, poUsers, pcUnitDisp
		
		poUsers = Factory.GetStaffObject()
		poUsers.GetLockedOutUsers()

		poPage = This.NewPageObject("admin:admin", "admin")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE WageTypeAdminPage()
		PRIVATE poPage

		IF !This.SelectData(Process.Licence, "wageType")
			This.AddError("Page Setup Failed!")
		ENDIF

		IF !This.SelectData(Process.Licence, "costcent")
			This.AddError("Page Setup Failed!")
		ENDIF

		IF !This.SelectData(Process.Licence, "allow")
			This.AddError("Page Setup Failed!")
		ENDIF

		poPage = This.NewPageObject("admin:wagetypes", "wageTypeAdmin")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*================================================================================*

	PROCEDURE ChangePasswordPage()
		PRIVATE poPage, poStaff
		PRIVATE pnMinPasswdLen, pnMaxPasswdLen, pnMinAlphaChars, pnMinNumericChars, pnMinOtherChars, plMixedCaseRequired

		pnMinPasswdLen = VAL(AppSettings.Get("passwdMinLength"))
		pnMaxPasswdLen = VAL(AppSettings.Get("passwdMaxLength"))
		pnMinAlphaChars = VAL(AppSettings.Get("passwdMinAlphaChars"))
		pnMinNumericChars = VAL(AppSettings.Get("passwdMinNumericChars"))
		pnMinOtherChars = VAL(AppSettings.Get("passwdMinOtherChars"))
		plMixedCaseRequired = EVALUATE(AppSettings.Get("passwdMixedCaseRequired"))
		poStaff = Factory.GetStaffObject()
		
		LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		
		IF !poStaff.Load(This.Employee)
			This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
		ELSE
			IF poStaff.oData.myPayCode == 0
				This.AddError("Payroll Users cannot change their passwords.")
			ENDIF
		ENDIF
		Session.SetSessionVar("ChangePasswordNonce", lcNonce)

		poPage = This.NewPageObject("home:chgPasswd", "changePassword")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*================================================================================*
	
	
	*================================================================================*

	PROCEDURE ResetPassword()

		LOCAL loStaff, lcNewPasswd0, lcNewPasswd1
		LOCAL lnMinPasswdLen, lnMaxPasswdLen, lnMinAlphaChars, lnMinNumericChars, lnMinOtherChars, llMixedCaseRequired
		LOCAL lnNumLetters, llUpperCase, llLowerCase, lnNumNumbers, lnNumOthers, lnCount, lcChar, lnCode
		LOCAL lcUserCode, lcTokenDBF
		loStaff = Factory.GetStaffObject()
		IF !loStaff.Load(This.Employee)
			This.AddError("Password not changed - Failed to load Employee record: " + loStaff.cErrorMsg)
		ELSE
			IF loStaff.oData.myPayCode == 0
				This.AddError("Payroll Users cannot change their passwords.")
			ELSE
				lcNewPasswd0 = Request.Form("newPasswd0")
				lcNewPasswd1 = Request.Form("newPasswd1")

				IF !(lcNewPasswd0 == lcNewPasswd1)	&& Grr..  != not the same as !(==) (i.e. there is no !== shorthand)
					This.AddError("New Passwords do not match.")
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF lcNewPasswd0 == TRANSFORM(This.Licence) OR lcNewPasswd0 == ALLTRIM(loStaff.oData.myEmail)
					This.AddError("New Passwords must be different your other login details.")
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				lnMinPasswdLen		= VAL(AppSettings.Get("passwdMinLength"))
				lnMaxPasswdLen		= VAL(AppSettings.Get("passwdMaxLength"))
				lnMinAlphaChars		= VAL(AppSettings.Get("passwdMinAlphaChars"))
				lnMinNumericChars	= VAL(AppSettings.Get("passwdMinNumericChars"))
				lnMinOtherChars		= VAL(AppSettings.Get("passwdMinOtherChars"))
				llMixedCaseRequired = EVALUATE(AppSettings.Get("passwdMixedCaseRequired"))

				IF LEN(lcNewPasswd0) < lnMinPasswdLen
					This.AddError(;
						"New Passwords is too short." + CRLF;
						+ "Must be at least " + TRANSFORM(lnMinPasswdLen);
						+ " character" + IIF(lnMinPasswdLen == 1, "", "s") + " long.";
					)
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF LEN(lcNewPasswd0) > lnMaxPasswdLen
					This.AddError(;
						"New Passwords is too long." + CRLF;
						+ "Must be no more than " + TRANSFORM(lnMaxPasswdLen);
						+ " character" + IIF(lnMaxPasswdLen== 1, "", "s") + " long.";
					)
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				lnNumLetters = 0
				lnNumNumbers = 0
				lnNumOthers = 0
				llUpperCase = .F.
				llLowerCase = .F.
				FOR lnCount = 1 TO LEN(lcNewPasswd0)
					lcChar = SUBSTR(lcNewPasswd0, lnCount, 1)
					lnCode = ASC(lcChar)
					DO CASE
					CASE lnCode >= ASC('A') AND lnCode <= ASC('Z')
						lnNumLetters = lnNumLetters + 1
						llUpperCase = .T.
					CASE lnCode >= ASC('a') AND lnCode <= ASC('z')
						lnNumLetters = lnNumLetters + 1
						llLowerCase = .T.
					CASE lnCode >= ASC('0') AND lnCode <= ASC('9')
						lnNumNumbers = lnNumNumbers + 1
					OTHERWISE
						lnNumOthers = lnNumOthers + 1
					ENDCASE
				ENDFOR

				IF lnNumLetters < lnMinAlphaChars
					This.AddError(;
						"Not enough letters used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinNumericChars);
						+ " letter" + IIF(lnMinNumericChars == 1, "", "s");
						+ IIF(llMixedCaseRequired, ", including both upperCase and lowerCase." , ".");
					)
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF llMixedCaseRequired AND !(llUpperCase AND llLowerCase)
					This.AddError("New Password must use both upperCase and lowerCase letters.")
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				IF lnNumNumbers < lnMinNumericChars
					This.AddError(;
						"Not enough numbers used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinNumericChars);
						+ " number" + IIF(lnMinNumericChars == 1, "", "s") + ".";
					)
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				IF lnNumOthers < lnMinOtherChars
					This.AddError(;
						"Not enough non-alphanumeric characters used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinOtherChars);
						+ " non-alphanumeric character" + IIF(lnMinOtherChars == 1, "", "s") + ".";
					)
					Response.Redirect("resetPasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				** Change the password...
				TRY
					IF THIS.selectdata(THIS.licence, "myStaff")
						SELECT myusername FROM mystaff WHERE mywebcode = THIS.employee INTO CURSOR tmpmywebusr
						UPDATE mystaff SET mypassword = lcnewpasswd0, mychanged = STUFF(mychanged, changed_password, 1, 'C') WHERE mywebcode = THIS.employee 
						SELECT tmpmywebusr
						GO TOP
						lcusercode=tmpmywebusr.myusername
						lctokendbf = ADDBS(PROCESS.cdatapath) +  + "\passtokens.dbf"
						IF !EMPTY(lcusercode)
							UPDATE (lctokendbf) SET tokenused=.T.  ;
								WHERE ALLTRIM(UPPER(tokenuser))==ALLTRIM(UPPER(lcusercode)) ;
								AND ALLTRIM(UPPER(tokencomp))==ALLTRIM(UPPER(THIS.licence))
						ENDIF
						USE IN SELECT("tmpmywebusr")
						THIS.adduserinfo("Password has been reset successfully.")
						This.SendPassUpdate()
				    ELSE
						THIS.adderror("Cannot reset password.")
					ENDIF
				CATCH
					THIS.adderror("Cannot reset password.")
				ENDTRY
			ENDIF
		ENDIF
		This.LogOut()
		RETURN
	ENDPROC

	*================================================================================*
	
	
	
	*================================================================================*

	PROCEDURE ResetPasswordPage()
		PRIVATE poPage, poStaff
		PRIVATE pnMinPasswdLen, pnMaxPasswdLen, pnMinAlphaChars, pnMinNumericChars, pnMinOtherChars, plMixedCaseRequired

		pnMinPasswdLen = VAL(AppSettings.Get("passwdMinLength"))
		pnMaxPasswdLen = VAL(AppSettings.Get("passwdMaxLength"))
		pnMinAlphaChars = VAL(AppSettings.Get("passwdMinAlphaChars"))
		pnMinNumericChars = VAL(AppSettings.Get("passwdMinNumericChars"))
		pnMinOtherChars = VAL(AppSettings.Get("passwdMinOtherChars"))
		plMixedCaseRequired = EVALUATE(AppSettings.Get("passwdMixedCaseRequired"))

		poStaff = Factory.GetStaffObject()
		
			
		IF !poStaff.Load(This.Employee)
			This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
		ELSE
			IF poStaff.oData.myPayCode == 0
				This.AddError("Payroll Users cannot change their passwords.")
			ENDIF
		ENDIF
		
		LOCAL lcToken
		lcToken = Request.QueryString("tID")
		llValidToken = Security.ValidateToken(lcToken,.F.)
		IF NOT llValidToken
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		poPage = This.NewPageObject("reset:reset", "resetPassword")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*================================================================================*


	*================================================================================*
	PROCEDURE CompanyPage()
		PRIVATE poPage, poStaff, poMessages, plNewsAccess, plNewsDeleteAccess

		IF This.Employee == 1 OR This.Employee == -999
			* Super-user account - has no access to normal company page - only the admin section
			Response.Redirect("AdminPage.si" + This.AppendMessages('?'))
		ELSE
			* Normal employee - show company page.
			poStaff = Factory.GetStaffObject()
			poMessages = Factory.GetMessagesObject()
			IF !poStaff.Load(This.Employee)
				This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
			ENDIF

			plNewsAccess = This.CheckRights("NEWS_ACCESS") AND This.CheckRights("NEWS_VIEW")
			plNewsDeleteAccess = This.CheckRights("NEWS_ACCESS") AND This.CheckRights("POST_NEWS") AND This.CheckRights("NEWS_DELETE")

			poPage = This.NewPageObject("home:main", "main")
			
			Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ContactInformationPage()
		PRIVATE poPage, poStaff
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager, plViewGroups
		PRIVATE plCanSave, plShowBank, plChangeBank, plShowUserDefined, plAusie, pcExtra, plChangeUD, plPNG
		PRIVATE pcBank, pcBranch, pcAccount, pcSuffix
		PRIVATE plShowBankDetails						&& 25/06/2010  CMGM  MSI 2010.02  Added plShowBankDetails

		plViewGroups = .F.
		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed!")
		ELSE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF

			IF !This.CheckAccess(pnCurrentStaff, plManager)
				This.AddError("You do not have access to this page.")
			ELSE
				poStaff = Factory.GetStaffObject()

				IF !poStaff.Load(pnCurrentStaff)
					This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
				ELSE
					*!* 08/03/2011  CMGM  MSI 2011.02  TTP6417  Reinstate CHANGEBANK from 2011.02: have to split access of contact & bank details
					*!* 20/08/2010  CMGM  MSI 2010.02  TTP5898  New option to change Employee Groups
					IF !plManager
						* If regular employee, only check for personal information
						plCanSave = ( This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee )
						plChangeBank = ( This.CheckRights("CHANGEBANK") AND pnCurrentStaff == This.Employee )
					ELSE
						* If a manager, need to check group.  Dummy Group ("My Details") have value of -1.
						plViewGroups = This.CheckRights("V_GROUP_DETAILS") OR This.CheckRights("C_GROUP_DETAILS")
						IF pnCurrentGroup < 0
							*  If editing own details, need to check for personal information
							plCanSave = ( This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee )
							plChangeBank = ( This.CheckRights("CHANGEBANK") AND pnCurrentStaff == This.Employee )
						ELSE
							*  If editing group details, need to check for "Change Group Contact Details"
							plCanSave = This.CheckRights("C_GROUP_DETAILS")
							plChangeBank = plCanSave
						ENDIF
					ENDIF

					*!* 25/06/2010  CMGM  MSI 2010.02  Always show the bank details section (even if there is none - will show blank fields);
					*!*                               CHANGEBANK is now redundant					
					*!*	plShowBank = !EMPTY(poStaff.oData.myBank) AND This.CheckRights("CHANGEBANK")
					plShowBank = .T.

					*!* 25/06/2010  CMGM  MSI 2010.02  Check if the person logged into the system is the owner of the details;
					*!*                                if "YES", make sure bank account number is viewable
					plShowBankDetails = .F.
					IF pnCurrentStaff == This.Employee
						plShowBankDetails = .T.
					ENDIF

					* HG 11/09/2009 TTP4169, 3471
					* added the checking for the rights of user-defined information
					* HG 28/10/2009 TTP4680, 4588
					* added a new viewing level control for user-defined information
					plShowUserDefined = This.CheckRights("V_UD_INFO") AND !(;
						EMPTY(poStaff.oData.myUdfL1d) AND EMPTY(poStaff.oData.myUdfL2d) AND EMPTY(poStaff.oData.myUdfD1d);
						AND EMPTY(poStaff.oData.myUdfD2d) AND EMPTY(poStaff.oData.myUdfC1d) AND EMPTY(poStaff.oData.myUdfC2d);
						AND EMPTY(poStaff.oData.myUdfC3d) AND EMPTY(poStaff.oData.myUdfN1d) AND EMPTY(poStaff.oData.myUdfM1d);
					)

					*!* 16/03/2011  CMGM  MSI 2011.02  TTP6417  Have to give access to employees' UDF details to managers
					IF !plManager
						* If regular employee, only check for "Change User-Defined Information"
						* HG 28/10/2009 TTP4680, 4588
						* added a seperate control for the right of changing user-defined information
						plChangeUD = This.CheckRights("CHANGEUD")
					ELSE
						* If a manager, need to check group.  Dummy Group ("My Details") have value of -1.
						IF pnCurrentGroup < 0
							*  If editing own details, need to check for "Change User-Defined Information"
							plChangeUD = ( This.CheckRights("CHANGEUD") AND pnCurrentStaff == This.Employee )
						ELSE
							*  If editing group details, need to check for "Change Group Contact Details"
							plChangeUD = This.CheckRights("C_GROUP_DETAILS")
						ENDIF
					ENDIF


					plAusie = This.IsAustralia()
					plPNG = FILE(This.CompanyDataPath() + "png.mem")

					IF plAusie
						pcBank		= SUBSTR(poStaff.oData.myBank, 1, 3)	&& nnn
						pcBranch	= SUBSTR(poStaff.oData.myBank, 5, 3)	&& ___-nnn
						IF plPNG
						   pcAccount	= SUBSTR(poStaff.oData.myBank, 9, 10)	&& ___-___-nnnnnnnnn
						ELSE
						   pcAccount	= SUBSTR(poStaff.oData.myBank, 9, 9)	&& ___-___-nnnnnnnnn
						ENDIF   
						pcSuffix	= ""
					ELSE
						pcBank		= SUBSTR(poStaff.oData.myBank, 1, 2)	&& nn
						pcBranch	= SUBSTR(poStaff.oData.myBank, 4, 4)	&& __-nnnn
						pcAccount	= SUBSTR(poStaff.oData.myBank, 9, 7)	&& __-____-nnnnnnn
						pcSuffix	= SUBSTR(poStaff.oData.myBank, 17, 3)	&& __-____-_______-nnn
					ENDIF

					* CM extra info
					pcExtra = ""
					IF !EMPTY(poStaff.oData.myXML)
						LOCAL poXml
						TRY
							poXml = CREATEOBJECT("MSXML2.DOMDocument")
							poXml.async = .F.
							poXml.validateOnParse = .F.
							IF poXml.LoadXml(poStaff.oData.myXml)
								pcExtra = poXml.documentElement.selectSingleNode("/contact").childNodes(0).text
							ENDIF
						CATCH
							* The extra info is optional so if it fails to load we don't care...
							&&NOTE: this means we can't tell the difference between a not-there and a there-but-can't-load..!
						ENDTRY
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		poPage = This.NewPageObject("home:contact", "contact")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE LocatorBoardPage()
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager
		PRIVATE pnStatuses, pnDateEnabledStatuses, pnTimeEnabledStatuses
		PRIVATE ARRAY paStatuses[1]
		PRIVATE ARRAY paDateEnabledStatuses[1]
		PRIVATE ARRAY paTimeEnabledStatuses[1]
		PRIVATE poStaff, pcDate, pcTime, plCanSave
		LOCAL lcTemp, ldDate

		IF !This.SelectData(This.Licence, "mystaff")
			This.AddError("Page Setup Failed!")
		ELSE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF

			IF !This.CheckAccess(pnCurrentStaff, plManager)
				This.AddError("You do not have access to this page.")
			ELSE
				plCanSave = This.CheckRights("C_LOCATOR")

				poStaff = Factory.GetStaffObject()

				IF !poStaff.Load(pnCurrentStaff)
					This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
				ELSE
					lcTemp = AppSettings.Get("locatorStatuses")
					pnStatuses = ALINES(paStatuses, lcTemp, 5)
					lcTemp = AppSettings.Get("locatorStatusesUsingDate")
					pnDateEnabledStatuses = ALINES(paDateEnabledStatuses, lcTemp, 5)
					lcTemp = AppSettings.Get("locatorStatusesUsingTime")
					pnTimeEnabledStatuses = ALINES(paTimeEnabledStatuses, lcTemp, 5)

					IF EMPTY(poStaff.oData.lbDueBack)
						ldDate = DATE()
					ELSE
						ldDate = poStaff.oData.lbDueBack
					ENDIF
					pcDate = DTOC(ldDate)
					pcTime = LEFT(TTOC(ldDate, 2), 5)

					* Show everyone except for the Admin user
					SELECT ALLTRIM(myStaff.mySurname) + ", " + ALLTRIM(myStaff.myName) as fullName, myWebCode, lbStatus, lbDueBack, lbNotes, LOWER(mySurname), LOWER(myName);
						FROM myStaff;
						WHERE myWebCode != pnCurrentStaff AND myWebCode >= WEBCODE_PAYROLLUSER_BOUNDARY;
						ORDER BY 6, 7;
						INTO CURSOR curLocator
				ENDIF
			ENDIF
		ENDIF

		poPage = This.NewPageObject("home:locator", "locatorBoard")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE PhonelistPage()
		PRIVATE poPage
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager

		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed!")
		ELSE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF
			IF plManager AND pnCurrentGroup<> -1 
				IF !This.GetEmployeesByGroupCode(pnCurrentGroup, "curstaff")
					This.AddError("PhoneList Setup Failed!")
				ELSE
					SELECT curStaff.myWebCode, curStaff.fullName, myStaff.myPhone, myStaff.myEmail, myStaff.myMobile, LOWER(myStaff.mySurname), LOWER(myStaff.myName);
						FROM curStaff JOIN myStaff ON curStaff.myWebCode = myStaff.myWebCode;
						INTO CURSOR curPhoneList;
						ORDER BY 6, 7
				ENDIF
			ELSE
				SELECT myWebCode, ALLTRIM(mySurname) + ", "+ ALLTRIM(myName) as fullName, myPhone, myEmail, myMobile, LOWER(mySurname), LOWER(myName);
					FROM myStaff;
					WHERE myWebCode >= WEBCODE_PAYROLLUSER_BOUNDARY;
					INTO CURSOR curPhoneList;
					ORDER BY 6, 7
			ENDIF
		ENDIF

		poPage = This.NewPageObject("home:phonelist", "phoneList")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE LeaveBalancesPage()
		PRIVATE poPage, poStaff, poBalances
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager, pcUpdated
		LOCAL ltDateTime, loEnts, loBals, loBalObj, llAusie, lnSortOffset, loUnsortedBals, lnI

		pcUpdated = ""

		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed!")
		ELSE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF

			IF !This.CheckAccess(pnCurrentStaff, plManager)
				This.AddError("You do not have access to this page.")
			ELSE
				poStaff = Factory.GetStaffObject()

				IF !poStaff.Load(pnCurrentStaff)
					This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
				ELSE
					lnSortOffset = ASC('A') - 1		&& 1 == 'A' etc
					llAusie = This.IsAustralia()

					loUnsortedBals = CREATEOBJECT("COLLECTION")

					IF This.CheckRights("LV_ANNUAL")
						loEnts = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_ANU_ENT")
							IF !llAusie
								loEnts.Add(poStaff.oData.myHpPerc,	"Entitlement %")
							ENDIF
							IF !EMPTY(poStaff.oData.myHpUnits)
								loEnts.Add(poStaff.oData.myHpEnt,	"Entitlement " + ALLTRIM(poStaff.oData.myHpUnits))
							ENDIF
							IF !llAusie
								loEnts.Add(poStaff.oData.myHpDate,	"Anniversary Date")
							ENDIF
						ENDIF

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_ANU_BAL")
							IF llAusie
								*IF This.CheckRights("LV_ANNUAL_ACCRUED")	//FIXME: shouldn't we be doing something like this for Ausie too?  and also tweaking the balance like below??
									loBals.Add(poStaff.oData.myHpOut,							"Carry Over")
								*ENDIF
								loBals.Add(poStaff.oData.myHpAccrue - poStaff.oData.myHpAdv,	"Year-To-Date")
								loBals.Add(poStaff.oData.myHpTotal,								"Balance")
							ELSE
								IF This.CheckRights("LV_ANNUAL_ACCRUED")
									loBals.Add(poStaff.oData.myHpAccrue,	ALLTRIM(poStaff.oData.myHpUnits) + " Accrued")
								ENDIF
								loBals.Add(poStaff.oData.myHpOut,			ALLTRIM(poStaff.oData.myHpUnits) + " Outstanding")
								loBals.Add(poStaff.oData.myHpAdv,			ALLTRIM(poStaff.oData.myHpUnits) + " Advanced")
								loBals.Add(;
									IIF(This.CheckRights("LV_ANNUAL_ACCRUED"), poStaff.oData.myHpTotal, poStaff.oData.myHpTotal - poStaff.oData.myHpAccrue),;
									"Balance";
								)
							ENDIF
						ENDIF

						loBalObj = This.NewLeaveBalObject("Annual Leave", loEnts, loBals, "MakeLeaveRequestPage.si?type=O")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'a')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF

					IF This.CheckRights("LV_SICK")
						loEnts = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_SIC_ENT")
							loEnts.Add(poStaff.oData.mySpEnt,		"Entitlement " + ALLTRIM(poStaff.oData.mySpUnits))
							IF !llAusie
								loEnts.Add(poStaff.oData.mySpDate,	"Anniversary Date")
								loEnts.Add(poStaff.oData.mySpMax,	"Max. " + ALLTRIM(poStaff.oData.mySpUnits) + " Entitlement")
							ENDIF
						ENDIF

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_SIC_BAL")
							IF llAusie
								loBals.Add(poStaff.oData.mySpOut,								"Carry Over")
								loBals.Add(poStaff.oData.mySpAccrue - poStaff.oData.mySpAdv,	"Year-To-Date")
								loBals.Add(poStaff.oData.mySpTotal,								"Balance")
							ELSE
								loBals.Add(poStaff.oData.mySpTotal,	"Remaining Balance")
							ENDIF
						ENDIF

						loBalObj = This.NewLeaveBalObject(IIF(llAusie, "Personal", "Sick") + " Leave", loEnts, loBals, "MakeLeaveRequestPage.si?type=S")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'b')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF

					IF This.CheckRights("LV_LONG")
						loEnts = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_LNG_ENT")
							loEnts.Add(poStaff.oData.myLslEnt,		"Entitlement " + ALLTRIM(poStaff.oData.myLslUnits))
							IF !llAusie
								loEnts.Add(poStaff.oData.myLslDate,	"Entitlement Date")
							ENDIF
						ENDIF

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_LNG_BAL")
							IF llAusie
								* HG 04/09/2009 TTP1977
								* changed the source of two types ("Prior to 16/08/1978" and "Between 16/08/1978 and 17/08/1993") of leave balances.
								* loBals.Add(poStaff.oData.myLslAccru,							"Prior to 16/08/1978")	&&QUESTION: is this correct?
								* loBals.Add(poStaff.oData.myLslAccru,							"Between 16/08/1978<br/>and 17/08/1993")
								loBals.Add(poStaff.oData.myLsl1978,								"Prior to 16/08/1978")
								loBals.Add(poStaff.oData.myLsl1993,								"Between 16/08/1978<br/>and 17/08/1993")
								loBals.Add(poStaff.oData.myLslOut,								"After 17/08/1993")
								loBals.Add(poStaff.oData.myLslAccru - poStaff.oData.myLslAdv,	"Year-To-Date")			&&QUESTION: given the above, this doesn't look right either..!
								loBals.Add(poStaff.oData.myLslTotal,							"Balance")
							ELSE
								loBals.Add(poStaff.oData.myLslTotal,	ALLTRIM(poStaff.oData.myLslUnits) + " Accrued")
							ENDIF
						ENDIF

						loBalObj = This.NewLeaveBalObject("Long Service Leave", loEnts, loBals, "MakeLeaveRequestPage.si?type=N")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'c')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF

					IF This.CheckRights("LV_ALT")
						loEnts = CREATEOBJECT("COLLECTION")

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_ALT_BAL")
							* Australia has lieu time of days/hours, NZ has alternative leave in days
							loBals.Add(poStaff.oData.myAltTotal,	"Remaining " + ALLTRIM(poStaff.oData.myAltUnits))
						ENDIF

						loBalObj = This.NewLeaveBalObject(IIF(llAusie, "Lieu " + ALLTRIM(poStaff.oData.myAltUnits), "Alternative Leave"), loEnts, loBals, "MakeLeaveRequestPage.si?type=" + IIF(llAusie, 'L', 'Z'))
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'd')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF

					IF llAusie AND This.CheckRights("LV_RDO")
						loEnts = CREATEOBJECT("COLLECTION")

						loBals = CREATEOBJECT("COLLECTION")
						loBals.Add(poStaff.oData.myRdoTotal,	"Remaining Hours")

						loBalObj = This.NewLeaveBalObject("Rostered Day Off", loEnts, loBals, "MakeLeaveRequestPage.si?type=R")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'e')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF
					
					LOCAL lbShiftLeaveForAU, lbShiftLeaveForNZ
					
					lbShiftLeaveForAU = (llAusie AND This.CheckRights("LV_SHIFT"))
					lbShiftLeaveForNZ = (!llAusie AND This.CheckRights("LV_SHIFT") AND !EMPTY(poStaff.oData.myShDate))

					IF lbShiftLeaveForAU OR lbShiftLeaveForNZ
						loEnts = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_SFT_ENT")
							loEnts.Add(poStaff.oData.myShEnt,	"Entitlement " + ALLTRIM(poStaff.oData.myShUnits))
							loEnts.Add(poStaff.oData.myShDate,	"Entitlement Date")
						ENDIF

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_SFT_BAL")
							IF This.CheckRights("LV_SHIFT_ACCRUED")
								loBals.Add(poStaff.oData.myShAccrue,	ALLTRIM(poStaff.oData.myShUnits) + " Accrued")
							ENDIF
							loBals.Add(poStaff.oData.myShOut,			ALLTRIM(poStaff.oData.myShUnits) + " Outstanding")
							loBals.Add(;
								IIF(This.CheckRights("LV_SHIFT_ACCRUED"), poStaff.oData.myShTotal, poStaff.oData.myShTotal - poStaff.oData.myShAccrue),;
								"Total " + ALLTRIM(poStaff.oData.myShUnits) + " Entitlement";
							)
						ENDIF

						loBalObj = This.NewLeaveBalObject(ALLTRIM(poStaff.oData.myShName), loEnts, loBals, "MakeLeaveRequestPage.si?type=F")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'f')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF
					
					LOCAL lbOthLeaveForAU, lbOthLeaveForNZ
					
					lbOthLeaveForAU = (llAusie AND This.CheckRights("LV_OTHER"))
					lbOthLeaveForNZ = (!llAusie AND This.CheckRights("LV_OTHER") AND !EMPTY(poStaff.oData.myOtDate))

					IF lbOthLeaveForAU OR lbOthLeaveForNZ
						loEnts = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_OTH_ENT")
							loEnts.Add(poStaff.oData.myOtEnt,	"Entitlement " + ALLTRIM(poStaff.oData.myOtUnits))
							loEnts.Add(poStaff.oData.myOtDate,	"Entitlement Date")
						ENDIF

						loBals = CREATEOBJECT("COLLECTION")
						IF This.CheckRights("LV_OTH_BAL")
							IF This.CheckRights("LV_OTHER_ACCRUED")
								loBals.Add(poStaff.oData.myOtAccrue,	ALLTRIM(poStaff.oData.myOtUnits) + " Accrued")
							ENDIF
							loBals.Add(poStaff.oData.myOtOut,			ALLTRIM(poStaff.oData.myOtUnits) + " Outstanding")
							loBals.Add(poStaff.oData.myOtAdv,			ALLTRIM(poStaff.oData.myOtUnits) + " Advanced")
							loBals.Add(;
								IIF(This.CheckRights("LV_OTHER_ACCRUED"), poStaff.oData.myOtTotal, poStaff.oData.myOtTotal - poStaff.oData.myOtAccrue),;
								"Total " + ALLTRIM(poStaff.oData.myOtUnits) + " Entitlement";
							)
						ENDIF

						loBalObj = This.NewLeaveBalObject(ALLTRIM(poStaff.oData.myOtName), loEnts, loBals, "MakeLeaveRequestPage.si?type=T")
						IF loBalObj.sortVal > 0
							loUnsortedBals.Add(loBalObj, CHR(lnSortOffset + loBalObj.sortVal) + 'g')	&& suffix is to break sorting equality uniformly.
						ENDIF
					ENDIF

					* 02/11/2009;TTP4702;JCF: For when the user has access to no types, we need to not crash when creating laKeys[]...
					IF loUnsortedBals.Count == 0
						* Nothing to sort so bail out
						poBalances = loUnsortedBals
					ELSE
						LOCAL ARRAY laKeys[loUnsortedBals.Count]
						FOR lnI = 1 TO loUnsortedBals.Count
							laKeys[lnI] = loUnsortedBals.GetKey(lnI)
						NEXT

						ASORT(laKeys, 1, -1, 1, 0)

						poBalances = CREATEOBJECT("COLLECTION")
						FOR lnI = 1 TO loUnsortedBals.Count
							poBalances.Add(loUnsortedBals.Item(laKeys[lnI]))
						NEXT
					ENDIF

					IF FILE("staffCount.dbf")
						* Staffcount is in the root folder.
						SELECT TOP 1 tDateTime FROM staffCount WHERE UPPER(ALLTRIM(cLicence)) == UPPER(ALLTRIM(TRANSFORM(This.Licence))) ORDER BY tDateTime DESC INTO CURSOR curLast
						ltDateTime = curLast.tDateTime
						USE IN SELECT("curLast")

						IF !EMPTY(ltDateTime)
							ltDateTime = This.GetLocalTime(ltDateTime)
							pcUpdated = "Last updated: " + CDOW(ltDateTime) + " " +DMY(ltDateTime) + " " + TTOC(ltDateTime, 2)
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		poPage = This.NewPageObject("leave:balances", "leaveBalances")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE LeaveCalendarPage()
		PRIVATE poPage
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager
		PRIVATE pnMoveDate, pdDate, pdStart, pdFinish, pcShow, poCodes
		LOCAL lcFilter

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "leaveRequestDays"))
			This.AddError("Page Setup Failed!")
		ELSE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF

			IF !This.CheckAccess(pnCurrentStaff, plManager)
				This.AddError("You do not have access to this page.")
			ELSE
				poCodes = This.GetLeaveCodes(pnCurrentStaff, .T.)

				pcShow = Request.QueryString("show")
				pnMoveDate = VAL(Request.QueryString("moveDate"))
				pdDate = CTOD(Request.QueryString("date"))
				
				IF !(VARTYPE(pcShow) == 'C' AND INLIST(pcShow, "all", "approved", "declined", "pending"))
					pcShow = "all"
				ENDIF

				IF VARTYPE(pnMoveDate) != 'N'
					pnMoveDate = 0
				ENDIF

				IF VARTYPE(pdDate) != 'D' OR EMPTY(pdDate)
					pdDate = DATE()
				ENDIF

				IF pnMoveDate == -1 OR pnMoveDate == 1
					* date clicked - move month
					pdDate = GOMONTH(pdDate, pnMoveDate)
				ENDIF

				poRetainList.SetEntry("show", pcShow)
				poRetainList.SetEntry("date", TRANSFORM(pdDate))

				pdStart = DATE(YEAR(pdDate), MONTH(pdDate), 1)
				pdFinish = GOMONTH(pdStart, 1) - 1

				* which employees are we generating for?
				IF plManager
					This.GetEmployeesByGroupCode(pnCurrentGroup, "curCal")
				ELSE
					SELECT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) as fullName;
						FROM myStaff WHERE myWebCode = This.Employee;
						INTO CURSOR curCal
				ENDIF

				DO CASE
					CASE pcShow == "approved"
						lcFilter = "AND leaveRequests.accepted AND !leaverequests.cancelled"
					CASE pcShow == "declined"
						lcFilter = "AND leaveRequests.declined"
					CASE pcShow == "pending"
						lcFilter = "AND !(leaveRequests.accepted OR leaveRequests.declined)"
					OTHERWISE
						lcFilter = "AND !leaverequests.cancelled"
				ENDCASE

				SELECT leaveRequests.*, leaveRequestDays.*, ICASE(leaveRequests.accepted, 0, leaveRequests.declined, 2, 1) AS sortOrder;
					FROM leaveRequests JOIN leaveRequestDays ON leaveRequests.id = leaveRequestDays.leaveReqId;
					WHERE BETWEEN(leaveRequestDays.date, pdStart, pdFinish) ;
					&lcFilter.;
					INTO CURSOR curDays;
					ORDER BY leaveRequests.employee, leaveRequestDays.date, sortOrder
			ENDIF
		ENDIF
		poPage = This.NewPageObject("leave:calendar", "leaveCalendar")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE MakeLeaveRequestPage()
		PRIVATE poPage, poStaff
		PRIVATE pnCurrentStaff, pnCurrentGroup, poRetainList, plManager
		PRIVATE pnStep, poCode, poCodes
		PRIVATE pnStaffHoursPerDay		&& 17/09/2010  CMGM  MSI 2010.03  TTP6014  Employee's Hours Per Standard Day
		PRIVATE ARRAY paDates[1]

        LOCAL lcTypeVal
        lcTypeVal=REQUEST.querystring("type")
        
        DO CASE
        CASE LEN(ALLTRIM(lcTypeVal))> 1
             This.AddError("Invalid page request!")
		CASE !This.SelectData(This.Licence, "myStaff")
			 This.AddError("Page Setup Failed!")
		OTHERWISE
			poRetainList = Factory.GetRetainListObject()

			plManager = .F.
			pnCurrentGroup = 0
			pnCurrentStaff = 0
			pnStaffHoursPerDay = 0		&& 17/09/2010  CMGM  MSI 2010.03  TTP6014  Employee's Hours Per Standard Day
			
			IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
				This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
			ENDIF

			IF !This.CheckAccess(pnCurrentStaff, plManager)
				This.AddError("You do not have access to this page.")
			ELSE
				poStaff = Factory.GetStaffObject()

				IF !poStaff.Load(pnCurrentStaff)
					This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
				ELSE
					pnStep = 1

					poCodes = This.GetLeaveCodes(pnCurrentStaff)

					IF poCodes.Count == 0
						This.AddError("You have no leave types accessible!")
						This.AddUserInfo("Contact your MyStaffInfo administrator to remedy this.")
					ELSE
						IF Request.Form("step") == "1"
							IF EMPTY(Request.Form("days"))
								This.AddValidationError("You must select at least one day to continue!")
								This.cErrorField = "days"
							ELSE
								ALINES(paDates, Request.Form("days"), 0, ',')

								pnStep = 2

								poCode = This.GetLeaveCode(Request.Form("type"), pnCurrentStaff)
								IF ISNULL(poCode)
									This.AddError("Can't get leaveCode for '" + Request.Form("type") + "'")
								ENDIF
								
								* 17/09/2010  CMGM  MSI 2010.03  TTP6014  Employee's Hours Per Standard Day
								pnStaffHoursPerDay = poStaff.oData.myHoursDay
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDCASE

		poPage = This.NewPageObject("leave:make", "makeLeaveRequest")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ViewLeaveRequestPage(tcPage, tnId)
		PRIVATE poPage, poStaff, pcShow, pnSelected, plFiltered, pcOrder, pnRowCount
		PRIVATE plManageMode, plMessageMode, plApproveMode, plDeclineMode	&& these are set in the page HEAD call
		LOCAL lcFilter, lcJoin

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "leaveRequestDays");
		  AND This.SelectData(This.Licence, "leaveRequestStatus"))
			This.AddError("Page Setup Failed!")
		ELSE
			poStaff = Factory.GetStaffObject()

			IF !poStaff.Load(This.Employee)
				This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
			ELSE
				pcOrder = This.SQLSanitise(Request.QueryString("order"))
				IF EMPTY(pcOrder) OR OCCURS(' ', pcOrder) > 1	&& sanitise the input
					pcOrder = "dateMade desc"
				ENDIF
				pcShow = Request.QueryString("show")
				DO CASE
				   CASE !EMPTY(VAL(Request.QueryString("id"))) AND VARTYPE(pcShow)== 'C' AND EMPTY(pcShow)
						pcShow = "all"
				   CASE !(VARTYPE(pcShow) == 'C' AND INLIST(pcShow, "all", "approved", "declined", "pending","unread","downloaded","imported","cancel requested","cancelled"))
					    pcShow = "pending"
				ENDCASE
				        	    
				lcJoin = ""
				DO CASE
					CASE pcShow == "approved"
						lcFilter = "AND lr.accepted AND (!lr.downloaded)"
					CASE pcShow == "declined"
						lcFilter = "AND lr.declined"
					CASE pcShow == "pending"
						lcFilter = "AND !(lr.accepted OR lr.declined)"
					CASE pcShow == "unread"
						lcFilter = "AND (EMPTY(lst.read) AND lst.to="+ALLTRIM(STR(This.Employee,15))+")"
						lcJoin = " JOIN leaveRequestStatus lst ON lr.id=lst.leaveReqID "
					CASE pcShow == "downloaded"
						lcFilter = "AND lr.downloaded AND !lr.imported AND !lr.cancelReq AND !lr.cancelled"
					CASE pcShow == "imported"
						lcFilter = "AND lr.imported"
					CASE pcShow == "cancel requested"
						lcFilter = "AND lr.cancelReq AND !lr.cancelled"
					CASE pcShow == "cancelled"
						lcFilter = "AND lr.cancelled"
					OTHERWISE
						lcFilter = ""
				ENDCASE

				IF VARTYPE(tcPage) == 'C' AND tcPage != "leave:sendMesg"
					lcWho = "lr.manager"
				ELSE
					lcWho = "lr.employee"
				ENDIF
				
				SELECT lr.*, lrd.unitType, SUM(lrd.units) AS units, MIN(lrd.date) AS minDate, MAX(lrd.date) AS maxDate,;
				ICASE(accepted, 0, declined, 2, 1) AS statusSort,;
				ICASE(UPPER(ALLTRIM(unitType)) == "DAYS", SUM(lrd.units) * 24, SUM(lrd.units)) AS unitsSort;
				FROM leaveRequests lr join leaveRequestDays lrd on lr.id = lrd.leaveReqId;
				&lcJoin. ; 
				WHERE &lcWho. = This.Employee;
				&lcFilter.;
				INTO CURSOR curRequests;
				GROUP BY lr.id;
				ORDER BY &pcOrder.


				pnRowCount = _TALLY

				pnSelected = EVL(tnId, VAL(Request.QueryString("id")))
				plFiltered = (pnSelected < 0)
				pnSelected = ABS(pnSelected)
			ENDIF
		ENDIF

		poPage = This.NewPageObject(IIF(VARTYPE(tcPage) == 'C', tcPage, "leave:view"), "viewLeaveRequest")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	&&NOTE: these 4 are very similar in that they all go to the viewLeaveRequest template but have different modes, and may require an ID to be passed.
	PROCEDURE ManageLeaveRequestPage()
		This.ViewLeaveRequestPage("leave:manage")
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE SendLeaveRequestMessagePage(tcPage)
		LOCAL lnId, llManagerMessage

		lnId = VAL(Request.QueryString("id"))
		llManagerMessage = !EMPTY(Request.QueryString("toEmp"))

		IF EMPTY(lnId)
			This.AddError("No Leave Request specified!")
		ENDIF

		This.ViewLeaveRequestPage(EVL(tcPage, IIF(llManagerMessage, "leave:sendMgrMesg", "leave:sendMesg")), lnId)
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE ApproveLeaveRequestPage()
		This.SendLeaveRequestMessagePage("leave:approve")
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE DeclineLeaveRequestPage()
		This.SendLeaveRequestMessagePage("leave:decline")
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE InboxPage()
		PRIVATE poPage, poStaff, poInbox, pcOrder, pnRowCount

		poStaff = Factory.getStaffObject()
		IF !poStaff.Load(This.Employee)
			This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
		ELSE
			pcOrder = This.SQLSanitise(Request.QueryString("order"))
			IF EMPTY(pcOrder) OR OCCURS(' ', pcOrder) > 1	&& sanitise the input further
				pcOrder = "meDate desc"
			ENDIF

			poInbox = Factory.GetMessagesObject()
			IF !poInbox.open()
				This.AddError("Failed to load Inbox: " + poInbox.cErrorMsg)
			ELSE
				poInbox.execute([SELECT meId, meFromName, meDate, meSubject, meComplete FROM myMessages WHERE meToId = Process.Employee and meType="IN" ORDER BY &pcOrder. INTO CURSOR curMessages])
				pnRowCount = _TALLY
			ENDIF
		ENDIF

		poPage = This.NewPageObject("messages:inbox", "inbox")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE OutboxPage()
		PRIVATE poPage, poStaff, poOutbox, pcOrder, pnRowCount

		poStaff = Factory.getStaffObject()
		IF !poStaff.Load(This.Employee)
			This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
		ELSE
			pcOrder = This.SQLSanitise(Request.QueryString("order"))
			IF EMPTY(pcOrder) OR OCCURS(' ', pcOrder) > 1	&& sanitise the input further
				pcOrder = "meDate desc"
			ENDIF

			poOutbox = Factory.GetMessagesObject()
			IF !poOutbox.open()
				This.AddError("Failed to load Outbox: " + poOutbox.cErrorMsg)
			ELSE
				poOutbox.execute([SELECT meId, meToName, meType, meDate, meSubject, meComplete FROM myMessages WHERE meFromId = Process.Employee and INLIST(meType, "OUT", "News") ORDER BY &pcOrder. INTO CURSOR curMessages])
				pnRowCount = _TALLY
			ENDIF
		ENDIF

		poPage = This.NewPageObject("messages:outbox", "outbox")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ReadMessagePage()
		PRIVATE poPage, poStaff, pcFromPage, pcType, lnMessageId, poMessage, pcFromPage

		poStaff = Factory.getStaffObject()
		IF !poStaff.Load(This.Employee)
			This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
		ELSE
			pnMessageId = VAL(Request.QueryString("meId"))
			IF EMPTY(pnMessageId)
				This.AddError("No Message to read!")
			ELSE
				pcFromPage = Request.QueryString("fromPage")
				IF '/' $ pcFromPage
					* Sanitise the return url to ensure it is local!
					pcFromPage = RIGHT(pcFromPage, LEN(pcFromPage) - RAT('/', pcFromPage))
				ENDIF

				poMessage = Factory.getMessagesObject()
				IF !poMessage.Load(pnMessageId)
					This.AddError("No Message to Read: " + poMessage.cError)
				ELSE
					IF !This.CheckMessageAccess(pnMessageId)
						This.AddError("You do not have access to this message!")
					ELSE
						pcType = ALLTRIM(poMessage.oData.meType)

						IF !(poMessage.oData.meComplete OR pcType == "Leave";
						  OR pcType == "News" AND "Outbox" $ pcFromPage)		&& if not Leave and not News from the Outbox page...
							* ...mark as read:	//NOTE: the first person to read a news item marks it read for everyone..!
							poMessage.oData.meComplete = .T.
							IF poMessage.Save()
								This.AddUserInfo("Marked as Read.")
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		poPage = This.NewPageObject("messages:read", "readMessage")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE SendMessagePage(toMessage, tcPageType)
		PRIVATE poPage, poMessage, pcPageType, poForm
		LOCAL lnTo

		poForm = null	&& initialise the storage so it's acccessible across calls for HEAD and MAIN

	
		IF VARTYPE(toMessage) == 'O' AND VARTYPE(tcPageType) == 'C'
			pcPageType = tcPageType
			poMessage = toMessage
		ELSE
			poMessage = Factory.GetMessagesObject()
			IF !poMessage.ConstructMessage(This.Employee)
				This.AddError("Could create new message: " + poMessage.cError)
			ELSE
				lnTo = VAL(Request.QueryString("id"))
				IF lnTo != 0
					poMessage.oData.meToId = lnTo
				ENDIF

				pcPageType = "Send"
			ENDIF
		ENDIF

		IF !This.CheckMessageAccess(poMessage.oData.meId)
			This.AddError("You do not have access to this message!")
		ENDIF

		poPage = This.NewPageObject("messages:send", "editMessage")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE SendNewsPage()
		PRIVATE poPage, poMessage

		poMessage = Factory.GetMessagesObject()
		IF poMessage.ConstructMessage(This.Employee)
			poMessage.oData.meDate = This.GetLocalTime(DATETIME())
		ELSE
			This.AddError("Could create new message: " + poMessage.cError)
		ENDIF

		poPage = This.NewPageObject("messages:news", "editNews")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*


	************************************************************************************
	PROCEDURE GetINISection(tcKey AS STRING, tcSection AS STRING, tcINI AS STRING)
	LOCAL lcBuffer, lcINI, lcSection

	DECLARE INTEGER GetPrivateProfileString IN Kernel32 STRING, STRING, STRING, STRING @, INTEGER, STRING

	lcBuffer = SPACE(255)

	* Note: ignoring return value:
	GetPrivateProfileString(tcSection, tcKey, "", @lcBuffer, LEN(lcBuffer), tcINI)

	lcBuffer = ALLTRIM(STRTRAN(lcBuffer, CHR(0), ""))

	RETURN lcBuffer
	ENDFUNC
	*****************************************************************************************
	

	PROCEDURE GroupSummaryPage()
			
        cCSVFile = ""
		this.GroupSummaryProc(.F.)
	ENDPROC
				
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE GroupSummaryCSV()
        	this.GroupSummaryProc(.T.)
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE GroupSummaryProc(tlPrint)
		PRIVATE poPage, poStaff
		PRIVATE pnCurrentPay, pnCurrentStaff, pnCurrentGroup, poRetainList, plManager, pcLinkToOut
		PRIVATE plAustralia, pcHpUnits, pcShName, pcShUnits, pcOtName, pcOtUnits, pcLslUnits, pcAltUnits
		PRIVATE pnPayCount, plPayOpen

		PRIVATE rcPayInfo, rcGroupInfo, rcStaffInfo, rcApprover, rdPayDate, rcCompName

		LOCAL lnI, lcType, lcPrefix, lcDescription, llCount, lnUnits, lnEntries, lnUnapprovedEntries, lcStatus, lcStatusDesc
		LOCAL lnAllApproved, lnNoneApproved, lnSomeApproved, lcLinkClass, lcToolTip, llAuthz, llAuthz2, lcFilter
		LOCAL lcCodeFilter, lnArgCount, lcJoiner

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "myPays");
		  AND This.SelectData(This.Licence, "myGroups");
		  AND This.SelectData(This.Licence, "allow");
		  AND This.SelectData(This.Licence, "timesheet"))
			This.AddError("Page Setup Failed!")
		ELSE
			poStaff = Factory.GetStaffObject()

			IF !poStaff.Load(This.Employee)
				This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
			ELSE
				poRetainList = Factory.GetRetainListObject()

				plManager = .F.
				pnCurrentGroup = 0
				pnCurrentStaff = 0
				IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList, .T.)	&& 19/11/2009;TTP4845;JCF: Default to group rather than My Details.
					This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
				ENDIF

				poRetainList.RemoveEntry("currentStaff")	&& unused as we are only showing Groups here.

				IF !This.CheckAccess(pnCurrentStaff, plManager)
					This.AddError("You do not have access to this page.")
				ELSE
					pnCurrentPay = 0
					plPayOpen = .F.
					pnPayCount = This.SetupPayControlData(@pnCurrentPay, @plPayOpen, poRetainList, "open")
					IF pnPayCount < 0
						This.AddError("PayControl Setup Failed!")
					ELSE
						IF pnPayCount == 0
							pnCurrentPay = -1
						ENDIF

						plAustralia = This.IsAustralia()

						pcHpUnits	= poStaff.oData.myHpUnits
						pcShName	= poStaff.oData.myShName
						pcShUnits	= poStaff.oData.myShUnits
						pcOtName	= poStaff.oData.myOtName
						pcOtUnits	= poStaff.oData.myOtUnits
						pcLslUnits	= poStaff.oData.myLslUnits
						pcAltUnits	= poStaff.oData.myAltUnits

						&& iconImg := "Green"|"Orange"|"Red"
						&& iconDesc := "All entries approved"|"Some entries approved"|"No entries approved"
						*!* 23/11/2009;TTP4868;JCF: Added the entries count to the cursor we create here so we can differentiate between no-entries and zero-total-entries.
						*!* 30/11/2009;TTP4874;JCF: Split the allowances up into more cols...
						CREATE CURSOR curSummary (;
							iconImg	C(6),;
							iconDesc C(21),;
							canApprove L,;
							canUnapprove L,;
							employee I,;
							empName	C(100),;
							classM C(16),	tipM C(60),		unitsM B(2),	entriesM I,;
							classW C(16),	tipW C(60),		unitsW B(2),	entriesW I,;
							classA1 C(16),	tipA1 C(60),	unitsA1 B(2),	entriesA1 I,;
							classA2 C(16),	tipA2 C(60),	unitsA2 B(2),	entriesA2 I,;
							classA3 C(16),	tipA3 C(60),	unitsA3 B(2),	entriesA3 I,;
							classS C(16),	tipS C(60),		unitsS B(2),	entriesS I,;
							classO C(16),	tipO C(60),		unitsO B(2),	entriesO I,;
							classF C(16),	tipF C(60),		unitsF B(2),	entriesF I,;
							classT C(16),	tipT C(60),		unitsT B(2),	entriesT I,;
							classN C(16),	tipN C(60),		unitsN B(2),	entriesN I,;
							classL C(16),	tipL C(60),		unitsL B(2),	entriesL I,;
							classR C(16),	tipR C(60),		unitsR B(2),	entriesR I,;
							classB C(16),	tipB C(60),		unitsB B(2),	entriesB I,;
							classP C(16),	tipP C(60),		unitsP B(2),	entriesP I,;
							classY C(16),	tipY C(60),		unitsY B(2),	entriesY I,;
							classZ C(16),	tipZ C(60),		unitsZ B(2),	entriesZ I,;
							classD C(16),	tipD C(60),		unitsD B(2),	entriesD I,;
							classR2 C(16),	tipR2 C(60),	unitsR2 B(2),	entriesR2 I,;
							classU C(16),	tipU C(60),	unitsU B(2),	entriesU I )

						IF This.GetEmployeesByGroupCode(pnCurrentGroup, "curStaff")
							* collect all TS records of this pay of selected employees
							SELECT timesheet.tsEmp, timesheet.tsUnits, timesheet.tsType, timesheet.tsCode, timesheet.tsApproved ;
								FROM timesheet ;
								JOIN curStaff ON timesheet.tsEmp = curStaff.myWebCode ;
								WHERE timesheet.tsPay = pnCurrentPay AND !timesheet.tsDownload ;
								INTO CURSOR curPayTimeSheet READWRITE 
							INDEX on tsEmp TAG tsEmp
							INDEX on tsCode TAG tsCode
							INDEX on tsType TAG tsType 

							SELECT curStaff
							SCAN FOR myWebCode >= WEBCODE_PAYROLLUSER_BOUNDARY
								SELECT curSummary
								APPEND BLANK

								REPLACE;
									employee	WITH curStaff.myWebCode;
									empName		WITH curStaff.fullName

								lnAllApproved = 0
								lnNoneApproved = 0
								lnSomeApproved = 0

								llCount = .T.

								FOR lnI = 1 to GS_COL_MAX
									lcPrefix = ""

									DO CASE
										CASE lnI == GS_COL_Timesheet
											lcType = 'M'
											lcDescription = "Timesheet"
											llAuthz = This.CheckRights("TS_TIMESHEET_V")
										CASE lnI == GS_COL_Wage
											lcType = 'W'
											lcDescription = "Wage"
											llAuthz = This.CheckRights("TS_WAGES_V")
										*!* 30/11/2009;TTP4874;JCF: Split the allowances up into more cols...
										CASE INLIST(lnI, GS_COL_Allowances_Amount, GS_COL_Allowances_Rate, GS_COL_Allowances_Units)
											lcType = 'A'
											lcPrefix = 'A' + TRANSFORM(lnI - GS_COL_Allowances_Amount + 1)	&& A1, A2, A3...
											DO CASE
												CASE lnI == GS_COL_Allowances_Amount
													lcDescription = "Amount-Allowances"
												CASE lnI == GS_COL_Allowances_Rate
													lcDescription = "Rate-Allowances"
												CASE lnI == GS_COL_Allowances_Units
													lcDescription = "Units-Allowances"
											ENDCASE
											llAuthz = This.CheckRights("TS_ALLOWANCES_V")
										CASE lnI == GS_COL_SickLeave
											lcType = 'S'
											lcDescription = IIF(plAustralia, "Personal Leave", "Sick Leave")
											llAuthz2 = This.CheckRights("TS_LEAVE_V")
											llAuthz = llAuthz2 AND This.CheckRights("LV_SICK")
											llCount = .T.
										CASE lnI == GS_COL_AnnualLeave
											lcType = 'O'
											lcDescription = "Annual Leave"
											llAuthz = llAuthz2 AND This.CheckRights("LV_ANNUAL")
										CASE lnI == GS_COL_ShiftLeave
											lcType = 'F'
											lcDescription = "Shift Leave"
											llAuthz = llAuthz2 AND This.CheckRights("LV_SHIFT") AND !EMPTY(poStaff.oData.myShDate)
										CASE lnI == GS_COL_OtherLeave
											lcType = 'T'
											lcDescription = "Other Leave"
											llAuthz = llAuthz2 AND This.CheckRights("LV_OTHER") AND !EMPTY(poStaff.oData.myOtDate)
										CASE lnI == GS_COL_LongServiceLeave
											lcType = 'N'
											lcDescription = "Long Service Leave"
											llAuthz = llAuthz2 AND This.CheckRights("LV_LONG")
										CASE lnI == GS_COL_LieuTime
											lcType = 'L'
											lcDescription = "Lieu Time"
											llCount = plAustralia
											llAuthz = llAuthz2 AND This.CheckRights("LV_ALT")
										CASE lnI == GS_COL_RDO
											lcType = 'R'
											lcDescription = "RDO"
											llAuthz = llAuthz2 AND This.CheckRights("LV_RDO")
										CASE lnI == GS_COL_BereavementLeave
											lcType = 'B'
											lcDescription = "Bereavement Leave"
											llCount = !plAustralia
											llAuthz = llAuthz2
										CASE lnI == GS_COL_PublicHoliday
											lcType = 'P'
											lcDescription = "Public Holiday"
										CASE lnI == GS_COL_AltLeaveAccrued
											lcType = 'Y'
											lcDescription = "Alternative Leave Accrued"
											llAuthz = llAuthz2 AND This.CheckRights("LV_ALT")
										CASE lnI == GS_COL_AltLeavePaid
											lcType = 'Z'
											lcDescription = "Alternative Leave Paid"
										CASE lnI == GS_COL_DaysPaid
											lcType = 'D'
											lcDescription = "Days Paid"
											llCount = .T.
											llAuthz = This.CheckRights("TS_OTHER_V")
										CASE lnI == GS_COL_RelevantDaysPaid
											lcPrefix = "R2"
											lcType = 'R'
											lcDescription = "Relevant Days Paid"
											llCount = !plAustralia
										CASE lnI=GS_COL_UnpaidLeave
										    lcPrefix = "U" 
										    lcType = 'U'
											lcDescription = "Unpaid Leave"
											llCount = .T.
									ENDCASE

									IF EMPTY(lcPrefix)
										lcPrefix = lcType
									ENDIF
									
									&&NOTE: the logic here around calc_how must match that in the payroll import code and that in the NewAllowanceObject() calculations around tnType!

									*!* 30/11/2009;TTP4874;JCF: Handle the calculations for the split Allowances cols specially by looking up the set of codes that apply for each col.  If none, that col gets zeros.
									DO CASE
										CASE lnI == GS_COL_Allowances_Amount
											lcCodeFilter = "(INLIST(tsCode"
											lnArgCount = 1
											lcJoiner = ", "
											SELECT allow
											SCAN FOR calc_how == 1
												lcCodeFilter = lcCodeFilter + lcJoiner + TRANSFORM(code)
												lcJoiner = ", "
												lnArgCount = lnArgCount + 1
												IF lnArgCount > 25
													lcJoiner = ") OR INLIST(tsCode, "
													lnArgCount = 1
												ENDIF
											ENDSCAN
											lcCodeFilter = lcCodeFilter + '))'

											IF lcCodeFilter = "(INLIST(tsCode))"
												lnUnits = 0
												lnEntries = 0
												lnUnapprovedEntries = 0
											ELSE
												SELECT curPayTimeSheet 
												COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnEntries
												lnUnits = 0
												lnUnapprovedEntries = 0
												IF lnEntries > 0
													SUM(tsUnits) FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnUnits  
													COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries 
												ENDIF  
*!*													SELECT timesheet
*!*													CALCULATE SUM(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnUnits
*!*													CALCULATE COUNT(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnEntries
*!*													CALCULATE COUNT(tsUnits) FOR tsemp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tstype == lcType AND !tsDownload AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries
											ENDIF

										CASE lnI == GS_COL_Allowances_Rate
											lcCodeFilter = "(INLIST(tsCode"
											lnArgCount = 1
											lcJoiner = ", "
											SELECT allow
											SCAN FOR INLIST(calc_how, 2, 3, 4, 8)
												lcCodeFilter = lcCodeFilter + lcJoiner + TRANSFORM(code)
												lcJoiner = ", "
												lnArgCount = lnArgCount + 1
												IF lnArgCount > 25
													lcJoiner = ") OR INLIST(tsCode, "
													lnArgCount = 1
												ENDIF
											ENDSCAN
											lcCodeFilter = lcCodeFilter + '))'

											IF lcCodeFilter == "(INLIST(tsCode))"
												lnUnits = 0
												lnEntries = 0
												lnUnapprovedEntries = 0
											ELSE
												SELECT curPayTimeSheet 
												COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnEntries
												lnUnits = 0
												lnUnapprovedEntries = 0
												IF lnEntries > 0
													IF lcType = "M"
														SUM(ROUND(tsUnits,2)) FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnUnits  
													ELSE 
														SUM(tsUnits) FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnUnits  
													ENDIF 
													COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries 
												ENDIF  
*!*													SELECT timesheet
*!*													IF lcType == "M"
*!*														CALCULATE SUM(ROUND(tsUnits,2)) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnUnits
*!*													ELSE
*!*														CALCULATE SUM(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnUnits
*!*													ENDIF
*!*													CALCULATE COUNT(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnEntries
*!*													CALCULATE COUNT(tsUnits) FOR tsemp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tstype == lcType AND !tsDownload AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries
											ENDIF

										CASE lnI == GS_COL_Allowances_Units
											lcCodeFilter = "(INLIST(tsCode"
											lnArgCount = 1
											lcJoiner = ", "
											SELECT allow
											SCAN FOR !INLIST(calc_how, 1, 2, 3, 4, 8)
												lcCodeFilter = lcCodeFilter + lcJoiner + TRANSFORM(code)
												lcJoiner = ", "
												lnArgCount = lnArgCount + 1
												IF lnArgCount > 25
													lcJoiner = ") OR INLIST(tsCode, "
													lnArgCount = 1
												ENDIF
											ENDSCAN
											lcCodeFilter = lcCodeFilter + '))'
											
											IF lcCodeFilter = "(INLIST(tsCode))"
												lnUnits = 0
												lnEntries = 0
												lnUnapprovedEntries = 0
											ELSE
												SELECT curPayTimeSheet 
												COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnEntries
												lnUnits = 0
												lnUnapprovedEntries = 0
												IF lnEntries > 0
													IF lcType = "M"
														SUM(ROUND(tsUnits,2)) FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnUnits  
													ELSE 
														SUM(tsUnits) FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. TO lnUnits  
													ENDIF 
													COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries 
												ENDIF  
*!*													IF lcType == "M"
*!*														CALCULATE SUM(ROUND(tsUnits,2)) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnUnits
*!*													ELSE
*!*														CALCULATE SUM(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnUnits
*!*													ENDIF
*!*													CALCULATE COUNT(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload AND &lcCodeFilter. TO lnEntries
*!*													CALCULATE COUNT(tsUnits) FOR tsemp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tstype == lcType AND !tsDownload AND &lcCodeFilter. AND !tsApproved TO lnUnapprovedEntries
											ENDIF

										OTHERWISE
												SELECT curPayTimeSheet 
												COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType TO lnEntries
												lnUnits = 0
												lnUnapprovedEntries = 0
												IF lnEntries > 0
													IF lcType = "M"
														SUM(ROUND(tsUnits,2)) FOR tsEmp = curStaff.myWebCode AND tsType = lcType TO lnUnits  
													ELSE 
														SUM(tsUnits) FOR tsEmp = curStaff.myWebCode AND tsType = lcType TO lnUnits  
													ENDIF 
													COUNT FOR tsEmp = curStaff.myWebCode AND tsType = lcType AND !tsApproved  TO lnUnapprovedEntries 
												ENDIF  
											*!* 30/11/2009;TTP4874;JCF: Do the normal calculations for all other columns
*!*												SELECT timesheet
*!*												IF lcType == "M"
*!*													CALCULATE SUM(ROUND(tsUnits,2)) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload TO lnUnits
*!*												ELSE
*!*													CALCULATE SUM(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload TO lnUnits
*!*												ENDIF
*!*												*!* 23/11/2009;TTP4868;JCF: removed assumption that units are always positive.  Instead, count the actual entries for the purpose of the status.
*!*												CALCULATE COUNT(tsUnits) FOR tsEmp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tsType == lcType AND !tsDownload TO lnEntries

*!*												lnUnapprovedEntries = 0
*!*												IF lnEntries > 0
*!*													CALCULATE COUNT(tsUnits) FOR tsemp == curStaff.myWebCode AND tsPay == pnCurrentPay AND tstype == lcType AND !tsDownload AND !tsApproved TO lnUnapprovedEntries
*!*												ENDIF
									ENDCASE

									DO CASE
										CASE lnEntries == 0
											lcLinkClass = "tsgsEmpty"
											lcToolTip = "No $ entries have been entered."
										CASE lnEntries == lnUnapprovedEntries
											lcLinkClass = "tsgsNoneApproved"
											lcToolTip = "No $ entries have been approved."
											IF llCount AND llAuthz
												lnNoneApproved = lnNoneApproved + 1
											ENDIF
										CASE lnUnapprovedEntries != 0
											lcLinkClass = "tsgsSomeApproved"
											lcToolTip = "Some $ entries have been approved."
											IF llCount AND llAuthz
												lnSomeApproved = lnSomeApproved + 1
											ENDIF
										OTHERWISE
											lcLinkClass = "tsgsAllApproved"
											IF llCount AND llAuthz
												lnAllApproved = lnAllApproved + 1
											ENDIF
											lcToolTip = "All $ entries have been approved."
									ENDCASE
									lcToolTip = STRTRAN(lcToolTip, "$", lcDescription)

									*!* 23/11/2009;TTP4868;JCF: Pass the count of entries as well as the sum total.
									REPLACE;
										("class" + lcPrefix)	WITH lcLinkClass,;
										("tip" + lcPrefix)		WITH lcToolTip,;
										("units" + lcPrefix)	WITH IIF(lcPrefix="M",ROUND(lnUnits,2),ROUND(lnUnits,2)),;
										("entries" + lcPrefix)	WITH lnEntries;
										IN curSummary
								ENDFOR

								DO CASE
									CASE lnNoneApproved > 0
										lcStatus = "Red"
										lcStatusDesc = "No entries approved."
									CASE lnSomeApproved > 0
										lcStatus = "Orange"
										lcStatusDesc = "Some entries approved."
									OTHERWISE
										lcStatus = "Green"
										lcStatusDesc = IIF(lnAllApproved == 0, "No entries.", "All entries approved.")
								ENDCASE

								REPLACE;
									iconImg			WITH lcStatus,;
									iconDesc		WITH lcStatusDesc;
									canApprove		WITH (lnNoneApproved + lnSomeApproved > 0);
									canUnapprove	WITH (lnAllApproved + lnSomeApproved > 0);
									IN curSummary
							ENDSCAN

							USE IN SELECT("curPayTimeSheet")
							*!* 30/11/2009;TTP4874;JCF: Add the extra col totals too..
							SELECT;
								SUM(unitsP) AS totalP, SUM(unitsM) AS totalM, SUM(unitsW) AS totalW, SUM(unitsA1) AS totalA1, SUM(unitsA2) AS totalA2, SUM(unitsA3) AS totalA3,;
								SUM(unitsS) AS totalS, SUM(unitsO) AS totalO, SUM(unitsF) AS totalF, SUM(unitsT) AS totalT, SUM(unitsN) AS totalN, SUM(unitsL) AS totalL,;
								SUM(unitsR) AS totalR, SUM(unitsB) AS totalB, SUM(unitsY) AS totalY, SUM(unitsZ) AS totalZ, SUM(unitsD) AS totalD, SUM(unitsR2) AS totalR2,;
								SUM(unitsU) AS totalU ;
								FROM curSummary;
								INTO CURSOR curSummaryTotals
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		rcGroupInfo = "Unknown Group"
		rcPayInfo = " "
		rcPayName = "Unknown Pay"
		rcStaffInfo = "Unknown Manager"
		rcApprover = "Unknown Manager"
		rdPayDate = { / / }
		rcCompName = "Company Name"

		IF pnCurrentPay > 0
			IF SEEK(pnCurrentPay, "myPays", "pay_pk")
				rcPayInfo = ALLTRIM(mypays.pay_name)
				rcPayName = "Starting " + CDOW(myPays.pay_date) + ", " + DMY(myPays.pay_date)
				rdPayDate = myPays.pay_date
			ENDIF
		ENDIF
				
		IF pnCurrentGroup > 0
			IF SEEK(pnCurrentGroup,"myGroups","grcode")
				rcGroupInfo = myGroups.grname
			ENDIF
		ENDIF

*!*			IF pnCurrentStaff > 0
*!*				IF SEEK(pnCurrentStaff,"myStaff","mywebcode")
*!*					rcStaffInfo = ALLTRIM(myStaff.mySurname)+","+ALLTRIM(myStaff.myname)
*!*					rcApprover = ALLTRIM(myStaff.myname)+" "+ALLTRIM(myStaff.mySurname)
*!*				ENDIF
*!*			ENDIF

       **JA 14/11/2012 replaced the above code with below to print correct manager name
		IF this.employee > 0
			IF SEEK(this.employee,"myStaff","mywebcode")
				rcStaffInfo = ALLTRIM(myStaff.mySurname)+","+ALLTRIM(myStaff.myname)
				rcApprover = ALLTRIM(myStaff.myname)+" "+ALLTRIM(myStaff.mySurname)
			ENDIF
		ENDIF
		**JA 14/11/2012
		
		IF tlPrint
			LOCAL lcFileName, lcPathName, lcOutputFile as String
			lcPathName = This.CompanyDataPath() + ADDBS("reports")
			IF NOT DIRECTORY(lcPathName)
				TRY
					MKDIR (lcPathName)
				CATCH
				ENDTRY
			ENDIF
					
			IF DIRECTORY(lcPathName)
				lcFileName = "GROUP_SUMMARY"
        			lcOutputFile = lcPathName + TRANSFORM(This.Employee) + "_" + lcFileName + ".CSV"
				*lcURL = STREXTRACT(UPPER(Request.GetCurrentUrl(.F.)),This.CompanyHTTPS(),"/GROUP")
				*lcURLCsv = This.CompanyHTTPS()+LOWER(ALLTRIM(lcURL))+"/GetCSV.si?date=" + lcFileName
				lcURLCsv = "http" + LOWER(STREXTRACT(UPPER(Request.GetCurrentUrl(.F.)),"HTTP","/GROUP")) + "/GetCSV.si?date=" + lcFileName

*!*					CREATE CURSOR curSummOut (Status C(21),Employee_Code I,Employee_Name C(100), Time B(2),;
*!*												Wages B(2),Allowance_Amount B(2),Allowance_Rate B(2),Allowance_Units B(2),;
*!*												Sick_Leave B(2),Annual_Leave B(2),Shift_Leave B(2),Other_Leave B(2),;
*!*												Long_Service_Leave B(2),Rostered_Day_Off B(2),Bereavement_Leave B(2),;
*!*												Public_Holiday_Leave B(2),Alternative_Leave_Accrued B(2),;
*!*												Alternative_Leave_Paid B(2),Hours_or_Days_Paid B(2),;
*!*												Relevant_Hours_or_Days_Paid B(2))
** JA 23/10/2012, include personal leave and lieu time 
				CREATE CURSOR curSummOut (Status C(21),Employee_Code I,Employee_Name C(100), Time B(2),;
											Wages B(2),Allowance_Amount B(2),Allowance_Rate B(2),Allowance_Units B(2),;
											Sick_Leave B(2), Personal_Leave B(2), Annual_Leave B(2),Shift_Leave B(2),Other_Leave B(2),;
											Long_Service_Leave B(2), Lieu_Hours B(2), Rostered_Day_Off B(2),Bereavement_Leave B(2),;
											Public_Holiday_Leave B(2),Alternative_Leave_Accrued B(2),;
											Alternative_Leave_Paid B(2),Hours_or_Days_Paid B(2),;
											Relevant_Hours_or_Days_Paid B(2), Unpaid_Hours B(2))																				
				SELECT curSummary
				GO top
				DO WHILE NOT EOF('curSummary')
					SELECT curSummOut
					APPEND BLANK

					nEmpCode = -1
					IF curSummary.employee > 0
						IF SEEK(curSummary.employee,"myStaff","mywebcode")
							nEmpCode = myStaff.myPayCode
						ENDIF
					ENDIF

*!*						replace	Status WITH CurSummary.icondesc,;
*!*								Employee_Code WITH nEmpCode,;
*!*								Employee_Name WITH curSummary.empname,;
*!*								Time WITH curSummary.unitsm,;
*!*								Wages WITH curSummary.unitsw,;
*!*								Allowance_Amount WITH curSummary.unitsa1,;
*!*								Allowance_Rate WITH curSummary.unitsa2,;
*!*								Allowance_Units WITH curSummary.unitsa3,;
*!*								Sick_Leave WITH curSummary.unitss,;
*!*								Annual_Leave WITH curSummary.unitso,;
*!*								Shift_Leave WITH curSummary.unitsf,;
*!*								Other_Leave WITH curSummary.unitst,;
*!*								Long_Service_Leave WITH curSummary.unitsn,;
*!*								Rostered_Day_Off WITH curSummary.unitsl,;
*!*								Bereavement_Leave WITH curSummary.unitsr,;
*!*								Public_Holiday_Leave WITH curSummary.unitsb,;
*!*								Alternative_Leave_Accrued WITH curSummary.unitsp,;
*!*								Alternative_Leave_Paid WITH curSummary.unitsy,;
*!*								Hours_or_Days_Paid WITH curSummary.unitsz,;
*!*								Relevant_Hours_or_Days_Paid WITH curSummary.unitsd
					* JAI & MY 19/10/2012 (US 9121, 9122) - column mixed up
					replace	Status WITH CurSummary.icondesc,;
							Employee_Code WITH nEmpCode,;
							Employee_Name WITH curSummary.empname,;
							Time WITH curSummary.unitsM,;
							Wages WITH curSummary.unitsW,;
							Allowance_Amount WITH curSummary.unitsa1,;
							Allowance_Rate WITH curSummary.unitsa2,;
							Allowance_Units WITH curSummary.unitsa3,;
							Sick_Leave WITH curSummary.unitsS,;
							Personal_Leave WITH curSummary.unitsS,; && JA code 23/10/2012
							Annual_Leave WITH curSummary.unitsO,;
							Shift_Leave WITH curSummary.unitsF,;
							Other_Leave WITH curSummary.unitsT,;
							Long_Service_Leave WITH curSummary.unitsN,;
							Lieu_Hours WITH curSummary.unitsL,;
							Rostered_Day_Off WITH curSummary.unitsR,;
							Bereavement_Leave WITH curSummary.unitsB,;
							Public_Holiday_Leave WITH curSummary.unitsP,;
							Alternative_Leave_Accrued WITH curSummary.unitsY,;
							Alternative_Leave_Paid WITH curSummary.unitsZ,;
							Hours_or_Days_Paid WITH curSummary.unitsD,;
							Relevant_Hours_or_Days_Paid WITH curSummary.unitsR, ;
							Unpaid_Hours WITH curSummary.unitsU
							
					SKIP IN 'curSummary'
				ENDDO
											
				SELECT curSummOut
				GO top
				IF NOT EOF()
	            	IF FILE(lcOutputFile)
    	        			DELETE FILE (lcOutputFile)
	    	       	ENDIF
					* Jai & MY 19/10/2012 (US 9121, 9122) relative columns for diff country
					IF plAustralia
						COPY TO (lcOutputFile) FIELDS EXCEPT Sick_Leave, Shift_Leave, Other_Leave, Alternative_Leave_Accrued, Alternative_Leave_Paid, Bereavement_Leave, Relevant_Hours_or_Days_Paid, Public_Holiday_Leave TYPE CSV
					ELSE
						COPY TO (lcOutputFile) FIELDS EXCEPT Rostered_Day_Off, Personal_Leave, Lieu_Hours TYPE CSV
					ENDIF 
	
					*This.AddUserInfo("CSV File Was Created")
					Response.Downloadfile(lcOutputfile,"application/csv","GroupSummary.csv")
					RETURN
					
				ENDIF
				
			ENDIF
		ENDIF

		poPage = This.NewPageObject("timesheets:summary", "time_groupSummary")
		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)

	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE TimeEntryPage(tlIsHistory)
			
        cCSVFile = ""
		this.TimeEntryProc(tlIsHistory,.F.)
	ENDPROC
				
	*--------------------------------------------------------------------------------*

	PROCEDURE TimeSheetCSV(tlIsHistory)
		this.TimeEntryProc(tlIsHistory,.T.)
	ENDPROC


    **JA code 1/11/2012, introduce separate procedure for History CSV
	PROCEDURE TimeHistoryCSV
		this.TimeEntryProc(.T.,.T.)
	ENDPROC

	*--------------------------------------------------------------------------------*

	&&NOTE: this page handles posts to support the buildup of formRows...
	PROCEDURE TimeEntryProc(tlIsHistory,tlCSV)
		PRIVATE poPage, poStaff, poTypes
		PRIVATE pnCurrentPay, pnCurrentStaff, pnCurrentGroup, pnCurrentTemplate, poRetainList, plManager
		PRIVATE pnEditId, pcShowDownloaded, pdStartDate, pdEndDate, pcApproved, plAusie, plUnitDisp
		PRIVATE pcAddType, pnAddCount, pcOpen, pnPayCount, plPayOpen, pcFocusField, pnPreviousPay
		PRIVATE poLeaveCodes, poOtherCodes, poAllowCodes, poWageCodes, poCostCentres, poJobCodes
		
		LOCAL lcApprovedFilter, lcRangeFilter, lcDownloadedFilter, lcAction, loType, lnFieldCount, lnI, lnJ, lnCount, lnAt
		LOCAL ltStart, ltEnd, ltBreak, lnBreakLen, lnUnits, lcUnitDisp
		LOCAL ARRAY laFields[13], laDefaults[13]
		LOCAL lDispFileName
		

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "myPays");
		  AND This.SelectData(This.Licence, "myGroups");
		  AND This.SelectData(This.Licence, "wageType");
		  AND This.SelectData(This.Licence, "allow");
		  AND This.SelectData(This.Licence, "costCent");
		  AND This.SelectData(This.Licence, "timesheet"))	&&TODO: add new table for job-whateverItIs
			This.AddError("Page Setup Failed!")
		ELSE
			poStaff = Factory.GetStaffObject()

			IF !poStaff.Load(This.Employee)
				This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
			ELSE
				poRetainList = Factory.GetRetainListObject()
			
				plManager = .F.
				pnCurrentGroup = 0
				pnCurrentStaff = 0
				pnPreviousPay = 0

				rcGroupInfo = "< My Details >"
				rcPayInfo = "Unknown Pay"
				rcPayName = " "
				rcStaffInfo = "Unknown Manager"
				rcApprover = "Unknown Manager"
				rdPayDate = { / / }
				rcCompName = "Company Name"

				IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
					This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
				ENDIF
	
			 	pnCurrentTemplate = -1
				nTemplateCnt = This.SetupTemplateControlData(@pnCurrentTemplate, poRetainList)
*					This.AddValidationError("TemplateControl Setup Failed!")	&& non-fatal error

				IF !This.CheckAccess(pnCurrentStaff, plManager, .T.)
					This.AddError("You do not have access to this page.")
				ELSE
					pcOpen = EVL(EVL(Request.Form("open"), Request.QueryString("open")), "open")
					pnCurrentPay = 0
					plPayOpen = .F.
					pnPayCount = This.SetupPayControlData(@pnCurrentPay, @plPayOpen, poRetainList, IIF(tlIsHistory, pcOpen, "open"))
					IF pnPayCount < 0
						This.AddError("PayControl Setup Failed!")
					ELSE
						pnPreviousPay = EVL(VAL(Request.QueryString("previousPay")),-1)
						IF pnPayCount == 0
							pnCurrentPay = -1
						ENDIF

						pcFocusField = ""
						pcAddType = ""
						plAusie = This.IsAustralia()
						poLeaveCodes = This.GetLeaveCodes(IIF(pnCurrentStaff == EVERYONE_OPTION, This.Employee, pnCurrentStaff))
						poOtherCodes = This.GetOtherCodes()
						poAllowCodes = This.GetAllowanceCodes()
						poWageCodes = This.GetWageCodes()
						poCostCentres = This.GetCostCentres()
						&&LATER: poJobCodes = This.GetJobCodes()

						pnEditId = VAL(Request.QueryString("edit"))
						SELECT timesheet
						LOCATE FOR tsId = pnEditId
						IF !FOUND()
							pnEditId = 0
						ELSE
							IF !(plManager OR pnCurrentStaff == This.Employee) OR tsDownload
								This.AddError("You do not have access to edit that entry!")
								pnEditId = 0
							ELSE
								IF tsApproved
									This.AddValidationError("You do not have permission to edit that entry!")
									pnEditId = 0
								ENDIF
							ENDIF
						ENDIF

						lcUnitDisp = AppSettings.Get("unitdisp")
						plUnitDisp = .F.
						IF ALLTRIM(lcUnitDisp) = "Decimal"
							plUnitDisp = .T.
					 	ENDIF

						pcAddType = ""
						pnAddCount = 0

						IF !tlIsHistory
							pcShowDownloaded = "no"
							pdStartDate = {}
							pdEndDate = {}
						ELSE
							pcShowDownloaded = Request.QueryString("downloaded")
							IF !INLIST(pcShowDownloaded, "yes", "no", "both")
								pcShowDownloaded = "no"
							ENDIF

							pdStartDate = CTOD(Request.QueryString("startDate"))
							pdEndDate = CTOD(Request.QueryString("endDate"))
						ENDIF

						pcApproved = EVL(Request.Form("approved"), Request.QueryString("approved"))

						DO CASE
							CASE pcApproved == "yes"
								lcApprovedFilter = "AND !ISNULL(tsApproved) AND tsApproved"
							CASE pcApproved == "no"
								lcApprovedFilter = "AND (ISNULL(tsApproved) OR !tsApproved)"
							OTHERWISE
								lcApprovedFilter = ""
								pcApproved = "both"
						ENDCASE

						DO CASE
							CASE pcShowDownloaded == "yes"
								lcDownloadedFilter = "AND !ISNULL(tsDownload) AND tsDownload"
							CASE pcShowDownloaded == "no"
								lcDownloadedFilter = "AND (ISNULL(tsDownload) OR !tsDownload)"
							OTHERWISE
								lcDownloadedFilter = ""
								pcShowDownloaded = "both"
						ENDCASE

						IF !EMPTY(pdStartDate)
							IF !EMPTY(pdEndDate)
								lcRangeFilter = "AND BETWEEN(tsDate, pdStartDate, pdEndDate)"
							ELSE
								lcRangeFilter = "AND tsDate >= pdStartDate"
							ENDIF
						ELSE
							IF !EMPTY(pdEndDate)
								lcRangeFilter = "AND tsDate <= pdEndDate"
							ELSE
								lcRangeFilter = ""
							ENDIF
						ENDIF

						poTypes = This.GetTimesheetTypes(.T., pnCurrentPay, pnCurrentGroup, pnCurrentStaff, lcApprovedFilter + ' ' + lcDownloadedFilter + ' ' + lcRangeFilter)

						IF !tlIsHistory
							* 02/11/2009;TTP4688;JCF: Added handling for pcFocusField so that the first editable field in the edited line gets focus on page load.
							IF !EMPTY(pnEditId)
								FOR lnI = 1 TO poTypes.Count
									loType = poTypes.Item(lnI)

									SELECT (loType.cursorName)
									LOCATE FOR tsId == pnEditId
									IF FOUND()
										EXIT
									ENDIF
								NEXT

								IF plManager
									IF EMPTY(pcFocusField) AND pnCurrentStaff == EVERYONE_OPTION
										pcFocusField = "staff"
									ENDIF
								ENDIF
								IF loType.showDate
									IF EMPTY(pcFocusField) AND !loType.readOnlyDate
										pcFocusField = "date"
									ENDIF
								ENDIF
								IF loType.showLeaveType
									IF EMPTY(pcFocusField) AND !loType.readOnlyLeaveType
										pcFocusField = "leaveType"
									ENDIF
								ENDIF
								IF loType.showOtherType
									IF EMPTY(pcFocusField) AND !loType.readOnlyOtherType
										pcFocusField = "otherType"
									ENDIF
								ENDIF
								IF loType.showCode
									IF EMPTY(pcFocusField) AND !loType.readOnlyCode
										pcFocusField = "code"
									ENDIF
								ENDIF
								IF loType.showStart
									IF EMPTY(pcFocusField) AND !loType.readOnlyStart
										pcFocusField = "start"
									ENDIF
								ENDIF
								IF loType.showEnd
									IF EMPTY(pcFocusField) AND !loType.readOnlyEnd
										pcFocusField = "end"
									ENDIF
								ENDIF
								IF loType.showBreak
									IF EMPTY(pcFocusField) AND !loType.readOnlyBreak
										pcFocusField = "break"
									ENDIF
								ENDIF
								IF loType.showUnits
									IF EMPTY(pcFocusField) AND !loType.readOnlyUnits
										pcFocusField = "units"
									ENDIF
								ENDIF
								IF loType.showReduce
									IF EMPTY(pcFocusField) AND !loType.readOnlyReduce
										pcFocusField = "units2"
									ENDIF
								ENDIF
								IF loType.showWageType
									IF EMPTY(pcFocusField) AND !loType.readOnlyWageType
										pcFocusField = "wageType"
									ENDIF
								ENDIF
								IF loType.showRateCode
									IF EMPTY(pcFocusField) AND !loType.readOnlyRateCode
										pcFocusField = "rateCode"
									ENDIF
								ENDIF
								IF loType.showCostCent
									IF EMPTY(pcFocusField) AND !loType.readOnlyCostCent
										pcFocusField = "costCentre"
									ENDIF
								ENDIF
								IF loType.showJobCode
									IF EMPTY(pcFocusField) AND !loType.readOnlyJobCode
										pcFocusField = "jobCode"
									ENDIF
								ENDIF
							ELSE
								pcAddType = Request.Form("addType")
								pnAddCount = VAL(Request.Form("count"))

								IF EMPTY(poTypes.GetKey(pcAddType))
									pcAddType = ""
									pnAddCount = 0
								ELSE
									loType = poTypes.Item(pcAddType)
									lnFieldCount = 0

									poAddValues = CREATEOBJECT("COLLECTION")

									lcAction = Request.Form("mode")

									* Don't soak up current values if changing type.
									IF !(LEFT(lcAction, 7) == "newType")
										FOR lnI = 1 TO pnAddCount
											IF plManager
												poAddValues.Add(Request.Form("staff_" + TRANSFORM(lnI)), "staff_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showDate
												poAddValues.Add(Request.Form("date_" + TRANSFORM(lnI)), "date_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showLeaveType
												poAddValues.Add(Request.Form("leaveType_" + TRANSFORM(lnI)), "leaveType_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showOtherType
												poAddValues.Add(Request.Form("otherType_" + TRANSFORM(lnI)), "otherType_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showCode
												poAddValues.Add(Request.Form("code_" + TRANSFORM(lnI)), "code_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showStart
												poAddValues.Add(Request.Form("start_" + TRANSFORM(lnI)), "start_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showEnd
												poAddValues.Add(Request.Form("end_" + TRANSFORM(lnI)), "end_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showBreak
												poAddValues.Add(Request.Form("break_" + TRANSFORM(lnI)), "break_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showUnits
												IF loType.showStart AND loType.showEnd AND loType.showBreak
													* 16/10/2009;TTP:4643;JCF: Timesheet type needs the units recalculated as the [shouldn't] be passed back by the browser since the field is disabled.
													ltStart	= CTOT(poAddValues.Item("start_" + TRANSFORM(lnI)))
													ltEnd	= CTOT(poAddValues.Item("end_" + TRANSFORM(lnI)))
													ltBreak	= CTOT(poAddValues.Item("break_" + TRANSFORM(lnI)))

													lnBreakLen = ltBreak - CTOT("00:00")

													IF ltEnd < ltStart
														* Crossed midnight
														ltEnd = ltEnd + 86400
													ENDIF

													lnUnits = ((ltEnd - ltStart) - lnBreakLen) / 3600
													IF lnUnits < 0 OR lnUnits > 24
														* This probably never fires unless they disable JavaScript...
														This.AddValidationError("Invalid Start/End/Break combination.")
													ENDIF

													poAddValues.Add(TRANSFORM(lnUnits), "units_" + TRANSFORM(lnI))
												ELSE
													poAddValues.Add(Request.Form("units_" + TRANSFORM(lnI)), "units_" + TRANSFORM(lnI))
												ENDIF
											ENDIF
											IF loType.showReduce
												poAddValues.Add(Request.Form("units2_" + TRANSFORM(lnI)), "units2_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showWageType
												poAddValues.Add(Request.Form("wageType_" + TRANSFORM(lnI)), "wageType_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showRateCode
												poAddValues.Add(Request.Form("rateCode_" + TRANSFORM(lnI)), "rateCode_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showCostCent
												poAddValues.Add(Request.Form("costCentre_" + TRANSFORM(lnI)), "costCentre_" + TRANSFORM(lnI))
											ENDIF
											IF loType.showJobCode
												poAddValues.Add(Request.Form("jobCode_" + TRANSFORM(lnI)), "jobCode_" + TRANSFORM(lnI))
											ENDIF
										NEXT
									ENDIF

									* 02/11/2009;TTP4688;JCF: Added handling for pcFocusField so that the first editable field in the "current" line gets focus on page load.
									IF plManager
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "staff_"
										laDefaults[lnFieldCount] = TRANSFORM(pnCurrentStaff)

										IF EMPTY(pcFocusField) AND pnCurrentStaff == EVERYONE_OPTION
											pcFocusField = "staff_"
										ENDIF
									ENDIF
									IF loType.showDate
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "date_"
										laDefaults[lnFieldCount] = DATE()

										IF EMPTY(pcFocusField) AND !loType.readOnlyDate
											pcFocusField = "date_"
										ENDIF
									ENDIF
									IF loType.showLeaveType
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "leaveType_"
										laDefaults[lnFieldCount] = ""

										IF EMPTY(pcFocusField) AND !loType.readOnlyLeaveType
											pcFocusField = "leaveType_"
										ENDIF
									ENDIF
									IF loType.showOtherType
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "otherType_"
										laDefaults[lnFieldCount] = ""

										IF EMPTY(pcFocusField) AND !loType.readOnlyOtherType
											pcFocusField = "otherType_"
										ENDIF
									ENDIF
									IF loType.showCode
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "code_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyCode
											pcFocusField = "code_"
										ENDIF
									ENDIF
									IF loType.showStart
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "start_"
										laDefaults[lnFieldCount] = "00:00"

										IF EMPTY(pcFocusField) AND !loType.readOnlyStart
											pcFocusField = "start_"
										ENDIF
									ENDIF
									IF loType.showEnd
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "end_"
										laDefaults[lnFieldCount] = "00:00"

										IF EMPTY(pcFocusField) AND !loType.readOnlyEnd
											pcFocusField = "end_"
										ENDIF
									ENDIF
									IF loType.showBreak
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "break_"
										laDefaults[lnFieldCount] = "00:00"

										IF EMPTY(pcFocusField) AND !loType.readOnlyBreak
											pcFocusField = "break_"
										ENDIF
									ENDIF
									IF loType.showUnits
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "units_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyUnits
											pcFocusField = "units_"
										ENDIF
									ENDIF
									IF loType.showReduce
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "units2_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyReduce
											pcFocusField = "units2_"
										ENDIF
									ENDIF
									IF loType.showWageType
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "wageType_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyWageType
											pcFocusField = "wageType_"
										ENDIF
									ENDIF
									IF loType.showRateCode
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "rateCode_"
										laDefaults[lnFieldCount] = '1'

										IF EMPTY(pcFocusField) AND !loType.readOnlyRateCode
											pcFocusField = "rateCode_"
										ENDIF
									ENDIF
									IF loType.showCostCent
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "costCentre_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyCostCent
											pcFocusField = "costCentre_"
										ENDIF
									ENDIF
									IF loType.showJobCode
										lnFieldCount = lnFieldCount + 1
										laFields[lnFieldCount] = "jobCode_"
										laDefaults[lnFieldCount] = '0'

										IF EMPTY(pcFocusField) AND !loType.readOnlyJobCode
											pcFocusField = "jobCode_"
										ENDIF
									ENDIF

									* if asked to, replicate a row in the temp table with consecutive dates, or add more rows to it, or explode an Everyone row etc...
									* 02/11/2009;TTP4688;JCF: Added logic to define what the "current" line means in each case below for pcFocusField.
									DO CASE
										CASE LEFT(lcAction, 7) == "newType"		&& may be newType or newType_<count>
											* add N rows of the new type
											IF '_' $ lcAction
												lnCount = VAL(SUBSTR(lcAction, 9))
											ELSE
												pnAddCount = 0
												lnCount = 1
											ENDIF

											FOR lnI = pnAddCount + 1 TO pnAddCount + lnCount
												FOR lnJ = 1 TO lnFieldCount
													IF laFields[lnJ] == "date_"
														poAddValues.Add(TRANSFORM(laDefaults[lnJ] + (lnI - pnAddCount - 1)), laFields[lnJ] + TRANSFORM(lnI))
													ELSE
														poAddValues.Add(laDefaults[lnJ], laFields[lnJ] + TRANSFORM(lnI))
													ENDIF
												NEXT
											NEXT

											* Focus on the first new row
											pcFocusField = pcFocusField + TRANSFORM(pnAddCount + 1)

											pnAddCount = pnAddCount + lnCount

										CASE LEFT(lcAction, 4) == "add_"
											* add N rows of the current type
											IF pnAddCount > 0
												* start the dates from just after the last one if present...
												lnCount = ASCAN(laFields, "date_")
												IF lnCount > 0
													laDefaults[lnCount] = CTOD(poAddValues.Item("date_" + TRANSFORM(pnAddCount))) + 1
												ENDIF
											ENDIF

											lnCount = VAL(SUBSTR(lcAction, 5))

											FOR lnI = pnAddCount + 1 TO pnAddCount + lnCount
												FOR lnJ = 1 TO lnFieldCount
													IF laFields[lnJ] == "date_"
														poAddValues.Add(TRANSFORM(laDefaults[lnJ] + (lnI - pnAddCount - 1)), laFields[lnJ] + TRANSFORM(lnI))
													ELSE
														poAddValues.Add(laDefaults[lnJ], laFields[lnJ] + TRANSFORM(lnI))
													ENDIF
												NEXT
											NEXT

											* Focus on first new row
											pcFocusField = pcFocusField + TRANSFORM(pnAddCount + 1)

											pnAddCount = pnAddCount + lnCount

										CASE LEFT(lcAction, 7) == "delete_"
											* delete the Nth row; bubble up other values
											lnCount = VAL(SUBSTR(lcAction, 8))

											pnAddCount = pnAddCount - 1

											FOR lnI = lnCount TO pnAddCount
												FOR lnJ = 1 TO lnFieldCount
													poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
													poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnI + 1)), laFields[lnJ] + TRANSFORM(lnI))
												NEXT
											NEXT

											* Focus on either the row below the one deleted if there is one, or the last row
											pcFocusField = pcFocusField + TRANSFORM(MIN(pnAddCount, lnCount))

										CASE LEFT(lcAction, 5) == "copy_"		&& copy_<at>_<count>
											* copy the Nth row, inserting at N+1 and pushing down.
											lnAt = VAL(SUBSTR(lcAction, 6))
											lnCount = VAL(SUBSTR(lcAction, 8 + FLOOR(LOG10(lnAt))))

											FOR lnI = pnAddCount + lnCount TO lnAt + lnCount + 1 STEP -1
												FOR lnJ = 1 TO lnFieldCount
													IF lnI <= pnAddCount
														poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
													ENDIF

													poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnI - lnCount)), laFields[lnJ] + TRANSFORM(lnI))
												NEXT
											NEXT

											FOR lnI = lnAt + 1 TO lnAt + lnCount
												FOR lnJ = 1 TO lnFieldCount
													IF lnI <= pnAddCount
														poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
													ENDIF

													IF laFields[lnJ] == "date_"
														poAddValues.Add(TRANSFORM(CTOD(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnAt))) + lnI - lnAt), laFields[lnJ] + TRANSFORM(lnI))
													ELSE
														poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnAt)), laFields[lnJ] + TRANSFORM(lnI))
													ENDIF
												NEXT
											NEXT

											* Focus on the first new row
											pcFocusField = pcFocusField + TRANSFORM(lnAt + 1)

											pnAddCount = pnAddCount + lnCount

										CASE LEFT(lcAction, 8) == "explode_"
											* copy the current row that is for Everyone once for each member of the current group

											lnAt = VAL(SUBSTR(lcAction, 9))
											IF poAddValues.Item("staff_" + TRANSFORM(lnAt)) != TRANSFORM(EVERYONE_OPTION)
												This.AddError("Can only explode a row set to " + EVERYONE_LABEL + ".")
											ELSE
												IF !This.GetEmployeesByGroupCode(pnCurrentGroup, "curGroupStaff")
													This.AddError("Cannot load current group!")
												ELSE
													loGroup = CREATEOBJECT("COLLECTION")
													lnCount = 0
													SELECT curGroupStaff
													SCAN
														IF curGroupStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curGroupStaff.myWebCode < 2	&&NOTE: always hiding PayrollUsers here; hide Admin user either way...
															LOOP
														ENDIF

														lnCount = lnCount + 1
														loGroup.Add(TRANSFORM(curGroupStaff.myWebCode), TRANSFORM(lnCount))
													ENDSCAN

													* set the staff field on the copied row to the first group member.
													poAddValues.Remove("staff_" + TRANSFORM(lnAt))
													poAddValues.Add(loGroup.Item(1), "staff_" + TRANSFORM(lnAt))

													* copy the Nth row, inserting at N+1 and pushing down.
													lnCount = lnCount - 1

													IF lnCount > 0
														FOR lnI = pnAddCount + lnCount TO lnAt + lnCount + 1 STEP -1
															FOR lnJ = 1 TO lnFieldCount
																IF lnI <= pnAddCount
																	poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
																ENDIF

																poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnI - lnCount)), laFields[lnJ] + TRANSFORM(lnI))
															NEXT
														NEXT

														FOR lnI = lnAt + 1 TO lnAt + lnCount
															FOR lnJ = 1 TO lnFieldCount
																IF lnI <= pnAddCount
																	poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
																ENDIF

																IF laFields[lnJ] == "staff_"
																	poAddValues.Add(loGroup.Item(lnI - lnAt + 1), laFields[lnJ] + TRANSFORM(lnI))
																ELSE
																	poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnAt)), laFields[lnJ] + TRANSFORM(lnI))
																ENDIF
															NEXT
														NEXT
													ENDIF

													* Focus on the row that was exploded - now the first new row
													pcFocusField = pcFocusField + TRANSFORM(lnAt)

													pnAddCount = pnAddCount + lnCount
												ENDIF
											ENDIF

										*!* 09/11/2009;TTP4555;JCF: handle the case when the user has pressed enter and submits the form with no specific action.
										CASE lcAction == ""	&& Happens if the AddEntries form is submitted by the user pressing the Enter key...need to fix the focused field in this case.
											pcFocusField = pcFocusField + '1'
									ENDCASE
								ENDIF
							ENDIF
						ENDIF

						rcGroupInfo = "< My Details >"
						rcPayInfo = "Unknown Pay"
						rcPayName = " "
						rcStaffInfo = "Unknown Manager"
						rcApprover = "Unknown Manager"
						rdPayDate = { / / }
						rcCompName = "Company Name"

						IF pnCurrentPay > 0
							IF SEEK(pnCurrentPay, "myPays", "pay_pk")
								rcPayInfo = ALLTRIM(mypays.pay_name)
								rcPayName = "Starting " + CDOW(myPays.pay_date) + ", " + DMY(myPays.pay_date)
								rdPayDate = myPays.pay_date
							ENDIF
						ENDIF
				
						IF pnCurrentGroup > 0
							IF SEEK(pnCurrentGroup,"myGroups","grcode")
								rcGroupInfo = myGroups.grname
							ENDIF
						ENDIF

						IF this.employee > 0
							IF SEEK(this.employee,"myStaff","mywebcode")
								rcStaffInfo = ALLTRIM(myStaff.mySurname)+","+ALLTRIM(myStaff.myname)
								rcApprover = ALLTRIM(myStaff.myname)+" "+ALLTRIM(myStaff.mySurname)
							ENDIF
						ENDIF

					ENDIF
				ENDIF
			ENDIF

*			WAIT WINDOW IIF(VARTYPE(lcAction)<>"C","lcAction = None","lcAction = "+lcAction)+CHR(13)+;
	    	        " pnCurrentGroup = "+ALLTRIM(STR(pnCurrentGroup,10,0))+CHR(13)+;
	        	    " pnCurrentStaff = "+ALLTRIM(STR(pnCurrentStaff,10,0))+CHR(13)+;
					" this.Employee = "+IIF(VARTYPE(this.Employee)="N",ALLTRIM(STR(this.employee,10,0)),"**")+CHR(13)+;
		            " Manager ["+IIF(plManager,"Yes","No")+"]"+CHR(13)+;
	        	    " pnPreviousPay = "+ALLTRIM(STR(pnPreviousPay,10,0))+CHR(13)+;
		            " pnAddCount = "+ALLTRIM(STR(pnAddCount,10,0))+CHR(13)+;
		            " pnEditId = "+ALLTRIM(STR(pnEditID,10,0))+' '+CHR(13)+;
		            " plUnitDisp = ["+IIF(plUnitDisp,"Yes","No")+"] "+CHR(13)+;
		            TIME() ;
		             NOWAIT noclear
*
		ENDIF

		IF tlCSV
			* JAI & MY  19/10/2012 (US 9134) export to csv with the correct columns ---------------------------
*!*				CREATE CURSOR TimeSheetOut (Date D,Start C(8), End C(8), Employee_Code I, First_Names C(50),Surname C(50),;
*!*											TYPE C(50),Units B(2),Cost_Centre_Code N(12,0),Cost_Centre_Name C(50),;
*!*											Allowance_Code I(4,0),Allowance_Name C(50),Wage_Type I(4,0),Wage_Name C(50),;
*!*											Rate_Code I(4,0), Break C(8),Units_To_Reduce B(2), Status C(15),Pay C(50))
			CREATE CURSOR TimeSheetOut ( ;
				Pay C(50), ;
				Group C(50), ;
				Employee_Code I, ;
				Employee_FirstNames C(50), ;
				Employee_LastName C(50),;
				Type C(50), ;
				Status C(15), ;
				Date D, ;
				Start C(8), ;
				End C(8), ;
				Break C(8), ;
				Units B(2), ;
				sub_type_Code I, ;
				sub_type_Name C(50), ;
				Cost_Centre_Code I, ;
				Cost_Centre_Name C(50),;
				Units_To_Reduce B(2),;
				tsDate D,; &&     JA 01/11/2012
				tsDownload L,; && JA 01/11/2012
				tsApproved L,; && JA 01/11/2012
				IsDownloaded C(3)) && JA 01/11/2012
			*---------------------------------------------------------------------------------------------------------------
			lnFld = AFIELDS(laFld,"timesheet")
			DIMENSION laFld(lnFld+1,18)
			FOR nFldCntr = 1 TO 18
				laFld(lnFld+1,nFldCntr) = laFld(lnFld,nFldCntr)
			ENDFOR

			laFld(lnFld+1,1) = "TSTYPECODE"
			laFld(lnFld+1,2) = "C"
			laFld(lnFld+1,3) = 1
			laFld(lnFld+1,4) = 0
			m.tsType = "U"
			
			CREATE CURSOR TimeSheetCSV FROM ARRAY laFld
			
			IF This.CheckRights("TS_TIMESHEET_V")
                m.tsTypeCode = "T"
				SELECT curTimes
				GO top
				DO WHILE NOT EOF()
					SCATTER memvar
					INSERT INTO TimeSheetCSV FROM memvar
					SKIP
				ENDDO
			ENDIF
			
			IF This.CheckRights("TS_WAGES_V")
                m.tsTypeCode = "W"
				SELECT curWages
				GO top
				DO WHILE NOT EOF()
					SCATTER memvar
					INSERT INTO TimeSheetCSV FROM memvar
					SKIP
				ENDDO
			ENDIF
			
			IF This.CheckRights("TS_LEAVE_V")
                m.tsTypeCode = "L"
				SELECT curLeave
				GO top
				DO WHILE NOT EOF()
					SCATTER memvar
					INSERT INTO TimeSheetCSV FROM memvar
					SKIP
				ENDDO
			ENDIF
			
			IF This.CheckRights("TS_ALLOWANCES_V")
                m.tsTypeCode = "A"
				SELECT curAllowances
				GO top
				DO WHILE NOT EOF()
					SCATTER memvar
					INSERT INTO TimeSheetCSV FROM memvar
					SKIP
				ENDDO
			ENDIF
			
			IF This.CheckRights("TS_OTHER_V")
                m.tsTypeCode = "O"
				SELECT curOther
				GO top
				DO WHILE NOT EOF()
					SCATTER memvar
					INSERT INTO TimeSheetCSV FROM memvar
					SKIP
				ENDDO
			ENDIF

			SELECT TimeSheetCSV
			GO top
			DO WHILE NOT EOF('TimeSheetCSV')
				SELECT TimeSheetOut
				APPEND BLANK

				nEmpCode = -1
				cEmpFirst = "Unknown"
				cEmpLast = "Unknown"
				cCostName = "Unknown"
				cAllowName = SPACE(1)
				cWageName = SPACE(1)
				cPayName = SPACE(1)
				
				IF TimeSheetCSV.tsemp > 0
					IF SEEK(TimeSheetCSV.tsemp,"myStaff","mywebcode")
						nEmpCode = myStaff.myPayCode
						cEmpFirst = ALLTRIM(myStaff.myname)
						cEmpLast = ALLTRIM(myStaff.mysurname)
					ENDIF    
			   ENDIF						

				IF TimeSheetCSV.tscostcent > 0
					IF SEEK(TimeSheetCSV.tscostcent,"CostCent","code")
						*cCostName = ALLTRIM(CostCent.Name)
						replace cost_centre_code WITH TimeSheetCSV.tscostcent, cost_centre_name WITH ALLTRIM(CostCent.Name) IN TimeSheetOut
					ENDIF
				**JA 23/10/2012, display cost centre name as 'Employee Default' for rest	
				ELSE	
				   replace cost_centre_name WITH "Employee Default" IN TimeSheetOut
				ENDIF

				IF TimeSheetCSV.tscode > 0
					IF SEEK(TimeSheetCSV.tscode,"Allow","code")
						*cAllowName = ALLTRIM(Allow.Name)
						replace sub_type_code WITH TimeSheetCSV.tscode, sub_type_name WITH ALLTRIM(Allow.Name) IN TimeSheetOut
					ENDIF
				ENDIF

				IF TimeSheetCSV.tswagetype > 0
					IF SEEK(TimeSheetCSV.tswagetype,"Wagetype","code")
						*cWageName = ALLTRIM(Allow.Name)
						replace sub_type_code WITH TimeSheetCSV.tswagetype, sub_type_name WITH ALLTRIM(Wagetype.Name) IN TimeSheetOut
					ENDIF
				ENDIF

				IF TimeSheetCSV.tspay > 0
					IF SEEK(TimeSheetCSV.tspay,"myPays","pay_pk")
						cPayName = ALLTRIM(myPays.Pay_Name)
					ENDIF
				ENDIF
		
				plAusie = This.IsAustralia()

				cTSType = SPACE(1)
				IF TimeSheetCSV.tstype == "M"
					cTSType = "TimeSheet"
					Replace Type WITH CTSType
				ENDIF
				IF TimeSheetCSV.tstype == "W"
					cTSType = "Wages"
				ENDIF
				IF TimeSheetCSV.tstype == "N"
					*cTSType = "Long Service Name"
					cTSType = "Leave"
					replace sub_type_name WITH "Long Service Leave" IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "Z"
					*cTSType = "Alternative Leave Paid"
					cTSType = "Leave"
					replace sub_type_name WITH "Alternative Leave Paid"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "A"
					cTSType = "Allowance"
				ENDIF
				IF TimeSheetCSV.tstype == "D"
					*cTSType = "Other Leave"
					cTSType = "Other"
					replace sub_type_name WITH ALLTRIM(poStaff.oData.myHpUnits) + " Paid for Holiday Pay"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "Y"
					*cTSType = "Alternative Leave Accrued"
					cTSType = "Leave"
					replace sub_type_name WITH "Alternative Leave Accrued"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "B"
					*cTSType = "Bereavement Leave"
					cTSType = "Leave"
					replace sub_type_name WITH "Bereavement Leave"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "U"
					cTSType = "Leave"
					replace sub_type_name WITH "Unpaid Leave"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "P"
					*cTSType = "Public Holiday Leave"
					cTSType = "Leave"
					replace sub_type_name WITH "Public Holiday Leave"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "T"
					*cTSType = "Shift Leave"
					cTSType = "Leave"
					replace sub_type_name WITH "Other Leave"  IN TimeSheetOut
				ENDIF
				**JA 23/10/2012, introduced a check for shift leave
				IF TimeSheetCSV.tstype == "F"
					*cTSType = "Shift Leave"
					cTSType = "Leave"
					replace sub_type_name WITH "Shift Leave"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "O"
					cTSType = "Leave"
					replace sub_type_name WITH IIF(plAusie,"Annual Leave","Holiday Pay")  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "S"
					cTSType = "Leave"
					replace sub_type_name WITH IIF(plAusie,"Personal Leave","Sick Leave")  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "L"
					*cTSType = "Lieu Time"
					cTSType = "Leave"
					replace sub_type_name WITH "Lieu Time"  IN TimeSheetOut
				ENDIF
				IF TimeSheetCSV.tstype == "R"
					*cTSType = "Rostered Day Off"
					cTSType = "Other"
					replace sub_type_name WITH IIF(plAusie,"Rostered Day Off", ALLTRIM(poStaff.oData.myHpUnits) + " Paid for Relevant Daily Rate")  IN TimeSheetOut
				ENDIF

*!*					replace	Date WITH TimeSheetCSV.tsdate,;
*!*							Employee_Code WITH nEmpCode,;
*!*							First_Names WITH cEmpFirst,;
*!*							SurName WITH cEmplast,;
*!*							Start WITH IIF(not EMPTY(TimeSheetCSV.tsstart),TTOC(TimeSheetCSV.tsstart,2),""),;
*!*							End WITH IIF(not EMPTY(TimeSheetCSV.tsfinish),TTOC(TimeSheetCSV.tsfinish,2),""),;
*!*							Break WITH IIF(not EMPTY(TimeSheetCSV.tsbreak),TTOC(TimeSheetCSV.tsbreak,2),""),;
*!*							TYPE WITH cTSType,;
*!*							Units WITH TimeSheetCSV.tsunits,;
*!*							Cost_Centre_Code WITH TimeSheetCSV.tscostcent,;
*!*							Cost_Centre_Name WITH cCostName,;
*!*							Allowance_Code WITH TimeSheetCSV.tscode,;
*!*							Allowance_Name WITH cAllowName,;
*!*							Wage_Type WITH TimeSheetCSV.tswageType,;
*!*							Wage_Name WITH cWageName,;
*!*							Rate_Code WITH TimeSheetCSV.tsratecode,;
*!*							Units_To_Reduce WITH TimesheetCSV.tsunits2,;
*!*							Status WITH IIF(TimeSheetCSV.tsapproved,"Approved","Unapproved"),;
*!*							Pay WITH cPayName
				replace	Date WITH TimeSheetCSV.tsdate,;
						Group WITH rcGroupInfo, ; 
						Employee_Code WITH nEmpCode,;
						Employee_FirstNames WITH cEmpFirst,;
						Employee_LastName WITH cEmplast,;
						Start WITH IIF(not EMPTY(TimeSheetCSV.tsstart),TTOC(TimeSheetCSV.tsstart,2),""),;
						End WITH IIF(not EMPTY(TimeSheetCSV.tsfinish),TTOC(TimeSheetCSV.tsfinish,2),""),;
						Break WITH IIF(not EMPTY(TimeSheetCSV.tsbreak),TTOC(TimeSheetCSV.tsbreak,2),""),;
						TYPE WITH cTSType,;
						Units WITH TimeSheetCSV.tsunits,;
						Units_To_Reduce WITH TimesheetCSV.tsunits2,;
						Status WITH IIF(TimeSheetCSV.tsapproved,"Approved","Unapproved"),;
						Pay WITH cPayName

   		       **JA code change 01/11/2012
  			   REPLACE  isDownloaded WITH IIF(TimeSheetCSV.tsDownload,"Yes","No ")
			   **
				SKIP IN 'TimeSheetCSV'
			ENDDO
											
			SELECT TimeSheetOut
			GO top
			
			IF NOT EOF()
			    *JA code change 01/11/2012
			    	LOCAL lcFileName, lcPathName, lcOutputFile 
				lcPathName = THIS.companydatapath() + ADDBS("reports")
				IF NOT DIRECTORY(lcPathName)
					TRY
						MKDIR (lcPathName)
					CATCH
						THIS.cerror = "Reports foldameer was not created."
					ENDTRY
				ENDIF

			   	IF tlIsHistory
			   		lcFileName = "TIMEHISTORY"
			   	ELSE
			   		lcFileName = "TIMESHEET"
				ENDIF 
				lcOutputFile = lcPathName + TRANSFORM(THIS.employee) + '_' + lcFileName + '.CSV'
						
				*JA code change 01/11/2012
				IF tlIsHistory
					lcURLpdf = "http" + LOWER(STREXTRACT(UPPER(Request.GetCurrentUrl(.F.)),"HTTP","/TIMEHISTORY")) + "/GetCSV.si?date="+lcFileName
				ELSE
					lcURLpdf = "http" + LOWER(STREXTRACT(UPPER(Request.GetCurrentUrl(.F.)),"HTTP","/TIMESHEET")) + "/GetCSV.si?date="+lcFileName
				ENDIF 
				
				IF DIRECTORY(lcPathName)
			            IF FILE(lcOutputFile)
		    		        	DELETE FILE (lcOutputFile)
	        		   	ENDIF
					COPY TO (lcOutputFile) TYPE CSV FIELDS EXCEPT tsDate, tsDownload, tsApproved
					*This.AddUserInfo("Timesheet CSV File Created")
					*JA code change 01/11/2012
					IF tlIsHistory
					  lDispFileName="TimesheetHistory.csv"
					ELSE
					  lDispFileName="Timesheet.csv"
					ENDIF
					
					Response.Downloadfile(lcOutputfile,"application/csv",lDispFileName)
					RETURN
					
				ENDIF
			ENDIF
		ENDIF

		poPage = This.NewPageObject("timesheets:" + IIF(tlIsHistory, "history", "entry"), "time_entry")
		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)

	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE TimeHistoryPage()
		This.TimeEntryPage(.T.)
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE ReportsPage()
		PRIVATE poPage, pnNumPayslips, pnNumReports

		pnNumPayslips = This.GetReports(.T., "sortedPayslips")
		pnNumReports = This.GetReports(.F., "sortedReports")

		poPage = This.NewPageObject("reports:reports", "reports")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)

		IF pnNumPayslips != 0
			SELECT sortedPayslips
			USE IN 0
		ENDIF
		IF pnNumReports != 0
			SELECT sortedReports
			USE IN 0
		ENDIF
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE PolicyDocumentsPage()
		PRIVATE poPage

		poPage = This.NewPageObject("docs:docs", "docs")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*================================================================================*

	* CM support custom applications...
	PROCEDURE SpecialPage()
		PRIVATE poPage

		poPage = This.NewPageObject("special:special", "special")
		Response.ExpandScript(This.CompanyDataPath() + "master" + SOURCE_EXT, Server.nScriptMode)

	ENDPROC

	*################################################################################*
#DEFINE TOC_SubPage_

	*> +define: SubPage
	FUNCTION SubPage(tcTemplate as String, tcSection as String, toArgs as Collection) as String
		LOCAL loResponse, lcOutput, lnI, lnJ, lcHeaders, lcHeader, lcValue, lcCode
		PRIVATE pcSection, poArgs
		
		IF !FILE(This.CompanyHtmlPath() + tcTemplate + IIF(Server.nScriptMode == 2, ".FXP", SOURCE_EXT))
			RETURN "[Failed to open template: " + tcTemplate + "]"
		ENDIF

		pcSection = tcSection		&& This is used in templates to specify which section is rendered.
		poArgs = toArgs				&& This is used in "control" templates to supply config params.
		IF ISNULL(poArgs) OR VARTYPE(poArgs) != "O"
			* Ensure this always exists and is a Key/Value collection...
			poArgs = CREATEOBJECT("COLLECTION")
			poArgs.Add("_", "_")
		ENDIF

		loResponse = CREATEOBJECT([WWC_PAGERESPONSE])

		* Expand the template in a new page object
		loResponse.ExpandScript(This.CompanyHtmlPath() + tcTemplate + SOURCE_EXT, Server.nScriptMode)

		* Get the rendered page; headers + content
		lcOutput = loResponse.Render()

		* Copy any headers and cookies to the real page
		lcHeaders = SUBSTR(lcOutput, 1, AT(CRLF + CRLF, lcOutput) + 1)	&& including the first of the CRLF completely!
		lnI = AT(CRLF, lcHeaders)
		DO WHILE lnI > 1
			lnJ = AT(':', SUBSTR(lcHeaders, 1, lnI))
			IF lnJ > 0
				lcHeader = SUBSTR(lcHeaders, 1, lnJ - 1)
				IF !(lcHeader == "Content-Type" OR lcHeader == "Content-Length" OR lcHeader == "RequestId")
					* Found a header not already repeated; add it...
					lcValue = SUBSTR(lcHeaders, lnJ + 2, lnI - lnJ - 2)
					Response.AppendHeader(lcHeader, lcValue)
				ENDIF
			ENDIF

			lcHeaders = SUBSTR(lcHeaders, lnI + 2)
			lnI = AT(CRLF, lcHeaders)
		ENDDO

		* Return the content (minus it's headers) to the page
		lcOutput = SUBSTR(lcOutput, AT(CRLF + CRLF, lcOutput) + 4)		&& Starts at the begining of the content (after the first blank line)

		RETURN lcOutput
	ENDFUNC

	*################################################################################*
#DEFINE TOC_NewObjects_

	&&TODO: move these to the Factory
	*> +define: NewObjects
	FUNCTION NewTabObject(tcTitle AS String, tlCurrent AS Boolean, tcURI AS String)
		LOCAL loTab

		loTab = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loTab, "title",		tcTitle)
		ADDPROPERTY(loTab, "current",	tlCurrent)
		ADDPROPERTY(loTab, "URI",		tcURI)

		RETURN loTab
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewPageObject(tcPageID AS String, tcTemplate AS String)
		LOCAL loPage

		loPage = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loPage, "ID",		tcPageID)
		ADDPROPERTY(loPage, "title",	"")
		ADDPROPERTY(loPage, "heading",	"")
		ADDPROPERTY(loPage, "template",	tcTemplate)

		RETURN loPage
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewLeaveBalObject(tcTitle AS String, toEntitlements AS Collection, toBalances AS Collection, tcApplyURL AS String)
		LOCAL loBal

		loBal = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loBal, "title",			tcTitle)
		ADDPROPERTY(loBal, "entitlements",	toEntitlements)
		ADDPROPERTY(loBal, "balances",		toBalances)
		ADDPROPERTY(loBal, "applyURL",		tcApplyURL)

		* calculate the size of the list of titles and headings, so the boxes can be sorted into height order.
		ADDPROPERTY(;
			loBal,;
			"sortVal",;
			IIF(toEntitlements.Count > 0, toEntitlements.Count + 1, 0);
				+ IIF(toBalances.Count > 0, toBalances.Count + 1, 0);
				+ IIF(toEntitlements.Count > 0 AND toBalances.Count > 0, 1, 0);
		)		&& Long Service Leave has a long title that must wrap so is taller...but only by 0.5 really, so leaving it out..: + IIF(tcTitle == "Long Service Leave", 1, 0);

		RETURN loBal
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewLeaveCodeObject(;
		tcCode AS String, tcDisplayCode AS String, tcName AS String, tcUnitsName AS String,;
		tlEnableReduce AS Boolean, tcUnitsHint AS String, tcReduceHint AS String,;
		tnBalanceValue AS String, tlViewBalance AS Boolean, tcBalanceStrPrefix AS String, tcBalanceStrSuffix)
		LOCAL loCode

		loCode = CREATEOBJECT("EMPTY")

		ADDPROPERTY(loCode, "code",				tcCode)
		ADDPROPERTY(loCode, "displayCode",		tcDisplayCode)
		ADDPROPERTY(loCode, "name",				tcName)
		ADDPROPERTY(loCode, "units",			tcUnitsName)
		ADDPROPERTY(loCode, "enableReduce",		tlEnableReduce)
		ADDPROPERTY(loCode, "unitsHint",		tcUnitsHint)
		ADDPROPERTY(loCode, "reduceHint",		IIF(tlEnableReduce, tcReduceHint, "N/A"))
		ADDPROPERTY(loCode, "balanceValue",		tnBalanceValue)
		ADDPROPERTY(loCode, "balanceString",	IIF(tlViewBalance, ALLTRIM(tcBalanceStrPrefix) + ' ' + TRANSFORM(tnBalanceValue) + ' ' + ALLTRIM(tcBalanceStrSuffix), ""))

		RETURN loCode
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewLeaveMessageObject(tnId, tnFrom, tnTo, tcFromName, tcToName, tcSubject, tcMessage, ttRead, ttSent)
		LOCAL loMessage

		loMessage = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loMessage, "id",		tnId)
		ADDPROPERTY(loMessage, "from",		tnFrom)
		ADDPROPERTY(loMessage, "to",		tnTo)
		ADDPROPERTY(loMessage, "fromName",	tcFromName)
		ADDPROPERTY(loMessage, "toName",	tcToName)
		ADDPROPERTY(loMessage, "subject",	tcSubject)
		ADDPROPERTY(loMessage, "message",	tcMessage)
		ADDPROPERTY(loMessage, "read",		ttRead)
		ADDPROPERTY(loMessage, "sent",		ttSent)

		RETURN loMessage
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewTimesheetTypeObject(tcId, tcTypeChar, tcTitle, tcAuthz, tcFields, tcCursorName, tnRowCount) AS Object
		LOCAL loTimesheetType, lcFieldsTest

		loTimesheetType = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loTimesheetType, "id",			tcId)
		ADDPROPERTY(loTimesheetType, "tsType",		tcTypeChar)
		ADDPROPERTY(loTimesheetType, "title",		tcTitle)
		ADDPROPERTY(loTimesheetType, "authz",		STRTRAN(tcAuthz, '.CheckRights', "Process.CheckRights"))
		ADDPROPERTY(loTimesheetType, "fields",		tcFields)
		ADDPROPERTY(loTimesheetType, "cursorName",	tcCursorName)
		ADDPROPERTY(loTimesheetType, "rowCount",	tnRowCount)

		lcFieldsTest = UPPER(tcFields)

		IF 'D' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showDate",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyDate",		'd' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showDate",			.F.)
		ENDIF

		IF 'T' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showLeaveType",		.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyLeaveType",	't' $ tcFields)

			* LeaveType value controls the text of the hints and the enablement of the units and reduce fields.
			ADDPROPERTY(loTimesheetType, "showUnitsHint",		.T.)
			ADDPROPERTY(loTimesheetType, "showReduceHint",		.T.)
		ELSE
			ADDPROPERTY(loTimesheetType, "showLeaveType",		.F.)

			* LeaveType value controls the text of the hints and the enablement of the units and reduce fields.
			ADDPROPERTY(loTimesheetType, "showUnitsHint",		.F.)
			ADDPROPERTY(loTimesheetType, "showReduceHint",		.F.)
		ENDIF

		IF 'O' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showOtherType",		.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyOtherType",	'o' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showOtherType",		.F.)
		ENDIF

		IF 'K' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showCode",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyCode",		'k' $ tcFields)

			* AllowanceCode controls the text of the hint and the enablement of the units field.
			loTimesheetType.showUnitsHint = .T.
		ELSE
			ADDPROPERTY(loTimesheetType, "showCode",			.F.)
		ENDIF

		IF 'S' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showStart",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyStart",		's' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showStart",			.F.)
		ENDIF

		IF 'F' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showEnd",				.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyEnd",			'f' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showEnd",				.F.)
		ENDIF

		IF 'B' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showBreak",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyBreak",		'b' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showBreak",			.F.)
		ENDIF

		IF 'U' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showUnits",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyUnits",		'u' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showUnits",			.F.)
		ENDIF

		IF 'R' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showReduce",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyReduce",		'r' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showReduce",			.F.)
		ENDIF

		IF 'W' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showWageType",		.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyWageType",	'w' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showWageType",		.F.)
		ENDIF

		IF 'A' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showRateCode",		.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyRateCode",	'a' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showRateCode",		.F.)
		ENDIF

		IF 'C' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showCostCent",		.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyCostCent",	'c' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showCostCent",		.F.)
		ENDIF

		IF 'J' $ lcFieldsTest
			ADDPROPERTY(loTimesheetType, "showJobCode",			.T.)
			ADDPROPERTY(loTimesheetType, "readOnlyJobCode",		'j' $ tcFields)
		ELSE
			ADDPROPERTY(loTimesheetType, "showJobCode",			.F.)
		ENDIF

		RETURN loTimesheetType
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION NewTemplateTypeObject(tcId, tcTypeChar, tcTitle, tcAuthz, tcFields, tcCursorName, tnRowCount) AS Object
		LOCAL loTemplateType, lcFieldsTest

		loTemplateType = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loTemplateType, "id",			tcId)
		ADDPROPERTY(loTemplateType, "tsType",		tcTypeChar)
		ADDPROPERTY(loTemplateType, "title",		tcTitle)
		ADDPROPERTY(loTemplateType, "authz",		STRTRAN(tcAuthz, '.CheckRights', "Process.CheckRights"))
		ADDPROPERTY(loTemplateType, "fields",		tcFields)
		ADDPROPERTY(loTemplateType, "cursorName",	tcCursorName)
		ADDPROPERTY(loTemplateType, "rowCount",		tnRowCount)

		lcFieldsTest = UPPER(tcFields)

		IF 'E' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showWeek",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyWeek",		'e' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showWeek",			.F.)
		ENDIF

		IF 'D' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showDay",				.T.)
			ADDPROPERTY(loTemplateType, "readOnlyDay",			'd' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showDay",				.F.)
		ENDIF

		IF 'T' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showLeaveType",		.T.)
			ADDPROPERTY(loTemplateType, "readOnlyLeaveType",	't' $ tcFields)

			* LeaveType value controls the text of the hints and the enablement of the units and reduce fields.
			ADDPROPERTY(loTemplateType, "showUnitsHint",		.T.)
			ADDPROPERTY(loTemplateType, "showReduceHint",		.T.)
		ELSE
			ADDPROPERTY(loTemplateType, "showLeaveType",		.F.)

			* LeaveType value controls the text of the hints and the enablement of the units and reduce fields.
			ADDPROPERTY(loTemplateType, "showUnitsHint",		.F.)
			ADDPROPERTY(loTemplateType, "showReduceHint",		.F.)
		ENDIF

		IF 'O' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showOtherType",		.T.)
			ADDPROPERTY(loTemplateType, "readOnlyOtherType",	'o' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showOtherType",		.F.)
		ENDIF

		IF 'K' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showCode",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyCode",		'k' $ tcFields)

			* AllowanceCode controls the text of the hint and the enablement of the units field.
			loTemplateType.showUnitsHint = .T.
		ELSE
			ADDPROPERTY(loTemplateType, "showCode",			.F.)
		ENDIF

		IF 'S' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showStart",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyStart",		's' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showStart",			.F.)
		ENDIF

		IF 'F' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showEnd",				.T.)
			ADDPROPERTY(loTemplateType, "readOnlyEnd",			'f' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showEnd",				.F.)
		ENDIF

		IF 'B' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showBreak",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyBreak",		'b' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showBreak",			.F.)
		ENDIF

		IF 'U' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showUnits",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyUnits",		'u' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showUnits",			.F.)
		ENDIF

		IF 'R' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showReduce",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyReduce",		'r' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showReduce",			.F.)
		ENDIF

		IF 'W' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showWageType",		.T.)
			ADDPROPERTY(loTemplateType, "readOnlyWageType",	'w' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showWageType",		.F.)
		ENDIF

		IF 'A' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showRateCode",		.T.)
			ADDPROPERTY(loTemplateType, "readOnlyRateCode",	'a' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showRateCode",		.F.)
		ENDIF

		IF 'C' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showCostCent",		.T.)
			ADDPROPERTY(loTemplateType, "readOnlyCostCent",	'c' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showCostCent",		.F.)
		ENDIF

		IF 'J' $ lcFieldsTest
			ADDPROPERTY(loTemplateType, "showJobCode",			.T.)
			ADDPROPERTY(loTemplateType, "readOnlyJobCode",		'j' $ tcFields)
		ELSE
			ADDPROPERTY(loTemplateType, "showJobCode",			.F.)
		ENDIF

		RETURN loTemplateType
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*@ NewAllowanceObject:
	 * Parameters:
	 *	tcCode:			The key value as a string - used on the generated select list as the option value
	 *	tcName:			The name of this entry - used on the generated select list as the option text
	 *	tnType:			The value from ALLOW.CALC_HOW - used to set up other properties
	 *	tnAmount:		The value from ALLOW.AMOUNT - for some calculation types this an override value
	 *	tnCostCentre:	The value from ALLOW.COST_CENT - when non-zero this is an override value for costCentre
	 * Returns:
	 *	A new Allowance object for the above settings.
	 *	The object has the following properties, which unless otherwise specified are copies of the input parameters:
	 *		code, name, type, amount, costCentre,
	 *		unitsHint: the text of the hint tooltip shown next to the units field when this is the selected allowance code,
	 *		enableUnits: boolean indicating if the units field should be disabled (and therefore overridden) when this is the selected allowance code,
	 *		unitsValue: the override value to populate the units field with when this is the selected allowance code.
	 * Notes:
	 *	Allowances now have enough information to support overriding of units/costCentre and supplying of hint text, so need their own object.
	FUNCTION NewAllowanceObject(tcCode, tcName, tnType, tnAmount, tnCostCentre, tcHide) AS Object
		LOCAL loCode, llEnableUnits, lcUnitsHint, lnUnitsValue

		loCode = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loCode, "code",			ALLTRIM(tcCode))
		ADDPROPERTY(loCode, "name",			ALLTRIM(tcName))
		ADDPROPERTY(loCode, "type",			tnType)
		ADDPROPERTY(loCode, "amount",		tnAmount)
		ADDPROPERTY(loCode, "costCentre",	tnCostCentre)	&& if nonEmpty, this is the override value and implies a disabled field
		ADDPROPERTY(loCode, "hide",			tcHide)

		* Copied from payroll's import.spr:
		*!* ALLOW.CALC_HOW:
		*!* 	1 = Fixed Dollar Amount
		*!* 			TTP:4173; 0 -> units; 0 -> rate; msiUnits -> amount unless the allowance has a specified amount, in which case we use that regardless.
		*!* 	2 = % of Wage & Salary
		*!* 			TTP:4555; 0 -> units; msiUnits -> rate; 0 -> amount
		*!* 	3 = Total Hours
		*!* 			TTP:4555; 0 -> units; msiUnits -> rate; 0 -> amount
		*!* 	4 = Equivalent Hours
		*!* 			TTP:4555; 0 -> units; msiUnits -> rate; 0 -> amount
		*!* 	5 = Specific Hours
		*!* 			msiUnits -> units; 0 -> rate; 0 -> amount
		*!* 	6 = Rated Units
		*!* 			msiUnits -> units; 0 -> rate; 0 -> amount
		*!* 	7 = Hourly Rate
		*!* 			msiUnits -> units; 0 -> rate; 0 -> amount
		*!* 	8 = % of Total Gross
		*!* 			TTP:4555; 0 -> units; msiUnits -> rate; 0 -> amount
		*!* 	9 = Motor Vehicle	[AU only]
		*!* 			msiUnits -> units; 0 -> rate; 0 -> amount
		*!* 	10 = Accommodation	[AU only]
		*!* 			msiUnits -> units; 0 -> rate; 0 -> amount

		lnUnitsValue = 0
		llEnableUnits = .T.

		&&NOTE: the logic here around tnType must match that in the payroll import code and that in the GroupSummaryPage() calculations around calc_how!

		IF tnType == 1
			* Fixed Dollar Amount: entered units are the amount, unless there is an override value.
			lnUnitsValue = tnAmount
			llEnableUnits = EMPTY(tnAmount)

			IF llEnableUnits
				lcUnitsHint = "Dollar Amount"
			ELSE
				lcUnitsHint = "Dollar Amount; Pre-set"
			ENDIF
		ELSE
			* Entered units are the rate, unless there is an override value.
			IF INLIST(tnType, 2, 3, 4, 8)
				lnUnitsValue = tnAmount
				llEnableUnits = EMPTY(tnAmount)

				IF llEnableUnits
					lcUnitsHint = "Rate"
				ELSE
					lcUnitsHint = "Rate; Pre-set"
				ENDIF
			ELSE
				* Entered units are units, and the value in tnAmount is the rate, not an override value.
				lcUnitsHint = "Units"
			ENDIF
		ENDIF

		ADDPROPERTY(loCode, "unitsHint",	lcUnitsHint)
		ADDPROPERTY(loCode, "enableUnits",	llEnableUnits)
		ADDPROPERTY(loCode, "unitsValue",	lnUnitsValue)

		RETURN loCode
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Used for all simple code types.
	FUNCTION NewCodeObject(tcCode, tcName) AS Object
		LOCAL loCode

		loCode = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loCode, "code",	ALLTRIM(tcCode))
		ADDPROPERTY(loCode, "name",	ALLTRIM(tcName))

		RETURN loCode
	ENDFUNC

	*--------------------------------------------------------------------------------*
	* Used for all simple code types. (With Hide)
	FUNCTION NewCodeObjectHide(tcCode, tcName, tcHide) AS Object
		LOCAL loCode

		loCode = CREATEOBJECT("EMPTY")
		ADDPROPERTY(loCode, "code",	ALLTRIM(tcCode))
		ADDPROPERTY(loCode, "name",	ALLTRIM(tcName))
		ADDPROPERTY(loCode, "hide",	ALLTRIM(tcHide))

		RETURN loCode
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*################################################################################*
#DEFINE TOC_Getters_

	*> +define: Getters
	FUNCTION GetMainTabs(tcPageID AS String) AS Collection
		* tcPageID is maintabid:subtabid sort of format
		LOCAL loTabs, loXML, loTab, loNode, lcAuthz, llAuthd, loSubNode
		LOCAL llAdmin, llPayrollUser, llManager, llFoundCurrent, llPassReset
		
		llAdmin = This.CheckRights("ADMIN")
		llPayrollUser = This.IsPayrollUser(This.Employee)
		llManager = This.IsManager(This.Employee)
		
		llPassReset = IIF(ALLTRIM(Session.GetSessionVar("isPassReset"))="Y",.T.,.F.)
			

		loTabs = CREATEOBJECT("COLLECTION")
		loXML = NEWOBJECT("MSXML2.DOMDocument")

		loXML.async = .F.
		loXML.Load(ADDBS(This.cDataPath) + "sitemap.xml")

		IF loXML.parseError.errorCode # 0
			This.AddError("Sitemap not found: " + TRANSFORM(loXML.parseError.errorCode) + ": " + loXML.parseError.reason)
			RETURN loTabs
		ENDIF

		llFoundCurrent = .F.

		loNode = loXML.documentElement.firstChild
		DO WHILE !ISNULL(loNode)
			* Loop thru the top-level siteMapNode elements...
			IF loNode.nodeName == "#comment"
				* skip comments as the code following requires complete siteMapNode's only
				loNode = loNode.nextSibling
				LOOP
			ENDIF

			lcAuthz = loNode.attributes.getNamedItem("authz").text
			llAuthd = .T.	&& default to having access if authz==""
			IF llPassReset AND UPPER(lcAuthz)<>"LLPASSRESET"
			  	 llAuthd = .F.
			ELSE   
				IF !(lcAuthz == "")	&& dam set exact off means I have to do it this way around!!
					* Run the authz macro to find out if the user has access
					 llAuthd = &lcAuthz
				ENDIF
			ENDIF
			IF llAuthd
				* Create a tab for this entry...
				** the tab is current if .id + ':' is in tcPageID - i.e. .id is the first part.
				loTab = This.NewTabObject(;
					loNode.attributes.getNamedItem("title").text,;
					(loNode.attributes.getNamedItem("id").text + ':') $ tcPageID,;
					loNode.attributes.getNamedItem("url").text;
				)

				IF loTab.current
					* we don't set the error straight away if the current tab is not auth'd as there may be an alternative branch that matches this pageName instead.
					llFoundCurrent = .T.
				ENDIF

				IF !ISNULL(loNode.attributes.getNamedItem("tabTitle"))
					loTab.title = loNode.attributes.getNamedItem("tabTitle").text
				ENDIF

				IF EMPTY(loTab.URI)
					* If the URL is empty, look it up as the first auth'd child link...
					loSubNode = loNode.firstChild
					DO WHILE !ISNULL(loSubNode)
						* Loop thru the child elements...
						IF loSubNode.nodeName == "#comment"
							* skip comments as the code following requires complete siteMapNode's only
							loSubNode = loSubNode.nextSibling
							LOOP
						ENDIF

						lcAuthz = loSubNode.attributes.getNamedItem("authz").text
						llAuthd = .T.	&& default to having access if authz==""
						IF !(lcAuthz == "")	&& dam set exact off means I have to do it this way around!!
							* Run the authz macro to find out if the user has access
							llAuthd = &lcAuthz
						ENDIF
						IF llAuthd
							* found a tab with access; copy it's URL and exit
							loTab.URI = loSubNode.attributes.getNamedItem("url").text
							loSubNode = .null.
						ELSE
							* try again...
							loSubNode = loSubNode.nextSibling
						ENDIF
					ENDDO
				ENDIF

				IF !EMPTY(loTab.URI) AND (loTab.current OR ISNULL(loNode.attributes.getNamedItem("hidden")))	&& ...and not current or hidden
					loTabs.add(loTab)
				ENDIF
			ENDIF

			loNode = loNode.nextSibling
		ENDDO

		IF !llFoundCurrent
			* User has no access to [any version of] the current main tab!
			This.AddError("You are not authorised to access this page!")
			* not returning here, so the user has the chance to be presented with somewhere to go..!
		ENDIF

		RETURN loTabs
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION GetSubTabs(tcPageID AS String, rcPageTitle AS String, rcPageHeading AS String, rcHelpPageURL AS String) AS Collection
		* tcPageID is maintabid:subtabid sort of format
		* rcPageTitle and rcPageHeading must be passed by reference
		LOCAL loTabs, loXML, loTab, loNode, loFoundNode, lcAuthz, llAuthd
		LOCAL llAdmin, llPayrollUser, llManager, llFoundCurrent, llPassReset

		rcPageTitle = "Error"
		rcPageHeading = "Error"
		rcHelpPageURL = ""

		llPassReset = IIF(ALLTRIM(Session.GetSessionVar("isPassReset"))="Y",.T.,.F.)

		llAdmin = This.CheckRights("ADMIN")
		llPayrollUser = This.IsPayrollUser(This.Employee)
		llManager = This.IsManager(This.Employee)

		loTabs = CREATEOBJECT("COLLECTION")
		loXML = NEWOBJECT("MSXML2.DOMDocument")

		loXML.async = .F.
		loXML.Load(ADDBS(This.cDataPath) + "sitemap.xml")

		IF loXML.parseError.errorCode # 0
			This.AddError("Sitemap not found: " + TRANSFORM(loXML.parseError.errorCode) + ": " + loXML.parseError.reason)
			RETURN loTabs
		ENDIF

		llFoundCurrent = .F.

		loNode = loXML.documentElement.firstChild
		loFoundNode = .null.
		DO WHILE !ISNULL(loNode)
			* Loop thru the top-level siteMapNode elements...
			IF loNode.nodeName == "#comment"
				* skip comments as the code following requires complete siteMapNode's only
				loNode = loNode.nextSibling
				LOOP
			ENDIF

			lcAuthz = loNode.attributes.getNamedItem("authz").text
			llAuthd = .T.	&& default to having access if authz==""
			IF !(lcAuthz == "")	&& dam set exact off means I have to do it this way around!!
				* Run the authz macro to find out if the user has access
				llAuthd = &lcAuthz
			ENDIF
			IF llAuthd
				IF (loNode.attributes.getNamedItem("id").text + ':') $ tcPageID
					* Found the current, authorised, main tab
					loFoundNode = loNode
					EXIT
				ENDIF
			ENDIF

			loNode = loNode.nextSibling
		ENDDO

		loNode = loFoundNode
		IF ISNULL(loNode)
			* User has no access to [any version of] the current main tab!
			IF !("You are not authorised to access this page!" $ This.cError)
				This.AddError("You are not authorised to access this page!")
			ENDIF
			RETURN loTabs	&& hide all subtabs in this case
		ELSE
			rcPageHeading = loNode.attributes.getNamedItem("title").text
		ENDIF

		loNode = loNode.FirstChild
		DO WHILE !ISNULL(loNode)
			* Loop thru the 2nd-level siteMapNode elements...
			IF loNode.nodeName == "#comment"
				* skip comments as the code following requires complete siteMapNode's only
				loNode = loNode.nextSibling
				LOOP
			ENDIF

			lcAuthz = loNode.attributes.getNamedItem("authz").text
			llAuthd = .T.	&& default to having access if authz==""
			IF !(lcAuthz == "")	&& dam set exact off means I have to do it this way around!!
				* Run the authz macro to find out if the user has access
				llAuthd = &lcAuthz
			ENDIF
			IF llAuthd
				* Create a tab for this entry...
				** the tab is current if ':' + .id is in tcPageID - i.e. .id is the second part.
				loTab = This.NewTabObject(;
					loNode.attributes.getNamedItem("title").text,;
					(':' + loNode.attributes.getNamedItem("id").text) $ tcPageID,;
					loNode.attributes.getNamedItem("url").text;
				)

				IF !ISNULL(loNode.attributes.getNamedItem("tabTitle"))
					loTab.title = loNode.attributes.getNamedItem("tabTitle").text
				ENDIF

				IF loTab.current
					* we don't set the error straight away if the current tab is not auth'd as there may be an alternative branch that matches this pageName instead.
					llFoundCurrent = .T.
					rcPageTitle = loTab.title
					IF rcPageHeading != loTab.title
						rcPageHeading = rcPageHeading + " > " + loTab.title
					ENDIF

					IF !ISNULL(loNode.attributes.getNamedItem("helpPage"))
						rcHelpPageURL = "/help/" + loNode.attributes.getNamedItem("helpPage").text
					ENDIF
				ENDIF

				IF loTab.current OR ISNULL(loNode.attributes.getNamedItem("hidden"))	&& ...not current or hidden
					loTabs.add(loTab)
				ENDIF
			ENDIF

			loNode = loNode.nextSibling
		ENDDO

		IF !llFoundCurrent
			* User has no access to this subtab!
			IF !("You are not authorised to access this page!" $ This.cError)
				This.AddError("You are not authorised to access this page!")
			ENDIF
			* not returning here, so the user has the chance to be presented with somewhere to go..!
		ENDIF

		RETURN loTabs
	ENDFUNC

	*-=---=---=---=---=---=---=---=---=---=---=---=---=---=---=---=---=---=---=---=--*

	FUNCTION GetQuickLinks() AS Collection
		LOCAL loLinks, loXML, loLink, loNode, loParentNode, lcAuthz, llAuthd, lnIndex
		LOCAL llIsCat, loCategories, loURLs, loTitles, loAttrib, lcKey, lcValue, lcCatKey, loSortKeys
		LOCAL llAdmin, llPayrollUser, llManager, llFoundCurrent, llPassReset

		llAdmin = This.CheckRights("ADMIN")
		llPayrollUser = This.IsPayrollUser(This.Employee)
		llManager = This.IsManager(This.Employee)
		llPassReset = IIF(ALLTRIM(Session.GetSessionVar("isPassReset"))="Y",.T.,.F.)


		loLinks = CREATEOBJECT("COLLECTION")

		loSortKeys = CREATEOBJECT("COLLECTION")
		loSortKeys.keySort = 2

		loCategories = CREATEOBJECT("COLLECTION")

		loURLs = CREATEOBJECT("COLLECTION")
		loURLs.keySort = 2
		loTitles = CREATEOBJECT("COLLECTION")
		loTitles.keySort = 2

		loXML = NEWOBJECT("MSXML2.DOMDocument")

		loXML.async = .F.
		loXML.Load(ADDBS(This.cDataPath) + "sitemap.xml")

		IF loXML.parseError.errorCode # 0
			This.AddError("Sitemap not found: " + TRANSFORM(loXML.parseError.errorCode) + ": " + loXML.parseError.reason)
			RETURN loLinks
		ENDIF

		loNode = loXML.documentElement.firstChild
		loParentNode = .null.
		DO WHILE !ISNULL(loNode)
			* Loop thru the top-level siteMapNode elements...
			IF loNode.nodeName == "#comment"
				* skip comments as the code following requires complete siteMapNode's only
				loNode = loNode.nextSibling
				IF ISNULL(loNode) AND !ISNULL(loParentNode)
					* Out of siblings and can go up
					loNode = loParentNode.nextSibling
					* Remember to do this to avoid looping
					loParentNode = .null.
				ENDIF
				LOOP
			ENDIF

			lcAuthz = loNode.attributes.getNamedItem("authz").text
			llAuthd = .T.	&& default to having access if authz==""
			IF !(lcAuthz == "")	&& dam set exact off means I have to do it this way around!!
				* Run the authz macro to find out if the user has access
				llAuthd = &lcAuthz
			ENDIF
			IF llAuthd
				* Check that the quickLink attribs exist
				loAttrib = loNode.attributes.getNamedItem("quickCat")
				IF !ISNULL(loAttrib)
					* Fix the value so it string-sorts safely
					lnIndex = AT(':', loAttrib.text)
					lcCatKey = RIGHT(PADL(TRANSFORM(VAL(SUBSTR(loAttrib.text, 1, lnIndex - 1))), 7, '0'), 8)
					lcValue = SUBSTR(loAttrib.text, lnIndex + 1)

					loAttrib = loNode.attributes.getNamedItem("quickTitle")
					IF !ISNULL(loAttrib)
						* We have both attribs and are auth'd, so can keep this one...
						IF loCategories.GetKey(lcCatKey) == 0
							* Add the category if we haven't seen it yet
							loCategories.Add(lcValue, lcCatKey)
							* Pad out the sortKey so that category keys sort ahead of category contents!
							loSortKeys.Add(lcCatKey, lcCatKey + SPACE(9))
						ENDIF

						* Fix the value so it string-sorts safely
						lnIndex = AT(':', loAttrib.text)
						lcKey = RIGHT(PADL(TRANSFORM(VAL(SUBSTR(loAttrib.text, 1, lnIndex - 1))), 7, '0'), 8)
						lcValue = SUBSTR(loAttrib.text, lnIndex + 1)

						* Create a global-order sort key for the link
						lcKey = lcCatKey + ':' + lcKey

						loTitles.Add(lcValue, lcKey)
						loURLs.Add(loNode.attributes.getNamedItem("url").text, lcKey)
						loSortKeys.Add(lcKey, lcKey)
					ENDIF
				ENDIF

				* Move to the next node to check, which since we have access to this node, is its children
				IF !ISNULL(loNode.firstChild)
					* Remember where we came from.
					loParentNode = loNode			&&NOTE: this will break if the sitemap is ever deeper than tab/subtab!
					loNode = loNode.firstChild
				ELSE
					* No children, (already looking at subTabs?), so move sideways, or up
					loNode = loNode.nextSibling
					IF ISNULL(loNode) AND !ISNULL(loParentNode)
						* Out of siblings and can go up
						loNode = loParentNode.nextSibling
						* Remember to do this to avoid looping
						loParentNode = .null.
					ENDIF
				ENDIF
			ELSE
				* When not auth'd, move sideways or up
				loNode = loNode.nextSibling
				IF ISNULL(loNode) AND !ISNULL(loParentNode)
					* Out of siblings and can go up
					loNode = loParentNode.nextSibling
					* Remember to do this to avoid looping
					loParentNode = .null.
				ENDIF
			ENDIF
		ENDDO

		loSortKeys.keySort = 2
		FOR EACH lcKey IN loSortKeys
			llIsCat = (AT(':', lcKey) == 0)
			loLinks.Add(This.NewTabObject(;
				IIF(llIsCat, loCategories.Item(lcKey), loTitles.Item(lcKey)),;
				.F.,;
				IIF(llIsCat, "", loURLs.Item(lcKey));
			))
		NEXT

		RETURN loLinks
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> fix
	&&LATER: need to know what is meant by "new timesheet" and add it here...
	FUNCTION GetAlerts(tcPageID AS String) AS Collection
		LOCAL loTabs, lnNewMail, lnNewRequests, lnNewMessages, lnNewMgrMessages, loMessages

		loTabs = CREATEOBJECT("COLLECTION")

		IF tcPageID == "home:main"
			* The home page has it's own form of alerts in the page, so we don't repeat them.
			* (It asks for them again, passing blank as tcPageID)
			RETURN loTabs
		ENDIF

		loMessages = Factory.GetMessagesObject()
		lnNewMail = loMessages.GetNewMessageCount(This.Employee)

		lnNewRequests = 0
		IF This.SelectData(This.Licence, "leaveRequests")
			SELECT leaveRequests
			CALCULATE CNT() TO lnNewRequests FOR manager == This.Employee AND !accepted AND !declined

			lnNewMessages = 0
			IF This.SelectData(This.Licence, "leaveRequestStatus")
				SELECT COUNT(y.id) AS messagesCount;
					FROM leaveRequests x;
					JOIN leaveRequestStatus y ON x.id == y.leaveReqID;
					WHERE x.employee == This.Employee AND y.to == This.Employee AND EMPTY(y.read);
					INTO CURSOR curNewMessages
				lnNewMessages = curNewMessages.messagesCount
				USE IN SELECT("curNewMessages")

				SELECT COUNT(y.id) AS messagesCount;
					FROM leaveRequests x;
					JOIN leaveRequestStatus y ON x.id == y.leaveReqID;
					WHERE x.manager == This.Employee AND y.to == This.Employee AND EMPTY(y.read) AND !(x.employee == This.Employee);
					INTO CURSOR curNewMgrMessages
				lnNewMgrMessages = curNewMgrMessages.messagesCount
				USE IN SELECT("curNewMgrMessages")
			ENDIF
		ENDIF

		IF lnNewMail <> 0
			loTabs.Add(This.NewTabObject(;
				TRANSFORM(FLOOR(lnNewMail)) + " new message" + IIF(lnNewMail > 1, "s.", "."),;
				"mail",;
				"InboxPage.si";
			))
		ENDIF

		IF lnNewRequests <> 0
			loTabs.Add(This.NewTabObject(;
				TRANSFORM(FLOOR(lnNewRequests)) + " unprocessed leave request" + IIF(lnNewRequests > 1, "s.", "."),;
				"leaveApproval",;
				"ManageLeaveRequestPage.si";
			))
		ENDIF

		IF lnNewMessages <> 0
			loTabs.Add(This.NewTabObject(;
				TRANSFORM(FLOOR(lnNewMessages)) + " unread leave request message" + IIF(lnNewMessages > 1, "s", "") + " from your manager.",;
				"leaveMessage",;
				"ViewLeaveRequestPage.si?order=dateMade+desc&show=unread";
			))
		ENDIF

		IF lnNewMgrMessages <> 0
			loTabs.Add(This.NewTabObject(;
				TRANSFORM(FLOOR(lnNewMgrMessages)) + " unread leave request message" + IIF(lnNewMgrMessages > 1, "s.", "."),;
				"leaveMessage",;
				"ManageLeaveRequestPage.si";
			))
		ENDIF

		RETURN loTabs
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetReports(tlPayslips, tcCursor) AS Integer
		LOCAL lnNumPayslips

		IF VARTYPE(tcCursor) != 'C' OR tcCursor == ""
			tcCursor = "sortedReports"
		ENDIF

		** CM File map for standard payslips
		** "<webcode>_PAYSLIP_<YYYYMMDD>.PDF" (generally, PAYSLIP is the report type, so for other reports this can be anything)
		lnNumPayslips = ADIR(laFiles, This.CompanyDataPath() + ADDBS("payslips") + TRANSFORM(This.Employee) + IIF(tlPayslips, "_PAY*SLIP*.PDF", "_*.pdf"))

		IF VARTYPE(lnNumPayslips) == 'N' AND lnNumPayslips != 0
			CREATE CURSOR fileList (;
				fileName C(254),;
				fileSize I,;
				dateMod D,;
				timeMod C(9),;
				attribs C(5);
			)
			APPEND FROM ARRAY laFiles
			SELECT fileName;
				FROM fileList;
				ORDER BY dateMod DESC, timeMod DESC;
				INTO CURSOR (tcCursor);
				WHERE tlPayslips OR !LIKE("*_PAY*SLIP*.PDF*", fileName)

			lnNumPayslips = _TALLY

			SELECT fileList
			USE IN 0
		ELSE
			lnNumPayslips = 0
			* This means the uses of the above cursors are not run elsewhere - we are ok not creating them.
		ENDIF

		RETURN lnNumPayslips
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetLeaveCodes(tnEmployee AS Integer, tlUnrestricted AS Boolean) AS Collection
		LOCAL loStaff, loCodes, loCode

		IF tnEmployee > 0
			loStaff = Factory.GetStaffObject()
			IF !loStaff.Load(EVL(tnEmployee, This.Employee))
				This.AddError("Cannot load LeaveCodes: " + loStaff.cErrorMsg)
				RETURN null
			ENDIF
		ENDIF
			
		loCodes = CREATEOBJECT("COLLECTION")
		* HG 17/09/2009 TTP707,4169
		* changed the codes used in checkrights() from the view rights to the request rights (the newly added ones from bit 81 to 91)
		DO CASE
		* AU
		CASE This.IsAustralia()
			* personal leave
			IF tnEmployee > 0
				* annual leave
				IF This.CheckRights("LR_TYPE_ANNUAL") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
							"O", "AL", "Annual Leave",ALLTRIM(loStaff.oData.myHpUnits),	.F.,;
							ALLTRIM(loStaff.oData.myHpUnits) + " Taken", "",;
							IIF(This.CheckRights("LV_ANNUAL_ACCRUED"), loStaff.oData.myHpTotal, loStaff.oData.myHpTotal - loStaff.oData.myHpAccrue),;
							This.CheckRights("LV_ANU_BAL"), "Annual Leave Balance:", loStaff.oData.myHpUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				IF This.CheckRights("LR_TYPE_SICK") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"S", "PL", "Personal Leave",				"Hours",							.T.,;
						"Hours Taken.  The number of hours that the employee would normally work on this particular day of the week.",;
						ALLTRIM(loStaff.oData.mySpUnits) + " To Reduce Entitlement.  For an employee on a five day week, this would be one-fifth of the annual sick leave entitlement hours.",;
						loStaff.oData.mySpTotal,;
						This.CheckRights("LV_SIC_BAL"), "Personal Leave Balance:", loStaff.oData.mySpUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF
				
				* long service leave
				IF This.CheckRights("LR_TYPE_LONG") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"N", "LSL", "Long Service Leave",			ALLTRIM(loStaff.oData.myLslUnits),	.F.,;
						ALLTRIM(loStaff.oData.myLslUnits) + " Taken", "",;
						loStaff.oData.myLslTotal,;
						This.CheckRights("LV_LNG_BAL"), "Long Service Balance:", loStaff.oData.myLslUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* lieu time
				IF This.CheckRights("LR_TYPE_LIEU") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"L", "LT", "Lieu Time",ALLTRIM(loStaff.oData.myAltUnits),.T.,;
						ALLTRIM(loStaff.oData.myAltUnits) + " Worked",;
						ALLTRIM(loStaff.oData.myAltUnits) + " Taken",;
						loStaff.oData.myAltTotal,;
						This.CheckRights("LV_ALT_BAL"), "Lieu Time Balance:", loStaff.oData.myAltUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* rostered day off
				IF This.CheckRights("LR_TYPE_ROSTERED") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"R", "RD", "Rostered Day Off",				"Hours",							.F.,;
						ALLTRIM(loStaff.oData.myAltUnits) + " Taken", "",;
						loStaff.oData.myRdoTotal,;
						.T., "Rostered Day Off Balance:", "Hours")
					loCodes.Add(loCode, loCode.code)
				ENDIF
				
				* other leave
				IF (This.CheckRights("LR_TYPE_OTHER") OR tlUnrestricted)
				    LOCAL lcOthLName
				    lcOthLName =  IIF(!EMPTY(ALLTRIM(loStaff.oData.myOTName)),ALLTRIM(loStaff.oData.myOTName),'Other Leave')
					loCode = This.NewLeaveCodeObject(;
						"T", "OT", lcOthLName, ALLTRIM(loStaff.oData.myOtUnits),	.F.,;
						ALLTRIM(loStaff.oData.myOtUnits) + " Taken", "",;
						IIF(This.CheckRights("LV_OTHER_ACCRUED"), loStaff.oData.myOtTotal, loStaff.oData.myOtTotal - loStaff.oData.myOtAccrue),;
						This.CheckRights("LV_OTH_BAL"), ALLTRIM(loStaff.oData.myOtName) + " Balance:", loStaff.oData.myOtUnits)
					loCodes.Add(loCode, loCode.code)
			    ENDIF
				
				* shift leave
				IF (This.CheckRights("LR_TYPE_SHIFT") OR tlUnrestricted)
					LOCAL lcShiftLName
				    lcShiftLName =  IIF(!EMPTY(ALLTRIM(loStaff.oData.mySHName)),ALLTRIM(loStaff.oData.mySHName),'Shift Leave')
					loCode = This.NewLeaveCodeObject(;
						"F", "SH", lcShiftLName,ALLTRIM(loStaff.oData.myShUnits),	.F.,;
						ALLTRIM(loStaff.oData.myShUnits) + " Taken", "",;
						IIF(This.CheckRights("LV_SHIFT_ACCRUED"), loStaff.oData.myShTotal, loStaff.oData.myShTotal - loStaff.oData.myShAccrue),;
						This.CheckRights("LV_SFT_BAL"), ALLTRIM(loStaff.oData.myShName) + " Balance:", loStaff.oData.myShUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF
				
				* unpaid leave
				IF (This.CheckRights("LR_UNPAID_LEAVE") OR tlUnrestricted)
					loCode = This.NewLeaveCodeObject(;
						"U", "UN", "Unpaid Leave", "Hours", .F.,;
						"Hours Taken", "",;
						0,;
						.F., "Unpaid Leave Balance:", "")
					loCodes.Add(loCode, loCode.code)
				ENDIF


			ELSE
				loCode = This.NewLeaveCodeObject(;
						"O", "AL", "Annual Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"S", "PL", "Personal Leave","Hours",.T.,;
						"Hours Taken",;
						"Days Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"N", "LSL", "Long Service Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)
				
				loCode = This.NewLeaveCodeObject(;
						"L", "LT", "Lieu Time","Units",.T.,;
						"Units Worked",;
						"Units Taken",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"R", "RD", "Rostered Day Off","Units",.T.,;
						"Days Taken",;
						"Days Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)
				
				loCode = This.NewLeaveCodeObject(;
						"U", "UN", "Unpaid Leave", "Hours", .F.,;
						"Hours Taken", "",;
						0,;
						.F., "Unpaid Leave Balance:", "")
					loCodes.Add(loCode, loCode.code)
				
			ENDIF
			
		* NZ
		OTHERWISE
			IF tnEmployee > 0
				* annual leave
				IF This.CheckRights("LR_TYPE_ANNUAL") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"O", "AL", "Annual Leave",ALLTRIM(loStaff.oData.myHpUnits),.F.,;
						ALLTRIM(loStaff.oData.myHpUnits) + " Taken", "",;
						IIF(This.CheckRights("LV_ANNUAL_ACCRUED"),loStaff.oData.myHpTotal,loStaff.oData.myHpTotal - loStaff.oData.myHpAccrue),;
						This.CheckRights("LV_ANU_BAL"), "Annual Leave Balance:", loStaff.oData.myHpUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* sick leave
				IF This.CheckRights("LR_TYPE_SICK") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"S", "SL", "Sick Leave","Hours",.T.,;
						"Hours Taken.  The number of hours that the employee would normally work on this particular day of the week.",;
						ALLTRIM(loStaff.oData.mySpUnits) + " To Reduce Entitlement.  For an employee on a five day week, this would be one-fifth of the annual sick leave entitlement hours.",;
						loStaff.oData.mySpTotal,;
						This.CheckRights("LV_SIC_BAL"), "Sick Leave Balance:", loStaff.oData.mySpUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* long service leave
				IF This.CheckRights("LR_TYPE_LONG") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"N", "LSL", "Long Service Leave",ALLTRIM(loStaff.oData.myLslUnits),	.F.,;
						ALLTRIM(loStaff.oData.myLslUnits) + " Taken", "",;
						loStaff.oData.myLslTotal,;
						This.CheckRights("LV_LNG_BAL"), "Long Service Balance:", loStaff.oData.myLslUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* bereavement leave
				IF This.CheckRights("LR_TYPE_BEREAVEMENT") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"B", "BL", "Bereavement Leave","Hours",.T.,;
						"Hours To Pay",;
						"Days Taken",;
						0, .F.)
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* public holiday
				IF This.CheckRights("LR_TYPE_PUBLIC_HOLIDAY") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"P", "PH", "Public Holiday","Hours",.T.,;
						"Hours To Pay",;
						"Days Not Worked",;
						loStaff.oData.myAltTotal, .F.) && Not showing it here as we would need to ADD to the balance the units selected.
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* alternative leave accrued
				IF This.CheckRights("LR_TYPE_ALTERNATIVE_ACCRUED") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"Y", "AA", "Alternative Leave Accrued",		ALLTRIM(loStaff.oData.myAltUnits),	.F.,;
						"Days Worked", "",;
						loStaff.oData.myAltTotal, .F.)	&& Not showing it here as we would need to ADD to the balance the units selected.
					loCodes.Add(loCode, loCode.code)
				ENDIF

				* alternative leave paid
				IF This.CheckRights("LR_TYPE_ALTERNATIVE_PAID") OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"Z", "AP", "Alternative Leave Paid",		"Hours",							.T.,;
						"Hours To Pay",;
						"Days To Reduce Entitlement",;
						loStaff.oData.myAltTotal,;
						This.CheckRights("LV_ALT_BAL"), "Alternative Leave Balance:", loStaff.oData.myAltUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF

		        * unpaid leave
				IF (This.CheckRights("LR_UNPAID_LEAVE") OR tlUnrestricted)
					loCode = This.NewLeaveCodeObject(;
						"U", "UN", "Unpaid Leave", "Hours", .F.,;
						"Hours Taken", "",;
						0,;
						.F., "Unpaid Leave Balance:", "")
					loCodes.Add(loCode, loCode.code)
				ENDIF
				

				* shift leave
				IF (This.CheckRights("LR_TYPE_SHIFT") AND !EMPTY(loStaff.oData.myShDate)) OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"F", "SH", ALLTRIM(loStaff.oData.mySHName),	ALLTRIM(loStaff.oData.myShUnits),	.F.,;
						ALLTRIM(loStaff.oData.myShUnits) + " Taken", "",;
						IIF(This.CheckRights("LV_SHIFT_ACCRUED"), loStaff.oData.myShTotal, loStaff.oData.myShTotal - loStaff.oData.myShAccrue),;
						This.CheckRights("LV_SFT_BAL"), ALLTRIM(loStaff.oData.myShName) + " Balance:", loStaff.oData.myShUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF
	
				* other leave
				IF (This.CheckRights("LR_TYPE_OTHER") AND !EMPTY(loStaff.oData.myOtDate)) OR tlUnrestricted
					loCode = This.NewLeaveCodeObject(;
						"T", "OT", ALLTRIM(loStaff.oData.myOTName),	ALLTRIM(loStaff.oData.myOtUnits),	.F.,;
						ALLTRIM(loStaff.oData.myOtUnits) + " Taken", "",;
						IIF(This.CheckRights("LV_OTHER_ACCRUED"), loStaff.oData.myOtTotal, loStaff.oData.myOtTotal - loStaff.oData.myOtAccrue),;
						This.CheckRights("LV_OTH_BAL"), ALLTRIM(loStaff.oData.myOtName) + " Balance:", loStaff.oData.myOtUnits)
					loCodes.Add(loCode, loCode.code)
				ENDIF
			ELSE
				loCode = This.NewLeaveCodeObject(;
						"O", "AL", "Annual Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"S", "SL", "Sick Leave","Hours",.T.,;
						"Hours Taken",;
						"Days Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
					"P", "PH", "Public Holiday","Hours",.T.,;
					"Hours To Pay",;
					"Days Not Worked",;
					0, .F.) && Not showing it here as we would need to ADD to the balance the units selected.
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"N", "LSL", "Long Service Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)
				
				loCode = This.NewLeaveCodeObject(;
						"B", "BL", "Bereavement Leave","Hours",.T.,;
						"Hours To Pay",;
						"Days Taken",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"Y", "AA", "Alternative Leave Accrued","Units",.T.,;
						"Days Worked",;
						"Days Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"Z", "AP", "Alternative Leave Paid","Hours",.T.,;
						"Hours To Pay",;
						"Days To Reduce Entitlement",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"F", "SH", "Shift Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)

				loCode = This.NewLeaveCodeObject(;
						"T", "OT", "Other Leave","Units",.T.,;
						"Units Taken",;
						"Units Not Worked",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)
				
				loCode = This.NewLeaveCodeObject(;
						"U", "UN", "Unpaid Leave","Hours",.F.,;
						"Hours Taken",;
						"",;
						0, .F.)
				loCodes.Add(loCode, loCode.code)
				
			ENDIF
		ENDCASE

		RETURN loCodes
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION GetLeaveCode(tcCode AS String, tnEmployee AS Integer, tlUnrestricted AS Boolean)
		LOCAL loCodes
		loCodes = This.GetLeaveCodes(tnEmployee, tlUnrestricted)
		IF !EMPTY(loCodes.GetKey(tcCode))
			RETURN loCodes.Item(tcCode)
		ELSE
			RETURN null
		ENDIF
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetWageCodes() AS Collection
		LOCAL loCodes, loCode

		loCodes = CREATEOBJECT("COLLECTION")

		IF !This.SelectData(This.Licence, "wageType")
			RETURN null
		ENDIF

		SELECT wageType
		SCAN FOR !((UPPER(ALLTRIM(wageType.name)) == "UNDEFINED") OR wageType.hide)
			loCode = This.NewCodeObject(TRANSFORM(wageType.code), ALLTRIM(wageType.name))
			IF ISNULL(loCode)
				RETURN null
			ENDIF
			loCodes.Add(loCode, loCode.code)
		ENDSCAN

		RETURN loCodes
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetCostCentres() AS Collection
		LOCAL loCodes, loCode

		IF !This.SelectData(This.Licence, "costCent")
			RETURN null
		ENDIF

		loCodes = CREATEOBJECT("COLLECTION")

		loCode = This.NewCodeObjectHide('0', "0 Employee Default","N")
		loCodes.Add(loCode, loCode.code)

		SELECT code, name, hide FROM costCent INTO CURSOR curCC ORDER BY code
		SELECT curCC
		SCAN
			loCode = This.NewCodeObjectHide(TRANSFORM(code), TRANSFORM(code) + ' ' + name, IIF(hide,"Y","N"))	&&NOTE: name prefixed with code here.
			loCodes.Add(loCode, loCode.code)
		ENDSCAN
		USE IN SELECT("curCC")

		RETURN loCodes
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetAllowanceCodes() AS Collection
		LOCAL loCodes, loCode

		IF !This.SelectData(This.Licence, "allow")
			RETURN null
		ENDIF

		loCodes = CREATEOBJECT("COLLECTION")

		SELECT code, name, calc_how, amount, cost_cent, hide FROM allow INTO CURSOR curAllow ORDER BY code
		SELECT curAllow
		SCAN
			loCode = This.NewAllowanceObject(TRANSFORM(code), name, calc_how, amount, cost_cent, IIF(hide,"Y","N"))
			loCodes.Add(loCode, loCode.code)
		ENDSCAN
		USE IN SELECT("curAllow")

		RETURN loCodes
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetOtherCodes() AS Collection
		LOCAL loCodes, loCode, loStaff

		loStaff = Factory.GetStaffObject()
		IF !loStaff.Load(This.Employee)
			This.AddError("Cannot load OtherCodes!")
			RETURN null
		ENDIF

		loCodes = CREATEOBJECT("COLLECTION")

		loCode = This.NewCodeObject('D', ALLTRIM(loStaff.oData.myHpUnits) + " Paid for Holiday Pay")
		loCodes.Add(loCode, loCode.code)

		IF !This.IsAustralia()
			loCode = This.NewCodeObject('R', ALLTRIM(loStaff.oData.myHpUnits) + " Paid for Relevant Daily Rate")
			loCodes.Add(loCode, loCode.code)
		ENDIF

		RETURN loCodes
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetOpenPayCount() AS Integer
		LOCAL lnCount, lcAlias

		lcAlias = ALIAS()

		lnCount = 0
		IF This.SelectData(This.Licence, "myPays")
			SELECT *;
				FROM myPays;
				WHERE pay_type == 2 AND pay_status == 1;
				INTO CURSOR curOpenPays
			lnCount = RECCOUNT("curOpenPays")
			USE IN SELECT("curOpenPays")
		ENDIF

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN lnCount
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION GetOpenTemplateCount() AS Integer
		LOCAL lnCount, lcAlias

		lcAlias = ALIAS()

		lnCount = 0
		IF This.SelectData(This.Licence, "myPays")
			SELECT * FROM myPays WHERE pay_type = 1 AND pay_status = 1 INTO CURSOR curOpenTemplates
			lnCount = RECCOUNT("curopentemplates")
			USE IN SELECT("curopentemplates")
		ENDIF

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN lnCount
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort
	FUNCTION GetDefaultCostCentre(tnEmployee, tcType) AS Integer
		LOCAL loStaff, lcXML, lcCode

		loStaff = Factory.GetStaffObject()
		IF !loStaff.Load(tnEmployee)
			RETURN null
		ELSE
			lcXML = STREXTRACT(loStaff.oData.myXML, "<DefaultCostCentres>", "</DefaultCostCentres>", 1, 1)

			lcCode = STREXTRACT(lcXML, "<EmployeeDefault>", "</EmployeeDefault>", 1, 1)

			RETURN FLOOR(VAL(lcCode))
		ENDIF
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetLeaveMessages(tnLeaveReqId)
		LOCAL loMessages, lcAlias

		loMessages = null

		lcAlias = ALIAS()

		IF This.SelectData(This.Licence, "leaveRequestStatus")
			loMessages = CREATEOBJECT("COLLECTION")
			ADDPROPERTY(loMessages, "unreadCount", 0)

			SELECT * FROM leaveRequestStatus WHERE leaveReqId = tnLeaveReqId INTO CURSOR curLeaveMessages
			SELECT curLeaveMessages
			SCAN
				loMessages.Add(This.NewLeaveMessageObject(id, from, to, fromName, toName, subject, message, read, sent))

				IF EMPTY(read) AND to == This.Employee
					loMessages.unreadCount = loMessages.unreadCount + 1
				ENDIF
			ENDSCAN

			USE IN SELECT("curLeaveMessages")
		ENDIF

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN loMessages
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	*> sort
	FUNCTION GetManagersCount() AS Integer
		LOCAL lnCount

		lnCount = -1
		IF This.SelectData(This.Licence, "myManage")
			SELECT DISTINCT maMyStaff FROM myManage INTO CURSOR curManagersCount
			lnCount = _TALLY
			USE IN SELECT("curManagersCount")
		ENDIF

		RETURN lnCount
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort
	FUNCTION GetGroupsForManager(tnManager AS Integer, tcResultCursor AS String) AS Boolean
		LOCAL llOk as Boolean
		LOCAL lcAlias as String

		llOk = .F.
		lcAlias = EVL(tcResultCursor, "curGroups")

		IF This.SelectData(This.Licence, "myGroups") AND This.SelectData(This.Licence, "myManage")
			SELECT grCode, grName;
				FROM myGroups JOIN myManage ON myGroups.grCode = myManage.maMyGroups;
				WHERE myManage.maMyStaff = tnManager;
				ORDER BY grName;
				INTO CURSOR (lcAlias)
			llOk = .T.
		ENDIF

		RETURN llOk
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*@ GetDefaultGroup:
	 * Parameters:
	 *	tnManager:	The employee we are finding the default group for.
	 * Returns:
	 *	The default group for this manager - specifically, the All Employees group if they have it, or the first group they do have if any, else the My Details "group".
	FUNCTION GetDefaultGroup(tnManager AS Integer) AS Integer
		LOCAL lnResult

		* Get the list of groups for the specified manager
		IF !This.GetGroupsForManager(tnManager, "curDefaultGroup")
			This.AddError("Can't get default group for manager.")
			lnResult = MY_DETAILS_GROUP
		ELSE
			* Always default to my details as it is always present
			lnResult = MY_DETAILS_GROUP
			SELECT curDefaultGroup
			SCAN FOR UPPER(ALLTRIM(grName)) <> "ALL EMPLOYEES"
				* If we haven't already, assign the current group as the default (i.e. grab the first one we see, if any)
				IF lnResult == MY_DETAILS_GROUP
					lnResult = grCode
				ENDIF
			ENDSCAN
		ENDIF

		RETURN lnResult
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort; test
	FUNCTION GetEmployeesByGroupCode(tnGroupCode AS Integer, tcResultCursor AS String) AS Boolean
		LOCAL llOK as Boolean
		LOCAL lcAlias as String

		llOk = .F.
		lcAlias = EVL(tcResultCursor, "curStaff")

		IF This.SelectData(This.Licence, "myStaff") AND This.SelectData(This.Licence, "myTeams")
			IF tnGroupCode == MY_DETAILS_GROUP
				* Fake group representing "myself"
				* CF 14/10/2008; added the same fields selected below here, so that both branches return the same cursor structure..
				SELECT DISTINCT myStaff.myWebCode, ALLTRIM(myStaff.mySurname) + ", " + ALLTRIM(myStaff.myName) AS fullName, LOWER(mySurname) AS mySurname, LOWER(myName) AS myName;
					FROM myStaff WHERE myWebCode = This.Employee;
					INTO CURSOR (lcAlias) READWRITE
			ELSE
				* Normal group code supplied
				* CM added check for 0 paycode to remove payrolluser bug where zero leave balances are displayed for an employee
				* that happens when the payroll user is included as a member of the group.
				SELECT DISTINCT myStaff.myWebCode, ALLTRIM(myStaff.mySurname) + ", " + ALLTRIM(myStaff.myName) AS fullName, LOWER(mySurname) AS mySurname, LOWER(myName) AS myName;
					FROM myStaff JOIN myTeams ON myStaff.myWebCode = myTeams.tmMyStaff;
					WHERE myTeams.tmMyGroups = tnGroupCode AND myWebCode <> 1 AND myPayCode <> 0;
					INTO CURSOR (lcAlias) READWRITE;
					ORDER BY 3, 4
			ENDIF
			llOk = .T.
		ENDIF

		RETURN llOk
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort
	* Puts a list of Employees into the given cursor, excluding the superUser.
	FUNCTION GetEmployees(tcResultCursor AS String, tlIncludeCurrent AS Boolean) AS Boolean
		IF This.SelectData(This.Licence, "myStaff")
			SELECT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName, LOWER(mySurname) AS mySurname, LOWER(myName) AS myName;
				FROM myStaff;
				WHERE (tlIncludeCurrent OR myWebCode != This.Employee) AND myWebCode != 1;
				INTO CURSOR (tcResultCursor);
				ORDER BY 3, 4

			RETURN .T.
		ENDIF

		RETURN .F.
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort; comment; test!
	* Puts a list of ??? into the given cursor
	FUNCTION GetGroupsByEmployeeCode(tnEmployee AS Integer, tcResultCursor AS String) AS Boolean
		IF This.SelectData(This.Licence, "myGroups");
		  AND This.SelectData(This.Licence, "myTeams")
			SELECT myGroups.* FROM myGroups JOIN myTeams ON grCode = tmMyGroups;
				ORDER BY grname;
				INTO CURSOR (tcResultCursor);		&& was hardcoded as: "curGroups"
				WHERE tmMyStaff = tnEmployee

			RETURN .T.
		ENDIF

		RETURN .F.
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> sort
	* Puts a list of payroll users that have the specified employee in one of their groups into the given cursor
	FUNCTION GetPayrollUsersByEmployeeCode(tnEmployee AS Integer, tcResultCursor AS String, tnExcludeID AS Integer) AS Boolean
		IF This.SelectData(This.Licence, "myTeams");
		  AND This.SelectData(This.Licence, "myManage");
		  AND This.SelectData(This.Licence, "myStaff")
			SELECT DISTINCT maMyStaff FROM myTeams JOIN myManage ON tmMyGroups = maMyGroups;
				WHERE tmMyStaff = tnEmployee;
				INTO CURSOR curManagersX

			* Now construct a list of names, codes etc.
			IF EMPTY(tnExcludeID)
				SELECT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName, LOWER(mySurname), LOWER(myName);
					FROM myStaff;
					JOIN curManagersX ON myWebCode = maMyStaff;
					INTO CURSOR (tcResultCursor);
					ORDER BY 3, 4
			ELSE
				SELECT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName, LOWER(mySurname), LOWER(myName);
					FROM myStaff JOIN curManagersX ON myWebCode = maMyStaff;
					WHERE myWebCode != tnExcludeID;
					INTO CURSOR (tcResultCursor);
					ORDER BY 3, 4
			ENDIF

			USE IN SELECT("curManagersX")

			RETURN .T.
		ENDIF

		RETURN .F.
	ENDFUNC

	*################################################################################*
#DEFINE TOC_Queries_

	*> +define: Queries
	FUNCTION CheckAccessForManager(tnStaffCode AS Integer, tnManager AS Integer)
		IF !(VARTYPE(tnManager) == 'N')
			tnManager = This.Employee
			IF !(tnStaffCode == tnManager OR This.IsManager(tnManager))
				RETURN .F.
			ENDIF
		ENDIF

		IF This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "myTeams");
		  AND This.SelectData(This.Licence, "myManage")
			IF tnStaffCode == This.Employee
				RETURN .T.
			ENDIF

			SELECT DISTINCT myStaff.myWebCode;
				FROM myStaff JOIN myTeams ON myStaff.myWebCode = myTeams.tmMyStaff;
				JOIN myManage ON myTeams.tmMyGroups = myManage.maMyGroups;
				WHERE myManage.maMyStaff = tnManager;
				AND myStaff.myWebCode == tnStaffCode;
				INTO CURSOR curCheckAccess
			USE IN SELECT("curCheckAccess")

			IF _TALLY > 0
				RETURN .T.
			ENDIF
		ENDIF

		RETURN .F.
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION CheckAccess(tnCurrentStaff AS Integer, tlManager AS Boolean, tlAllowEveryone AS Boolean)
		IF !(tlManager OR tnCurrentStaff == This.Employee);
		  OR (tlManager AND !(tlAllowEveryone AND tnCurrentStaff == EVERYONE_OPTION OR This.CheckAccessForManager(tnCurrentStaff)))
			This.AddError("You do not have access to this page.")
			This.AddUserInfo("Your attempt to access has been logged.")		&& Strictly, this is only true if the following Load() works, but we might as well scare them...

			LOCAL loIPStuff, loStaff, lcSubject, lcMessage

			lcWorkArea = SELECT()
			loStaff = Factory.GetStaffObject()

			*loIPStuff = CREATEOBJECT("wwIPStuff")
			lcSubject = "Attempted Illegal Access"
			lcMessage = "  User: " + TRANSFORM(This.Employee) + CRLF
			IF loStaff.Load(This.Employee)
				lcMessage = lcMessage + "  Name: " + loStaff.FullName + CRLF
			ENDIF
			lcMessage = lcMessage + "  Company: " + TRANSFORM(This.Licence) + CRLF
			lcMessage = lcMessage + "  Page: *" + SUBSTR(Request.GetCurrentURL(.F.), 5) + CRLF	&& using HTTP, but stripping it anyway
			lcMessage = lcMessage + "  Attempted Access User: " + TRANSFORM(tnCurrentStaff) + CRLF
			IF tnCurrentStaff != This.Employee AND loStaff.Load(tnCurrentStaff)
				lcMessage = lcMessage + "  Attempted Access Name: " + loStaff.FullName + CRLF
			ENDIF

			*loIPStuff.SendMailAsync(;
				*AppSettings.Get("mailserver"),;
				*AppSettings.Get("updateName"),;
				*AppSettings.Get("update_adr"),;
				*"admin@mystaffinfo.com",;			--magic!
				*"", lcSubject, lcMessage, "", "";
			*)

			lcMessage = "----------------------------------" + CRLF + lcSubject + ':' + CRLF + lcMessage
			STRTOFILE(lcMessage, This.CompanyDataPath() + "illegalAccesses.txt", .T.)

			RETURN .F.
		ENDIF

		RETURN .T.
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment!, test
	FUNCTION CheckMessageAccess(tnMessageID AS Integer) AS Boolean
		LOCAL loMessages, llOK

		llOK = .F.

		loMessages = Factory.GetMessagesObject()
		IF loMessages.Load(tnMessageID)
			IF EMPTY(loMessages.oData.meToID) OR loMessages.oData.meToID == This.Employee OR loMessages.oData.meFromID == This.Employee	&& new or news message, or to or from me...
				llOK = .T.
			ENDIF
		ENDIF

		RETURN llOK
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION IsAdminUser(tnEmployee AS Integer) AS Boolean
		LOCAL loStaff as Object
		LOCAL llAdmin as Boolean

		IF INLIST(tnEmployee, 0, -999)	&& Logged in via fat client or dev user
			RETURN .T.
		ENDIF

		llAdmin = .F.
		loStaff = Factory.GetStaffObject()
		IF loStaff.Load(tnEmployee)
			llAdmin = loStaff.oData.myAdmin
		ENDIF

		RETURN llAdmin
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION IsManager(tnEmployee AS Integer) AS Boolean
		* CM 16/08/2005 Using original method now
		* User is a manager if they are looged by fat client
		* or the user code is marked as a manager
		LOCAL llManager as Boolean

		DO CASE
		CASE tnEmployee = 0
			* Remote administration - head office application
			llManager = .T.
		CASE tnEmployee = -999
			* Development login - elevated rights
			llManager = .T.
		OTHERWISE
			* Check employee record
			llManager = .F.
			IF This.SelectData(This.Licence, "myStaff")
				IF SEEK(tnEmployee, "myStaff", "myWebCode")
					llManager = myStaff.myManager
				ENDIF
			ENDIF
		ENDCASE

		RETURN llManager
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION CheckRights(tcArea AS String) AS Boolean
		RETURN This.CheckRightsFor(This.Employee, tcArea)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> test!
	FUNCTION CheckRightsFor(tnEmployee, tcArea)
		* CM add basic user security functionality to the system
		LOCAL llAllowed, lnWorkArea, loRights, lcFile, lcAlias

		lcAlias = ALIAS()

		llAllowed = .T.
		tcArea = LOWER(ALLTRIM(tcArea))
		DO CASE
		CASE tnEmployee = -999	&& dev user - allowed everywhere
		CASE tcArea == "main"	&& main section - always allowed...
			IF tnEmployee == 1		&& ...except for super user account
				llAllowed = .F.
			ENDIF
		CASE tcArea == "logout"	&& always allowed

		CASE tcArea == "admin"
			llAllowed = This.IsAdminUser(tnEmployee)

		CASE tcArea == "maint"
			llAllowed = This.IsAdminUser(tnEmployee) OR This.IsManager(tnEmployee)

		CASE tcArea == "manager"
			llAllowed = This.IsManager(tnEmployee)

		OTHERWISE
			llAllowed = .F.

			* This.SelectData(This.Licence, "myStaff")	BUG: not checking for success!	--not needed as the following opens the table where needed anyway..?
			loRights = NEWOBJECT("SecurityFacade", "ComaccOnlineSecurity.vcx")
			llAllowed = loRights.Get(tnEmployee, UPPER(tcArea))

			IF tcArea == "timesheet"
				llAllowed = llAllowed AND FILE(This.CompanyDataPath() + "timesheet.mem")
			ENDIF
		ENDCASE

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN llAllowed
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION IsPayrollUser(tnWebCode AS Integer)
		LOCAL loStaff, llPayrollUser

		loStaff = Factory.GetStaffObject()
		IF loStaff.Load(tnWebCode)
			llPayrollUser = (loStaff.oData.mypaycode == 0 AND loStaff.oData.myWebCode > 1)
		ENDIF

		RETURN llPayrollUser
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Does this copany have access to this module?
	FUNCTION HasModule(tcModuleName AS String) AS Boolean
		RETURN FILE(This.CompanyDataPath() + FORCEEXT(ALLTRIM(tcModuleName), "mem"))
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> test
	FUNCTION IsAustralia()
		LOCAL llAustralia, lcAlias

		lcAlias = ALIAS()

		IF This.SelectData(This.Licence, "myStaff")
			GO TOP IN myStaff
			llAustralia = (LOWER(ALLTRIM(myStaff.myCountry)) == "australia")
		ENDIF

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN llAustralia
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> test
	PROTECTED FUNCTION IsTemplatePay(tnPayNumber AS Integer) AS Boolean
		LOCAL llTemplatePay as Boolean, lcAlias

		lcAlias = ALIAS()

		llTemplatePay = .F.

		IF This.SelectData(This.Licence, "myPays") AND SEEK(tnPayNumber, "myPays", "pay_pk")
			llTemplatePay = (myPays.pay_type == 1)
		ENDIF

		IF !EMPTY(lcAlias)
			SELECT (lcAlias)
		ENDIF

		RETURN llTemplatePay
	ENDFUNC

	*################################################################################*
#DEFINE TOC_Data_

	*> +define: Data
	* Returns a path already ending in a '\'
	FUNCTION CompanyDocumentsPath()
		* full path to the database directory... (payslips are under here, etc.)
		RETURN ADDBS(This.CompanyDataPath() + "Documents")
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Returns a path already ending in a '\'
	FUNCTION CompanyDataPath()
		* full path to the database directory... (payslips are under here, etc.)
		RETURN ADDBS(This.cDataPath) + ADDBS(ALLTRIM(TRANSFORM(This.Licence)))
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Returns a path already ending in a '\'
	FUNCTION CompanyHtmlPath()
		* full path to the docroot directory... (scripts are here, and company docs URI's are relative to this)
		DO CASE
		CASE "devserver" $ LOWER(SYS(0))
			* Test/Staging Server
			RETURN ADDBS(This.cHtmlPagePath)
		CASE VERSION(2) = 0 && This.oRequest.IsLinkSecure()
			* Live server
			RETURN STRTRAN(ADDBS(This.cHtmlPagePath), "http:", "https:", 1, 1, 1)
		OTHERWISE
			* Development machine, probably running file mode in VFP
			RETURN ADDBS(This.cHtmlPagePath)
		ENDCASE
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Path to where the templates (HTML pages) are stored.
	* Returns a path already ending in a '\'
	FUNCTION GetTemplatePagePath()
		RETURN ADDBS(This.CompanyHtmlPath() + "Templates")
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION ListDocuments(tcIndent, tcPath, tcParent) As String
		LOCAL lcPath, lnFiles, lnFile, lcItem, lcName, lcParent, lcOutput, lnI, lcJoiner
		LOCAL ARRAY laFiles[1, 5]

		IF EMPTY(tcPath)
			* first entrance
			tcPath = FULLPATH(This.CompanyDocumentsPath())
			tcParent = "0"
		ENDIF

		lcOutput = "";

		lcPath = ADDBS(tcPath) + "*.*"
		lcParent = EVL(tcParent, "root")

		lnFiles = ADIR(laFiles, lcPath, 'D')

		* Sort so directories are at the top:
		** Prepend a type char:
		FOR lnFile = 1 TO lnFiles
			IF 'D' $ laFiles(lnFile, 5)
				laFiles(lnFile, 1) = 'D' + laFiles(lnFile, 1)
			ELSE
				laFiles(lnFile, 1) = 'F' + laFiles(lnFile, 1)
			ENDIF
		ENDFOR

		ASORT(laFiles, 1, -1, 0, 1)

		** Strip the type prefix
		FOR lnFile = 1 TO lnFiles
			laFiles(lnFile, 1) = RIGHT(laFiles(lnFile, 1), LEN(laFiles(lnFile, 1)) - 1)
		ENDFOR

		lnI = 0
		lcJoiner = ""
		FOR lnFile = 1 TO lnFiles
			* Make sure we do not recurse up the structure
			IF INLIST(laFiles(lnFile, 1), '.', "..")
				LOOP
			ENDIF

			* Create a name for the item
			lcItem = tcParent + '_' + TRANSFORM(lnI)
			lnI = lnI + 1

			* Is this a folder?
			IF 'D' $ laFiles(lnFile, 5)
				lcOutput = lcOutput + lcJoiner + tcIndent + '["' +PROPER(laFiles(lnFile, 1)) + '", "' + lcItem + '", [' + CRLF
				lcOutput = lcOutput + This.ListDocuments(tcIndent + CHR(9), ADDBS(tcPath) + laFiles(lnFile, 1), lcItem)
				lcOutput = lcOutput + tcIndent + ']]'
			ELSE
				lcStrip = LOWER(FULLPATH(This.CompanyDocumentsPath()))
				lcPath = LOWER(FULLPATH(tcPath))
				lcLink = ADDBS(STRTRAN(lcPath, lcStrip, ""))
				lcOutput = lcOutput + lcJoiner + tcIndent + '["' +PROPER(laFiles(lnFile, 1)) + '", "GetDocument.si?doc=' + This.URLEscape(STRTRAN(lcLink + PROPER(laFiles(lnFile, 1)), '\', '/')) + '"]'
			ENDIF
			lcJoiner = ',' + CRLF
		ENDFOR

		RETURN lcOutput + CRLF
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION GetTimesheetTypes(tlSelectData, tnCurrentPay, tnCurrentGroup, tnCurrentStaff, tcFilter) AS Object
		LOCAL lcFilter2, loTypes, llAusie, loLeaveCodes, lnI, llRates, lcEmpWhere

		llAusie = This.IsAustralia()

		IF tlSelectData AND !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Cannot get timesheet types!")
			RETURN null
		ENDIF

		TRY
			SELECT myStaff
			LOCATE FOR myRates
			llRates = FOUND()
		CATCH
			llRates = .F.
		ENDTRY

		IF !tlSelectData
			tcFilter = ""
		ENDIF

		loTypes = CREATEOBJECT("COLLECTION")
		
		LOCAL lnBOUNDARY as Integer
		lnBOUNDARY = MAX(WEBCODE_PAYROLLUSER_BOUNDARY, 2)
		
		IF tlSelectData AND tnCurrentStaff == EVERYONE_OPTION
			IF !This.GetEmployeesByGroupCode(tnCurrentGroup, "curGroupStaff")
				This.AddError("Cannot load current group!")
				RETURN loTypes	&& bail out on error
			ENDIF
			lcEmpWhere = "tsEmp in (Select myWebCode from curGroupStaff Where myWebCode >= lnBOUNDARY)"
		*!*	lcEmpWhere = "INLIST(tsEmp"						&& 01/07/2010  CMGM  TTP5692  Error is caused by INLIST: can only take 25 expr including the search expr
*!*				lcEmpWhere = ""									&& 01/07/2010  CMGM  TTP5692  Replace it by native SQL IN()

*!*				SELECT curGroupStaff
*!*				SCAN
*!*					IF curGroupStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curGroupStaff.myWebCode < 2	&&NOTE: always hiding PayrollUsers here; hide Admin user either way...
*!*						LOOP
*!*					ENDIF

*!*					lcEmpWhere = lcEmpWhere + "," + TRANSFORM(curGroupStaff.myWebCode)
*!*				ENDSCAN
*!*				
*!*				IF EMPTY(_TALLY)
*!*					lcEmpWhere = ".T."								&& 25/02/2011  CMGM  2011.02  TTP6615  Fix incorrect sql "tsEmp IN ()" below due to PayrollUsers being hidden
*!*				ELSE
*!*				*!*	lcEmpWhere = lcEmpWhere + ")"					&& 01/07/2010  CMGM  TTP5692  
*!*					lcEmpWhere = SUBSTR(lcEmpWhere, 2)				&& 01/07/2010  CMGM  TTP5692  Remove the initial ","
*!*					lcEmpWhere = "tsEmp IN (" + lcEmpWhere + ")"	&& 01/07/2010  CMGM  TTP5692  Build SQL
*!*				ENDIF
			
		ELSE
			lcEmpWhere = "tsEmp == tnCurrentStaff"
		ENDIF

		&&LATER: add J to the below when a JobCode is needed...
		IF This.CheckRights("TS_TIMESHEET_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere. ;
					AND tsPay == tnCurrentPay ;
					AND tmId == 0 ;
					AND tsType == 'M';
					&tcFilter.;
					INTO CURSOR curTimes;
					ORDER BY tsDate, tsStart
			ENDIF

			loTypes.Add(This.NewTimesheetTypeObject("time", 'M', "Timesheet", [.CheckRights("TS_TIMESHEET_V")], IIF(llRates, "DSFBuWAC", "DSFBuWC"), "curTimes", _TALLY), "time")
		ENDIF

		IF This.CheckRights("TS_WAGES_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere.;
					AND tsPay == tnCurrentPay ;
					AND tmId == 0 ;
					AND tsType == 'W';
					&tcFilter.;
					INTO CURSOR curWages;
					ORDER BY tsDate, tsStart
			ENDIF
			loTypes.Add(This.NewTimesheetTypeObject("wages", 'W', "Wages", [.CheckRights("TS_WAGES_V")], IIF(llRates, "DUWAC", "DUWC"), "curWages", _TALLY), "wages")
		ENDIF

		IF This.CheckRights("TS_LEAVE_V")
			loLeaveCodes = This.GetLeaveCodes(IIF(!tlSelectData OR tnCurrentStaff == EVERYONE_OPTION, This.Employee, tnCurrentStaff))

			lcFilter2 = tcFilter + " AND INLIST(tsType"

			FOR lnI = 1 TO loLeaveCodes.Count
				lcFilter2 = lcFilter2 + ", '" + loLeaveCodes.Item(lnI).code + "'"
			NEXT

			lcFilter2 = lcFilter2 + ")"

			IF !("INLIST(tsType)" $ lcFilter2)
				IF tlSelectData
					SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere.;
					AND tsPay == tnCurrentPay ;
					AND tmId == 0 ;
					&lcFilter2.;
					INTO CURSOR curLeave;
					ORDER BY tsDate, tsStart
				ENDIF

				loTypes.Add(This.NewTimesheetTypeObject("leave", 'S', "Leave", [.CheckRights("TS_LEAVE_V")], "DTURC", "curLeave", _TALLY), "leave")	&& this doesn't need more authz as it's left out if no leaveTypes are available
			ENDIF
		ENDIF

		IF This.CheckRights("TS_ALLOWANCES_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE &lcEmpWhere.;
					AND tsPay == tnCurrentPay ;
					AND tmId == 0 ;
					AND tsType == 'A';
					&tcFilter.;
					INTO CURSOR curAllowances;
					ORDER BY tsDate, tsStart
			ENDIF

			loTypes.Add(This.NewTimesheetTypeObject("allowances", 'A', "Allowances", [.CheckRights("TS_ALLOWANCES_V")], "DKUC", "curAllowances", _TALLY), "allowances")
		ENDIF

		IF This.CheckRights("TS_OTHER_V")
			IF tlSelectData
				IF llAusie
					lcFilter2 = tcFilter + " AND INLIST(tsType, 'D')"
				ELSE
					lcFilter2 = tcFilter + " AND INLIST(tsType, 'R', 'D')"
				ENDIF

				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE &lcEmpWhere. ;
					AND tmId == 0 ;
					AND tsPay == tnCurrentPay;
					&lcFilter2.;
					INTO CURSOR curOther;
					ORDER BY tsDate, tsStart
			ENDIF

			loTypes.Add(This.NewTimesheetTypeObject("other", 'D', "Other", [.CheckRights("TS_OTHER_V")], "DOU", "curOther", _TALLY), "other")
		ENDIF

		RETURN loTypes
	ENDFUNC

	*================================================================================*

	PROCEDURE CollectTimeEntryFormData(toType, tcSuffix, tnCurrentStaff, tlManager, rnStaff, rdDate,;
				roLeaveCode, roOtherCode, roAllowCode, rtStart, rtEnd, rtBreak, rnUnits, rnReduce,;
				roWageCode, rnRateCode, roCostCentCode,roJobCode)

		LOCAL loCodes, lcValue

		IF !EMPTY(tcSuffix)
			tcSuffix = '_' + TRANSFORM(tcSuffix)
		ENDIF

		IF tlManager
			rnStaff = VAL(Request.Form("staff" + tcSuffix))
		ENDIF

		IF toType.showDate AND !toType.readOnlyDate
			rdDate = CTOD(Request.Form("date" + tcSuffix))
		ENDIF

		IF toType.showLeaveType AND !toType.readOnlyLeaveType
			lcValue = Request.Form("leaveType" + tcSuffix)
			loCodes = This.GetLeaveCodes(tnCurrentStaff)		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roLeaveCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showOtherType AND !toType.readOnlyOtherType
			lcValue = Request.Form("otherType" + tcSuffix)
			loCodes = This.GetOtherCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roOtherCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showCode AND !toType.readOnlyCode
			lcValue = Request.Form("code" + tcSuffix)
			loCodes = This.GetAllowanceCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roAllowCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showStart AND !toType.readOnlyStart
			rtStart = CTOT(Request.Form("start" + tcSuffix))
		ENDIF

		IF toType.showEnd AND !toType.readOnlyEnd
			rtEnd = CTOT(Request.Form("end" + tcSuffix))
		ENDIF

		IF toType.showBreak AND !toType.readOnlyBreak
			rtBreak = CTOT(Request.Form("break" + tcSuffix))
		ENDIF

		IF toType.showUnits AND !toType.readOnlyUnits
			* handle the override values on allowances where applicable as they will [should] not have been posted
			IF toType.id == "allowances" AND !ISNULL(roAllowCode) AND !EMPTY(roAllowCode.unitsValue)
				rnUnits = roAllowCode.unitsValue
			ELSE
				rnUnits = VAL(Request.Form("units" + tcSuffix))
			ENDIF
		ENDIF

		IF toType.showReduce AND !toType.readOnlyReduce
			IF !toType.showLeaveType OR roLeaveCode.enableReduce
				rnReduce = VAL(Request.Form("units2" + tcSuffix))
			ENDIF
		ENDIF

		IF toType.showWageType AND !toType.readOnlyWageType
			lcValue = Request.Form("wageType" + tcSuffix)
			loCodes = This.GetWageCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roWageCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showRateCode AND !toType.readOnlyRateCode
			rnRateCode = VAL(Request.Form("rateCode" + tcSuffix))
		ENDIF

		IF toType.showCostCent AND !toType.readOnlyCostCent
			* handle the override values on allowances where applicable as they will [should] not have been posted
			IF toType.id == "allowances" AND !ISNULL(roAllowCode) AND !EMPTY(roAllowCode.costCentre)
				lcValue = TRANSFORM(roAllowCode.costCentre)
			ELSE
				lcValue = Request.Form("costCentre" + tcSuffix)
			ENDIF
			loCodes = This.GetCostCentres()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roCostCentCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

*		IF toType.showJobCode AND !toType.readOnlyJobCode
*			lcValue = Request.Form("jobCode" + tcSuffix)
*			loCodes = This.GetJobCodes()		&& This handles the authz for us.
*			IF !EMPTY(loCodes.GetKey(lcValue))
*				roJobCode = loCodes.Item(lcValue)
*			ENDIF
*		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	FUNCTION SaveSingleTimeEntry(toType, tnId_nCurrentPay,;
									tlIsManger, tnStaff,;
									tdDate, toLeaveCode, toOtherCode,;
									toAllowCode, ttStart, ttEnd, ttBreak, tnUnits,;
									tnReduce, toWageCode, tnRateCode, toCostCentCode,;
									toJobCode,tnCurrentGroup, rnRowCount) AS Boolean

		LOCAL lnBreakLen, llError, loGroup, lnMaxUnitsValue

		*!* 16/11/2009;TTP4791;JCF: Added the no-limit effect to other as well, and centralised the check.
		*!* 19/11/2009;TTP4852,4854;JCF: altered max "uncapped" limit to avoid a rounding issue that caused an out of band value and hence a NumericOverflow error on import.
		lnMaxUnitsValue = IIF(INLIST(toType.id, "wages", "allowances", "other"), 999998.99, 24)

		rnRowCount = 1

		llError = .F.

		* Validate...
		IF tlIsManger
			IF EMPTY(tnStaff)
				This.AddValidationError("Employee not specified.")
				llError = .T.
			ELSE
				IF tnStaff == EVERYONE_OPTION
					IF !This.GetEmployeesByGroupCode(tnCurrentGroup, "curGroupStaff")
						This.AddError("Cannot load current group!")
						RETURN .F.	&& bail out on error
					ENDIF

					loGroup = CREATEOBJECT("COLLECTION")
					rnRowCount = 0
					SELECT curGroupStaff
					SCAN
						IF curGroupStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curGroupStaff.myWebCode < 2	&&NOTE: always hiding PayrollUsers here; hide Admin user either way...
							LOOP
						ENDIF

						rnRowCount = rnRowCount + 1
						loGroup.Add(curGroupStaff.myWebCode)
					ENDSCAN
				ELSE
					&&TODO: validate emp is in [manager's | selected] group
				ENDIF
			ENDIF
		ENDIF
		IF toType.showDate AND !toType.readOnlyDate
			IF EMPTY(tdDate)
				This.AddValidationError("Missing or invalid Date.")
				llError = .T.
			ELSE
				&&MAYBE: check it's within the current pay?
			ENDIF
		ENDIF
		IF toType.showLeaveType AND !toType.readOnlyLeaveType
			IF ISNULL(toLeaveCode)
				This.AddValidationError("Missing or invalid LeaveType.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showOtherType AND !toType.readOnlyOtherType
			IF ISNULL(toOtherCode)
				This.AddValidationError("Missing or invalid OtherType.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showCode AND !toType.readOnlyCode
			IF ISNULL(toAllowCode)
				This.AddValidationError("Missing or invalid Code.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showStart AND !toType.readOnlyStart
			IF EMPTY(ttStart)
				This.AddValidationError("Missing or invalid Start.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showEnd AND !toType.readOnlyEnd
			IF EMPTY(ttEnd)
				This.AddValidationError("Missing or invalid End.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showBreak AND !toType.readOnlyBreak
			IF EMPTY(ttBreak)
				This.AddValidationError("Missing or invalid Break.")
				llError = .T.
			ENDIF
		ENDIF
		IF !(EMPTY(ttStart) OR EMPTY(ttEnd) OR EMPTY(ttBreak))
			lnBreakLen = ttBreak - CTOT("00:00")

			IF ttEnd < ttStart
				* Crossed midnight
				ttEnd = ttEnd + 86400
			ENDIF

			tnUnits = ((ttEnd - ttStart) - lnBreakLen) / 3600
			IF tnUnits < 0 OR tnUnits > 24
				This.AddValidationError("Invalid Start/End/Break combination.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showUnits AND !toType.readOnlyUnits
			* Only validate if the field was enabled - in this case, for all types other than allowances, or for allowances where the selected code has the field enabled.
			IF !toType.showCode OR toAllowCode.enableUnits
				IF tnUnits < -lnMaxUnitsValue OR tnUnits > lnMaxUnitsValue	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					*!* 16/11/2009;TTP4791;JCF: Added the missing TRANSFORM() so the following line actually works; Centralised the logic in the same way as the template.
					This.AddValidationError("Units out of range " + TRANSFORM(-lnMaxUnitsValue) + " to " + TRANSFORM(lnMaxUnitsValue) + '.')
					llError = .T.
				ENDIF
			ENDIF
		ENDIF
		IF toType.showReduce AND !toType.readOnlyReduce
			* Only validate if the field was enabled - in this case, for all types other than leave, or for leave where the selected code has the field enabled.
			IF !toType.showLeaveType OR toLeaveCode.enableReduce
				IF tnReduce < -24 OR tnReduce > 24	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					This.AddValidationError("Reduce out of range -24 to 24.")
					llError = .T.
				ENDIF
			ELSE
				tnReduce = 0
			ENDIF
		ENDIF
		IF toType.showWageType AND !toType.readOnlyWageType
			IF ISNULL(toWageCode)
				This.AddValidationError("Missing or invalid WageType.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showRateCode AND !toType.readOnlyRateCode
			IF tnRateCode < 1 OR tnRateCode > 9
				This.AddValidationError("Invalid RateCode (not between 1 and 9 inclusive).")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showCostCent AND !toType.readOnlyCostCent
			IF ISNULL(toCostCentCode)
				This.AddValidationError("Missing or invalid CostCentre.")
				llError = .T.
			ENDIF
		ENDIF
		IF toType.showJobCode AND !toType.readOnlyJobCode
			IF ISNULL(toJobCode)
				This.AddValidationError("Missing or invalid JobCode.")
				llError = .T.
			ENDIF
		ENDIF

		IF !llError
			* Save...
			FOR lnI = 1 TO rnRowCount
				SELECT timesheet
				IF tnId_nCurrentPay < 0
					* ...new
					APPEND BLANK
				ELSE
					* ...existing
					LOCATE
					LOCATE FOR tsId == tnId_nCurrentPay
					IF !FOUND()
						This.AddError("Entry not found!")
						RETURN .F.
					ENDIF
				ENDIF

				replace tmId WITH 0
				
				IF tnId_nCurrentPay < 0
					* tsType will get corrected later if needed.
					REPLACE;
						tsType WITH toType.tsType;
						tsPay  WITH -tnId_nCurrentPay IN timesheet
				ENDIF

				IF tlIsManger OR tnId_nCurrentPay < 0
					* always save for new rows, and when manager for edited rows.
					IF tnStaff == EVERYONE_OPTION
						REPLACE tsEmp WITH loGroup.Item(lnI) IN timesheet
					ELSE
						REPLACE tsEmp WITH tnStaff IN timesheet
					ENDIF
				ENDIF
				IF toType.showDate AND !toType.readOnlyDate
					REPLACE tsDate WITH tdDate IN timesheet
				ENDIF
				IF toType.showLeaveType AND !toType.readOnlyLeaveType
					REPLACE tsType WITH toLeaveCode.code IN timesheet
				ENDIF
				IF toType.showOtherType AND !toType.readOnlyOtherType
					REPLACE tsType WITH toOtherCode.code IN timesheet
				ENDIF
				IF toType.showCode AND !toType.readOnlyCode
					REPLACE tsCode WITH VAL(toAllowCode.code) IN timesheet
				ENDIF
				IF toType.showStart AND !toType.readOnlyStart
					REPLACE;
						tsStart		WITH ttStart;
						tsFinish	WITH ttEnd;
						tsBreak		WITH ttBreak;
						tsUnits		WITH tnUnits;
						IN timesheet
				ENDIF
				IF toType.showUnits AND !toType.readOnlyUnits
					REPLACE tsUnits WITH tnUnits IN timesheet
				ENDIF
				IF toType.showReduce AND !toType.readOnlyReduce
					REPLACE tsUnits2 WITH tnReduce IN timesheet
				ENDIF
				IF toType.showWageType AND !toType.readOnlyWageType
					REPLACE tsWageType WITH VAL(toWageCode.code) IN timesheet
				ENDIF
				IF toType.showRateCode AND !toType.readOnlyRateCode
					REPLACE tsRateCode WITH tnRateCode IN timesheet
				ENDIF
				IF toType.showCostCent AND !toType.readOnlyCostCent
					REPLACE tsCostCent WITH VAL(toCostCentCode.code) IN timesheet
				ENDIF
				IF toType.showJobCode AND !toType.readOnlyJobCode
					REPLACE tsXXXX WITH VAL(toJobCode.code) IN timesheet	&&LATER: Add correct dbFieldName...
				ENDIF
			NEXT

			IF tnId_nCurrentPay >= 0
				This.AddUserInfo(toType.title + " Entry Saved.")
			ENDIF
		ENDIF

		RETURN !llError
	ENDFUNC

	*--------------------------------------------------------------------------------*

	PROCEDURE RetiainTimeEntry(;
		toType, toRetainList, tcSuffix,;
		tlIsManger, tnStaff,;
		tdDate, toLeaveCode, toOtherCode, toAllowCode, ttStart, ttEnd, ttBreak, tnUnits, tnReduce, toWageCode, toCostCentCode, toJobCode;
	)
		LOCAL loCodes, lcValue

		IF !EMPTY(tcSuffix)
			tcSuffix = '_' + TRANSFORM(tcSuffix)
		ENDIF

		IF tlIsManger
			toRetainList.SetEntry("staff" + tcSuffix, TRANSFORM(tnStaff))
		ENDIF
		IF toType.showDate AND !toType.readOnlyDate
			toRetainList.SetEntry("date" + tcSuffix, TRANSFORM(tdDate))
		ENDIF
		IF toType.showLeaveType AND !toType.readOnlyLeaveType AND !ISNULL(toLeaveCode)
			toRetainList.SetEntry("leaveType" + tcSuffix, toLeaveCode.code)
		ENDIF
		IF toType.showOtherType AND !toType.readOnlyOtherType AND !ISNULL(toOtherCode)
			toRetainList.SetEntry("otherType" + tcSuffix, toOtherCode.code)
		ENDIF
		IF toType.showCode AND !toType.readOnlyCode AND !ISNULL(toAllowCode)
			toRetainList.SetEntry("code" + tcSuffix, toAllowCode.code)
		ENDIF
		IF toType.showStart AND !toType.readOnlyStart
			toRetainList.SetEntry("start" + tcSuffix, TRANSFORM(ttStart))
		ENDIF
		IF toType.showEnd AND !toType.readOnlyEnd
			toRetainList.SetEntry("end" + tcSuffix, TRANSFORM(ttEnd))
		ENDIF
		IF toType.showBreak AND !toType.readOnlyBreak
			toRetainList.SetEntry("break" + tcSuffix, TRANSFORM(ttBreak))
		ENDIF
		IF toType.showUnits AND !toType.readOnlyUnits
			toRetainList.SetEntry("units" + tcSuffix, TRANSFORM(tnUnits))
		ENDIF
		IF toType.showReduce AND !toType.readOnlyReduce
			IF !toType.showLeaveType OR roLeaveCode.enableReduce
				toRetainList.SetEntry("units2" + tcSuffix, TRANSFORM(tnReduce))
			ENDIF
		ENDIF
		IF toType.showWageType AND !toType.readOnlyWageType AND !ISNULL(toWageCode)
			toRetainList.SetEntry("wageType" + tcSuffix, toWageCode.code)
		ENDIF
		IF toType.showCostCent AND !toType.readOnlyCostCent AND !ISNULL(toCostCentCode)
			toRetainList.SetEntry("costCentre" + tcSuffix, toCostCentCode.code)
		ENDIF
		IF toType.showJobCode AND !toType.readOnlyJobCode AND !ISNULL(toJobCode)
			toRetainList.SetEntry("jobCode" + tcSuffix, toJobCode.code)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*
	* 17/05/2010  CMGM  MRD 4.2.2.2  New function to copy the selected previous pay timesheet entries into
	*                                the selected current pay.
	* 20/04/2012  RAJ   CARD 4878,4880,4881,4882 - Added new Staff Selectors
	
	PROCEDURE LoadPreviousPayTimesheetToCurrentPayTimesheet()
		LOCAL lnPreviousPay, lnCurrentPay, lnCurrentStaff, llAllOk, llManager
		LOCAL lnDaysDiff, ldNewDate, ldCurrentPay, ldPreviousPay, lnCurrentGroup

		lnDaysDiff = 0
		ldNewDate = {//}
		ldCurrentPay = {//}
		ldPreviousPay = {//}

		lnLastTSID = 0

		lnCurrentPay = VAL(Request.QueryString("currentPay"))
		lnPreviousPay = VAL(Request.QueryString("previousPay"))
		lnCurrentStaff = VAL(Request.QueryString("currentStaff"))
		lnCurrentGroup = VAL(Request.QueryString("currentGroup"))
		
		* MY - 22/11/2012 - select <everyone> then select <my details>, lnCurrentStaff = -1 which should be this.employee
		* maybe need a fix on the page when selecting <my details>
		IF lnCurrentGroup = -1 AND lnCurrentStaff = -1
			lnCurrentStaff = This.employee
		ENDIF 
	
		llManager = This.IsManager(This.Employee)
		IF lnCurrentStaff == 0
			lnCurrentStaff = This.Employee
		ENDIF

		IF !(This.SelectData(This.Licence, "myPays");
		  AND This.SelectData(This.Licence, "timesheet"))
			This.AddError("Load Timesheet Setup Failed!")
		ELSE
			IF EMPTY(lnPreviousPay)
				This.AddError("No Previous Pay selected!")
			ELSE
				llAllOk = .t.
				IF lnCurrentStaff == -1
					IF lnCurrentGroup > 0
						IF !llManager OR !This.CheckRights("PREV_PAY_APPLY_M")
							This.AddError("No Rights To Apply Previous Pay to a Group!")
							llAllOk = .f.
						ENDIF
					ENDIF
				ENDIF
				
				IF llAllOk
					CREATE CURSOR curStaff (tsEmp I(4,0))
					SELECT curStaff
					INDEX on tsEmp TAG tsEmp
					SET ORDER TO tsEmp
					GO top

					IF lnCurrentStaff == -1
						IF lnCurrentGroup > 0
							IF this.GetEmployeesByGroupCode(lnCurrentGroup, "curEmpGrp")
								GO TOP IN "curEmpGrp"
								DO WHILE NOT EOF("curEmpGrp")
									m.tsEmp = curEmpGrp.mywebcode
									IF NOT SEEK(m.tsEmp,"curStaff")
										INSERT INTO curStaff FROM memvar
									ENDIF
									SKIP IN "curEmpGrp"
								ENDDO
							ENDIF
						ENDIF
					ENDIF
	
					IF lnCurrentStaff > 0
						m.tsEmp = lnCurrentStaff
						IF NOT SEEK(m.tsEmp,"curStaff")
							INSERT INTO curStaff FROM memvar
						ENDIF
					ENDIF
					
					* Get the last/highest timesheet ID
					SELECT MAX(tsid) AS MAX_TSID FROM timesheet INTO CURSOR curMaxTimesheet

					SELECT curMaxTimesheet
					lnLastTSID = curMaxTimesheet.MAX_TSID

					* Get previous pay start date
					SELECT myPays
					LOCATE FOR pay_pk = lnPreviousPay
					IF !FOUND() 
						* This should not happen anyway
						This.AddError("Could not find Previous Pay!")
						llAllOk = .f.
					ELSE
						ldPreviousPay = myPays.pay_date
					ENDIF

					* Get current pay start date
					SELECT myPays
					LOCATE FOR pay_pk = lnCurrentPay
					IF !FOUND()
						* This should not happen anyway
						This.AddError("Could not find Current Pay!")
						llAllOk = .f.
					ELSE
						ldCurrentPay = myPays.pay_date
					ENDIF

					SELECT curStaff
					GO top
					IF EOF("curStaff")
						This.AddError("Nothing To Save!")
						llAllOk = .f.
					ENDIF

					IF llAllOk
						SELECT timesheet
						GO top

						lnFld = AFIELDS(laFld,"timesheet")
						CREATE CURSOR curPreviousPaySave FROM ARRAY laFld
						SELECT curPreviousPaySave
						GO top

						* Get the difference in days
						lnDaysDiff = ldCurrentPay - ldPreviousPay

						* Update the previous pay timesheet copy with the following:
						* 1) use new current pay
						* 2) use the new dates
						* 3) use the new IDs
                		* 4) make status 'Unapproved'
	                	* 5) Set downloaded to 'Not Done'
						SELECT curStaff
						GO top
						DO WHILE NOT EOF("curStaff")
						* Create a copy of the previous pay timesheet entries
							IF USED("curPreviousPayTS")
								USE IN "curPreviousPayTS"
							ENDIF

							SELECT * FROM timesheet ;
								WHERE (tsPay == lnPreviousPay) AND (tsEmp == curStaff.tsEmp);
																INTO CURSOR curPreviousPayTS READWRITE
							SELECT curPreviousPayTS				
							GO top
							IF NOT EOF("curPreviousPayTS")
								DO WHILE NOT EOF("curPreviousPayTS")
									lnLastTSID = lnLastTSID + 1
									ldNewDate = tsDate + lnDaysDiff
						
									REPLACE tsPay  WITH lnCurrentPay
									REPLACE tsDate WITH ldNewDate
									REPLACE tsID   WITH lnLastTSID

									* 09/06/2010  CMGM  TTP 5609  Entries to be copied should be 'Unapproved' by default
									REPLACE tsDownLoad WITH .F.
									* 18/04/2012  RAJ Entries to be copied should be 'Not Downloaded' by default
									REPLACE tsApproved WITH .F.
									SKIP IN "curPreviousPayTS"
								ENDDO
		
								* Insert previous pay timesheet copy into the temporoary save table
								SELECT curPreviousPayTS
								GO top
								INSERT INTO curPreviousPaySave SELECT * FROM curPreviousPayTS
							ENDIF
							
							SKIP IN "curStaff"
						ENDDO

						SELECT curPreviousPaySave
						GO top
*						BROWSE
						* Insert previous pay timesheet copy into the temporoary save table
						IF NOT EOF("curPreviousPaySave")
							INSERT INTO timesheet SELECT * FROM curPreviousPaySave
							* Display success message
							* JA 23/10/2012, changed loaded to applied, refer US 7731
							*This.AddUserInfo("Previous pay starting " + DMY(ldPreviousPay) + " is successfully loaded into the current pay.")
							This.AddUserInfo("Previous pay starting " + DMY(ldPreviousPay) + " is successfully applied into the current pay.")
						ELSE
							This.AddError("Nothing To Save!")
						ENDIF

						* Housekeeping
						USE IN SELECT("curEmpGrp")
						USE IN SELECT("curStaff")
						USE IN SELECT("curMaxTimesheet")
						USE IN SELECT("curPreviousPayTS")
						USE IN SELECT("curPreviousPaySave")
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		* Retain parameters
		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentPay", TRANSFORM(VAL(Request.QueryString("currentPay"))))
		loRetainList.SetEntry("previousPay", TRANSFORM(VAL(Request.QueryString("previousPay"))))
		loRetainList.SetEntry("currentGroup", TRANSFORM(VAL(Request.QueryString("currentGroup"))))
		loRetainList.SetEntry("currentStaff", TRANSFORM(VAL(Request.QueryString("currentStaff"))))
		
		lcFrom = Request.QueryString("from")		
		
		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		DO CASE
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
		ENDCASE

	ENDPROC


	*################################################################################*
#DEFINE TOC_Actions_

	*> +define: Actions
	PROCEDURE LogOut()
		Session.SetSessionVar("licence", 0)
		Session.SetSessionVar("employee", 0)
		Session.SetSessionVar("prevLoginDT", "")
		Session.SetSessionVar("IsPassReset", 0)
		This.Licence = 0
		This.Employee = 0
		This.AddUserInfo("Logout Successful.")
		Response.Redirect("CompanyPage.si" + This.AppendMessages('?'))		&& not using SOURCE_EXT on purpose
	ENDPROC

	*================================================================================*

	PROCEDURE ChangePassword()
		LOCAL loStaff, lcOldPasswd, lcNewPasswd0, lcNewPasswd1
		LOCAL lnMinPasswdLen, lnMaxPasswdLen, lnMinAlphaChars, lnMinNumericChars, lnMinOtherChars, llMixedCaseRequired
		LOCAL lnNumLetters, llUpperCase, llLowerCase, lnNumNumbers, lnNumOthers, lnCount, lcChar, lnCode
		loStaff = Factory.GetStaffObject()
		LOCAL lcNonce
		lcNonce = Session.GetSessionVar("ChangePasswordNonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		IF !loStaff.Load(This.Employee)
			This.AddError("Password not changed - Failed to load Employee record: " + loStaff.cErrorMsg)
		ELSE
			IF loStaff.oData.myPayCode == 0
				This.AddError("Payroll Users cannot change their passwords.")
			ELSE
				lcOldPasswd = Request.Form("oldPasswd")
				lcNewPasswd0 = Request.Form("newPasswd0")
				lcNewPasswd1 = Request.Form("newPasswd1")

				IF !(ALLTRIM(loStaff.oData.myPassword) == ALLTRIM(lcOldPasswd))
					This.AddError("Your current password is not correct and you have been logged out. Please try again.")
					This.LogOut()
					RETURN
				ENDIF
				IF !(lcNewPasswd0 == lcNewPasswd1)	&& Grr..  != not the same as !(==) (i.e. there is no !== shorthand)
					This.AddError("New Passwords do not match.")
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF lcOldPasswd == lcNewPasswd0
					This.AddError("New Password must be different from Old Password.")
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF lcNewPasswd0 == TRANSFORM(This.Licence) OR lcNewPasswd0 == ALLTRIM(loStaff.oData.myEmail)
					This.AddError("New Passwords must be different your other login details.")
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				lnMinPasswdLen		= VAL(AppSettings.Get("passwdMinLength"))
				lnMaxPasswdLen		= VAL(AppSettings.Get("passwdMaxLength"))
				lnMinAlphaChars		= VAL(AppSettings.Get("passwdMinAlphaChars"))
				lnMinNumericChars	= VAL(AppSettings.Get("passwdMinNumericChars"))
				lnMinOtherChars		= VAL(AppSettings.Get("passwdMinOtherChars"))
				llMixedCaseRequired = EVALUATE(AppSettings.Get("passwdMixedCaseRequired"))

				IF LEN(lcNewPasswd0) < lnMinPasswdLen
					This.AddError(;
						"New Passwords is too short." + CRLF;
						+ "Must be at least " + TRANSFORM(lnMinPasswdLen);
						+ " character" + IIF(lnMinPasswdLen == 1, "", "s") + " long.";
					)
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF LEN(lcNewPasswd0) > lnMaxPasswdLen
					This.AddError(;
						"New Passwords is too long." + CRLF;
						+ "Must be no more than " + TRANSFORM(lnMaxPasswdLen);
						+ " character" + IIF(lnMaxPasswdLen== 1, "", "s") + " long.";
					)
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				lnNumLetters = 0
				lnNumNumbers = 0
				lnNumOthers = 0
				llUpperCase = .F.
				llLowerCase = .F.
				FOR lnCount = 1 TO LEN(lcNewPasswd0)
					lcChar = SUBSTR(lcNewPasswd0, lnCount, 1)
					lnCode = ASC(lcChar)
					DO CASE
					CASE lnCode >= ASC('A') AND lnCode <= ASC('Z')
						lnNumLetters = lnNumLetters + 1
						llUpperCase = .T.
					CASE lnCode >= ASC('a') AND lnCode <= ASC('z')
						lnNumLetters = lnNumLetters + 1
						llLowerCase = .T.
					CASE lnCode >= ASC('0') AND lnCode <= ASC('9')
						lnNumNumbers = lnNumNumbers + 1
					OTHERWISE
						lnNumOthers = lnNumOthers + 1
					ENDCASE
				ENDFOR

				IF lnNumLetters < lnMinAlphaChars
					This.AddError(;
						"Not enough letters used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinNumericChars);
						+ " letter" + IIF(lnMinNumericChars == 1, "", "s");
						+ IIF(llMixedCaseRequired, ", including both upperCase and lowerCase." , ".");
					)
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF
				IF llMixedCaseRequired AND !(llUpperCase AND llLowerCase)
					This.AddError("New Password must use both upperCase and lowerCase letters.")
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				IF lnNumNumbers < lnMinNumericChars
					This.AddError(;
						"Not enough numbers used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinNumericChars);
						+ " number" + IIF(lnMinNumericChars == 1, "", "s") + ".";
					)
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				IF lnNumOthers < lnMinOtherChars
					This.AddError(;
						"Not enough non-alphanumeric characters used." + CRLF;
						+ "Must use at least " + TRANSFORM(lnMinOtherChars);
						+ " non-alphanumeric character" + IIF(lnMinOtherChars == 1, "", "s") + ".";
					)
					Response.Redirect("changePasswordPage.si" + This.AppendMessages('?'))
					RETURN
				ENDIF

				** Change the password...
				IF This.SelectData(This.Licence, "myStaff")
					UPDATE myStaff SET myPassword = lcNewPasswd0, myChanged = STUFF(myChanged, CHANGED_PASSWORD, 1, 'C') WHERE myWebCode = This.Employee
					This.AddUserInfo("Your password was changed successfully, please login again.")
					This.LogOut()
					RETURN
				ELSE
					This.AddError("Cannot change password.")
				ENDIF
			ENDIF
		ENDIF

		Response.Redirect("ChangePasswordPage.si" + This.AppendMessages('?'))
	ENDPROC

	*--------------------------------------------------------------------------------*
	* Called by forgot password page.
	
	PROCEDURE GetPassword()
		LOCAL lcCompany, lcUserCode, loError

		lcCompany = UPPER(ALLTRIM(Request.Form("licence")))
		lcUserCode = UPPER(ALLTRIM(Request.Form("email")))	&& this is coming from a reworked login page, so is still called email.

		IF EMPTY(lcCompany) OR !This.SelectData(lcCompany, "myStaff")
		    This.ReplaceError("")
			This.AddUserInfo("Details have been emailed.")
			Security.LogEvent("Failure", "Forgot Password", lcCompany, lcUserCode)
		ELSE
			LOCATE FOR UPPER(ALLTRIM(myUserName)) == lcUserCode

			IF !FOUND()
				This.AddUserInfo("Details have been emailed.")
				Security.LogEvent("Failure", "Forgot Password", lcCompany, lcUserCode)
			ELSE

				* We are not currently logged in.  Company ID is set to the company licence on login so we have to repeat the code here:
				AppSettings.cDataPath = ADDBS(ADDBS(AppSettings.cRootPath) + lcCompany)

				lcReplyTo = AppSettings.Get("update_reply")
				lcServerURL = This.GetINISection("serverURL","Server",FULLPATH("server.ini"))
				IF EMPTY(lcServerURL)
				   lcServerURL = "https://mystaffinfo.myob.com/"
				ENDIF
				** create tokens
				LOCAL lcToken, lcURL
				
				lcToken = This.CreateTokens(lcCompany,lcUserCode)
                IF !EMPTY(lcToken)
					* Email user
					lcUrl=ALLTRIM(lcServerUrl)+"ResetPasswordPage.si?tID="+lcToken
					loIPStuff = CREATEOBJECT("wwIPStuff")
					lcSubject = "MyStaffInfo - Requested Password"
					lcMessage = "You (or someone else) has requested to reset your password. " + CRLF + CRLF + "Your user name: " + lcUsercode + CRLF +;
					CRLF + "Please use this url " + lcURL + " to reset your password." + CRLF + CRLF + ;
					"This url will expire by " + ALLTRIM(TTOC(This.GetTokenExpiry(lcToken))) +"."
					TRY
					    DO wwdotnetbridge
						LOCAL lobridge AS wwdotnetbridge
						lobridge=CREATEOBJECT("wwDotNetBridge","V4")
						lcerror=SecureEmail(lobridge,SMTP_SERVER,.T.,SMTP_USER,SMTP_USER_PASSWORD,SMTP_SENDER_NAME,SMTP_SENDER_EMAIL,ALLTRIM(myStaff.myEmail),lcSubject, lcMessage, "", .F.,lcReplyTo, 101, This.Licence)
						IF EMPTY(lcerror)
							This.AddUserInfo("Details have been emailed.")
						ELSE
						    This.AddError(lcerror)
						ENDIF	
					CATCH TO loError
					This.AddError(TRANSFORM(loError.ErrorNo) + ": " + loError.Message + "; " + loError.Details)
					ENDTRY
					Security.LogEvent("Success", "Forgot Password", lcCompany, lcUserCode)
				ELSE
		           This.ReplaceError("Unable to send password reset email. Please try again later.")
				   Security.LogEvent("Failure", "Unable to send password reset tokens", lcCompany, lcUserCode)				
				ENDIF
			ENDIF
		ENDIF

		* Log out, just in case
		Session.SetSessionVar("licence", 0)
		Session.SetSessionVar("employee", 0)
		Session.SetSessionVar("prevLoginDT","")
		Session.SetSessionVar("IsPassReset","N")
		This.Licence = 0
		This.Employee = 0

		This.LoginPage(This.IsError())
	ENDPROC


	*================================================================================*
	
	
	
	***********************************************************************************************
	PROCEDURE SimpleEnc
	***********************************************************************************************
	* Decrypt user password.
	LPARAMETERS Encrypted
	LOCAL Decrypted
 
	IF EMPTY(Encrypted)
		Decrypted = SPACE(8)
	ELSE
		Decrypted = CHRTRAN(Encrypted, 'MNBVCXZLKJHGFDSAPOIUYTREWQ1234567890', ;
			 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghij')
	ENDIF
 
	RETURN Decrypted
	***********************************************************************************************



	
	***********************************************************************************************
	PROCEDURE createtokens(pccompany AS STRING, pcusercode AS STRING)
	***********************************************************************************************
	LOCAL lctokendbf, lctokenstr, lcevilchars
	lctokenstr=""
	TRY
		lctokendbf = ADDBS(PROCESS.cdatapath) +  + "\passtokens.dbf"
		IF !FILE(lctokendbf)
			CREATE TABLE (lctokendbf) FREE (tokenid INTEGER AUTOINC, tokencomp V(50), tokenuser V(100), tokenstr V(200), tokenexp T, tokenused l)
			INDEX ON ALLTRIM(tokenstr) TAG tokenstr
			USE 
		ENDIF
		SELECT tokenstr FROM (lctokendbf) WHERE ALLTRIM(UPPER(tokenuser))==ALLTRIM(UPPER(pcusercode)) AND ;
		  ALLTRIM(UPPER(tokencomp))==ALLTRIM(UPPER(pccompany)) AND !tokenused AND tokenexp>=DATETIME() ;
		   INTO CURSOR tmptoken
		IF _TALLY=0
			lctokenstr = THIS.guidgen(1)
			lcevilchars="<>&{}:-*~@#$,();."
			IF !EMPTY(lctokenstr)
				lctokenstr=This.SimpleEnc(pccompany)+lctokenstr+This.SimpleEnc(pcusercode)+STR(YEAR(DATE()),4)+SYS(3)+ALLTRIM(STR(MONTH(DATE()),2))+SYS(3)+ALLTRIM(STR(DAY(DATE()),2))+STRTRAN(TIME(),":","")
			    lctokenstr=ALLTRIM(CHRTRAN(lctokenstr,lcevilchars, ""))
				INSERT INTO (lctokendbf) (tokencomp,tokenuser,tokenstr,tokenexp) VALUES (pccompany, pcusercode, lctokenstr, DATETIME()+86400)
			ENDIF
		ELSE
		   SELECT tmptoken
		   GO TOP
		   lctokenstr=ALLTRIM(tmptoken.tokenstr)
		   USE IN SELECT("tmptoken")
		ENDIF	
	CATCH
		lctokenstr=""
	ENDTRY
	RETURN lctokenstr
	*================================================================================*
	
	*================================================================================*	
	PROCEDURE GetTokenExpiry(pcToken AS String)
	*================================================================================*	
	LOCAL lctokendbf, retVal
	retVal=DATETIME()
	TRY
	   lctokendbf = ADDBS(PROCESS.cdatapath) +  + "\passtokens.dbf"
	   SELECT tokenexp FROM (lctokendbf) WHERE ALLTRIM(tokenstr)==ALLTRIM(pctoken) INTO ARRAY tmpTokenExp
	   IF _TALLY>0
	      retVal=tmpTokenExp[1]
	   ENDIF
	CATCH
       ** do nothing
	ENDTRY
	RETURN retVal
	
	
	*================================================================================*	
	
	*================================================================================*
	PROCEDURE GUIDGen 
	*================================================================================*
	LPARAMETERS tn_mode as Integer

	LOCAL ; 
		lc_guid_return as String, ; 
		lc_buffer as String, ; 
		ln_result as Integer, ; 
		lc_GUID as String 

	DECLARE Integer CoCreateGuid ; 
	   IN ole32.dll ; 
	   String@ pguid 

	lc_GUID = SPACE(16) && 16 Byte = 128 Bit 
	ln_result = CoCreateGuid(@lc_GUID) 

	IF tn_mode = 0 
		lc_guid_return = lc_GUID 
	ELSE 
		lc_buffer = SPACE(78) 

		DECLARE Integer StringFromGUID2 ; 
		    IN ole32.dll ; 
		    String  pguid, ; 
		    String  @lpszBuffer, ; 
		    Integer cbBuffer 

		ln_result = StringFromGUID2(lc_GUID,@lc_buffer,LEN(lc_buffer)/2) 
		lc_guid_return = STRCONV((LEFT(lc_buffer,(ln_result-1)*2)),6) 
	ENDIF 


	RETURN lc_guid_return
ENDFUNC
	


	*================================================================================*
	
	

	*================================================================================*
	PROCEDURE SendPassUpdate
	*================================================================================*
	LOCAL lcServer, lcUpdateName, lcUpdateAdr
	
	lcReplyTo = AppSettings.Get("update_reply")
	
	lcSubject = "MyStaffInfo - Password has been changed"
	lcMessage = "You (or someone else) has successfully changed your existing password for MyStaffInfo. " + CRLF + CRLF 
	TRY
	    DO wwdotnetbridge
		LOCAL lobridge AS wwdotnetbridge
		lobridge=CREATEOBJECT("wwDotNetBridge","V4")
		lcerror=SecureEmail(lobridge,SMTP_SERVER,.T.,SMTP_USER,SMTP_USER_PASSWORD,SMTP_SENDER_NAME,SMTP_SENDER_EMAIL,ALLTRIM(myStaff.myEmail),lcSubject, lcMessage, "", .F.,lcReplyTo, 100, This.Licence)
		IF EMPTY(lcerror)
		   This.AddUserInfo("Details have been emailed.")
		ELSE
		   This.AddError(lcerror)
		ENDIF	
	CATCH TO loError
		   This.AddError(TRANSFORM(loError.ErrorNo) + ": " + loError.Message + "; " + loError.Details)
	ENDTRY
	RETURN
    *================================================================================*


	
	
	*================================================================================*
	PROCEDURE SaveAdminSettings()
	*================================================================================*
	
		LOCAL lnCount, lnVars, lcType, lcName, luValue


	    LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
*!*			IF NOT llValidNonce
*!*				This.AddUserInfo("Page access has expired.")
*!*				This.LogOut()
*!*				RETURN
*!*			ENDIF
		
		IF !This.IsAdminUser(This.Employee) OR EMPTY(Request.Form("save_admin_settings"))
			This.AddError("You are not autorised to access this page!")
			This.AddUserInfo("Your attempt has been logged.")
			This.LogReport("SaveAdminSettings", "Attempted illegal access")
		ELSE
			DIMENSION laVars[1, 2]

			lnVars = Request.aFormVars(@laVars)
			FOR lnCount = 1 TO lnVars
				IF LEFT(laVars[lnCount, 1], 6) == "_type_"
					lcName = SUBSTR(laVars[lnCount, 1], 7)

					
					lcType = laVars[lnCount, 2]
					luValue = Request.Form(lcName)

					DO CASE
						CASE lcType $ "NFIBY"
							luValue = VAL(luValue)
						CASE lcType $ "DT"
							luValue = CTOD(luValue)
						CASE lcType $ 'L'
							luValue = !EMPTY(luValue)
					ENDCASE
					IF (INLIST(ALLTRIM(LOWER(lcName)),'updatename','update_adr', 'update_reply') OR ;
					  LEFT(ALLTRIM(LOWER(lcName)),7)='message' OR  LEFT(ALLTRIM(LOWER(lcName)),7)='subject') AND VARTYPE(luValue)='C'
					  luValue = this.HTMLGetSafeAntiXSS(luvalue)
					  IF EMPTY(ALLTRIM(luValue))
					     luValue = This.GetDefValueForNotices(ALLTRIM(LOWER(lcName)))
					  ENDIF
					ENDIF  
					IF !(LOWER(lcName) == "mailserver") OR This.Employee == -999
						IF LOWER(lcName) == "releaselocks"
							This.ReleaseLockedUsers()
							This.AdduserInfo("Locked user accounts have been released.")
						ELSE
							AppSettings.Put(lcName, luValue)
						ENDIF 	
					ELSE
						This.AddValidationError("Attempt to alter mailserver denied and logged!")
						This.LogReport("SaveAdminSettings", "Attempted to alter mailserver")
					ENDIF
				ENDIF
			ENDFOR

			This.AddUserInfo("Settings Saved.")
		ENDIF

		Response.Redirect("AdminPage.si" + This.AppendMessages('?'))
	ENDPROC

	*--------------------------------------------------------------------------------*
	
	PROCEDURE GetDefValueForNotices
	LPARAMETERS lcInput
	LOCAL lcRetValue
	DO CASE
	   CASE  LEFT(ALLTRIM(LOWER(lcInput)),7)='message'
	         lcRetValue="You have new updates on the mystaffinfo website."+CHR(13)+CHR(13)+"Please visit https://mystaffinfo.myob.com to view them."+CHR(13)+CHR(13)
	   CASE  LEFT(ALLTRIM(LOWER(lcInput)),7)='subject'
	         lcRetValue="MyStaffInfo - New Updates"+CHR(13)+CHR(13)	
	   CASE  ALLTRIM(lcInput)='updatename'
	         lcRetValue='MyStaffInfo Administrator'
	   CASE  ALLTRIM(lcInput)='update_adr'
	         lcRetValue='MyStaffInfoAdmin@myob.com'	
	   CASE  ALLTRIM(lcInput)='update_reply'
	         lcRetValue='noreply.mystaffinfo@myob.com'	
	   OTHERWISE
	         lcRetValue=''
	ENDCASE
	RETURN lcRetValue                                      
	
	ENDPROC
	

	PROCEDURE ReleaseLockedUsers()
		LOCAL loUsers, lcMessage, lcUser
		loUsers = Factory.GetStaffObject()
		loUsers.GetLockedOutUsers()
		
		loUsers.load(This.Employee)
		lcUser = loUsers.odata.myusername
		
		* process and log each one
		IF USED("TLockedOut")
			SELECT TLockedOut
			SCAN
				lcMessage = ;
					ALLTRIM(TLockedOut.FullName) + ;
					", Locked at " + ;
					TRANSFORM(TLockedOut.mylocked) + ;
					" after " + ;
					TRANSFORM(TLockedOut.myattempts) + ;
					" attempts."
				Security.LogEvent("Unlock", "Admin", This.Licence, lcUser, lcMessage)
			ENDSCAN
		ENDIF	
		
		loUsers.ReleaseLocks()
	ENDPROC 

	*--------------------------------------------------------------------------------*

	PROCEDURE SaveAdminTypes()
		LOCAL llAllowOk

		LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		
		IF !This.SelectData(Process.Licence, "wageType") OR !This.SelectData(Process.Licence, "costcent") OR;
			!This.SelectData(Process.Licence, "allow")
			This.AddError("Page Setup Failed!")
		ELSE
			SELECT wagetype
			SCAN
				REPLACE Hide WITH EMPTY(Request.Form("chk" + TRANSFORM(code)))
			ENDSCAN

			SELECT costcent
			SCAN
				REPLACE Hide WITH EMPTY(Request.Form("chkcost" + TRANSFORM(code)))
			ENDSCAN

			llAllowOk = .f.
			IF "chkall"$request.cformvars
				llAllowOk = .t.
				SELECT allow
				SCAN
					REPLACE Hide WITH EMPTY(Request.Form("chkall" + TRANSFORM(code)))
				ENDSCAN
			ENDIF

			IF llAllowOk					
				This.AddUserInfo("Wage / Allowance / Cost Centre Types Saved.")
			ELSE
				This.AddUserInfo("Wage / Cost Centre Types Saved. At least one Allowance must be selected.")
			ENDIF
		ENDIF

		Response.Redirect("WageTypeAdminPage.si" + This.AppendMessages('?'))
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ChkAppSettings(tcName,tuDefVal)
		LOCAL puValue

		puValue = AppSettings.Get(tcName)
		IF TYPE("puValue") = "C"
			IF NOT EMPTY(puValue)
				puValue = EVALUATE("puValue")
			ENDIF
		ENDIF
	
		IF TYPE("puValue") <> TYPE("tuDefVal")
			AppSettings.Put(tcName, tuDefVal)
		ENDIF
	ENDPROC

	*================================================================================*

	PROCEDURE SaveContactInformation()
		LOCAL loStaff, lnCurrentGroup, loRetainList, llManager, llCanSave, llChangeBank, llShowUserDefined, llChangeUserDefined, llCanView
		LOCAL lcBank, lnBank, lnBranch, lnAccount, lnSuffix, llShowBank, llShowBank, llAusie, llSaveDetails, llSaveBank, llSaveUserDefinedDetails, llProblem
		PRIVATE pnCurrentStaff, pcEmail

		pnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
		lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentStaff", pnCurrentStaff)
		loRetainList.SetEntry("currentGroup", lnCurrentGroup)

		llManager = This.IsManager(This.Employee)

		loStaff = Factory.GetStaffObject()

        LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		* 02/11/2k9;TTP4715;JCF: reworked security settings checks so change is dependant on view and saving UserDefinedInfo is separate from normal contact details saving.
		IF !loStaff.Load(pnCurrentStaff)
			This.AddError("Failed to update Employee Record: " + loStaff.cErrorMsg)
		ELSE
			llCanView = This.CheckRights("V_DETAILS")
			llShowBank = .T.	&& Always show the bank details section (even if there is none - will show blank fields)
			llShowUserDefined = llCanView AND This.CheckRights("V_UD_INFO") AND !(;
				EMPTY(loStaff.oData.myUdfL1d) AND EMPTY(loStaff.oData.myUdfL2d) AND EMPTY(loStaff.oData.myUdfD1d);
				AND EMPTY(loStaff.oData.myUdfD2d) AND EMPTY(loStaff.oData.myUdfC1d) AND EMPTY(loStaff.oData.myUdfC2d);
				AND EMPTY(loStaff.oData.myUdfC3d) AND EMPTY(loStaff.oData.myUdfN1d) AND EMPTY(loStaff.oData.myUdfM1d);
			)


			*!* 21/03/2010  CMGM  MSI 2011.01  TTP6417  Replaced below
			*!*	*!* 09/08/2010  CMGM  MSI 2010.02  TTP5898  New option to change Employee Groups
			*!*	*!*	*!* 14/07/2010  CMGM  MSI 2010.02  Managers now can save employee details within their group
			*!*	*!*	*!*	llCanSave = llCanView AND This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee
			*!*	*!*	llCanSave = llCanView AND This.CheckRights("C_DETAILS") AND (pnCurrentStaff == This.Employee OR llManager)
			*!*	llCanSave = llCanView AND ( ( This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee ) OR;
			*!*	                            ( This.CheckRights("C_GROUP_DETAILS") AND llManager) )
			*!*	llChangeBank = llShowBank AND This.CheckRights("CHANGEBANK") AND (pnCurrentStaff == This.Employee OR llManager)
			*!*	llChangeUserDefined = llShowUserDefined AND This.CheckRights("CHANGEUD")


			*!* 21/03/2010  CMGM  MSI 2011.01  TTP6417  Standardised rules for Contact Details changes
			IF !llManager
				* If regular employee, only check for personal information
				llCanSave = ( This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee )
				llChangeBank = ( This.CheckRights("CHANGEBANK") AND pnCurrentStaff == This.Employee )
				llChangeUserDefined = This.CheckRights("CHANGEUD")
			ELSE
				* If a manager, need to check group.  Dummy Group ("My Details") have value of -1.
				IF lnCurrentGroup < 0
					*  If editing own details, need to check for personal information
					llCanSave = ( This.CheckRights("C_DETAILS") AND pnCurrentStaff == This.Employee )
					llChangeBank = ( This.CheckRights("CHANGEBANK") AND pnCurrentStaff == This.Employee )
					llChangeUserDefined = ( This.CheckRights("CHANGEUD") AND pnCurrentStaff == This.Employee )
				ELSE
					*  If editing group details, need to check for "Change Group Contact Details"
					llCanSave = This.CheckRights("C_GROUP_DETAILS")
					llChangeBank = llCanSave
					llChangeUserDefined = llCanSave
				ENDIF
			ENDIF


			IF !(llCanSave OR llChangeBank OR llChangeUserDefined)
				This.AddError("You do not have access to alter this information!")
			ELSE
				llAusie = This.IsAustralia()

				llProblem = .F.

				llSaveBank = .F.
				IF llChangeBank
					llSaveBank = .T.
					* Only validate bank account if user can change it.

					lnBank = VAL(Request.Form("bank"))
					lnBranch = VAL(Request.Form("branch"))
					lnAccount = VAL(Request.Form("account"))

					IF llAusie
					   IF FILE(This.CompanyDataPath() + "png.mem")
						  lcBank = PADL(ALLTRIM(Request.Form("bank")), 3, '0') + "-";
							+ PADL(ALLTRIM(Request.Form("branch")), 3, '0') + "-";
							  + PADL(ALLTRIM(Request.Form("account")), 10, '0')
						ELSE
						  lcBank = PADL(ALLTRIM(Request.Form("bank")), 3, '0') + "-";
							+ PADL(ALLTRIM(Request.Form("branch")), 3, '0') + "-";
							  + PADL(ALLTRIM(Request.Form("account")), 9, '0')
						ENDIF	  
							
					ELSE
						lnSuffix = VAL(Request.Form("suffix"))
						lcBank = ;
							PADL(ALLTRIM(Request.Form("bank")), 2, '0') + "-";
							+ PADL(ALLTRIM(Request.Form("branch")), 4, '0') + "-";
							+ PADL(ALLTRIM(Request.Form("account")), 7, '0') + "-";
							+ PADL(ALLTRIM(Request.Form("suffix")), 3, '0')
					ENDIF

					DO CASE
						CASE llAusie
							* Nothing to check
						CASE !This.ValidateBankAccount(lnBank, lnBranch, lnAccount, lnSuffix)
							This.AddValidationError("Bank Account number is not valid.")
							llSaveBank = .F.
							llProblem = .T.
					ENDCASE

					IF llSaveBank AND !(lcBank == ALLTRIM(loStaff.oData.myBank))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_BANK, 1, 'C')
						loStaff.oData.myBank = lcBank
					ENDIF
				ENDIF

				llSaveUserDefinedDetails = .F.
				IF llChangeUserDefined
					* Save UDF fields
					* L1, L2, D1, D2, C1, C2, C3, N1, M1
					** Booleans **
					LOCAL luField
					IF !EMPTY(loStaff.oData.myUdfL1D)
						luField = IIF(UPPER(ALLTRIM(Request.Form("udfL1"))) == "ON", .T., .F.)
						IF luField <> loStaff.oData.myUdfL1
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_L1, 1, 'C')
							loStaff.oData.myUdfL1 = luField
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF
					IF !EMPTY(loStaff.oData.myUdfL2D)
						luField = IIF(UPPER(ALLTRIM(Request.Form("udfL2"))) == "ON", .T., .F.)
						IF luField <> loStaff.oData.myUdfL2
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_L2, 1, 'C')
							loStaff.oData.myUdfL2 = luField
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF

					** Dates **
					IF !EMPTY(loStaff.oData.myUdfD1D)
						IF !(EVALUATE("{" + Request.Form("udfD1") + "}") == loStaff.oData.myUdfD1)
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_D1, 1, 'C')
							loStaff.oData.myUdfD1 = EVALUATE("{" + Request.Form("udfd1") + "}")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF
					IF !EMPTY(loStaff.oData.myUdfD2D)
						IF !(EVALUATE("{" + Request.Form("udfD2") + "}") == loStaff.oData.myUdfD2)
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_D2, 1, 'C')
							loStaff.oData.myUdfD2 = EVALUATE("{" + Request.Form("udfd2") + "}")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF

					** Strings **
					IF !EMPTY(loStaff.oData.myUdfC1D)
						IF !(ALLTRIM(Request.Form("udfC1")) == ALLTRIM(loStaff.oData.myUdfC1))
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_C1, 1, 'C')
							loStaff.oData.myUdfC1 = Request.Form("udfc1")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF
					IF !EMPTY(loStaff.oData.myUdfC2D)
						IF !(ALLTRIM(Request.Form("udfC2")) == ALLTRIM(loStaff.oData.myUdfC2))
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_C2, 1, 'C')
							loStaff.oData.myUdfC2 = Request.Form("udfc2")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF
					IF !EMPTY(loStaff.oData.myUdfC3D)
						IF !(ALLTRIM(Request.Form("udfC3")) == ALLTRIM(loStaff.oData.myUdfC3))
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_C3, 1, 'C')
							loStaff.oData.myUdfC3 = Request.Form("udfc3")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF

					** Numbers **
					IF !EMPTY(loStaff.oData.myUdfN1D)
						IF !(VAL(Request.Form("udfN1")) == loStaff.oData.myUdfN1)
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_N1, 1, 'C')
							loStaff.oData.myUdfN1 = VAL(Request.Form("udfn1"))
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF

					** Memos **
					IF !EMPTY(loStaff.oData.myUdfM1D)
						IF !(ALLTRIM(Request.Form("udfM1")) == ALLTRIM(loStaff.oData.myUdfM1))
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_UDF_M1, 1, 'C')
							loStaff.oData.myUdfM1 = Request.Form("udfm1")
							llSaveUserDefinedDetails = .T.
						ENDIF
					ENDIF
				ENDIF

				llSaveDetails = .F.
				IF llCanSave
					&&TODO: validate lengths!
					IF !(ALLTRIM(Request.Form("slip_name")) == ALLTRIM(loStaff.oData.mySlipName))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_SLIPNAME, 1, 'C')
						loStaff.oData.mySlipName = Request.Form("slip_name")
						llSaveDetails = .T.
					ENDIF
					IF !(ALLTRIM(Request.Form("address")) == ALLTRIM(loStaff.oData.myAddress))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_ADDRESS, 1, 'C')
						loStaff.oData.myAddress = Request.Form("address")
						llSaveDetails = .T.
					ENDIF
					IF llAusie
						IF !(ALLTRIM(Request.Form("address2")) == ALLTRIM(loStaff.oData.myAddress2))
							loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_ADDRESS2, 1, 'C')
							loStaff.oData.myAddress2 = Request.Form("address2")
							llSaveDetails = .T.
						ENDIF
					ENDIF
					IF !(ALLTRIM(Request.Form("suburb")) == ALLTRIM(loStaff.oData.mySuburb))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_SUBURB, 1, 'C')
						loStaff.oData.mySuburb	= Request.Form("suburb")
						llSaveDetails = .T.
					ENDIF
					IF !(ALLTRIM(Request.Form("city")) == ALLTRIM(loStaff.oData.myCity))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_CITY, 1, 'C')
						loStaff.oData.myCity		= Request.Form("city")
						llSaveDetails = .T.
					ENDIF
					IF !(ALLTRIM(Request.Form("phone")) == ALLTRIM(loStaff.oData.myPhone))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_PHONE, 1, 'C')
						loStaff.oData.myPhone		= Request.Form("phone")
						llSaveDetails = .T.
					ENDIF
					IF !(VAL(Request.Form("postcode")) == (loStaff.oData.myPostCode))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_POSTCODE, 1, 'C')
						loStaff.oData.myPostCode		= VAL(Request.Form("postcode"))
						llSaveDetails = .T.
					ENDIF
					IF !(ALLTRIM(Request.Form("mobile")) == ALLTRIM(loStaff.oData.myMobile))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_MOBILE, 1, 'C')
						loStaff.oData.myMobile		= Request.Form("mobile")
						llSaveDetails = .T.
					ENDIF

					pcEmail = ALLTRIM(Request.Form("email"))
					IF EMPTY(pcEmail)
						This.AddValidationError("Email address can not be empty.")
						llSaveDetails = .F.
						llProblem = .T.
					ENDIF
					loStaff.Query("myWebCode FROM myStaff WHERE myWebCode != pnCurrentStaff AND UPPER(ALLTRIM(myEmail)) == UPPER(ALLTRIM(pcEmail)) INTO CURSOR curDupes")
					IF USED("curDupes") AND RECCOUNT("curDupes") != 0
						This.AddValidationError("This email address has already been used by another employee.")
						llSaveDetails = .F.
						llProblem = .T.
					ENDIF
					
					IF !(ALLTRIM(Request.Form("email")) == ALLTRIM(loStaff.oData.myEmail))
						loStaff.oData.myChanged = STUFF(loStaff.oData.myChanged, CHANGED_EMAIL, 1, 'C')
						loStaff.oData.myEmail = ALLTRIM(Request.Form("email"))
						loStaff.oData.myUserName = ALLTRIM(Request.Form("email"))
						llSaveDetails = .T.
					ENDIF
					
					
				ENDIF

				IF (llSaveDetails OR llSaveBank OR llSaveUserDefinedDetails) AND !llProblem
					IF !loStaff.Save()
						This.AddError("User details save failed: " + loStaff.cErrorMsg)
					ELSE
						This.AddUserInfo("User details saved.")
					ENDIF
				ELSE
					This.AddUserInfo("User details unchanged - nothing to save.")
				ENDIF
			ENDIF
		ENDIF

		Response.Redirect("ContactInformationPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE SaveLocation()
		LOCAL lnCurrentStaff, lnCurrentGroup, loRetainList, loStaff, lcStatus, lcDueDate, lcDueTime, ldDateTime, llCanSave

		lnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
		lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentStaff", lnCurrentStaff)
		loRetainList.SetEntry("currentGroup", lnCurrentGroup)

		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed.")
		ELSE
			llCanSave = This.CheckRights("C_LOCATOR")

			loStaff = Factory.GetStaffObject()
			IF !loStaff.Load(lnCurrentStaff)
				This.AddError("Failed to update Employee Record: " + loStaff.cErrorMsg)
			ELSE
				IF !This.CheckAccess(lnCurrentStaff, This.IsManager(This.Employee))
					This.AddError("You do not have access to this page.")
				ELSE
					IF !This.CheckRights("C_LOCATOR")
						This.AddError("You do not have access to alter this information!")
					ELSE
						lcStatus = ALLTRIM(Request.Form("status"))
						lcDueDate = Request.Form("dueDate")
						lcDueTime = Request.Form("dueTime")

						IF EMPTY(lcDueDate)
							lcDueDate = DTOC(DATE())
						ENDIF
						IF EMPTY(lcDueTime)
							lcDueTime = TTOC(DATETIME(), 2)
						ENDIF

						ldDateTime = DATETIME(;
							VAL(SUBSTR(lcDueDate, 7, 4)),;
							VAL(SUBSTR(lcDueDate, 4, 2)),;
							VAL(SUBSTR(lcDueDate, 1, 2)),;
							VAL(SUBSTR(lcDueTime, 1, 2)),;
							VAL(SUBSTR(lcDueTime, 4, 2)),;
							0;
						)

						loStaff.oData.lbStatus	= lcStatus
						loStaff.oData.lbDueBack	= ldDateTime
						loStaff.oData.lbNotes	= ALLTRIM(Request.Form("notes"))

						* save the changes
						IF !loStaff.Save()
							This.AddError("Failed to Save changes.")
						ELSE
							This.AddUserInfo("Changes Saved.")
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		Response.Redirect("LocatorBoardPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
	ENDPROC

	*================================================================================*

	PROCEDURE MakeLeaveRequest()
		LOCAL lcDates, lcDate, lcType, lnTo, lnUnits, lnI, lnNumDates, loLeaveCode, lnEmployee
		LOCAL lnMaxUnits, ltTime, loStaff, lcEmployeeName, lcManagerName, lnTotalUnits, lnDaysNoMedical
		LOCAL ARRAY laDates[1]

		lcDates = Request.Form("days")
		IF EMPTY(lcDates)
			This.AddValidationError("You must select at least one day to continue!")
		ENDIF

		lnEmployee = EVL(VAL(Request.Form("currentStaff")), This.Employee)
		IF !This.CheckAccessForManager(lnEmployee)
			* Making this a fatal error by not including the MESSAGE_VALIDATION_PREFIX since this can only happen if people are hacking the formContents or have an old page.
			This.AddError("You do not have access to request leave for that Employee!")		&&TODO: log this?  (We aren't calling CheckAccess so it's not logged for us...)
		ENDIF

		lnTo = VAL(Request.Form("to"))
		IF !This.CheckAccessForManager(lnEmployee, lnTo)
			* Making this a fatal error by not including the MESSAGE_VALIDATION_PREFIX since this can only happen if people are hacking the formContents or have an old page.
			This.AddError("The selected to-manager is not a manager of the " + IIF(lnEmployee == This.Employee, "current", "selected") + " employee!")		&&TODO: log this?  (We aren't calling CheckAccess so it's not logged for us...)
		ENDIF

		lcType = Request.Form("type")
		loLeaveCode = This.GetLeaveCode(lcType, lnEmployee)
		IF ISNULL(loLeaveCode)
			This.AddValidationError("Invalid leave type selected!  Can't find leaveCode for '" + lcType + "'")
			lnMaxUnits = 8		&& to avoid things breaking below, stopping us from seeing the collected errors
			IF !EMPTY(This.cErrorField)
				This.cErrorField = "type"
			ENDIF
		ELSE
			lnMaxUnits = IIF(UPPER(loLeaveCode.units) == "HOURS", 24, 1)
		ENDIF

		IF !EMPTY(lcDates)		&& otherwise we'd get a spurious empty-date error as well
			lnNumDates = ALINES(laDates, lcDates, 0, ',')

			LOCAL ARRAY laUnits[lnNumDates]

			FOR lnI = 1 TO lnNumDates
				lcDate = laDates[lnI]
				laDates[lnI] = CTOD(lcDate)
				IF EMPTY(laDates[lnI])
					This.AddValidationError("Invalid day selected: '" + lcDate + "'")
				ELSE
					laUnits[lnI] = VAL(Request.Form('d' + DTOC(laDates[lnI], 1)))
					IF laUnits[lnI] <= 0 OR laUnits[lnI] > lnMaxUnits
						This.AddValidationError("Units out of range for " + lcDate + "!")
						IF !EMPTY(This.cErrorField)
							This.cErrorField = 'd' + DTOC(laDates[lnI], 1)
						ENDIF
					ENDIF
				ENDIF
			NEXT
		ENDIF

		* HG 01/09/2009 TTP1328 get the value of sick daysNoMedical when it's personal leave; or set it to 0 for the rest of leave types
		IF lcType = "S"
			lnDaysNoMedical = VAL(Request.Form("daysNoMedical"))
		ELSE
			lnDaysNoMedical = 0
		ENDIF

		IF This.IsError()
			* Go back and handle any errors, fatal or otherwise...
			This.MakeLeaveRequestPage()
			RETURN
		ENDIF

		* Got everything we need, so on with the request-making...
		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "leaveRequestDays");
		  AND This.SelectData(This.Licence, "leaveRequestStatus"))
			This.AddError("Page setup failed - Cannot create request!")
			This.MakeLeaveRequestPage()
		ELSE
			loStaff = Factory.GetStaffObject()

			IF !loStaff.Load(lnEmployee)
				This.AddError("Failed to load LeaveRequest Employee record: " + loStaff.cErrorMsg)
				This.MakeLeaveRequestPage()
			ELSE
				lcEmployeeName = loStaff.fullName

				IF !loStaff.Load(lnTo)
					This.AddError("Failed to load LeaveRequest Manager record: " + loStaff.cErrorMsg)
					This.MakeLeaveRequestPage()
				ELSE
					lcManagerName = loStaff.fullName
					lcComments = This.htmlsanitiser(this.URLUnescape(Request.Form("comments")))

					* Create a new leave request
					* HG 01/09/2009 TTP 1328 Added new column sick_nomed
					ltTime = This.GetLocalTime(DATETIME())
					SELECT leaveRequests
					APPEND BLANK
					REPLACE;
						leaveCode	WITH loLeaveCode.code,;
						leaveType	WITH loLeaveCode.name,;
						employee	WITH lnEmployee,;
						manager		WITH lnTo,;
						empName		WITH lcEmployeeName,;
						manName		WITH lcManagerName,;
						comment		WITH This.EnforceMaxWordLength(lcComments, MAX_WORD_LENGTH),;
						dateMade	WITH ltTime;
						sick_nomed	WITH lnDaysNoMedical;
						IN leaveRequests

					lnId = leaveRequests.id

					* Populate the requested dates/units
					lnTotalUnits = 0
					FOR lnI = 1 TO lnNumDates
						SELECT leaveRequestDays
						APPEND BLANK
						REPLACE;
							leaveReqID	WITH lnId,;
							date		WITH laDates[lnI],;
							units		WITH laUnits[lnI],;
							unitType	WITH loLeaveCode.units;
							IN leaveRequestDays

						lnTotalUnits = lnTotalUnits + laUnits[lnI]
					NEXT

					* And set the intial status message
					SELECT leaveRequestStatus
					APPEND BLANK
					REPLACE;
						leaveReqID	WITH lnId,;
						from		WITH lnEmployee,;
						to			WITH lnTo,;
						fromName	WITH lcEmployeeName,;
						toName		WITH lcManagerName,;
						subject		WITH loLeaveCode.name + " - " + TRANSFORM(lnTotalUnits) + " " + loLeaveCode.units + ".",;
						message		WITH This.EnforceMaxWordLength(lcComments, MAX_WORD_LENGTH),;
						sent		WITH ltTime;
						IN leaveRequestStatus

					This.SendSiteUpdatedEmailTo(lnTo, SENDMAIL_LEAVEREQUEST_NEW)

					* Go back to leave request page
					This.AddUserInfo("Leave request has been saved.")

					Response.Redirect("ViewLeaveRequestPage.si?id=" + TRANSFORM(lnId) + This.AppendMessages("&show=pending&"))
					
				ENDIF
			ENDIF
		ENDIF
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	* Action the approve/decline...
	PROCEDURE SendLeaveResponse()
		LOCAL lnId, llManage, lcMode, lnTo, lcSubject, lcMessage, loStaff, lcFromName, lcToName

		lnId = VAL(Request.Form("id"))
		lcMode = Request.Form("mode")
		llManage = !EMPTY(Request.Form("manage"))

		IF !INLIST(lcMode, "-1", '0', '1')
			This.AddError("Invalid mode.")
		ELSE
			IF INLIST(lcMode, "-1", '1')
				llManage = .T.
			ENDIF

			IF !(This.SelectData(This.Licence, "myStaff");
			  AND This.SelectData(This.Licence, "leaveRequests");
			  AND This.SelectData(This.Licence, "leaveRequestStatus"))
				This.AddError("Page setup failed - Cannot delete request!")
			ELSE
				SELECT leaveRequests
				LOCATE FOR id == lnId
				IF !FOUND()
					This.AddError("Cannot find request to delete!")
				ELSE
					lnTo = VAL(Request.Form("to"))

					loStaff = Factory.GetStaffObject()
					IF !loStaff.Load(This.Employee)
						This.AddError("Failed to load from-employee record: " + loStaff.cErrorMsg)
					ELSE
						lcFromName = loStaff.fullName
					ENDIF

					IF !loStaff.Load(lnTo)
						This.AddError("Failed to load to-employee record: " + loStaff.cErrorMsg)
					ELSE
						lcToName = loStaff.fullName
					ENDIF

					IF !(EMPTY(lcFromName) OR EMPTY(lcToName))
						* Can only send to manager if on view page and only to employee on manage page, and only to the one listed in the request...
						IF !(;
							!llManage AND leaveRequests.employee == This.Employee AND leaveRequests.manager == lnTo;
							OR llManage AND leaveRequests.employee == lnTo AND leaveRequests.manager == This.Employee;
						  )
							This.AddError("You do not have access to this leave request!")
						ELSE
							DO CASE
								CASE lcMode == '0' AND !((leaveRequests.employee == This.Employee OR leaveRequests.manager == This.Employee) AND This.CheckRights([LR_MYSEND]))
									This.AddError("You do not have access to send this leave request message!")
								CASE lcMode == '1' AND !(leaveRequests.manager == This.Employee AND This.CheckRights([LR_ACCEPT]))
									This.AddError("You do not have access to approve this request!")
								CASE lcMode == '-1' AND !(leaveRequests.manager == This.Employee AND This.CheckRights([LR_DECLINE]))
									This.AddError("You do not have access to decline this request!")
							OTHERWISE
								lcSubject = IIF(lcMode == '1', "Leave Request Accepted", IIF(lcMode == '-1', "Leave Request Declined", Request.Form("subject")))
								IF EMPTY(lcSubject)
									This.AddValidationError("Subject cannot be blank.")
								ENDIF

								lcMessage = This.htmlsanitiser(Request.Form("comments"))
								IF EMPTY(lcMessage) AND lcMode == "-1"
									This.AddValidationError("Comments cannot be blank when declining a request.")
								ENDIF

								IF !This.IsError()
									SELECT leaveRequests

									DO CASE
										CASE lcMode == '1'
											REPLACE;
												accepted WITH .T.;
												declined WITH .F.;
												IN leaveRequests

											This.SendSiteUpdatedEmailTo(lnTo, SENDMAIL_LEAVEREQUEST_ACCEPTED, lcMessage)
										CASE lcMode == "-1"
											REPLACE;
												accepted WITH .F.;
												declined WITH .T.;
												IN leaveRequests

											This.SendSiteUpdatedEmailTo(lnTo, SENDMAIL_LEAVEREQUEST_DECLINED, lcMessage)
									ENDCASE

									* -------------------------------------------------------------------------------------------------------------------------
									* 01/03/2011  CMGM  MSI 2010.03,2011.02  TTP5722,6619  
									* If we are processing (approving or declining) leave requests, select all messages attached to the original leave request.
									* Flag all these messages to 'READ.'
									*
									* NOTE(S):
									* - The original leave request is the first one sent from EMPLOYEE to MANAGER.
									* - lcMode == '0' means 'SEND' messages only: no processing
									* -------------------------------------------------------------------------------------------------------------------------
									IF NOT (lcMode == '0')
										SELECT Id ;
										FROM leaveRequestStatus ;
										WHERE leaveRequestStatus.leaveReqId = lnId AND ;
												leaveRequestStatus.from = lnTo AND ;			&& employee
													leaveRequestStatus.to = This.Employee	;	&& manager
										INTO CURSOR curLeaveRequestStatus

										* ...Did not find anything to process
										IF EMPTY(_TALLY)
											This.AddError("Cannot find request to mark read!")
										ENDIF
							
										* 01/03/2011  CMGM  MSI 2010.03,2011.02  TTP5722,6619  ...Now mark all the messages as 'read.'
										SELECT curLeaveRequestStatus
										SCAN
											SELECT leaveRequestStatus
											REPLACE read WITH DATETIME() FOR leaveRequestStatus.ID = curLeaveRequestStatus.ID
											*	This.AddUserInfo("Message marked as read.")
										ENDSCAN
																		
										USE IN SELECT("curLeaveRequestStatus")
									ENDIF
									

									* Create new message									
									SELECT leaveRequestStatus
									APPEND BLANK
									REPLACE;
										leaveReqID	WITH lnId,;
										from		WITH This.Employee,;
										to			WITH lnTo,;
										fromName	WITH lcFromName,;
										toName		WITH lcToName,;
										subject		WITH lcSubject,;
										message		WITH lcMessage,;
										sent		WITH DATETIME();
										IN leaveRequestStatus

									This.AddUserInfo(ICASE(lcMode == '1', "Leave Approved", lcMode == "-1", "Leave Declined", "Leave") + " message sent.")
								ENDIF
							ENDCASE
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		* Isolate the request if sending a mesage or if there was errors, but not if sending an approve/decline
		IF This.IsValidationError()
			Response.Redirect(ICASE(lcMode == '1', "ApproveLeaveRequest", lcMode == "-1", "DeclineLeaveRequest", "SendLeaveRequestMessage") + "Page.si?id=-" + TRANSFORM(lnLeaveId) + This.AppendMessages("&"))
		ELSE
			Response.Redirect(IIF(llManage, "Manage", "View") + "LeaveRequestPage.si?id=" + IIF(lcMode == '0', '-', "") + TRANSFORM(lnId) + This.AppendMessages("&show=pending&"))
		ENDIF
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE DeleteLeaveRequest()
		LOCAL lnId, llManage, lcQueryString, lnPos, lcNonce

		lnId = VAL(Request.QueryString("id"))
		llManage = !EMPTY(Request.QueryString("manage"))

		* 13/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (deleting record)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "leaveRequestDays");
		  AND This.SelectData(This.Licence, "leaveRequestStatus"))
			This.AddError("Page setup failed - Cannot delete request!")
		ELSE
			SELECT leaveRequests
			LOCATE FOR id == lnId
			IF !FOUND()
				This.AddError("Cannot find request to delete!")
			ELSE
			    DO CASE
				CASE !((leaveRequests.employee == This.Employee OR leaveRequests.manager == This.Employee) AND This.CheckRights([LR_DELETE]))
					This.AddError("You do not have access to delete this request!")
				OTHERWISE
					DELETE FROM leaveRequestDays WHERE leaveReqId == lnId
					DELETE FROM leaveRequestStatus WHERE leaveReqId == lnId
					DELETE FROM leaveRequests WHERE id == lnId

					This.AddUserInfo("Leave request deleted.")
				ENDCASE
			ENDIF
		ENDIF

		lcQueryString = Request.QueryString()
		lnPos = AT("id=", lcQueryString)
		lcQueryString = LEFT(lcQueryString, lnPos - 1) + SUBSTR(lcQueryString, lnPos + LEN("&id=" + TRANSFORM(lnId)))

		Response.Redirect(IIF(llManage, "Manage", "View") + "LeaveRequestPage.si?" + lcQueryString + This.AppendMessages('&'))
	ENDPROC


	*--------------------------------------------------------------------------------*
	PROCEDURE CancelLeaveRequest()
		LOCAL lnId, llManage, lcQueryString, lnPos, lcNonce

		lnId = VAL(Request.QueryString("id"))
		llManage = !EMPTY(Request.QueryString("manage"))

		* 13/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (deleting record)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "mycancelreq")) 
			This.AddError("Page setup failed - Cannot cancel request!")
		ELSE
			SELECT leaveRequests
			LOCATE FOR id == lnId
			IF !FOUND()
				This.AddError("Cannot find request to cancel!")
			ELSE
			    DO CASE
				CASE !((leaveRequests.employee == This.Employee OR leaveRequests.manager == This.Employee))
					This.AddError("You do not have access to cancel this request!")
				OTHERWISE
                    UPDATE leaveRequests SET cancelReq=.T. WHERE id == lnId 
                    INSERT INTO mycancelreq (mlcode,mldate,cancelby) VALUES (lnId,DATETIME(),this.employee)
					This.AddUserInfo("Cancel Leave has been requested")
				ENDCASE
			ENDIF
		ENDIF

		lcQueryString = Request.QueryString()
		lnPos = AT("id=", lcQueryString)
		lcQueryString = LEFT(lcQueryString, lnPos - 1) + SUBSTR(lcQueryString, lnPos + LEN("&id=" + TRANSFORM(lnId)))

		Response.Redirect(IIF(llManage, "Manage", "View") + "LeaveRequestPage.si?" + lcQueryString + This.AppendMessages('&'))
	ENDPROC

	*--------------------------------------------------------------------------------*



	*--------------------------------------------------------------------------------*

	PROCEDURE ToggleMessageRead()
		LOCAL lnId, lnLeaveId, llManage, lcQueryString, lnPos

		lnId = VAL(Request.QueryString("id"))
		llManage = !EMPTY(Request.QueryString("manage"))

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "leaveRequests");
		  AND This.SelectData(This.Licence, "leaveRequestStatus"))
			This.AddError("Page setup failed - Cannot delete request!")
		ELSE
			SELECT leaveRequestStatus
			LOCATE FOR id == lnId
			IF !FOUND()
				This.AddError("Cannot find request message to mark read!")
			ELSE
				lnLeaveId = leaveRequestStatus.leaveReqId

				SELECT leaveRequests
				LOCATE FOR id == lnLeaveId
				IF !FOUND()
					This.AddError("Cannot find request to mark read!")
				ELSE
					IF !(leaveRequests.employee == This.Employee OR leaveRequests.manager == This.Employee)
						This.AddError("You do not have access to this request!")
					ELSE
						SELECT leaveRequestStatus
						IF EMPTY(read)
							REPLACE read WITH DATETIME() FOR id == lnId

							This.AddUserInfo("Message marked as read.")
						ELSE
							REPLACE read WITH {} FOR id == lnId

							This.AddUserInfo("Message marked as unread.")
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		lcQueryString = Request.QueryString()
		lnPos = AT("id=", lcQueryString)
		lcQueryString = LEFT(lcQueryString, lnPos - 1) + SUBSTR(lcQueryString, lnPos + LEN("&id=" + TRANSFORM(lnId)))
		lcQueryString = lcQueryString + "&id=" + TRANSFORM(lnLeaveId)

		Response.Redirect(IIF(llManage, "Manage", "View") + "LeaveRequestPage.si?" + lcQueryString + This.AppendMessages("&"))
	ENDPROC

	*================================================================================*

	PROCEDURE DeleteReport()
		LOCAL lcFileToDelete, lcNonce

		lcFileToDelete = ALLTRIM(Request.QueryString("delete"))

		* 13/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (deleting record)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		IF EMPTY(lcFileToDelete)
			This.AddError("Nothing to delete!")
		ELSE
			** ChrisF:	made the delete link more secure.
			**			now it looks the same as the view link - it hides the This.Employee part.
			DELETE FILE (This.CompanyDataPath() + ADDBS("payslips") + TRANSFORM(This.Employee) + '_' + lcFileToDelete + ".pdf")
			This.AddUserInfo("Report has been deleted.")
		ENDIF

		Response.Redirect("ReportsPage.si" + This.AppendMessages('?'))
	ENDPROC

	*================================================================================*

	PROCEDURE SendMessage()
		LOCAL loMessageIn, lnMeToId, lcMeSubject, lcMeMessage, loStaff, lcMeToName, lcMeFromName, loMessageOut

		loMessageIn	= Factory.GetMessagesObject()
		loMessageOut = Factory.GetMessagesObject()
		loStaff		= Factory.GetStaffObject()
		lnMeToId	= VAL(Request.Form("metoid"))
		lcMeSubject = Request.Form("mesubject")
		lcMeMessage = Request.Form("memessage")

		lcMeToName	= ""
		
		LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		
		IF loStaff.Load(lnMeToId)
			lcMeToName = loStaff.fullName
		ENDIF

		lcMeFromName = ""
		IF loStaff.Load(This.Employee)
			lcMeFromName = loStaff.fullName
		ENDIF

		IF loMessageIn.GetBlankRecord() AND loMessageOut.GetBlankRecord()
			loMessageIn.oData.meToId		= lnMeToId
			loMessageIn.oData.meToName		= lcMeToName
			loMessageIn.oData.meFromId		= This.Employee
			loMessageIn.oData.meFromName	= lcMeFromName
			loMessageIn.oData.meSubject		= this.htmlsanitiser(lcMeSubject)
			loMessageIn.oData.meMessage		= this.htmlsanitiser(lcMeMessage)
			loMessageIn.oData.meDate		= This.GetLocalTime(DATETIME())
			loMessageIn.oData.meType		= "IN"
			loMessageIn.CopyTo(loMessageOut)
			loMessageOut.oData.meType		= "OUT"

			IF loMessageIn.Save() AND loMessageOut.Save()
				This.AddUserInfo("Message sent successfully.")
				This.SendSiteUpdatedEmailTo(lnMeToId, SENDMAIL_MESSAGE)
			ELSE
				This.AddError("Message not sent.")
			ENDIF
		ELSE
			This.AddError("Could not create new message.")
		ENDIF

		Response.Redirect("InboxPage.si" + This.AppendMessages('?'))
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE SendNews()
		LOCAL loMessage, lcMeSubject, lcMeMessage, loStaff, lcMeFromName

		loMessage	= Factory.GetMessagesObject()
		loStaff		= Factory.GetStaffObject()
		lcMeSubject = Request.Form("mesubject")
		lcMeMessage = Request.Form("memessage")
		lcMeFromName = ""

		LOCAL lcNonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF
		
		IF loStaff.Load(This.Employee)
			lcMeFromName = loStaff.fullName
		ENDIF

		IF loMessage.GetBlankRecord()
			loMessage.oData.meFromId	= This.Employee
			loMessage.oData.meFromName	= lcMeFromName
			loMessage.oData.meSubject	= this.htmlsanitiser(lcMeSubject)
			loMessage.oData.meMessage	= this.htmlsanitiser(lcMeMessage)
			loMessage.oData.meDate		= This.GetLocalTime(DATETIME())
			loMessage.oData.meType		= "News"

			IF loMessage.Save()
				This.AddUserInfo("News sent successfully.")

				loStaff.Execute([SELECT myWebCode FROM myStaff INTO CURSOR tmpMail])
				IF USED("tmpMail")
					SELECT tmpMail
					SCAN
						This.SendSiteUpdatedEmailTo(tmpMail.myWebCode, SENDMAIL_NEWS)
					ENDSCAN
				ENDIF
			ELSE
				This.AddError("News not sent.")
			ENDIF
		ELSE
			This.AddError("Could not create news message: " + loMessage.cError)
		ENDIF

		Response.Redirect("CompanyPage.si" + This.AppendMessages('?'))
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ReplyMessage()
		LOCAL lnMessageId, loMessage, lcNonce

		* 14/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (replying messages)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnMessageId = VAL(This.URLUnEscape(Request.QueryString("id")))
		IF EMPTY(lnMessageId)
			This.AddError("No message to reply to!")
			This.SendMessagePage()
			RETURN
		ENDIF

		loMessage = Factory.GetMessagesObject()

		IF loMessage.ConstructReply(lnMessageId)
			loMessage.oData.meDate = This.GetLocalTime(DATETIME())

			This.SendMessagePage(loMessage, "Reply")
		ELSE
			This.AddError("Could not reply to message: " + loMessage.cError)
			This.SendMessagePage()
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE ForwardMessage()
		LOCAL lnMessageId, loMessage, lcNonce

		* 14/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (forwarding messages)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnMessageId = VAL(This.URLUnEscape(Request.QueryString("id")))
		IF EMPTY(lnMessageId)
			This.AddError("No message to forward!")
			This.SendMessagePage()
			RETURN
		ENDIF

		loMessage = Factory.GetMessagesObject()

		IF loMessage.ConstructForward(lnMessageId)
			loMessage.oData.meDate = This.GetLocalTime(DATETIME())

			This.SendMessagePage(loMessage, "Forward")
		ELSE
			This.AddError("Could not forward message: " + loMessage.cError)
			This.SendMessagePage()
		ENDIF
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	PROCEDURE DeleteMessage()
		LOCAL lnMessageId, lcDestination, loMessage, lcNonce

		* 13/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (deleting record)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnMessageId = VAL(Request.QueryString("meId"))
		lcDestination = This.URLUnEscape(Request.QueryString("fromPage"))
		loMessage = Factory.GetMessagesObject()

		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed.")
		ELSE
			IF !EMPTY(lnMessageId)
				IF loMessage.Load(lnMessageId)
					IF This.CheckRights("can_delete")
						IF loMessage.Delete()
							This.AddUserInfo("Message deleted.")
						ELSE
							This.AddError("Could not delete message.")
						ENDIF
					ELSE
						This.AddError("Not allowed to Delete this message.")
					ENDIF
				ELSE
					This.AddError("No message to Delete: " + loMessage.cError)
				ENDIF
			ELSE
				This.AddError("No message to Delete.")
			ENDIF
		ENDIF

		IF !EMPTY(lcDestination)
			Response.Redirect(lcDestination + This.AppendMessages(IIF('?' $ lcDestination, '&', '?')))
		ELSE
			Response.Redirect("InboxPage.si" + This.AppendMessages('?'))
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE DeleteNewsItem()
		LOCAL lnMessageId, lcDestination, loMessage, lcNonce

		* 13/07/2011  CMGM  2011.03  Stratsec APP-06  Force user to EITHER authenticate again OR add nonce before performing sensitive actions (deleting record)
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnMessageId = VAL(Request.QueryString("meId"))
		lcDestination = This.URLUnEscape(Request.QueryString("fromPage"))
		loMessage = Factory.GetMessagesObject()

		IF !This.SelectData(This.Licence, "myStaff")
			This.AddError("Page Setup Failed.")
		ELSE
			IF !EMPTY(lnMessageId)
				IF loMessage.Load(lnMessageId)
					IF This.CheckRights("NEWS_DELETE")
						IF loMessage.Delete()
							This.AddUserInfo("News item deleted.")
						ELSE
							This.AddError("Could not delete news item.")
						ENDIF
					ELSE
						This.AddError("Not allowed to delete this news item.")
					ENDIF
				ELSE
					This.AddError("No news item to Delete: " + loMessage.cError)
				ENDIF
			ELSE
				This.AddError("No news item to delete.")
			ENDIF
		ENDIF

		IF !EMPTY(lcDestination)
			Response.Redirect(lcDestination + This.AppendMessages(IIF('?' $ lcDestination, '&', '?')))
		ELSE
			Response.Redirect("CompanyPage.si" + This.AppendMessages('?'))
		ENDIF
	ENDPROC

	*================================================================================*

	PROCEDURE UnapproveTimesheetEntry(tlApprove)
		LOCAL lnId, lnCurrentStaff, lcFrom, lcValue
		LOCAL lcNonce, llValidNonce

		*!* 17/03/2011  CMGM  MSI 2011.02  TTP6637  Added nonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnId = VAL(Request.QueryString("edit"))

		IF !This.SelectData(This.Licence, "timesheet")
			This.AddError("Page Setup Failed.")
		ELSE
			SELECT timesheet
			LOCATE FOR tsId == lnId
			IF !FOUND()
				This.AddError("Entry not found!")
			ELSE
				lnCurrentStaff = VAL(Request.QueryString("currentStaff"))

				IF timesheet.tsEmp != lnCurrentStaff AND lnCurrentStaff != EVERYONE_OPTION	&& If we are looking at Everyone, we can't check this anyway.
					This.AddError("Incorrect entry ownership!")
				ELSE
					This.UnapproveTimesheetEmp(tlApprove, "tsId == " + TRANSFORM(lnId))
					RETURN
				ENDIF
			ENDIF
		ENDIF

		* If we got here it's due to an error, so no point retaining anything.
		lcFrom = Request.QueryString("from")

		DO CASE
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + This.AppendMessages('?'))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + This.AppendMessages('?'))
		ENDCASE
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE ApproveTimesheetEntry()
		This.UnapproveTimesheetEntry(.T.)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE UnapproveTimesheetEmp(tlApprove, tcFilter)
		LOCAL lnCurrentStaff, lcFrom

		IF VARTYPE(tcFilter) != 'C'
			tcFilter = ""
		ENDIF

		lnCurrentStaff = VAL(Request.QueryString("currentStaff"))

		IF EMPTY(lnCurrentStaff)
			This.AddError("No Employee selected!")
		ELSE
			IF !This.CheckAccess(lnCurrentStaff, This.IsManager(This.Employee), .T.)	&& Allow Everyone option
				This.AddError("You do not have access to this page.")
			ELSE
				IF lnCurrentStaff != EVERYONE_OPTION	&& If this is -1 (Everyone), then the ID check that already has been set by [Un]ApproveTimesheetEntry is enough, as the button for calling this directly goes away when viewing Everyone.
					tcFilter = tcFilter + IIF(EMPTY(tcFilter), "", " AND ") + "tsEmp == " + TRANSFORM(lnCurrentStaff)
				ENDIF
				This.UnapproveTimesheetGroup(tlApprove, tcFilter)
				RETURN
			ENDIF
		ENDIF

		* If we got here it's due to an error, so no point retaining anything.
		lcFrom = Request.QueryString("from")

		DO CASE
			CASE lcFrom == "gs"
				Response.Redirect("GroupSummaryPage.si" + This.AppendMessages('?'))
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + This.AppendMessages('?'))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + This.AppendMessages('?'))
		ENDCASE
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE ApproveTimesheetEmp()
		This.UnapproveTimesheetEmp(.T.)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE UnapproveTimesheetGroup(tlApprove, tcFilter)
		LOCAL lnCurrentGroup, lnCurrentPay, lnPos, llIsError, lcFrom, lcValue
		LOCAL lcNonce, llValidNonce

		*!* 17/03/2011  CMGM  MSI 2011.02  TTP6637  Added nonce
		lcNonce = Request.QueryString("nonce")
		llValidNonce = This.NonceIsValid(lcNonce)
		IF NOT llValidNonce
			This.AddUserInfo("Page access has expired.")
			This.LogOut()
			RETURN
		ENDIF

		lnCurrentPay = VAL(Request.QueryString("currentPay"))
		lnCurrentGroup = VAL(Request.QueryString("currentGroup"))

		IF !This.IsManager(This.Employee)
			This.AddError("You do not have access to this function!")
		ELSE
			IF !This.CheckRights(IIF(EMPTY(tcFilter), "TS_GROUP", "TS_EMPLOYEE") + IIF(tlApprove, '_A', '_D'))
				This.AddError("You do not have access to " + IIF(tlApprove, "Approve", "Unapprove") + " this " + IIF(EMPTY(tcFilter), "Group", "Employee") + "!")
			ELSE
				IF !This.GetGroupsForManager(This.Employee, "curGroups")
					This.AddError("Page Setup Failed!")
				ELSE
					SELECT curGroups
					LOCATE FOR grCode == lnCurrentGroup
					IF !FOUND() AND lnCurrentGroup != MY_DETAILS_GROUP
						This.AddError("Cannot find Employee record!")
					ELSE
						IF !(This.SelectData(This.Licence, "myPays");
						  AND This.SelectData(This.Licence, "timesheet"))
							This.AddError("Page Setup Failed.")
						ELSE
							SELECT myPays
							LOCATE FOR pay_type == 2 AND pay_status == 1 AND pay_pk == lnCurrentPay
							IF !FOUND()
								This.AddError("Selected Pay not found or not open!")
							ELSE
								IF !This.GetEmployeesByGroupCode(lnCurrentGroup, "curStaff")
									This.AddError("Cannot Load Group!")
								ELSE
									IF VARTYPE(tcFilter) != 'C'
										tcFilter = ".T."
									ENDIF
									
									LOCAL tspayFilter
								    tspayfilter = " tspay = " + ALLTRIM(STR(lncurrentpay,20))
								    
									UPDATE timesheet SET tsApproved = tlApprove;
										WHERE &tcFilter.;
										AND !tsDownload;
										AND tsEmp IN (SELECT myWebCode FROM curStaff) AND &tspayfilter

									This.AddUserInfo(IIF("tsId" $ tcFilter, "", "All ") + IIF(tcFilter == ".T.", "Group", "Employee") + IIF("tsId" $ tcFilter, " entry ", "'s entries ") + IIF(tlApprove, "Approved", "Unapproved") + ".")
								ENDIF
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentPay", lnCurrentPay)
		loRetainList.SetEntry("currentGroup", lnCurrentGroup)

		IF !EMPTY(tcFilter)	&& Need this check as when called directly it's not defined!
			lnPos = AT("tsEmp == ", tcFilter)
			IF lnPos > 0
				* If we have a specific employee we are dealing with, that is the one we want to go back to looking at
				loRetainList.SetEntry("currentStaff", TRANSFORM(VAL(SUBSTR(tcFilter, lnPos + 9))))
			ELSE
				* Otherwise, get it from the URL as per usual (covers the case of Everyone)
				loRetainList.SetEntry("currentStaff", TRANSFORM(VAL(Request.QueryString("currentStaff"))))
			ENDIF
		ELSE
			* Otherwise, get it from the URL as per usual (covers the case of Everyone)
			loRetainList.SetEntry("currentStaff", TRANSFORM(VAL(Request.QueryString("currentStaff"))))
		ENDIF

		lcFrom = Request.QueryString("from")

		IF !(lcFrom == "gs")
			lcValue = Request.QueryString("approved")
			IF INLIST(lcValue, "yes", "no", "both")
				loRetainList.SetEntry("approved", lcValue)
			ENDIF

			IF lcFrom == "history"
				loRetainList.SetEntry("startDate", TRANSFORM(CTOD(Request.QueryString("startDate"))))
				loRetainList.SetEntry("endDate", TRANSFORM(CTOD(Request.QueryString("endDate"))))

				lcValue = Request.QueryString("open")
				IF INLIST(lcValue, "open", "closed", "both")
					loRetainList.SetEntry("open", lcValue)
				ENDIF

				lcValue = Request.QueryString("downloaded")
				IF INLIST(lcValue, "yes", "no", "both")
					loRetainList.SetEntry("downloaded", lcValue)
				ENDIF
			ENDIF
		ENDIF

		lcValue = Request.QueryString("type")

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		DO CASE
			CASE lcFrom == "gs"
				Response.Redirect("GroupSummaryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcValue)), '#' + lcValue + "Entries", ""))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcValue)), '#' + lcValue + "Entries", ""))
		ENDCASE
	ENDPROC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	PROCEDURE ApproveTimesheetGroup()
		This.UnapproveTimesheetGroup(.T.)
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE DeleteTimeRow()
		LOCAL lcFrom, lnId, lnCurrentStaff, lnCurrentPay, loRetainList, lcValue, llIsError

		lnCurrentPay = VAL(Request.QueryString("currentPay"))
		lnCurrentStaff = EVL(VAL(Request.QueryString("currentStaff")), This.Employee)

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lnId = VAL(Request.QueryString("edit"))
			IF EMPTY(lnId)
				This.AddError("No entry to delete!")
			ELSE

				SELECT timesheet
				LOCATE FOR tsId == lnId AND tsPay == lnCurrentPay AND (tsEmp == lnCurrentStaff OR lnCurrentStaff == EVERYONE_OPTION)	&& ...for current emp if or Everyone
				IF !FOUND()
					This.AddError("Cannot find entry to delete!")
				ELSE
					IF tsDownload
						This.AddError("Cannot delete downloaded entry.")
					ELSE
						IF !This.CheckAccess(lnCurrentStaff, This.IsManager(This.Employee), .T.)	&& Allow Everyone option
							This.AddError("You do not have access to this page.")
						ELSE
							SELECT timesheet
							IF tsEmp != lnCurrentStaff AND lnCurrentStaff != EVERYONE_OPTION	&& If we are looking at Everyone, we can't check this anyway
								This.AddError("You do not have access to that entry!")
							ELSE
								DELETE FROM timesheet WHERE tsId == lnId

								This.AddUserInfo("Entry deleted.")
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentPay", lnCurrentPay)
		loRetainList.SetEntry("currentGroup", VAL(Request.QueryString("currentGroup")))
		loRetainList.SetEntry("currentStaff", lnCurrentStaff)

		lcFrom = Request.Form("from")

		IF !(lcFrom == "gs")
			lcValue = Request.QueryString("approved")
			IF INLIST(lcValue, "yes", "no", "both")
				loRetainList.SetEntry("approved", lcValue)
			ENDIF

			IF lcFrom == "history"
				loRetainList.SetEntry("startDate", TRANSFORM(CTOD(Request.QueryString("startDate"))))
				loRetainList.SetEntry("endDate", TRANSFORM(CTOD(Request.QueryString("endDate"))))

				lcValue = Request.QueryString("open")
				IF INLIST(lcValue, "open", "closed", "both")
					loRetainList.SetEntry("open", lcValue)
				ENDIF

				lcValue = Request.QueryString("downloaded")
				IF INLIST(lcValue, "yes", "no", "both")
					loRetainList.SetEntry("downloaded", lcValue)
				ENDIF
			ENDIF
		ENDIF

		lcValue = Request.QueryString("type")

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		DO CASE
			CASE lcFrom == "gs"
				Response.Redirect("GroupSummaryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcValue)), '#' + lcValue + "Entries", ""))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcValue)), '#' + lcValue + "Entries", ""))
		ENDCASE
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE SaveTimeEntry()
		LOCAL lcType, loType, loTypes, loRetainList, lcFrom, lcValue, lnId, llManager, llIsError, lnCurrentStaff, lnCurrentGroup, lnRowCount
		LOCAL lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, lnRateCode, loCostCentCode, loJobCode

		llManager = This.IsManager(This.Employee)

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lcType = Request.Form("type")
			loTypes = This.GetTimesheetTypes(.F.)
			IF EMPTY(loTypes.GetKey(lcType))
				This.AddError("Unknown entry type!")
			ELSE
				loType = loTypes.Item(lcType)

				lnId = VAL(Request.Form("edit"))

				lnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
				lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)

				SELECT timesheet
				LOCATE FOR tsId == lnId
				IF !FOUND()
					This.AddError("Entry not found!")
				ELSE
					lnStaff			= This.Employee
					ldDate			= {}
					loLeaveCode		= null
					loOtherCode		= null
					loAllowCode		= null
					ltStart			= {}
					ltEnd			= {}
					ltBreak			= {}
					lnUnits			= 0
					lnReduce		= 0
					loWageCode		= null
					lnRateCode		= 0
					loCostCentCode	= null
					loJobCode		= null

					* If lnCurrentStaff == Everyone, we must pass the current user for use in looking up LeaveCodes.
					This.CollectTimeEntryFormData(loType, "", IIF(lnCurrentStaff == EVERYONE_OPTION, This.Employee, lnCurrentStaff), llManager, @lnStaff, @ldDate, @loLeaveCode, @loOtherCode, @loAllowCode, @ltStart, @ltEnd, @ltBreak, @lnUnits, @lnReduce, @loWageCode, @lnRateCode, @loCostCentCode, @loJobCode)
					This.SaveSingleTimeEntry(loType, lnId, llManager, lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, lnRateCode, loCostCentCode, loJobCode, lnCurrentGroup, @lnRowCount)
				ENDIF
			ENDIF
		ENDIF

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentPay", TRANSFORM(VAL(Request.Form("currentPay"))))
		loRetainList.SetEntry("currentGroup", TRANSFORM(EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)))
		loRetainList.SetEntry("currentStaff", TRANSFORM(EVL(VAL(Request.Form("currentStaff")), This.Employee)))

		lcValue = Request.Form("approved")
		IF INLIST(lcValue, "yes", "no", "both")
			loRetainList.SetEntry("approved", lcValue)
		ENDIF

		lcFrom = Request.Form("from")

		IF lcFrom == "history"
			loRetainList.SetEntry("startDate", TRANSFORM(CTOD(Request.Form("startDate"))))
			loRetainList.SetEntry("endDate", TRANSFORM(CTOD(Request.Form("endDate"))))

			lcValue = Request.QueryString("open")
			IF INLIST(lcValue, "open", "closed", "both")
				loRetainList.SetEntry("open", lcValue)
			ENDIF

			lcValue = Request.Form("downloaded")
			IF INLIST(lcValue, "yes", "no", "both")
				loRetainList.SetEntry("downloaded", lcValue)
			ENDIF
		ENDIF

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		DO CASE
			CASE lcFrom == "history"
				Response.Redirect("TimeHistoryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcType)), '#' + lcType + "Entries", ""))
			OTHERWISE
				Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcType)), '#' + lcType + "Entries", ""))
		ENDCASE
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE SaveNewTimeEntries()
		LOCAL loTypes, lcType, loType, lcValue, lnCount, lnI, lnSaved, llManager, lnCurrentPay, llIsError, lnErrored, lnCurrentStaff, lnCurrentGroup, lnRowCount
		LOCAL lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, lnRateCode, loCostCentCode, loJobCode

		llManager = This.IsManager(This.Employee)
		lnErrored = 0
		loRetainList = Factory.GetRetainListObject()

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lcType = Request.Form("addType")
			loTypes = This.GetTimesheetTypes(.F.)
			IF EMPTY(loTypes.GetKey(lcType))
				This.AddError("Unknown entry type!")
				loType = null
			ELSE
				loType = loTypes.Item(lcType)

				lnCount = VAL(Request.Form("count"))
				lnCurrentPay = VAL(Request.Form("currentPay"))

				lnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
				lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)

				IF lnCount < 1
					This.AddError("Nothing to save!")
				ELSE
					lnSaved = 0

					FOR lnI = 1 TO lnCount
						lnStaff			= This.Employee
						ldDate			= {}
						loLeaveCode		= null
						loOtherCode		= null
						loAllowCode		= null
						ltStart			= {}
						ltEnd			= {}
						ltBreak			= {}
						lnUnits			= 0
						lnReduce		= 0
						loWageCode		= null
						lnRateCode		= 0
						loCostCentCode	= null
						loJobCode		= null

						* If lnCurrentStaff == Everyone, we must pass the current user for use in looking up LeaveCodes.
						This.CollectTimeEntryFormData(loType, lnI, IIF(lnCurrentStaff == EVERYONE_OPTION, This.Employee, lnCurrentStaff), llManager, @lnStaff, @ldDate, @loLeaveCode, @loOtherCode, @loAllowCode, @ltStart, @ltEnd, @ltBreak, @lnUnits, @lnReduce, @loWageCode, @lnRateCode, @loCostCentCode, @loJobCode)
						IF This.SaveSingleTimeEntry(loType, -lnCurrentPay, llManager, lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, lnRateCode, loCostCentCode, loJobCode, lnCurrentGroup, @lnRowCount)
							lnSaved = lnSaved + lnRowCount
						ELSE
							lnErrored = lnErrored + 1
							* Removed as the URL gets too long...
							* This.RetiainTimeEntry(loType, loRetainList, lnErrored, llManager, lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, loCostCentCode, loJobCode)
						ENDIF
					NEXT
				ENDIF
			ENDIF
		ENDIF

		IF !ISNULL(loType)
			IF lnSaved == 1
				This.AddUserInfo("1 " + loType.title + " Entry Saved.")
			ELSE
				IF lnSaved > 0
					This.AddUserInfo(TRANSFORM(lnSaved) + ' ' + loType.title + " Entries Saved.")
				ENDIF
			ENDIF
		ENDIF

		loRetainList.SetEntry("currentPay", TRANSFORM(lnCurrentPay))
		loRetainList.SetEntry("currentGroup", TRANSFORM(EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)))
		loRetainList.SetEntry("currentStaff", TRANSFORM(EVL(VAL(Request.Form("currentStaff")), This.Employee)))
		loRetainList.SetEntry("addType", lcType)
		loRetainList.SetEntry("count", "0")	&& Removed TRANSFORM(lnErrored) as the URL gets too long.

		lcValue = Request.Form("approved")
		IF INLIST(lcValue, "yes", "no", "both")
			loRetainList.SetEntry("approved", lcValue)
		ENDIF

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcType)), '#' + lcType + "Entries", ""))
	ENDPROC


*--------------------------------------------------------------------------------*
FUNCTION GeneratePDF(tcReportFile) as String
	LOCAL loSession, lnRetVal, lcFileName, lcPath
	lcPath = UPPER(this.licence)+"\REPORTS\"
	
	lcFileName = FORCEEXT(lcPath+"Report" + SYS(2015),"PDF")
	loSession = EVALUATE([xfrx("XFRX#INIT")])
	lnRetVal = loSession.SetParams(lcFileName, , .T., , , ,"PDF")
	
	IF lnRetVal = 0
	    loSession.ProcessReport(tcReportFile)
		loSession.finalize()
    ELSE
		lcFileName = ""
    ENDIF

	RETURN (lcFileName)
ENDFUNC

*################################################################################*
#DEFINE TOC_MessageUtils_

	*> +define: MessageUtils
	* Append an info message.
	PROCEDURE AddUserInfo(tcInfo, tcJoiner)
		IF !(VARTYPE(tcJoiner) == 'C')
			tcJoiner = MESSAGE_JOINER
		ENDIF

		IF EMPTY(This.cInfo)
			This.cInfo = tcInfo
		ELSE
			This.cInfo = This.cInfo + tcJoiner + tcInfo
		ENDIF

		* 09/03/2011  CMGM  MSI 2011.02  TTP6638  Save the message in a session variable (AppendMessages() clears the property)
		Session.SetSessionVar("cInfo",This.cInfo)
	ENDPROC

	
	*--------------------------------------------------------------------------------*

	* Replace the info message.  All calls to this should be documented as to why!
	PROCEDURE ReplaceUserInfo(tcInfo)
		This.cInfo = tcInfo
	ENDPROC

	*--------------------------------------------------------------------------------*

	* Append an error message.
	PROCEDURE AddError(tcError, tcErrorField, tcJoiner)
		IF !(VARTYPE(tcJoiner) == 'C')
			tcJoiner = MESSAGE_JOINER
		ENDIF

		IF VARTYPE(tcErrorField) == 'C' AND !EMPTY(tcErrorField)
			This.cErrorField = tcErrorField
		ENDIF

		IF EMPTY(This.cError)
			This.cError = tcError
		ELSE
			This.cError = This.cError + tcJoiner + tcError
		ENDIF

		* 09/03/2011  CMGM  MSI 2011.02  TTP6638  Save the message in a session variable (AppendMessages() clears the property)
		Session.SetSessionVar("cError",This.cError)
	ENDPROC

	*--------------------------------------------------------------------------------*

	* Replace the error message.  All calls to this should be documented as to why!
	PROCEDURE ReplaceError(tcError)
		This.cError = tcError
	ENDPROC

	*--------------------------------------------------------------------------------*

	PROCEDURE AddValidationError(tcError)
		This.AddError(MESSAGE_VALIDATION_PREFIX + tcError)
	ENDPROC

	*--------------------------------------------------------------------------------*

	FUNCTION IsUserInfo() as Boolean
		RETURN !EMPTY(This.cInfo)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION IsError() as Boolean
		RETURN !EMPTY(This.cError)
	ENDFUNC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	FUNCTION IsFatalError() as Boolean
		LOCAL lnI
		LOCAL ARRAY laErrors[1]

		FOR lnI = 1 TO ALINES(laErrors, STRTRAN(This.cError, MESSAGE_TRIM, ""), 0, MESSAGE_SPLIT_CHAR)
			IF !EMPTY(laErrors[lnI]) AND AT(MESSAGE_VALIDATION_PREFIX, laErrors[lnI]) == 0
				RETURN .T.
			ENDIF
		NEXT

		RETURN .F.
	ENDFUNC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	FUNCTION IsValidationError() as Boolean
		RETURN This.IsError() AND !This.IsFatalError()
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION GetErrorFieldName() as String
		LOCAL lcErrField as String

		lcErrField = This.cErrorField
		This.cErrorField = ""

		RETURN lcErrField
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> HTML
	* For putting the messages into a page
	FUNCTION CheckForMessages(tcErrorCSSClass as String, tcInfoCSSClass as String, tlSuppressJavascript as Boolean) as String
		LOCAL lcOutput as String
		LOCAL lcQueryString as String
		LOCAL lcSessionMessage as String

		lcOutput = ""
		lcQueryString = ""
		lcSessionMessage = ""

		IF !((EMPTY(This.cError) AND EMPTY(This.cInfo)) OR tlSuppressJavascript)
			* if we are going to output the toUserMessages div due to current-page messages, soak up QueryString ones too, so we don't double-up the IDs
			IF !EMPTY(Request.QueryString("errorMessage"))
				This.AddError(This.HTMLEscape(Request.QueryString("errorMessage")))
			ENDIF
			IF !EMPTY(Request.QueryString("userInfo"))
				This.AddUserInfo(This.HTMLEscape(Request.QueryString("userInfo")))
			ENDIF
		ENDIF

		* Check for outstanding error conditions.
		IF !EMPTY(This.cError)

			This.cError = STRTRAN(This.cError, "VALIDATION:", "")

			** The close link requires javascript so use javascript to output it...
			lcOutput = lcOutput + [<table id="] + tcErrorCSSClass + ["><tr><td class="i"><img src="assets/messages/error.png";
				 alt="warning"/></td><td role="alert">] + This.HTMLEscape2(This.cError) + [</td>]
			IF !tlSuppressJavascript
				lcOutput = lcOutput + [<td class="x"><script type="text/javascript">] + CRLF;
					+ CHR(9) + [document.write('<' + 'a href="javascript:closeMessage(\'] + tcErrorCSSClass + [\')" class="xx"><' + '/a>');] + CRLF;
					+ [</script></td>] + CRLF
			ENDIF
			lcOutput = lcOutput + [</tr></table>] + CRLF
			This.cError = ""
		ENDIF

		* Check for outstanding information, warnings etc.
		IF !EMPTY(This.cInfo)
			lcClickUrl = ""
			IF TYPE("lcURLCsv") == "C"
				IF NOT EMPTY("cURLCsv")
					lcClickUrl = [ - <a href="]+lcURLCsv+[">Click Here To Download The File</a>]
				ENDIF
			ENDIF
								
			IF TYPE("lcURLpdf") == "C"
				IF NOT EMPTY("cURLpdf")
					lcClickUrl = [ - <a href="]+lcURLpdf+[">Click Here To Download The File</a>]
				ENDIF
			ENDIF

			** The close link requires javascript so use javascript to output it...
			lcOutput = lcOutput + [<table id="] + tcInfoCSSClass + ["><tr><td class="i"><img src="assets/messages/userInfo.png";
					 alt="info"/></td><td role="status">] + This.HTMLEscape2(This.cInfo)+ lcClickURL + [</td>]
			IF !tlSuppressJavascript
				lcOutput = lcOutput + [<td class="x"><script type="text/javascript">] + CRLF;
					+ CHR(9) + [document.write('<' + 'a href="javascript:closeMessage(\'] + tcInfoCSSClass + [\')" class="xx"><' + '/a>');] + CRLF;
					+ [</script></td>] + CRLF
			ENDIF
			lcOutput = lcOutput +[</tr></table>]

			This.cInfo = ""
		ENDIF

		* If there is output to be displayed - package it correctly.
		IF !EMPTY(lcOutput)
			lcOutput = [<div id="toUserMessages"] + IIF(tlSuppressJavascript, [ style="border:2px dashed red;padding:0px;clear: both;"], "") + [>] + lcOutput + [</div>]
			IF tlSuppressJavascript
				* not entirely sure why this is needed...
				lcOutput = [&nbsp;] + lcOutput
			ENDIF
		ELSE
			* 09/03/2011  CMGM  MSI 2011.02  TTP6638  Before displaying the message, we need to make sure that what was originally appended in the URL
			*                                         has not changed by comparing it with the session variable.  Need to checked errors first then infos.
			lcQueryString = Request.QueryString("errorMessage")
			lcSessionMessage = Session.GetSessionVar("cError")
			IF EMPTY(lcQueryString)
				lcQueryString = Request.QueryString("userInfo")
				lcSessionMessage = Session.GetSessionVar("cInfo")
			ENDIF
			
			IF !tlSuppressJavascript
				IF lcQueryString == lcSessionMessage
					* no current-page messages so let the javascript code handle QueryString ones if present.
					lcOutput = [			<script type="text/javascript">] + CRLF;
							 + [				CheckForMessageText();] + CRLF;
							 + [			</script>] + CRLF
				ENDIF
			ENDIF
		ENDIF
		
		RETURN lcOutput
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* For putting the messages onto a URL
	FUNCTION AppendMessages(tcJoiner)
		LOCAL lcOutput, lcJoiner

		lcOutput = ""
		lcJoiner = tcJoiner

		IF !EMPTY(This.cError)
			This.cError = STRTRAN(This.cError, "VALIDATION:", "")

			lcOutput = lcJoiner + "errorMessage=" + This.URLEscape(This.cError)
			This.cError = ""
			lcJoiner = "&"
		ENDIF

		IF !EMPTY(This.cInfo)
			lcOutput = lcOutput + lcJoiner + "userInfo=" + This.URLEscape(This.cInfo)
			This.cInfo = ""
		ENDIF

		RETURN lcOutput
	ENDFUNC

	* For putting the messages onto a URL
	FUNCTION AppendMessagesLink(tcJoiner)
		LOCAL lcOutput, lcJoiner

		lcOutput = ""
		lcJoiner = tcJoiner

		IF !EMPTY(This.cInfo)
			lcOutput = lcOutput + lcJoiner + "userInfo=" + ALLTRIM(This.cInfo)
			This.cInfo = ""
		ENDIF

		RETURN lcOutput
	ENDFUNC



	*################################################################################*
#DEFINE TOC_Utils_

	*> +define: Utils

	* Make the constant readable elsewhere...
	FUNCTION AccessKeysEnabled() AS Boolean
		RETURN ACCESSKEYS_ENABLED
	ENDFUNC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	* Manages a page's global accessKey list to avoid doubleups.  Returns false if there is a clash or if accessKeys are disabled.
	* - tcItem must be a uniquely-identifying string for the item getting the accessKey.
	* - if tlHidden is true, this item is not currently visible (e.g. modal content) so will always allocate, but the lookup from the key to the item can be overridden by a nonHidden item.
	* - tlSuppressError is for internal use only and should not be passed.
	FUNCTION SetAccessKey(tcItem, tcKey, tlHidden, tlSuppressError) AS Boolean
		IF !ACCESSKEYS_ENABLED
			RETURN .F.
		ENDIF

		* Check if the key is already used...
		IF This.oAccessKeyList.GetKey(tcKey) == 0
			* If this item already had a key, remove it first:
			IF This.oAccessItemList.GetKey(tcItem) != 0
				This.oAccessItemList.Remove(tcItem)
			ENDIF

			* Store the linkages:
			This.oAccessKeyList.Add(tcItem,			tcKey)
			This.oAccessItemList.Add(tcKey,			tcItem)
			This.oAccessHiddenList.Add(tlHidden,	tcItem)

			RETURN .T.
		ENDIF

		* Key already in use...

		* ...on the same item?
		IF This.oAccessKeyList.Item(tcKey) == tcItem
			* Setting the same key on the same item has no effect other than maybe updating the hidden flag to false...
			IF This.oAccessHiddenList.Item(tcItem) AND !tlHidden
				This.oAccessHiddenList.Remove(tcItem)
				This.oAccessHiddenList.Add(tlHidden, tcItem)
			ENDIF

			RETURN .T.
		ENDIF

		* Deal with the clash...
		IF !tlHidden
			IF This.oAccessHiddenList.Item(This.oAccessKeyList.Item(tcKey))
				* The item using it is hidden but we are not - override...
				This.oAccessKeyList.Remove(tcKey)

				* If this item already had a key, remove it first:
				IF This.oAccessItemList.GetKey(tcItem) != 0
					This.oAccessItemList.Remove(tcItem)
					This.oAccessHiddenList.Remove(tcItem)
				ENDIF

				* Store the linkages:
				This.oAccessKeyList.Add(tcItem,			tcKey)
				This.oAccessItemList.Add(tcKey,			tcItem)
				This.oAccessHiddenList.Add(tlHidden,	tcItem)

				RETURN .T.
			ELSE
				IF tlSuppressError
					RETURN .F.
				ENDIF

				LOCAL lcText

				lcText = "Can't set AccessKey for "

				IF LEN(tcItem) > ACCESSKEY_ERROR_TRIMLEN
					lcText = lcText + LEFT(tcItem, ACCESSKEY_ERROR_TRIMLEN) + "..."
				ELSE
					lcText = lcText + tcItem
				ENDIF

				lcText = lcText + " to '" + tcKey + "'; already in use by "

				tcItem = This.oAccessKeyList.Item(tcKey)
				IF LEN(tcItem) > ACCESSKEY_ERROR_TRIMLEN
					lcText = lcText + LEFT(tcItem, ACCESSKEY_ERROR_TRIMLEN) + "..."
				ELSE
					lcText = lcText + tcItem
				ENDIF

				This.AddError(lcText)
				RETURN .F.
			ENDIF
		ELSE
			* Can have doubleups for items that are hidden (e.g. in modal content), so just add a link for the item - leave the key link as it was...

			* If this item already had a key, remove it first:
			IF This.oAccessItemList.GetKey(tcItem) != 0
				This.oAccessItemList.Remove(tcItem)
				This.oAccessHiddenList.Remove(tcItem)
			ENDIF

			* Store the item linkages:
			This.oAccessItemList.Add(tcKey, tcItem)
			This.oAccessHiddenList.Add(tlHidden, tcItem)

			RETURN .T.
		ENDIF
	ENDFUNC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	* Allocates the next available accessKey from the letters/numbers in the title and formats the title with it underlined and returns the key,
	* or returns the empty string if all the letters/numbers in this title are used or if accessKeys are disabled.
	* - tcItem and tlHidden are as per SetAccessKey()
	* - rcText should be passed by reference if you want it edited to have the allocated key underlined (optional).
	FUNCTION SetNextAccessKey(tcItem, rcText, tlHidden) AS String
		LOCAL lnI, lcChar, llInEntity, llInTag

		IF !ACCESSKEYS_ENABLED
			RETURN ""
		ENDIF

		llInEntity = .F.
		llInTag = .F.

		rcText = STRTRAN(STRTRAN(rcText, "<u>", ""), "</u>", "")

		FOR lnI = 1 TO LEN(rcText)
			lcChar = SUBSTR(rcText, lnI, 1)

			IF llInEntity
				IF lcChar == ';'
					llInEntity = .F.
				ENDIF

				LOOP
			ENDIF
			IF llInTag
				IF lcChar == '>'
					llInTag = .F.
				ENDIF

				LOOP
			ENDIF

			IF lcChar == '&'
				llInEntity = .T.
				LOOP
			ENDIF
			IF lcChar == '<'
				llInTag = .T.
				LOOP
			ENDIF

			IF !(UPPER(lcChar) >= 'A' AND UPPER(lcChar) <= 'Z' OR lcChar >= '0' AND lcChar <= '9')
				* Only work with letters or numbers
				LOOP
			ENDIF

			* Check the key:
			IF This.oAccessKeyList.GetKey(lcChar) != 0
				* Used already...
				IF tlHidden OR !This.oAccessHiddenList.Item(This.oAccessKeyList.Item(lcChar))
					LOOP
				*ELSE
					* Allow override if we are not hidden and the item using this key is...
				ENDIF
			ENDIF

			IF This.SetAccessKey(tcItem, UPPER(lcChar), tlHidden, .T.)
				rcText = STUFF(rcText, lnI, 1, "<u>" + lcChar + "</u>")
				RETURN UPPER(lcChar)
			ENDIF
		NEXT

		RETURN ""
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	* 12/01/2011  CMGM  TTP6452  New function that will sanitise strings. It will
	*                            remove whitespaces first then remove the ff characters:
	*  &  - ampersand
	*  <  - left angle bracket
	*  >  - right angle bracket
	*  /  - forward slash
	*  '  - single quotation mark
	*  "  - double quotation mark
	*  \  - backslash
	*  ;  - semicolon
	* ..  - ellipses (2 dots only, for directory traversing)
	*
	* 31/03/2011  CMGM  TTP6452  Added tnStringType as parameter
	*                            1 = URL, saves all "tags" and "code within tags" into
	*                                     an array then removes them from the string.
	*                                     afterwards, removes all invalid remnant chars
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	FUNCTION SanitiseString(tcText, tnStringType) AS String
		IF EMPTY(tcText)
			RETURN tcText
		ENDIF
		
		IF EMPTY(tnStringType)
			* Init otherwise CASE clause will fail
			tnStringType = 0
		ENDIF

		tcText = ALLTRIM(tcText)

		DO CASE
		CASE tnStringType = 1
			* URL
			LOCAL lnI, w, x, y, z, lcTemp
			DIMENSION laInvalidCode[1]
			y = 1	&& Same start as x, do not change -- will mess up STREXTRACT in FOR clause below
			z = 1	&& Same start as x

			* Get the code within "<" and ">" include the delimiters
			lnI = LEN(tcText)
			FOR x=1 TO lnI
				IF z < lnI
					lcTemp = STREXTRACT(tcText, [<], [>], y, 4) 
					IF EMPTY(lcTemp)
						EXIT
					ELSE
						DIMENSION laInvalidCode[y]
						laInvalidCode[y] = lcTemp
						y = y + 1
						z = z + LEN(lcTemp)
					ENDIF
					* Increment x so don't have to loop so many times
					IF z > x
						x = z
					ENDIF
				ENDIF
			ENDFOR

			* Init the vars appropriately because we need to go thru the whole string again
			w = ALEN(laInvalidCode) + 1
			y = 1	&& do not change to 0 -- will mess up STREXTRACT in FOR clause below
			z = 1	&& Same start as x
			dummy = 1
			
			* Get the code within e.g. <script>...code here...</script>
			FOR x=1 TO lnI
				IF z < lnI
					lcTemp = STREXTRACT(tcText, [>], [<], y)
					IF EMPTY(lcTemp)
						IF dummy = 1
							* Ignore the first blank
							y = y + 1
							dummy = 2
						ELSE
							EXIT
						ENDIF
					ELSE
						DIMENSION laInvalidCode[w]
						laInvalidCode[w] = lcTemp
						w = w + 1
						y = y + 1
						z = z + LEN(lcTemp)
					ENDIF
					* Increment x so don't have to loop so many times
					IF z > x
						x = z
					ENDIF
				ENDIF
			ENDFOR

			* Now remove all invalid code in string
			IF NOT EMPTY(laInvalidCode)
				FOR x=1 TO ALEN(laInvalidCode)
					tcText = STRTRAN(tcText, laInvalidCode[x], "")
				ENDFOR
			ENDIF

			* Remove other invalid / remnant characters
			tcText = CHRTRAN(tcText, ['"<>],"")

		OTHERWISE
			* Common string
			tcText = CHRTRAN(tcText, [&<>/'"\;] ,"")
			tcText = STRTRAN(tcText, "..", "")
		ENDCASE
		
		RETURN tcText
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION SQLSanitise(tcText) AS String
		RETURN STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(tcText, ["], ""), "[", ""), "]", ""), "'", ""), CHR(9), ""), CHR(10), ""), CHR(13), "")
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	* Adjust current date/time value by the delay specified in the company setup.
	FUNCTION GetLocalTime(tdDateTime)
		* The server time is set to NZ standard time.  Need to adjust only for AU for now.
		LOCAL lnDelay, ltTime

		lnDelay = VAL(AppSettings.Get("timedelay"))

		IF lnDelay != 0
			IF VARTYPE(tdDateTime) == 'T'
				ltTime = tdDateTime + 3600 * lnDelay
			ELSE
				ltTime = tdDateTime + lnDelay / 24
			ENDIF
		ELSE
			ltTime = tdDateTime
		ENDIF

		RETURN ltTime
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION StripHTML(tcHTML AS String, tlLeaveInLinebreaks AS Boolean) As String
		LOCAL lcOutput, llInTag, lcChar, lnI, lnStart, lcTag

		llInTag = .F.
		lcOutput = ""

		FOR lnI = 1 TO LEN(tcHTML)
			lcChar = SUBSTR(tcHTML, lnI, 1)
			IF llInTag
				IF lcChar == '>'
					llInTag = .F.
					lcTag = SUBSTR(tcHTML, lnStart, lnI - lnStart + 1)
					* keep br tags but strip any styling etc.
					IF LOWER(LEFT(lcTag, 3)) == "<br" AND INLIST(SUBSTR(lcTag, 4, 1), ' ', '/') AND RIGHT(lcTag, 2) == "/>"
						IF tlLeaveInLinebreaks
							lcOutput = lcOutput + "<br/>"
						ELSE
							lcOutput = lcOutput + ' '	&& to stop things running together too much.
						ENDIF
					ENDIF
				ENDIF
			ELSE
				IF lcChar == '<'
					llInTag = .T.
					lnStart = lnI
				ELSE
					lcOutput = lcOutput + lcChar
				ENDIF
			ENDIF
		NEXT

		RETURN lcOutput
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	&&NOTE: does not close all open tags not already closed in snip;  for best results, strip them all out first.
	FUNCTION WordTrim(tcText, tnLength, tcAdd) As String
		LOCAL lcResult, lnLength

		* Early bailout check
		IF LEN(tcText) <= tnLength
			RETURN tcText
		ENDIF

		* Trim to the correct length
		lnLength = tnLength - LEN(tcAdd)
		lcResult = SUBSTR(tcText, 1, lnLength)

		*!* 24/11/2009;TTP4863;JCF: Fix an infinite loop bug when a shown news item has a very long first word.
		* If the first word is longer than the length, we can't split before it so chop it instead... (We need this as otherwise the following loop never exits!!!)
		IF OCCURS(' ', lcResult) + OCCURS(CHR(9), lcResult) + OCCURS(CHR(10), lcResult) == 0	&& ...contains no occurances of the 3 boundary chars for the below loop...
			RETURN lcResult + tcAdd
		ENDIF

		* Find the last break between words within the required length
		DO WHILE !INLIST(SUBSTR(lcResult, LEN(lcResult), 1), ' ', CHR(9), CHR(10))			&& not including CHR(13) so that CRLF's are not split.
			lcResult = SUBSTR(lcResult, 1, LEN(lcResult) - 1)
		ENDDO

		* Trim back to the end of the previous word
		DO WHILE INLIST(SUBSTR(lcResult, LEN(lcResult), 1), ' ', CHR(9), CHR(10), CHR(13))	&& including CHR(13) so that we eat all the whitespace.
			lcResult = SUBSTR(lcResult, 1, LEN(lcResult) - 1)
		ENDDO

		RETURN lcResult + tcAdd
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Breaks too-long words with spaces
	FUNCTION EnforceMaxWordLength(tcText, tnLength) AS String
		LOCAL lnI, lnCount, lnEnd, lcChar

		lnI = 1
		lnEnd = LEN(tcText)
		lnCount = 0

		DO WHILE lnI <= lnEnd
			lcChar = SUBSTR(tcText, lnI, 1)
			IF lcChar == CHR(9) OR lcChar == CHR(10) OR lcChar == CHR(13) OR lcChar == ' '
				* Reset for the next word
				lnCount = 0
			ELSE
				lnCount = lnCount + 1

				IF lnCount > tnLength
					* Split the word by inserting a space - don't care if it's next to another one (e.g. if this is the end of the word) as duplicate spaces collapse in HTML
					tcText = STUFF(tcText, lnI, 0, ' ')

					* Ensure our length is updated so we end at the right place.	&& 25/11/2009;TTP4868;JCF: bugfix for an issue found when we used this code in the payroll payslip note screen - do not increment lnI here - it makes the length-check be out by one for all but the first long word when splitting large blocks of unbroken text.
					lnEnd = lnEnd + 1

					* Reset for the next word
					lnCount = 0
				ENDIF
			ENDIF

			lnI = lnI + 1
		ENDDO

		RETURN tcText
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION URLEscape(tcText)
		LOCAL lnIndex, lcOutput, lcChar, lnCharVal

		lcOutput = ""
		FOR lnIndex = 1 TO LEN(tcText)
			lcChar = SUBSTR(tcText, lnIndex, 1)
			lnCharVal = ASC(lcChar)
			DO CASE
			CASE lnCharVal == ASC(' ')													&& ' ' -> '+'
				lcOutput = lcOutput + '+'
			CASE lnCharVal < ASC('0') OR BETWEEN(lnCharVal, ASC('9') + 1, ASC('A') - 1);
					OR BETWEEN(lnCharVal, ASC('Z') + 1, ASC('a') - 1) OR lnCharVal > ASC('z')
				lcOutput = lcOutput + '%' + SUBSTR(TRANSFORM(lnCharVal, "@0"), 9, 2)	&& outOfBandChar -> '%XX'
			OTHERWISE																	&& inBandChar; no change
				lcOutput = lcOutput + lcChar
			ENDCASE
		ENDFOR

		RETURN lcOutput
	ENDFUNC

	*--------------------------------------------------------------------------------*

	FUNCTION URLUnEscape(tcText)
		LOCAL lnIndex, lcOutput, lnCharVal, lcTemp

		lcTemp = STRTRAN(tcText, '+', ' ')						&& '+' -> ' '
		lnIndex = AT('%', lcTemp)								&& is there any "%XX" in there?
		IF EMPTY(lnIndex)
			lcOutput = lcTemp										&& no; keep all the rest, stop
		ELSE
			lcOutput = SUBSTR(lcTemp, 1, lnIndex - 1)				&& yes; keep upto it
			DO WHILE !EMPTY(lnIndex)
				lcTemp = SUBSTR(lcTemp, lnIndex + 1)					&& remove kept and '%' from temp
				lnCharVal = VAL("0x" + SUBSTR(lcTemp, 1, 2))			&& encoded char is now the 1st 2 digits
				lcOutput = lcOutput + CHR(lnCharVal)
				lcTemp = SUBSTR(lcTemp, 3)								&& remove the 2 digits
				lnIndex = AT('%', lcTemp)								&& is there any "%XX" left in there?
				IF EMPTY(lnIndex)
					lcOutput = lcOutput + lcTemp							&& no; add all the rest, stop
				ELSE
					lcOutput = lcOutput + SUBSTR(lcTemp, 1, lnIndex - 1)	&& yes; keep upto it
				ENDIF
			ENDDO
		ENDIF

		RETURN lcOutput
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Make the text safe for use between ""'s, e.g. in <input type="text" value="(here)"> etc
	FUNCTION InputFieldEscape(tcText) AS String
		* (Replace all '"' with '&#34;')
		RETURN STRTRAN(tcText, '"', "&#34;")
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Make the text safe for display on an HTML page, or in a <textarea>
	FUNCTION HTMLEscape(tcText, tlEncodeLinebreaks) AS String
		IF tlEncodeLinebreaks
			* used when we are outputting (assumed valid) HTML or text that contains linebreaks...	(It's no wrong to double-escape, one with .T. and one with .F.)
			tcText = STRTRAN(tcText, CRLF, "<br/>")
		ELSE
			* (Replace all {'&', '<', '>', '"', "'"} with {"&amp;", "&lt;", "&gt;", "&quot", "&#39;"} - in that order, or else!)
			* Include forward slash as it helps end an HTML entity
			tcText = STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(TRANSFORM(tcText), '&', "&amp;"), '<', "&lt;"), '>', "&gt;"), '"', "&quot;"), ['], "&#39;")
			tcText = STRTRAN(tcText,[/],"&#47;")
		ENDIF
		RETURN tcText
	ENDFUNC

	*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *

	* Escape everything but &nbsp; entities
	FUNCTION HTMLEscape2(tcText) AS String
		RETURN STRTRAN(This.HTMLEscape(tcText), "&amp;nbsp;", "&nbsp;")
	ENDFUNC


    FUNCTION HTMLGetSafeAntiXSS(tcText) AS String
    	DO wwdotnetbridge
		LOCAL lobridge AS wwdotnetbridge
		LOCAL retString AS String
		TRY
			lobridge = CREATEOBJECT("wwDotNetBridge","V4")
			=lobridge.loadassembly("EXOESdotnet.dll")
			esnet=lobridge.createinstance("EXOESdotnet.GenUtils")
			retString=  esnet.GetSafeHtml(tcText)
		CATCH
		    retString=tcText
		ENDTRY
		esnet=null
		lobridge=null
		RETURN retString
    
    ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE htmlsanitiser
		PARAMETERS lcString
		IF VARTYPE(lcString)<>"C"
			RETURN ""
		ENDIF
		IF LEN(lcString)>5000
			RETURN "<br>This html message exceeded the maximum allowed length of 5000 characters</br>"
		ENDIF
		LOCAL lcOutput,lnI
		TRY
			lcOutput = This.HTMLGetSafeAntiXSS(lcString)
		CATCH
			lcOutput = ""
			FOR lnI = 1 TO LEN(lcString)
				lcChar = SUBSTR(lcString, lnI, 1)
				DO CASE
				CASE lcChar = '&'
					lcOutput = lcOutput + '&amp;'
				CASE lcChar = '<'
					lcOutput = lcOutput + '&lt;'
				CASE lcChar = '>'
					lcOutput = lcOutput + '&gt;'
				CASE lcChar = '"'
					lcOutput = lcOutput + '&quot;'
				CASE lcChar = "'"
					lcOutput = lcOutput + '&#x27;'
				CASE lcChar = '/'
					lcOutput = lcOutput + '&#x2F;'
				OTHERWISE
					lcOutput = lcOutput + lcChar
				ENDCASE
			ENDFOR
		ENDTRY
		RETURN lcOutput
	ENDFUNC
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	* Safely read a value from a Key/Value Collection, returning the supplied default if it is not present.
	FUNCTION GetConfigValue(toCollection as Collection, tcKey as String, tuDefault as Variant) As Variant
		LOCAL lnIndex
		
		lnIndex = toCollection.GetKey(tcKey)
		IF lnIndex > 0
			RETURN toCollection.Item(lnIndex)
		ELSE
			RETURN tuDefault
		ENDIF
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Safely remove a value from a Key/Value Collection.
	PROCEDURE RemoveConfigValue(toCollection as Collection, tcKey as String)
		LOCAL lnIndex

		lnIndex = toCollection.GetKey(tcKey)
		IF lnIndex > 0
			toCollection.Remove(tcKey)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	* Safely write a value to a Key/Value Collection, removing it first if it was already present.
	PROCEDURE SetConfigValue(toCollection as Collection, tcKey as String, tuValue as Variant)
		LOCAL lnIndex

		This.RemoveConfigValue(toCollection, tcKey)
		toCollection.Add(tuValue, tcKey)
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	* MIME types...
	FUNCTION GetDocType(tcFileName)
		LOCAL lcExt, lcContent
		lcExt = UPPER(JUSTEXT(tcFileName))
		* Always CONTENT TYPE/SUB TYPE
		* Valid content types are application, audio, image, message, model, multipart, text, video
		* Subtypes are variable
		DO CASE
		CASE lcExt = "PDF"
			lcContent = "application/pdf"
		CASE lcExt = "DOC"
			lcContent = "application/msword"
		CASE lcExt = "RTF"
			lcContent = "application/rtf"
		CASE lcExt = "XLS"
			lcContent = "application/vnd.ms-excel"
		CASE lcExt = "XML"
			lcContent = "application/xml"
		CASE lcExt = "ZIP"
			lcContent = "application/zip"
		CASE lcExt = "PNG"
			lcContent = "image/png"
		CASE lcExt = "JPEG"
			lcContent = "image/jpeg"
		CASE lcExt = "GIF"
			lcContent = "image/gif"
		CASE lcExt = "HTML" OR lcExt = "HTM"
			lcContent = "text/html"
		OTHERWISE
			*lcContent = "application/octet-stream"
			lcContent = "text/plain"
		ENDCASE
		RETURN lcContent
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	FUNCTION ValidateBankAccount(tnBank, tnBranch, tnAccount, tnSuffix)
		LOCAL BS, I, NUMS, SUM, J, BSOK

		BS = PADL(ALLTRIM(STR(tnBranch)), 4, "0");
			+ PADL(ALLTRIM(STR(tnAccount)), 7, "0");
			+ PADL(ALLTRIM(STR(tnSuffix)), 3, "0")

		DO CASE
		CASE !INLIST(tnBank, 4, 5, 7, 8, 9, 10, 25, 26, 28, 29) AND tnBank <= 38
			* Modulus 11 checks.
			IF tnAccount < 990000
				BSOK = (MOD(;
					  VAL(SUBSTR(BS,  1, 1)) * 6;
					+ VAL(SUBSTR(BS,  2, 1)) * 3;
					+ VAL(SUBSTR(BS,  3, 1)) * 7;
					+ VAL(SUBSTR(BS,  4, 1)) * 9;
					+ VAL(SUBSTR(BS,  6, 1)) * 10;
					+ VAL(SUBSTR(BS,  7, 1)) * 5;
					+ VAL(SUBSTR(BS,  8, 1)) * 8;
					+ VAL(SUBSTR(BS,  9, 1)) * 4;
					+ VAL(SUBSTR(BS, 10, 1)) * 2;
					+ VAL(SUBSTR(BS, 11, 1)) * 1,;
					11;
				) == 0)
			ELSE
				BSOK = (MOD(;
					  VAL(SUBSTR(BS,  6, 1)) * 10;
					+ VAL(SUBSTR(BS,  7, 1)) * 5;
					+ VAL(SUBSTR(BS,  8, 1)) * 8;
					+ VAL(SUBSTR(BS,  9, 1)) * 4;
					+ VAL(SUBSTR(BS, 10, 1)) * 2;
					+ VAL(SUBSTR(BS, 11, 1)) * 1,;
					11;
				) == 0)
			ENDIF

		CASE tnBank == 9
			DIMENSION NUMS[5]

			NUMS[1] = VAL(SUBSTR(BS,  8, 1)) * 5
			NUMS[2] = VAL(SUBSTR(BS,  9, 1)) * 4
			NUMS[3] = VAL(SUBSTR(BS, 10, 1)) * 3
			NUMS[4] = VAL(SUBSTR(BS, 11, 1)) * 2
			NUMS[5] = VAL(SUBSTR(BS, 14, 1)) * 1

			FOR J = 1 TO 2
				FOR I = 1 TO 5
					IF LEN(ALLTRIM(STR(NUMS[I]))) == 2
						NUMS[I] = MOD(NUMS[I], 10) + FLOOR(NUMS[I] / 10)
					ENDIF
				ENDFOR
			ENDFOR
			BSOK = (MOD(NUMS[1] + NUMS[2] + NUMS[3] + NUMS[4] + NUMS[5], 11) == 0)

		CASE tnBank == 8 OR tnBank == 32
			BSOK = (MOD(;
				  VAL(SUBSTR(BS,  5, 1)) * 7;
				+ VAL(SUBSTR(BS,  6, 1)) * 6;
				+ VAL(SUBSTR(BS,  7, 1)) * 5;
				+ VAL(SUBSTR(BS,  8, 1)) * 4;
				+ VAL(SUBSTR(BS,  9, 1)) * 3;
				+ VAL(SUBSTR(BS, 10, 1)) * 2;
				+ VAL(SUBSTR(BS, 11, 1)) * 1,;
				11;
			) == 0)

		CASE tnBank == 25 OR tnBank == 33
			BSOK = (MOD(;
				  VAL(SUBSTR(BS,  5, 1)) * 1;
				+ VAL(SUBSTR(BS,  6, 1)) * 7;
				+ VAL(SUBSTR(BS,  7, 1)) * 3;
				+ VAL(SUBSTR(BS,  8, 1)) * 1;
				+ VAL(SUBSTR(BS,  9, 1)) * 7;
				+ VAL(SUBSTR(BS, 10, 1)) * 3;
				+ VAL(SUBSTR(BS, 11, 1)) * 1,;
				10;
			) == 0)

		CASE tnBank = 26 OR tnBank = 29
			DIMENSION NUMS[10]

			NUMS[1]  = VAL(SUBSTR(BS,  5, 1)) * 1
			NUMS[2]  = VAL(SUBSTR(BS,  6, 1)) * 3
			NUMS[3]  = VAL(SUBSTR(BS,  7, 1)) * 7
			NUMS[4]  = VAL(SUBSTR(BS,  8, 1)) * 1
			NUMS[5]  = VAL(SUBSTR(BS,  9, 1)) * 3
			NUMS[6]  = VAL(SUBSTR(BS, 10, 1)) * 7
			NUMS[7]  = VAL(SUBSTR(BS, 11, 1)) * 1
			NUMS[8]  = VAL(SUBSTR(BS, 12, 1)) * 3
			NUMS[9]  = VAL(SUBSTR(BS, 13, 1)) * 7
			NUMS[10] = VAL(SUBSTR(BS, 14, 1)) * 1

			SUM = 0
			FOR J = 1 TO 2
				FOR I = 1 TO 10
					IF LEN(ALLTRIM(STR(NUMS[I]))) = 2
						NUMS[I] = MOD(NUMS[I], 10) + FLOOR(NUMS[I] / 10)
					ENDIF
					IF J = 2
						SUM = SUM + NUMS[I]
					ENDIF
				ENDFOR
			ENDFOR
			BSOK = (MOD(SUM, 10) == 0)

		OTHERWISE
			BSOK = .F.
		ENDCASE

		BSOK = BSOK AND tnAccount > 0 AND (tnBranch > 0 OR tnBank == 9)

		RETURN BSOK
	ENDFUNC

	*################################################################################*
#DEFINE TOC_ServeUtils_

	*> +define: ServeUtils; fix
	* Send the user a file (if logged in) - a.k.a. pretend to be a secure webserver
	FUNCTION ServeDoc(tcFileName, tcContentType)
		IF FILE(tcFileName)

			Response.TransmitFile(tcFileName, tcContentType)

			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> rename: ServePDFReport; test!; fix
	PROCEDURE ServeReport()
		LOCAL lcReportName, lcReportFile

		lcReportName = STRTRAN(Request.QueryString("report"), '\', "")	&& no subdirs supported!

		&&FIXME: Security Issue: the following should probably be length-clamped, but to what?
		lcReportFile = This.CompanyDataPath() + ALLTRIM(lcReportName) + ".pdf"

		Response.TransmitFile(lcReportFile, "application/pdf")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> rename: ServeExcelReport; test!; fix
	PROCEDURE ServeFile()
		LOCAL lcReportName, lcReportFile

		lcReportName = STRTRAN(Request.QueryString("report"), '\', "")	&& no subdirs supported!

		&&FIXME: Security Issue: the following should probably be length-clamped, but to what?  And restricted to a fixed extension, since the MIME type is forced...
		lcReportFile = This.CompanyDataPath() + ALLTRIM(lcReportName)

		Response.TransmitFile(lcReportFile, "application/vnd.ms-excel")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> fix
	* Serve a file, or go back to the PolicyDocumentsPage to display an error message
	PROCEDURE GetDocument()
		LOCAL loDoc, lnDocId, lcContent

		&&FIXME: Security Issue: the following should probably be length-clamped, but to what?  And maybe restricted to a fixed set of extensions?
		lcDoc = This.CompanyDocumentsPath() + This.SanitiseString(Request.QueryString("doc"))

		IF !This.ServeDoc(lcDoc, This.GetDocType(lcDoc))
			This.AddError("Could not return document.")
			This.PolicyDocumentsPage()
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> fix
	* Serve a payslip PDF file, or go back to the ReportsPage to display an error message
	PROCEDURE GetReport()
		* Input expected is the date-part of the filename - no stealing another emplyoees PDF's here, mister!
		LOCAL lcDateParameter, lcFileName

		&&FIXME: Security Issue: the following should probably be length-clamped, but to what?
		lcDateParameter = This.SanitiseString(Request.QueryString("date"))

		IF !EMPTY(lcDateParameter)
			lcFileName = This.CompanyDataPath() + ADDBS("payslips") + TRANSFORM(This.Employee) + '_' + lcDateParameter + ".pdf"
			IF This.ServeDoc(lcFileName, "application/pdf")
				* don't want the ReportsPage on the end of this, so:
				RETURN
			ELSE
				This.AddError("No report available for that date: " + lcDateParameter)
			ENDIF
		ELSE
			This.AddError("No date specified.")
		ENDIF

		This.ReportsPage()
	ENDPROC

	*--------------------------------------------------------------------------------*
	PROCEDURE GetReportLink(tcReportEXT)
		LOCAL lcLinkName, lcFileName, lcDownFile as String
		lcLinkName = Request.QueryString("date")
		
		IF !EMPTY(lcLinkName)
			* save as filename
			lcDownFile = lcLinkName + "_" + DTOS(DATE())+STRTRAN(TIME(),":","") + "." + tcReportEXT 
			* source file link
			lcFileName = This.companydatapath() + "reports\" + TRANSFORM(This.Employee) + "_" + lcLinkName + "." + tcReportEXT
			IF FILE(lcFileName)
 				Response.DownLoadFile(lcFileName, "application/" + tcReportEXT, lcDownFile)
			ENDIF
		ENDIF

	ENDPROC
	*--------------------------------------------------------------------------------*
	PROCEDURE GetPDF()
		This.GetReportLink("pdf")
	ENDPROC
	*--------------------------------------------------------------------------------*
	PROCEDURE GetCSV()
		This.GetReportLink("csv")
	ENDPROC
*!*		*--------------------------------------------------------------------------------*
*!*		PROCEDURE GetPDF()
*!*			LOCAL lcFileName, lcDownFile

*!*			lcFileName = Request.QueryString("PDFFile")
*!*	 		lcDownFile = Request.QueryString("PDFOut")
*!*	 
*!*			IF !EMPTY(lcFileName)
*!*				IF FILE(lcFileName)
*!*	 				Response.DownLoadFile(lcFileName, "application/pdf",lcDownFile)
*!*				ENDIF
*!*			ENDIF

*!*		ENDPROC
	*--------------------------------------------------------------------------------*

	*> fix
	* Serve a payslip PDF file, or go back to the ReportsPage to display an error message
*!*		PROCEDURE GetCSV()
*!*			LOCAL lcFileName, lcDownFile

*!*			lcFileName = Request.QueryString("CSVFile")
*!*	 		lcDownFile = Request.QueryString("CSVOut")
*!*	 
*!*			IF !EMPTY(lcFileName)
*!*				IF FILE(lcFileName)
*!*	 				Response.DownLoadFile(lcFileName, "application/csv",lcDownFile)
*!*				ENDIF
*!*			ENDIF

*!*		ENDPROC

	*################################################################################*
#DEFINE TOC_Database_

	*> +define: Database
	FUNCTION SelectData(tnLicence, tcDataSource)
		IF EMPTY(tcDataSource) OR !DIRECTORY(This.CompanyDataPath())
			RETURN .F.
		ENDIF

		LOCAL lcAlias, llOK, lcDataBase, lcPath, loException, lcTemp

		llOK = .F.

		* Parse any dbc names from the data source.
		lcDatabase = ""
		IF '!' $ tcDataSource
			lcTemp = CHRTRAN(tcDataSource, '!', '.')
			lcDataBase = ALLTRIM(JUSTSTEM(lcTemp))
			lcAlias = JUSTEXT(lcTemp)
		ELSE
			lcAlias = tcDataSource
		ENDIF

		* Get old path
		lcPath = SET("PATH")

		TRY
			* Set path to point at this companys data
			SET PATH TO (ADDBS(This.cDataPath) + TRANSFORM(tnLicence))
			* Open stuff...
			IF !EMPTY(lcDataBase)
				OPEN DATABASE (lcDatabase) SHARED
			ENDIF
			IF !USED(lcAlias)
				* Note: the NODATA - must remember to requery a view ourselves when we call This..!
				USE (tcDataSource) AGAIN NODATA IN 0
			ENDIF
		CATCH TO loException
			This.AddError("Failed to open " + tcDataSource  + ";<br>" + TRANSFORM(loException.ErrorNo) + ": " + loException.Message)
		FINALLY
			* Put the path back
			SET PATH TO (lcPath)
		ENDTRY

		IF USED(lcAlias)
			** Note: a lot of the code relies on the fact that SelectData() includes a SELECT here..:
			SELECT (lcAlias)
			llOK = .T.
		ENDIF

		RETURN llOK
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*

	*!* 19/11/2009;TTP4845;JCF: Added tlDefaultToGroup.
	*@ SetupStaffGroupControlData:
	 * Parameters:
	 *	rlManager:			Passed by reference; Set to true if the current user is a manager.
	 *	rnCurrentGroup:		Passed by reference; Set to the currently selected group, or to the default group (see tlDefaultToGroup).
	 *	rnCurrentStaff:		Passed by reference; Set to the currently selected employee, or to the current employee if possible (e.g. if a member of the current group), otherwise, to the first employee in the group.
	 *	toRetainList:		The RetainList object to save state into.
	 *	tlDefaultToGroup:	Optional; If true, default to the All Employees group if present, otherwise the first group, rather than the My Details "group".  Defaults to false.
	 * Returns:
	 *	True on success, false on error.  Note that if the current user is not a manager then the the control will be skipped so we return false.
	 * Notes:
	 *	Doesn't work if the control uses non-default formField names!  It expects them to be called currentStaff and currentGroup only.
	FUNCTION SetupStaffGroupControlData(rlManager AS Boolean, rnCurrentGroup AS Integer, rnCurrentStaff AS Integer, toRetainList, tlDefaultToGroup AS Boolean) AS Boolean
		rlManager = This.IsManager(This.Employee)

		IF !rlManager
			* Non-managers can only see their own record, but that is not an error.
			&&NOTE: doing it this way means if they try to hack the page nothing happens, rather than being warned off... (Otherwise this needs to be after the QueryString reads and we would need to check the input against the allowed...)
			rnCurrentGroup = MY_DETAILS_GROUP
			rnCurrentStaff = This.Employee
			RETURN .T.
		ENDIF

		rnCurrentStaff = EVL(;
			VAL(Request.Form("currentStaff")),;
			EVL(;
				VAL(Request.QueryString("currentStaff")),;
				This.Employee;
			);
		)
		*!* 19/11/2009;TTP4845;JCF: Added the ability to ask for the default group to not be My Details, and if that is the case, to prefer the All Employees group.
		rnCurrentGroup = EVL(;
			VAL(Request.Form("currentGroup")),;
			EVL(;
				VAL(Request.QueryString("currentGroup")),;
				IIF(;
					tlDefaultToGroup,;
					This.GetDefaultGroup(This.Employee),;
					MY_DETAILS_GROUP;
				);
			);
		)

		*!* 19/11/2009;TTP4845;JCF: removed redundent code.

		IF rnCurrentStaff != EVERYONE_OPTION	&& if not Everyone...
			* Current staff is based on the selected group only
			* So get the first employee for the selected group:
			IF This.GetEmployeesByGroupCode(rnCurrentGroup, "curStaff")
				GO TOP IN curStaff
				* Does this employee exist in the current group?
				SELECT curStaff
				LOCATE FOR myWebCode == rnCurrentStaff
				IF !FOUND()
					LOCATE
					rnCurrentStaff = curStaff.myWebCode
				ENDIF

				USE IN SELECT("curStaff")
				USE IN SELECT("curGroups")

				IF EMPTY(rnCurrentStaff)
					rnCurrentStaff = This.Employee
				ENDIF
			ELSE
				rlManager = .F.		&& so that the control is skipped
			ENDIF
		ENDIF

		toRetainList.SetEntry("currentStaff", rnCurrentStaff)
		toRetainList.SetEntry("currentGroup", rnCurrentGroup)

		RETURN rlManager
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	* Sets up everything needed for the PayControl.	(Note: doesn't work if the control uses a non-default formField name!)
	FUNCTION SetupPayControlData(rnCurrentPay, rlPayOpen, toRetainList, tcPayStatus) AS Integer
		LOCAL lcFilter, lnCount

		IF !This.SelectData(This.Licence, "myPays")
			RETURN -1	&& no data!
		ELSE
			IF VARTYPE(tcPayStatus) != 'C' OR !INLIST(tcPayStatus, "all", "open", "closed")
				tcPayStatus = "open"
			ENDIF

			DO CASE
				CASE tcPayStatus == "all"
					lcFilter = ".T."
				CASE tcPayStatus == "open"
					lcFilter = "pay_status == 1"
				CASE tcPayStatus == "closed"
					lcFilter = "pay_status == 2"
			ENDCASE

			* Select pays only, not templates:
			SELECT *;
				FROM myPays;
				WHERE &lcFilter. AND pay_type == 2;
				INTO CURSOR curPays;
				ORDER BY pay_name

			lnCount = _TALLY

			rnCurrentPay = EVL(VAL(Request.Form("currentPay")), VAL(Request.QueryString("currentPay")))

			IF VARTYPE(rnCurrentPay) != 'N' OR EMPTY(rnCurrentPay)
				rnCurrentPay = -1
			ENDIF

			SELECT curPays
			LOCATE FOR pay_pk == rnCurrentPay

			IF !FOUND()
				* Set to the first in the list
				SELECT curPays
				LOCATE
				rnCurrentPay = curPays.pay_pk
			ENDIF

			rlPayOpen = (curPays.pay_status == 1)

			toRetainList.SetEntry("currentPay", rnCurrentPay)

			RETURN lnCount
		ENDIF
	ENDFUNC


	*################################################################################*
#DEFINE TOC_Logging_

	*> +define: Logging
	PROCEDURE CheckSiteRegistration()
		LOCAL lcReg

		lcReg = ""

		IF FILE(This.CompanyDataPath() + "core.mem")
			lcReg = lcReg + "coretrue"
		ELSE
			lcReg = lcReg + "corefalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "employeemessaging.mem")
			lcReg = lcReg + "employeemessagingtrue"
		ELSE
			lcReg = lcReg + "employeemessagingfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "leavemanagement.mem")
			lcReg = lcReg + "leavemanagementtrue"
		ELSE
			lcReg = lcReg + "leavemanagementfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "payslips.mem")
			lcReg = lcReg + "payslipstrue"
		ELSE
			lcReg = lcReg + "payslipsfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "reporting.mem")
			lcReg = lcReg + "reportingtrue"
		ELSE
			lcReg = lcReg + "reportingfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "timesheet.mem")
			lcReg = lcReg + "timesheettrue"
		ELSE
			lcReg = lcReg + "timesheetfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "documents.mem")
			lcReg = lcReg + "documentstrue"
		ELSE
			lcReg = lcReg + "documentsfalse"
		ENDIF

		IF FILE(This.CompanyDataPath() + "timeclock.mem")
			lcReg = lcReg + "timeclocktrue"
		ELSE
			lcReg = lcReg + "timeclockfalse"
		ENDIF

		* Removed...
		*IF FILE(This.CompanyDataPath() + "special.mem")
		*	lcReg = lcReg + "specialtrue"
		*ELSE
			lcReg = lcReg + "specialfalse"	&& might as well still send This.
		*ENDIF

		Response.Write(lcReg)
	ENDPROC

	*--------------------------------------------------------------------------------*

	* /// <summary>
	 */// userInfo is a SYS(0) command run at the client end. We log it here so we have as much
	 */// information stored about the user attempting to upload/download information as we can get.
	 */// logs the machine name on the network, as well as the currently logged in user.
	 */// </summary>
	PROCEDURE LogSysZero(userInfo as String)
		LOCAL logInfo as String

		TEXT TO logInfo TEXTMERGE NOSHOW PRETEXT 7
			Date: << TRANSFORM(This.GetLocalTime(DATETIME())) >> | Licence: << TRANSFORM(This.Licence) >> | User: << TRANSFORM(userInfo) >> | Upload Staff
			---
		ENDTEXT

		STRTOFILE(logInfo, "upload.log", .T.)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	PROCEDURE LogReport(tcFile, tcText)
		LOCAL lcLog

		lcLog = This.CompanyDataPath() + "reports.dbf"
		IF !FILE(lcLog)
			CREATE TABLE (lcLog) FREE (cFile C(200), cText C(200), tTime T, nWebCode I)
			USE IN SELECT(lcLog)
		ENDIF

		IF This.SelectData(This.Licence, "reports")
			INSERT INTO reports (cFile, cText, tTime, nWebCode) VALUES (tcFile, tcText, This.GetLocalTime(DATETIME()), This.Employee)
		ENDIF
	ENDPROC

	*################################################################################*
#DEFINE TOC_Events_

	*> +define: Events
	* Event for adding new records during synchronise process.
	FUNCTION AddRecord(tcTable as String, tuKey as Variant)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Event for handling changed records during synchronise process.
	FUNCTION ChangeRecord(tcTable as String, tuKey as Variant)
		DO CASE
		CASE LOWER(tcTable) == "mystaff"
			This.SendSiteUpdatedEmailTo(tuKey, SENDMAIL_STAFF)
		ENDCASE
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Event for deleting records during synchronise process.
	FUNCTION DeleteRecord(tcTable as String, tuKey as Variant)
	ENDFUNC

	*################################################################################*
#DEFINE TOC_EventHandlers_

	*> +define: EventHandlers; comment
	FUNCTION OnAddRecord(tcTable as String, tuKey as Variant)
		DO CASE
		CASE LOWER(tcTable) == "mystaff"
			This.SendSiteUpdatedEmailTo(tuKey, SENDMAIL_STAFF)
		ENDCASE
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	FUNCTION OnChangeRecord(tcTable as String, tuKey as Variant)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	FUNCTION OnDeleteRecord(tcTable as String, tuKey as Variant)
	ENDFUNC

	*################################################################################*
#DEFINE TOC_Upload_

	FUNCTION ExcludeField(tcTable, tcField)
		LOCAL lcField, lcTable, lcFields

		lcTable = LOWER(ALLTRIM(tcTable))
		lcField = LOWER(ALLTRIM(tcField))
		DO CASE
		CASE lcTable == "mystaff"
			lcFields = "|mypaycode|mypayroll|mywebcode|mycountry|mybirth|mytaxnum|mytaxcode|myfreq|myfreqdesc|mystart|myslipname|myaddress|"
			lcFields = lcFields + "mysuburb|mycity|myphone|myhpperc|myhpent|myhpdate|myspent|myspdate|mylslent|mylsldate|myshent|myshdate|myotent|"
			lcFields = lcFields + "myotdate|mymanager|mypostcode|myrdototal|myxml|myudfl1|myudfl1d|myudfl2|myudfl2d|myudfd1|myudfd1d|myudfd2|myudfd2d|"
			lcFields = lcFields + "myudfc1|myudfc1d|myudfc2|myudfc2d|myudfc3|myudfc3d|myudfn1|myudfn1d|myudfm1|myudfm1d|mymobile|mybank|myaddress2|"
		ENDCASE

		RETURN '|' + lcField + '|' $ lcFields	&& CF 10/10/2008: safer this way - no accidental substring matches
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	* Gets the XML uploaded content
	FUNCTION GetUploadPackage(tcPackage as String) as String
		LOCAL lcXml, lcStatus, lcUserInfo, llUseCompression, lcTime, lcZipFile, lcPath, lcFile, loIp, lcUploadFile

		lcXml = Request.Form("xml")
		lcStatus = ""
		lcUserInfo = Request.Form("syszero")
		This.LogSysZero(lcUserInfo)

		llUseCompression = LOWER(Request.Form("zip")) == "true"
		IF llUseCompression
			lcTime = TTOC(DATETIME(), 1)
			lcZipFile = This.CompanyDataPath() + tcPackage + lcTime + ".zip"
			lcPath = ADDBS(FULLPATH(This.CompanyDataPath() + lcTime))
			lcFile = lcPath + tcPackage + ".xml"
			lcUploadFile = Request.Form("uploadfile")

			STRTOFILE(lcUploadFile, lcZipFile, 0)

			loIp = NEWOBJECT("DynaZip", "ComaccOnlineSecurity.vcx")
			TRY
				IF FILE(lcZipFile)
					* unzip the file
					IF loIp.UnzipFiles(lcZipFile, lcPath) == 0
						lcXml = FILETOSTR(lcFile)
					ENDIF
				ENDIF
			CATCH
				&&BUG: silent catch-all's should be documented!
			FINALLY
				DELETE FILE (lcZipFile)
				DELETE FILE (lcFile)
				loIp = null
				TRY
					RD (lcPath)
				CATCH
					&&BUG: silent catch-all's should be documented!
				ENDTRY
			ENDTRY
		ENDIF

		RETURN lcXml
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	PROCEDURE UploadGeneric(lcTable, lcKey, lcPackage)
		LOCAL lcXML, loData, lcField, i, llChanged

		lcXML = This.GetUploadPackage(lcPackage)

		USE IN SELECT("curXML")
		USE IN SELECT("curSource")

		IF EMPTY(lcXML)
			Response.Write("Problem: Could not open the XML data.")
			RETURN
		ENDIF

		XMLTOCURSOR(lcXML, "curXML")
		IF USED("curXML")
			IF !This.SelectData(This.Licence, lcTable)
				Response.Write("Problem: Could not open the table: " + This.cError)
				This.cError = ""
				RETURN
			ENDIF

			* Remove deleted records from the website
			SELECT &lcKey. FROM (lcTable) WHERE &lcKey. NOT IN (SELECT &lcKey. FROM curXML) INTO CURSOR curDeleted
			SELECT curDeleted
			SCAN
				SELECT (lcTable)
				LOCATE FOR &lcKey. == curDeleted.&lcKey.
				IF FOUND(lcTable)
					* this record needs to be removed
					DELETE IN (lcTable)
					RAISEEVENT(Process, "DeleteRecord", lcTable, curDeleted.&lcKey.)
				ENDIF
			ENDSCAN
			USE IN SELECT("curDeleted")

			* Now add any new records and process the existing ones
			SELECT curXML
			SCAN
				SELECT curXML
				SCATTER NAME loData MEMO
				SELECT (lcTable)
				LOCATE FOR &lcKey. == curXML.&lcKey.
				IF !FOUND(lcTable)
					* This is a new record
					SELECT (lcTable)
					APPEND BLANK
					GATHER NAME loData MEMO
					IF lcTable == "allow" OR lcTable == "costcent" OR lcTable == "wagetype"
					  replace hide WITH .f.
					ENDIF
					RAISEEVENT(Process, "AddRecord", lcTable, loData.&lcKey.)
				ELSE
					* This is an existing record - compare fields
					llChanged = .F.
					SELECT (lcTable)
					FOR i = 1 TO FCOUNT(lcTable)
						lcField = FIELD(i)
						TRY
							* Some fields may be mismatched on purpose - ignore the errors
							IF &lcField. != loData.&lcField. AND UPPER(ALLTRIM(lcField)) <> "HIDE"
								* field has changed - ignore 'Hide Fields - Web Only'
								REPLACE (lcField) WITH loData.&lcField. IN (lcTable)
								IF !This.ExcludeField(lcTable, lcField)	&&QUESTION: should this be around the replace too?  If not, document it!
									llChanged = .T.
								ENDIF
							ENDIF
						CATCH
							&&BUG: catch-all's need to be documented!
						ENDTRY
					ENDFOR
					IF llChanged
						RAISEEVENT(Process, "ChangeRecord", lcTable, loData.&lcKey.)
					ENDIF
				ENDIF
			ENDSCAN
		ENDIF

		USE IN SELECT("curXML")

		Response.Write("OK")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadWageTypes()
		This.UploadGeneric("wagetype", "code", "uploadwagetypes")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadAllowances()
		This.UploadGeneric("allow", "code", "uploadallowances")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadCostCentres()
		This.UploadGeneric("costcent", "code", "uploadcostcentres")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadMyTeams()
		This.UploadGeneric("myteams", "tmcode", "uploadmyteams")
	ENDPROC


	*> comment; sort
	PROCEDURE UploadImported()
		LOCAL lcImpFileName,emailto,emailFor
		lcImpFileName=ADDBS(This.cDataPath) + TRANSFORM(This.Licence) + "\myimported.dbf"
		IF !FILE(lcImpFileName)
			CREATE TABLE (lcImpFileName) FREE (impID I AUTOINC, mlcode I, mldaycode I,mldate T, impType C(1))
			INDEX ON mlcode TAG mlcode 
			INDEX ON impID TAG impID	
		ENDIF
		This.UploadGeneric("myimported", "impid", "uploadimported")
		
		** Update Leaverequests table as imported
		lStat = This.SelectData(This.Licence, "leaverequests")
		lStat = This.SelectData(This.Licence, "myimported")
		lStat = This.SelectData(This.Licence, "mycancelreq")
		UPDATE leaverequests FROM myimported WHERE leaverequests.id=myimported.mlcode AND myimported.imptype="P" SET imported=.T.
		SELECT a.id,a.employee,a.manager,CAST(0 as integer) as cancelby ;
		  FROM leaverequests a, myimported b WHERE a.id=b.mlcode AND b.imptype="C" AND a.cancelled=.F. ;
		    INTO CURSOR mytempcanceldays READWRITE

		INDEX ON ID TAG ID
		
		UPDATE mytempcanceldays FROM mycancelreq WHERE mytempcanceldays.id=mycancelreq.mlcode SET cancelby=mycancelreq.cancelby
		
		UPDATE leaverequests FROM myimported WHERE leaverequests.id=myimported.mlcode AND myimported.imptype="C" SET cancelled=.T.
		
		SELECT id,employee,manager,cancelby FROM mytempcanceldays GROUP BY id INTO CURSOR mytempcancel

        SELECT mytempcancel
        SCAN
            DO CASE
            CASE mytempcancel.manager=mytempcancel.employee AND mytempcancel.cancelby=mytempcancel.employee && mangager cancelled his own leave
				emailTo=mytempcancel.manager
				emailFor="Self"
            CASE mytempcancel.manager<>mytempcancel.employee AND mytempcancel.cancelby=mytempcancel.employee &&employee cancelled his own leave
				emailTo=mytempcancel.manager
				emailFor="Manager"
            CASE mytempcancel.manager<>mytempcancel.employee AND mytempcancel.cancelby=mytempcancel.manager &&manager cancelled employee's leave
				emailTo=mytempcancel.employee
				emailFor="Employee"
			OTHERWISE
				emailTo=mytempcancel.employee
				emailFor="Employee"
			ENDCASE	
			IF !EMPTY(emailTo)
				This.SendCancelLeaveEmailTo(emailTo, SENDMAIL_LEAVEREQUEST_NEW,emailFor)
			ENDIF	
		ENDSCAN	
		USE IN SELECT("mytempcanceldays")	 
		USE IN SELECT("mytempcancel")	 

		
	ENDPROC
	
	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadMyGroups()
		This.UploadGeneric("mygroups", "grcode", "uploadmygroups")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort
	PROCEDURE UploadMyManage()
		This.UploadGeneric("mymanage", "macode", "uploadmymanage")
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment; sort; test
	PROCEDURE UploadMyPays()
		THIS.UploadGeneric("mypays", "pay_pk", "uploadmypays")
		* CM screw referential integrity
		*        =MESSAGEBOX("Lic = "+this.licence+" Flag = "+IIF(lStat,"Ok,","NOT Ok"),0)
		*This.SelectData(This.Licence, "mypays")	BUG: not checking for success!
		** Process MyPays Here
		lStat = THIS.SelectData(THIS.Licence, "timesheet")
		SELECT timesheet
		SET ORDER TO tmid
		GO TOP
		SELECT mypays
		GO TOP
		DO WHILE NOT EOF("mypays")
			IF (mypays.pay_orig > 0) AND (mypays.pay_type == 1)
				DELETE FROM timesheet WHERE (timesheet.tmid == mypays.pay_pk)
				SELECT * FROM timesheet WHERE timesheet.tmid == mypays.pay_orig INTO CURSOR curTemTime READWRITE
				SELECT curTemTime
				REPLACE ALL tmid WITH mypays.pay_pk
				INSERT INTO timesheet SELECT * FROM curTemTime
				USE IN 'curTemTime'
			ENDIF
			SKIP IN 'mypays'
		ENDDO

		SELECT mypays
		REPLACE ALL pay_orig WITH 0
		GO TOP
		IF This.Licence == "67932"
			* don't clean up template records for SPCA
		ELSE
			DELETE FROM timesheet WHERE BETWEEN(tmid,1,99998) AND ;
				(tmid NOT IN (SELECT pay_pk FROM mypays WHERE mypays.pay_type = 1))
		ENDIF
	ENDPROC		

	*--------------------------------------------------------------------------------*

	*> comment; test
	* Upload staff has to notify users as well - so is not using generic process directly
	PROCEDURE UploadStaff()
		This.UploadGeneric("myStaff", "myWebCode", "uploadStaff")
		TRY
			LOCAL lnStaffCount

			IF This.SelectData(This.Licence, "myStaff") AND	This.SelectData(This.Licence, "timesheet")
				SELECT myStaff

				CALCULATE CNT() TO lnStaffCount

				IF !FILE("StaffCount.DBF")
					CREATE TABLE StaffCount (cLicence C(100), tDateTime T, nCount I)
					USE IN SELECT("StaffCount")
				ENDIF

				IF !USED("StaffCount")
					USE StaffCount SHARED IN 0
				ENDIF

				INSERT INTO StaffCount (cLicence, tDateTime, nCount) VALUES (This.Licence, This.GetLocalTime(DATETIME()), lnStaffCount)
			ENDIF
		CATCH TO loErr
			&&BUG: catch-all's need to be documented!  Also, now the above is checking like it should, it's probably not as needed...
			*Response.Write(outputException(loErr))		--I wonder where my outputException() went - it was useful in that it got the correct info out of them...
		ENDTRY

*		IF USED("myStaff")
*			IF USED("timesheet")
*				DELETE FROM timesheet WHERE BETWEEN(tmId,1,99999) AND ;
*										(tsEmp NOT IN (SELECT myWebCode FROM myStaff))
*			ENDIF
*		ENDIF
		
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment
	* Upload reports has to ??? - so is not using generic process directly
	PROCEDURE UploadReports()
		LOCAL uploadStatus as String, userInfo as String, lcString, lcFile, lcFileName, lcFileBuffer, lnStaffId

		lcFile = ""
		lcFileBuffer = Request.GetMultiPartFile("file", @lcFile)
		lcFileName = Request.GetMultipartFormVar("filename")

		uploadStatus = "" && error description (if any)
		userInfo = Request.GetMultipartFormVar("syszero")	&& we send a SYS(0) from the remote server to track who is using this function

		* First step - log the attempted upload
		This.LogSysZero(userInfo)

		* Test the checksum
		IF LEN(lcFileBuffer) != 0
			* Put the file somewhere...
			TRY
				lcFile = This.CompanyDataPath() + "PAYSLIPS\" + ALLTRIM(lcFileName)
				File2Var(lcFile, lcFileBuffer)
			CATCH
				uploadStatus = "File could not be saved to the company reports folder"
			ENDTRY
		ELSE
			uploadStatus = "File checksum does not match, file has not been uploaded"
		ENDIF
		IF !EMPTY(uploadStatus)
			Response.Write(uploadStatus)
			RETURN
		ENDIF
		uploadStatus = "OK"

		lnStaffId = FLOOR(VAL(lcFileName))
		IF lnStaffId != 0
			This.SendSiteUpdatedEmailTo(lnStaffId, SENDMAIL_PAYSLIP)
		ENDIF

		Response.Write(uploadStatus)

		* This.CheckReportCounts(lnStaffId)	&&QUESTION: why was this removed?  Document it!
	ENDPROC

	*################################################################################*
#DEFINE TOC_Downloads_

	*> +define: Downloads; comment!
	PROCEDURE Poll()
		LOCAL lcXml, lcPath, loTimeSheet, lcUserInfo, loIp, lcZipFile, lcXmlFile, llUnApproved, llDownloadAgain	&&, lcFile, llAllowed

		lcXml = ""
		lcUserInfo = Request.Form("syszero")
		llUnApproved = IIF(Request.Form("unapproved") == "true", .T., .F.)
		llDownloadAgain = IIF(Request.Form("downloadagain") == "true", .T., .F.)

		This.LogSysZero(lcUserInfo)
		*lcFile = This.CompanyDataPath() + "timeclock.mem"
		*llAllowed = FILE(lcFile)

		IF This.CheckRights("admin") &&AND llAllowed
			IF This.SelectData(This.Licence, "timesheet") AND This.SelectData(This.Licence, "mypays")
				LOCAL lcAgain, lcApprove

				IF llUnApproved
					lcApprove = " AND .T. "			&&QUESTION: should this rather be "AND (ISNULL(tsApproved) OR !tsApproved)", or does llUnApproved mean included rather than only-these?
				ELSE
					lcApprove = " AND tsApproved "
				ENDIF

				IF llDownloadAgain
					lcAgain = " AND .T. "
				ELSE
					lcAgain = " AND !tsDownload "
				ENDIF

				SELECT tsID, tsDate, tsStart, tsFinish, tsEmp, tsCostCent, tsBreak, tsUnits, tsApproved, tsDownload;
				FROM timesheet JOIN myPays ON timesheet.tsPay = myPays.pay_pk;
				WHERE myPays.pay_type = 2 &lcApprove. &lcAgain.;
				INTO CURSOR curPoll

				IF USED("curPoll") AND RECCOUNT("curPoll") != 0
					CURSORTOXML("curPoll", "lcXml", 1, 4, 0, '1')
					USE IN SELECT("curPoll")
				ELSE
					lcXml = "Problem: There are no available times on the website to download."
				ENDIF
				UPDATE timesheet SET tsDownload = .T. WHERE !tsDownload AND tsApproved	&&QUESTION: should this be in the above if?
			ELSE
				lcXml = "Problem: Could not open the times information on the website."
			ENDIF
		ELSE
			lcXml = "Problem: Bad Login" + TRANSFORM(This.Employee)
		ENDIF

		IF LOWER(Request.Form("zip")) == "true"
			* Compress the return files.
			loIp = NEWOBJECT("DynaZip", "ComaccOnlineSecurity.vcx")
			lcZipFile = ADDBS(FULLPATH(This.CompanyDataPath())) + "curpoll.zip"
			lcXmlFile = FORCEEXT(lcZipFile, "xml")
			TRY
				IF FILE(lcZipFile)
					DELETE FILE (lcZipFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			TRY
				IF FILE(lcXmlFile)
					DELETE FILE (lcXmlFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			STRTOFILE(lcXml, lcXmlFile)
			IF loIp.ZipFiles(lcZipFile, lcXmlFile) == 0
				* Cool it worked
				Response.TransmitFile(FORCEEXT(lcZipFile, "zip"), "application/zip")
			ELSE
				lcXml = "Problem: Could not compress the return file (timeclock)."
				Response.Write(lcXML)
			ENDIF
		ELSE
			Response.Write(lcXML)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment!
	PROCEDURE DownloadStaff() as String
		LOCAL lcXml as String, lcPath as String, loStaff, loIp, lcZipFile, lcXmlFile

		IF This.CheckRights("admin")
			lcXml = ""
			loStaff = Factory.GetStaffObject()
			loStaff.Open()
			loStaff.Execute([select myWebCode,myPayCode,myChanged,mySlipName,myAddress,mySuburb,myCity,myPhone,myEmail,myPostCode,myPassword,myMobile,myAddress2,myUdfl1,myUdfl2,myUdfd1,myUdfd2,myUdfc1,myUdfc2,myUdfc3,myUdfn1,myUdfm1,myBank from myStaff into cursor curDownloadStaff])
			IF USED("curDownloadStaff")
				CURSORTOXML("curDownloadStaff", "lcXml", 1, 4, 0, '1')
				* reset changed flag
				loStaff.Execute([update myStaff set myChanged = ""])
			ELSE
				lcXml = "Problem: There are no employees available on the website to download"
			ENDIF
			USE IN SELECT("curDownloadStaff")
		ELSE
			lcXml = "Problem: Bad Login" + TRANSFORM(This.Employee)
		ENDIF

		IF LOWER(Request.Form("zip")) == "true"
			* Compress the return files.
			loIp = NEWOBJECT("DynaZip", "ComaccOnlineSecurity.vcx")
			lcZipFile = ADDBS(FULLPATH(This.CompanyDataPath())) + "staff.zip"
			lcXmlFile = FORCEEXT(lcZipFile, "xml")

			TRY
				IF FILE(lcZipFile)
					DELETE FILE (lcZipFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			TRY
				IF FILE(lcXmlFile)
					DELETE FILE (lcXmlFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			STRTOFILE(lcXml, lcXmlFile)
			IF loIp.ZipFiles(lcZipFile, lcXmlFile) = 0
				* Cool it worked
				Response.TransmitFile(FORCEEXT(lcZipFile, "zip"), "application/zip")
				RETURN
			ELSE
				lcXml = "Problem: Could not compress the return file (staff)."
				Response.Write(lcXML)
			ENDIF
		ELSE
			Response.Write(lcXML)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> fix
	PROCEDURE DownloadLeaveRequests() as String
		LOCAL lcXml, lcPath, lcUserInfo, lcFile, llAllowed, loIp, lcZipFile, lcXmlFile

		lcXml = ""
		lcUserInfo = Request.Form("syszero")
		This.LogSysZero(lcUserInfo)

		lcFile = This.CompanyDataPath() + "leavemanagement.mem"
		llAllowed = FILE(lcFile)

		* HG 01/09/2009 TTP1328 added sick_nomed
		IF This.CheckRights("admin") AND llAllowed
			IF This.SelectData(This.Licence, "leaveRequests") AND This.SelectData(This.Licence, "leaveRequestDays")	AND This.SelectData(This.Licence, "leaveRequestStatus") AND This.SelectData(This.Licence, "mycancelreq")
				SELECT;
					leaveRequests.id, leaveRequests.sick_nomed, leaveRequestDays.id as dayID, leaverequests.employee, leavecode,;
					date, units, unitType, leaveRequestStatus.from, leaveRequestStatus.sent, leaverequests.cancelReq, leaverequests.cancelled,CAST(0 as I) as CancelBy;
				FROM;
					leaveRequests;
						JOIN leaveRequestDays ON leaveRequests.id == leaveRequestDays.leaveReqID;
						JOIN leaveRequestStatus ON leaveRequests.id == leaveRequestStatus.leaveReqID;
				WHERE;
					(!downloaded AND accepted AND "Accepted" $ leaveRequestStatus.subject) ;
				UNION ALL ;
				SELECT;
					leaveRequests.id, leaveRequests.sick_nomed, leaveRequestDays.id as dayID, leaverequests.employee, leaverequests.leavecode,;
					leaveRequestDays.date, leaveRequestDays.units, leaveRequestDays.unitType, leaveRequests.employee as from, mycancelreq.mldate as sent, ;
					leaverequests.cancelReq, leaverequests.cancelled, mycancelreq.CancelBy;
				FROM;
					leaveRequests;
						JOIN leaveRequestDays ON leaveRequests.id == leaveRequestDays.leaveReqID;
						JOIN mycancelreq ON leaveRequests.id == mycancelreq.mlcode;
						WHERE leaverequests.cancelReq AND !leaverequests.cancelled ;
				INTO cursor curRequests READWRITE

				IF USED("curRequests")
					IF RECCOUNT("curRequests") = 0
						INSERT INTO curRequests (id) VALUES (0)
						DELETE FROM curRequests WHERE id = 0
					ENDIF
					CURSORTOXML("curRequests", "lcXml", 1, 4, 0, '1')
					USE IN SELECT("curRequests")
				ELSE
					lcXml = "Problem: There are no available leave requests on the website to download."
				ENDIF
				UPDATE leaveRequests SET downloaded = .T. WHERE !downloaded AND accepted		&&QUESTION: should this be subject to the above if?
			ELSE
				lcXml = "Problem: Could not open the leave request information on the website."
			ENDIF
		ELSE
			lcXml = "Problem: Bad Login" + TRANSFORM(This.Employee)
		ENDIF

		IF LOWER(Request.Form("zip")) == "true"
			* Compress the return files.
			loIp = NEWOBJECT("DynaZip", "ComaccOnlineSecurity.vcx")
			lcZipFile = ADDBS(FULLPATH(This.CompanyDataPath())) + "leaverequests.zip"
			lcXmlFile = FORCEEXT(lcZipFile, "xml")

			&&FIXME: the other places this is done the following are done in TRY..CATCH's - why not here? (or why there?)
			IF FILE(lcZipFile)
				DELETE FILE (lcZipFile)
			ENDIF
			IF FILE(lcXmlFile)
				DELETE FILE (lcXmlFile)
			ENDIF

			STRTOFILE(lcXml, lcXmlFile)
			IF loIp.ZipFiles(lcZipFile, lcXmlFile) == 0
				* Cool, it worked
				Response.TransmitFile(FORCEEXT(lcZipFile, "zip"), "application/zip")
			ELSE
				lcXml = "Problem: Could not compress the return file (Leave Requests)."
				Response.Write(lcXML)
			ENDIF
		ELSE
			Response.Write(lcXML)
		ENDIF
	ENDPROC

	*--------------------------------------------------------------------------------*

	*> comment!
	PROCEDURE DownloadTimeSheets()
		LOCAL lcXml, lcPath, loTimeSheet, lcUserInfo, ldDateFrom, ldDateTo, lcFile, llAllowed, llAll, lnBatch, loIp, lcZipFile, lcXmlFile, lcDates

		lcXml = ""
		ldDateFrom = CTOD(Request.Form("DateFrom"))
		ldDateTo = CTOD(Request.Form("DateTo"))
		llAll = UPPER(ALLTRIM(Request.Form("DownloadAll"))) == ".T."
		lnBatch = FLOOR(VAL(Request.Form("Batch")))

		DO CASE
		CASE EMPTY(ldDateFrom) AND EMPTY(ldDateTo)
			lcDates = " .T. "
		CASE EMPTY(ldDateFrom)
			lcDates = [ tsDate >= ldDateFrom ]
		CASE EMPTY(ldDateTo)
			lcDates = [ tsDate <= ldDateTo ]
		OTHERWISE
			lcDates = [ BETWEEN(tsDate, ldDateFrom, ldDateTo) ]
		ENDCASE

		lcUserInfo = Request.Form("syszero")
		This.LogSysZero(lcUserInfo)

		lcFile = This.CompanyDataPath() + "timesheet.mem"
		llAllowed = FILE(lcFile)

		IF This.CheckRights("admin") AND llAllowed
			IF This.SelectData(This.Licence, "TimeSheet") AND This.SelectData(This.Licence, "myStaff")
				IF llAll
					SELECT * FROM timesheet WHERE tsPay = lnBatch AND tsEmp != 0 AND tsApproved AND &lcDates. INTO CURSOR curTimesheets
				ELSE
					SELECT * FROM timesheet WHERE tsPay = lnBatch AND tsEmp != 0 AND tsApproved AND &lcDates. AND !tsDownload INTO CURSOR curTimesheets
				ENDIF
				IF USED("curTimesheets") AND RECCOUNT("curTimesheets") != 0
					CURSORTOXML("curTimesheets", "lcXml", 1, 4, 0, '1')
					USE IN SELECT("curTimesheets")
				ELSE
					lcXml = "Problem: There are no available timesheets on the website to download."
				ENDIF
				IF llAll
					UPDATE timesheet SET tsDownload = .T. WHERE tsPay = lnBatch AND tsEmp != 0 AND &lcDates. AND tsApproved
				ELSE
					UPDATE timesheet SET tsDownload = .T. WHERE tsPay = lnBatch AND tsEmp != 0 AND &lcDates. AND tsApproved AND !tsDownload
				ENDIF
			ELSE
				lcXml = "Problem: Could not open the timesheet information on the website."
			ENDIF
		ELSE
			lcXml = "Problem: Bad Login" + TRANSFORM(This.Employee)
		ENDIF

		IF LOWER(Request.Form("zip")) == "true"
			* Compress the return files.
			loIp = NEWOBJECT("DynaZip", "ComaccOnlineSecurity.vcx")
			lcZipFile = ADDBS(FULLPATH(This.CompanyDataPath())) + "timesheets.zip"
			lcXmlFile = FORCEEXT(lcZipFile, "xml")

			TRY
				IF FILE(lcZipFile)
					DELETE FILE (lcZipFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			TRY
				IF FILE(lcXmlFile)
					DELETE FILE (lcXmlFile)
				ENDIF
			CATCH
				&&BUG: catch-all's should be documented!
			ENDTRY

			STRTOFILE(lcXml, lcXmlFile)
			IF loIp.ZipFiles(lcZipFile, lcXmlFile) = 0
				* Cool, it worked
				Response.TransmitFile(FORCEEXT(lcZipFile, "zip"), "application/zip")
			ELSE
				lcXml = "Problem: Could not compress the return file (timesheets)."
				Response.Write(lcXML)
			ENDIF
		ELSE
			Response.Write(lcXML)
		ENDIF
	ENDPROC

	*################################################################################*
#DEFINE TOC_Emails_

	*> +define: Emails; fix?
	PROCEDURE SendSiteUpdatedEmailTo(tnEmployee, tnType, tnMessage)
		LOCAL lcWorkArea, loIPStuff as wwipstuff, loStaff, llSendMail, llResult, lcText, loError
		LOCAL lcMessage

		IF VARTYPE(tnType) != 'N' OR tnType = 0
			RETURN
		ENDIF

		lcWorkArea = SELECT()
		loStaff = Factory.GetStaffObject()

		llSendMail = EVALUATE(AppSettings.Get("sendMail" + ALLTRIM(STR(tnType))))

		IF llSendMail AND loStaff.Load(tnEmployee)


			lcRecipient = ALLTRIM(loStaff.oData.myEmail)
			lcSubject = AppSettings.Get("subject" + ALLTRIM(STR(tnType)))
			lcMessage = AppSettings.Get("message" + ALLTRIM(STR(tnType)))
			lcReplyTo = AppSettings.Get("update_reply")

			IF !FILE("checkEmail.mem")	&&NOTE: this file lives in the root folder, sibling the company directories.
				TRY
				    DO wwdotnetbridge
				    LOCAL lobridge AS wwdotnetbridge
					lobridge=CREATEOBJECT("wwDotNetBridge","V4")
					*lcerror=SecureEmail(lobridge,SMTP_SERVER,.T.,SMTP_USER,SMTP_USER_PASSWORD,SMTP_SENDER_NAME,SMTP_SENDER_EMAIL,lcRecipient,lcSubject, lcMessage, "", .T.,lcReplyTo, tnType)
					lcerror=SecureEmail(lobridge,SMTP_SERVER,.T.,SMTP_USER,SMTP_USER_PASSWORD,SMTP_SENDER_NAME,SMTP_SENDER_EMAIL,lcRecipient,lcSubject, lcMessage, "", .F.,lcReplyTo, tnType, This.Licence)
					IF !EMPTY(lcerror)
					    This.AddError(lcerror)
					ENDIF	
				CATCH TO loError
					This.AddError(TRANSFORM(loError.ErrorNo) + ": " + loError.Message + "; " + loError.Details)
				ENDTRY
			ELSE
				LOCAL llResult
				llResult = loIpStuff.SendMail()
				IF !llResult
					STRTOFILE(TRANSFORM(DATETIME()) + ' ' + loIpStuff.cErrorMsg + CRLF, This.CompanyDataPath() + "emailErrors.txt", 1)
				ENDIF
			ENDIF
		ENDIF

		SELECT (lcWorkArea)
	ENDPROC
	
	
	
	PROCEDURE SendCancelLeaveEmailTo(tnEmployee, tnType, tcFor)
		LOCAL lcWorkArea, loIPStuff as wwipstuff, loStaff, llSendMail, llResult, lcText, loError
		LOCAL lcMessage

		IF VARTYPE(tnType) != 'N' OR tnType = 0
			RETURN
		ENDIF

		lcWorkArea = SELECT()
		loStaff = Factory.GetStaffObject()

		llSendMail = EVALUATE(AppSettings.Get("sendMail" + ALLTRIM(STR(tnType))))

		IF llSendMail AND loStaff.Load(tnEmployee)
	

			lcRecipient = ALLTRIM(loStaff.oData.myEmail)
			lcSubject = "mystaffinfo.com - Cancelled Leave Request"
			lcMessage = AppSettings.Get("message" + ALLTRIM(STR(tnType)))
			lcReplyTo = AppSettings.Get("update_reply")

			DO CASE
			CASE ALLTRIM(UPPER(tcfor))="EMPLOYEE"
				lcMessage = "Your leave request on the MyStaffInfo website has been cancelled by the Manager. Please visit https://mystaffinfo.myob.com to view the details."
			CASE ALLTRIM(UPPER(tcfor))="MANAGER"
				lcMessage = "A leave request you approved on the MyStaffInfo website has been cancelled by the Employee who requested it. Please visit https://mystaffinfo.myob.com to view the details."
			CASE ALLTRIM(UPPER(tcfor))="SELF"
				lcMessage = "Your leave request on the MyStaffInfo website has been cancelled by you. Please visit https://mystaffinfo.myob.com to view the details."
			OTHERWISE
				lcMessage = "Some leave request on the MyStaffInfo website has been cancelled. Please visit https://mystaffinfo.myob.com to view the details."
			ENDCASE

			IF !FILE("checkEmail.mem")	&&NOTE: this file lives in the root folder, sibling the company directories.
				TRY
					DO wwdotnetbridge
				    LOCAL lobridge AS wwdotnetbridge
					lobridge=CREATEOBJECT("wwDotNetBridge","V4")
					lcerror=SecureEmail(lobridge,SMTP_SERVER,.T.,SMTP_USER,SMTP_USER_PASSWORD,SMTP_SENDER_NAME,SMTP_SENDER_EMAIL,lcRecipient,lcSubject, lcMessage, "", .F.,lcReplyTo, tnType, This.Licence)
					IF !EMPTY(lcerror)
					    This.AddError(lcerror)
					ENDIF	

				CATCH TO loError
					This.AddError(TRANSFORM(loError.ErrorNo) + ": " + loError.Message + "; " + loError.Details)
				ENDTRY
			ELSE
				LOCAL llResult
				llResult = loIpStuff.SendMail()
				IF !llResult
					STRTOFILE(TRANSFORM(DATETIME()) + ' ' + loIpStuff.cErrorMsg + CRLF, This.CompanyDataPath() + "emailErrors.txt", 1)
				ENDIF
			ENDIF
		ENDIF

		SELECT (lcWorkArea)
	ENDPROC


	*################################################################################*
#DEFINE TOC_HTML_Controls_

	&&NOTE: these don't add CSS classNames, and if they need to it needs to be from a parameter... (maybe with a defined default)

	*> comment
	FUNCTION StaffDropDown(;
		tcName AS String, tuSelectedValue AS Variant, tnGroup AS Integer,;
		tcExtraAttribs AS String, tlHidePayrollUsers AS Boolean, tlIncludeCurrent AS Boolean,;
		tnIndent AS Integer, rnPrevRef AS Integer, rnNextRef AS Integer,;
		tlEveryoneEntry AS Boolean) AS String
		
		LOCAL lcSelect AS String
		LOCAL lcOnChange AS String
		LOCAL lcIndent AS String
		LOCAL lnPrevCode AS Integer
		LOCAL lnCount AS Integer
		LOCAL loOptions AS Object

		IF VARTYPE(tnGroup) != 'N'
			* List all staff
			This.GetEmployees("curStaff", tlIncludeCurrent)
			lcOnChange = ""
		ELSE
			* Group based
			lcOnChange = [ onchange="SubmitForm(this.form);"]
			IF !This.GetEmployeesByGroupCode(tnGroup, "curStaff")
				RETURN ""	&& bail out on error
			ENDIF
		ENDIF

		lnCount = 0

		IF tcExtraAttribs == "webFormDataQuery"
			loOptions = CREATEOBJECT("COLLECTION")

			IF tlEveryoneEntry
				loOptions.Add(EVERYONE_LABEL, TRANSFORM(EVERYONE_OPTION))
			ENDIF

			SELECT curStaff
			SCAN
				IF tlHidePayrollUsers AND curStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curStaff.myWebCode < 2	&& hide Admin user either way...
					LOOP
				ENDIF

				loOptions.Add(ALLTRIM(fullName), TRANSFORM(myWebCode))
			ENDSCAN
			USE IN SELECT("curstaff")

			RETURN loOptions
		ELSE
			IF EMPTY(tnIndent)
				tnIndent = 0
			ENDIF
			lcIndent = STRTRAN(SPACE(tnIndent), ' ', CHR(9))

			rnPrevRef = 0	&& These 2 are passed by reference if at all
			rnNextRef = 0

			lnPrevCode = 0

			lcSelect = lcIndent + [<select name="] + ALLTRIM(tcName) + '"' + lcOnChange + IIF(!EMPTY(tcExtraAttribs), ' ' + ALLTRIM(tcExtraAttribs), "") + [>] + CRLF

			lnCount = 0

			IF tlEveryoneEntry
				lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(EVERYONE_OPTION) + ["] + IIF(tuSelectedValue == EVERYONE_OPTION, [ selected="selected"], "") + [>] + This.HTMLEscape(EVERYONE_LABEL) + [</option>] + CRLF

				* Signal that the next arrow must be enabled for the first real entry
				lnPrevCode = EVERYONE_OPTION

				lnCount = lnCount + 1
			ENDIF

			SELECT curStaff
			SCAN
				IF tlHidePayrollUsers AND curStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curStaff.myWebCode < 2	&& hide Admin user either way...
					LOOP
				ENDIF

				lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(curstaff.mywebcode) + ["] + IIF(curStaff.myWebCode == tuSelectedValue, [ selected="selected"], "") + [>] + ALLTRIM(curStaff.fullName) + [</option>] + CRLF

				IF curStaff.myWebCode == tuSelectedValue
					* If the current staff is selected, save the previous one if any to rnPrevRef
					rnPrevRef = lnPrevCode
				ENDIF
				IF lnPrevCode == tuSelectedValue
					* If the previous staff was selected or if the Everyone entry is in use and was selected, save the currnet one if any to rnNextRef
					rnNextRef = curStaff.myWebCode
				ENDIF

				* Lastly, copy the current code
				lnPrevCode = curStaff.myWebCode

				lnCount = lnCount + 1
			ENDSCAN

			IF rnPrevRef == 0 AND rnNextRef == 0 AND lnCount > 1
				* current staff is not in the group - form will default to the first in the list, so we may need to fix rnNextRef
				lnCount = IIF(tlEveryoneEntry, 1, 0)	&& This means that the initial default will be the Everyone entry if it is present.

				SELECT curStaff
				SCAN
					IF tlHidePayrollUsers AND curStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curStaff.myWebCode < 2	&& hide Admin user either way...
						LOOP
					ENDIF

					IF lnCount == 1
						* Get the second entry for next
						rnNextRef = curStaff.myWebCode
						EXIT
					ENDIF

					lnCount = lnCount + 1
				ENDSCAN
			ENDIF

			lcSelect = lcSelect + lcIndent + [</select>] + CRLF
			USE IN SELECT("curstaff")

			IF lnCount <= 0
				lnCurrentStaff = 0
			ENDIF

			RETURN lcSelect
		ENDIF
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	* More accurately, Managers-of-current-user-DropDown..:
	FUNCTION ManagersDropDown(tcName, tnSelected, tnExcludeID, tcExtraAttribs)
		RETURN This.ManagersForDropDown(This.Employee, tcName, tnSelected, tnExcludeID, tcExtraAttribs)
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	FUNCTION ManagersForDropDown(tnEmployee, tcName, tnSelected, tnExcludeID, tcExtraAttribs)
		LOCAL lcOutput, loOptions

		lcOutput = ""

		IF This.GetPayrollUsersByEmployeeCode(tnEmployee, "curManagers", tnExcludeID)
			IF tcExtraAttribs == "webFormDataQuery"
				* Return the list of options for the webForm control
				loOptions = CREATEOBJECT("COLLECTION")

				IF USED("curManagers")
					SELECT curManagers
					SCAN
						loOptions.Add(ALLTRIM(fullName), TRANSFORM(myWebCode))
					ENDSCAN
				ENDIF

				RETURN loOptions
			ELSE
				* Construct the drop down list
				lcOutput = [<select name="] + ALLTRIM(tcName) + ["] + IIF(!EMPTY(tcExtraAttribs), ' ' + ALLTRIM(tcExtraAttribs), "") + [>] + CRLF

				IF USED("curManagers")
					SELECT curManagers
					SCAN
						lcOutput = lcOutput + [	<option value="] + TRANSFORM(myWebCode) + ["] + IIF(myWebCode == tnSelected, [ selected="selected"], "") + [>] + ALLTRIM(fullName) + [</option>] + CRLF
					ENDSCAN
				ENDIF

				lcOutput = lcOutput + [</select>] + CRLF
			ENDIF
		ENDIF

		RETURN lcOutput
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> comment
	FUNCTION AllManagersDropDown(tcName, tnSelected, tnExcludeID, tcExtraAttribs)
		LOCAL lcOutput, loOptions

		IF This.SelectData(This.Licence, "myManage")
			IF EMPTY(tnExcludeID)
				SELECT DISTINCT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) as FullName;
					FROM myStaff JOIN myManage ON maMyStaff == myWebCode;
					INTO CURSOR curAllManagers;
					ORDER BY FullName
			ELSE
				SELECT DISTINCT myWebCode, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) as FullName;
					FROM myStaff JOIN myManage ON maMyStaff == myWebCode;
					WHERE maMyStaff != tnExcludeID;
					INTO CURSOR curAllManagers;
					ORDER BY FullName
			ENDIF
		ENDIF

		IF tcExtraAttribs == "webFormDataQuery"
			* Return the list of options for the webForm control
			loOptions = CREATEOBJECT("COLLECTION")

			IF USED("curAllManagers")
				SELECT curAllManagers
				SCAN
					loOptions.Add(ALLTRIM(fullName), TRANSFORM(myWebCode))
				ENDSCAN
			ENDIF

			RETURN loOptions
		ELSE
			* construct the drop down list
			lcOutput = [<select name="] + ALLTRIM(tcName) + ["] + IIF(!EMPTY(tcExtraAttribs), ' ' + ALLTRIM(tcExtraAttribs), "") + [>] + CRLF

			IF USED("curAllManagers")
				SELECT curAllManagers
				SCAN
					lcOutput = lcOutput + [	<option value="] + TRANSFORM(myWebCode) + ["] + IIF(myWebCode == tnSelected, [ selected="selected"], "") + [>] + ALLTRIM(FullName) + [</option>] + CRLF
				ENDSCAN
				USE IN SELECT("curAllManagers")
			ENDIF

			lcOutput = lcOutput + [</select>]

			RETURN lcOutput
		ENDIF
	ENDFUNC

	*--------------------------------------------------------------------------------*

	* Selects the list of available security groups for the currently logged in user.  Generates nothing if there are none.
	FUNCTION GroupDropDown(tcName, tuSelectedValue, tcExtraAttribs, tlStatic, tlShowMyDetailsForPayrollUsers, tnIndent) as String
		* Selects the list of available security groups for the currently logged in user.
		LOCAL lcSelect, lcScript, llPayrollUser, lcIndent

		IF EMPTY(tnIndent)
			tnIndent = 0
		ENDIF
		lcIndent = STRTRAN(SPACE(tnIndent), ' ', CHR(9))

		* This.GetGroupsByEmployeeCode(This.Employee, "curGroups")
		This.GetGroupsForManager(This.Employee, "curGroups")

		lcScript = ""
		IF !tlStatic
			* Include javascript sumbit of form
			lcScript = [ onchange="SubmitForm(this.form);"]
		ENDIF

		lcSelect = lcIndent + [<select name="] + ALLTRIM(tcName) + '"' + IIF(EMPTY(tcExtraAttribs), "", ' ' + ALLTRIM(tcExtraAttribs)) + lcScript + [>] + CRLF

		llPayrollUser = This.IsPayrollUser(This.Employee)
		IF !llPayrollUser OR tlShowMyDetailsForPayrollUsers
			lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(MY_DETAILS_GROUP) + ["] + IIF(tuSelectedValue == MY_DETAILS_GROUP, [ selected="selected"], "") + [>] + This.HTMLEscape(MY_DETAILS_LABEL) + [</option>] + CRLF
		ENDIF

		SELECT curGroups
		SCAN
			lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(grCode) + ["] + IIF(grCode == tuSelectedValue, [ selected="selected"], "") + [>] + This.HTMLEscape(ALLTRIM(grName)) + [</option>] + CRLF
		ENDSCAN

		lcSelect = lcSelect + lcIndent + [</select>] + CRLF
		USE IN SELECT("curGroups")

		RETURN lcSelect
	ENDFUNC


    ******************************

	* Selects the list of available security groups for the currently logged in user.  Generates nothing if there are none.
	FUNCTION repgroupdropdown(tcname, tuselectedvalue, tcextraattribs, tlstatic, tlshowmydetailsforpayrollusers, tnindent) AS STRING
	* Selects the list of available security groups for the currently logged in user.
	LOCAL lcselect, lcscript, llpayrolluser, lcindent
	LOCAL tccurgroupopt, tNSelGroup, tSSelGroup
	
	
	tccurgroupopt = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrGrOpt")))>0,ALLTRIM(REQUEST.querystring("CurrGrOpt")),"Selected")
	tNSelGroup = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrGroup")))>0,EVALUATE(ALLTRIM(REQUEST.querystring("CurrGroup"))),0)
	tSSelGroup = ALLTRIM(STR(tNSelGroup,10))
	
	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF
	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))
	
	poretainlist = factory.getretainlistobject()


	IF UPPER(tccurgroupopt)="SELECTED"
	* This.GetGroupsByEmployeeCode(This.Employee, "curGroups")
		THIS.getgroupsformanager(THIS.employee, "curGroups")

		lcscript = ""
		IF !tlstatic
	       * Include javascript sumbit of form
			lcscript = [ onchange="SubmitForm(this.form);"]
		ENDIF

		lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' + IIF(EMPTY(tcextraattribs), "", ' ' + ALLTRIM(tcextraattribs)) + lcscript + [>] + crlf

		llpayrolluser = THIS.ispayrolluser(THIS.employee)
		IF !llpayrolluser OR tlshowmydetailsforpayrollusers
			lcselect = lcselect + lcindent + [	<option value="] + TRANSFORM(my_details_group) + ["] + IIF(tNSelGroup == -1, [ selected="selected"], "") + [>] + THIS.htmlescape(my_details_label) + [</option>] + crlf
		ENDIF

		SELECT curgroups
		SCAN
			lcselect = lcselect + lcindent + [	<option value="] + TRANSFORM(grcode) + ["] + IIF( ALLTRIM(STR(grcode,10)) == tSSelGroup, [ selected="selected"], "") + [>] + THIS.htmlescape(ALLTRIM(grname)) + [</option>] + crlf
			
		ENDSCAN

		lcselect = lcselect + lcindent + [</select>] + crlf
		USE IN SELECT("curGroups")

	ELSE
		lcselect = "-- All groups selected --"
	ENDIF
	RETURN lcselect
	ENDFUNC


	*--------------------------------------------------------------------------------*
	FUNCTION reptypedropdown(tcname, tuselectedvalue, tlstatic, tnindent) AS STRING

	LOCAL lcselect, lcscript, llpayrolluser, lcindent

	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF
	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))

	lcscript = ""


	IF !tlstatic
* Include javascript sumbit of form
		lcscript = [ onchange="SubmitForm(this.form);"]
	ENDIF

	lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' +  "" + lcscript + [>] + crlf

	lcselect = lcselect + lcindent + [	<option value="All"] + IIF(UPPER(tuselectedvalue) == "ALL", [ selected="selected"], "") + [>] + THIS.htmlescape("All") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Batch"] + IIF(UPPER(tuselectedvalue) == "BATCH", [ selected="selected"], "") + [>] + THIS.htmlescape("Batch") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Template"] + IIF(UPPER(tuselectedvalue) == "TEMPLATE", [ selected="selected"], "") + [>] + THIS.htmlescape("Template") + [</option>] + crlf

	lcselect = lcselect + lcindent + [</select>] + crlf


	RETURN lcselect


	ENDFUNC



	*--------------------------------------------------------------------------------*
	FUNCTION repapprovedropdown(tcname, tuselectedvalue, tlstatic, tnindent) AS STRING

	LOCAL lcselect, lcscript, llpayrolluser, lcindent, lccurapprove

	lccurapprove = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepApprove")))>0,ALLTRIM(REQUEST.querystring("CurrentRepApprove")),"Selected")

	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF
	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))

	lcscript = ""

	IF !tlstatic
* Include javascript sumbit of form
		lcscript = [ onchange="SubmitForm(this.form);"]
	ENDIF

	lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' +  "" + lcscript + [>] + crlf

	lcselect = lcselect + lcindent + [	<option value="Both"] + IIF(UPPER(lcCurApprove) == "BOTH", [ selected="selected"], "") + [>] + THIS.htmlescape("Both") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Yes"] + IIF(UPPER(lcCurApprove) == "YES", [ selected="selected"], "") + [>] + THIS.htmlescape("Yes") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="No"] + IIF(UPPER(lcCurApprove) == "NO", [ selected="selected"], "") + [>] + THIS.htmlescape("No") + [</option>] + crlf

	lcselect = lcselect + lcindent + [</select>] + crlf


	RETURN lcselect


	ENDFUNC

*--------------------------------------------------------------------------------*


*--------------------------------------------------------------------------------*
	FUNCTION repstatusdropdown(tcname, tuselectedvalue, tlstatic, tnindent) AS STRING

	LOCAL lcselect, lcscript, llpayrolluser, lcindent

	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF
	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))

	lcscript = ""

	IF !tlstatic
* Include javascript sumbit of form
		lcscript = [ onchange="SubmitForm(this.form);"]
	ENDIF

	lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' +  "" + lcscript + [>] + crlf

	lcselect = lcselect + lcindent + [	<option value="All"] + IIF(UPPER(tuselectedvalue) == "ALL", [ selected="selected"], "") + [>] + THIS.htmlescape("All") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Open"] + IIF(UPPER(tuselectedvalue) == "OPEN", [ selected="selected"], "") + [>] + THIS.htmlescape("Open") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Closed"] + IIF(UPPER(tuselectedvalue) == "CLOSED", [ selected="selected"], "") + [>] + THIS.htmlescape("Closed") + [</option>] + crlf

	lcselect = lcselect + lcindent + [</select>] + crlf


	RETURN lcselect


	ENDFUNC


*--------------------------------------------------------------------------------*


*--------------------------------------------------------------------------------*
	FUNCTION repgroupoptdropdown(tcname, tuselectedvalue, tlstatic, tnindent) AS STRING

	LOCAL lcselect, lcscript, llpayrolluser, lcindent, tccurgroupopt

	tccurgroupopt = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrGrOpt")))>0,ALLTRIM(REQUEST.querystring("CurrGrOpt")),"Selected")

	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF
	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))

	lcscript = ""

	IF !tlstatic
* Include javascript sumbit of form
		lcscript = [ onchange="SubmitForm(this.form);"]
	ENDIF

	lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' +  "" + lcscript + [>] + crlf

	lcselect = lcselect + lcindent + [	<option value="All"] + IIF(UPPER(tuselectedvalue) == "ALL", [ selected="selected"], "") + [>] + THIS.htmlescape("All") + [</option>] + crlf
	lcselect = lcselect + lcindent + [	<option value="Selected"] + IIF(UPPER(tccurgroupopt) == "SELECTED", [ selected="selected"], "") + [>] + THIS.htmlescape("Selected") + [</option>] + crlf

	lcselect = lcselect + lcindent + [</select>] + crlf


	RETURN lcselect


	ENDFUNC



	
	*--------------------------------------------------------------------------------*
	* JA 06/11/2012 Function to populate batch and template name

	FUNCTION paytempdropdown(tcname AS STRING, ruselectedvalue AS INTEGER, tlstatic, tnindent, tcextraattribs) AS STRING
	LOCAL tcstatus, tctype, tccurrentrep
	LOCAL lcselect, lcbatchselect, lctempselect, lccompselect, lctype, lcstatus, lcscript, lcindent, lnfirstvalue, llfound
	tctype = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepType")))>1,ALLTRIM(REQUEST.querystring("CurrentRepType")),"Batch")
	tcstatus = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepStatus")))>1,ALLTRIM(REQUEST.querystring("CurrentRepStatus")),"Open")
	tccurrentrep = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepName")))>0,ALLTRIM(REQUEST.querystring("CurrentRepName")),"-1")

	IF EMPTY(tnindent)
		tnindent = 0
	ENDIF

	lcindent = STRTRAN(SPACE(tnindent), ' ', CHR(9))
    lcSelect = ""
      
	IF THIS.selectdata(THIS.licence, "myPays")
		DO CASE
		CASE UPPER(tctype) == "ALL"
			lctype = ".T."
		CASE UPPER(tctype) == "BATCH"
			lctype = "pay_type == 2"
		CASE UPPER(tctype) == "TEMPLATE"
			lctype = "pay_type == 1"
		OTHERWISE
			lctype = "pay_type == 2"
		ENDCASE
		DO CASE
		CASE UPPER(tcstatus) == "ALL"
			lcstatus = ".T."
		CASE UPPER(tcstatus) == "OPEN"
			lcstatus = "pay_status == 1"
		CASE UPPER(tcstatus) == "CLOSED"
			lcstatus = "pay_status == 2"
		OTHERWISE
			lcstatus = "pay_status == 1"
		ENDCASE
		IF !tlstatic
	        *Include javascript sumbit of form
			lcscript = [ onchange="SubmitForm(this.form);"]
		ELSE
			lcscript = ""
		ENDIF
		SELECT *;
			FROM mypays;
			WHERE &lctype. AND &lcstatus.;
			INTO CURSOR curpays;
			ORDER BY pay_name
		IF RECCOUNT("curPays") != 0
			lnfirstvalue = 0
			llfound = .F.
			lcselect = lcindent + [<select name="] + ALLTRIM(tcname) + '"' + lcscript + IIF(EMPTY(tcextraattribs), "", ' ' + ALLTRIM(tcextraattribs)) + [>] + crlf
			SELECT curpays
			SCAN
				IF !EMPTY(lnfirstvalue)
					lnfirstvalue = pay_pk
				ENDIF
				lcselect = lcselect + lcindent + [	<option value="] + TRANSFORM(pay_pk) + ["]
				IF pay_pk == VAL(tccurrentrep)
					lcselect = lcselect + [ selected="selected"]
					llfound = .T.
				ENDIF
				lcselect = lcselect + [>] + ALLTRIM(pay_name) + [</option>] + crlf
			ENDSCAN
			lcselect = lcselect + lcindent + [</select>] + crlf
			IF !llfound
				ruselectedvalue = lnfirstvalue
			ENDIF
		ENDIF
		USE IN SELECT("curPays")
	ENDIF
	IF EMPTY(lcSelect)
	   lcSelect = "-- There are no records --"
	ENDIF
	RETURN lcselect
	ENDFUNC

	*--------------------------------------------------------------------------------*
	
	

	* Selects the list of available pays.  Generates nothing if there are none matching.
	* ruSelectedValue is set to the value of the first listed Pay if any iff it is not found in the list and it was passed by reference.
	FUNCTION PayDropDown(tcName as String, ruSelectedValue as Integer, tcType, tcStatus, tlStatic, tnIndent, tcExtraAttribs) as String
		LOCAL lcSelect, lcType, lcStatus, lcScript, lcIndent, lnFirstValue, llFound

		IF EMPTY(tnIndent)
			tnIndent = 0
		ENDIF
		lcIndent = STRTRAN(SPACE(tnIndent), ' ', CHR(9))

		lcSelect = ""

		IF This.SelectData(This.Licence, "myPays")
			DO CASE
				CASE tcType == "all"
					lcType = ".T."
				CASE tcType == "pay"
					lcType = "pay_type == 2"
				CASE tcType == "template"
					lcType = "pay_type == 1"
				OTHERWISE
					lcType = "pay_type == 2"
			ENDCASE

			tcStatus = EVL(tcStatus, "open")
			DO CASE
				CASE tcStatus == "all"
					lcStatus = ".T."
				CASE tcStatus == "open"
					lcStatus = "pay_status == 1"
				CASE tcStatus == "closed"
					lcStatus = "pay_status == 2"
				OTHERWISE
					lcStatus = "pay_status == 1"
			ENDCASE

			IF !tlStatic
				* Include javascript sumbit of form
				lcScript = [ onchange="SubmitForm(this.form);"]
			ELSE
				lcScript = ""
			ENDIF

			SELECT *;
				FROM myPays;
				WHERE &lcType. AND &lcStatus.;
				INTO CURSOR curPays;
				ORDER BY pay_name

			IF RECCOUNT("curPays") != 0
				lnFirstValue = 0
				llFound = .F.
				lcSelect = lcIndent + [<select name="] + ALLTRIM(tcName) + '"' + lcScript + IIF(EMPTY(tcExtraAttribs), "", ' ' + ALLTRIM(tcExtraAttribs)) + [>] + CRLF
				SELECT curPays
				SCAN
					IF !EMPTY(lnFirstValue)
						lnFirstValue = pay_pk
					ENDIF
					lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(pay_pk) + ["]
					IF pay_pk == ruSelectedValue
						lcSelect = lcSelect + [ selected="selected"]
						llFound = .T.
					ENDIF
					lcSelect = lcSelect + [>] + ALLTRIM(pay_name) + [</option>] + CRLF
				ENDSCAN
				lcSelect = lcSelect + lcIndent + [</select>] + CRLF

				IF !llFound
					ruSelectedValue = lnFirstValue
				ENDIF
			ENDIF

			USE IN SELECT("curPays")
		ENDIF

		RETURN lcSelect
	ENDFUNC

	*--------------------------------------------------------------------------------*
	* 17/05/2010  CMGM  MRD 4.2.2.2  New function to created the "Previous Pays" dropdown list
	* Selects the list of available previous pays including the current pay.  Generates "<Select>" if there are none matching.
	FUNCTION PreviousPayDropDown(tcName as String, ruSelectedValue as Integer, tcType, tcStatus, tlStatic, tnIndent, tcExtraAttribs) as String
		LOCAL lcSelect, lcType, lcStatus, lcScript, lcIndent, lnFirstValue, llFound

		IF EMPTY(tnIndent)
			tnIndent = 0
		ENDIF
		lcIndent = STRTRAN(SPACE(tnIndent), ' ', CHR(9))

		lcSelect = ""

		IF This.SelectData(This.Licence, "myPays")
			DO CASE
				CASE tcType == "all"
					lcType = ".T."
				CASE tcType == "pay"
					lcType = "pay_type == 2"
				CASE tcType == "template"
					lcType = "pay_type == 1"
				OTHERWISE
					lcType = "pay_type == 2"
			ENDCASE

			tcStatus = EVL(tcStatus, "all")
			DO CASE
				CASE tcStatus == "all"
					lcStatus = ".T."
				CASE tcStatus == "open"
					lcStatus = "pay_status == 1"
				CASE tcStatus == "closed"
					lcStatus = "pay_status == 2"
				OTHERWISE
					lcStatus = ".T."
			ENDCASE

			IF !tlStatic
				* Include javascript sumbit of form
				lcScript = [ onchange="SubmitForm(this.form);"]
			ELSE
				lcScript = ""
			ENDIF
	
			*MY 24/10/2012 (US7737)
			SELECT *;
				FROM myPays;
				WHERE &lcType. AND &lcStatus.;
				INTO CURSOR curPreviousTmp;
				ORDER BY pay_date desc

			lnFld = AFIELDS(laFld,"myPays")
			CREATE CURSOR curPreviousPays FROM ARRAY laFld

			lnCntr = 0
			SELECT curPreviousTmp
			GO top
			DO WHILE NOT EOF() AND lnCntr < 30
 				lnCntr = lnCntr + 1
				SCATTER memvar
				INSERT INTO curPreviousPays FROM memvar
				SKIP
			ENDDO
			USE IN 'curPreviousTmp'
			
			SELECT curPreviousPays
			GO top

			lcSelect = lcIndent + [<select name="] + ALLTRIM(tcName) + '"' + lcScript + IIF(EMPTY(tcExtraAttribs), "", ' ' + ALLTRIM(tcExtraAttribs)) + [>] + CRLF

			* Will at least display "<Select>"
            IF tcType <> "template"
  				lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(SELECT_VALUE) + ["] + ;
  							IIF(ruSelectedValue == SELECT_VALUE, [ selected="selected"], "") + [>] + ;
 							This.HTMLEscape(SELECT_LABEL) + [</option>] + CRLF
            ELSE
  				lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(SELECT_VALUE) + ["] + IIF(ruSelectedValue == SELECT_VALUE, [ selected="selected"], "") + [>] + This.HTMLEscape(DEF_LABEL_TEMPLATE) + [</option>] + CRLF
            ENDIF

			IF RECCOUNT("curPreviousPays") != 0
				lnFirstValue = 0
				llFound = .F.
				
				SELECT curPreviousPays
				SCAN
					IF !EMPTY(lnFirstValue)
						lnFirstValue = pay_pk
					ENDIF
					lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(pay_pk) + ["]
					IF pay_pk == ruSelectedValue
						lcSelect = lcSelect + [ selected="selected"]
						llFound = .T.
					ENDIF
					lcSelect = lcSelect + [>] + ALLTRIM(pay_name) + [</option>] + CRLF
				ENDSCAN

				IF !llFound
					ruSelectedValue = lnFirstValue
				ENDIF
			ENDIF
			lcSelect = lcSelect + lcIndent + [</select>] + CRLF
		
			USE IN SELECT("curPreviousPays")
		ENDIF

		RETURN lcSelect
	ENDFUNC

	*================================================================================*
	* 30/03/2012  RAJ - Add new Template Functions
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE TemplateDropDown(tcName as String, ruSelectedValue as Integer, tcType, tcStatus,;
	                           tlStatic, tnIndent, tcExtraAttribs, tcCaller, lnCurrentStaff) as String

		LOCAL lcSelect, lcType, lcStatus, lcScript, lcIndent, lnFirstValue, llFound, lnChk

		IF EMPTY(tnIndent)
			tnIndent = 0
		ENDIF

		lcIndent = STRTRAN(SPACE(tnIndent), ' ', CHR(9))
		lcSelect = ""

		IF This.SelectData(This.Licence, "myPays") AND This.SelectData(This.Licence, "timesheet")

			DO CASE
				CASE tcType == "all"
					lcType = ".T."
				CASE tcType == "pay"
					lcType = "pay_type == 2"
				CASE tcType == "template"
					lcType = "pay_type == 1"
				OTHERWISE
					lcType = "pay_type == 2"
			ENDCASE

			tcStatus = EVL(tcStatus, "all")
			DO CASE
				CASE tcStatus == "all"
					lcStatus = ".T."
				CASE tcStatus == "open"
					lcStatus = "pay_status == 1"
				CASE tcStatus == "closed"
					lcStatus = "pay_status == 2"
				OTHERWISE
					lcStatus = ".T."
			ENDCASE

			IF !tlStatic
				* Include javascript sumbit of form
				lcScript = [ onchange="SubmitForm(this.form);"]
			ELSE
				lcScript = ""
			ENDIF

			IF used('CurTemplates')
				USE IN 'CurTemplates'
			ENDIF

			IF tcCaller == "time_entry"
			* MY - 22/11/2012 - fix template refresh issue for <my details> and <every one> 
*!*					IF lnCurrentStaff > 0
*!*						IF This.IsManager(This.Employee) AND This.CheckRights("TEM_APPLY_M")
*!*							SELECT * FROM myPays WHERE (pay_type == 1 AND pay_status == 1) AND;
*!*									(pay_pk IN (SELECT tmId FROM timesheet WHERE tsEmp == lnCurrentStaff));
*!*									 INTO CURSOR CurTemplates ORDER BY pay_name READWRITE
*!*						ENDIF
*!*					ELSE
*!*						IF This.IsManager(This.Employee) AND This.CheckRights("TEM_APPLY_M")
*!*							SELECT * FROM myPays WHERE (pay_type == 1 AND pay_status == 1) AND;
*!*								(pay_pk IN (SELECT tmId FROM timesheet)) ORDER BY pay_name;
*!*								 INTO CURSOR CurTemplates READWRITE
*!*						ENDIF
*!*					ENDIF

				IF This.IsManager(This.Employee) AND This.CheckRights("TEM_APPLY_M")
					IF pnCurrentGroup = -1
						*<my deatails> for manager
						lnCurrentStaff = This.employee
					ENDIF 
					IF lnCurrentStaff > 0
						SELECT * FROM myPays WHERE (pay_type == 1 AND pay_status == 1) AND;
								(pay_pk IN (SELECT tmId FROM timesheet WHERE tsEmp == lnCurrentStaff));
								 INTO CURSOR CurTemplates ORDER BY pay_name READWRITE
					ELSE
						* <everyone> select all templates used by one of nominated group member, not all groups
						SELECT * FROM myPays ;
							WHERE pay_type = 1 AND pay_status = 1 AND ;
								pay_pk In ( ;
									SELECT tmid FROM timesheet ;
										JOIN myteams ON timesheet.tsemp = myteams.tmmystaff  ;
										JOIN mygroups ON myteams.tmmygroups = mygroups.grcode AND mygroups.grcode = pnCurrentGroup)  ;
							ORDER BY pay_name;
							INTO CURSOR CurTemplates READWRITE
					ENDIF
				ENDIF 

				IF NOT USED('CurTemplates')
					SELECT mypays
					lnFldCnt = aFields(laFld,"mypays")
					CREATE CURSOR curTemplates FROM ARRAY laFld
				ENDIF
			ELSE
				SELECT * FROM myPays WHERE (pay_type == 1 AND pay_status == 1);
						 INTO CURSOR CurTemplates ORDER BY pay_name READWRITE
			ENDIF

						
			* Will at least display "<My Template>" - maybe!!!! (Only for Timesheet Entry)
			llFound = .F.
			IF tcCaller == "time_entry"
				IF lnCurrentStaff > 0
					lnChk = MY_TEMPLATE
					SET ORDER TO tmId IN 'timesheet'
					IF SEEK(lnChk,'timesheet')
						DO WHILE NOT EOF('timesheet') AND (timesheet.tmId == lnChk)
							IF (timesheet.tsEmp == lnCurrentStaff)
								llFound = .T.
								EXIT
							ENDIF
							SKIP IN 'timesheet'
						ENDDO
					ENDIF
				ELSE
					llFound = .T.
				ENDIF
			ELSE
				llFound = .T.
			ENDIF
			
			IF llFound
				SELECT curTemplates
				APPEND BLANK
				replace pay_pk     WITH MY_TEMPLATE
				replace pay_name   WITH DEF_LABEL_TEMPLATE
				replace pay_status WITH 1
				replace pay_type   WITH 1
				replace pay_orig   WITH 0
			ENDIF

			IF RECCOUNT("curTemplates") > 0
				lcSelect = lcIndent + [<select name="] + ALLTRIM(tcName) + '"' + lcScript +;
								 IIF(EMPTY(tcExtraAttribs), "", ' ' + ALLTRIM(tcExtraAttribs)) + [>] + CRLF
			
				lnFirstValue = 0
				llFound = .F.
				
				SELECT curTemplates
				GO top
				DO WHILE NOT EOF()
					IF !EMPTY(lnFirstValue)
						lnFirstValue = pay_pk
					ENDIF
					lcSelect = lcSelect + lcIndent + [	<option value="] + TRANSFORM(pay_pk) + ["]
					IF pay_pk == ruSelectedValue
						lcSelect = lcSelect + [ selected="selected"]
						llFound = .T.
					ENDIF
					lcSelect = lcSelect + [>] + This.HTMLEscape(ALLTRIM(pay_name)) + [</option>] + CRLF
					SKIP
				ENDDO

				IF !llFound
					GO top
					ruSelectedValue = pay_pk
					lcKey = [	<option value="] + TRANSFORM(pay_pk) + ["]
					lcSelect = STRTRAN(lcSelect,lcKey,lcKey + [ selected="selected"])
				ENDIF
				lcSelect = lcSelect + lcIndent + [</select>] + CRLF
			ENDIF
			
			USE IN SELECT("curTemplates")

		ENDIF

		RETURN lcSelect
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE GetDayName(tnDay) as String

		IF tnDay = 1
			RETURN "Monday"
		ENDIF
		
		IF tnDay = 2
			RETURN "Tuesday"
		ENDIF

		IF tnDay = 3
			RETURN "Wednesday"
		ENDIF

		IF tnDay = 4
			RETURN "Thursday"
		ENDIF

		IF tnDay = 5
			RETURN "Friday"
		ENDIF

		IF tnDay = 6
			RETURN "Saturday"
		ENDIF

		IF tnDay = 7
			RETURN "Sunday"
		ENDIF
	
		RETURN ""
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE GetDayNumber(tcDay) as Integer

		IF UPPER(ALLTRIM(tcDay)) == "MONDAY"
			RETURN 1
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "TUESDAY"
			RETURN 2
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "WEDNESDAY"
			RETURN 3
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "THURSDAY"
			RETURN 4
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "FRIDAY"
			RETURN 5
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "SATURDAY"
			RETURN 6
		ENDIF
		IF UPPER(ALLTRIM(tcDay)) == "SUNDAY"
			RETURN 7
		ENDIF

		RETURN -1
	ENDPROC

	
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	&&NOTE: this page handles posts to support the buildup of formRows...

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	&&NOTE: this page handles posts to support the buildup of formRows...
	PROCEDURE timesheetreportpage (tbReport)
	PRIVATE popage, postaff, potypes
	PRIVATE pncurrgroup, poretainlist, plmanager, pncurrentreptype, pncurrentrepstatus, pncurrentrepapprove, pncurrgropt, pncurrentname
	PRIVATE pcrepopt
	PRIVATE poleavecodes, poothercodes, poallowcodes, powagecodes, pocostcentres, pojobcodes

	LOCAL lcaction, lotype, lnfieldcount, lni, lnj, lncount, lnat, lnPay_PK
	LOCAL ARRAY lafields[13], ladefaults[13]

	IF !(THIS.selectdata(THIS.licence, "myStaff");
			AND THIS.selectdata(THIS.licence, "myPays");
			AND THIS.selectdata(THIS.licence, "myGroups");
			AND THIS.selectdata(THIS.licence, "wageType");
			AND THIS.selectdata(THIS.licence, "allow");
			AND THIS.selectdata(THIS.licence, "costCent");
			AND THIS.selectdata(THIS.licence, "timesheet"))	&&TODO: add new table for job-whateverItIs
		THIS.adderror("Page Setup Failed!")
	ELSE
		postaff = factory.getstaffobject()

		IF !postaff.LOAD(THIS.employee)
			THIS.adderror("Failed to load Employee record: " + postaff.cerrormsg)
		ELSE
			poretainlist = factory.getretainlistobject()

			plmanager = .F.
			pncurrgroup = 0

			pncurrentreptype = IIF(LEN(ALLTRIM(REQUEST.querystring("CurrentRepType")))>0,ALLTRIM(REQUEST.querystring("CurrentRepType")),"Batch")
			pncurrentrepstatus = IIF(LEN(ALLTRIM(REQUEST.querystring("CurrentRepStatus")))>0,ALLTRIM(REQUEST.querystring("CurrentRepStatus")),"Open")
			pncurrentrepapprove = IIF(LEN(ALLTRIM(REQUEST.querystring("CurrentRepApprove")))>0,ALLTRIM(REQUEST.querystring("CurrentRepApprove")),"Both")
			pncurrgropt = IIF(LEN(ALLTRIM(REQUEST.querystring("CurrGrOpt")))>0,ALLTRIM(REQUEST.querystring("CurrGrOpt")),"Selected")
			pncurrgroup = IIF(LEN(ALLTRIM(REQUEST.querystring("CurrGroup")))>0,EVALUATE(ALLTRIM(REQUEST.querystring("CurrGroup"))),-1)
			pncurrentrepname = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepName")))>0,ALLTRIM(REQUEST.querystring("CurrentRepName")),"-1")
			
			this.Timesheet_PayType = 0
			IF EMPTY(pncurrentreptype) OR pncurrentreptype = "Batch"
				this.Timesheet_PayType = 2
			ELSE
				lnPay_PK = IIF (LEN(ALLTRIM(REQUEST.querystring("CurrentRepName")))>0,ALLTRIM(REQUEST.querystring("CurrentRepName")),"-1")
				SELECT mypays
				IF VAL(lnPay_PK) > 0
					LOCATE FOR pay_pk = VAL(lnPay_PK)
					IF FOUND()
						this.Timesheet_PayType = pay_type
					ENDIF
				ELSE  
					LOCATE 
					IF FOUND()
						this.Timesheet_PayType = pay_type
					ENDIF 
				ENDIF
			ENDIF
		ENDIF
	ENDIF

	IF tbReport
		LOCAL lcOutputFile, lcmycountry  AS STRING
		LOCAL reportsession  AS OBJECT
		LOCAL returnvalue  AS INTEGER

		* prepare data
		PROCESS.selectdata(PROCESS.licence,  "timesheet")
		PROCESS.selectdata(PROCESS.licence,  "costcent")
		PROCESS.selectdata(PROCESS.licence,  "allow")
		PROCESS.selectdata(PROCESS.licence,  "wagetype")
		PROCESS.selectdata(PROCESS.licence,  "mystaff")
		PROCESS.selectdata(PROCESS.licence,  "mygroups")
		PROCESS.selectdata(PROCESS.licence,  "myteams")
		* prepare parameter
		SELECT mystaff
		LOCATE FOR  mywebcode > 	1000000
		lcmycountry = mystaff.mycountry
		* prepare report cursor
		LOCAL lcMsg, lcFileName, lcPathName AS STRING
		lcMsg = THIS.TimesheetReportProc(PROCESS.licence, lcmycountry, THIS.employee)

		IF NOT EMPTY(lcMsg)
			PROCESS.cerror =  lcMsg
		ELSE
			* report output
			lcPathName = THIS.companydatapath() + ADDBS("reports")
			IF NOT DIRECTORY(lcPathName)
				TRY
					MKDIR (lcPathName)
				CATCH
					THIS.cerror = "Reports folder was not created."
				ENDTRY
			ENDIF
			lcFileName = "Timesheet_Report"
			lcOutputFile = lcPathName + TRANSFORM(THIS.employee) + '_' + lcFileName + '.PDF'
			IF DIRECTORY(lcPathName)
				* generate PDF file
				reportsession =  xfrx("XFRX#INIT")
				returnvalue =  reportsession.setparams(lcOutputFile,  ,  .T.,  ,  ,  ,  "PDF")
				IF returnvalue =  0
					reportsession.processreport("timesheet.frx")
					reportsession.finalize()
				ELSE
					lcOutputFile = 	""
				ENDIF

				IF FILE(lcOutputFile)
					lcvalues = "Timesheet" + IIF(curMyPays.pay_type = 1, " Template", "") + " Report." + curMyPays.Reportvalues
					THIS.logreport(lcOutputFile, lcvalues)
					deletefiles(lcOutputFile, 300)
					*lcURL = STREXTRACT(UPPER(REQUEST.GetCurrentUrl(.F.)),This.CompanyHTTPS(),"/TIMESHEET")
					*lcURLpdf = This.CompanyHTTPS()+LOWER(ALLTRIM(lcURL))+"/GetPDF.si?date="+lcFileName
					*lcURLpdf = "http" +LOWER(STREXTRACT(UPPER(Request.GetCurrentUrl(.F.)),"HTTP","/TIMESHEET")) + "/GetPDF.si?date="+lcFileName
					*THIS.AddUserInfo("Report Was Created Successfully")
					Response.Downloadfile(lcOutputFile,"application/pdf","TimeSheetReport.pdf")
					RETURN
				ELSE
					THIS.cerror = "Report was not created."
				ENDIF
			ENDIF 
		ENDIF
	ENDIF

	tbReport = .F. && reset Report variable

	popage = THIS.newpageobject("timesheets:" + "ts_report", "time_report")
	response.expandscript(THIS.companyhtmlpath() + "master" + source_ext, SERVER.nscriptmode)

	ENDPROC
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	
	
	PROCEDURE TimesheetReportPDF
		This.TimesheetReportPage(.T.)
	ENDPROC 
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	PROCEDURE timesheetreportproc(tclicence, tcmycountry, tnstaff_code)
	* MY - 02/11/2012 - prepare timesheet report cursor
	
	LOCAL tntransid, tcapproved, tcselectgroup, tngroupid, lctype, lcstatus, tnreptype, tnrepstatus
	
	tntransid=-1
	tngroupid=-1

	tntransid = IIF (LEN(ALLTRIM(REQUEST.querystring("RepNameParamValue")))>0,EVALUATE(ALLTRIM(REQUEST.querystring("RepNameParamValue"))),-1)
	tcapproved = IIF (LEN(ALLTRIM(REQUEST.querystring("RepApproveParamValue")))>0,ALLTRIM(REQUEST.querystring("RepApproveParamValue")),"Both")
	tcselectgroup = IIF (LEN(ALLTRIM(REQUEST.querystring("RepGrOptParamValue")))>0,ALLTRIM(REQUEST.querystring("RepGrOptParamValue")),"Selected")
	tngroupid = IIF (LEN(ALLTRIM(REQUEST.querystring("GroupParamValue")))>0,EVALUATE(ALLTRIM(REQUEST.querystring("GroupParamValue"))),-1)
	tnreptype = IIF (LEN(ALLTRIM(REQUEST.querystring("RepTypeParamValue")))>0,ALLTRIM(REQUEST.querystring("RepTypeParamValue")),"Batch")
	tnrepstatus = IIF (LEN(ALLTRIM(REQUEST.querystring("RepStatusParamValue")))>0,ALLTRIM(REQUEST.querystring("RepStatusParamValue")),"Open")

	LOCAL lcapproved, lcgroup AS STRING
	PRIVATE lcvalues
	lcvalues = "Group: "

	DO CASE
	CASE UPPER(tcapproved) ==  "BOTH" OR UPPER(tcapproved) ==  "ALL"
		lcapproved =  " .T. "
	CASE UPPER(tcapproved) ==  "YES"
		lcvalues =  lcvalues +  "Approved Times Only. "
		lcapproved =  " timesheet.tsapproved "
	CASE UPPER(tcapproved) ==  "NO"
		lcapproved =  " NOT timesheet.tsapproved "
	OTHERWISE
		lcapproved =  " .T. "
	ENDCASE

	IF UPPER(tcselectgroup) == "ALL"
		lcgroup = " .T. "
	ELSE
		lcgroup = " tmmygroups = " +  TRANSFORM(tngroupid) +  " "
	ENDIF

	IF tntransid <= 0 && Handles 0 and -1 for default values
		lctype=""
		lcstatus=""

		DO CASE
		CASE UPPER(tnreptype) == "ALL"
			lctype = ".T."
		CASE UPPER(tnreptype) == "BATCH"
			lctype = "pay_type == 2"
		CASE UPPER(tnreptype) == "TEMPLATE"
			lctype = "pay_type == 1"
		OTHERWISE
			lctype = "pay_type == 2"
		ENDCASE
		DO CASE
		CASE UPPER(tnrepstatus) == "ALL"
			lcstatus = ".T."
		CASE UPPER(tnrepstatus) == "OPEN"
			lcstatus = "pay_status == 1"
		CASE UPPER(tnrepstatus) == "CLOSED"
			lcstatus = "pay_status == 2"
		OTHERWISE
			lcstatus = "pay_status == 1"
		ENDCASE

		SELECT *;
			FROM mypays;
			WHERE &lctype. AND &lcstatus.;
			INTO CURSOR curpays;
			ORDER BY pay_name
			
		SELECT curpays
		GO TOP
		IF NOT EOF()
			tntransid = curpays.pay_pk
		ENDIF

	ENDIF
	IF UPPER(tcselectgroup)= "ALL"

		SELECT mystaff.*,mygroups.grname  ;
			FROM mystaff  ;
			JOIN myteams ON mystaff.mywebcode = myteams.tmmystaff  ;
			JOIN mygroups ON myteams.tmmygroups = mygroups.grcode;
			INTO CURSOR curgroupedstaff
		lcvalues =  lcvalues +  "All Groups. "
	ELSE

		IF tngroupid = -1
			lcvalues = lcvalues +  "My Details Only. "
			SELECT mystaff.*,  SPACE(10)  AS  grname  ;
				FROM  mystaff  ;
				WHERE  mystaff.mywebcode =  tnstaff_code ;
				INTO  CURSOR  curgroupedstaff  READWRITE
		ELSE
			SELECT mystaff.*,mygroups.grname  ;
				FROM mystaff  ;
				JOIN myteams ON mystaff.mywebcode = myteams.tmmystaff  ;
				JOIN mygroups ON myteams.tmmygroups = mygroups.grcode;
				WHERE &lcgroup. ;
				INTO CURSOR curgroupedstaff
			lcvalues =  lcvalues +  ALLTRIM(curgroupedstaff.grname) +  ". "
		ENDIF

	ENDIF

	SELECT * FROM  curgroupedstaff ORDER BY grname INTO  CURSOR  curstaff
	USE IN  SELECT("curgroupedstaff")

	* Transaction types
	CREATE CURSOR  curtypes  (typecode c(1), typename c(50), leavename c(50))
	INDEX ON typecode TAG 	typecode
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("M", "Timesheet", "")
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("W", "Wages", "")
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("A",  "Allowances",  "")
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("S",  "Leave",  "Sick Leave")
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("O",  "Leave",  "Annual Leave")
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("F",  "Leave",  ALLTRIM(mystaff.myshname))
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("T",  "Leave",  ALLTRIM(mystaff.myotname))
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("N",  "Leave",  "Long Service Leave" )
	INSERT INTO curtypes (typecode, typename, leavename)  VALUES ("U",  "Leave",  "Unpaid Leave" )

	IF tcmycountry =  "Australia"
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("L",  "Leave", "Lieu Time")
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("R",  "Leave", "Rostered Day Off")
	ELSE
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("B",  "Leave", "Bereavement Leave")
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("P",  "Leave", "Public Holiday")
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("Y",  "Leave", "Alternative Leave Accrued")
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("Z",  "Leave", "Alternative Leave Paid")
		INSERT INTO curtypes (typecode, typename, leavename) VALUES ("R",  "Other", IIF(EMPTY(mystaff.myhpunits), "", ALLTRIM(mystaff.myhpunits) + " ") + "Paid for Relevant Daily Rate")
	ENDIF

	INSERT INTO curtypes (typecode, typename, leavename) VALUES ("D", "Other", IIF(EMPTY(mystaff.myhpunits), "", ALLTRIM(mystaff.myhpunits) + " ") + "Paid for Holiday Pay")

	* get mater record of batch or template
	SELECT *, lcvalues AS reportvalues, ;
		"Timesheet " + IIF(pay_type = 1, "Template", "Summary") + " Report" AS reporttitle, ;
		IIF(pay_type = 1, "Template Name: " + pay_name, "Date: " + DTOC(pay_date) + SPACE(20) + "Batch Name: " + pay_name) AS reportdetail, ;
		"Company ID: " + TRANSFORM(tclicence) AS licence ;
		FROM mypays ;
		WHERE pay_pk = tntransid ;
		INTO CURSOR curmypays

	DO CASE
	CASE curmypays.pay_type = 2
	* create batch cursor
		SELECT grname, ALLTRIM(mysurname)+", " + ALLTRIM(myname) AS myname, timesheet.tsdate, curtypes.typename, curtypes.leavename, ;
			NVL(ALLOW.NAME, SPACE(20)) AS allowname, NVL(wagetype.NAME, SPACE(50)) AS wagename,  ;
			TTOC(timesheet.tsstart,2) AS tsstart, ;
			curmypays.pay_name, curmypays.pay_status, curmypays.pay_type, curmypays.pay_date, ;
			timesheet.tstype, curstaff.mycountry,;
			TTOC(timesheet.tsfinish, 2) AS tsfinish, ;
			TTOC(timesheet.tsbreak, 2) AS tsbreak, timesheet.tsunits,;
			timesheet.tscostcent, timesheet.tscode, timesheet.tswagetype, timesheet.tsunits2,	;
			NVL(costcent.NAME,"Default Cost Centre ") AS costname, ;
			000000000.00 AS totaldays,	000000000.00 AS totalhours, myaltunits,myotunits, myshunits, mylslunits, myspunits,;
			myhpunits,  000000000.00 AS reducdays,	000000000.00 AS reduchours, ;
			timesheet.tsapproved, timesheet.tsdownload, EVL(timesheet.tsratecode, 1) AS tsratecode ;
			FROM timesheet  ;
			JOIN curtypes ON timesheet.tstype = curtypes.typecode  ;
			JOIN curmypays ON timesheet.tspay = curmypays.pay_pk ;
			JOIN curstaff ON timesheet.tsemp = curstaff.mywebcode;
			LEFT OUTER JOIN costcent ON timesheet.tscostcent = costcent.CODE ;
			LEFT OUTER JOIN ALLOW ON timesheet.tscode = ALLOW.CODE  ;
			LEFT OUTER JOIN wagetype ON timesheet.tswagetype = wagetype.CODE ;
			WHERE &lcapproved. ;
			ORDER BY 1,2,3,4,5,6,7,8 ;
			INTO CURSOR curreport READWRITE

	CASE curmypays.pay_type = 1
	* create template cursor
		SELECT grname, ALLTRIM(mysurname)+", " + ALLTRIM(myname) AS myname, timesheet.tsweek, timesheet.tsdaynbr, curtypes.typename, curtypes.leavename, ;
			NVL(ALLOW.NAME, SPACE(20)) AS allowname, ;
			NVL(wagetype.NAME, SPACE(50)) AS wagename,  ;
			TTOC(timesheet.tsstart,2) AS tsstart,  ;
			curmypays.pay_name, curmypays.pay_status, curmypays.pay_type, curmypays.pay_date, timesheet.tsdate, ;
			timesheet.tstype, curstaff.mycountry,;
			TTOC(timesheet.tsfinish, 2) AS tsfinish, ;
			TTOC(timesheet.tsbreak, 2) AS tsbreak, timesheet.tsunits,;
			timesheet.tscostcent, timesheet.tscode, timesheet.tswagetype, timesheet.tsunits2, ;
			NVL(costcent.NAME,"Default Cost Centre ") AS costname, ;
			000000000.00 AS totaldays,	;
			000000000.00 AS totalhours, ;
			myaltunits,myotunits, myshunits, mylslunits, myspunits,;
			myhpunits,  000000000.00 AS reducdays,	000000000.00 AS reduchours, ;
			timesheet.tsapproved, timesheet.tsdownload, EVL(timesheet.tsratecode, 1) AS tsratecode, timesheet.tsday ;
			FROM timesheet  ;
			JOIN curtypes ON timesheet.tstype = curtypes.typecode  ;
			JOIN curmypays ON timesheet.tmid = curmypays.pay_pk ;
			JOIN curstaff ON timesheet.tsemp = curstaff.mywebcode;
			LEFT OUTER JOIN costcent ON timesheet.tscostcent = costcent.CODE ;
			LEFT OUTER JOIN ALLOW ON timesheet.tscode = ALLOW.CODE  ;
			LEFT OUTER JOIN wagetype ON timesheet.tswagetype = wagetype.CODE ;
			ORDER BY 1,2,3,4,5,6,7,8,9 ;
			INTO CURSOR curreport READWRITE
	OTHERWISE
		RETURN "No template or batch was selected."
	ENDCASE

	SELECT curreport
	SCAN
		DO CASE
		CASE INLIST(tstype, "M", "W")
			REPLACE totalhours WITH totalhours + tsunits IN curreport
		CASE tstype = "A"
		CASE tstype = "S"
			IF myspunits = "Hours"
				REPLACE totalhours WITH totalhours + tsunits IN curreport
				REPLACE reduchours WITH reduchours + tsunits2 IN curreport
			ELSE
				REPLACE totalhours WITH totalhours + tsunits IN curreport
				REPLACE reducdays WITH reducdays + tsunits2 IN curreport
			ENDIF
		CASE tstype = "O"
			IF myhpunits = "Hours"
				REPLACE totalhours WITH totalhours + tsunits IN curreport
			ELSE
				REPLACE totaldays WITH totaldays + tsunits IN curreport
			ENDIF
		CASE tstype = "F"
			IF myshunits = "Hours"
				REPLACE totalhours WITH totalhours + tsunits IN curreport
			ELSE
				REPLACE totaldays WITH totaldays + tsunits IN curreport
			ENDIF
		CASE tstype = "T"
			IF myotunits = "Hours"
				REPLACE totalhours WITH totalhours + tsunits IN curreport
			ELSE
				REPLACE totaldays WITH totaldays + tsunits IN curreport
			ENDIF
		CASE tstype = "N"
			IF mylslunits = "Hours"
				REPLACE totalhours WITH totalhours + tsunits IN curreport
			ELSE
				REPLACE totaldays WITH totaldays + tsunits IN curreport
			ENDIF
		CASE mycountry = "Australia" .AND. tstype = "L"
		CASE mycountry = "Australia" .AND. tstype = "R"
		CASE mycountry <> "Australia" .AND. tstype = "R"
		CASE tstype = "D"
		CASE tstype = "B"
			REPLACE totalhours WITH totalhours + tsunits IN curreport
			REPLACE reducdays WITH reducdays + tsunits2 IN curreport
		CASE tstype = "U"
			REPLACE totalhours WITH totalhours + tsunits IN curreport
		CASE tstype = "P"
			REPLACE totalhours WITH totalhours + tsunits IN curreport
			REPLACE reducdays WITH reducdays + tsunits2 IN curreport
		CASE tstype = "Y"
			REPLACE totaldays WITH totaldays + tsunits IN curreport
		CASE tstype = "Z"
			REPLACE totalhours WITH totalhours + tsunits IN curreport
			REPLACE reducdays WITH reducdays + tsunits2 IN curreport
		ENDCASE
	ENDSCAN

	SELECT curreport
	GO TOP
 
    IF EOF()
		RETURN "Report was not created. No information found."
	ELSE
		*REPORT FORM TimeSheet.frx PREVIEW NOCONSOLE 
		RETURN ""
	ENDIF

	
	ENDPROC
	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	&&NOTE: this page handles posts to support the buildup of formRows...
	PROCEDURE TemplateEntryPage
		PRIVATE poPage, poStaff, poTypes, pnTemplateCount, pcAddWeek, pcAddDay, pnCurrentDay
		PRIVATE pnCurrentStaff, pnCurrentGroup, pnCurrentTemplate, poRetainList, plManager
		PRIVATE pnEditId, pcShowDownloaded, pdStartDate, pdEndDate, pcApproved, plAusie
		PRIVATE pcAddType, pnAddcount, pcOpen, pnPayCount, plPayOpen, pcFocusField, pcApproved
		PRIVATE poLeaveCodes, poOtherCodes, poAllowCodes, poWageCodes, poCostCentres, poJobCodes
		PRIVATE pcAddGroup, pcStartDay, pnStartDay, plUnitDisp
		
		LOCAL lcAction, loType, lnFieldCount,lnI, lnJ, lnCount, lnAt, ltStart, ltEnd, ltBreak, lnBreakLen
		LOCAL lnUnits, llEditOk, lcUnitDisp
		LOCAL ARRAY laFields[13], laDefaults[13]

		pnEditId = 0
		pnAddCount = 0
		pnCurrentDay = 0
		pcApproved = .f.
		pcStartDay = "Monday"
		pnStartDay = 0
		plPayOpen = .f.

		rcGroupInfo = "< My Details >"
		rcStaffInfo = "Unknown Manager"
		rcCompName = "Company Name"
		rcTemplate = "<My Template>"

		IF !(This.SelectData(This.Licence, "myStaff");
		  AND This.SelectData(This.Licence, "myPays");
		  AND This.SelectData(This.Licence, "myGroups");
		  AND This.SelectData(This.Licence, "wageType");
		  AND This.SelectData(This.Licence, "allow");
		  AND This.SelectData(This.Licence, "costCent");
		  AND This.SelectData(This.Licence, "timesheet"))	&&TODO: add new table for job-whateverItIs
			This.AddError("Page Setup Failed!")
		ELSE
			poStaff = Factory.GetStaffObject()

			IF !poStaff.Load(This.Employee)
				This.AddError("Failed to load Employee record: " + poStaff.cErrorMsg)
			ELSE
				poRetainList = Factory.GetRetainListObject()
			
				plManager = .F.
				pnCurrentGroup = 0
				pnCurrentStaff = 0
			 	pnCurrentTemplate = 0
				IF !This.SetupStaffGroupControlData(@plManager, @pnCurrentGroup, @pnCurrentStaff, poRetainList)
					This.AddValidationError("StaffGroupControl Setup Failed!")	&& non-fatal error
				ENDIF

				llEditOk = .t.
				lcAction = Request.Form("mode")
				pnEditId = VAL(Request.QueryString("edit"))

				IF VARTYPE(lcAction) == 'C'
					IF lcAction = "refilter" AND pnEditId > 0
						IF !This.CheckRights("TEM_GROUP_M") AND !This.CheckRights("TEM_EMPLOYEE_M")
							llEditOk = .f.
						ENDIF

						IF llEditOk
							IF !This.CheckAccess(pnCurrentStaff, This.IsManager(This.Employee), .T.)	&& Allow Everyone option
								llEditOk = .f.
							ENDIF
						ENDIF
				
						IF llEditOk
							IF pnCurrentTemplate == MY_TEMPLATE
								IF This.Employee <> pnCurrentStaff
									llEditOk = .f.
								ENDIF
							ENDIF
						ENDIF
					ENDIF
				ENDIF

				IF !llEditOk
					This.AddError("You do not have access to edit that entry!")
					pnEditId = 0
				ELSE
				 	pnCurrentTemplate = 0
					IF !This.SetupTemplateControlData(@pnCurrentTemplate, poRetainList)
						This.AddValidationError("TemplateControl Setup Failed!")	&& non-fatal error
				 	ENDIF

					pcFocusField = ""
					pcAddType = ""
					pcAddWeek = ""
					pcAddDay = ""

					IF VARTYPE(pcAddGroup) <> "C"
						pcAddGroup = "Employee"
					ENDIF

					plAusie = This.IsAustralia()
					poLeaveCodes = This.GetLeaveCodes(IIF(pnCurrentStaff == EVERYONE_OPTION, This.Employee, pnCurrentStaff),.T.)
					poOtherCodes = This.GetOtherCodes()
					poAllowCodes = This.GetAllowanceCodes()
					poWageCodes = This.GetWageCodes()
					poCostCentres = This.GetCostCentres()
					&&LATER: poJobCodes = This.GetJobCodes()

					pnEditId = VAL(Request.QueryString("edit"))
					lcUnitDisp = AppSettings.Get("unitdisp")
					plUnitDisp = .F.
					IF ALLTRIM(lcUnitDisp) = "Decimal"
						plUnitDisp = .T.
				 	ENDIF

					pnTemplateCount = 0
					SELECT mypays
					COUNT TO pnTemplateCount FOR mypays.pay_type = 1

					pcAddType = ""
					pcShowDownloaded = "no"

					pnCurrentPay = 0
					poTypes = This.GetTemplateTypes(.T., pnCurrentTemplate, pnCurrentGroup, pnCurrentStaff,"")
	
					* 02/11/2009;TTP4688;JCF: Added handling for pcFocusField so that the first editable field in the edited line gets focus on page load.
					IF !EMPTY(pnEditId)
						FOR lnI = 1 TO poTypes.Count
							loType = poTypes.Item(lnI)
							SELECT (loType.cursorName)
							LOCATE FOR tsId == pnEditId
							IF FOUND()
								EXIT
							ENDIF
						NEXT

						IF loType.showWeek
							IF EMPTY(pcFocusField) AND !loType.readOnlyWeek
								pcFocusField = "week"
							ENDIF
						ENDIF
						IF loType.showDay
							IF EMPTY(pcFocusField) AND !loType.readOnlyDay
								pcFocusField = "day"
							ENDIF
						ENDIF
						IF loType.showLeaveType
							IF EMPTY(pcFocusField) AND !loType.readOnlyLeaveType
								pcFocusField = "leaveType"
							ENDIF
						ENDIF
						IF loType.showOtherType
							IF EMPTY(pcFocusField) AND !loType.readOnlyOtherType
								pcFocusField = "otherType"
							ENDIF
						ENDIF
						IF loType.showCode
							IF EMPTY(pcFocusField) AND !loType.readOnlyCode
								pcFocusField = "code"
							ENDIF
						ENDIF
						IF loType.showStart
							IF EMPTY(pcFocusField) AND !loType.readOnlyStart
								pcFocusField = "start"
							ENDIF
						ENDIF
						IF loType.showEnd
							IF EMPTY(pcFocusField) AND !loType.readOnlyEnd
								pcFocusField = "end"
							ENDIF
						ENDIF
						IF loType.showBreak
							IF EMPTY(pcFocusField) AND !loType.readOnlyBreak
								pcFocusField = "break"
							ENDIF
						ENDIF
						IF loType.showUnits
							IF EMPTY(pcFocusField) AND !loType.readOnlyUnits
								pcFocusField = "units"
							ENDIF
						ENDIF
						IF loType.showReduce
							IF EMPTY(pcFocusField) AND !loType.readOnlyReduce
								pcFocusField = "units2"
							ENDIF
						ENDIF
						IF loType.showWageType
							IF EMPTY(pcFocusField) AND !loType.readOnlyWageType
								pcFocusField = "wageType"
							ENDIF
						ENDIF
						IF loType.showRateCode
							IF EMPTY(pcFocusField) AND !loType.readOnlyRateCode
								pcFocusField = "rateCode"
							ENDIF
						ENDIF
						IF loType.showCostCent
							IF EMPTY(pcFocusField) AND !loType.readOnlyCostCent
								pcFocusField = "costCentre"
							ENDIF
						ENDIF
						IF loType.showJobCode
							IF EMPTY(pcFocusField) AND !loType.readOnlyJobCode
								pcFocusField = "jobCode"
							ENDIF
						ENDIF
					ELSE
						pcAddWeek = Request.Form("addWeek")
						pcAddDay  = Request.Form("addDay")
						pcAddType = Request.Form("addType")
						pcAddGroup = Request.Form("addgroup")
						pnAddcount = VAL(Request.Form("count"))
						pnCurrentDay = VAL(Request.Form("currentday"))
						pcStartDay = Request.Form("startDay")
						IF EMPTY(pcStartDay)
							pcStartDay = "Monday"
						ENDIF

						pnStartDay = this.getdaynumber(pcStartDay)
						pnStartDay = pnStartDay - 1
						IF pnStartDay < 0
							pnStartDay = 0
						ENDIF
						
						IF EMPTY(poTypes.GetKey(pcAddType))
							pcAddType = ""
							pnAddcount = 0
						ELSE
							loType = poTypes.Item(pcAddType)
							lnFieldCount = 0
							poAddValues = CREATEOBJECT("COLLECTION")
							* Don't soak up current values if changing type.
							IF !(LEFT(lcAction, 7) == "newType")
								FOR lnI = 1 TO pnAddcount
									IF loType.showWeek
										poAddValues.Add(Request.Form("week_" + TRANSFORM(lnI)), "week_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showDay
										poAddValues.Add(Request.Form("day_" + TRANSFORM(lnI)), "day_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showLeaveType
										poAddValues.Add(Request.Form("leaveType_" + TRANSFORM(lnI)), "leaveType_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showOtherType
										poAddValues.Add(Request.Form("otherType_" + TRANSFORM(lnI)), "otherType_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showCode
										poAddValues.Add(Request.Form("code_" + TRANSFORM(lnI)), "code_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showStart
										poAddValues.Add(Request.Form("start_" + TRANSFORM(lnI)), "start_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showEnd
										poAddValues.Add(Request.Form("end_" + TRANSFORM(lnI)), "end_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showBreak
										poAddValues.Add(Request.Form("break_" + TRANSFORM(lnI)), "break_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showUnits
										IF loType.showStart AND loType.showEnd AND loType.showBreak
											* 16/10/2009;TTP:4643;JCF: Timesheet type needs the units recalculated as the [shouldn't] be passed back by the browser since the field is disabled.
											ltStart	= CTOT(poAddValues.Item("start_" + TRANSFORM(lnI)))
											ltEnd	= CTOT(poAddValues.Item("end_" + TRANSFORM(lnI)))
											ltBreak	= CTOT(poAddValues.Item("break_" + TRANSFORM(lnI)))
											lnBreakLen = ltBreak - CTOT("00:00")
												IF ltEnd < ltStart
													* Crossed midnight
													ltEnd = ltEnd + 86400
												ENDIF
												lnUnits = ((ltEnd - ltStart) - lnBreakLen) / 3600
												IF lnUnits < 0 OR lnUnits > 24
													* This probably never fires unless they disable JavaScript...
													This.AddValidationError("Invalid Start/End/Break combination.")
												ENDIF
											poAddValues.Add(TRANSFORM(lnUnits), "units_" + TRANSFORM(lnI))
										ELSE
											poAddValues.Add(Request.Form("units_" + TRANSFORM(lnI)), "units_" + TRANSFORM(lnI))
										ENDIF
									ENDIF
									IF loType.showReduce
										poAddValues.Add(Request.Form("units2_" + TRANSFORM(lnI)), "units2_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showWageType
										poAddValues.Add(Request.Form("wageType_" + TRANSFORM(lnI)), "wageType_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showRateCode
										poAddValues.Add(Request.Form("rateCode_" + TRANSFORM(lnI)), "rateCode_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showCostCent
										poAddValues.Add(Request.Form("costCentre_" + TRANSFORM(lnI)), "costCentre_" + TRANSFORM(lnI))
									ENDIF
									IF loType.showJobCode
										poAddValues.Add(Request.Form("jobCode_" + TRANSFORM(lnI)), "jobCode_" + TRANSFORM(lnI))
									ENDIF
								NEXT
							ENDIF
	
							* 02/11/2009;TTP4688;JCF: Added handling for pcFocusField so that the first editable field in the "current" line gets focus on page load.
							IF loType.showWeek
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "week_"
								laDefaults[lnFieldCount] = "1"
								IF EMPTY(pcFocusField) AND !loType.readOnlyWeek
									pcFocusField = "week_"
								ENDIF
							ENDIF
							IF loType.showDay
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "day_"
								laDefaults[lnFieldCount] = "Monday"
								pnCurrentDay = 1
								IF EMPTY(pcFocusField) AND !loType.readOnlyDay
									pcFocusField = "day_"
								ENDIF
							ENDIF
							IF loType.showLeaveType
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "leaveType_"
								laDefaults[lnFieldCount] = ""
									IF EMPTY(pcFocusField) AND !loType.readOnlyLeaveType
									pcFocusField = "leaveType_"
								ENDIF
							ENDIF
							IF loType.showOtherType
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "otherType_"
								laDefaults[lnFieldCount] = ""
									IF EMPTY(pcFocusField) AND !loType.readOnlyOtherType
									pcFocusField = "otherType_"
								ENDIF
							ENDIF
							IF loType.showCode
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "code_"
								laDefaults[lnFieldCount] = '0'
									IF EMPTY(pcFocusField) AND !loType.readOnlyCode
									pcFocusField = "code_"
								ENDIF
							ENDIF
							IF loType.showStart
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "start_"
								laDefaults[lnFieldCount] = "00:00"
								IF EMPTY(pcFocusField) AND !loType.readOnlyStart
									pcFocusField = "start_"
								ENDIF
							ENDIF
							IF loType.showEnd
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "end_"
								laDefaults[lnFieldCount] = "00:00"
								IF EMPTY(pcFocusField) AND !loType.readOnlyEnd
									pcFocusField = "end_"
								ENDIF
							ENDIF
							IF loType.showBreak
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "break_"
								laDefaults[lnFieldCount] = "00:00"
								IF EMPTY(pcFocusField) AND !loType.readOnlyBreak
									pcFocusField = "break_"
								ENDIF
							ENDIF
							IF loType.showUnits
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "units_"
								laDefaults[lnFieldCount] = '0'
								IF EMPTY(pcFocusField) AND !loType.readOnlyUnits
									pcFocusField = "units_"
								ENDIF
							ENDIF
							IF loType.showReduce
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "units2_"
								laDefaults[lnFieldCount] = '0'
								IF EMPTY(pcFocusField) AND !loType.readOnlyReduce
									pcFocusField = "units2_"
								ENDIF
							ENDIF
							IF loType.showWageType
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "wageType_"
								laDefaults[lnFieldCount] = '0'
								IF EMPTY(pcFocusField) AND !loType.readOnlyWageType
									pcFocusField = "wageType_"
								ENDIF
							ENDIF
							IF loType.showRateCode
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "rateCode_"
								laDefaults[lnFieldCount] = '1'
								IF EMPTY(pcFocusField) AND !loType.readOnlyRateCode
									pcFocusField = "rateCode_"
								ENDIF
							ENDIF
							IF loType.showCostCent
									lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "costCentre_"
								laDefaults[lnFieldCount] = '0'
								IF EMPTY(pcFocusField) AND !loType.readOnlyCostCent
									pcFocusField = "costCentre_"
								ENDIF
							ENDIF
							IF loType.showJobCode
								lnFieldCount = lnFieldCount + 1
								laFields[lnFieldCount] = "jobCode_"
								laDefaults[lnFieldCount] = '0'
								IF EMPTY(pcFocusField) AND !loType.readOnlyJobCode
									pcFocusField = "jobCode_"
								ENDIF
							ENDIF
							* if asked to, replicate a row in the temp table with consecutive dates, or add more rows to it, or explode an Everyone row etc...
							* 02/11/2009;TTP4688;JCF: Added logic to define what the "current" line means in each case below for pcFocusField.
							DO CASE
								CASE LEFT(lcAction, 7) == "newType"		&& may be newType or newType_<count>
									* add N rows of the new type
									IF '_' $ lcAction
										lnCount = VAL(SUBSTR(lcAction, 9))
									ELSE
										pnAddcount = 0
										lnCount = 1
									ENDIF
									FOR lnI = pnAddcount + 1 TO pnAddcount + lnCount
										FOR lnJ = 1 TO lnFieldCount
											poAddValues.Add(laDefaults[lnJ], laFields[lnJ] + TRANSFORM(lnI))
										NEXT
									NEXT
									* Focus on the first new row
									pcFocusField = pcFocusField + TRANSFORM(pnAddcount + 1)
	
									pnAddcount = pnAddcount + lnCount
								
								CASE LEFT(lcAction, 4) == "add_"
									* add N rows of the current type

									lnCount = VAL(SUBSTR(lcAction, 5))
									pnCurrentDay = 0
									IF pnAddCount > 0
										lnCount = ASCAN(laFields, "day_")
										IF lnCount > 0
											lcTemp = poAddValues.Item("day_" + TRANSFORM(pnAddCount))
											pnCurrentDay = this.getdaynumber(lcTemp)
										ENDIF
									ENDIF

									lnCount = VAL(SUBSTR(lcAction, 5))
									pnCurrentDay = pnStartDay
									FOR lnI = pnAddcount + 1 TO pnAddcount + lnCount
										FOR lnJ = 1 TO lnFieldCount
											IF laFields[lnJ] == "day_"
												pnCurrentDay = pnCurrentDay + 1
												IF pnCurrentDay > 7
													pnCurrentDay = 1
												ENDIF
												lcTemp = this.getdayname(pnCurrentDay)
												poAddValues.Add(lcTemp, laFields[lnJ] + TRANSFORM(lnI))
											ELSE
												poAddValues.Add(laDefaults[lnJ], laFields[lnJ] + TRANSFORM(lnI))
											ENDIF
										NEXT
									NEXT
	
									* Focus on first new row
									pcFocusField = pcFocusField + TRANSFORM(pnAddcount + 1)
	
									pnAddcount = pnAddcount + lnCount
	
								CASE LEFT(lcAction, 7) == "delete_"
									* delete the Nth row; bubble up other values
									lnCount = VAL(SUBSTR(lcAction, 8))
									pnAddcount = pnAddcount - 1
									FOR lnI = lnCount TO pnAddcount
										FOR lnJ = 1 TO lnFieldCount
											poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
											poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnI + 1)), laFields[lnJ] + TRANSFORM(lnI))
										NEXT
									NEXT
	
									* Focus on either the row below the one deleted if there is one, or the last row
									pcFocusField = pcFocusField + TRANSFORM(MIN(pnAddcount, lnCount))
	
								CASE LEFT(lcAction, 5) == "copy_"		&& copy_<at>_<count>
									* copy the Nth row, inserting at N+1 and pushing down.
									lnAt = VAL(SUBSTR(lcAction, 6))
									lnCount = VAL(SUBSTR(lcAction, 8 + FLOOR(LOG10(lnAt))))
	
									FOR lnI = pnAddcount + lnCount TO lnAt + lnCount + 1 STEP -1
										FOR lnJ = 1 TO lnFieldCount
											IF lnI <= pnAddcount
												poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
											ENDIF
	
											poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnI - lnCount)), laFields[lnJ] + TRANSFORM(lnI))
										NEXT
									NEXT

									FOR lnI = lnAt + 1 TO lnAt + lnCount
										FOR lnJ = 1 TO lnFieldCount
											IF lnI <= pnAddcount
												poAddValues.Remove(laFields[lnJ] + TRANSFORM(lnI))
											ENDIF
	
											poAddValues.Add(poAddValues.Item(laFields[lnJ] + TRANSFORM(lnAt)), laFields[lnJ] + TRANSFORM(lnI))
										NEXT
									NEXT

									* Focus on the first new row
									pcFocusField = pcFocusField + TRANSFORM(lnAt + 1)

									pnAddcount = pnAddcount + lnCount

								*!* 09/11/2009;TTP4555;JCF: handle the case when the user has pressed enter and submits the form with no specific action.
								CASE lcAction == ""	&& Happens if the AddEntries form is submitted by the user pressing the Enter key...need to fix the focused field in this case.
									pcFocusField = pcFocusField + '1'
							ENDCASE
						ENDIF
					
						rcGroupInfo = "< My Details >"
						rcStaffInfo = "Unknown Manager"
						rcCompName = "Company Name"
						rcTemplate = "<My Template>"

						IF pnCurrentTemplate > 0
							IF SEEK(pnCurrentTemplate,"myPays","pay_pk")
								rcTemplate = myPays.pay_name
							ENDIF
						ENDIF

						IF pnCurrentGroup > 0
							IF SEEK(pnCurrentGroup,"myGroups","grcode")
								rcGroupInfo = myGroups.grname
							ENDIF
						ENDIF

						IF pnCurrentStaff > 0
							IF SEEK(pnCurrentStaff,"myStaff","mywebcode")
								rcStaffInfo = ALLTRIM(myStaff.mySurname)+","+ALLTRIM(myStaff.myname)
							ENDIF
						ENDIF

					ENDIF
				ENDIF
			ENDIF
		ENDIF

*		WAIT WINDOW IIF(VARTYPE(lcAction)<>"C","lcAction = None","lcAction = "+lcAction)+CHR(13)+;
		            " pnCurrentTemplate = "+ALLTRIM(STR(pnCurrentTemplate,10,0))+CHR(13)+;
		            " pnCurrentGroup = "+ALLTRIM(STR(pnCurrentGroup,10,0))+CHR(13)+;
		            " pnCurrentStaff = "+ALLTRIM(STR(pnCurrentStaff,10,0))+CHR(13)+;
 					" this.Employee = "+IIF(VARTYPE(this.Employee)="N",ALLTRIM(STR(this.employee,10,0)),"**")+CHR(13)+;
		            " Manager ["+IIF(plManager,"Yes","No")+"]"+CHR(13)+;
		            " pnAddCount = "+ALLTRIM(STR(pnAddCount,10,0))+CHR(13)+;
		            IIF(VARTYPE(pcAddGroup)<>"C","pcAddGroup = none","pcAddGroup = "+pcaddGroup)+CHR(13)+;
		            " pnCurrentDay = "+ALLTRIM(STR(pnCurrentDay,10,0))+CHR(13)+;
		            " pnEditId = "+ALLTRIM(STR(pnEditID,10,0))+' '+CHR(13)+;
		            " pcStartDay = "+pcStartDay+chr(13)+;
		            TIME() ;
		             NOWAIT noclear
*		             
		poPage = This.NewPageObject("timesheets:t_entry", "template_entry")

		Response.ExpandScript(This.CompanyHtmlPath() + "master" + SOURCE_EXT, Server.nScriptMode)
	ENDPROC

	*--------------------------------------------------------------------------------*
	PROCEDURE CollectTemplateEntryFormData(toType, tcSuffix, tnCurrentStaff,;
											rnWeek, rcDay, roLeaveCode,;
											roOtherCode, roAllowCode, rtStart, rtEnd, rtBreak,;
											rnUnits, rnReduce, roWageCode, rnRateCode,;
											roCostCentCode, roJobCode)
											
		LOCAL loCodes, lcValue

		IF !EMPTY(tcSuffix)
			tcSuffix = '_' + TRANSFORM(tcSuffix)
		ELSE
			tcSuffix = ""
		ENDIF

		IF toType.showWeek AND !toType.readOnlyWeek
			rnWeek = VAL(Request.Form("week" + tcSuffix))
		ENDIF

		IF toType.showDay AND !toType.readOnlyDay
			rcDay = Request.Form("day" + tcSuffix)
		ENDIF

		IF toType.showLeaveType AND !toType.readOnlyLeaveType
			lcValue = Request.Form("leaveType" + tcSuffix)
			loCodes = This.GetLeaveCodes(-1,.T.)		&& Turn Off authz
			IF !EMPTY(loCodes.GetKey(lcValue))
				roLeaveCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showCode AND !toType.readOnlyCode
			lcValue = Request.Form("code" + tcSuffix)
			loCodes = This.GetAllowanceCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roAllowCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showOtherType AND !toType.readOnlyOtherType
			lcValue = Request.Form("otherType" + tcSuffix)
			loCodes = This.GetOtherCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roOtherCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showStart AND !toType.readOnlyStart
			rtStart = CTOT(Request.Form("start" + tcSuffix))
		ENDIF

		IF toType.showEnd AND !toType.readOnlyEnd
			rtEnd = CTOT(Request.Form("end" + tcSuffix))
		ENDIF
		IF toType.showBreak AND !toType.readOnlyBreak
			rtBreak = CTOT(Request.Form("break" + tcSuffix))
		ENDIF

		IF toType.showUnits AND !toType.readOnlyUnits
			* handle the override values on allowances where applicable as they will [should] not have been posted
			IF toType.id == "allowances" AND !ISNULL(roAllowCode) AND !EMPTY(roAllowCode.unitsValue)
				rnUnits = roAllowCode.unitsValue
			ELSE
				rnUnits = VAL(Request.Form("units" + tcSuffix))
			ENDIF
		ENDIF

		IF toType.showReduce AND !toType.readOnlyReduce
			IF !toType.showLeaveType OR roLeaveCode.enableReduce
				rnReduce = VAL(Request.Form("units2" + tcSuffix))
			ENDIF
		ENDIF

		IF toType.showWageType AND !toType.readOnlyWageType
			lcValue = Request.Form("wageType" + tcSuffix)
			loCodes = This.GetWageCodes()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roWageCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

		IF toType.showRateCode AND !toType.readOnlyRateCode
	
			rnRateCode = VAL(Request.Form("rateCode" + tcSuffix))
		ENDIF

		IF toType.showCostCent AND !toType.readOnlyCostCent
			* handle the override values on allowances where applicable as they will [should] not have been posted
			IF toType.id == "allowances" AND !ISNULL(roAllowCode) AND !EMPTY(roAllowCode.costCentre)
				lcValue = TRANSFORM(roAllowCode.costCentre)
			ELSE
				lcValue = Request.Form("costCentre" + tcSuffix)
			ENDIF
			loCodes = This.GetCostCentres()		&& This handles the authz for us.
			IF !EMPTY(loCodes.GetKey(lcValue))
				roCostCentCode = loCodes.Item(lcValue)
			ENDIF
		ENDIF

*		IF toType.showJobCode AND !toType.readOnlyJobCode
*			lcValue = Request.Form("jobCode" + tcSuffix)
*			loCodes = This.GetJobCodes()		&& This handles the authz for us.
*			IF !EMPTY(loCodes.GetKey(lcValue))
*				roJobCode = loCodes.Item(lcValue)
*			ENDIF
*		ENDIF
		
		RETURN
		
	ENDPROC

	*--------------------------------------------------------------------------------*
	FUNCTION SaveSingleTemplateEntry(toType, tnCurrentTemplate, tnStaff, tnWeek, tcDay, toLeaveCode,;
										toOtherCode, toAllowCode, ttStart, ttEnd, ttBreak,;
										tnUnits, tnReduce, toWageCode, tnRateCode,;
										toCostCentCode, toJobCode) AS Boolean

		LOCAL lnBreakLen, llError, loGroup, lnMaxUnitsValue

		*!* 16/11/2009;TTP4791;JCF: Added the no-limit effect to other as well, and centralised the check.
		*!* 19/11/2009;TTP4852,4854;JCF: altered max "uncapped" limit to avoid a rounding issue that caused an out of band value and hence a NumericOverflow error on import.
		lnMaxUnitsValue = IIF(INLIST(toType.id, "wages", "allowances", "other"), 999998.99, 24)

		llError = .F.

		* Validate...

		IF toType.showWeek AND !toType.readOnlyWeek
			IF EMPTY(tnWeek)
				This.AddValidationError("Missing or invalid Week.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showDay AND !toType.readOnlyDay
			IF EMPTY(tcDay)
				This.AddValidationError("Missing or invalid Day.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showLeaveType AND !toType.readOnlyLeaveType
			IF ISNULL(toLeaveCode)
				This.AddValidationError("Missing or invalid LeaveType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showOtherType AND !toType.readOnlyOtherType
			IF ISNULL(toOtherCode)
				This.AddValidationError("Missing or invalid OtherType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showCode AND !toType.readOnlyCode
			IF ISNULL(toAllowCode)
				This.AddValidationError("Missing or invalid Code.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showStart AND !toType.readOnlyStart
			IF EMPTY(ttStart)
				This.AddValidationError("Missing or invalid Start.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showEnd AND !toType.readOnlyEnd
			IF EMPTY(ttEnd)
				This.AddValidationError("Missing or invalid End.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showBreak AND !toType.readOnlyBreak
			IF EMPTY(ttBreak)
				This.AddValidationError("Missing or invalid Break.")
				llError = .T.
			ENDIF
		ENDIF

		IF !(EMPTY(ttStart) OR EMPTY(ttEnd) OR EMPTY(ttBreak))
			lnBreakLen = ttBreak - CTOT("00:00")

			IF ttEnd < ttStart
				* Crossed midnight
				ttEnd = ttEnd + 86400
			ENDIF

			tnUnits = ((ttEnd - ttStart) - lnBreakLen) / 3600
			IF tnUnits < 0 OR tnUnits > 24
				This.AddValidationError("Invalid Start/End/Break combination.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showUnits AND !toType.readOnlyUnits
			* Only validate if the field was enabled - in this case, for all types other than allowances, or for allowances where the selected code has the field enabled.
			IF !toType.showCode OR toAllowCode.enableUnits
				IF tnUnits < -lnMaxUnitsValue OR tnUnits > lnMaxUnitsValue	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					*!* 16/11/2009;TTP4791;JCF: Added the missing TRANSFORM() so the following line actually works; Centralised the logic in the same way as the template.
					This.AddValidationError("Units out of range " + TRANSFORM(-lnMaxUnitsValue) + " to " + TRANSFORM(lnMaxUnitsValue) + '.')
					llError = .T.
				ENDIF
			ENDIF
		ENDIF

		IF toType.showReduce AND !toType.readOnlyReduce
			* Only validate if the field was enabled - in this case, for all types other than leave, or for leave where the selected code has the field enabled.
			IF !toType.showLeaveType OR toLeaveCode.enableReduce
				IF tnReduce < -24 OR tnReduce > 24	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					This.AddValidationError("Reduce out of range -24 to 24.")
					llError = .T.
				ENDIF
			ELSE
				tnReduce = 0
			ENDIF
		ENDIF

		IF toType.showWageType AND !toType.readOnlyWageType
			IF ISNULL(toWageCode)
				This.AddValidationError("Missing or invalid WageType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showRateCode AND !toType.readOnlyRateCode
			IF tnRateCode < 1 OR tnRateCode > 9
				This.AddValidationError("Invalid RateCode (not between 1 and 9 inclusive).")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showCostCent AND !toType.readOnlyCostCent
			IF ISNULL(toCostCentCode)
				This.AddValidationError("Missing or invalid CostCentre.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showJobCode AND !toType.readOnlyJobCode
			IF ISNULL(toJobCode)
				This.AddValidationError("Missing or invalid JobCode.")
				llError = .T.
			ENDIF
		ENDIF
		
		IF !llError		
			SELECT TimeSheet

			REPLACE tsType   WITH toType.tsType
			REPLACE tmId     WITH tnCurrentTemplate
			REPLACE tsPay    WITH 0
			REPLACE tsDate   WITH DATE()
			REPLACE tsEmp    WITH tnStaff
			REPLACE tsDayNbr WITH this.getdaynumber(tcDay)
	
			IF toType.showWeek AND !toType.readOnlyWeek
				REPLACE tsWeek WITH tnWeek
			ENDIF
			IF toType.showDay AND !toType.readOnlyDay
				REPLACE tsDay WITH tcDay
			ENDIF
			IF toType.showLeaveType AND !toType.readOnlyLeaveType
				REPLACE tsType WITH toLeaveCode.code
			ENDIF
			IF toType.showOtherType AND !toType.readOnlyOtherType
				REPLACE tsType WITH toOtherCode.code
			ENDIF
			IF toType.showCode AND !toType.readOnlyCode
				REPLACE tsCode WITH VAL(toAllowCode.code)
			ENDIF
			IF toType.showStart AND !toType.readOnlyStart
				REPLACE	tsStart WITH ttStart, tsFinish WITH ttEnd, tsBreak WITH ttBreak, tsUnits WITH tnUnits
			ENDIF
			IF toType.showUnits AND !toType.readOnlyUnits
				REPLACE tsUnits WITH tnUnits
			ENDIF
			IF toType.showReduce AND !toType.readOnlyReduce
				REPLACE tsUnits2 WITH tnReduce
			ENDIF
			IF toType.showWageType AND !toType.readOnlyWageType
				REPLACE tsWageType WITH VAL(toWageCode.code)
			ENDIF
			IF toType.showRateCode AND !toType.readOnlyRateCode
				REPLACE tsRateCode WITH tnRateCode
			ENDIF
			IF toType.showCostCent AND !toType.readOnlyCostCent
				REPLACE tsCostCent WITH VAL(toCostCentCode.code)
			ENDIF
*			IF toType.showJobCode AND !toType.readOnlyJobCode
*				REPLACE tsXXXX WITH VAL(toJobCode.code)	&&LATER: Add correct dbFieldName...
*			ENDIF
		ENDIF
		
		RETURN !llError
	ENDFUNC

	*--------------------------------------------------------------------------------*
	FUNCTION TempSingleTemplateEntry(toType, tnCurrentTemplate, tnStaff, tnWeek, tcDay, toLeaveCode,;
										toOtherCode, toAllowCode, ttStart, ttEnd, ttBreak,;
										tnUnits, tnReduce, toWageCode, tnRateCode,;
										toCostCentCode, toJobCode,;
										tnCurrentGroup, rnRowCount) AS Boolean

		LOCAL lnBreakLen, llError, loGroup, lnMaxUnitsValue

		*!* 16/11/2009;TTP4791;JCF: Added the no-limit effect to other as well, and centralised the check.
		*!* 19/11/2009;TTP4852,4854;JCF: altered max "uncapped" limit to avoid a rounding issue that caused an out of band value and hence a NumericOverflow error on import.
		lnMaxUnitsValue = IIF(INLIST(toType.id, "wages", "allowances", "other"), 999998.99, 24)

		rnRowCount = 1

		llError = .F.

		* Validate...

		IF toType.showWeek AND !toType.readOnlyWeek
			IF EMPTY(tnWeek)
				This.AddValidationError("Missing or invalid Week.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showDay AND !toType.readOnlyDay
			IF EMPTY(tcDay)
				This.AddValidationError("Missing or invalid Day.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showLeaveType AND !toType.readOnlyLeaveType
			IF ISNULL(toLeaveCode)
				This.AddValidationError("Missing or invalid LeaveType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showOtherType AND !toType.readOnlyOtherType
			IF ISNULL(toOtherCode)
				This.AddValidationError("Missing or invalid OtherType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showCode AND !toType.readOnlyCode
			IF ISNULL(toAllowCode)
				This.AddValidationError("Missing or invalid Code.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showStart AND !toType.readOnlyStart
			IF EMPTY(ttStart)
				This.AddValidationError("Missing or invalid Start.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showEnd AND !toType.readOnlyEnd
			IF EMPTY(ttEnd)
				This.AddValidationError("Missing or invalid End.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showBreak AND !toType.readOnlyBreak
			IF EMPTY(ttBreak)
				This.AddValidationError("Missing or invalid Break.")
				llError = .T.
			ENDIF
		ENDIF

		IF !(EMPTY(ttStart) OR EMPTY(ttEnd) OR EMPTY(ttBreak))
			lnBreakLen = ttBreak - CTOT("00:00")

			IF ttEnd < ttStart
				* Crossed midnight
				ttEnd = ttEnd + 86400
			ENDIF

			tnUnits = ((ttEnd - ttStart) - lnBreakLen) / 3600
			IF tnUnits < 0 OR tnUnits > 24
				This.AddValidationError("Invalid Start/End/Break combination.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showUnits AND !toType.readOnlyUnits
			* Only validate if the field was enabled - in this case, for all types other than allowances, or for allowances where the selected code has the field enabled.
			IF !toType.showCode OR toAllowCode.enableUnits
				IF tnUnits < -lnMaxUnitsValue OR tnUnits > lnMaxUnitsValue	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					*!* 16/11/2009;TTP4791;JCF: Added the missing TRANSFORM() so the following line actually works; Centralised the logic in the same way as the template.
					This.AddValidationError("Units out of range " + TRANSFORM(-lnMaxUnitsValue) + " to " + TRANSFORM(lnMaxUnitsValue) + '.')
					llError = .T.
				ENDIF
			ENDIF
		ENDIF

		IF toType.showReduce AND !toType.readOnlyReduce
			* Only validate if the field was enabled - in this case, for all types other than leave, or for leave where the selected code has the field enabled.
			IF !toType.showLeaveType OR toLeaveCode.enableReduce
				IF tnReduce < -24 OR tnReduce > 24	&& 19/11/2009;4810;JCF: negative limit same as -max now rather than 0.
					This.AddValidationError("Reduce out of range -24 to 24.")
					llError = .T.
				ENDIF
			ELSE
				tnReduce = 0
			ENDIF
		ENDIF

		IF toType.showWageType AND !toType.readOnlyWageType
			IF ISNULL(toWageCode)
				This.AddValidationError("Missing or invalid WageType.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showRateCode AND !toType.readOnlyRateCode
			IF tnRateCode < 1 OR tnRateCode > 9
				This.AddValidationError("Invalid RateCode (not between 1 and 9 inclusive).")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showCostCent AND !toType.readOnlyCostCent
			IF ISNULL(toCostCentCode)
				This.AddValidationError("Missing or invalid CostCentre.")
				llError = .T.
			ENDIF
		ENDIF

		IF toType.showJobCode AND !toType.readOnlyJobCode
			IF ISNULL(toJobCode)
				This.AddValidationError("Missing or invalid JobCode.")
				llError = .T.
			ENDIF
		ENDIF
		
		IF !llError		
			SELECT curTimeSheet
			APPEND BLANK

			REPLACE tsType WITH toType.tsType
			replace tmId   WITH tnCurrentTemplate
			REPLACE tsPay  WITH 0
			REPLACE tsDate WITH DATE()
			REPLACE tsEmp  WITH tnStaff
			REPLACE tsDayNbr WITH this.getdaynumber(tcDay)
	
			IF toType.showWeek AND !toType.readOnlyWeek
				REPLACE tsWeek WITH tnWeek
			ENDIF
			IF toType.showDay AND !toType.readOnlyDay
				REPLACE tsDay WITH tcDay
			ENDIF
			IF toType.showLeaveType AND !toType.readOnlyLeaveType
				REPLACE tsType WITH toLeaveCode.code
			ENDIF
			IF toType.showOtherType AND !toType.readOnlyOtherType
				REPLACE tsType WITH toOtherCode.code
			ENDIF
			IF toType.showCode AND !toType.readOnlyCode
				REPLACE tsCode WITH VAL(toAllowCode.code)
			ENDIF
			IF toType.showStart AND !toType.readOnlyStart
				REPLACE;
					tsStart		WITH ttStart;
					tsFinish	WITH ttEnd;
					tsBreak		WITH ttBreak;
					tsUnits		WITH tnUnits
			ENDIF
			IF toType.showUnits AND !toType.readOnlyUnits
				REPLACE tsUnits WITH tnUnits
			ENDIF
			IF toType.showReduce AND !toType.readOnlyReduce
				REPLACE tsUnits2 WITH tnReduce
			ENDIF
			IF toType.showWageType AND !toType.readOnlyWageType
				REPLACE tsWageType WITH VAL(toWageCode.code)
			ENDIF
			IF toType.showRateCode AND !toType.readOnlyRateCode
				REPLACE tsRateCode WITH tnRateCode
			ENDIF
			IF toType.showCostCent AND !toType.readOnlyCostCent
				REPLACE tsCostCent WITH VAL(toCostCentCode.code)
			ENDIF
*			IF toType.showJobCode AND !toType.readOnlyJobCode
*				REPLACE tsXXXX WITH VAL(toJobCode.code)	&&LATER: Add correct dbFieldName...
*			ENDIF
		ENDIF
		
		RETURN !llError
	ENDFUNC

	*--------------------------------------------------------------------------------*
	PROCEDURE SaveNewTemplateEntries()
		LOCAL loTypes, lcType, loType, lcValue, lnCount, lnEditID, lnSaved, llManager, lnCurrentPay, llIsError
		LOCAL lnErrored, lnCurrentStaff, lnCurrentGroup, lnRowCount, lnStaff, lnWeek, lcDay
		LOCAL loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce
		LOCAL loWageCode, lnRateCode, loCostCentCode, loJobCode, lnEditId, lcAddGroup, lnCurrentTemplate

		llManager = This.IsManager(This.Employee)
		lnErrored = 0
		loRetainList = Factory.GetRetainListObject()

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lcType = Request.Form("addType")
			loTypes = This.GetTemplateTypes(.F.)
			IF EMPTY(loTypes.GetKey(lcType))
				This.AddError("Unknown entry type!")
				loType = null
			ELSE

				loType = loTypes.Item(lcType)

				lnCount = VAL(Request.Form("count"))
				lnEditID = EVL(VAL(Request.Form("edit")),0)
				lnCurrentTemplate = EVL(VAL(Request.Form("currentTemplate")), 0)
				lnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
				lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)
				lcAddGroup =  EVL(Request.Form("addGroup"), "")

				CREATE CURSOR curStaff (tsEmp I(4,0))
				SELECT curStaff
				INDEX on tsEmp TAG tsEmp
				SET ORDER TO tsEmp
				GO top

				IF ALLTRIM(LOWER(lcAddGroup)) == "allgroups"
					lnCurrentGroup = -1
					IF llManager
						IF this.GetGroupsForManager(this.employee, "curManGrp")
							SELECT curManGrp
							GO top
							DO WHILE NOT EOF("curManGrp")
								IF this.GetEmployeesByGroupCode(curManGrp.grCode, "curEmpGrp")
									GO TOP IN "curEmpGrp"
									DO WHILE NOT EOF("curEmpGrp")
										m.tsemp = curEmpGrp.mywebcode
										IF NOT SEEK(m.tsemp,"curStaff")
											INSERT INTO curStaff FROM memvar
										ENDIF
										SKIP IN "curEmpGrp"
									ENDDO
								ENDIF
								SKIP IN "curManGrp"
							ENDDO
						ENDIF
					ENDIF
				ENDIF

				IF ALLTRIM(LOWER(lcAddGroup)) == "employee"
					IF lnCurrentStaff == -1
						IF lnCurrentGroup > 0
							IF this.GetEmployeesByGroupCode(lnCurrentGroup, "curEmpGrp")
								GO TOP IN "curEmpGrp"
								DO WHILE NOT EOF("curEmpGrp")
									m.tsemp = curEmpGrp.mywebcode
									IF NOT SEEK(m.tsemp,"curStaff")
										INSERT INTO curStaff FROM memvar
									ENDIF
									SKIP IN "curEmpGrp"
								ENDDO
							ENDIF
						ENDIF
					ENDIF
				ENDIF

				IF USED("curManGrp")
					USE IN "curManGrp"
				ENDIF

				IF USED("curEmpGrp")
					USE IN "curEmpGrp"
				ENDIF
				
				IF ALLTRIM(LOWER(lcAddGroup)) == "thisgroup"
					IF lnCurrentGroup > 0
						IF this.GetEmployeesByGroupCode(lnCurrentGroup, "curEmpGrp")
							GO TOP IN "curEmpGrp"
							DO WHILE NOT EOF("curEmpGrp")
								m.tsemp = curEmpGrp.mywebcode
								IF NOT SEEK(m.tsemp,"curStaff")
									INSERT INTO curStaff FROM memvar
								ENDIF
								SKIP IN "curEmpGrp"
							ENDDO
						ENDIF
					ENDIF
				ENDIF
		
				IF lnCurrentStaff > 0
	 				m.tsEmp = lnCurrentStaff
					IF NOT SEEK(m.tsEmp,"curStaff")
						INSERT INTO curStaff FROM memvar
					ENDIF
				ENDIF
					
				IF USED("curEmpGrp")
					USE IN "curEmpGrp"
				ENDIF

				SELECT curStaff
				GO top
				IF EOF()
					lnCount = 0
				ENDIF

				lnSaved = 0

				IF lnCount < 1
					This.AddError("Nothing to save!")
				ELSE
					lnSaved = 0
					
					SELECT timesheet
					GO top
					lnFld = AFIELDS(laFld,"timesheet")
					CREATE CURSOR curTimeSheet FROM ARRAY laFld
					SELECT curTimeSheet
					GO top

					FOR lnI = 1 TO lnCount
						loLeaveCode		= null
						loOtherCode		= null
						loAllowCode		= null
						ltStart			= {}
						ltEnd			= {}
						ltBreak			= {}
						lnUnits			= 0
						lnReduce		= 0
						loWageCode		= null
						lnRateCode		= 0
						loCostCentCode	= null
						loJobCode		= null

						SELECT curStaff
						GO top
						
						DO WHILE NOT EOF("curStaff")

							This.CollectTemplateEntryFormData(loType, lnI, curStaff.tsEmp, @lnWeek, @lcDay,;
															 @loLeaveCode,	@loOtherCode, @loAllowCode,;
															 @ltStart, @ltEnd,	@ltBreak, @lnUnits,;
															 @lnReduce, @loWageCode, @lnRateCode,;
															 @loCostCentCode, @loJobCode)

							IF This.TempSingleTemplateEntry(loType, lnCurrentTemplate, curStaff.tsEmp, lnWeek,;
															 lcDay,	loLeaveCode, loOtherCode, loAllowCode,;
															 ltStart, ltEnd, ltBreak, lnUnits, lnReduce,;
															 loWageCode, lnRateCode, loCostCentCode,;
															 loJobCode, lnCurrentGroup, @lnRowCount)	

								lnSaved = lnSaved + 1
							ELSE
								lnErrored = lnErrored + 1
								* Removed as the URL gets too long...
								* This.RetiainTimeEntry(loType, loRetainList, lnErrored, llManager, lnStaff, ldDate, loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce, loWageCode, loCostCentCode, loJobCode)
							ENDIF

							SKIP IN "curStaff"

						ENDDO
					ENDFOR
				ENDIF
			ENDIF
		ENDIF

		IF USED("curStaff")
			USE IN "curStaff"
		ENDIF

		IF USED("curTimeSheet")
			SELECT curTimeSheet
			GO top
			DO WHILE NOT EOF("curTimeSheet")
				SCATTER memvar
				m.tsId = 0
				INSERT INTO timesheet FROM memvar
				SKIP IN "curTimeSheet"
			ENDDO
			USE IN "curTimeSheet"
		ENDIF

		loRetainList.SetEntry("currentTemplate", TRANSFORM(EVL(VAL(Request.Form("currentTemplate")), NO_TEMPLATE)))
		loRetainList.SetEntry("currentGroup", TRANSFORM(EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)))
		loRetainList.SetEntry("currentStaff", TRANSFORM(EVL(VAL(Request.Form("currentStaff")), This.Employee)))
		loRetainList.SetEntry("addgroup", TRANSFORM(EVL(VAL(Request.Form("addgroup")), lcAddGroup)))
		loRetainList.SetEntry("addType", lcType)
		loRetainList.SetEntry("count", "0")	&& Removed TRANSFORM(lnErrored) as the URL gets too long.

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		IF !ISNULL(loType)
			IF lnSaved == 1
				This.AddUserInfo("Template Entry Saved.")
			ELSE
				IF lnSaved > 0
					This.AddUserInfo(TRANSFORM(lnSaved) + ' ' + " Template Entries Saved.")
				ENDIF
			ENDIF
		ENDIF

		Response.Redirect("TemplateEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcType)), '#' + lcType + "Entries", ""))
	ENDPROC

	*--------------------------------------------------------------------------------*
	PROCEDURE SaveTemplateEntry()
		LOCAL lcType, loType, loTypes, loRetainList, lcFrom, lcValue, lnEditId, llManager, llIsError
		LOCAL lnCurrentStaff, lnCurrentGroup, lnRowCount, lnStaff, lnWeek, lcDay, loLeaveCode
		LOCAL loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce
		LOCAL loWageCode, lnRateCode, loCostCentCode, loJobCode, lnCurrentTemplate

		llManager = This.IsManager(This.Employee)

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lcType = Request.Form("type")
			loTypes = This.GetTemplateTypes(.F.)
			IF EMPTY(loTypes.GetKey(lcType))
				This.AddError("Unknown entry type!")
			ELSE
				loType = loTypes.Item(lcType)

				lnEditID = EVL(VAL(Request.Form("edit")),0)
				lnCurrentTemplate = EVL(VAL(Request.Form("currentTemplate")), 0)
				lnCurrentStaff = EVL(VAL(Request.Form("currentStaff")), This.Employee)
				lnCurrentGroup = EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)

				SELECT timesheet
				LOCATE FOR tsId == lnEditId
				IF !FOUND()
					This.AddError("Entry not found!")
				ELSE
					lnStaff			= lnCurrentStaff
					ldDate			= {}
					loLeaveCode		= null
					loOtherCode		= null
					loAllowCode		= null
					ltStart			= {}
					ltEnd			= {}
					ltBreak			= {}
					lnUnits			= 0
					lnReduce		= 0
					loWageCode		= null
					lnRateCode		= 0
					loCostCentCode	= null
					loJobCode		= null

					This.CollectTemplateEntryFormData(loType, "", lnStaff, @lnWeek, @lcDay,;
														 @loLeaveCode,	@loOtherCode, @loAllowCode,;
														 @ltStart, @ltEnd,	@ltBreak, @lnUnits,;
														 @lnReduce, @loWageCode, @lnRateCode,;
														 @loCostCentCode, @loJobCode)


					llIsError = !This.SaveSingleTemplateEntry(loType, lnCurrentTemplate, lnStaff, lnWeek,;
														 lcDay,	loLeaveCode, loOtherCode, loAllowCode,;
														 ltStart, ltEnd, ltBreak, lnUnits, lnReduce,;
														 loWageCode, lnRateCode, loCostCentCode,;
														 loJobCode)
				ENDIF
			ENDIF
		ENDIF
		
		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentTemplate", TRANSFORM(EVL(VAL(Request.Form("currentTemplate")), NO_TEMPLATE)))
		loRetainList.SetEntry("currentGroup", TRANSFORM(EVL(VAL(Request.Form("currentGroup")), MY_DETAILS_GROUP)))
		loRetainList.SetEntry("currentStaff", TRANSFORM(EVL(VAL(Request.Form("currentStaff")), This.Employee)))

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		IF !ISNULL(loType)
			IF !llIsError
				This.AddUserInfo("Template Entry Saved.")
			ELSE
				This.AddUserInfo("Template Entry NOT Saved.")
			ENDIF			
		ENDIF
		
		Response.Redirect("TemplateEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&'))

		RETURN
		
	ENDPROC

	*--------------------------------------------------------------------------------*
	PROCEDURE DeleteTemplateRow()
		LOCAL lcFrom, lnId, lnCurrentStaff, lnCurrentTemplate, loRetainList, lcValue, llIsError

		lnCurrentTemplate = VAL(Request.QueryString("currentTemplate"))
		lnCurrentStaff = EVL(VAL(Request.QueryString("currentStaff")), This.Employee)

		IF !(This.SelectData(This.Licence, "timesheet");
		  AND This.SelectData(This.Licence, "myStaff"))
			This.AddError("Page Setup Failed!")
		ELSE
			lnId = VAL(Request.QueryString("edit"))
			IF EMPTY(lnId)
				This.AddError("No entry to delete!")
			ELSE

				SELECT timesheet
				LOCATE FOR tsId == lnId AND tmId <> 0 ;
								AND (tsEmp == lnCurrentStaff OR lnCurrentStaff == EVERYONE_OPTION) && ...for current emp if or Everyone
				IF !FOUND()
					This.AddError("Cannot find entry to delete!")
				ELSE
					IF tsDownload
						This.AddError("Cannot delete downloaded entry.")
					ELSE
						IF !This.CheckAccess(lnCurrentStaff, This.IsManager(This.Employee), .T.)	&& Allow Everyone option
							This.AddError("You do not have access to this page.")
						ELSE
							SELECT timesheet
							IF tsEmp != lnCurrentStaff AND lnCurrentStaff != EVERYONE_OPTION	&& If we are looking at Everyone, we can't check this anyway
								This.AddError("You do not have access to that entry!")
							ELSE
								DELETE FROM timesheet WHERE tsId == lnId

								This.AddUserInfo("Entry deleted.")
							ENDIF
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		loRetainList = Factory.GetRetainListObject()
		loRetainList.SetEntry("currentTemplate", lnCurrentTemplate)
		loRetainList.SetEntry("currentGroup", VAL(Request.QueryString("currentGroup")))
		loRetainList.SetEntry("currentStaff", lnCurrentStaff)

		lcFrom = Request.Form("from")

		lcValue = Request.QueryString("type")

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		Response.Redirect("TemplateEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError OR EMPTY(lcValue)), '#' + lcValue + "Entries", ""))
	ENDPROC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	FUNCTION GetTemplateTypes(tlSelectData, tnCurrentTemplate, tnCurrentGroup, tnCurrentStaff, tcFilter) AS Object
		LOCAL lcFilter2, loTypes, llAusie, loLeaveCodes, lnI, llRates, lcEmpWhere

		llAusie = This.IsAustralia()

		IF tlSelectData AND !(This.SelectData(This.Licence, "timesheet");
						  AND This.SelectData(This.Licence, "myStaff");
						  AND This.SelectData(This.Licence, "myPays"))
			This.AddError("Cannot get template data!")
			RETURN null
		ENDIF

		TRY
			SELECT myStaff
			LOCATE FOR myRates
			llRates = FOUND()
		CATCH
			llRates = .F.
		ENDTRY

		IF !tlSelectData
			tcFilter = ""
		ENDIF

		loTypes = CREATEOBJECT("COLLECTION")

		IF tlSelectData AND tnCurrentStaff == EVERYONE_OPTION
			IF !This.GetEmployeesByGroupCode(tnCurrentGroup, "curGroupStaff")
				This.AddError("Cannot load current group!")
				RETURN loTypes	&& bail out on error
			ENDIF

		*!*	lcEmpWhere = "INLIST(tsEmp"						&& 01/07/2010  CMGM  TTP5692  Error is caused by INLIST: can only take 25 expr including the search expr
			lcEmpWhere = ""									&& 01/07/2010  CMGM  TTP5692  Replace it by native SQL IN()

			SELECT curGroupStaff
			SCAN
				IF curGroupStaff.myWebCode < WEBCODE_PAYROLLUSER_BOUNDARY OR curGroupStaff.myWebCode < 2	&&NOTE: always hiding PayrollUsers here; hide Admin user either way...
					LOOP
				ENDIF

				lcEmpWhere = lcEmpWhere + "," + TRANSFORM(curGroupStaff.myWebCode)
			ENDSCAN
			
			IF EMPTY(_TALLY)
				lcEmpWhere = ".T."								&& 25/02/2011  CMGM  2011.02  TTP6615  Fix incorrect sql "tsEmp IN ()" below due to PayrollUsers being hidden
			ELSE
			*!*	lcEmpWhere = lcEmpWhere + ")"					&& 01/07/2010  CMGM  TTP5692  
				lcEmpWhere = SUBSTR(lcEmpWhere, 2)				&& 01/07/2010  CMGM  TTP5692  Remove the initial ","
				lcEmpWhere = "tsEmp IN (" + lcEmpWhere + ")"	&& 01/07/2010  CMGM  TTP5692  Build SQL
			ENDIF
		ELSE
			lcEmpWhere = "tsEmp == tnCurrentStaff"
		ENDIF

		&&LATER: add J to the below when a JobCode is needed...
		*IF This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M")
		**JA check additional security rights similar to time entry
		IF (This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M")) AND  This.CheckRights("TS_TIMESHEET_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere.;
					AND ((tmId IN (SELECT pay_pk FROM mypays WHERE pay_type == 1)) OR tmId == MY_TEMPLATE) ;
					AND tsType = 'M' AND tmId == tnCurrentTemplate ;
					&tcFilter.;
					INTO CURSOR curTimes;
					ORDER BY tsWeek,tsID asc
					&&ORDER BY tsWeek,tsDayNbr,tsStart,tsID asc
			ENDIF

			loTypes.Add(This.NewTemplateTypeObject("time", 'M', "Timesheet", [.CheckRights("TS_TIMESHEET_V")], IIF(llRates, "EDSFBuWAC", "EDSFBuWC"), "curTimes", _TALLY), "time")
			
		ENDIF
        **JA check additional security rights similar to time entry
		*IF This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M")
		IF (This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M")) AND This.CheckRights("TS_WAGES_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere.;
					AND ((tmId IN (SELECT pay_pk FROM mypays WHERE pay_type == 1)) OR tmId == MY_TEMPLATE) ;
					AND tsType = 'W' AND tmId == tnCurrentTemplate ;
					&tcFilter.;
					INTO CURSOR curWages;
					ORDER BY tsWeek,tsID asc
			ENDIF
			loTypes.Add(This.NewTemplateTypeObject("wages", 'W', "Wages", [.CheckRights("TS_WAGES_V")], IIF(llRates, "EDUWAC", "EDUWC"), "curWages", _TALLY), "wages")
		ENDIF
        **JA check additional security rights similar to time entry
        **IF This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M") 
		IF (This.CheckRights("TEM_GROUP_M") OR This.CheckRights("TEM_EMPLOYEE_M")) AND This.CheckRights("TS_LEAVE_V")

			loLeaveCodes = This.GetLeaveCodes(IIF(!tlSelectData OR tnCurrentStaff == EVERYONE_OPTION, This.Employee, tnCurrentStaff),.T.)

			lcFilter2 = tcFilter + " AND INLIST(tsType"

			FOR lnI = 1 TO loLeaveCodes.Count
				lcFilter2 = lcFilter2 + ", '" + loLeaveCodes.Item(lnI).code + "'"
			NEXT

			lcFilter2 = lcFilter2 + ")"

			IF !("INLIST(tsType)" $ lcFilter2)
				IF tlSelectData
					SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE;
					&lcEmpWhere.;
					AND ((tmId IN (SELECT pay_pk FROM mypays WHERE pay_type == 1)) OR tmId == MY_TEMPLATE) ;
					AND tmId == tnCurrentTemplate ;
					&lcFilter2.;
					INTO CURSOR curLeave;
					ORDER BY tsWeek,tsID asc
				ENDIF
				loTypes.Add(This.NewTemplateTypeObject("leave", 'S', "Leave", [.CheckRights("TS_LEAVE_V")], "EDTURC", "curLeave", _TALLY), "leave")	&& this doesn't need more authz as it's left out if no leaveTypes are available
			ENDIF
		ENDIF

		IF This.CheckRights("TS_ALLOWANCES_V")
			IF tlSelectData
				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE &lcEmpWhere. ;
					AND ((tmId IN (SELECT pay_pk FROM mypays WHERE pay_type == 1)) OR tmId == MY_TEMPLATE) ;
					AND tsType = 'A' AND tmId == tnCurrentTemplate ;
					&tcFilter.;
					INTO CURSOR curAllowances;
					ORDER BY tsWeek,tsID asc
			ENDIF
			loTypes.Add(This.NewTemplateTypeObject("allowances", 'A', "Allowances", [.CheckRights("TS_ALLOWANCES_V")], "EDKUC", "curAllowances", _TALLY), "allowances")
		ENDIF

		IF This.CheckRights("TS_OTHER_V")
			IF tlSelectData
				IF llAusie
					lcFilter2 = tcFilter + " AND INLIST(tsType, 'D')"
				ELSE
					lcFilter2 = tcFilter + " AND INLIST(tsType, 'R', 'D')"
				ENDIF

				SELECT *, ALLTRIM(mySurname) + ", " + ALLTRIM(myName) AS fullName;
					FROM timesheet;
					JOIN myStaff ON tsEmp = myWebCode;
					WHERE &lcEmpWhere. ;
					AND ((tmId IN (SELECT pay_pk FROM mypays WHERE pay_type == 1)) OR tmId == MY_TEMPLATE) ;
					AND tmId == tnCurrentTemplate;
					&lcFilter2.;
					INTO CURSOR curOther;
					ORDER BY tsWeek,tsID asc
			ENDIF
			loTypes.Add(This.NewTemplateTypeObject("other", 'D', "Other", [.CheckRights("TS_OTHER_V")], "EDOU", "curOther", _TALLY), "other")
		ENDIF

		RETURN loTypes
	ENDFUNC

	*-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=*
	* Sets up everything needed for the TemplateControl.	(Note: doesn't work if the control uses a non-default formField name!)
	FUNCTION SetupTemplateControlData(rnCurrentTemplate, poRetainList) AS Integer
		LOCAL lcFilter, llManager

		IF !This.SelectData(This.Licence, "myPays")
			RETURN .f.	&& no data!
		ELSE
			tcPayStatus = "open" && Only Return 'Open' Templates
			DO CASE
				CASE tcPayStatus == "all"
					lcFilter = ".T."
				CASE tcPayStatus == "open"
					lcFilter = "pay_status == 1"
				CASE tcPayStatus == "closed"
					lcFilter = "pay_status == 2"
			ENDCASE

			* Select templates only, 
			SELECT mypays
			lnFldCnt = aFields(laFld,"mypays")
			CREATE CURSOR curTemplates FROM ARRAY laFld
			
			llManager = This.IsManager(This.Employee)
			llRights = .f.
			IF llManager 
				llRights = This.CheckRights("TEM_GROUP_M")
			ENDIF
						
			IF llRights
				SELECT curTemplates
				GO top
			
				SELECT mypays
				SET ORDER TO pay_name
				GO top
				lnMaxTemp = 0
				DO WHILE NOT EOF()
					SELECT mypays
					IF pay_type == 1
						IF pay_status == 1
							SELECT mypays
							SCATTER memvar
							IF pay_pk > lnMaxTemp
								lnMaxTemp = pay_pk
							ENDIF
							SELECT curTemplates
							INSERT INTO curTemplates FROM memvar
						ENDIF
					ENDIF
					SELECT mypays
					SKIP
				ENDDO
			ENDIF
					
			* Will at least display "<My Template>"
			SELECT curTemplates
			APPEND BLANK
*			replace pay_pk     WITH lnMaxTemp + 1
			replace pay_pk     WITH MY_TEMPLATE
			replace pay_name   WITH DEF_LABEL_TEMPLATE
			replace pay_status WITH 1
			replace pay_type   WITH 1
			replace pay_orig   WITH 0

			rnCurrentTemplate = EVL(VAL(Request.Form("currentTemplate")), VAL(Request.QueryString("currentTemplate")))

			IF VARTYPE(rnCurrentTemplate) != 'N' OR EMPTY(rnCurrentTemplate)
				rnCurrentTemplate = NO_TEMPLATE
			ENDIF

			SELECT curTemplates
			GO top
			IF EOF()
				rnCurrentTemplate = 0
			ELSE
				IF rnCurrentTemplate < 0
					rnCurrentTemplate = pay_pk
				ENDIF
			ENDIF

			llManager = This.IsManager(This.Employee)
			llRights = .f.
			IF llManager 
				llRights = This.CheckRights("TEM_GROUP_M")
			ENDIF
			
			IF NOT llRights
				rnCurrentTemplate = MY_TEMPLATE
			ENDIF

			poRetainList.SetEntry("currentTemplate", rnCurrentTemplate)

			RETURN .t.
		ENDIF
		WAIT clear
		
	ENDFUNC

	*--------------------------------------------------------------------------------*
	PROCEDURE ApplyTemplateToTimeSheet()
		LOCAL lcValue, lnCount, lnEditID, lnSaved, llManager, lnCurrentPay, llIsError, ldCurrentPayDate
		LOCAL lnErrored, lnCurrentStaff, lnCurrentGroup, lnRowCount, lnStaff, lnWeek, lcDay, llAllOk
		LOCAL loLeaveCode, loOtherCode, loAllowCode, ltStart, ltEnd, ltBreak, lnUnits, lnReduce
		LOCAL loWageCode, lnRateCode, loCostCentCode, loJobCode, lnEditId, lcAddGroup, lnCurrentTemplate
		LOCAL lnStartDay, lcErrMess

		lcErrMess = "Nothing To Save!"
		llManager = This.IsManager(This.Employee)
		lnErrored = 0
		llAllOk = .t.

		loRetainList = Factory.GetRetainListObject()

		IF !(This.SelectData(This.Licence, "timesheet");
			  AND This.SelectData(This.Licence, "myStaff");
			  AND This.SelectData(This.Licence, "myPays"))
			This.AddError("Page Setup Failed!")
		ELSE
			lnCurrentTemplate = EVL(VAL(Request.QueryString("currentTemplate")), NO_TEMPLATE)
			lnCurrentStaff = EVL(VAL(Request.QueryString("currentStaff")), This.Employee)
			lnCurrentGroup = EVL(VAL(Request.QueryString("currentGroup")), MY_DETAILS_GROUP)
			lnCurrentPay = EVL(VAL(Request.QueryString("currentPay")), -1)

			ldCurrentPayDate = {  /  /  }
			IF lnCurrentPay > 0
				SELECT mypays
				SET ORDER TO pay_pk
				IF SEEK(lnCurrentPay,"mypays")
					ldCurrentPayDate = mypays.pay_date
				ENDIF
			ENDIF

			llAllOk = .t.
			lnSaved = 0
			lnStartDay = 0

			IF EMPTY(ldCurrentPayDate)
				llAllOk = .f.
			ELSE
				lnStartDay = DOW(ldCurrentPayDate,2)
			ENDIF

			IF llAllOk
				CREATE CURSOR curStaff (tsEmp I(4,0))
				SELECT curStaff
				INDEX on tsEmp TAG tsEmp
				SET ORDER TO tsEmp
				GO top

				IF lnCurrentStaff == -1
					IF lnCurrentGroup > 0
						IF this.GetEmployeesByGroupCode(lnCurrentGroup, "curEmpGrp")
							GO TOP IN "curEmpGrp"
							DO WHILE NOT EOF("curEmpGrp")
								m.tsemp = curEmpGrp.mywebcode
								IF NOT SEEK(m.tsemp,"curStaff")
									INSERT INTO curStaff FROM memvar
								ENDIF
								SKIP IN "curEmpGrp"
							ENDDO
						ENDIF
					ENDIF
				ENDIF
		
				IF lnCurrentStaff > 0
					m.tsEmp = lnCurrentStaff
					IF NOT SEEK(m.tsEmp,"curStaff")
						INSERT INTO curStaff FROM memvar
					ENDIF
				ENDIF
				
				IF USED("curEmpGrp")
					USE IN "curEmpGrp"
				ENDIF

				SELECT curStaff
				GO top
				IF EOF()
					llAllOk = .f.
				ENDIF

				IF llAllOk
					SELECT timesheet
					GO top
					lnFld = AFIELDS(laFld,"timesheet")
					CREATE CURSOR curTimeSheet FROM ARRAY laFld
					SELECT curTimeSheet
					INDEX on ALLTRIM(STR(tsWeek,1))+ALLTRIM(STR(tsDayNbr,1)) TAG main
					INDEX on DTOS(tsdate) TAG tsDate
					GO top
				
					poAllowCodes = This.GetAllowanceCodes()
					poOtherCodes = This.GetOtherCodes()
					poWageCodes = This.GetWageCodes()

					cTimeChk = "M"
					cOtherChk = ""

					FOR lnJ = 1 TO poOtherCodes.Count
						loCode = poOtherCodes.Item(lnJ)
						cOtherChk = cOtherChk + ALLTRIM(loCode.code)
					ENDFOR
	
					SELECT timesheet
					SET ORDER to tmId
					GO top
					IF SEEK(lnCurrentTemplate)
						DO WHILE NOT EOF("timesheet") AND timesheet.tmId == lnCurrentTemplate
							IF SEEK(timesheet.tsEmp,"curStaff")
								SELECT timesheet
								SCATTER memvar
								m.tmid = 0
								m.tsdate = {  /  /  }
								m.tsApproved = .f.
								m.tsDownload = .f.

								cLeaveChk = ""
								poLeaveCodes = This.GetLeaveCodes(timesheet.tsEmp, .F.)

								FOR lnJ = 1 TO poLeaveCodes.Count
									loCode = poLeaveCodes.Item(lnJ)
									cLeaveChk = cLeaveChk + ALLTRIM(loCode.code)
								ENDFOR
								
								* MY 15-11-2012 remove security check as it should not be checked here
								lOk = .F.
								IF m.tstype == cTimeChk
									*IF This.CheckRightsFor(timesheet.tsEmp,"TS_TIMESHEET_V")
										lOk = .T.
									*ENDIF
								ENDIF
			
								IF m.tstype == "W"
									*IF This.CheckRightsFor(timesheet.tsEmp,"TS_WAGES_V")
										FOR lnJ = 1 TO poWageCodes.Count
											loCode = poWageCodes.Item(lnJ)
											IF m.tswagetype == VAL(loCode.code)
												lOk = .T.
											ENDIF
										ENDFOR
									*ELSE
									*	lOk = .F.
									*ENDIF
								ENDIF

								IF m.tstype == "A"
									*IF This.CheckRightsFor(timesheet.tsEmp,"TS_ALLOWANCES_V")
										FOR lnJ = 1 TO poAllowCodes.Count
											loCode = poAllowCodes.Item(lnJ)
											IF m.tscode == VAL(loCode.code)
												lOk = .T.
											ENDIF
										ENDFOR
									*ELSE
									*	lOk = .F.
									*ENDIF
								ENDIF						

								IF m.tstype <> "W" AND m.tstype <> "A" AND m.tstype <> cTimeChk
									IF NOT EMPTY(cLeaveChk)
										*IF This.CheckRightsFor(timesheet.tsEmp,"TS_LEAVE_V")
											IF m.tsType$cLeaveChk
												lOk = .T.
											ENDIF
										*ENDIF
									ENDIF
								
									IF NOT EMPTY(cOtherChk)
										*IF This.CheckRightsFor(timesheet.tsEmp,"TS_OTHER_V")
											IF m.tsType$cOtherChk
												lOk = .T.
											ENDIF
										*ENDIF
									ENDIF
								ENDIF
											
								IF lOk
									INSERT INTO curTimeSheet FROM memvar
								ENDIF
							ENDIF
							SKIP IN 'timesheet'
						ENDDO
					ENDIF

					SELECT curTimeSheet
					SET ORDER TO
					GO top
					IF EOF()
						llAllOk = .f.
					ELSE
						SELECT curStaff
						GO top
						DO WHILE NOT EOF("curStaff")
							SELECT curTimeSheet
							SET FILTER TO tsEmp == curStaff.tsEmp
							GO top
							IF NOT EOF("curTimeSheet")
								lnWeek = 1
								ldActualDate = ldCurrentPayDate
								FOR lnLoopWeek = 1 TO 6
									lnStartDayWork = lnStartDay
									lnWeek = lnLoopWeek				
									FOR lnLoopDay = 1 TO 7
										GO TOP IN 'curTimeSheet'
										DO WHILE NOT EOF("curTimeSheet")
											IF tsWeek == lnWeek
												IF tsDayNbr == lnStartDayWork
													IF EMPTY(curTimeSheet.tsDate)
														REPLACE tsPay  WITH lnCurrentPay
														REPLACE tsDate WITH ldActualDate
													ENDIF
												ENDIF
											ENDIF
											SKIP IN "curTimeSheet"
										ENDDO
										lnStartDayWork = lnStartDayWork + 1
										IF lnStartDayWork > 7
											lnStartDayWork = 1
										ENDIF
										ldActualDate = ldActualDate + 1
									ENDFOR
								ENDFOR
							ENDIF
							SKIP IN "curStaff"
						ENDDO
					ENDIF
					SELECT curTimeSheet
					SET ORDER TO
					SET FILTER TO
					GO top
				ENDIF
			ENDIF
			
			IF !llAllOk
				This.AddError(lcErrMess)
			ENDIF
		ENDIF

		IF USED("curStaff")
			USE IN "curStaff"
		ENDIF

		IF USED("curTimeSheet")
			SELECT curTimeSheet
			GO top
*			BROWSE
			IF llAllOk
				lnSaved = RECCOUNT("curTimeSheet")
				INSERT INTO timesheet SELECT * FROM curTimeSheet
			ENDIF
			USE IN "curTimeSheet"
		ENDIF
		
		loRetainList.SetEntry("currentTemplate", TRANSFORM(EVL(VAL(Request.QueryString("currentTemplate")), NO_TEMPLATE)))
		loRetainList.SetEntry("currentGroup", TRANSFORM(EVL(VAL(Request.QueryString("currentGroup")), MY_DETAILS_GROUP)))
		loRetainList.SetEntry("currentStaff", TRANSFORM(EVL(VAL(Request.QueryString("currentStaff")), This.Employee)))
		loRetainList.SetEntry("currentPay", TRANSFORM(EVL(VAL(Request.QueryString("currentPay")), -1)))

		llIsError = This.IsError()	&& must cache this as AppendMessages() will empty any errors.

		IF llAllOk		
			IF lnSaved == 1
				This.AddUserInfo("Timesheet Entry Saved.")
			ELSE
				IF lnSaved > 1
					This.AddUserInfo(TRANSFORM(lnSaved) + ' ' + " Timesheet Entries Saved.")
				ENDIF
			ENDIF
		ENDIF
		
		Response.Redirect("TimeEntryPage.si" + loRetainList.SaveState(.F., .F., .T.) + This.AppendMessages('&') + IIF(!(llIsError), '#' + "Timesheet Entries", ""))
	ENDPROC

	*--------------------------------------------------------------------------------*
	PROCEDURE CheckTemplates() as Boolean
		LOCAL lnCurrentStaff, llManager, llApplyGrp, llApplyMy, lnOk

		IF !(This.SelectData(This.Licence, "timesheet"))
			This.AddError("Page Setup Failed!")
		ENDIF
		
		lnCurrentStaff = This.Employee
		llManager = This.IsManager(This.Employee)
		llApplyGrp = .f.
		llApplyMy = .f.		

		IF llManager 
			llApplyGrp = This.CheckRights("TEM_APPLY_M")
		ENDIF

		IF This.CheckRights("TEM_APPLY_E")
			SELECT * FROM timesheet WHERE (tsEmp == lnCurrentStaff AND tmiD == 99999);
					 INTO CURSOR ChkTemplates READWRITE

			SELECT ChkTemplates
			GO top
			IF !EOF("ChkTemplates")
				llApplyMy = .t.
			ENDIF
			
			USE IN 'ChkTemplates'
		ENDIF

		lnOk = .f.
		IF llApplyGrp OR llApplyMy
			lnOk = .t.
		ENDIF

		RETURN(lnOk)
	ENDPROC
	
	
	*################################################################################*
#DEFINE TOC_Nonce_Controls_

	*--------------------------------------------------------------------------------*

	*> Creates nonce
	FUNCTION NonceCreate() as String
		LOCAL lcNonce

		lcNonce = ""
		lcNonce = This.NonceGenerateHash()

		RETURN lcNonce
	ENDFUNC


	*--------------------------------------------------------------------------------*

	*> Creates nonce hash using current time and employee code
	FUNCTION NonceGenerateHash() as String
		LOCAL lcNonce, i
		LOCAL lcHash1, lcHash2

		lcNonce = ""

		* SYS(2) is seconds after midnight
		i = VAL(SYS(2))
		i = CEILING(VAL(SYS(2)) / NONCE_DURATION)

		* Convert to HEX
		lcHash1 = TRANSFORM(i, '@0x')
		lcHash1 = SUBSTR(lcHash1, 3)

		* Use employee number as well
		lcHash2	= TRANSFORM((VAL(Session.GetSessionVar("employee"))), '@0x') 
		lcHash2 = SUBSTR(lcHash2, 3)

		lcNonce = lcHash1 + lcHash2
		RETURN lcNonce
	ENDFUNC

	*--------------------------------------------------------------------------------*

	*> Validates if nonce is valid
	FUNCTION NonceIsValid(tcNonce) as Boolean
		LOCAL llNonceIsValid

		IF This.NonceGenerateHash() == tcNonce
			llNonceIsValid = .T.
		ENDIF
		
		RETURN llNonceIsValid
	ENDFUNC

	*--------------------------------------------------------------------------------*
	FUNCTION FindInList(teExpr, tcCheckList)
		*MY 18/10/2012 - Foxpro inlist function can can include up to 25 expressions. This function has no limit of expressions
		IF PCOUNT() <= 1
			RETURN .f.
		ENDIF

		LOCAL lnLoop, lcValue as Integer
		FOR lnLoop = 1 TO GETWORDCOUNT(tcCheckList, ",")
			lcValue = GETWORDNUM(tcCheckList, lnLoop, ",")
			IF VARTYPE(teExpr) = VARTYPE(lcValue)
				IF teExpr = lcValue
					RETURN .t.
				ENDIF
			ELSE
				IF TRANSFORM(teExpr) = TRANSFORM(lcValue)
					RETURN .t.
				ENDIF
			ENDIF 		
		ENDFOR 
		RETURN .f.
	ENDFUNC
	*--------------------------------------------------------------------------------*
	FUNCTION ShowTime(tnUnitDisp, tnTotal)
		*MY 18/10/2012 - Show time in correct format (fix 60:60 issue)
		* tnUnitDisp = 1: show as number 35.50
		* tnUnitDisp <> 1: show as time 35:30
		IF tnUnitDisp = 1
			IF tnTotal < 10
				RETURN "0" + ALLTRIM(STR(tnTotal,9,2))
			ELSE
				RETURN ALLTRIM(STR(tnTotal,9,2))
			ENDIF 
		ELSE
			LOCAL lcDeci as String
			LOCAL lnDeci as Number
			lnDeci = tnTotal - INT(tnTotal)
			IF lnDeci = 0
				lcDeci = "00"
			ELSE
				lcDeci = PADL(TRANSFORM(FLOOR(lnDeci * 60)), 2, "0")
				IF lcDeci = "60"
					lcDeci = "00"
				ENDIF
			ENDIF
			IF tnTotal < 10
				RETURN "0" + TRANSFORM(FLOOR(tnTotal)) + ":" + lcDeci
			ELSE
				RETURN TRANSFORM(FLOOR(tnTotal)) + ":" + lcDeci
			ENDIF  
		ENDIF
	ENDFUNC 
	*--------------------------------------------------------------------------------*
	FUNCTION CheckApplyTemplates(tlManager, tnCurrentGroup, tnCurrentStaff)
		IF tlManager and This.CheckRights("TEM_APPLY_M")
			RETURN .t.
		ELSE
			IF NOT This.CheckRights("TEM_APPLY_E")
				RETURN .f.
			ELSE 
				IF tnCurrentGroup = MY_DETAILS_GROUP
					RETURN .t.
				ELSE
					IF This.CheckEmptyGroup(tnCurrentGroup) = 0
						* no staff in the group
						RETURN .f.
					ELSE 
						IF tnCurrentStaff = This.Employee
							RETURN .t.
						ELSE
							RETURN .f.
						ENDIF 
					ENDIF
				ENDIF 
			ENDIF 
		ENDIF
		RETURN .f.
	ENDFUNC 
	*--------------------------------------------------------------------------------*
	FUNCTION CheckApplyPreviousPay(tlManager, tnCurrentGroup, tnCurrentStaff)
		IF tlManager and This.CheckRights("PREV_PAY_APPLY_M")
			RETURN .t.
		ELSE
			IF NOT This.CheckRights("PREV_PAY_APPLY_E")
				RETURN .f.
			ELSE 
				IF tnCurrentGroup = MY_DETAILS_GROUP
					RETURN .t.
				ELSE
					* check if the employee is in this group
					IF This.CheckEmptyGroup(tnCurrentGroup) = 0
						* no staff in the group
						RETURN .f.
					ELSE 
						IF tnCurrentStaff = This.Employee
							RETURN .t.
						ELSE
							RETURN .f.
						ENDIF 
					ENDIF
				ENDIF 
			ENDIF 
		ENDIF
		RETURN .f.
	ENDFUNC 
	*--------------------------------------------------------------------------------*
	FUNCTION CheckEmptyGroup(tnCurrentGroup)
	* return staff count 
		SELECT * ;
			FROM myteams ;
			WHERE tmmygroups = tnCurrentGroup ;
			INTO CURSOR curGroupStaffCount 
		RETURN _tally 
	ENDFUNC
	*--------------------------------------------------------------------------------*

ENDDEFINE
