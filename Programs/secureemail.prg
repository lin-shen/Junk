PARAMETERS losmtpbridge, pcmailserver, plusessl, pcusername, pcpassword, pcsender, pcsenderemail, pcrecipient, pcsubject, pcmessage, pcattachment, plsendemailasync, pcreplyto, pnEmailType, peLicence
LOCAL lcerror, lcLog, lcLicence

IF NOT EmailAddressValidation(pcRecipient)
	lcerror = "Invalid email address"
	RETURN lcerror
ENDIF

* ES-7156 Create SMTP log - we now using Popeye API, instead of SMTP
DO CASE
CASE TYPE("peLicence") = "C"		
	lcLicence = peLicence
CASE TYPE("peLicence") = "N"		
	lcLicence = TRANSFORM(peLicence)
OTHERWISE
	lcLicence = ""
ENDCASE

LOCAL FoxJson
FoxJson = NEWOBJECT('FoxJson', 'Fox_Json.prg')
FoxJson.StartJson()

	FoxJson.AddValue('subject', pcsubject)

	FoxJson.AddChild('custom')
		FoxJson.AddValue('text', 'Email on Popeye')
		FoxJson.AddValue('html', '<html><body>' + pcmessage + '</body></html>')
	FoxJson.EndChild()
	
	
	FoxJson.AddChild('from')
		FoxJson.AddValue('name', 'MyStaffInfo Admin')
		IF !EMPTY(pcreplyto) AND VARTYPE(pcreplyto)="C"
			FoxJson.AddValue('email', pcreplyto)
		ELSE 
			FoxJson.AddValue('email', 'MyStaffInfoAdmin@myob.com')
		ENDIF 
	FoxJson.EndChild()
	
	FoxJson.AddChild('to', .t.)
		FoxJson.AddChild('')
			FoxJson.AddValue('name', '')
			FoxJson.AddValue('email', pcrecipient)
		FoxJson.EndChild()
	FoxJson.EndChild(.t.)
	
*!*		IF !EMPTY(pcattachment) AND VARTYPE(pcattachment)="C"
*!*			FoxJson.AddValue('attachments', pcattachment)
*!*		ENDIF 
		
FoxJson.EndChild(.f., .t.)

LOCAL winhttpclient

TRY
	winhttpclient=CREATEOBJECT("WinHTTP.WinHTTPRequest.5.1")
	winhttpclient.OPEN("POST", "https://api.popeye.myob.com/email")
	winhttpclient.setrequestheader("x-myobapi-key", "2c040ff83d714faba5fbf714")
	winhttpclient.setrequestheader("content-type", "application/json")
	winhttpclient.setrequestheader("port", "587")
	winhttpclient.SEND(FoxJson.JsonText)
	
	lcLog = TTOC(DATETIME(), 1) + "," + lcLicence + "," + TRANSFORM(pnEmailType) 
CATCH TO ex	
	lcLog = TTOC(DATETIME(), 1) + "," + lcLicence + "," + TRANSFORM(pnEmailType) + ", Popeye error:" + ex.message
	lcerror = ex.message
ENDTRY

winhttpclient=NULL
FoxJson = null 
STRTOFILE(lcLog + CHR(13), "msi_email_log.txt", 1)	

RETURN lcerror

**************************************************************
FUNCTION EmailAddressValidation
**************************************************************
* This function returns True if the email address passed is a
* valid email address. Else it returns false. It creates vbscript
* class object for Regular Expression and test the passed string
* for the patterned assigned to its pattern property.
LPARAMETERS lcEmailAddress

* Do not test if empty.
IF EMPTY(lcEmailAddress)
	RETURN .F.
ENDIF

lcEmailAddress = ALLTRIM(lcEmailAddress)

LOCAL lAtPos, lPart1Email, lPart2Email

lAtPos= AT("@",lcEmailAddress)

lPart1Email = SUBSTR(lcEmailAddress,1,lAtPos-1)
lPart2Email = SUBSTR(lcEmailAddress,lAtPos+1,LEN(lcEmailAddress))


IF lAtPos=0 OR LEN(lcEmailAddress)=1 OR LEN(ALLTRIM(lPart1Email))=0 OR  LEN(ALLTRIM(lPart2Email))=0
	RETURN .F.
ELSE
	RETURN .T.
ENDIF

