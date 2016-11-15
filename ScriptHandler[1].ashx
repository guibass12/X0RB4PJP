option explicit
'=== Constant to Show or Hide Error messages
Const SHOW_SM_ERROR_MESSAGES                                = "False"
Const SM_SMART_MESSAGING_CLIENT_DB_NAME                     = "SmartMessaging.db"
Const SM_UIPROVIDER_ENUM_OBJECT                             = "SMUIMgrEnum"
Const SM_PROP_UIMGR_NAME                                    = "2"
Const SM_PROP_PROVIDER_APPLICATION_MGR                      = "6"
Const SM_PROP_PROVIDER_SUBSCRIPTION_MGR                     = "5"
Const SM_PROVIDER_NAME                                      = "2"
Const MC_PROP_APPLICATION_ID                                = "1"
Const MC_PROP_APPLICATION_NAME                              = "2"
Const MC_PROP_APPLICATION_VERSION                           = "4"
Const MC_PROP_APPLICATION_INSTALL_LCID						= "5"
Const MC_PROP_SUB_MGR_OBJ                                   = "0"
Const SM_PROP_PROVIDER_ENUM_CLIENTDB_OBJECT                 = "3"
Const SM_PROP_PROVIDER_ENUM_SYSTEM_MGR                      = "4"
Const SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT                 = "5"
Const SM_PROP_PROVIDER_ENUM_FILESYSTEM_OBJECT               = "6"
Const SM_CLIENTDB_OBJECT_CREATION_FAILED    				= "0"
Const SM_CLIENTDB_OBJECT_CREATION_SUCCESS					= "1"
Const SM_CLIENTDB_VERSION_LOCATION					        = "HKLM\Software\McAfee\MSM"  ' Pre-WSS 12.0
Const SM_CLIENTDB_VERSION_KEY_NAME					        = "SMVersion"
Const SM_VERSION_LOCATION       					        = "HKLM\Software\McAfee\MSM"  ' Pre-WSS 12.0
Const SM_REG_LOCATION                                       = "HKLM\Software\McAfee\Platform\MSM" ' WSS12.0+
Const SM_WEBUPDATE_LOCATION                                 = "HKLM\software\McAfee\MSC\Update\Webupdate"
Const SM_VERSIONID_KEY_NAME                                 = "VerID"
Const SM_VERSION_KEY_NAME       					        = "SMVersion"
Const SM_OOBE_LOCATION       					            = "HKLM\Software\McAfee\MSC\OEM\Manage\RGW\ActWiz"
Const SM_OOBE_KEY_NAME       					            = "OOBEModel"
Const SM_OOBE_ACTIVATIONDATE_KEY_NAME       				= "SvrActDate"
Const SM_DEFAULT_VERSION                                    = "1"
Const SM_NO_VERSION                                         = ""
Const SM_APP_VERSION_SPLITTER                               = "."
Const SM_VERSIONMGR_OBJECT_CREATION_SUCCESS                 = "1"
Const SM_VERSIONMGR_OBJECT_CREATION_FAILED                  = "0"
Const SM_CLIENTDB_DEFAULT_VERSION                           = "1"
Const SM_CLIENTDB_NO_VERSION                                = ""
Const SM_CLIENTDB_MISSING_DB                                = "0"
Const SM_CLIENTDB_DB_OPEN                                   = "1"
Const SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL                     = "1"
Const SM_CLIENTDB_QUERY_EXEC_FAIL                           = "0"
Const SM_CLIENTDB_ALL_TABLES_CREATED_SUCCESS                = "1"
Const SM_CLIENTDB_ALL_TABLES_CREATED_FAILED                 = "0"
Const SM_CLIENTDB_VERSION_VALIDATION_SUCCESS                = "1"
Const SM_CLIENTDB_VERSION_VALIDATION_FAILED                 = "0"
Const SM_DEFAULT_SYNC_FLAG_VALUE                            = "0"
Const SM_CLIENTDB_SYNC_MESSAGE_URL                          = "http://us.mcafee.com"
Const SM_MESSAGE_LOG_TABLE_NAME                             = "MessageLog"
Const SM_MESSAGE_LOG_COLUMNS                                = "MessageViewHistoryID, ActionTypeID,Actionvalue, ActionURL,SyncFlag,CreatedDate"
Const SM_MESSAGE_SYNC_LOG_TABLE_NAME                        = "MessageSyncLog"
Const SM_MESSAGE_SYNC_LOG_COLUMNS                           = "SyncURL, ResultTypeID, NumMessage,SyncFlag, CreatedDate"
Const SM_MESSAGE_VIEW_HISTORY_TABLE_NAME                    = "MessageViewHistory"
Const SM_MESSAGE_VIEW_HISTORY_COLUMNS                       = "MessageID,Category,Priority,SyncFlag,AcctID,Affid,LCID,ProductKey,createddate"
Const SM_CLIENTDB_OPERATION_OK                              = 1
Const SM_CLIENTDB_OPERATION_FAILED                          = 2
Const SM_CLIENTDB_ACCESS_DENIED                             = 3
Const SM_CLIENTDB_NO_MORE_DATA                              = 4
Const SM_CLIENTDB_DATA_FOR_CURRENT_ROW_READY                = 5

Const SM_PROVIDER_OBJECT_CREATION_SUCCESS                   = 1
Const SM_PROVIDER_OBJECT_CREATION_FAILED                    = 0
Const SM_GET_INSTALLED_APPS_SUCCESS                         = 1
Const SM_GET_INSTALLED_APPS_FAILED                          = 0
Const SM_APPINFO_OBJECT_CREATION_SUCCESS                    = 1
Const SM_APPINFO_OBJECT_CREATION_FAILED                     = 0
Const SM_SUBMGR_OBJECT_CREATION_SUCCESS                     = 1
Const SM_SUBMGR_OBJECT_CREATION_FAILED                      = 0
Const SM_SYSTEMMGR_PROVIDER                                 = "SysMgrObject"
Const SM_SYSMGR_OBJECT_CREATION_SUCCESS                     = 1
Const SM_SYSMGR_OBJECT_CREATION_FAILED                      = 0
Const SM_UIMGR_NETWORKDRIVES_EXIST                          = "Yes"
Const SM_UIMGR_NETWORKDRIVES_MISSING                        = "No"
Const SM_SYSMGR_ALLOWED_IDLE_TIME_IN_MSEC                    = 60000
Const SM_SYSMGR_SYSTEM_IDLE                                 = 0
Const SM_SYSMGR_SYSTEM_IN_USE                               = 1

Const SM_SYSTEM_IN_SCREEN_SAVER_MODE                        = true
Const SM_SYSTEM_NOT_IN_SCREEN_SAVER_MODE                    = false

Const SM_SYSTEM_IN_FULL_SCREEN_MODE                        = true
Const SM_SYSTEM_NOT_IN_FULL_SCREEN_MODE                    = false
Const SM_OS_IS_WIN8_OR_GREATER                             = True

Const SM_REGISTRY_OBJECT_CREATION_SUCCESS                   = 1
Const SM_REGISTRY_OBJECT_CREATION_FAILED                    = 0

Const SM_TASKSCHEDULER_OBJECT_CREATION_SUCCESS              = 1
Const SM_TASKSCHEDULER_OBJECT_CREATION_FAILED               = 0
Const SM_TASKSCHEDULER_TASK_CREATION_SUCCESS                = 1
Const SM_TASKSCHEDULER_TASK_CREATION_FAILED                 = 0

Const SM_SUBMGR_APPID                                       = "appid"
Const SM_SUBMGR_AFFID                                       = "affid"
Const SM_SUBMGR_APPCODE                                     = "app_code"
Const SM_SUBMGR_ACCTID                                      = "accnt_id"
Const SM_SUBMGR_PRODUCTKEY                                  = "product_key"
Const SM_SUBMGR_EXPIREDATE                                  = "settings"
Const SM_SUBMGR_EXTENDEDEXPIREDATE                          = "extended_expiry_dt"
Const SM_SUBMGR_PERPETUAL                                   = "perpetual"
Const SM_SUBMGR_TRIAL                                       = "trial"
Const SM_SUBMGR_BACKEND                                     = "backend"
Const SM_SUBMGR_SYNCURL                                     = "sync_url"
Const SM_SUBMGR_WEBSITEURL                                  = "website"
Const SM_SUBMGR_LANGCODE                                    = "applang"
Const SM_SUBMGR_PKGID                                       = "package_id"
Const SM_SUBMGR_PKGNAME                                     = "package_name"
Const SM_SUBMGR_RENEWALURL                                  = "renewalurl"
Const SM_SUBMGR_FLAGS                                       = "flags"
Const SM_SUBMGR_EMAIL                                       = "e_mail"
Const SM_SUBMGR_LICENSECOUNT                                = "num_licenses"
Const SM_SUBMGR_INSTALLDT                                   = "installed"

Const SM_TASKSCHEDULER_SHELL                                = "WScript.Shell"
Const SM_TASKSCHEDULER_NAME                                 = "Task1"
Const SM_TASKSCHEDULER_TRIGGER_FREQUENCY                    = "3600"
Const SM_TASKSCHEDULER_TASK_PROCESS_PATH                    = "c:\\My.exe"
Const SM_TASKSCHEDULER_TASK_PROCESS_PARAMS                  = "http://sm.mcafee.com/syncmessage.aspx"
Const SM_TASKSCHEDULER_ACTIVETIMEBIAS_PATH                  = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
Const SM_SUBINFO_TOTAL_KNOWN_APPDATA_PROP_COUNT		        = "1"
Const SM_SHOW_MESSAGE_LIMIT                                 = "1"
Const SM_MAINT_FLAG_ENABLED                                 = False
Const SM_DB_HIT_LIMIT                                       =   1

Const SM_UICONTAINER_DEFAULT_HEIGHT                         = "600"
Const SM_UICONTAINER_DEFAULT_WIDTH                          = "450"
Const SM_UICONTAINER_OBJECT_CREATION_SUCCESS                = 1
Const SM_UICONTAINER_OBJECT_CREATION_FAILED                 = 0
Const SM_UICONTAINER_CENTERED                               = 0
Const SM_UICONTAINER_BOTTOM_RIGHT_CORNER                    = 1
Const SM_UICONTAINER_MANUAL                                 = 2

Const SM_INSTRUMENTATION_OBJECT_CREATION_FAILED             = 0
Const SM_INSTRUMENTATION_OBJECT_CREATION_SUCCESS            = 1

Const SM_INSTRU_STATUS_FAILED				                = False
Const SM_INSTRU_STATUS_SUCCESS				                = True
Const SM_GETUTCDTTIME_FAILED				                = "0"
Const SM_GETUTCDTTIME_SUCCESS				                = "1"
Const SM_GETUTC_DATE_TIME_FN_RESPONSE_HANDLER		        = "GetServerUTCDtTime"
Const SM_GETUTC_DATE_TIME_FN_REQUEST_URL		            = "https://sm.mcafee.com/InstruCommunicator.asmx/GetUTCTime"
Const SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER		        = "GetInstruResponse"
Const SM_SEND_INSTRU_XML_FN_RESPONSE_TIMEOUT		        = 5000
Const SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME    	        = "InstrumentationXML="
Const SM_SEND_INSTRU_XML_FN_REQUEST_URL		    	        = "https://sm.mcafee.com/InstruCommunicator.asmx/SetReportingInstru"
Const SM_SEND_SMARTINSTRU_XML_FN_REQUEST_URL		    	        = "https://sm.mcafee.com/InstruCommunicator.asmx/SetSmartReportingInstru"

Const SM_MACHINE_CONFIG_OS                                  = "os="
Const SM_MACHINE_CONFIG_OS_VERSION_ID                       = " osv="
Const SM_MACHINE_CONFIG_MEM_SIZE_MB                         = " MemorySizeInMB="
Const SM_MACHINE_CONFIG_PROCESSOR_FMLY                      = " ProcessorFamily="
Const SM_MACHINE_CONFIG_PROCESSOR_VRSN                      = " ProcessorVersion="
Const SM_MACHINE_CONFIG_PROCESSOR_SPEED_GHZ                 = " ProcessorSpeedGHZ="
Const SM_MACHINE_CONFIG_NO_OF_PROCESSORS                    = " NumberOfProcessors="
Const SM_MACHINE_CONFIG_NETWORK_DRIVES = " NetworkDrives="
Const SM_MACHINE_CONFIG_CSP_CLIENT_ID = " CspClientId="
Const SM_MACHINE_CONFIG_HARDWARE_ID                         = " HardwareId="
Const SM_MACHINE_CONFIG_SOFTWARE_ID                         = " SoftwareId="

Const SM_CLIENT_CONFIG_APPID                                = "app_id="
Const SM_CLIENT_CONFIG_APPCODE                              = " app_code="
Const SM_CLIENT_CONFIG_PERPETUAL                            = " perpetual="
Const SM_CLIENT_CONFIG_TRIAL                                = " trial="
Const SM_CLIENT_CONFIG_ACCTID                               = " acct_id="
Const SM_CLIENT_CONFIG_EXP_DT                               = " expire_dt="
Const SM_CLIENT_CONFIG_BACKEND                              = " backend="
Const SM_CLIENT_CONFIG_SYNC_URL                             = " sync_url="
Const SM_CLIENT_CONFIG_WEB_SITE                             = " web_site="
Const SM_CLIENT_CONFIG_AFFID                                = " aff_id="
Const SM_CLIENT_CONFIG_BUILDID                              = " build_id="
Const SM_CLIENT_CONFIG_LANG_CODE                            = " lang_code="
Const SM_CLIENT_CONFIG_PKGID                                = " pkg_id="
Const SM_CLIENT_CONFIG_PKGNM                                = " pkg_name="
Const SM_CLIENT_CONFIG_RNWL_URL                             = " renewal_url="
Const SM_CLIENT_CONFIG_REG_STATUS                           = " reg_status="
Const SM_CLIENT_CONFIG_EMAIL                                = " e_mail="
Const SM_CLIENT_CONFIG_INSTALLED_DT                         = " installed_date="
Const SM_CLIENT_CONFIG_LIC_QTY                              = " lic_qty="
Const SM_CLIENT_CONFIG_PROD_KEY                             = " product_key="
Const SM_CLIENT_CONFIG_INSTALLED_DYS                        = " installed_days="
Const SM_CLIENT_CONFIG_ACTIVATION_DYS                       = " activation_days="
Const SM_CLIENT_CONFIG_ACTIVATION_DATE                       = " activation_date="
Const SM_CLIENT_CONFIG_OOBE                                 = " oobe="
Const SM_CLIENT_CONFIG_IPT                                  = " ipt="
Const SM_CLIENT_CONFIG_VERSIONID                            = " versionid="
Const SM_CLIENT_CONFIG_LOCALEID                             = " locale_id="
Const SM_CLIENT_CONFIG_EULA                                 = " eula="
Const SM_CLIENT_OOBE_CODE                                   = "3"
Const SM_CLIENT_CONFIG_FORCEUPGRADE                         = " force_upgrade=" 'Added for force upgrade 
Const SM_CLIENT_CONFIG_DAYSAFTERINSTALL                     = " days_after_install=" 'Added for SM Optimization 
Const SM_CLIENT_CONFIG_ACTIVATIONGRACEPERIOD                = " activation_grace_period=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_UNREGEXPIRED                         = " unreg_expired=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_UNREGEXPIREINDAYS                    = " unreg_expire_in_days=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_EXTENDEDEXPIRYDATE                   = " extended_expiry_dt=" ' Added for Windows 8 change
Const SM_CLIENT_CONFIG_DISABLEDDATE                         = " disabled_in_days=" ' Added for Windows 8 change
Const SM_CLIENT_CONFIG_SAFMESSAGE                           = " sa_message="        ' Site Advisor Freemium Message Applicable
Const SM_CLIENT_CONFIG_COHORTID                           = " cohort_definition_id="  

'=== Sub mgmt II constants
Const SM_PROP_PROVIDER_VERSION                              = "1"
Const SM_SUBMGMT2_OBJ_PROVIDER_VERSION                      = "2.0"
Const SM_PROP_PROVIDER_SUBMGMT2                             = "7"
Const SM_GET_SUBMGMT2_SUCCESS                               = 1
Const SM_GET_SUBMGMT2_FAILED                                = 0
''==== Object Creation
Const SM_OBJECT_CREATION_FAILED    			                = "0"
Const SM_OBJECT_CREATION_SUCCESS				            = "1"

''==== MSC IPT
Const SM_PROP_PROVIDER_MSCIPT                               = "8"
Const PLATFORMDSP                                           = "PlatformDspObj"
Const PRODUCT_KEY                                           = "pk"
Const IPTVALIDTILL                                          = "HKEY_CURRENT_USER\Software\MCAFEE\MSC\IPTValidTill"

'==== OOBE V3
Const SM_OOBE_V3_LOCATION       					        = "HKLM\Software\McAfee\OemInfo\Activation\OneClick\QueryParams"
Const SM_CLIENT_OOBE_V3_CODE                                = "8"
Const SM_CLIENT_OOBE_V3_VERSION                             = "OCVersion"
Const SM_CLIENT_OOBE_V3_ENABLE                              = "OCEnable"
Const SM_CLIENT_OOBE_V3_EULASTATE                           = "EULAState"
Const SM_CLIENT_OOBE_V3V4_EULADATE                          = "EulaDt"
Const SM_CLIENT_OOBE_V3_VERSION_NUMBER                      = "3.0"
Const SM_CLIENT_APPINFO_LOCATION                            = "HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams"
Const SM_CLIENT_OOBE_V3_IPRDONE                             = "IPRdone"

'===== AA: Site Advisor Freemium Project Variables
Const SM_CLIENT_SAFREEMIUM_LOCATION                         = "HKLM\SOFTWARE\McAfee\SiteAdvisor"
Const SM_CLIENT_SAF_VERSION                                 = "FullVersion"
Const SM_CLIENT_SAINSTALL_FOLDER                            = "Install Dir"
Const SM_CLIENT_SAFINSTALLER_FILE                           = "saUi.exe"
Const SM_CLIENT_SAFUICNTARGS                                = " /saFreemium /postinstallpage"
Const SM_CLIENT_SAFSMOFFER_STATUS                           = "IsSAFreemiumOfferAccepted"
'===== AA: Site Advisor Freemium Project Variables end

Const SM_CLIENT_MSC_LOCATION                                = "HKLM\SOFTWARE\McAfee\MSC"
 

'==== Smart Messaging Optimization
Const SM_CLIENT_UICNTPATH_LOCATION                          = "HKLM\SOFTWARE\McAfee\MSC\OEM\Manage\RGW"
Const SM_CLIENT_UI_CONTAINER_PATH                           = "UICntPath"

''=== Virus Alert
Const REG_PATH_MSC_SECURITYREPORTS = "HKLM\Software\McAfee\MSC\Settings\Securityreports"
Const REG_KEY_THREATMETER = "threatMeter"

'==== Windows 8 Metro
Const REG_KEY_DISABLEDDATE                                  = "disabled_after_uac_date"

Const SM_PROP_METROTOASTYPE                                 = "toasttype"
Const SM_PROP_METROMSGHEADER                                = "msgheader"
Const SM_PROP_METROMSGBODY                                  = "msgbody"
Const SM_PROP_ISMETROVALID                                  = "isvalid"

Const SM_SUBMGR_METRODISPLAYMODE_UNKNOWN = 0
Const SM_SUBMGR_METRODISPLAYMODE_DESKTOP = 1
Const SM_SUBMGR_METRODISPLAYMODE_METRO = 2
Const SM_SUBMGR_METRODISPLAYMODE_STARTUP = 3

Const SM_SUBMGR_METROCLOSEREASON_UNDEFINED = -1
Const SM_SUBMGR_METROCLOSEREASON_CLICKED = 0
Const SM_SUBMGR_METROCLOSEREASON_DISMISSED = 1
Const SM_SUBMGR_METROCLOSEREASON_TIMEOUT = 2
Const SM_SUBMGR_METROCLOSEREASON_ERROR = 3

Const SM_METRO_TEMPLATE_UNDEFINED = -1
Const SM_METRO_TEMPLATE_TOASTTEXT01 = 4
Const SM_METRO_TEMPLATE_TOASTTEXT02 = 5

' A/B Testing before the A/B Framework was implemented.
Const SM_CLIENT_MESSAGEID_LOCATION                         = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation" 	'
Const SM_CLIENT_MESSAGEID_KEY		                       = "MessageID"
Const SM_CLIENT_INPRODUCT_LOCATION                         = "HKLM\SOFTWARE\McAfee\MSC\Settings\InProductTransaction"
Const SM_CLIENT_IPT_ENABLE_KEY                             = "enable"
Const SM_CLIENT_IPT_LCID                                   = "lcid"
Const SM_CLIENT_APPINFO_LOCATION_SUBSTITUTE                = "HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"
Const SM_CLIENT_IPT_BUILD                                  = "build"
Const SM_CLIENT_INPRODUCT_UPDATE_VERSION                   = "UpdateVersion"
Const SM_CLIENT_INPRODUCT_SUBSTATUS                        = "SubStatus"
Const SM_CLIENT_INPRODUCT_PARENTAFFID                      = "ParentAffId"

' Cohort registry

Const SM_CLIENT_COHORT_SM_LOCATION = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation\cohort\smid"
Const SM_CLIENT_COHORT_RULE_LOCATION = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation\cohort\cohortapplied"
Const SM_CLIENT_COHORT_ID = "cohortexecuted"


    Dim oXMLDoc, oXMLHTTP
    Dim m_UTCDateTime, m_InstruStatus
 '=== This will get the size of the Providerobject array
    Function Length(ByRef vArray)
        Dim intIndex
        Dim intCount
    
        intCount = 0
        For intIndex = 0 To UBound(vArray)
            If Not(IsNull(vArray(intIndex))) And Not(IsEmpty(vArray(intIndex)))Then
                If Not(vArray(intIndex) Is Nothing) Then
                    intCount = intCount + 1
                End If    
            End If
        Next
        Length = intCount
    End Function
    '=== This will be used to show/hide the msgbox.
    '=== it will be used 
    Function ShowErrorMessage(strErrorMessage)
        Dim blnShowMsg
        blnShowMsg  = SHOW_SM_ERROR_MESSAGES
        If CBool(blnShowMsg) Then
            MsgBox(strErrorMessage)
        End If
    End Function
            
    Function HideWindow()
        Dim objExternal,objUICont
	    Set objExternal = window.external
	    Set objUICont = objExternal.GetParam ("McUIContainer")
	    Call objUICont.Hide() 
		Call objUICont.Terminate()
		'window.close
    End Function
    
    Function GetInternetConnectedState(objProviderEnum)
        Dim intResult
        Dim bInternetState
        Dim objSMSystemData
        
        Set objSMSystemData = New CSMSystemData
        bInternetState    = false
        
        If Not(objProviderEnum Is Nothing) Then
            If  Not(objSMSystemData Is Nothing) Then
                intResult   = objSMSystemData.InitializeSystemObject(objProviderEnum)
                If intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then
                    bInternetState = objSMSystemData.IsConnectedToInternet
                Else    
                    ShowErrorMessage "GetInternetConnectedState:objSMSystemData initialization failed"
	            End If
	        Else
	            ShowErrorMessage "GetInternetConnectedState:objSMSystemData is  empty"  
            End If
        Else
            ShowErrorMessage "GetInternetConnectedState:objProviderEnum is  empty"  
        End If
        
        Set objSMSystemData = Nothing
        
        GetInternetConnectedState = bInternetState
     End Function
    	
    Function MakeHTTPRequest(RequestURL, ResponseHandlerFunc, strEnvelope)
        On Error Resume Next
		set oXMLHTTP = createobject("msxml2.xmlhttp.3.0")
        set oXMLDoc = createobject("msxml2.domdocument")
        
        ShowErrorMessage("strEnvelope = " & strEnvelope)
        ShowErrorMessage("Calling WebService To Get UTC Time")
        oXMLHTTP.onreadystatechange = getRef(ResponseHandlerFunc)
       
        If Len(Trim(RequestURL)) > 0 Then
            call oXMLHTTP.open("POST",RequestURL,false)
            call oXMLHTTP.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
            If Len(Trim(strEnvelope)) > 0 Then
               setTimeout "AbortRequest(oXMLHTTP)",SM_SEND_INSTRU_XML_FN_RESPONSE_TIMEOUT
            End If
            call oXMLHTTP.send(strEnvelope)
            
        End If
    End Function
    
    Function AbortRequest(oXMLHTTP)
        ShowErrorMessage("Before Aborting")
        oXMLHTTP.abort '=== To abort the httprequest if it takes long time to respond
        m_InstruStatus = SM_INSTRU_STATUS_FAILED
        ShowErrorMessage("After Aborting")
    End Function

    Function SendInstruXML(strInstruXML)
         m_InstruStatus = SM_INSTRU_STATUS_FAILED
        dim strEnvelope : strEnvelope = SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME & strInstruXML
        Call MakeHTTPRequest(SM_SEND_INSTRU_XML_FN_REQUEST_URL, SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER, strEnvelope)
		SendInstruXML = m_InstruStatus
    End Function

    Function SendSmartInstruXML(strInstruXML)
         m_InstruStatus = SM_INSTRU_STATUS_FAILED
        dim strEnvelope : strEnvelope = SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME & strInstruXML
        Call MakeHTTPRequest(SM_SEND_SMARTINSTRU_XML_FN_REQUEST_URL, SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER, strEnvelope)
		SendSmartInstruXML = m_InstruStatus
    End Function

    Function GetInstruResponse()
        If(oXMLHTTP.readyState = 4) Then
            dim szResponse: szResponse = oXMLHTTP.responseText
            call oXMLDoc.loadXML(szResponse)
            If(oXMLDoc.parseError.errorCode <> 0) Then
                call ShowErrorMessage(oXMLDoc.parseError.reason)
		        m_InstruStatus = SM_INSTRU_STATUS_FAILED
            else
                m_InstruStatus = oXMLDoc.getElementsByTagName("boolean")(0).childNodes(0).text
                call ShowErrorMessage(oXMLDoc.getElementsByTagName("boolean")(0).childNodes(0).text)
            End If
        End If
	
    End Function
    
    Function GetUTCDateTime()
        Dim RequestURL
        '===RequestURL = "https://sm.mcafee.com/smartservice/smartmessaginginstrumentation.asmx/GetUTCTime"
        Call MakeHTTPRequest(SM_GETUTC_DATE_TIME_FN_REQUEST_URL, SM_GETUTC_DATE_TIME_FN_RESPONSE_HANDLER, "")
        GetUTCDateTime = m_UTCDateTime
    End Function

    Sub GetServerUTCDtTime()
        If(oXMLHTTP.readyState = 4) Then
            dim szResponse: szResponse = oXMLHTTP.responseText
            call oXMLDoc.loadXML(szResponse)
            If(oXMLDoc.parseError.errorCode <> 0) Then
                call ShowErrorMessage(oXMLDoc.parseError.reason)
		m_UTCDateTime = SM_GETUTCDTTIME_FAILED
            else
                m_UTCDateTime = oXMLDoc.getElementsByTagName("string")(0).childNodes(0).text
                call ShowErrorMessage(oXMLDoc.getElementsByTagName("string")(0).childNodes(0).text)
            End If
        End If
    End Sub

   Private Function GetQueryStringValueFromURL(strActionURL,qrystr)
        Dim intCIDPos, qryVal
        Dim intEqualPos, intAmpPos
        Dim intUrlLen
        qryVal = ""
        ShowErrorMessage "getcid"
        '===temporary
       '' Dim testurl: testurl = "http://home.mcafee.com/root/campaign.aspx?cid=45856&abc=1"
        If Len(Trim(strActionURL)) > 0 Then
            intCIDPos = Instr(1,strActionURL,qrystr) '=== get the cid position
            If CInt(intCIDPos) > 0 Then
                intEqualPos = Instr(intCIDPos,strActionURL,"=") '=== get the position of equal to
                If intEqualPos > 0 Then
                    If Instr(intEqualPos+1, strActionURL, "&", 1) > 0 Then 
                        intAmpPos = Instr(intEqualPos, strActionURL, "&", 1)    
                        qryVal = Mid(strActionURL, intEqualPos+1, (intAmpPos - (intEqualPos+1))) ' retrieve the variable value
                    Else
                        intUrlLen = Len(strActionURL)
                        qryVal = Mid(strActionURL, intEqualPos+1, (intUrlLen - (intEqualPos))) ' retrieve the variable value                            
                    End If    
                End If
            End If
        End If
        
        GetQueryStringValueFromURL = qryVal
    End Function
    
    Function GetCurrentDate(strServerUTCTime)
        Dim strFinalCurrentDate
        Dim strDay, strMonth, strYr
        
        If Len(Trim(strServerUTCTime)) > 0 Then
            strFinalCurrentDate = strServerUTCTime
        Else
            strFinalCurrentDate = Date
        End If
        
        
        strDay = FormatTwoDigit(Day(strFinalCurrentDate))
        strMonth = FormatTwoDigit(Month(strFinalCurrentDate))
        strYr = Year(strFinalCurrentDate)
        
    
        strFinalCurrentDate = strMonth & "/" & strDay & "/" &  strYr
    
        GetCurrentDate = strFinalCurrentDate
    End Function
    
    Function FormatTwoDigit(strValue)
        If Len(Trim(strValue)) = 1 Then
            strValue = "0" & strValue
        End If    
        FormatTwoDigit = strValue   
    End Function

    Function GetBrowserVersion()
        Dim regEx, userAgent, CurrentMatches
		Set regEx = new RegExp
		userAgent = Navigator.UserAgent 
		userAgent = Mid(userAgent, InStr(userAgent, "MSIE ") + 5)
		regEx.Pattern = "(\d+\.\d+)"
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.MultiLine = True
		Set CurrentMatches = regEx.Execute(userAgent)
		If CurrentMatches.Count > 0 Then
			If IsNumeric(CurrentMatches(0)) Then
				GetBrowserVersion = CDbl(CurrentMatches(0))
				Exit Function
			End If
		End If
		GetBrowserVersion = 6
    End Function

    Function GetIptValidTill()
        Dim l_ObjShell
        Dim l_IptValidTill
        set l_ObjShell = CreateObject(SM_TASKSCHEDULER_SHELL) 
        On Error Resume Next
        l_IptValidTill = l_ObjShell.RegRead(IPTVALIDTILL) 
        If Err.Number <> 0 Then
            Err.Clear
            GetIptValidTill = ""
        Else
            GetIptValidTill = l_IptValidTill
        End If
     End Function


Class CSMProviderEnum

Private m_ProvArray()
Private m_objExternal
Private m_objProvEnum

    Private Sub Class_Initialize
        On Error Resume Next 
    
	    Dim objProv
	    Dim ProviderArrayLen
	    
	    Set m_objExternal   = Nothing
	    Set m_objProvEnum   = Nothing
	    Set objProv         = Nothing
	    
	    ReDim m_ProvArray(0)
	    
	    ProviderArrayLen    = 0
	    
	    If m_objExternal Is Nothing Then
	     	If IsObject( window.external ) And Not(window.external Is Nothing) Then
	            Set m_objExternal = window.external
	        Else
	            ShowErrorMessage("SMProviderEnum:Unable to get the window.external")  
	        End if
	    End If 
	    
	    If IsObject(m_objExternal) And Not m_objExternal Is Nothing Then
            Set m_objProvEnum = m_objExternal.GetParam("SMProviderEnum")
	    End If
	    
	    If IsObject(m_objProvEnum) And Not m_objProvEnum Is Nothing Then
	        m_objProvEnum.Reset
	        
	        Do
                Set objProv = Nothing
                Set objProv = m_objProvEnum.Next()
                
                If Not (objProv Is Nothing) Then
                    ProviderArrayLen = Length(m_ProvArray)
                    If ProviderArrayLen > 0 Then
	                    ReDim Preserve m_ProvArray(ProviderArrayLen) 
	                End if
	                Set m_ProvArray(ProviderArrayLen) = objProv  
	            End If
            Loop Until objProv Is Nothing
	        
            If Err.number > 0 Then
               If Length(m_ProvArray) = 0 Then
                ShowErrorMessage("SMProviderEnum Initialize: " & Err.Description) 
               End If
            End If


''''	        For Each objProv In m_objProvEnum
''''	            If objProv.IsEnabled()Then
''''	                ProviderArrayLen = Length(m_ProvArray)
''''	                If ProviderArrayLen > 0 Then
''''	                    ReDim Preserve m_ProvArray(ProviderArrayLen) 
''''	                End if
''''	                 m_ProvArray(ProviderArrayLen) = objProv       
''''	                
''''	            End If
''''	        Next
        Else
            ShowErrorMessage("SMProviderEnum:Unable to get the Provider Enum object")
	    End If
	End Sub

	Public Property Get ProviderSize()
		ProviderSize = Length(m_ProvArray)
	End Property
	
	Public Property Get ProviderEnumObject()
	    Set ProviderEnumObject = m_objProvEnum    
	End property
	
	Public Function GetProviderObject(ByRef intRetVal, intObjNumber)
		    On Error Resume Next
		    
		    Dim objNxt
		    Set objNxt = Nothing
		    
		    If Length(m_ProvArray) > 0 Then
		        Set objNxt      = m_ProvArray(intObjNumber)
		        intRetVal       = SM_PROVIDER_OBJECT_CREATION_SUCCESS
		    Else
		        ShowErrorMessage("SMProviderEnum:Unable to get the Provider object")
		        intRetVal = SM_PROVIDER_OBJECT_CREATION_FAILED    
		    End If
		    
		    Set GetProviderObject = objNxt
	End Function
	
   Private Sub Class_Terminate
	
    	If Not IsEmpty(m_objProvEnum) Then
    	    If Not m_objProvEnum Is Nothing Then
	            set m_objProvEnum = nothing
	        End If
	    End If
	
	    If Not IsEmpty(m_objExternal) Then
	        If Not m_objExternal Is Nothing Then
	            set m_objExternal = nothing
	        End If
	    End If
	
	End Sub


End Class

' CSMAppData 'inherits' from CSMDATA and is used for the AppData information 
''
Class CSMAppData

    Private m_arrInstalledApps()
	Private m_objAppMgr
	
    Private Sub Class_Initialize
        ReDim m_arrInstalledApps(0)
		Set m_objAppMgr = Nothing
	End Sub
	
	Public Property Get InstalledApplicationSize
        Dim intInstalledAppCount
        intInstalledAppCount =  Length(m_arrInstalledApps)
        If intInstalledAppCount > 0 Then
            InstalledApplicationSize = intInstalledAppCount   
        Else
            InstalledApplicationSize = 0
        End If
	End Property	

    Public Function InitializeAppMgrObject(objProvider)
	    On Error Resume Next
	   
	   Dim intResult
	   
	   intResult = 0
	   
	    If IsObject(objProvider) And Not(objProvider Is Nothing) Then   
	        
	        If m_objAppMgr Is Nothing Then
	            Set m_objAppMgr = objProvider.GetProperty(SM_PROP_PROVIDER_APPLICATION_MGR)
	            
	            If IsObject(m_objAppMgr) And Not(m_objAppMgr Is Nothing) Then
	                ShowErrorMessage("SMAppData:Application Manager object created")
	            Else
	                intResult = SM_GET_INSTALLED_APPS_FAILED
	                ShowErrorMessage("SMAppData:Unable to get the Application Manager object1")
	            End if
	        End If
	        
	        If IsObject(m_objAppMgr) And Not(m_objAppMgr Is Nothing) Then
	            Call SetInstalledAppsInArray(m_objAppMgr)
	            
	            If Length(m_arrInstalledApps) > 0 then
	              intResult = SM_GET_INSTALLED_APPS_SUCCESS 
	              ShowErrorMessage("SMAppData:Created the Installed Application array")
	            Else
	              intResult = SM_GET_INSTALLED_APPS_FAILED
	              ShowErrorMessage("SMAppData:Unable to get the Installed Application array")
	            End If
	        Else
	            intResult = SM_GET_INSTALLED_APPS_FAILED    
	            ShowErrorMessage("SMAppData:Unable to get the Application Manager object2")
	        End If
	    
	    Else
	        intResult = SM_GET_INSTALLED_APPS_FAILED
	        ShowErrorMessage("SMAppData:Unable to get the Provider object")
	    End If
	
	    InitializeAppMgrObject = intResult 
	End Function
	
	Public Function GetAppInfoObject(ByRef intRetVal,intAppInfoObjectNumber)
        On Error Resume Next	
        
        Dim objAppInfo
        Set objAppInfo = Nothing
        
        If Length(m_arrInstalledApps) > 0 Then
            Set objAppInfo = m_arrInstalledApps(intAppInfoObjectNumber)
            intRetVal = SM_APPINFO_OBJECT_CREATION_SUCCESS
            GetAppInfoObject = objAppInfo
        Else
             ShowErrorMessage("SMApps:Unable to get the AppInfo object")
            intRetVal = SM_APPINFO_OBJECT_CREATION_FAILED
        End If
        
	    Set GetAppInfoObject = objAppInfo    
	End Function
	
	Private Sub SetInstalledAppsInArray(objAppMgr)
	    On Error Resume Next
	    
	    Dim objAppInfo
	    Dim arrAppInfo()
	    Dim NoOfInstalledApps
	    Set objAppInfo = Nothing
	    ReDim arrAppInfo(0)
	    If IsObject(objAppMgr) And Not objAppMgr Is Nothing Then
	        objAppMgr.Reset
	        
	        Do
                Set objAppInfo = Nothing
                Set objAppInfo = objAppMgr.Next()
                NoOfInstalledApps = Length(m_arrInstalledApps)
	            If NoOfInstalledApps > 0 Then
	                ReDim Preserve m_arrInstalledApps(NoOfInstalledApps) 
	            End if
	            Set m_arrInstalledApps(NoOfInstalledApps) = objAppInfo  
            Loop Until objAppInfo Is Nothing
	        
            If Err.number > 0 Then
               If Length(m_arrInstalledApps) = 0 Then
                ShowErrorMessage("SMAppData GetInstalledApps: " & Err.Description) 
               End If
            End If
        End If
       
	End Sub

	Private Sub Class_Terminate
		If Not m_objAppMgr Is Nothing Then
			Set m_objAppMgr = Nothing
		End If
	End Sub

End Class


Class CSMSubscriptionData

    Private m_objSubMgr
    Private m_objSubInfo
    Private m_strBuildID 
    Private intReturn
    
    Private Sub Class_Initialize
        Set m_objSubMgr     = Nothing
		Set m_objSubInfo    = Nothing
	End Sub
	
	Public Property Get AppId()
	    Dim strAppId
	    
        strAppId = m_objSubInfo.GetProperty(SM_SUBMGR_APPID, intReturn)    
        
        If Len(Trim(strAppId)) > 0 Then
            AppId      = strAppId
        Else
            AppId = ""
        End If
    End Property

    Public Property Get AppCode()
        Dim strAppCode
        
        strAppCode = m_objSubInfo.GetProperty(SM_SUBMGR_APPCODE, intReturn)    
        
        If Len(Trim(strAppCode)) > 0 Then
            AppCode         = strAppCode
        Else
            AppCode = ""
        End If
    End Property

    Public Property Get PerpetualInfo()
        Dim strPerpetual
        
        strPerpetual = m_objSubInfo.GetProperty(SM_SUBMGR_PERPETUAL, intReturn)
        
        If Len(Trim(strPerpetual)) > 0 Then
            PerpetualInfo = strPerpetual
        Else
            PerpetualInfo = ""
        End If
    End Property
    
    Public Property Get TrialInfo()
        Dim strTrial
        
        strTrial = m_objSubInfo.GetProperty(SM_SUBMGR_TRIAL, intReturn)    
        
        If Len(Trim(strTrial)) > 0 Then
            TrialInfo = strTrial
        Else
            TrialInfo = ""
        End If
    End Property
    
    Public Property Get AccountID()
        Dim strAcctId
        strAcctId = m_objSubInfo.GetProperty(SM_SUBMGR_ACCTID, intReturn)    
        
        If Len(Trim(strAcctId)) > 0 Then
            AccountID = strAcctId
        Else
            AccountID = ""
        End If
        
        '=== temporary
        'AccountID = "976431"
        
    End Property
    
    Public Property Get ExpiryDate()
        Dim strExpDate,strDay, strMonth, strYear 
        
        strExpDate  = m_objSubInfo.GetProperty(SM_SUBMGR_EXPIREDATE, intReturn)    
        strExpDate  = GetCorrectDate(strExpDate)
        
        If Len(Trim(strExpDate)) > 0 Then
            ExpiryDate = strExpDate
        Else
            ExpiryDate = ""
        End If
    End Property

    Public Property Get ExtendedExpiryDate()
        Dim strExtendedExpiryDate

        strExtendedExpiryDate = m_objSubInfo.GetProperty(SM_SUBMGR_EXTENDEDEXPIREDATE, intReturn)
        strExtendedExpiryDate = GetCorrectDate(strExtendedExpiryDate)

        If Len(Trim(strExtendedExpiryDate)) > 0 Then
            ExtendedExpiryDate = strExtendedExpiryDate
        Else
            ExtendedExpiryDate = ""
        End If
    End Property
  
    Public Property Get Backend()
        Dim strBackend
        strBackend = m_objSubInfo.GetProperty(SM_SUBMGR_BACKEND, intReturn)    
        
        If Len(Trim(strBackend)) > 0 Then
            Backend = strBackend
        Else
            Backend = ""
        End If
    End Property

    Public Property Get SyncURL()
        Dim strSyncURL
        strSyncURL = m_objSubInfo.GetProperty(SM_SUBMGR_SYNCURL, intReturn)    
        
        If Len(Trim(strSyncURL)) > 0 Then
            SyncURL = strSyncURL
        Else
            SyncURL = ""
        End If
    End Property

    Public Property Get WebsiteURL()
        Dim strWebsiteURL
        strWebsiteURL = m_objSubInfo.GetProperty(SM_SUBMGR_WEBSITEURL, intReturn)    
        
        If Len(Trim(strWebsiteURL)) > 0 Then
            WebsiteURL = strWebsiteURL
        Else
            WebsiteURL = ""
        End If
    End Property

    Public Property Get AffID()
        Dim strAffId, strAff
        
        strAffId = m_objSubInfo.GetProperty(SM_SUBMGR_AFFID, intReturn)    
        If Len(Trim(strAffId)) > 0 Then
            If Instr(strAffid, "-") > 0  Then
                strAff = Mid(strAffid, 1, (Instr(strAffid,"-") -1))
                m_strBuildID = Mid(strAffid, (Instr(strAffid,"-") + 1), (Len(strAffid) - Instr(strAffid,"-")) )
            Else
                strAff  = strAffId  
            End If
            AffID = strAff
        Else
            AffID = ""
        End If        
        
    End Property

     '=================================================================
    'Acer SKU Swap
    '=================================================================
    Public Property Let AffID(value)
        Dim intResult
        intResult = m_objSubInfo.SetProperty(SM_SUBMGR_AFFID,value)
        intResult = m_objSubInfo.Commit()
    End Property
    
    '=================================================================
        
    Public Property Get BuildID()
        If Len(Trim(m_strBuildID)) > 0 Then
            BuildID = m_strBuildID
        Else
            BuildID = ""
        End If
    End Property

    Public Property Get LangCode()
        Dim strLangCode
        
        strLangCode = m_objSubInfo.GetProperty(SM_SUBMGR_LANGCODE, intReturn)    
        
        If Len(Trim(strLangCode)) > 0 Then
            LangCode = strLangCode
        Else
            LangCode = ""
        End If        
    End Property

    Public Property Get PkgID()
        Dim strPackageID
        strPackageID = m_objSubInfo.GetProperty(SM_SUBMGR_PKGID, intReturn)    
        
        If Len(Trim(strPackageID)) > 0 Then
            PkgID = strPackageID
        Else
            PkgID = ""
        End If
        
       ' PkgID = "274"
    End Property

    '==========================
    'Acer SKU Swap PkgID - Let 
    '==========================
     Public Property Let PkgID(value)
        Dim intResult
        intResult = m_objSubInfo.SetProperty(SM_SUBMGR_PKGID,value)
        intResult = m_objSubInfo.Commit()
    End Property



    Public Property Get PkgName()
        Dim strPkgName
        
        strPkgName = m_objSubInfo.GetProperty(SM_SUBMGR_PKGNAME, intReturn)    
        
        If Len(Trim(strPkgName)) > 0 Then
            PkgName = strPkgName
        Else
            PkgName = ""
        End If        
    End Property

	 Public Property Get RenewalURL()
        Dim strRenewalURL
        strRenewalURL = m_objSubInfo.GetProperty(SM_SUBMGR_RENEWALURL, intReturn)    
        
        If Len(Trim(strRenewalURL)) > 0 Then
            RenewalURL = strRenewalURL
        Else
            RenewalURL = ""
        End If
    End Property

    Public Property Get AppFlags()
        Dim strFlags
        
        strFlags = m_objSubInfo.GetProperty(SM_SUBMGR_FLAGS, intReturn)    
        If Len(Trim(strFlags)) > 0 Then
            If (1 And Clng(strFlags)) = 1 Then
                strFlags = "1" '=== Means Un-Registered
            Else
                strFlags = "0" '=== Means Registered
            End If
            
            AppFlags = strFlags
        Else
            AppFlags = "0" '=== Means Registered
        End If        
    End Property
	
	 Public Property Get Email()
        Dim strEmail
        
        strEmail = m_objSubInfo.GetProperty(SM_SUBMGR_EMAIL, intReturn)    
        
        If Len(Trim(strEmail)) > 0 Then
            Email = strEmail
        Else
            Email = ""
        End If        
    End Property

	 Public Property Get InstalledDate()
        Dim strInstalledDate
       
        strInstalledDate    = m_objSubInfo.GetProperty(SM_SUBMGR_INSTALLDT, intReturn)    
        strInstalledDate    = GetCorrectDate(strInstalledDate)
        
        If Len(Trim(strInstalledDate)) > 0 Then
            InstalledDate = strInstalledDate
        Else
            InstalledDate = ""
        End If        
    End Property

	 Public Property Get NumberOfLicenses()
        Dim strNumberOfLicense
        strNumberOfLicense = m_objSubInfo.GetProperty(SM_SUBMGR_LICENSECOUNT, intReturn)    
        
        If Len(Trim(strNumberOfLicense)) > 0 Then
            NumberOfLicenses = strNumberOfLicense
        Else
            NumberOfLicenses = "0"
        End If
    End Property

    Public Property Get ProductKey()
        Dim strPrdtKey
        strPrdtKey = m_objSubInfo.GetProperty(SM_SUBMGR_PRODUCTKEY, intReturn)    
        
        If Len(Trim(strPrdtKey)) > 0 Then
            ProductKey = strPrdtKey
        Else
            ProductKey = ""
        End If
    End Property    
    
    Public Property Get InstalledDays()
        Dim strInstalledDate
        Dim dtGracePeriodDate, dtCurrentDate
        Dim strInstalledDays
        
        strInstalledDate = InstalledDate 
        
        If Len(Trim(strInstalledDate)) > 0 Then
            Dim strServerDateTime, arrServerDateTime
            strServerDateTime = document.forms(0).hdnCurrentUTCTime.Value
            If strServerDateTime <> "" Then
                arrServerDateTime = Split(strServerDateTime, " ")
                dtCurrentDate = CDate(arrServerDateTime(0))
            Else
                dtCurrentDate = Date
            End If
            
            dtGracePeriodDate = CDate(strInstalledDate) + 15   
            If dtCurrentDate <= dtGracePeriodDate Then
               strInstalledDays = "1"
            Else
               strInstalledDays = "0"
            End If
        Else
            strInstalledDays = "0"
        End If   
        InstalledDays = strInstalledDays
    End Property

    'Added for Smart Messaging Optimization 
    Public Property Get DaysAfterInstall()
        DaysAfterInstall = GetDaysAfterInstall()
    End Property
    'End Smart Messaging Optimization 

    'Added for SM OOBE V3 changes
    Public Property Get ActivationGracePeriod()
        Dim strInstalledDate
        Dim dtCurrentDate, dtActGracePeriod
        Dim intDayDiff
        Dim strGracePeriod

        strInstalledDate = InstalledDate

        dtCurrentDate = Date
        dtActGracePeriod = CDate(strInstalledDate)

        intDayDiff = DateDiff("d",dtActGracePeriod, dtCurrentDate)   

        If intDayDiff > 5 Then
            strGracePeriod = "0" 'Beyond Activation Grace Period
        Else
            strGracePeriod = "1" ' Within Activation Grace Period
        End If

        ActivationGracePeriod = strGracePeriod
    End Property

   
   
    Public Function InitializeSubMgrObject(objProvider, strAppId, strAppCode)
      
        On Error Resume Next

        Dim intResult, retSubInfo
        Dim objIntermediateSubMgr
        
        Set objIntermediateSubMgr = Nothing
        
        intResult = 0

        If Not(objProvider Is Nothing) Then   
            If m_objSubMgr Is Nothing Then
                Set m_objSubMgr = objProvider.GetProperty(SM_PROP_PROVIDER_SUBSCRIPTION_MGR)
            End If
        Else
           intResult = SM_GET_INSTALLED_APPS_FAILED
            ShowErrorMessage("SMSubscriptionData:Unable to get the Provider object")
        End If
            
        If IsObject(m_objSubMgr) And Not(m_objSubMgr Is Nothing) Then
            If (Len(Trim(strAppId))> 0) Then  
                
                If (Len(Trim(strAppCode)) > 0) Then
                    strAppCode = ""
                End If
            
                Set objIntermediateSubMgr = m_objSubMgr.GetProperty(MC_PROP_SUB_MGR_OBJ)  
                'If objIntermediateSubMgr Not Is Nothing Then
                    Set m_objSubInfo = objIntermediateSubMgr.GetSubInfo(strAppId, strAppCode, retSubInfo)
                'Else
                'intResult = SM_GET_INSTALLED_APPS_FAILED    
                'ShowErrorMessage("SMAppData:Unable to get either strAppID or strAppCode ") 
                'End If
            Else
                intResult = SM_GET_INSTALLED_APPS_FAILED    
                ShowErrorMessage("SMSubscriptionData:Unable to get either strAppID or strAppCode ") 
            End If
        Else
            intResult = SM_GET_INSTALLED_APPS_FAILED    
            ShowErrorMessage("SMSubscriptionData:Unable to get the Application Manager object2")
        End If
        
        If IsObject(m_objSubInfo) And Not(m_objSubInfo Is Nothing) Then
        intResult = SM_SUBMGR_OBJECT_CREATION_SUCCESS    
        ShowErrorMessage("SMSubscriptionData:SubInfo object is successfully created")      
        End If

        If Err.number > 0 Then
        If intResult = SM_GET_INSTALLED_APPS_FAILED  Then
        ShowErrorMessage("SMSubscriptionData:SubInfo object: " & Err.Description)     
        End If
        End If

        Set objIntermediateSubMgr = Nothing
        InitializeSubMgrObject = intResult 
    End Function
	
    Private Function GetCorrectDate(strDate)
        Dim strFinalDate,strDay, strMonth, strYear 

        If Len(Trim(strDate)) > 0 Then
            strYear     = Mid(strDate,1,4)
            strMonth    = Mid(strDate,5,2)
            strDay      = Mid(strDate,7,2)
            If (strMonth <> "") And (strDay <> "") And (strYear <> "") Then
                If Len(strMonth) <> 2 Then
                    strMonth = "0" & strMonth
                End If
                If Len(strDay) <> 2 Then
                    strDay = "0" & strDay
                End If
                strFinalDate  = strMonth & "/" & strDay & "/" & strYear
            Else
                strFinalDate    = ""
            End If
        Else
        strFinalDate    = ""
        End If

        GetCorrectDate = strFinalDate
    End Function

    'Added for Smart Messaging Optimization 
    Private Function GetDaysAfterInstall() 
        Dim dtInstalledDate, dtCurrentDate
        Dim intDayDiff
        Dim strInstalledDays
        
        dtInstalledDate = InstalledDate 
        
        If Len(Trim(dtInstalledDate)) > 0 Then
            dtCurrentDate = Date
            intDayDiff = DateDiff("d",CDate(dtInstalledDate), dtCurrentDate)    
        End If   
        GetDaysAfterInstall = intDayDiff
    End Function
    'End of Smart Messaging Optimization 

    Private Sub Class_Terminate
		If Not IsEmpty(m_objSubMgr) Then
		    If Not m_objSubMgr Is Nothing Then
			    Set m_objSubMgr = Nothing
		    End If
		End If
		If Not IsEmpty(m_objSubInfo) Then
		    If Not m_objSubInfo Is Nothing Then
		        Set m_objSubInfo = Nothing
		    End If
		End If
	End Sub
	
End Class



Class CSMUIContainer
    
    Private m_ObjExternal, m_ObjUIContainer
    Private m_ObjUIContainerEnum

    
    
    Private Sub Class_Initialize
        Set m_ObjExternal           = window.external
        Set m_ObjUIContainer        = Nothing
        Set m_ObjUIContainerEnum    = Nothing
   	End Sub

'    Public Function DisplayUI(UIHeight, UIWidth)
'        Dim objExternal, objUICont
'        Set objUICont   = Nothing
'        Set objExternal = Nothing
'        Set objExternal = window.external

'        If Len(Trim(UIHeight)) = 0 Then
'            UIHeight = SM_UICONTAINER_DEFAULT_HEIGHT
'        End If
'        
'        If Len(Trim(UIWidth)) = 0 Then
'            UIWidth = SM_UICONTAINER_DEFAULT_WIDTH
'        End If
'        
'        If Not IsEmpty(m_ObjExternal) And Not m_ObjExternal Is Nothing Then
'            Set m_ObjUIContainer = m_ObjExternal.GetParam ("McUIContainer")
'            If Not IsEmpty(m_ObjUIContainer) And Not m_ObjUIContainer Is Nothing Then
'                Call m_ObjUIContainer.SetTitle("McAfee")
'                Call m_ObjUIContainer.SetSize (UIHeight, UIWidth)
'                Call m_ObjUIContainer.ClearWindowStyle(12582912)
'                Call m_ObjUIContainer.Center()
'                Call m_ObjUIContainer.Show()
'            End if
'        End If
'    End Function

    Public Function InitializeUIContainerObject()
        Dim intResult
        Dim objUIMgrObject
        Dim strUIProviderName
        
        Set objUIMgrObject = Nothing
        intResult = SM_UICONTAINER_OBJECT_CREATION_FAILED
        If Not (m_ObjExternal Is Nothing) Then
            Set m_ObjUIContainerEnum    = m_ObjExternal.GetParam(SM_UIPROVIDER_ENUM_OBJECT)
            'Set m_ObjUIContainer        = m_ObjExternal.GetParam("McUIContainer")    
            If Not(m_ObjUIContainerEnum Is Nothing) Then
                m_ObjUIContainerEnum.Reset
    	        Do
                    Set objUIMgrObject = m_ObjUIContainerEnum.Next()
                    If Not (objUIMgrObject Is Nothing) Then
                        strUIProviderName = objUIMgrObject.GetProperty(SM_PROP_UIMGR_NAME)    
                        If strUIProviderName = "MSC" Then
                            Set m_ObjUIContainer = objUIMgrObject    
                            Exit Do
                        End If
                    End If
                Loop Until objUIMgrObject Is Nothing
                
                intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS
                ShowErrorMessage("SMUIContainer:UIContainer Enum object created")
            Else
                intResult = SM_UICONTAINER_OBJECT_CREATION_FAILED
                ShowErrorMessage("SMUIContainer:Unable to get UIContainer Enum object")
            End If
        Else
            intResult = SM_UICONTAINER_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMUIContainer:Unable to get window.external object")
        End If
        
        InitializeUIContainerObject = intResult
    End Function

     Public Function DisplayUI(strWidth,strHeight)
        ShowErrorMessage("Calling UIContainerEnum width " & strWidth)
        
	    ShowErrorMessage("Calling UIContainerEnum strHeight " & strHeight)
	    If (Len(Trim(strWidth)) = 0) Then
	        strWidth = "360"    
	    End If
	    If (Len(Trim(strHeight)) = 0) Then
	        strHeight = "399"    
	    End If
	    If Not(m_ObjUIContainerEnum Is Nothing) Then
            Call window.external.GetParam("McUIContainer").ClearWindowStyle(12582912)   
            Call m_ObjUIContainerEnum.ShowContainer(SM_UICONTAINER_BOTTOM_RIGHT_CORNER,strWidth,strHeight,0,0)
	    End If
	End Function
	
	Public Function GetLastInputInfo()
	    Dim LastInputInfoInMilisec
	    LastInputInfoInMilisec = "0"
	    If Not (m_ObjUIContainer is Nothing) Then
	        LastInputInfoInMilisec = m_ObjUIContainer.GetLastInputInfo()
	    End If
	    ShowErrorMessage("LastInputInfo is " & LastInputInfoInMilisec)
	    GetLastInputInfo = LastInputInfoInMilisec
	End Function
	
	Public Function GetIfSystemInScreenSaverMode()
	    Dim bIsScreenSaverMode
	    bIsScreenSaverMode = false
	    If Not (m_ObjUIContainer is Nothing) Then
	        bIsScreenSaverMode = m_ObjUIContainer.IsScreenSaverMode()
	    End If
	    ShowErrorMessage("bIsScreenSaverMode is " & bIsScreenSaverMode)
	    GetIfSystemInScreenSaverMode = bIsScreenSaverMode
	End Function
	
	Public Function GetIfSystemInFullScreenMode()
	    Dim bIsGamingMode
	    bIsGamingMode = false
	    If Not (m_ObjUIContainer is Nothing) Then
	        bIsGamingMode = m_ObjUIContainer.IsGamingMode()
	    End If
	    ShowErrorMessage("bIsGamingMode is " & bIsGamingMode)
	    GetIfSystemInFullScreenMode = bIsGamingMode
	End Function	
	
	
	
	Public Function NetworkDrivesExist()
	   On Error resume next
	    Dim arrNetworkDrives
	    Dim blnDriveExist
	    Dim strNetworkContent
	    'MsgBox "Networkdrives"
	    blnDriveExist = SM_UIMGR_NETWORKDRIVES_MISSING
	    If Not (m_ObjUIContainer is Nothing) Then
	        arrNetworkDrives = m_ObjUIContainer.GetNetworkDrives()
	        If UBound(arrNetworkDrives) > 0 Then
	            blnDriveExist = SM_UIMGR_NETWORKDRIVES_EXIST
	        End If
	    End If
	    
	    NetworkDrivesExist = blnDriveExist
	End Function
    
    Private Sub Class_Terminate      
        
        If Not (m_ObjUIContainer is Nothing) Then 
            Set m_ObjUIContainer = Nothing
        End If
            
        If Not (m_ObjUIContainerEnum is Nothing) Then 
            Set m_ObjUIContainerEnum = Nothing
        End If
        
    End Sub   	
End Class


Class CSMSystemData
    
    Private m_objExternal
    Private m_objSystemInfo
    Private m_arrProcessorInfo
    Private m_NetworkDrivesExist
    Private m_LastInputInfoInSec
    Private m_intNoOfProcessors
    Private m_strProcessorFamily    '=== eg: intel (vendor)
    Private m_strProcessorVersion   '=== eg: CPU Name "Centrino"
    Private m_strProcessorSpeed
    Private m_blnNetworkDrives
    
    
    Private Sub Class_Initialize
        Set m_objExternal     = window.external
        Set m_objSystemInfo   = Nothing
       
    End Sub
    
    Public Property Get IsConnectedToInternet()
        Dim bInternetState: bInternetState = false
        
        bInternetState    = m_objSystemInfo.InternetConnectedState()
        IsConnectedToInternet   = bInternetState
     End Property	 

     Public Property Get MachineOsId()
        On Error Resume Next

        Dim strMachineOSNumber : strMachineOSNumber = m_objSystemInfo.GetOSVersion()
        If Len(Trim(strMachineOSNumber)) > 0 And IsNumeric(strMachineOSNumber) Then
            MachineOsId = CInt(strMachineOSNumber)
        Else
            MachineOsId = 0
        End If
     End Property
    
     Public Property Get MachineOS()
        on error resume next
        Dim strMachineOSNumber, strMachineOS
        strMachineOSNumber    = m_objSystemInfo.GetOSVersion()
        If Len(Trim(strMachineOSNumber)) > 0 Then
            Select Case strMachineOSNumber
                Case "1"
                    strMachineOS = "Windows 95"
                Case "2"
                    strMachineOS = "Windows 98"
                Case "3"
                    strMachineOS = "Windows NT"
                Case "4"
                    strMachineOS = "Windows 2000"
                Case "5"
                    strMachineOS = "Windows ME"
                Case "6"
                    strMachineOS = "Windows XP 32bit"
                Case "7"
                    strMachineOS = "Windows Vista 32bit"
                Case "8"
                    strMachineOS = "Windows XP 64bit"
                Case "9"
                    strMachineOS = "Windows Vista 64bit"
                Case "10"
                    strMachineOS = "Windows 7 32bit"
                Case "11"
                    strMachineOS = "Windows 7 64bit"
                Case "12"
                    strMachineOS = "Windows8_32bit"
                Case "13"
                    strMachineOS = "Windows8_64bit"
                Case "14"
                    strMachineOS = "Windows8_32bit" ' Actually Windows8.1_32bit
                Case "15"
                    strMachineOS = "Windows8_64bit" ' Actually Windows8.1_64bit
                Case "16"
                    strMachineOS = "Windows10_32bit"
                Case "17"
                    strMachineOS = "Windows10_64bit"
                Case else
                    strMachineOS = ""
            End Select
        
            MachineOS   = strMachineOS
        Else
            MachineOS = ""
        End If
     End Property
     
     Public Property Get MachineHardwareID()
        Dim strHardwareID
        strHardwareID    = m_objSystemInfo.GetMachineID()	    
        'strHardwareID    = m_objSystemInfo.GetProperty(MC_PROP_SYSTEMINFO_HARDWAREID)
        If Len(Trim(strHardwareID)) > 0 Then
            MachineHardwareID   = strHardwareID
        Else
            MachineHardwareID = ""
        End If
     End Property

     Public Property Get MachineSoftwareID()
        Dim strSoftwareID
        strSoftwareID    = m_objSystemInfo.GetWindowsID()
        'strSoftwareID    = m_objSystemInfo.GetProperty(MC_PROP_SYSTEMINFO_SOFTWAREID)
        If Len(Trim(strSoftwareID)) > 0 Then
            MachineSoftwareID   = strSoftwareID
        Else
            MachineSoftwareID = ""
        End If
     End Property	 

     Public Property Get MachineMemorySize()  '=== Check whether we get that in MB, as API needs that in MB
        on error resume next
        Dim strMachineMemory
        strMachineMemory    = m_objSystemInfo.GetSystemMemorySize()
        If Len(Trim(strMachineMemory)) > 0 Then
            MachineMemorySize   = GetMemorySizeInMB(strMachineMemory)
        Else
            MachineMemorySize = ""
        End If
     End Property     
     
     Public Property Get NumberOfProcessors()
        If m_intNoOfProcessors > 0 Then
            NumberOfProcessors   = m_intNoOfProcessors
        Else
            NumberOfProcessors = 0
        End If
     End Property	 
     
     Public Property Get ProcessorFamily()
        If Len(Trim(m_strProcessorFamily)) > 0 Then
            ProcessorFamily   = m_strProcessorFamily
        Else
            ProcessorFamily = ""
        End If
     End Property		 

     Public Property Get ProcessorVersion()
        If Len(Trim(m_strProcessorVersion)) > 0 Then
            ProcessorVersion   = m_strProcessorVersion
        Else
            ProcessorVersion = ""
        End If
     End Property	
     
     Public Property Get ProcessorSpeedGHZ()
        Dim strProcSpeedGHZ
        strProcSpeedGHZ = ""
        If Len(Trim(m_strProcessorSpeed)) > 0 Then
            strProcSpeedGHZ = GetSpeedInGHZ(m_strProcessorSpeed)
            ProcessorSpeedGHZ   = strProcSpeedGHZ
        Else
            ProcessorSpeedGHZ = 0
        End If
     End Property	
     
     Public Property Get NetworkDrivesExist()
        NetworkDrivesExist  = m_NetworkDrivesExist
     End Property

     Public Property Get IsSystemInUse()
        Dim intLastInputInfoInMSec
        Dim intSystemInUse, intResult
        Dim objSMUIContainer
         
        Set objSMUIContainer    = New CSMUIContainer
        intLastInputInfoInMSec   = 0
        If Not objSMUIContainer is Nothing Then
            intResult = objSMUIContainer.InitializeUIContainerObject()
            If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
                intLastInputInfoInMSec = objSMUIContainer.GetLastInputInfo  
            End If
        End If    
        intSystemInUse          = SM_SYSMGR_SYSTEM_IDLE
        ShowErrorMessage "intLastInputInfoInMSec is " & intLastInputInfoInMSec
        If intLastInputInfoInMSec > 0 Then
            ShowErrorMessage "intLastInputInfoInMSec is " & intLastInputInfoInMSec
            If intLastInputInfoInMSec < SM_SYSMGR_ALLOWED_IDLE_TIME_IN_MSEC Then
                intSystemInUse  = SM_SYSMGR_SYSTEM_IN_USE
            Else
                intSystemInUse  = SM_SYSMGR_SYSTEM_IDLE
            End If
        Else
            intSystemInUse  = SM_SYSMGR_SYSTEM_IN_USE
        End If
        ShowErrorMessage "intSystemInUse is " & intSystemInUse
        IsSystemInUse = intSystemInUse
     End Property    

     Public Function InitializeSystemObject(objProviderEnum)
      
        On Error Resume Next
        
        Dim ProviderEnumObject 
        Dim intResult, retSubInfo
        Dim testArray
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objSystemInfo = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_SYSTEM_MGR)
            
            If Not (m_objSystemInfo Is Nothing) Then
               
                testArray = m_objSystemInfo.GetProcessorInfo()
                m_arrProcessorInfo  = m_objSystemInfo.GetProcessorInfo()
                
                '=== Newly added code below
                Dim objSMUIContainer
                Set objSMUIContainer    = New CSMUIContainer
                
                If Not objSMUIContainer is Nothing Then
                    intResult = objSMUIContainer.InitializeUIContainerObject()
                    
                    If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
                        m_LastInputInfoInSec    = objSMUIContainer.GetLastInputInfo  
                        m_NetworkDrivesExist    = objSMUIContainer.NetworkDrivesExist    
                    End If
                    
                End If    
                '=== Newly added code above
                If Length(m_arrProcessorInfo) = 0 Then
                    ShowErrorMessage("SMSystemData:Unable to get the Processor Info")
                    intResult = SM_SYSMGR_OBJECT_CREATION_FAILEDs
                Else
                    SetProcessorInfo()
                End If
                intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS
            End If
        Else
            intResult = SM_SYSMGR_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMAppData:Unable to get the Provider object")
        End If
        
        InitializeSystemObject = intResult 
    End Function
    
    Private Sub SetProcessorInfo()
        on Error Resume next
        
        Dim intNoOfProcessors
        Dim objProcessorInfo
        
        Set objProcessorInfo = Nothing
        intNoOfProcessors = Length(m_arrProcessorInfo)
        
        If intNoOfProcessors > 0 Then
            Set objProcessorInfo = m_arrProcessorInfo(0)
            m_intNoOfProcessors = intNoOfProcessors
            If Not (objProcessorInfo Is Nothing) Then
                m_strProcessorFamily    = objProcessorInfo.GetCPUFamily()
                m_strProcessorVersion   = objProcessorInfo.GetCPUModel()  
                m_strProcessorSpeed     = objProcessorInfo.GetCPUSpeed()  
            End If
        End If
    End Sub
    
    Public Function GetMemorySizeInMB(strMemorySizeInKB)
        Dim lMemorySizeMB
        Dim intMBEquiv
        lMemorySizeMB = 0
        intMBEquiv = 1024
        If Len(Trim(strMemorySizeInKB)) > 0 Then
            lMemorySizeMB = Round(CDBL(strMemorySizeInKB) / intMBEquiv)
        End If
        
        GetMemorySizeInMB = lMemorySizeMB
    End Function

    ' Added for Smart Messaging Optimization 
    Public Function LaunchApplication(strExePath, strCmdLine)
        Dim intResult
        
        intResult = window.external.LaunchApplication(cstr(strExePath), cstr(strCmdLine))

        LaunchApplication = intResult
    End Function

    Public Function LaunchSAApplication(objProviderEnum, strExePath, strCmdLine)
        On Error Resume Next
        
        Dim ProviderEnumObject 
        Dim intResult, retSubInfo
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objSystemInfo = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_SYSTEM_MGR)
            If Not (m_objSystemInfo Is Nothing) Then
                intResult  = m_objSystemInfo.RunProgram(strExePath, strCmdLine)
                If intResult <> 0 Then
                    ShowErrorMessage("SMSystemData:Unable to launch SA")
                    intResult = SM_SYSMGR_OBJECT_CREATION_FAILEDs
                End If
                intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS
            End If
        Else
            intResult = SM_SYSMGR_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMAppData:Unable to get the Provider object")
        End If
        
        LaunchSAApp = intResult
    End Function
    ' End of Smart Messaging Optimization 
    
    Private Function GetSpeedInGHZ(SpeedInMB)
        Dim lSpeedGHZ
        Dim lGHZEquiv
        lSpeedGHZ = 0
        lGHZEquiv = 1000
        
        lSpeedGHZ = Round(CDBL(SpeedInMB) / lGHZEquiv)
        GetSpeedInGHZ = lSpeedGHZ
    End Function
    
     Private Sub Class_Terminate
        If Not IsEmpty(m_objSystemInfo) Then
            If Not m_objSystemInfo Is Nothing Then
                Set m_objSystemInfo = Nothing
            End If
        End If
        If Not IsEmpty(m_objExternal) Then
            If Not m_objExternal Is Nothing Then
                Set m_objExternal = Nothing
            End If
        End If
    End Sub


End Class


Class CSMClientDB

    Private m_objDBMgr
    
    Private Sub Class_Initialize()
        Set m_objDbMgr      = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        If Not IsEmpty(m_objDbMgr) Then
            If Not m_objDbMgr Is Nothing Then
                Set m_objDbMgr = Nothing
            End If
        End If
    End Sub

    Public Function InitializeClientDBObject(objProviderEnum)
   	    On Error Resume Next
        
        Dim ProviderEnumObject 
        Dim intResult
        Dim strDbOpenStatus
        
        intResult = 0
        strDbOpenStatus = SM_CLIENTDB_MISSING_DB
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objDbMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_CLIENTDB_OBJECT)
            
            If Not (m_objDbMgr Is Nothing) Then
                strDbOpenStatus = OpenDB()
                If strDBOpenStatus = SM_CLIENTDB_DB_OPEN Then
                    intResult = SM_CLIENTDB_OBJECT_CREATION_SUCCESS
                    ShowErrorMessage("ClientDB:ClientDB object created")
                Else
                    intResult = SM_CLIENTDB_OBJECT_CREATION_FAILED
                    ShowErrorMessage("ClientDB:Unable to open the Client DB")    
                End if
            Else
                intResult = SM_CLIENTDB_OBJECT_CREATION_FAILED
                ShowErrorMessage("ClientDB:Unable to get the ClientDB object")
            End if
        Else
            intResult = SM_CLIENTDB_OBJECT_CREATION_FAILED
            ShowErrorMessage("ClientDB:Unable to get the Provider object")
        End If
   	    
''''   	    If Not (ProviderEnumObject is Nothing) Then
''''   	        ProviderEnumObject = Nothing
''''   	    End If
   	
   	    InitializeClientDBObject = intResult
   	End Function

    Private Function OpenDB()
        Dim strResult
        Dim strDBOpen
        
        strDBOpen   = SM_CLIENTDB_MISSING_DB
	    If Not IsEmpty(m_objDbMgr) Then
	        If Not m_objDbMgr Is Nothing Then
                strResult = m_objDbMgr.Open(SM_SMART_MESSAGING_CLIENT_DB_NAME)
     
                If CStr(strResult) = SM_CLIENTDB_MISSING_DB Then '=== SQLite error or missing database
                  ShowErrorMessage("Unable to create/open Client database")
                  strDBOpen = SM_CLIENTDB_MISSING_DB  
                ElseIf CStr(strResult) = SM_CLIENTDB_DB_OPEN Then '=== SQLite Successful result
                    strDBOpen = SM_CLIENTDB_DB_OPEN
                End if
            Else
                ShowErrorMessage("ClientDB object is empty")
                strDBOpen = SM_CLIENTDB_MISSING_DB
            End If
        Else
            ShowErrorMessage("ClientDB object is empty")
            strDBOpen = SM_CLIENTDB_MISSING_DB    
         End If
         
         OpenDB  = strDBOpen
    End Function
   
    Public Function SyncDB()
        Dim strClientDBVersion
        Dim bTablesCreated
        
        bTablesCreated = SM_CLIENTDB_ALL_TABLES_CREATED_FAIL
        
        strClientDBVersion = GetClientSMVersion()
        
        If Len(Trim(strClientDBVersion)) > 0 Then
            

        
        Else
            bTablesCreated = CreateAllTables()
            If bTablesCreated = SM_CLIENTDB_ALL_TABLES_CREATED_SUCCESS Then
                SetClientSMVersion(SM_CLIENTDB_DEFAULT_VERSION)
            End If
        End If
    
    End Function
    
    Public Function	SelectRecords(strTableName, strColumnList, strCondition)
        Dim strSelectQry
        Dim intResult
        intResult = SM_CLIENTDB_QUERY_EXEC_FAIL
        strSelectQry = ""
        
        If (Len(Trim(strColumnList)) > 0) And (Len(Trim(strTableName)) > 0) Then
            strSelectQry = "SELECT " & strColumnList & " FROM " & strTableName
            
            If Len(Trim(strCondition)) > 0 Then
               strSelectQry = strSelectQry & " WHERE " & strCondition  
            End If
        End If
        intResult = ExecuteQueries(strSelectQry)
        SelectRecords = intResult
    End Function
    
    Public Function	InsertRecord(strTableName, strColumnList, strColumnValues)
        Dim strInsertQry
        Dim intResult
        intResult = SM_CLIENTDB_QUERY_EXEC_FAIL
        strInsertQry = ""
        
        If (Len(Trim(strColumnList)) > 0) And (Len(Trim(strTableName)) > 0) And (Len(Trim(strColumnValues)) > 0)Then
            strInsertQry = "INSERT INTO " & strTableName & " ( " & strColumnList & " ) VALUES ( " & strColumnValues & " )"
        End If
        
        intResult = ExecuteQueries(strInsertQry)
        InsertRecord = intResult
    End Function

    Public Function UpdateColumns(tableName,setCondition,whereCondition)
        Dim updateQry,resUpdate
        updateQry = ""
        
        If Len(Trim(tableName)) = 0 Then
            UpdateColumns = SM_CLIENTDB_QUERY_EXEC_FAIL
            Exit Function
        End If

        If Len(Trim(setCondition)) = 0 Then
            UpdateColumns = SM_CLIENTDB_QUERY_EXEC_FAIL
            Exit Function
        End If

        If Len(Trim(whereCondition)) = 0 Then
            UpdateColumns = SM_CLIENTDB_QUERY_EXEC_FAIL
            Exit Function
        End If

        updateQry = "UPDATE " & tableName & " SET " & setCondition

        If Len(Trim(whereCondition)) > 0 Then
            updateQry = updateQry &  " WHERE " & whereCondition
        End If
        resUpdate = ExecuteQueries(updateQry)
        UpdateColumns = resUpdate
    End Function
    
    Public Function UpdateRecord(strTableName, strColumnName, strColumnValue, strCondition)
        Dim strUpdateQry
        Dim intResult
        strUpdateQry = ""
        
        If (Len(Trim(strColumnName)) > 0) And (Len(Trim(strTableName)) > 0) And (Len(Trim(strColumnValue)) > 0)Then
            strUpdateQry = "UPDATE " & strTableName & " SET " & strColumnName & " = " & strColumnValue 
            
            If Len(Trim(strCondition)) > 0 Then
               strUpdateQry = strUpdateQry & " WHERE " & strCondition  
            End If
        End If
        
        intResult = ExecuteQueries(strUpdateQry)
        UpdateRecord = intResult
    End Function
    
    Public Function	DeleteRecord(strTableName, strCondition)
        Dim strDeleteQry
        
        strDeleteQry = ""
        
        If (Len(Trim(strTableName)) > 0)Then
            strDeleteQry = "DELETE FROM " & strTableName
            
            If Len(Trim(strCondition)) > 0 Then
               strDeleteQry = strDeleteQry & " WHERE " & strCondition  
            End If
        End If
        
        DeleteRecord = strDeleteQry
    End Function
    
    '=== This function should be called only after executing
    '=== the count(*) query to get the number of rows exist
    Public Function GetNumberOfRows()
     On Error Resume Next
        Dim intNumOfRows

        intNumOfRows = 0
        If (Not IsEmpty(m_objDbMgr)) And (Not m_objDbMgr Is Nothing) Then
            intNumOfRows = m_objDbMgr.GetData(0)   
        End If
        
        If Len(Trim(intNumOfRows)) > 0 Then
            intNumOfRows = CInt(intNumOfRows)    
        Else
            intNumOfRows = 0
        End If
        
        If Err.number > 0 Then
            ShowErrorMessage("SMCLientDB:Unable to get the records count for the executed query")
        End If
        GetNumberOfRows = intNumOfRows
    End Function
    
    Public Sub PointToNextRow()
        On Error Resume Next
        If (Not IsEmpty(m_objDbMgr)) And (Not m_objDbMgr Is Nothing) Then
            m_objDbMgr.Step()
        End If
        If Err.number > 0 Then
            ShowErrorMessage("SMCLientDB:Unable to point to next record")
        End If
    End Sub
    
    Public Function GetData(intColumnNo)
        On Error Resume Next
        Dim strColData
        
        strColData = ""
        If (Not IsEmpty(m_objDbMgr)) And (Not m_objDbMgr Is Nothing) Then
            strColData = m_objDbMgr.GetData(intColumnNo)
        Else
            strColData = ""
        End If   
        
        If Err.number > 0 Then
            ShowErrorMessage("SMCLientDB:Unable to get data for this column")
        End If     
        GetData = strColData
    End Function

    Public Function ExecuteQueries(strQry)
        Dim intQrySuccessful, bDBopen
      
        intQrySuccessful = SM_CLIENTDB_QUERY_EXEC_FAIL
        
        If Len(Trim(strQry)) > 0 Then
            If Not (m_objDBMgr Is Nothing) Then
                intQrySuccessful = m_objDBMgr.Exec(strQry)
            Else
                intQrySuccessful = SM_CLIENTDB_QUERY_EXEC_FAIL
            End If
        Else
            intQrySuccessful = SM_CLIENTDB_QUERY_EXEC_FAIL      
        End If
        
        If intQrySuccessful = SM_CLIENTDB_OPERATION_OK Or intQrySuccessful = SM_CLIENTDB_NO_MORE_DATA _
           Or intQrySuccessful = SM_CLIENTDB_DATA_FOR_CURRENT_ROW_READY Then
            intQrySuccessful = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL
        Else
            intQrySuccessful = SM_CLIENTDB_QUERY_EXEC_FAIL
        End If
        
        ExecuteQueries = intQrySuccessful
    End Function  
    
    Function ValidateDBHitStatus(intHitLimit, strServerUTCTime)
        Dim objSMClientDB
        Dim strTodaysDate, strDBStatus
        Dim strHitLimitColumn, strHitLimitCondition
        Dim intResult, intNumOfRows
        Dim blnExceededDBHitLimit
        
        If Len(Trim(strServerUTCTime)) > 0 Then
            strTodaysDate           = GetCurrentDate(strServerUTCTime)
        Else
            strTodaysDate           = GetCurrentDate(Date())
        End If
        
        blnExceededDBHitLimit   = False
        intNumOfRows            = 0
        intResult               = 0
        
        strHitLimitColumn = "Count(MessageSyncLogID) as NumOfHits"
        strHitLimitCondition = "CreatedDate like '" & strTodaysDate & "%'" 
        intResult   = SelectRecords(SM_MESSAGE_SYNC_LOG_TABLE_NAME, strHitLimitColumn, strHitLimitCondition)
        If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
            intNumOfRows = GetNumberOfRows()
            If CInt(intNumOfRows) > 0 Then
                If CInt(intHitLimit) > intNumOfRows Then
                    blnExceededDBHitLimit   = False   
                Else
                    blnExceededDBHitLimit   = True
                End If
            Else
                blnExceededDBHitLimit   = False
            End If
        Else
            ShowErrorMessage("SyncMessagevbs:Unable to get the DBHitLimit")  
        End If
        
        ValidateDBHitStatus = blnExceededDBHitLimit
  End Function

  Public Function DeleteAllTables()
    
    Dim result,numOfRows,count
    Dim tableName,dropTable


    result = SelectRecords("sqlite_master", " count(name)", " type='table'")        ''Get the number of rows for the select query
    If result = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
	    numOfRows = GetNumberOfRows
        
        If numOfRows = 0 Then           ''There are no client tables in the DB
            DeleteAllTables = True      
            Exit Function
        End If        

		For count = 0 To numOfRows -1                        
            result = SelectRecords("sqlite_master", " name", " type='table'")           ''get the list of all tables in clientdb
            If result = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                dropTable = "DROP TABLE " & GetData(0)
                result = ExecuteQueries(dropTable)
                If result <> SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                    DeleteAllTables = False
                    Exit Function
                End If
            End If
		    If (count + 1) < numOfRows Then
                PointToNextRow
		    End If
		Next
    End If

    If ((numOfRows = 0) Or (result = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL)) Then
        DeleteAllTables = True
    Else
        DeleteAllTables = False
    End If

  End Function

End Class


Class CSMRegistry
    
    Private m_objRegistryMgr
    
    Private Sub Class_Initialize
        Set m_objRegistryMgr    = Nothing
    End Sub
    
    Public Function InitializeRegistryObject(objProviderEnum)
    On Error Resume Next
        ShowErrorMessage "Initialize Function Called"
        Dim ProviderEnumObject 
        Dim intResult
        
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objRegistryMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
            
            If Not (m_objRegistryMgr Is Nothing) Then
                intResult = SM_REGISTRY_OBJECT_CREATION_SUCCESS
                ShowErrorMessage("SMRegistry:Registry object created")
            Else
                intResult = SM_REGISTRY_OBJECT_CREATION_FAILED
                ShowErrorMessage("SMRegistry:Unable to get the Registry object")
            End if
        Else
            intResult = SM_REGISTRY_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMRegistry:Unable to get the Provider object")
        End If
    
        InitializeRegistryObject = intResult
    End Function
    
    Private Property Get SMRegistryPath
        Dim bIsRegPathFound : bIsRegPathFound = False
        
        If Not m_objRegistryMgr Is Nothing Then
            bIsRegPathFound = m_objRegistryMgr.RegKeyPresent(SM_REG_LOCATION)
            If CBool(bIsRegPathFound) Then
                ' SM Registry location used by WSS 12.0 onward
                SMRegistryPath = SM_REG_LOCATION
            Else
                ' SM Registry location used previous to WSS 12.0
                SMRegistryPath = SM_CLIENTDB_VERSION_LOCATION
            End If            
        Else
            ShowErrorMessage("Unable to create Registry object") 
        End If
    End Property    
    
    Public Function GetClientSMVersion()
       Dim bResult, strSMVersion
       
       bResult = False
       strSMVersion = ""
            
            If Not m_objRegistryMgr Is Nothing Then
              bResult = m_objRegistryMgr.RegKeyPresent(SMRegistryPath)
               
               If CBool(bResult) Then
                   '=== TODO: Check whether this will get the Version value correctly.
                   '=== RegQueryValue(BSTR bstrKeyPath,BSTR bstrValue,VARIANT *pvValue) 
                    strSMVersion = m_objRegistryMgr.RegQueryValue(SMRegistryPath, SM_VERSION_KEY_NAME)
               Else
                    strSMVersion = SM_NO_VERSION
               End If
            Else   
               ShowErrorMessage("Unable to create Registry object")
            End If
        
        GetClientSMVersion = strSMVersion
    End Function

    Public Sub SetClientSMVersion(strSMVersionValue)
      
        Dim bCreateKey, bSetValue, bResult
        bCreateKey  = False
        bResult     = False
        bSetValue   = False
       
        If Not m_objRegistryMgr Is Nothing Then
            bResult = m_objRegistryMgr.RegKeyPresent(SMRegistryPath)
            If bResult Then
                '=== TODO: Check whether this parameter is passed correctly.
                '=== RegSetValue(BSTR bstrKeyPath,BSTR bstrKeyValueName,VARIANT vValue,BOOL *pbResult) 
                bSetValue = m_objRegistryMgr.RegSetValue(SMRegistryPath, SM_VERSION_KEY_NAME, Trim(strSMVersionValue))
                If CBool(bSetValue) Then
                    ShowErrorMessage("ResgistryValue for SMVersion is created")  
                Else
                    ShowErrorMessage("Unable to set Client SMVersion in Registry")
                End If
            Else
                ShowErrorMessage("MSM is missing in registry")
            End If
        Else
            ShowErrorMessage("Unable to create Registry object")       
        End If
    End Sub
    
    'ARPU Fix Start
    
    Public Sub SetRegKeyValueCstr(strKeyName,strValueName,strValue,bObfuscate)
        Dim bOldObfuscate,bSetValue
        bSetValue = False
        If Not m_objRegistryMgr Is Nothing Then
           bOldObfuscate = m_objRegistryMgr.Obfuscate
           m_objRegistryMgr.Obfuscate = bObfuscate
           bSetValue = m_objRegistryMgr.RegSetValue(strKeyName,strValueName, Trim(strValue))
          If CBool(bSetValue) Then
               ShowErrorMessage("Registry Entry " & strKeyName & " " & strValueName & " " & " is set to " & m_objRegistryMgr.RegQueryValue(strKeyName, strValueName) )
           Else
               ShowErrorMessage("Unable to set Registry entry " & strKeyName & " " & strValueName)
           End If            
           m_objRegistryMgr.Obfuscate = bOldObfuscate
        End If
    End Sub
    
    Public Function GetRegKeyValueName(strKeyName,strValueName,bObfuscate)
'        ShowErrorMessage("Registry Entry " & strKeyName & " " & strValueName)
        Dim bOldObfuscate,strValue,bResult
        strValue = ""
        bResult  = False
        If Not m_objRegistryMgr Is Nothing Then
            bOldObfuscate = m_objRegistryMgr.Obfuscate             
            m_objRegistryMgr.Obfuscate = bObfuscate
            strValue = m_objRegistryMgr.RegQueryValue(strKeyName, strValueName)
            ShowErrorMessage("Aff Id " & strValue)
            m_objRegistryMgr.Obfuscate = bOldObfuscate
'        Else
'           ShowErrorMessage("Registry Mgr is nothing")
        End If
'        ShowErrorMessage("Aff Id " & strValue)
        GetRegKeyValueName = strValue 
    End Function   	
    
    'ARPU Fix End  
    Public Function GetVersionId()
        Dim bResult, versionId
        bResult = False
        versionId = ""
                   
        If Not m_objRegistryMgr Is Nothing Then
            bResult = m_objRegistryMgr.RegKeyPresent(SM_WEBUPDATE_LOCATION)
               
            If CBool(bResult) Then
                versionId = m_objRegistryMgr.RegQueryValue(SM_WEBUPDATE_LOCATION,SM_VERSIONID_KEY_NAME)
            Else
                versionId = ""
            End If
        Else   
            ShowErrorMessage("Unable to create Registry object")
        End If

        GetVersionId = versionId

    End Function
    
     Private Sub Class_Terminate        
        If Not(m_objRegistryMgr is Nothing) Then
            Set m_objRegistryMgr    = Nothing
        End if
    End Sub   	

End Class


Class CSMDBVersionMgr
    
    Private m_ObjDbMgr
    
    Private Sub Class_Initialize()
        Set m_ObjDbMgr = Nothing
    End Sub

    Private Sub Class_Terminate()
        If Not (m_ObjDbMgr is Nothing) Then
            Set m_ObjDbMgr = Nothing
        End If
    End Sub
    
    Public Function InitializeVersionMgr(objProvider)
        Dim intResult
           
        Set m_ObjDbMgr = New CSMClientDB
        intResult = m_ObjDbMgr.InitializeClientDBObject(objProvider)
        
        If intResult = SM_CLIENTDB_OBJECT_CREATION_SUCCESS Then
            intResult = SM_VERSIONMGR_OBJECT_CREATION_SUCCESS
        Else
            intResult = SM_VERSIONMGR_OBJECT_CREATION_FAILED    
        End If
        InitializeVersionMgr = intResult
    End Function
   
    Public Function ExecuteSMVersionScripts(objRegistryMgr, strClientSMVersion, strCurrentSMVersion)
        On Error Resume Next
        
        Dim strVersionName, strScriptExecutionStatus
        Dim strScriptVersion
        Dim intCnt, intResult
        
        If (Len(Trim(strClientSMVersion)) > 0) And (Len(Trim(strCurrentSMVersion)) > 0) Then
            For intCnt = (CInt(strClientSMVersion)+1) To CInt(strCurrentSMVersion)
                strVersionName  = "strScriptExecutionStatus = ExecuteVersion" & intCnt & "Scripts()"
                Execute(strVersionName)
                
                If strScriptExecutionStatus = SM_CLIENTDB_ALL_TABLES_CREATED_FAILED Then
                    strScriptVersion = "ExecuteVersion" & intCnt & "Scripts failed"
                    ShowErrorMessage("SMVersionMgr:Function " & strScriptVersion)  
                    strScriptExecutionStatus = strScriptVersion
                    Exit For    
                End If
                
                 '=== Newly added code
                     Call objRegistryMgr.SetClientSMVersion(intCnt)
            Next

        Else
            ShowErrorMessage("ExecuteSMVersionScripts:clientSMVersion or Server's SM Version is empty")     
        End If
        If Err.number > 0 Then
            ShowErrorMessage("ExecuteSMVersionScripts:Function " & strScriptVersion & " failed")  
        End If
        
        ExecuteSMVersionScripts = strScriptExecutionStatus
    End Function

    Public Function ExecuteVersion1Scripts()
        Dim strQryToCreateTbl
        Dim bMsgSyncLogQrySuccess, bMsgViewHistoryQrySuccess, bMsgLogQrySuccess
        Dim QryExecStatus 
        QryExecStatus = SM_CLIENTDB_ALL_TABLES_CREATED_FAILED
        If Not (m_ObjDbMgr Is Nothing ) Then
            strQryToCreateTbl = GetCreateMessageSyncLogQry()
            
            bMsgSyncLogQrySuccess = m_ObjDbMgr.ExecuteQueries(strQryToCreateTbl)

            If bMsgSyncLogQrySuccess = SM_CLIENTDB_QUERY_EXEC_FAIL then
                ShowErrorMessage("Unable to create SyncLog Table")      
            End If
            
            strQryToCreateTbl = ""
            strQryToCreateTbl = GetCreateMessageViewHistoryQry()
            bMsgViewHistoryQrySuccess = m_ObjDbMgr.ExecuteQueries(strQryToCreateTbl)

            If bMsgViewHistoryQrySuccess = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                strQryToCreateTbl = ""
                strQryToCreateTbl = GetCreateMessageLogQry()
                bMsgLogQrySuccess = m_ObjDbMgr.ExecuteQueries(strQryToCreateTbl)
                
                If bMsgLogQrySuccess = SM_CLIENTDB_QUERY_EXEC_FAIL then
                    ShowErrorMessage("Unable to create MessageLog Table")      
                End If
                
            Else
                ShowErrorMessage("Unable to create View History Table")      
            End If   

            If bMsgSyncLogQrySuccess = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL and _
               bMsgViewHistoryQrySuccess = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL and _
               bMsgLogQrySuccess = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                
                QryExecStatus = SM_CLIENTDB_ALL_TABLES_CREATED_SUCCESS
            Else
                QryExecStatus = SM_CLIENTDB_ALL_TABLES_CREATED_FAILED
            End If
        End If  
        
         If Err.number > 0 Then
            ShowErrorMessage "Error is " & Err.Description
        End If
        
        ExecuteVersion1Scripts = QryExecStatus
        
    End Function

    Public Function ExecuteVersion2Scripts()
        Dim intResult
        intResult = SM_CLIENTDB_QUERY_EXEC_FAIL
        If Not (m_ObjDbMgr Is Nothing ) Then
         '   intResult = m_objDbMgr.InsertRecord("MessageViewHistory", "MessageID,Category,Priority,SyncFlag,AcctID,Affid,LCID,ProductKey,createddate", "802,'Informational',2,0,0,0,1033,'','05/29/2009 21:36:25'")
        End If
        ExecuteVersion2Scripts = intResult
    End Function
    
    
 Public Function ExecuteSMVersionValidationScripts(objRegistryMgr, strSMClientVersion, strCurrentSMVersion)
        On Error Resume Next
	    
        Dim strValidationVersionName, strScriptExecutionStatus
        Dim strScriptVersion, strNewClientVersion
        Dim intCnt, intResult
		
		Dim strSMVersionKeyExists, strExistingClientVersion
		
		strExistingClientVersion = strSMClientVersion
		
        If Len(Trim(strCurrentSMVersion)) > 0 Then
            For intCnt = (CInt(strExistingClientVersion)+1) To CInt(strCurrentSMVersion)
                strValidationVersionName  = "strScriptExecutionStatus = ExecuteVersion" & intCnt & "ValidationScripts()"
                Execute(strValidationVersionName)
                
                If strScriptExecutionStatus = SM_CLIENTDB_VERSION_VALIDATION_FAILED Then
                    strScriptVersion = "ExecuteVersion" & intCnt & "ValidationScripts failed"
                    ShowErrorMessage("SMVersionMgr:Function " & strScriptVersion)  
                    Exit For    
                End If
                
                 '=== Newly added code
                     Call objRegistryMgr.SetClientSMVersion(intCnt)
            Next
		Else
			ShowErrorMessage("ExecuteSMVersionValidationScripts:ServerVersion is empty")     
		End If
        If Err.number > 0 Then
            ShowErrorMessage("ExecuteSMVersionValidationScripts:Function " & strScriptVersion & " failed")  
			strScriptExecutionStatus = SM_CLIENTDB_VERSION_VALIDATION_FAILED
        End If
        
		If Len(Trim(strScriptExecutionStatus)) = 0 Then
			strNewClientVersion = objRegistryMgr.GetClientSMVersion
			
			If strNewClientVersion = strCurrentSMVersion Then
				strScriptExecutionStatus = SM_CLIENTDB_VERSION_VALIDATION_SUCCESS
				
			Else
				strScriptExecutionStatus = SM_CLIENTDB_VERSION_VALIDATION_FAILED
			End If
		End If
		
        ExecuteSMVersionValidationScripts = strScriptExecutionStatus
    End Function
	
	Public Function ExecuteVersion1ValidationScripts()
        Dim strQryVersionValidation
        Dim bVersionValidationStatus
        Dim QryExecStatus 
		
        QryExecStatus = SM_CLIENTDB_VERSION_VALIDATION_FAILED
        If Not (m_ObjDbMgr Is Nothing ) Then
            strQryVersionValidation = GetVersionValidationScripts()
            
            bVersionValidationStatus = m_ObjDbMgr.ExecuteQueries(strQryVersionValidation)

            If bVersionValidationStatus = SM_CLIENTDB_QUERY_EXEC_FAIL then
                ShowErrorMessage("ExecuteVersion1ValidationScripts:Unable to find the MessageSynclog Table")      
				QryExecStatus = SM_CLIENTDB_QUERY_EXEC_FAIL
			Else
				QryExecStatus = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL
            End If
        Else
				QryExecStatus = SM_CLIENTDB_QUERY_EXEC_FAIL
			    ShowErrorMessage("ExecuteVersion1ValidationScripts:Unable to create m_ObjDbMgr object")      
        End If  
        
         If Err.number > 0 Then
			QryExecStatus = SM_CLIENTDB_QUERY_EXEC_FAIL
            ShowErrorMessage "ExecuteVersion1ValidationScripts: Error is " & Err.Description
        End If
        
        ExecuteVersion1ValidationScripts = QryExecStatus
        
    End Function

End Class


    Public Function GetCreateMessageSyncLogQry()
        Dim strCreatemsgSyncQry
        
        strCreatemsgSyncQry  =                        " CREATE TABLE MessageSyncLog("
        strCreatemsgSyncQry  = strCreatemsgSyncQry &  " MessageSyncLogID INTEGER PRIMARY KEY NOT NULL,"
        strCreatemsgSyncQry  = strCreatemsgSyncQry &  " SyncURL Varchar(50) NOT NULL,"
        strCreatemsgSyncQry  = strCreatemsgSyncQry &  " ResultTypeID int NOT NULL,"
        strCreatemsgSyncQry  = strCreatemsgSyncQry &  " NumMessage int, SyncFlag int NOT NULL,"
        strCreatemsgSyncQry  = strCreatemsgSyncQry &  " CreatedDate TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL);"
        
        GetCreateMessageSyncLogQry = strCreatemsgSyncQry
    End Function
    
    Public Function GetCreateMessageViewHistoryQry()
        Dim strMessageViewHistoryQry
        
        strMessageViewHistoryQry  =                             " CREATE TABLE MessageViewHistory("
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " MessageViewHistoryID INTEGER PRIMARY KEY NOT NULL,"
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " MessageID int NOT NULL, Category Varchar(50) NOT NULL,"
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " Priority int NOT NULL, SyncFlag int NOT NULL,"
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " Affid int NOT NULL, LCID int NOT NULL, AcctID int, "
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " ProductKey Varchar(50),"
        strMessageViewHistoryQry  = strMessageViewHistoryQry &  " CreatedDate TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL);"
        
        GetCreateMessageViewHistoryQry = strMessageViewHistoryQry
    End Function
    
    Public Function GetCreateMessageLogQry()
        Dim strMessageLogQry
        strMessageLogQry  = ""
        strMessageLogQry  =                     " CREATE TABLE MessageLog("
        strMessageLogQry  = strMessageLogQry &  " MessageLogID INTEGER PRIMARY KEY NOT NULL,"
        strMessageLogQry  = strMessageLogQry &  " MessageViewHistoryID int NOT NULL,"
        strMessageLogQry  = strMessageLogQry &  " ActionTypeID int NOT NULL,"
        strMessageLogQry  = strMessageLogQry &  " ActionValue int NOT NULL,"
        strMessageLogQry  = strMessageLogQry &  " ActionURL Varchar(50),"
        strMessageLogQry  = strMessageLogQry &  " SyncFlag int NOT NULL," 
        strMessageLogQry  = strMessageLogQry &  " CreatedDate TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,"
        strMessageLogQry  = strMessageLogQry &  " FOREIGN KEY (MessageViewHistoryID) REFERENCES MessageViewHistory(MessageViewHistoryID));"
        
        GetCreateMessageLogQry = strMessageLogQry
    End Function
	
	Public Function GetVersionValidationScripts()
		Dim strValidationQry
        strValidationQry  = "Select count(*) from MessageSyncLog;"
		GetVersionValidationScripts = strValidationQry
	End Function


Class CSMInstrumentation
    Private m_objSMAppData, m_objSMSystemData
    Private m_strSoftwareID, m_strHardWareID
    Private m_strAcctID, m_strInstalledLCID
    
    Private Sub Class_Initialize()
        Set m_objSMAppData      = Nothing
        Set m_objSMSystemData   = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        If Not IsEmpty(m_objSMAppData) Then
            If Not m_objSMAppData Is Nothing Then
                Set m_objSMAppData = Nothing
            End If
        End If
        
        If Not IsEmpty(m_objSMSystemData) Then
            If Not m_objSMSystemData Is Nothing Then
                Set m_objSMSystemData = Nothing
            End If
        End If
        
    End Sub

     Public Function InitializeInstrumentationObject(objProviderEnum)
       ' On Error Resume Next
        Dim intProvSize, intClientObjResult, intAppDataObjResult
        Dim intRetVal, intIndex, intSystemDataObjResult, intInstruResult
        Dim objProvider
        
        intIndex    = 0
        intInstruResult = SM_INSTRUMENTATION_OBJECT_CREATION_FAILED
        intRetVal   = 0
        
        Set m_objSMSystemData   = New CSMSystemData
        Set m_objSMAppData      = New CSMAppData
        
        If Not(objProviderEnum Is Nothing) Then
            intProvSize = objProviderEnum.ProviderSize
            If intProvSize > 0 Then
                intSystemDataObjResult   = m_objSMSystemData.InitializeSystemObject(objProviderEnum)
                Set objProvider = objProviderEnum.GetProviderObject(intRetVal, intIndex)
                If intRetVal = SM_PROVIDER_OBJECT_CREATION_SUCCESS Then
                    If IsObject(objProvider) And Not(objProvider Is Nothing) Then
                        intAppDataObjResult = m_objSMAppData.InitializeAppMgrObject(objProvider)
                    End If
                End If
               If (intSystemDataObjResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS) _
                  And (intAppDataObjResult = SM_GET_INSTALLED_APPS_SUCCESS)Then 
                  intInstruResult = SM_INSTRUMENTATION_OBJECT_CREATION_SUCCESS
               End If
            End If
        End If
        
        If Err.number > 0 Then
            ShowErrorMessage "Error is Test : " & Err.Description
        End If
        
        InitializeInstrumentationObject = intInstruResult
     End Function
     
        Public Function GetInstruXML(objClientDB, objProviderEnum, strCSPClientID)
       ' On Error Resume Next
       
        Dim intRetVal, intIndex
        Dim strColumnNames, strJoinTableNames, strAffID, strActionXML
        Dim strCondition, strProductKey, strMsgActionDt
        Dim strMsgID, strActionURL, strMsgViewDt, strActionTypeID
        Dim strInstruXML, strMsgXML, strCID, strAcctID, strLCID, strRuleId
        Dim intResult, intNumOfRows, intCount
        Dim intMsgViewCount, intMsgActionCount
        Dim objProvider
        Set objProvider = Nothing
        Call SetSoftwareAndHardWareID()
		
		Dim intEulaStatus, strDevEntitlementId, strMachineGuid,strReleaseName, strBuildId, strEulaDt
		
		Dim m_objSMRegistry, ProviderEnumObject,bOldObfuscate
		
		Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
		Set m_objSMRegistry = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)

        intEulaStatus = m_objSMRegistry.RegQueryValue(SM_OOBE_V3_LOCATION, SM_CLIENT_OOBE_V3_EULASTATE)
        If  Len(Trim(intEulaStatus)) = 0 Then
            intEulaStatus = "0"
        End If

        strEulaDt = m_objSMRegistry.RegQueryValue(SM_CLIENT_APPINFO_LOCATION, SM_CLIENT_OOBE_V3V4_EULADATE)
        If  Len(Trim(strEulaDt)) = 0 Then
            strEulaDt = ""
        End If


		bOldObfuscate = m_objSMRegistry.Obfuscate	
		m_objSMRegistry.Obfuscate = true
        strDevEntitlementId = m_objSMRegistry.RegQueryValue("HKLM\SOFTWARE\McAfee\MSC\McInfo2\MAA2.0", "DeviceEntitlementId")
        If   Len(Trim(strDevEntitlementId)) = 0 Then
            strDevEntitlementId = ""
        End If
        m_objSMRegistry.Obfuscate = bOldObfuscate
		
		
		strMachineGuid = m_objSMRegistry.RegQueryValue("HKLM\SOFTWARE\Microsoft\Cryptography", "MachineGuid")
        If  Len(Trim(strMachineGuid)) = 0 Then
            strMachineGuid = ""
        End If

		strReleaseName = m_objSMRegistry.RegQueryValue("HKLM\SOFTWARE\McAfee\Msc", "ReleaseName")
        If  Len(Trim(strReleaseName)) = 0 Then
            strReleaseName = ""
        End If

         strBuildId = m_objSMRegistry.RegQueryValue("HKLM\SOFTWARE\McAfee\Msc\AppInfo\Substitute", "affid")
        If  Len(Trim(strBuildId)) > 0 Then
            Dim arrAffBuild: arrAffBuild = Split(strBuildId, "-", -1)
			
            If UBound(arrAffBuild) > 0 Then
                strBuildId = arrAffBuild(1)
            Else
                strBuildId = "0"
            End If
        End If
       
        strInstruXML        = "<InstrumentationXML SoftwareID='" & Replace(m_strSoftwareID, "_", "")  & "' HardwareID= '" &  m_strHardWareID 
		strInstruXML = strInstruXML & "' OSVersion='" & m_objSMSystemData.MachineOS & "' EulaStatus='" & Cstr(intEulaStatus) & "' DeviceEntitlementId='" & strDevEntitlementId & "' MachineGUID='" & strMachineGuid & "' CSPClientId='" & strCSPClientID  & "' WSSVersion='" & strReleaseName & "' EULADt='" & strEulaDt & "'>"
		
        strColumnNames      = "mv.MessageID, mv.acctid, mv.AffId, mv.lcid, mv.ProductKey, mv.CreatedDate, ml.ActionURL, ml.ActionTypeID, ml.CreatedDate "
	    strJoinTableNames	= "MessageViewHistory mv LEFT JOIN MessageLog ml ON mv.MessageViewHistoryID = ml.MessageViewHistoryID"    
	    strCondition        = " mv.SyncFlag = 0"
	    intMsgViewCount     = 1 
        If Not(objClientDB Is Nothing) Then
	        intResult = objClientDB.SelectRecords(strJoinTableNames, "count(*)", strCondition)
    
	        If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                intNumOfRows = objClientDB.GetNumberOfRows
	            If intNumOfRows > 0 Then
                   intResult = objClientDB.SelectRecords(strJoinTableNames, strColumnNames, strCondition)	                
                    If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                       For intCount = 0 To intNumOfRows -1
                            strMsgID        = objClientDB.GetData(0)
                            strAcctID       = objClientDB.GetData(1)
                            strAffID        = objClientDB.GetData(2)
                            strLCID         = objClientDB.GetData(3)
                            strProductKey   = objClientDB.GetData(4)
                            strMsgViewDt    = objClientDB.GetData(5)
                            strActionURL    = objClientDB.GetData(6)
                            strActionTypeID = objClientDB.GetData(7)
                            strMsgActionDt  = objClientDB.GetData(8)
                            If Len(Trim(strActionURL)) > 0 Then
                                strCID = GetCampaignID(strActionURL)
                            End If
                            If Len(Trim(strMsgActionDt)) > 0 Then
                                intMsgActionCount = 1
                            Else
                                intMsgActionCount = 0    
                            End If
''                            Set objProvider = objProviderEnum.GetProviderObject(intRetVal, intIndex)
''                            If intRetVal = SM_PROVIDER_OBJECT_CREATION_SUCCESS Then
''                                If IsObject(objProvider) And Not(objProvider Is Nothing) Then
''                                    Call GetAcctIDAndInstalledLCID(strProductKey, objProvider)         
''                                Else
''	                                ShowErrorMessage("SMInstrumentation:Unable to get AcctID and LCID")
''                                End If
''                            Else
''                            End If

							' Get the Cohort Rule executed for the SM ID
                            strRuleId = m_objSMRegistry.RegQueryValue(SM_CLIENT_COHORT_SM_LOCATION & "\" & strMsgID, SM_CLIENT_COHORT_ID)
							
                            If Len(Trim(strMsgXML))= 0 Then
                                strMsgXML = "<Message id='" & strMsgID & "' acctid='" & strAcctID & "'"
                            Else
                                strMsgXML = strMsgXML & "<Message id='" & strMsgID & "' acctid='" & strAcctID & "'"
                            End If
                            strMsgXML = strMsgXML &  " viewcount='" & intMsgViewCount & "' msgviewdt= '" & strMsgViewDt & "'"
                            strMsgXML = strMsgXML & " affid='" & strAffID &"' lcid='" & strLCID & "' buildid='" & strBuildId & "'"
							 If Len(Trim(strRuleId)) > 0 Then
                                strMsgXML = strMsgXML & " cohortDefinitionId='" & strRuleId & "'"
                            ELSE
                                strMsgXML = strMsgXML & " cohortDefinitionId='0'"
                            End If
                            strMsgXML = strMsgXML & " productkey='" & strProductKey &"' >"
                            strActionXML = ""
                            If intMsgActionCount > 0 Then     
                                strActionXML = strActionXML & "<Msglog actiondt='" & strMsgActionDt &"' actioncount='" & intMsgActionCount &"'"
                                strActionXML = strActionXML & " actiontype='" & strActionTypeID & "' cid='" & strCID &"' />"
                            End If
                            
                            If Len(Trim(strActionXML)) > 0 Then
                                strMsgXML = strMsgXML & strActionXML & "</Message>"
                            Else
                                strMsgXML = strMsgXML & "</Message>"    
                            End If
                        
                        If (intCount + 1) < intNumOfRows Then
                            objClientDB.PointToNextRow
                        End If
                        
                     Next
                   End If
                End If
	        End If
        End If
	    
	    strInstruXML = strInstruXML & strMsgXML & "</InstrumentationXML>"  
	    
	    If intNumOfRows = 0 Then
		    strInstruXML = ""	
	    End If
	    
	    
	    If Err.number > 0 Then
            ShowErrorMessage "Error is : " & Err.Description
        End If
        GetInstruXML = strInstruXML
    End Function

''''    Public Function GetInstruXML(objClientDB, objProviderEnum)
''''       ' On Error Resume Next
''''        Dim intRetVal, intIndex
''''        Dim strColumnNames, strJoinTableNames, strAffID
''''        Dim strCondition, strProductKey, strMsgActionDt
''''        Dim strMsgID, strActionURL, strMsgViewDt, strActionTypeID
''''        Dim strInstruXML, strMsgXML, strCID
''''        Dim intResult, intNumOfRows, intCount
''''        Dim intMsgViewCount, intMsgActionCount
''''        Dim objProvider
''''        Set objProvider = Nothing
''''        Call SetSoftwareAndHardWareID()
''''        strInstruXML        = "<InstrumentationXML SoftwareID='" & m_strSoftwareID & "' HardwareID= '" &  m_strHardWareID & "'>"
''''        strColumnNames      = "mv.MessageID, mv.AffId, mv.ProductKey, mv.CreatedDate, ml.ActionURL, ml.ActionTypeID, ml.CreatedDate "
''''	    strJoinTableNames	= "MessageViewHistory mv LEFT JOIN MessageLog ml ON mv.MessageViewHistoryID = ml.MessageViewHistoryID"    
''''	    strCondition        = " mv.SyncFlag = 0"
''''	    intMsgViewCount     = 1 
''''	    If Not(objClientDB Is Nothing) Then
''''	        intResult = objClientDB.SelectRecords(strJoinTableNames, "count(*)", strCondition)
''''	        
''''	        If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
''''	            intNumOfRows = objClientDB.GetNumberOfRows
''''	            If intNumOfRows > 0 Then
''''                   intResult = objClientDB.SelectRecords(strJoinTableNames, strColumnNames, strCondition)	                
''''                    If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
''''                        For intCount = 0 To intNumOfRows -1
''''                            strMsgID        = objClientDB.GetData(0)
''''                            strAffID        = objClientDB.GetData(1)
''''                            strProductKey   = objClientDB.GetData(2)
''''                            strMsgViewDt    = objClientDB.GetData(3)
''''                            strActionURL    = objClientDB.GetData(4)
''''                            strActionTypeID = objClientDB.GetData(5)
''''                            strMsgActionDt  = objClientDB.GetData(6)
''''                            If Len(Trim(strActionURL)) > 0 Then
''''                                strCID = GetCampaignID strActionURL
''''                            End If
''''                            If Len(Trim(strMsgActionDt)) > 0 Then
''''                                intMsgActionCount = 1
''''                            Else
''''                                intMsgActionCount = 0    
''''                            End If
''''                            Set objProvider = objProviderEnum.GetProviderObject(intRetVal, intIndex)
''''                            If intRetVal = SM_PROVIDER_OBJECT_CREATION_SUCCESS Then
''''                                If IsObject(objProvider) And Not(objProvider Is Nothing) Then
''''                                    Call GetAcctIDAndInstalledLCID(strProductKey, objProvider)         
''''                                Else
''''	                                ShowErrorMessage("SMInstrumentation:Unable to get AcctID and LCID")
''''                                End If
''''                            Else
''''                            End If
''''                            If Len(Trim(strMsgXML))= 0 Then
''''                                strMsgXML = "<Message id='" & strMsgID & "' cid='" & strCID & "'"
''''                            Else
''''                                strMsgXML = strMsgXML & "<Message id='" & strMsgID & "' cid='" & strCID & "'"
''''                            End If
''''                            strMsgXML = strMsgXML & " viewcount='" & intMsgViewCount & "' actiontype= '" & strActionTypeID & "'"
''''                            strMsgXML = strMsgXML & " actiondt='" & strMsgActionDt &"' actioncount='" & intMsgActionCount &"'"
''''                            strMsgXML = strMsgXML & " msgviewdt='" & strMsgViewDt &"' acctid='" & m_strAcctID &"'"
''''                            strMsgXML = strMsgXML & " affid='" & strAffID &"' lcid='" & m_strInstalledLCID &"'"
''''                            strMsgXML = strMsgXML & " productkey='" & strProductKey &"' />"
''''                        
''''                        If (intCount + 1) < intNumOfRows Then
''''                            objClientDB.PointToNextRow
''''                        End If
''''                        
''''                     Next
''''                       strInstruXML = strInstruXML & strMsgXML & "</InstrumentationXML>"   
''''                    End If
''''                End If
''''	            
''''	        End If
''''	    End If
''''	    
''''	    If Err.number > 0 Then
''''            ShowErrorMessage "Error is : " & Err.Description
''''        End If
''''        
''''        GetInstruXML = strInstruXML
''''    End Function
    
    Private Function GetCampaignID(strActionURL)
        Dim intCIDPos, cidValue
        Dim intEqualPos, intAmpPos
        Dim intUrlLen
        cidValue = ""
        ShowErrorMessage "getcid"
        '===temporary
       '' Dim testurl: testurl = "http://home.mcafee.com/root/campaign.aspx?cid=45856&abc=1"
        If Len(Trim(strActionURL)) > 0 Then
            intCIDPos = Instr(1,strActionURL,"cid") '=== get the cid position
            If CInt(intCIDPos) > 0 Then
                intEqualPos = Instr(intCIDPos,strActionURL,"=") '=== get the position of equal to
                If intEqualPos > 0 Then
                    If Instr(intEqualPos+1, strActionURL, "&", 1) > 0 Then 
                        intAmpPos = Instr(intEqualPos, strActionURL, "&", 1)    
                        cidValue = Mid(strActionURL, intEqualPos+1, (intAmpPos - (intEqualPos+1))) ' retrieve the variable value
                    Else
                        intUrlLen = Len(strActionURL)
                        cidValue = Mid(strActionURL, intEqualPos+1, (intUrlLen - (intEqualPos))) ' retrieve the variable value                            
                    End If    
                End If
            End If
        End If
        
        GetCampaignID = cidValue
    End Function
    
    Private Function SetSoftwareAndHardWareID()
        
        If Not(m_objSMSystemData Is Nothing) Then
            m_strSoftwareID = m_objSMSystemData.MachineSoftwareID
            m_strHardwareID = m_objSMSystemData.MachineHardWareID            
        Else
            m_strSoftwareID    = ""
            m_strHardwareID    = ""
        End If
    End Function
    
    Private Function GetAcctIDAndInstalledLCID(strProductKey, objProvider)
        Dim intInstalledAppCount, intAppIndex
        Dim intReturnAppInfo, intRetSubInfo
        Dim strAppID, strAppCode
        Dim strAcctID, strCurrentProductKey
        Dim objSMAppInfo, objSMSubscriptionData
        intInstalledAppCount        = 0
        Set objSMAppInfo            = Nothing
        Set objSMSubscriptionData   = Nothing
        
        If Len(Trim(strProductKey)) > 0 Then
            intInstalledAppCount = m_objSMAppData.InstalledApplicationSize
            If intInstalledAppCount > 0 Then
                For intAppIndex = 0 To intInstalledAppCount -1
                    Set objSMAppInfo = m_objSMAppData.GetAppInfoObject(intReturnAppInfo, intAppIndex)   
                    
                    If intReturnAppInfo = SM_APPINFO_OBJECT_CREATION_SUCCESS Then
                        strAppID      	=  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_ID)     
                        strAppCode    	=  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_NAME)
                        m_strInstalledLCID	=  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_INSTALL_LCID)
                        Set objSMSubscriptionData = New CSMSubscriptionData
                        intRetSubInfo = objSMSubscriptionData.InitializeSubMgrObject(objProvider, strAppID, strAppCode)
			            If intRetSubInfo = SM_SUBMGR_OBJECT_CREATION_SUCCESS Then
                            strCurrentProductKey = objSMSubscriptionData.ProductKey
                            If  strCurrentProductKey = strProductKey Then
                                m_strAcctID = objSMSubscriptionData.AccountID
                            Else
                                m_strAcctID = ""
                            End If    
                        End If
                    End If
                Next
            Else
                 ShowErrorMessage "SyncMessage:intInstalledAppCount is empty"
            End If
        End If
    End Function


End Class

Dim intProviderSize
Dim m_strAppID, m_strAppCode
Dim m_strInstalledLCID, m_strAppVersion
Dim m_objClientDB, m_objProviderEnum, m_objInstrumentation
Dim m_strOOBE, m_strActivationDate, m_strActivationDays, m_IsOOBEExists
Dim m_IsOOBEV3Exists, m_OOBEV3Enabled, m_OOBEV3EulaAccepted, m_OOBEV3V4EulaDate, m_OOBEV3RegStatus, m_IsOOBEV3Version, m_IsIprDone
Dim m_strDaysAfterInstall
Dim m_objFileSystem, m_objSystemInfo
Dim m_activationGracePeriod, m_expiryDate
Dim m_osID

Set m_objClientDB       = Nothing
Set m_objProviderEnum   = Nothing
Set m_objInstrumentation  = Nothing
Set m_strDaysAfterInstall = ""
intProviderSize = 0

Function AppendApplicationAttributes(msgReqXml)
    
    ShowErrorMessage "Calling Append Application Attributes"

    Dim xmlDocument 
    Set xmlDocument = CreateObject("Microsoft.XMLDOM")

   If xmlDocument Is Nothing Then
        ShowErrorMessage "Microsoft.XMLDOM is Nothing"
        Set xmlDocument = CreateObject("Msxml2.DOMDocument.6.0")        
    End If

    If xmlDocument Is Nothing Then
        ShowErrorMessage "Msxml2.DOMDocument.6.0 is Nothing"
        Set xmlDocument = CreateObject("Msxml2.DOMDocument.3.0")        
    End If

    If xmlDocument Is Nothing Then
        ShowErrorMessage "Xml DOM Is nothing"
        Exit Function
    End If

    ShowErrorMessage "Xml DOM Is not nothing"

    xmlDocument.loadXML(msgReqXml)

    If xmlDocument.childNodes.length = 0 Then   'MessageRequest
        ShowErrorMessage "Message request childNodes is empty"
        AppendApplicationAttributes = msgReqXml
        Exit Function
    End If

    If xmlDocument.childNodes(0).childNodes.length <= 1 Then    'Applications
        ShowErrorMessage "No Applications Nodes"
        AppendApplicationAttributes = msgReqXml
        Exit Function
    End If

    Dim mscAppVersion
    mscAppVersion = xmlDocument.childNodes(0).childNodes(1).attributes(1).value '<Applications core_app_version
    ''mscAppVersion = 11

    mscAppVersion = CDbl(mscAppVersion)
    If mscAppVersion < 11 Then
        ShowErrorMessage "Not an Emerald Build"
        AppendApplicationAttributes = msgReqXml
        Exit Function
    End If   

    Dim xml,nodeIdx    
    Dim isIPTEnabled,versionId
            
    isIPTEnabled = IsIPTInterfaceEnabled()
    versionId = GetAppConfigVersion()

    For nodeIdx = 0 To (xmlDocument.childNodes(0).childNodes(1).childNodes.length - 1)
        Set xml = xmlDocument.childNodes(0).childNodes(1).childNodes(nodeIdx)
        If Not xml Is Nothing Then
            xml.setAttribute "ipt",isIPTEnabled
            xml.setAttribute "versionid",versionId
        Else
            ShowErrorMessage "For loop exception"
        End If
    Next

    AppendApplicationAttributes = xmlDocument.xml

End Function

Function GetAppConfigVersion()
    
    Dim versionId,objRegistryMgr,result

    If intProviderSize > 0 Then
        Set objRegistryMgr = New CSMRegistry
        If Not objRegistryMgr Is Nothing Then
            result = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If result = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                versionId = objRegistryMgr.GetVersionId
            End If
        Else
            ShowErrorMessage "Registy Mgr is not present"
        End If
    End If

    GetAppConfigVersion = versionId

End Function

Function IsIPTInterfaceEnabled()
    Dim mscIpt,isInstCreated
    
    Set mscIpt = New CSMMscIpt
    isInstCreated = mscIpt.Initialize(m_objProviderEnum)
    
    If isInstCreated = SM_OBJECT_CREATION_FAILED Then
        ShowErrorMessage "CSMMscIpt object creation failed"
        IsIPTInterfaceEnabled = "false"
        Exit Function
    End If

    If mscIpt.IsIPTEnabled() Then
        ShowErrorMessage "IPT Enabled True"
        IsIPTInterfaceEnabled = "true"
    Else
        ShowErrorMessage "IPT Enabled False"
        IsIPTInterfaceEnabled = "false"
    End If  
End Function



Function Main()
    Dim bDbHitLimitReached, bInternetState, bVersionChkStatus
    Dim objHdnLCID, objHdnAppVersion
    Dim strMsgReqXML, strDBHitLimit, strServerUTCDtTime
    Dim intSystemStatus, intDBHitLimit
    Dim strMaintFlag, strSMVersion, strFilterXML
    Dim       strInstruXML, strInstruResponse
    Dim IsSystemInFullScreenMode, IsSystemInScreenSaverMode
    Call  SetGlobalObjects()    
    
    Set objHdnLCID          = document.forms(0).hdnLCID
    Set objHdnAppVersion    = document.forms(0).hdnApplicationVersion

    Dim m_objRegistryMgr, ProviderEnumObject
    Set ProviderEnumObject = m_objProviderEnum.ProviderEnumObject
    Set m_objRegistryMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
    ' Added for SM Optimization project 
    Set m_objFileSystem = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_FILESYSTEM_OBJECT)
    Set m_objSystemInfo = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_SYSTEM_MGR)
    ' End of SM Optimization project 
    '/// OOBE Starts
    m_IsOOBEExists = False
    m_IsOOBEExists = m_objRegistryMgr.RegKeyPresent(SM_OOBE_LOCATION)

    If m_IsOOBEExists Then
        m_strOOBE = m_objRegistryMgr.RegQueryValue(SM_OOBE_LOCATION, SM_OOBE_KEY_NAME)
        m_strActivationDate = m_objRegistryMgr.RegQueryValue(SM_OOBE_LOCATION, SM_OOBE_ACTIVATIONDATE_KEY_NAME)

        If m_strActivationDate <> "" Then
            Dim strServerDateTime, arrServerDateTime, dtCurrentDate
            strServerDateTime = document.forms(0).hdnCurrentUTCTime.Value
            If strServerDateTime <> "" Then
                arrServerDateTime = Split(strServerDateTime, " ")
                dtCurrentDate = CDate(arrServerDateTime(0))
            Else
                dtCurrentDate = Date
            End If

            m_strActivationDate = CDate(GetCorrectDate(m_strActivationDate))
            m_strActivationDays = DateDiff("d", m_strActivationDate, dtCurrentDate)

        End If
    End If
    '/// OOBE Ends

    '/// OOBE v3 Starts

    m_IsOOBEV3Exists = False
    m_IsOOBEV3Exists = m_objRegistryMgr.RegKeyPresent(SM_OOBE_V3_LOCATION)
    If m_IsOOBEV3Exists Then
        m_IsOOBEV3Version = m_objRegistryMgr.RegQueryValue(SM_OOBE_V3_LOCATION, SM_CLIENT_OOBE_V3_VERSION)
        If m_IsOOBEV3Version = SM_CLIENT_OOBE_V3_VERSION_NUMBER Then
            m_IsOOBEV3Exists = True
            m_OOBEV3Enabled = m_objRegistryMgr.RegQueryValue(SM_OOBE_V3_LOCATION, SM_CLIENT_OOBE_V3_ENABLE)
            If m_OOBEV3Enabled = "1" Then
                m_OOBEV3EulaAccepted = m_objRegistryMgr.RegQueryValue(SM_OOBE_V3_LOCATION, SM_CLIENT_OOBE_V3_EULASTATE)
                m_IsIprDone = m_objRegistryMgr.RegQueryValue(SM_CLIENT_APPINFO_LOCATION, SM_CLIENT_OOBE_V3_IPRDONE)
                'RR 3311 changes
                m_OOBEV3V4EulaDate = m_objRegistryMgr.RegQueryValue(SM_CLIENT_APPINFO_LOCATION, SM_CLIENT_OOBE_V3V4_EULADATE)
            End If
       Else
            m_IsOOBEV3Exists = False
        End If
    End If

    '/// OOBE v3 Ends

    'Enable IPT By Config XML'

        Dim bIsIPTUpdateEnabled
        bIsIPTUpdateEnabled = document.getElementById("hdnIsIptUpdateEnabled").value
        If Len(Trim(bIsIPTUpdateEnabled)) > 0 And bIsIPTUpdateEnabled = "1" Then
            UpdateIptByConfig()
        End If 
     'End Enable IPT By Config XML   

    ' Set Virus Alert Threat Level - Start
    Dim strThreatLevel, obfuscate, bSetValue
    Dim smConfigSet, smConfigKey, smConfig
    Dim telemetryForceFlow, DeviceIdcookieExists
    Dim affid, lcid, pkgid, affbuild
    Dim i
    Dim match
    
  
    DeviceIdcookieExists = "0"
    telemetryForceFlow = "0"
    match = false

    bSetValue = false
    strThreatLevel = Trim(document.forms(0).hdnThreatLevel.value)
    If Len(strThreatLevel) <> 0 And IsNumeric(strThreatLevel) Then
        obfuscate = m_objRegistryMgr.Obfuscate
        m_objRegistryMgr.Obfuscate = false
        bSetValue = m_objRegistryMgr.RegSetValue(REG_PATH_MSC_SECURITYREPORTS, REG_KEY_THREATMETER, CLng(strThreatLevel))
        m_objRegistryMgr.Obfuscate = obfuscate
    End If
    ' Set Virus Alert Threat Level - End


    strMaintFlag = document.forms(0).hdnMaintFlagValue.value
    strSMVersion = document.forms(0).hdnCurrentSMVersion.value
    strDBHitLimit = document.forms(0).hdnNumDbHit.value
    '===strServerUTCTime = document.forms(0).hdnCurrentUTCTime.Value

    Dim bIsEnableIPTBySegment
    Dim segmentRegistryMgr, intSegmentRegObj
    Dim strSegmentRegVersion
    Dim strSegmentAffId, strSegmentLCID, strSegmentSubStatus, strSegmentParentAffId
    Dim isUnRegistered, isTrial, isExpired
    Dim bSettingChanged
	Dim bSegmentMatch, bShowPosted

    bIsEnableIPTBySegment = document.getElementById("hdnEnableIPTbySegment").value

    If Len(Trim(bIsEnableIPTBySegment)) > 0 And bIsEnableIPTBySegment = "1" Then
        Set segmentRegistryMgr = New CSMRegistry
        If Not segmentRegistryMgr Is Nothing Then
           intSegmentRegObj = segmentRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
           If intSegmentRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
              strSegmentRegVersion = Trim(segmentRegistryMgr.GetRegKeyValueName(SM_CLIENT_INPRODUCT_LOCATION, SM_CLIENT_INPRODUCT_UPDATE_VERSION, False))
              strSegmentAffId = Split(Trim(segmentRegistryMgr.GetRegKeyValueName(SM_CLIENT_APPINFO_LOCATION, SM_SUBMGR_AFFID, False)), "-", -1)(0)				                             
              strSegmentLCID = segmentRegistryMgr.GetRegKeyValueName(SM_CLIENT_APPINFO_LOCATION, SM_CLIENT_IPT_LCID, False)
              strSegmentSubStatus = Trim(segmentRegistryMgr.GetRegKeyValueName(SM_CLIENT_INPRODUCT_LOCATION, SM_CLIENT_INPRODUCT_SUBSTATUS, False))
              strSegmentParentAffId = Trim(segmentRegistryMgr.GetRegKeyValueName(SM_CLIENT_INPRODUCT_LOCATION, SM_CLIENT_INPRODUCT_PARENTAFFID, False))
           End If
        End If


     
            Dim intProvSize, intAppDataObjResult
            Dim intRetVal, intIndex
            Dim objProvider
            Dim m_SMSystemData, m_SMAppData

            Dim intInstalledAppCount, intAppIndex,intSystemDataObjResult
            Dim intReturnAppInfo, intRetSubInfo
            Dim strAppID, strAppCode
            Dim strAcctID, strCurrentProductKey
            Dim objSMAppInfo, objSMSubscriptionData
            intInstalledAppCount        = 0
            intIndex    = 0
            intRetVal   = 0

            Set m_SMSystemData   = New CSMSystemData
            Set m_SMAppData      = New CSMAppData

            If Not(m_objProviderEnum Is Nothing) Then
                intProvSize = m_objProviderEnum.ProviderSize
                If intProvSize > 0 Then
                    intSystemDataObjResult   = m_SMSystemData.InitializeSystemObject(m_objProviderEnum)
                    Set objProvider = m_objProviderEnum.GetProviderObject(intRetVal, intIndex)
                    If intRetVal = SM_PROVIDER_OBJECT_CREATION_SUCCESS Then
                        If IsObject(objProvider) And Not(objProvider Is Nothing) Then
                            intAppDataObjResult = m_SMAppData.InitializeAppMgrObject(objProvider)
                        End If
                    End If
                End If
            End If

            Dim strAppFlags, strTrialInfo, strExpireDate

        intInstalledAppCount = m_SMAppData.InstalledApplicationSize
        If intInstalledAppCount > 0 Then
		    For intAppIndex = 0 To intInstalledAppCount -1
		    Set objSMAppInfo = m_SMAppData.GetAppInfoObject(intReturnAppInfo, intAppIndex)   
							
		    If intReturnAppInfo = SM_APPINFO_OBJECT_CREATION_SUCCESS Then
			    strAppID      	=  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_ID)     
			    strAppCode    	=  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_NAME)
                Set objSMSubscriptionData = New CSMSubscriptionData
			    intRetSubInfo = objSMSubscriptionData.InitializeSubMgrObject(objProvider, strAppID, strAppCode)
			    If intRetSubInfo = SM_SUBMGR_OBJECT_CREATION_SUCCESS Then
				    strAppFlags = objSMSubscriptionData.AppFlags
                    strTrialInfo = objSMSubscriptionData.TrialInfo
                    strExpireDate = objSMSubscriptionData.ExpiryDate
			    End If
		    End If
		    Next
        End If

        isUnRegistered = strAppFlags

        If (strTrialInfo = "1") Then
            isTrial = "1"
        Else 
            isTrial = "0"
        End if 
    
        Dim dtExpiryDt, dtCurrentDt
        Dim strSubStatus

        dtExpiryDt = CDate(strExpireDate)
        dtCurrentDt = Date

        If dtCurrentDt <= dtExpiryDt Then ' Not Expired
           isExpired = "0"
        Else
           isExpired = "1" ' Expired
        End If

        strSubStatus = isUnRegistered & "#" & isTrial & "#" & isExpired

        If (strSegmentSubStatus = "") Then 
            bSettingChanged = true
        Else 
            If (strSegmentSubStatus <>  strSubStatus) Then
                bSettingChanged = true
            End If 
        End If

    
        Dim strSegmentConfig, strSegmentConfigSet, strSegmentConfigKey, affIdDict, affParam, strAffKey
	    bSegmentMatch = false
	    strSegmentConfig = document.forms(0).hdnSegmentConfig.value
        Set affIdDict = CreateObject("Scripting.Dictionary")

        If Len(Trim(strSegmentConfig)) > 0 Then
            strSegmentConfigSet = split(strSegmentConfig, ";")
            For each affParam in strSegmentConfigSet
                If Len(Trim(affParam)) > 0 Then
                    strAffKey = Split(affParam, "|", -1)
                    If Len(Trim(strAffKey(0))) > 0 Then
                        affIdDict.Add strAffKey(0), strAffKey(0)
                    End If
                End If
            Next

            For i=0 to ubound(strSegmentConfigSet)
                strSegmentConfigKey = split(strSegmentConfigSet(i), "|")
                If (strSegmentParentAffId  = "") Then
                     bSegmentMatch = true
                     Exit For
                End If 
                If (affIdDict.Exists(Trim(strSegmentAffId)) AND Trim(strSegmentConfigKey(1)) = strSegmentLCID) Then
                        If (strSegmentRegVersion = "") OR (bSettingChanged = true) Then
                           bSegmentMatch = true
                           Exit For
                        ElseIf ((CDbl(strSegmentConfigKey(2)) - CDbl(strSegmentRegVersion)) > 0) OR (bSettingChanged = true) Then 
                            bSegmentMatch = true
	                        Exit For
                       End If
                 End If
                 
                 If (affIdDict.Exists(Trim(strSegmentParentAffId)) AND Trim(strSegmentConfigKey(1)) = strSegmentLCID) Then
                        If (strSegmentRegVersion = "") OR (bSettingChanged = true) Then
                           bSegmentMatch = true
                           Exit For
                        ElseIf ((CDbl(strSegmentConfigKey(2)) - CDbl(strSegmentRegVersion)) > 0) OR (bSettingChanged = true) Then 
                            bSegmentMatch = true
	                        Exit For
                       End If
                 End If
            Next
        End If 
    
        bShowPosted = false
    End If
    bDbHitLimitReached = true
    bInternetState = true
    If Len(Trim(strMaintFlag)) > 0 Then
        If CBool(strMaintFlag) = False Then
            bVersionChkStatus = CheckSMVersion(strSMVersion)
            If bVersionChkStatus Then           
                strServerUTCDtTime  = document.forms(0).hdnCurrentUTCTime.Value
                If (Len(Trim(strDBHitLimit)) > 0) Then
                    intDBHitLimit = CInt(strDBHitLimit)
                Else
                    intDBHitLimit  = SM_DB_HIT_LIMIT '=== default hit limit   
                End If

                bDbHitLimitReached = CheckIfDbHitLimitReached(intDBHitLimit, strServerUTCDtTime)
                
                If bDbHitLimitReached = False Then
                    strFilterXML = GetFilteringXML()
                    IsSystemInFullScreenMode = CheckIfInFullScreenMode() 
                    IsSystemInScreenSaverMode = CheckIfInScreenSaverMode()
                     
                    ''for sm telemetry
                    If(document.forms(0).hdnTelemetryConfigOn.Value = "1" And len(Trim(document.forms(0).hdnAffSitePkgConfig.value)) > 0) Then 
                        If ((IsSystemInFullScreenMode = SM_SYSTEM_NOT_IN_FULL_SCREEN_MODE) And (IsSystemInScreenSaverMode = SM_SYSTEM_NOT_IN_SCREEN_SAVER_MODE)) Then 
                            telemetryForceFlow = "0"      ''go to original flow, show sm
                        Else
                          If (document.forms(0).hdnDeviceIdCookieExists.Value = "0") Then
                            telemetryForceFlow = "1"      ''will not show sm. will return after setting PPrst cookie
                          End If
                        End If
                    End If
                   
                    If ((IsSystemInFullScreenMode = SM_SYSTEM_NOT_IN_FULL_SCREEN_MODE) And (IsSystemInScreenSaverMode = SM_SYSTEM_NOT_IN_SCREEN_SAVER_MODE)) OR (telemetryForceFlow = "1") Then 
                        intSystemStatus = CheckIfSystemInUse()
                        If intSystemStatus = SM_SYSMGR_SYSTEM_IN_USE Then
                            If (telemetryForceFlow <> "1") Then
                                strInstruXML = m_objInstrumentation.GetInstruXML(m_objClientDB, m_objProviderEnum, GetCSPClientId())   
                            End If

                            If Len(Trim(strInstruXML)) > 0 Then
                                bInternetState = GetInternetConnectedState(m_objProviderEnum)
                                If bInternetState Then
                                 ' strInstruResponse = SendInstruXML(strInstruXML)
                                  strInstruResponse = SendSmartInstruXML(strInstruXML)
                                Else
                                    Call HideWindow()
                                    ShowErrorMessage "Before getting InstruXML:Internet connection not available"
                                End If

                                If Len(Trim(strInstruResponse)) > 0 Then
                                    If CBool(strInstruResponse) = SM_INSTRU_STATUS_SUCCESS Then
                                        Call UpdateSyncFlag()
                                    End If
                                Else
                                    ShowErrorMessage "Response from webservice for instrumentation is empty"
                                End If
                            Else If (telemetryForceFlow <> "1") Then
                                ShowErrorMessage "strInstruXML is empty as there is nothing to be sent"
                            End If 
                            End If

                            m_strInstalledLCID = ""
                            strMsgReqXML = GetMessageRequestXML()
                            'strMsgReqXML = AppendApplicationAttributes(strMsgReqXML)
                            
                            If Not(objHdnLCID Is Nothing) Then
                                objHdnLCID.value = GetLCID
                            End If
                             

                            If Not(objHdnAppVersion Is Nothing) Then
                                objHdnAppVersion.value = GetAppVersion
                            End If
                            If (Len(Trim(strMsgReqXML)) > 0 ) Then
                                strMsgReqXML = Escape(strMsgReqXML)
                            End If
                            
                            If (Len(Trim(strFilterXML)) > 0) Then
                                strFilterXML = Escape(strFilterXML)
                            End If

                            '/// IPT - Suppress SM Alert
                            If IsIptInProgress() = True Then
                                Call InsertMsgSyncLogValues()
                                Call HideWindow()
                                Exit Function
                            End If
                            '/// End IPT - Suppress SM Alert

                            document.forms(0).hdnAffid.value = GetAffId()
                            bInternetState = GetInternetConnectedState(m_objProviderEnum)
                           
                           'CHECK FOR MATCH BEGIN
                           'IF telemetryForceFlow = "1" AND NOT MATCH THEN RETURN- DONT DO POST. 
                           If (document.forms(0).hdnTelemetryConfigOn.Value = "1") AND (telemetryForceFlow = "1") Then
                                If(Len(Trim(document.forms(0).hdnAffid.value)) > 0 And Len(Trim(document.forms(0).hdnLCID.value)) > 0 And Len(Trim(document.forms(0).hdnPkgId.value)) > 0) Then
                                    
                                    affbuild = split(document.forms(0).hdnAffid.value,"-")
                                    affid = affbuild(0)
                                    lcid = document.forms(0).hdnLCID.value
                                    pkgid = document.forms(0).hdnPkgId.value
                               
                                    smConfig = document.forms(0).hdnAffSitePkgConfig.value
                               
                                    smConfigSet = split(smConfig,"|")
                                    
                                    For i=0 to ubound(smConfigSet)
                                        smConfigKey = split(smConfigSet(i), ":",3)
                                        If(smConfigKey(0)=trim(affid) AND (smConfigKey(1)=trim(pkgid) OR smConfigKey(1)="-1") AND (smConfigKey(2)=trim(lcid) OR smConfigKey(2)="-1")) Then
                                            match = true
                                            Exit For
                                        End If
                                    Next
                                End If
                               
                                If (match = false) Then
                                       Exit Function
                                End If
                           End If
                           'CHECK FOR MATCH END
                            If bInternetState Then
                                If document.forms(0).name = "TestClient" Then
                                    document.forms(0).clientRequestXml.Value = Unescape(strMsgReqXML)
                                    document.forms(0).filterXml.value = Unescape(strFilterXML)
                                    ShowUI()
                                Else
                                    bShowPosted = true
                                    document.forms(0).hdnSegmentMatch.Value = bSegmentMatch
                                    document.forms(0).hdnTelemetryFlow.Value = telemetryForceFlow
                                    document.forms(0).hdnFilterXML.Value = strFilterXML
                                    document.forms(0).hdnClientRequestXML.value = strMsgReqXML
                                    document.forms(0).action = "ShowMessage.aspx"
                                    document.forms(0).submit
                                End If
                            Else
                                ShowErrorMessage "Internet connection not available"
                               ' Call HideWindow()
                            End If
                        Else
                            ShowErrorMessage "System is not in use"  
                            'Call HideWindow()
                        End If
                    Else
                        ShowErrorMessage "System is either in Gaming or Screensaver mode"  
                       ' Call HideWindow()
                    End If
                Else
                    ShowErrorMessage "DBHitLimit crossed " 
                   ' Call HideWindow()
                End If
        Else
            ShowErrorMessage "SMVersion check failed"  
            'Call HideWindow()
        End If    
        Else
            ShowErrorMessage "System is in maintenance mode"  
           ' Call HideWindow()
        End If   
    Else
        ShowErrorMessage "Maintenance mode value is empty"  
       ' Call HideWindow()
    End If
    
    Dim bForcePost
    bForcePost= false
    If Len(Trim(bIsEnableIPTBySegment)) > 0 And bIsEnableIPTBySegment = "1" Then
        If (bShowPosted = false) AND (bSegmentMatch = true) Then
            Dim  strClientRequest, strFilterRequest
        
            If (Len(Trim(strMsgReqXML)) > 0 ) Then
                strClientRequest = Escape(strMsgReqXML)
            Else 
                strClientRequest = Escape(GetMessageRequestXML())
            End If
        
            If (Len(Trim(strFilterXML)) > 0) Then
                strFilterRequest = Escape(strFilterXML)
            Else
               strFilterRequest = Escape(GetFilteringXML())
            End If
       
            bForcePost = true
            document.forms(0).hdnSegmentMatch.Value = bSegmentMatch
            document.forms(0).hdnTelemetryFlow.Value = "1"
            document.forms(0).hdnClientRequestXML.value = strClientRequest
            document.forms(0).action = "ShowMessage.aspx"
            document.forms(0).submit
        End if
    End if

    If (bShowPosted = false) AND (bForcePost = false) Then
       Call HideWindow()
    End if 

    If Not(m_objClientDB Is Nothing) Then
        Set m_objClientDB = Nothing
    End If

    If Not(m_objProviderEnum Is Nothing) Then
        Set m_objProviderEnum = Nothing
    End If
    Main = strMsgReqXML
End Function

  Function SetGlobalObjects()
    Dim intResult, intClientDBResult,intInstruResult
    
    Set m_objClientDB         = New CSMClientDB
    Set m_objProviderEnum     = New CSMProviderEnum
    Set m_objInstrumentation  = New CSMInstrumentation
    If Not(m_objProviderEnum Is Nothing) Then
        intProviderSize = m_objProviderEnum.ProviderSize
        If intProviderSize > 0 Then
            intClientDBResult = m_objClientDB.InitializeClientDBObject(m_objProviderEnum)
            intInstruResult = m_objInstrumentation.InitializeInstrumentationObject(m_objProviderEnum)    
        Else
            intResult = 0
            ShowErrorMessage "SetGlobalObjects:intProviderSize is  zero"  
        End If
    Else
        ShowErrorMessage "SetGlobalObjects:m_objProviderEnum is  empty"  
    End If
    
    If (intClientDBResult = SM_CLIENTDB_OBJECT_CREATION_SUCCESS) _
        And (intInstruResult = SM_INSTRUMENTATION_OBJECT_CREATION_SUCCESS)Then
        intResult = 1    
    End If
     
    SetGlobalObjects =  intResult  
 End Function

Function IsIptInProgress()
    Dim l_AppVersion 
    l_AppVersion = GetAppVersion()
    If l_AppVersion <> "" Then
        l_AppVersion = CDbl(l_AppVersion)
        If l_AppVersion >= 11 Then            
            Dim l_IptValidTillInSeconds
            l_IptValidTillInSeconds = GetIptValidTill()
            document.forms(0).hdnIptValidTill.Value = l_IptValidTillInSeconds
            If l_IptValidTillInSeconds <> "" Then
                l_IptValidTillInSeconds = CLng(l_IptValidTillInSeconds)
                Dim l_CurrentUtcInSeconds
                l_CurrentUtcInSeconds = document.forms(0).hdnCurrentUTCTimeInSeconds.Value
                l_CurrentUtcInSeconds = CLng(l_CurrentUtcInSeconds)
                If l_CurrentUtcInSeconds <= l_IptValidTillInSeconds Then
                    IsIptInProgress = True
                    Exit Function
                End If
            End If
        End If
    End If
    IsIptInProgress = False    
 End Function


 Function GetLCID()
    If Len(Trim(m_strInstalledLCID)) > 0 Then
        GetLCID = m_strInstalledLCID
    Else
        GetLCID = ""
    End If
End Function

Function GetAppVersion()
    Dim strAppVersion, strFinalAppVersion
    ShowErrorMessage "GetAppVersion:m_strAppVersion is  " & m_strAppVersion
    If Len(Trim(m_strAppVersion)) > 0 Then
        If Instr(1,m_strAppVersion,SM_APP_VERSION_SPLITTER) > 0 Then
            strAppVersion = Split(m_strAppVersion,SM_APP_VERSION_SPLITTER)
            strFinalAppVersion = strAppVersion(0) & SM_APP_VERSION_SPLITTER & strAppVersion(1)
        Else
            strFinalAppVersion = m_strAppVersion
        End If
        
        ShowErrorMessage "GetAppVersion:strFinalAppVersion is  " & strFinalAppVersion
        GetAppVersion = strFinalAppVersion
    Else
        GetAppVersion = ""
    End If
End Function

Function CheckIfDbHitLimitReached(intDBHitLimit, strServerUTCDtTime)
    Dim strTodaysDate
    Dim bExceededLimit
    Dim strHitLimitColumn, strHitLimitCondition
    Dim intChkResult, intFinalResult
    bExceededLimit = True
    If Not(m_objClientDB Is Nothing) Then
        bExceededLimit = m_objClientDB.ValidateDBHitStatus(intDBHitLimit, strServerUTCDtTime) 
    End If
    CheckIfDbHitLimitReached = bExceededLimit
End Function

Function CheckSMVersion(strSMServerVersion)
    Dim objRegistryMgr, objVersionMgr
   Dim bVersionChkStatus
    Dim intResult, intVersionResult, intTblCreationRslt, intVersionValidationRslt
    Dim strSMClientVersion
    ShowErrorMessage "CheckSMVersion"
    Set objVersionMgr = Nothing
    Set objRegistryMgr = Nothing
    bVersionChkStatus = False
   
        If intProviderSize > 0 Then
            Set objRegistryMgr = New CSMRegistry
    
            If Not objRegistryMgr Is Nothing Then
                intResult = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
                
                If intResult = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                    strSMClientVersion = objRegistryMgr.GetClientSMVersion
                    
                    Set objVersionMgr = New CSMDBVersionMgr
                    intVersionResult = objVersionMgr.InitializeVersionMgr(m_objProviderEnum)
                    If intVersionResult = SM_VERSIONMGR_OBJECT_CREATION_SUCCESS Then
                        If Len(Trim(strSMClientVersion)) = 0 Then
                            '=== Add here ClientSMVersionKey Validation call
                            strSMClientVersion = 0 '=== default value
                            intVersionValidationRslt = objVersionMgr.ExecuteSMVersionValidationScripts(objRegistryMgr, Trim(strSMClientVersion), Trim(strSMServerVersion))
                             strSMClientVersion = objRegistryMgr.GetClientSMVersion 
                            If Len(Trim(strSMClientVersion)) = 0 or strSMClientVersion = 0 Then
                                strSMClientVersion = 0
                            '===Else
                            '===    strSMClientVersion = objRegistryMgr.GetClientSMVersion  
                            End If
                        End If
                    
                        If Trim(strSMClientVersion) <>  Trim(strSMServerVersion) Then
                        '=== original location below for initializing objversionmgr
    '''                        Set objVersionMgr = New CSMDBVersionMgr
    '''                        intVersionResult = objVersionMgr.InitializeVersionMgr(m_objProviderEnum)
    '''                        If intVersionResult = SM_VERSIONMGR_OBJECT_CREATION_SUCCESS Then
                               intTblCreationRslt = objVersionMgr.ExecuteSMVersionScripts(objRegistryMgr, Trim(strSMClientVersion), Trim(strSMServerVersion))
                                If intTblCreationRslt = SM_CLIENTDB_ALL_TABLES_CREATED_SUCCESS Then
                                   '=== Call objRegistryMgr.SetClientSMVersion(strSMServerVersion)
                                    bVersionChkStatus = true
                                End If
    '''                        Else
    '''                            bVersionChkStatus = false    
    '''                        End If
                        Else
                            bVersionChkStatus = true    
                        End If 
                    Else
                        bVersionChkStatus = false    
                    End If    
                End If
            End If
        End If
   
    If Not(objRegistryMgr Is Nothing) Then
        Set objRegistryMgr = Nothing
    End If

    If Not(objVersionMgr Is Nothing) Then
        Set objVersionMgr = Nothing
    End If
    
    CheckSMVersion = bVersionChkStatus

End Function

Function CheckIfInFullScreenMode()
    Dim objSMUIContainer
    Dim bInFullScreenMode, intResult
    bInFullScreenMode = SM_SYSTEM_NOT_IN_FULL_SCREEN_MODE
    Set objSMUIContainer    = New CSMUIContainer
    
    
    If Not objSMUIContainer is Nothing Then
        intResult = objSMUIContainer.InitializeUIContainerObject()
        If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
            bInFullScreenMode = objSMUIContainer.GetIfSystemInFullScreenMode()  
        End If
    End If
        
     CheckIfInFullScreenMode = bInFullScreenMode   
 End Function

Function CheckIfInScreenSaverMode()
    Dim objSMUIContainer
    Dim bInScreenSaverMode, intResult
    bInScreenSaverMode = SM_SYSTEM_NOT_IN_SCREEN_SAVER_MODE
    Set objSMUIContainer    = New CSMUIContainer
    
    
    If Not objSMUIContainer is Nothing Then
        intResult = objSMUIContainer.InitializeUIContainerObject()
        If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
            bInScreenSaverMode = objSMUIContainer.GetIfSystemInScreenSaverMode()  
        End If
    End If
        
     CheckIfInScreenSaverMode = bInScreenSaverMode   
 End Function

Function CheckIfSystemInUse()
    Dim objSMSystemData
    Dim blnSystemInUse
    Dim intSystemInUse , intResult
    Dim objHdnClientExceptions
    
    Set objHdnClientExceptions  = Nothing
    Set objSMSystemData         = Nothing
    Set objSMSystemData         = New CSMSystemData

''    Set objHdnClientExceptions = document.forms(0).hdnClientSideExceptions
    
    If intProviderSize > 0 Then
        intSystemInUse = SM_SYSMGR_SYSTEM_IDLE
        ShowErrorMessage "IsSysteminuse"
        If  Not(objSMSystemData Is Nothing) Then
            intResult   = objSMSystemData.InitializeSystemObject(m_objProviderEnum)
            
            If intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then
                intSystemInUse  = objSMSystemData.IsSystemInUse
            Else
                If Not (objHdnClientExceptions Is Nothing) Then
                    objHdnClientExceptions.value = "Client database creation failed"
                End If
            End If
        Else
            ShowErrorMessage "System is not in use"
            intSystemInUse  = SM_SYSMGR_SYSTEM_IDLE
        End If    
    End If

    If Not(objSMSystemData Is Nothing) Then
        Set objSMSystemData = Nothing
    End If

   CheckIfSystemInUse   = intSystemInUse 
 End Function
  
'''' Function CreateNewTasks()
''''    Dim objTaskScheduler
''''    Dim intRetInitialize, intRetTaskStatus
''''    Dim strTaskSchedulerDefaultExecTime
''''    
''''    intRetInitialize                = 0
''''    intRetTaskStatus                = 0
''''    strTaskSchedulerDefaultExecTime = Now()  
''''    Set objTaskScheduler    = New CSMTaskScheduler
''''    
''''    If Not IsEmpty(objTaskScheduler) Then
''''        If Not objTaskScheduler Is Nothing Then
''''            intRetInitialize   = objTaskScheduler.InitializeTaskSchedulerObject()
''''            
''''            If intRetInitialize = SM_TASKSCHEDULER_OBJECT_CREATION_SUCCESS Then
''''                intRetTaskStatus = objTaskScheduler.CreateTaskScheduler(SM_TASKSCHEDULER_NAME, strTaskSchedulerDefaultExecTime, SM_TASKSCHEDULER_TRIGGER_FREQUENCY)    
''''            Else
''''                intReturnTaskStatus = SM_TASKSCHEDULER_TASK_CREATION_FAILED                
''''            End If
''''        Else
''''            intReturnTaskStatus = SM_TASKSCHEDULER_TASK_CREATION_FAILED
''''        End If
''''    Else
''''        intReturnTaskStatus = SM_TASKSCHEDULER_TASK_CREATION_FAILED
''''    End If    
'''' End Function
Function UpdateIptByConfig()  
    Dim strIptUpdateConfig
	strIptUpdateConfig = document.forms(0).hdnEnableIpt.value	
	Dim Params, Param, SubParams
	Dim Key,e_objDict
	Set e_objDict = CreateObject("Scripting.Dictionary")
	If Len(Trim(strIptUpdateConfig)) > 0 Then
		Params = Split(strIptUpdateConfig, "|", -1)
		For each Param in Params
			If Len(Trim(Param)) > 0 Then				
				SubParams  = Split(Trim(Param),";",-1)				
				Key = SubParams(0) & ";" & SubParams(1) & ";" & SubParams(2) & ";" & SubParams(3)                                
				e_objDict.Add Key,SubParams(4)
			End If
		Next	
	End if			
    Dim  objUpdateIptRegistryMgr,intUpdateIptRegObj,strIptRegAffId,strIptInstallLCID,strIptAppVersion,registryValue,appVersionList
    registryValue = ""       
    If e_objDict.Count > 0 Then
        Set objUpdateIptRegistryMgr = New CSMRegistry
        If Not objUpdateIptRegistryMgr Is Nothing Then                    
            intUpdateIptRegObj = objUpdateIptRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If intUpdateIptRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then                                                
                strIptRegAffId = Split(Trim(objUpdateIptRegistryMgr.GetRegKeyValueName(SM_CLIENT_APPINFO_LOCATION,SM_SUBMGR_AFFID,False)),"-",-1)(0)				                             
                strIptInstallLCID = objUpdateIptRegistryMgr.GetRegKeyValueName(SM_CLIENT_APPINFO_LOCATION,SM_CLIENT_IPT_LCID,False)                                                               
                appVersionList = Split(Trim(CStr(objUpdateIptRegistryMgr.GetRegKeyValueName(SM_CLIENT_APPINFO_LOCATION_SUBSTITUTE,SM_CLIENT_IPT_BUILD,False))),".",-1)		                                
                if ubound(appVersionList)=2 Then
                    strIptAppVersion = appVersionList(0) & ";" & appVersionList(1)                
                Else
                    strIptAppVersion=""
                End If
                registryValue = CStr(strIptInstallLCID) & ";" & CStr(strIptRegAffId) & ";" & CStr(strIptAppVersion)                                
            End if
        End if                 
    End if      
    If registryValue <> "" Then                   
        If e_objDict.Exists(registryValue) Then
           'change registry value based on the value send
           Dim strUpdateIpt
           Dim ProviderEnumObjectUpdateIpt
           Dim objUpdateIptRegistryMgr2
           Dim intCurrentValue
           Dim bSetValueIptUpdate
           bSetValueIptUpdate =  False
           strUpdateIpt = e_objDict(registryValue)
                If Len(Trim(strUpdateIpt)) > 0 Then                
                    If Not(m_objProviderEnum Is Nothing) Then
                        Set ProviderEnumObjectUpdateIpt = m_objProviderEnum.ProviderEnumObject
                        Set objUpdateIptRegistryMgr2 = ProviderEnumObjectUpdateIpt.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)

                        If Not (objUpdateIptRegistryMgr2 Is Nothing) Then                            
                            intCurrentValue = objUpdateIptRegistryMgr2.RegQueryValue(SM_CLIENT_INPRODUCT_LOCATION, SM_CLIENT_IPT_ENABLE_KEY)                            
                            If  (intCurrentValue = "") Or (CInt(strUpdateIpt) <> CInt(intCurrentValue)) Then
                                bSetValueIptUpdate = objUpdateIptRegistryMgr2.RegSetValue(SM_CLIENT_INPRODUCT_LOCATION, SM_CLIENT_IPT_ENABLE_KEY, CLng(strUpdateIpt))
                            End If
                        Else
                            ShowErrorMessage "UpdateIptByConfig: objUpdateIptRegistryMgr2 is empty"  
                        End if 
                    Else
                        ShowErrorMessage "UpdateIptByConfig: m_objProviderEnum is empty"  
                    End If
                Else
                        ShowErrorMessage "UpdateIptByConfig: strUpdateIpt is empty"  
                End If 
           'end change of registry value           
        End If
    End If
End Function

Function GetMessageRequestXML()
        'On Error Resume Next
        Dim objProvider
        Dim objSMAppData, objSMAppInfo
        Dim objSMSubscriptionData
        Dim intInstalledAppCount
        Dim arrInstalledApps
        Dim intIndex, intAppIndex
        Dim intReturn, intReturnAppInfo, intRetVal, intRetSubInfo 
        Dim strMessageRequestXML, strMachineConfigInfo
        Dim strSubInfo, strProviderName, strEachSubInfo
        Dim dtExpiredDate, dtCurrentDate, intOOBEExpireInDays, strExpired
        
        '=== Sub Mgmt II
        Dim flgGotPkgId
        
        Set objSMAppData            = Nothing
        Set objSMAppInfo            = Nothing
        Set objProvider             = Nothing
        Set objSMSubscriptionData   = Nothing
        intReturnAppInfo            = 0
        intInstalledAppCount        = 0
        intReturn                   = 0
        intRetVal                   = 0
        intRetSubInfo               = 0
        strSubInfo                  = ""
        strEachSubInfo              = ""
        strMessageRequestXML        = "<MessageRequest>"
        flgGotPkgId                             = "0"
        ShowErrorMessage "GetMessageRequestXML"
        
        ' ARPU Fix Start
        ShowErrorMessage "ARPU Fix "
        Dim  arpuRegistryMgr,intRegObj,bARPU
        bARPU = False
        Set arpuRegistryMgr = New CSMRegistry
        If Not arpuRegistryMgr Is Nothing Then
            intRegObj = arpuRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                Dim strRegAffid
                strRegAffid = arpuRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
                If Len(Trim(strRegAffid)) > 0 Then
                    If Trim(strRegAffid) = "0" Then
                        bARPU = True                    
                    End If
                Else
                   bARPU = True
                End If
            End If
        End If
        ' ARPU Fix End 

        ' SBT Fix Start
        ShowErrorMessage "SBT Fix "
        Dim  sbtRegistryMgr,intSBTRegObj,bSBT
        bSBT = False
        Set sbtRegistryMgr = New CSMRegistry
        If Not sbtRegistryMgr Is Nothing Then
            intSBTRegObj = sbtRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If intSBTRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                Dim strSBTRegAffid
                strSBTRegAffid = sbtRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
                If Len(Trim(strSBTRegAffid)) > 0 Then
                    If Trim(strSBTRegAffid) = "338" Then
                        bSBT = True    
                    End If
                Else
                   bSBT = False
                End If
            End If
        End If
        ' SBT Fix End                                     

  ' Dell Eula Check -- Begin                                                                                                                                                                             
                                       
  ShowErrorMessage "dell Fix "
  Dim  dellRegistryMgr,intdellRegObj
        
  Set dellRegistryMgr = New CSMRegistry
                                
  If Not dellRegistryMgr Is Nothing Then
            intdellRegObj = dellRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If intdellRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                
				Dim strdellRegAffid
				strdellRegAffid = dellRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
                                                
				Dim affbuild,affid
				affbuild = split(strdellRegAffid,"-")
				affid = affbuild(0)
                                                
				If Len(Trim(affid)) > 0 Then
                   If (Trim(affid) = "105") OR (Trim(affid) = "977") OR (Trim(affid) = "978") OR (Trim(affid) = "826") Then
                       
                                                                                   
					   Dim strdellRegOCEnable
					   Dim strdellRegOCEULAState
					   strdellRegOCEnable = dellRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","OCEnable",False)
					   strdellRegOCEULAState = dellRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","EULAState",False)
					   
                       If (Trim(strdellRegOCEnable) = "1") AND (Trim(strdellRegOCEULAState) = "0") Then
                             ' create the following two registry entries at below paths: HKLM\SOFTWARE\McAfee\MSC\OEM: OOBEDone,HKLM\SOFTWARE\McAfee\OemInfo:jeula
                             ' Update the OOBEDone value to 1 and jeula value to 1                                                                                                             
                                                                                                                 
                                                                                                                                                                                                                               
								Dim m_DellobjRegistryMgr, ProviderEnumObjectDell								
								Set ProviderEnumObjectDell = m_objProviderEnum.ProviderEnumObject							   
								Set m_DellobjRegistryMgr = ProviderEnumObjectDell.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
                                                                                                                
								Dim bDellOldObfuscate,bDellSetValue
								bDellSetValue = False
								bDellOldObfuscate = m_DellobjRegistryMgr.Obfuscate
								m_DellobjRegistryMgr.Obfuscate = False
								bDellSetValue = m_DellobjRegistryMgr.RegSetValue("HKLM\SOFTWARE\McAfee\OemInfo", "JEULA", CLng(1))
								bDellSetValue = m_DellobjRegistryMgr.RegSetValue("HKLM\SOFTWARE\McAfee\MSC\OEM", "OOBEDone", CLng(1))
								bDellSetValue = m_DellobjRegistryMgr.RegSetValue("HKLM\SOFTWARE\McAfee\PartnerCore\OEM", "OOBEDone", CLng(1))
								m_DellobjRegistryMgr.Obfuscate = bOldObfuscate
				
								' Call the mcautoreg.exe with /autoactivate 
								Dim  intDellLaunchAutoReg
								Dim  strDellMcAutoRegPath  
								Dim  strDellUpdtMcAutoRegPath
								
								strDellMcAutoRegPath = m_objFileSystem.GetCommonFilesPath() & "\mcafee\actwiz\mcautoreg.exe"  
								strDellUpdtMcAutoRegPath = Replace(strDellMcAutoRegPath,"Program Files","Progra~1")							   
								
								intDellLaunchAutoReg = m_objSystemInfo.RunProgramAndWait(strDellUpdtMcAutoRegPath, "/autoactivate")
							   
                        End If
					End If
                                                                  
                End If
            End If
        End If
 ' Dell Eula Check -- End
    

        ' AffId Swap Start
        Dim strAffIdSwapConfig, objDict
        Dim Params, Param
        Dim objKeyValue

                                Set objDict = CreateObject("Scripting.Dictionary")
        strAffIdSwapConfig = document.forms(0).hdnAffIdSwap.value
                                
        If Len(Trim(strAffIdSwapConfig)) > 0 Then
            Params = Split(strAffIdSwapConfig, ";", -1)
            For each Param in Params
                If Len(Trim(Param)) > 0 Then
                    objKeyValue = Split(Param, "=", -1)
                    objDict.Add objKeyValue(0), objKeyValue(1)
                End If 
            Next
        End if
        
        ShowErrorMessage "AffId Swap"
        Dim affIdRegistryMgr, intAffIdRegObj, bAffIdSwap
        Dim strAffIdRegAffId
        bAffIdSwap = False
        Set affIdRegistryMgr = New CSMRegistry
        If Not affIdRegistryMgr Is Nothing Then
            intAffIdRegObj = affIdRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
            If intAffIdRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                    strAffIdRegAffId = affIdRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
                    If Len(Trim(strAffIdRegAffId)) > 0 Then
                        If objDict.Exists(Trim(strAffIdRegAffId)) Then
                            bAffIdSwap = True    
                        End If
                   Else
                       bAffIdSwap = False
                    End If
            End If
        End If
              'AffId Swap End

                '======================================================
                'Acer Sku Change - Start                                
                '======================================================
                                Dim strAcerAffidArr
                                strAcerAffidArr=Split(strAffIdRegAffId,"-")
                                If strAcerAffidArr(0) ="662"  Then
                                                Dim isAcerSkuSwapped
                                                isAcerSkuSwapped=affIdRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\PartnerCore","IsSwapped",False)
                                                                If ((Len(Trim(isAcerSkuSwapped))=0) Or (isAcerSkuSwapped ="0")) Then
                                                                                Call AcerSwapSKUForMLSToMIS()
                                                                End If 
                                End If

                '======================================================
                'Acer Sku Change - End
                '======================================================
                        
        If intProviderSize > 0 Then
            strMachineConfigInfo        = GetMachineConfigInfo(m_objProviderEnum)
            'Added for Smart Messaging Optimization Project 
            Call GetUICntDetails()
            'End of Smart Messaging Optimization Project 
            strMessageRequestXML        = strMessageRequestXML & strMachineConfigInfo
            For intIndex = 0 To (intProviderSize -1)
                Set objProvider = m_objProviderEnum.GetProviderObject(intRetVal, intIndex)
                
                If intRetVal = SM_PROVIDER_OBJECT_CREATION_SUCCESS Then
                    
                    If IsObject(objProvider) And Not(objProvider Is Nothing) Then
                       
                        Set objSMAppData = New CSMAppData
                        intReturn = objSMAppData.InitializeAppMgrObject(objProvider)
                        strProviderName = objProvider.Getproperty(SM_PROVIDER_NAME)
                        
                        If Len(Trim(strProviderName)) = 0 Then
                            strProviderName = ""
                        End If
                        
                        If intReturn = SM_GET_INSTALLED_APPS_SUCCESS Then
                            intInstalledAppCount = objSMAppData.InstalledApplicationSize
                            If intInstalledAppCount > 0 Then
                                For intAppIndex = 0 To intInstalledAppCount -1
                                    Set objSMAppInfo = objSMAppData.GetAppInfoObject(intReturnAppInfo, intAppIndex)   
                                    
                                    If intReturnAppInfo = SM_APPINFO_OBJECT_CREATION_SUCCESS Then
                                        m_strAppID            =  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_ID)     
                                        m_strAppCode          =  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_NAME)
                                        If Trim(strProviderName) = "MSC" Then
                                            If m_strAppID = "MSC" Then
                                                m_strAppVersion    =  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_VERSION)
                                            End If
                                        Else
                                            m_strAppVersion    =  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_VERSION)
                                        End if
                                        m_strInstalledLCID    =  objSMAppInfo.GetProperty(MC_PROP_APPLICATION_INSTALL_LCID)
                                        Set objSMSubscriptionData = New CSMSubscriptionData
                                        intRetSubInfo = objSMSubscriptionData.InitializeSubMgrObject(objProvider, m_strAppID, m_strAppCode)
                                        If intRetSubInfo = SM_SUBMGR_OBJECT_CREATION_SUCCESS Then
                                            ' Msgbox "Above pkgid syncmessage3 intAppIndex: " & intAppIndex
                                           '=== Sub Mgmt II
                                           If flgGotPkgId = "0" Then
                                              ' Msgbox "before getting pkgid"
                                               ''document.frmSyncMessage.hdnPkgId.value  = objSMSubscriptionData.PkgID
                                               document.forms(0).hdnPkgId.value  = objSMSubscriptionData.PkgID                                              
                                              ' Msgbox "after getting pkgid: " & document.frmSyncMessage.hdnPkgId.value
                                               flgGotPkgId = "1"
                                           End If
                                            
                                            If Len(Trim(m_strAppVersion)) = 0 Then
                                                m_strAppVersion = ""
                                            End If                                           
                                            
                                            ' ARPU Fix Start
                                            If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                ShowErrorMessage "Registry value set call for 442 and objSMSubscriptionData.AffID is " & objSMSubscriptionData.AffID
                                                If (bARPU = True) And (objSMSubscriptionData.AffID = "442") Then
                                                    ShowErrorMessage "Inside set 442 call"
                                                    CAll arpuRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\queryparams"),  CStr("affid"), CStr("442"),False)
                                                    Call arpuRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),  CStr("affid"),  CStr("442"),False)
                                                    bARPU = False
                                                End If
                                            Else
                                                ShowErrorMessage "Registrymgr is not valid"
                                            End If
                                            Set arpuRegistryMgr = Nothing
                                            ' ARPU Fix End

                                            ' SBT Fix Start
                                            If intSBTRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                ShowErrorMessage "Registry value set call for 338 and objSMSubscriptionData.AffID is " & objSMSubscriptionData.AffID
                                                If (bSBT = True) And (objSMSubscriptionData.AffID = "0") Then
                                                    ShowErrorMessage "Inside set 338 call"
                                                    CAll sbtRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\queryparams"),  CStr("affid"), CStr("0"),False)
                                                    Call sbtRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),  CStr("affid"),  CStr("0"),False)
                                                    bSBT = False
                                               End If
                                            Else
                                                ShowErrorMessage "Registrymgr is not valid"
                                            End If
                                            Set sbtRegistryMgr = Nothing
                                            ' SBT Fix End

                                            'RR-143 : SBT migrated users switch back project 
                                            ShowErrorMessage "SBT Revert Fix"
                                            If objSMSubscriptionData.AffID = "338" Then
                                                Dim sbtRevertRegistryMgr, intSBTRevertRegObj
                                                Set sbtRevertRegistryMgr = New CSMRegistry
                                                If Not sbtRevertRegistryMgr Is Nothing Then
                                                    intSBTRevertRegObj = sbtRevertRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
                                                    If intSBTRevertRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                        Dim strSBTRevertRegAffid
                                                        strSBTRevertRegAffid = sbtRevertRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
                                                        If Len(Trim(strSBTRevertRegAffid)) > 0 Then
                                                            If Trim(strSBTRevertRegAffid) = "0" Then
                                                                ShowErrorMessage "Inside set 0 call"
                                                                Call sbtRevertRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\queryparams"),  CStr("affid"), CStr("338"),False)
                                                                Call sbtRevertRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),  CStr("affid"),  CStr("338"),False)
                                                            End If
                                                        End If
                                                    Else
                                                        ShowErrorMessage "Registrymgr is not valid"
                                                    End If
                                                End If
                                                Set sbtRevertRegistryMgr = Nothing
                                            End If  
                                            'RR-143 : SBT migrated users switch back project 
                                                                                                                                                                                
                                                                                                                                                                                
                                                                                                                                                                                
                                                                                                                                                                                
                                                                                                                                                                                

                                            ' AffId Swap Fix Start
                                            If intAffIdRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                ShowErrorMessage "AffId Swap Registry value set call and objSMSubscriptionData.AffID is " & objSMSubscriptionData.AffID
                                               'If (bAffIdSwap = True) And (objSMSubscriptionData.AffID = objDict.Item(strAffIdRegAffId)) Then
                                               If (bAffIdSwap = True) Then
                                                    Dim strTargetAffId
                                                    strTargetAffId = CStr(objDict.Item(strAffIdRegAffId))
                                                    ShowErrorMessage "Inside set " & objDict.Item(strAffIdRegAffid) & " call"
                                                    Call affIdRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\queryparams"),  CStr("affid"), strTargetAffId ,False)
                                                    Call affIdRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),  CStr("affid"), strTargetAffId ,False)
                                                    bAffIdSwap = False
                                                End If
                                            Else
                                               ShowErrorMessage "Registrymgr is not valid"
                                            End If

                                             Set affIdRegistryMgr = Nothing
                                           
                                            'AffId Swap Fix End

     
                                            '=============================================================
                                            'Acer Sku Swap - Update Mcsubdb Info - Start
                                            '=============================================================
                                                                                                                                                                                
                                                                                                                                                                                If objSMSubscriptionData.AffID = "662" Then
                                                                                                                                                                                  
                                                                                                                                                                                    Dim isAcerSkuSwappedMcsub
                                                Dim strMISAffIdMcsub,strMLSlangMcsub
                                                                                                                                                                                    Dim acerswapRegistryMgr
                                                                                                                                                                                                Dim intAcerswapRegistryObj
                                                Set acerswapRegistryMgr = New CSMRegistry
                                                                                                                                                                                                
                                                                                                                                                                                                If Not acerswapRegistryMgr Is Nothing Then
                                                               intAcerswapRegistryObj = acerswapRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
                                                                   If intAcerswapRegistryObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                                                                                                                                                                                                   isAcerSkuSwappedMcsub=acerswapRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\PartnerCore","IsSwapped",False)
                                                                        If isAcerSkuSwappedMcsub ="1" Then 
                                                                             strMLSlangMcsub=  acerswapRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","lang",False)
                                                                                                      
                                                                                                                                                                                                                                                                                                                  if((UCase(strMLSlangMcsub)="ZH-CN")) Then
                                                                                        strMISAffIdMcsub="662-287"
                                                                                                                                                                                                                              End If
                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                      if((UCase(strMLSlangMcsub)="ZH-TW")) Then
                                                                                      strMISAffIdMcsub="662-288"
                                                                                                                                                                                                                                                      End If
                                                                                                                                                                                                                                                        objSMSubscriptionData.AffID=strMISAffIdMcsub
                                                                                objSMSubscriptionData.PkgID="232"
                                                                                                                                                                                                                                ShowErrorMessage "Accer is Swapped SKU is 1 updated the MCSUBDB with Affid: "&  strMISAffIdMcsub
                                                                        End If
                                                                
                                                                                                                                                                                                                                                                                Else
                                                                                                                                                                                                                                                                                                ShowErrorMessage "Acer Registrymgr is not valid"                          
                                                                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                End If
                                                                                                                                                                                                                                Set acerswapRegistryMgr =Nothing
                                                                                                                                                                                End If                                    
                                                                                                                                                                                  
                                            
                                            '=============================================================
                                            'Acer Sku Swap - Update Mcsubdb Info - End
                                            '=============================================================
                                                                                                                                                      
                                            ' Sony EMEA Fix for incorrect installation
                                            ShowErrorMessage "Sony EMEA Fix for incorrect installation"
                                            If objSMSubscriptionData.AffID = "649" Then
                                                Dim sonyRegistryMgr, intSonyRegObj
                                                Set sonyRegistryMgr = New CSMRegistry
                                                If Not sonyRegistryMgr Is Nothing Then
                                                    intSonyRegObj = sonyRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
                                                    If intSonyRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
                                                                                                
                                                        ' Read the Affid and Termfix registy entry            
                                                        Dim strSonyRegAffid, strTermFix
                                                        strSonyRegAffid = sonyRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute","affid",False)
                                                                                                
                                                        strTermFix = sonyRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute","termfix",False)
                                                                                                
                                                        If Len(Trim(strTermFix)) = 0 Then
                                                            strTermFix = "0"
                                                        End If
                                                                                                
                                                        If Len(Trim(strSonyRegAffid)) = 0 Then
                                                            ' If Affid is empty and Termfix is 0 then read the details from Oeminfo\GUID and assign to MSC registry location
                                                            If Trim(strSonyRegAffid) = "" And strTermFix = "0" Then
                                                                ShowErrorMessage "Inside set 649 call"
                                                                Dim strSonyGUIDAffid, strSonyGUIDDomain, strSonyGUIDTerms, strSonyGUID, strSonyGUIDRegPath
                                                                Dim arrGUIDs 
                                                                Dim intSonyIndex, bGUIDExists
                                                                                                              
                                                                Dim m_objRegistryMgr, ProviderEnumObject
                                                                Set ProviderEnumObject = m_objProviderEnum.ProviderEnumObject
                                                                Set m_objRegistryMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
                                                                
                                                                'GUIDs of the package offered for Sony EMEA
                                                                '272 - 64664DDB-306E-4205-A5DB-04EA49A4D188, 273 -D7977B19-3445-4853-966D-44D9F4DFF658, 274-E2178637-5670-480F-AB51-F850B9B31594, 396-2311401C-4665-433C-B201-FABAA0EB5A2A
                                                                arrGUIDs = Array("64664DDB-306E-4205-A5DB-04EA49A4D188", "D7977B19-3445-4853-966D-44D9F4DFF658", "E2178637-5670-480F-AB51-F850B9B31594", "2311401C-4665-433C-B201-FABAA0EB5A2A")
                                                                                                                                                                           
                                                                'Check if GUID is present in the registry      
                                                                For intSonyIndex = 0 To UBound(arrGUIDs)
                                                                    bGUIDExists = m_objRegistryMgr.RegKeyPresent("HKLM\SOFTWARE\McAfee\OemInfo\" & arrGUIDs(intSonyIndex))
                                                                    'If present take that value for registy update
                                                                    If bGUIDExists = True Then
                                                                        strSonyGUID = arrGUIDs(intSonyIndex)
                                                                        Exit For
                                                                    End If
                                                                Next
                                                                                                              
                                                                'Form the registry path with the GUID
                                                                strSonyGUIDRegPath = CStr("HKLM\SOFTWARE\McAfee\OemInfo\" & Trim(strSonyGUID) )
                                                                ' Read the Affid, domain and terms from GUID location
                                                                strSonyGUIDAffid = sonyRegistryMgr.GetRegKeyValueName(strSonyGUIDRegPath & "\substitute","affid",False)
                                                                strSonyGUIDDomain = sonyRegistryMgr.GetRegKeyValueName(strSonyGUIDRegPath & "\substitute","domain",False)    
                                                                strSonyGUIDTerms = sonyRegistryMgr.GetRegKeyValueName(strSonyGUIDRegPath & "\substitute","terms",False)   
                                                                                                              
                                                                'Check the Terms value for quotes.  If present then affected machine
                                                                If InStr(strSonyGUIDTerms, Chr(34)) > 0 Then
                                                                                                                                
                                                                    'Replace the extra quotes in the Term value
                                                                    strSonyGUIDTerms = Replace(strSonyGUIDTerms, chr(34),"")
                                                                                                              
                                                                    'Update the GUID registry location with the affid and domain
                                                                    Call sonyRegistryMgr.SetRegKeyValueCstr( CStr(strSonyGUIDRegPath), CStr("affid"),  CStr(Trim(strSonyGUIDAffid)), True)
                                                                    Call sonyRegistryMgr.SetRegKeyValueCstr( CStr(strSonyGUIDRegPath), CStr("terms"),  CStr(Trim(strSonyGUIDTerms)), True)
                                                                    Call sonyRegistryMgr.SetRegKeyValueCstr( CStr(strSonyGUIDRegPath & "\substitute"), CStr("terms"),  CStr(Trim(strSonyGUIDTerms)), False)
                                                                    Call sonyRegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\OemInfo\substitute"), CStr("terms"),  CStr(Trim(strSonyGUIDTerms)), False)
                                                                                                                     
                                                                    'Update the OOBEDone value to 0 and then run the McOEMMgr.exe
                                                                    Dim bOldObfuscate,bSetValue
                                                                    bSetValue = False
                                                                    bOldObfuscate = m_objRegistryMgr.Obfuscate
                                                                    m_objRegistryMgr.Obfuscate = False
                                                                    bSetValue = m_objRegistryMgr.RegSetValue("HKLM\SOFTWARE\McAfee\MSC\OEM", "OOBEDone", CLng(0))
                                                                    m_objRegistryMgr.Obfuscate = bOldObfuscate
                                                                                                              
                                                                    ' Call the McOEMMgr.exe
                                                                    Dim intLaunchResult, intLaunchAutoReg
                                                                    Dim strMcOemMgrPath, strMcAutoRegPath
                                                                    
                                                                    strMcOemMgrPath = m_objSystemInfo.GetMISPPath() & "mcoemmgr.exe"
                                                                    strMcAutoRegPath = m_objSystemInfo.GetMISPPath() & "mcautoreg.exe"                                                                        

                                                                    ' Call the SystemInfo.Runprogram method to run the mcOEMMGR.exe
                                                                    intLaunchResult = m_objSystemInfo.RunProgramAndWait(strMcOemMgrPath, "")
                                                                    intLaunchAutoReg = m_objSystemInfo.RunProgram(strMcAutoRegPath, "")
                                                                    
                                                                    Call sonyRegistryMgr.SetRegKeyValueCstr(CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),CStr("termfix"),1,False)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            ' Sony EMEA Fix for incorrect installation End
                                                                                        
                                            If Len(Trim(strEachSubInfo)) > 0 Then
                                                strEachSubInfo  = strEachSubInfo & GetSubInfoXML(objSMSubscriptionData)
                                            Else
                                                strEachSubInfo = GetSubInfoXML(objSMSubscriptionData)
                                            End If
                                            
                                        End If
                                    End If
                                Next
                            Else
                                 ShowErrorMessage "SyncMessage:intInstalledAppCount is empty"
                            End If
                        End If
                        
                    End if
                End If
                If Len(Trim(strSubInfo)) = 0 Then
                    strSubInfo  = "<Applications type=" & chr(34) & strProviderName & chr(34) & " core_app_version=" & chr(34) & m_strAppVersion & chr(34)
                    'MsgBox("XML m_IsOOBEV3Exists " & m_IsOOBEV3Exists)
                    'MsgBox("XML m_OOBEV3Enabled " & m_OOBEV3Enabled)
                    If m_IsOOBEV3Exists And m_OOBEV3Enabled = "1" Then
                        'MsgBox("Inside XML")
                        'If m_OOBEV3EulaAccepted = "1" Then
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_OOBE & chr(34) & SM_CLIENT_OOBE_V3_CODE & chr(34)
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_EULA & chr(34) & m_OOBEV3EulaAccepted & chr(34)
                        'End If
                    ElseIf m_strOOBE = "1" Then
                        strSubInfo = strSubInfo & SM_CLIENT_CONFIG_OOBE & chr(34) & SM_CLIENT_OOBE_CODE & chr(34)
                        If m_strActivationDate <> "" Then
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_ACTIVATION_DYS & chr(34) & m_strActivationDays & chr(34)
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_ACTIVATION_DATE & chr(34) & m_strActivationDate & chr(34)
                        End If
                    End IF

                    ' IPT

                    Dim mscAppVersion
                    mscAppVersion = CDbl(m_strAppVersion)
                    If mscAppVersion >= 11 Then
                        
                        Dim isIPTEnabled,versionId, localeId
                        isIPTEnabled = IsIPTInterfaceEnabled()
                        versionId = GetAppConfigVersion()
                        localeId = GetLCID()
                        'versionId = "0#0#11#0#1"

                        strSubInfo = strSubInfo & SM_CLIENT_CONFIG_VERSIONID & chr(34) & versionId & chr(34)
                        Dim l_IsIpt : l_IsIpt = "0"
'                        If isIPTEnabled Then
'                            l_IsIpt = "1"
'                        End If

                        'If GetIEBrowserVersion() >= 7 Then
                        If GetBrowserVersion() >= 7 Then
                            l_IsIpt = "1"
                        End If
                        strSubInfo = strSubInfo & SM_CLIENT_CONFIG_IPT & chr(34) & l_IsIpt & chr(34)
                        strSubInfo = strSubInfo & SM_CLIENT_CONFIG_LOCALEID & chr(34) & localeId & chr(34)
                   End If   

                   If mscAppVersion >= 10 Then
                        ' Added for Force Upgrade 
                        Dim forceUpgrade
                        forceUpgrade = GetForceUpgrade()
                   
                       If forceUpgrade <> "" AND forceUpgrade <> "0"   Then
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_FORCEUPGRADE & chr(34) & forceUpgrade & chr(34)
                       End If
                       ' End of Force Upgrade changes 

                       ' Added for SM Optimization 
                       Dim dayAfterInstall
                       dayAfterInstall = GetLastDayofInstall()

                       strSubInfo = strSubInfo & SM_CLIENT_CONFIG_DAYSAFTERINSTALL & chr(34) & dayAfterInstall & chr(34)
                       ' End of SM Optimization 
                   End If

                    If mscAppVersion >= 11 AND m_IsOOBEV3Exists And m_OOBEV3Enabled = "1" Then
                        If m_OOBEV3EulaAccepted = "0" Then
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_ACTIVATIONGRACEPERIOD & chr(34) & m_activationGracePeriod & chr(34)
                        End If
                        If m_OOBEV3EulaAccepted = "1" Then
                            dtExpiredDate = CDate(m_expiryDate)
                            dtCurrentDate = Date
                            If dtCurrentDate <= dtExpiredDate Then ' Not Expired
                                strExpired = "0"
                                intOOBEExpireInDays = DateDiff("d",dtCurrentDate, dtExpiredDate)
                            Else
                                strExpired = "1" ' Expired
                                intOOBEExpireInDays = DateDiff("d",dtExpiredDate, dtCurrentDate)
                            End If
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_UNREGEXPIRED & chr(34) & strExpired & chr(34)
                            strSubInfo = strSubInfo & SM_CLIENT_CONFIG_UNREGEXPIREINDAYS & chr(34) & intOOBEExpireInDays & chr(34)

                            'RR 3311 changes
                            If m_OOBEV3V4EulaDate <> "" Then
                                m_OOBEV3V4EulaDate = CDate(GetCorrectDate(m_OOBEV3V4EulaDate))
                                m_strActivationDays = DateDiff("d", m_OOBEV3V4EulaDate, dtCurrentDate)
                                strSubInfo = strSubInfo & SM_CLIENT_CONFIG_ACTIVATION_DYS & chr(34) & m_strActivationDays & chr(34)
                            End If
                        End If
                    End If

                    ' Windows 8
                    If IsUserSegmentationSupported() Then
                        Dim intDisabledDate : intDisabledDate = GetDisabledDate()
                        If (Len(Trim(intDisabledDate)) <> 0) Then
                            Dim strServerTime : strServerTime = document.forms(0).hdnCurrentUTCTime.Value

                            If IsNumeric(intDisabledDate) And IsDate(strServerTime) Then
                                Dim strDisabledInDays : strDisabledInDays = DetermineDisabledDuration(intDisabledDate, strServerTime)
                                strSubInfo = strSubInfo & SM_CLIENT_CONFIG_DISABLEDDATE & Chr(34) & strDisabledInDays & Chr(34)
                            Else
                                ShowErrorMessage("Unable to calculate disabled date difference with server date" & strServerTime & " and client date " & strDisabledDate & ".")
                            End If    
                        End If                    
                    End If
                   
                    '   SA Freemium Project code start
                    Dim strSAFreemiumEligible
                    strSAFreemiumEligible       = "1"
                    strSAFreemiumEligible = GetSAFUICntDetails()
                    strSubInfo = strSubInfo & SM_CLIENT_CONFIG_SAFMESSAGE & chr(34) & strSAFreemiumEligible & chr(34)
                    ' SA Freemium Project code End

					 Dim strCohortIds
                    strCohortIds       = ""
                    strCohortIds = GetCohortRuleIds()
                    strSubInfo = strSubInfo & SM_CLIENT_CONFIG_COHORTID & chr(34) & strCohortIds & chr(34)
					
                    strSubInfo = strSubInfo & ">"
                End if
                strSubInfo = strSubInfo & strEachSubInfo
                'MsgBox(strSubInfo)
            Next
        Else
            ShowErrorMessage "SyncMessage:ProviderEnum object is empty or provider size is zero"
        End If  
        strSubInfo  = strSubInfo & "</Applications>" & vbcrlf
		' Add "<MessageID value=x />" if the MessageID is available in the registry.
		strSubInfo = GetMessageId(strSubInfo)
		strMessageRequestXML    = strMessageRequestXML & strSubInfo
        strMessageRequestXML    = strMessageRequestXML & "</MessageRequest>"
    
        Set objSMAppData            = Nothing
        Set objSMAppInfo            = Nothing
        Set objProvider             = Nothing
        '=== Set m_objProviderEnum         = Nothing
        Set objSMSubscriptionData   = Nothing

        'MsgBox(strMessageRequestXML)
        GetMessageRequestXML    = strMessageRequestXML
  End Function
  
  ' This function is used for A/B testing before the A/B framework was implemented.
  Function GetMessageId(strSubInfo)
	On Error Resume Next
	
	Dim messageIdRegistryMgr
	Set messageIdRegistryMgr = New CSMRegistry
	
	' This object is initialized in the SetGlobalObjects function but still needs to be checked that it hasn't already been set to Nothing
	If Not m_objProviderEnum Is Nothing Then
		If Not messageIdRegistryMgr Is Nothing Then
			Dim intMessageIdRegObj
			
			intMessageIdRegObj = messageIdRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
			If intMessageIdRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
				Dim messageId
				
				' Get registry value.
				messageId = messageIdRegistryMgr.GetRegKeyValueName(SM_CLIENT_MESSAGEID_LOCATION,SM_CLIENT_MESSAGEID_KEY,False)
				
				' Check if the value to the specified key exists.
				If Not IsNull(messageId) And Len(messageId) > 0 And Err.Number = 0 Then
					' Add the MessageID to the XML.
					strSubInfo = strSubInfo & "<MessageID value=" & Chr(34) & messageId & Chr(34) & "/>" & vbcrlf
					
					ShowErrorMessage "MessageID was found in the registry"
				Else
					ShowErrorMessage "MessageID was not found in the registry"
				End If
			End If
		End If
	Else
		ShowErrorMessage "m_objProviderEnum is null"
	End If
	
	Set messageIdRegistryMgr = Nothing

	GetMessageId = strSubInfo
  End Function
  
  Function GetAffId()
    Dim objRegistryMgr, intRegObj
    Set objRegistryMgr = New CSMRegistry

    If Not objRegistryMgr Is Nothing Then
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
            GetAffId = objRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)
            Exit Function
        End If
    End If
    GetAffId = "0"
  End Function

  ' Added for Force Upgrade 
  Function GetForceUpgrade()
    Dim objRegistryMgr, intRegObj
    Set objRegistryMgr = New CSMRegistry
    
    If Not objRegistryMgr Is Nothing Then
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
            GetForceUpgrade = objRegistryMgr.GetRegKeyValueName("HKLM\Software\McAfee\MSC\AppInfo\Substitute\QueryParams","forceupgrade",False)
            Exit Function
        End If
    End If
    GetForceUpgrade = ""
  End Function
  ' End of Force Upgrade

  ' Windows 8
  Function GetDisabledDate()
    Dim objRegistryMgr, intRegObj
    Set objRegistryMgr = New CSMRegistry
    
    If Not objRegistryMgr Is Nothing Then
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
            GetDisabledDate = objRegistryMgr.GetRegKeyValueName("HKLM\Software\McAfee\MSC", REG_KEY_DISABLEDDATE, False)
            Exit Function
        End If
    End If
    GetDisabledDate = ""

    Set objRegistryMgr = Nothing
  End Function

  Function DetermineDisabledDuration(disabledDate, serverTime)
    Dim dtmDisabledDate 
    Dim dtmServerTime

    dtmDisabledDate = EpochToDate(disabledDate)
    dtmServerTime = serverTime

    DetermineDisabledDuration = DateDiff("d", dtmDisabledDate, dtmServerTime) 
  End Function

  Function EpochToDate(intEpoch)
    EpochToDate = DateAdd("s", intEpoch, "01/01/1970 00:00:00")
  End Function

  Function DateToEpoch(strDate)
    DateToEpoch = DateDiff("s", "01/01/1970 00:00:00", strDate)
  End Function

  Function GetIEBrowserVersion()
    Dim objRegistryMgr, intRegObj
    Set objRegistryMgr = New CSMRegistry
    Dim l_RegistryValue
    If Not objRegistryMgr Is Nothing Then
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)
        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
            l_RegistryValue = objRegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\Microsoft\Internet Explorer\Version Vector","IE",False)
            If IsNumeric(l_RegistryValue) Then
                GetIEBrowserVersion = CDbl(l_RegistryValue)
                Exit Function
            End If
        End If
    End If
    GetIEBrowserVersion = "6"
  End Function

  Function GetMachineConfigInfo(m_objProviderEnum)
   On Error Resume Next
   ShowErrorMessage "GetMachineConfigInfo"
    Dim objSMSystemData
    Dim strMachineConfigXML
    
    Set objSMSystemData = Nothing
    Set objSMSystemData = New CSMSystemData
    intResult   = 0
    
    If  Not(objSMSystemData Is Nothing) Then
        
        intResult   = objSMSystemData.InitializeSystemObject(m_objProviderEnum)
        If intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then
'''            strMachineConfigXML = "<MachineConfig OS=" & chr(34)& objSMSystemData.MachineOS & chr(34) & " MemorySizeInMB=" & chr(34) & objSMSystemData.MachineMemorySize & chr(34)
'''            strMachineConfigXML = strMachineConfigXML & " ProcessorFamily=" & chr(34) & objSMSystemData.ProcessorFamily  & chr(34) & " ProcessorVersion=" & chr(34) & objSMSystemData.ProcessorVersion & chr(34)
'''            strMachineConfigXML = strMachineConfigXML   & " ProcessorSpeedGHZ=" & chr(34) & objSMSystemData.ProcessorSpeedGHZ & chr(34) & " NumberOfProcessors= " & chr(34) & objSMSystemData.NumberOfProcessors & chr(34)
'''            strMachineConfigXML = strMachineConfigXML   & " NetworkDrives=" & chr(34)& objSMSystemData.NetworkDrivesExist & chr(34) & " HardwareId=" & chr(34) & objSMSystemData.MachineHardwareID & chr(34)
'''            strMachineConfigXML = strMachineConfigXML   & " SoftwareId=" & chr(34) & objSMSystemData.MachineSoftwareID & chr(34) & "/>"
               strMachineConfigXML = "<MachineConfig " & SM_MACHINE_CONFIG_OS & chr(34)& objSMSystemData.MachineOS & chr(34) & SM_MACHINE_CONFIG_OS_VERSION_ID & chr(34)& objSMSystemData.MachineOsId & chr(34) & SM_MACHINE_CONFIG_MEM_SIZE_MB & chr(34) & objSMSystemData.MachineMemorySize & chr(34)
               strMachineConfigXML = strMachineConfigXML & SM_MACHINE_CONFIG_PROCESSOR_FMLY & chr(34) & objSMSystemData.ProcessorFamily  & chr(34) & SM_MACHINE_CONFIG_PROCESSOR_VRSN & chr(34) & objSMSystemData.ProcessorVersion & chr(34)
               strMachineConfigXML = strMachineConfigXML   & SM_MACHINE_CONFIG_PROCESSOR_SPEED_GHZ & chr(34) & objSMSystemData.ProcessorSpeedGHZ & chr(34) & SM_MACHINE_CONFIG_NO_OF_PROCESSORS & chr(34) & objSMSystemData.NumberOfProcessors & chr(34)
               strMachineConfigXML = strMachineConfigXML   & SM_MACHINE_CONFIG_NETWORK_DRIVES & chr(34)& objSMSystemData.NetworkDrivesExist & chr(34) & SM_MACHINE_CONFIG_HARDWARE_ID & chr(34) & objSMSystemData.MachineHardwareID & chr(34)
            strMachineConfigXML = strMachineConfigXML & SM_MACHINE_CONFIG_CSP_CLIENT_ID & chr(34) & GetCSPClientId & chr(34)
            strMachineConfigXML = strMachineConfigXML & SM_MACHINE_CONFIG_SOFTWARE_ID & chr(34) & objSMSystemData.MachineSoftwareID & chr(34) & "/>"

               '=== Sub Mgmt II.
               'document.frmSyncMessage.hdnWindowsId.value   = objSMSystemData.MachineSoftwareID
               'document.frmSyncMessage.hdnHardwareId.value  = objSMSystemData.MachineHardwareID
               
               document.forms(0).hdnWindowsId.value   = objSMSystemData.MachineSoftwareID
               document.forms(0).hdnHardwareId.value  = objSMSystemData.MachineHardwareID               
               
               m_osID = objSMSystemData.MachineOsId
               
        Else
            intResult = SM_SYSMGR_OBJECT_CREATION_FAILURE
            ShowErrorMessage "SyncMessage:Unable to Get Machine Config Info"      
        End If
    Else
        intResult = SM_SYSMGR_OBJECT_CREATION_FAILURE
   End If
    
    Set objSMSystemData = Nothing
    GetMachineConfigInfo    = strMachineConfigXML 
  End Function
   

  Function GetSubInfoXML(objSMSubscriptionData)
    Dim strSubInfoXML
    ShowErrorMessage "GetSubInfoXML"
    strSubInfoXML = ""
    

    If Not (objSMSubscriptionData Is Nothing) And Not ISEmpty(objSMSubscriptionData) Then
        Dim l_AccountId, l_RegistrationStatus, l_Trial, l_AffId
        
        l_AccountId = objSMSubscriptionData.AccountID
        
'        If l_AccountId = "" Then
'            l_AccountId = 0
'        Else
'            l_AccountId = CLng(l_AccountId)
'        End If

        l_RegistrationStatus = objSMSubscriptionData.AppFlags
        l_Trial = objSMSubscriptionData.TrialInfo
        l_AffId = objSMSubscriptionData.AffID

        If m_IsOOBEV3Exists And m_OOBEV3Enabled = "1" Then
            'MsgBox(l_RegistrationStatus)
            'If l_AccountId > 0 Then
            If m_IsIprDone = "1" Then
                l_RegistrationStatus = "0"
            Else If m_OOBEV3EulaAccepted = "1" Then ' If EULA Accepted
                    l_RegistrationStatus = "4"
                Else
                    l_RegistrationStatus = "7" ' If EULA not Accepted
                End If
            End If
        End If

        ' Added for Smart Messaging Optimization 
        If Len(Trim(m_strDaysAfterInstall)) > 0 Then
            m_strDaysAfterInstall  = m_strDaysAfterInstall & "," & objSMSubscriptionData.DaysAfterInstall
        Else
            m_strDaysAfterInstall  = objSMSubscriptionData.DaysAfterInstall
        End If
        ' End of Smart Messaging Optimization 

        m_activationGracePeriod = objSMSubscriptionData.ActivationGracePeriod
        m_expiryDate = objSMSubscriptionData.ExpiryDate

'        If l_AccountId = 0 Then
'            l_AccountId = ""
'        End If

        Dim strPkgName
        ' Encode Packge name before appending to XML string
        strPkgName = HTML_Encode(objSMSubscriptionData.PkgName)

        strSubInfoXML = "<Application " & SM_CLIENT_CONFIG_APPID    & chr(34)   & objSMSubscriptionData.AppId               & chr(34) & SM_CLIENT_CONFIG_APPCODE        & chr(34)   & objSMSubscriptionData.AppCode         & chr(34)    
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_PERPETUAL  & chr(34)   & objSMSubscriptionData.PerpetualInfo       & chr(34) & SM_CLIENT_CONFIG_TRIAL          & chr(34)   & l_Trial                               & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_ACCTID     & chr(34)   & l_AccountId                               & chr(34) & SM_CLIENT_CONFIG_EXP_DT         & chr(34)   & objSMSubscriptionData.ExpiryDate      & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_BACKEND    & chr(34)   & objSMSubscriptionData.Backend             & chr(34) & SM_CLIENT_CONFIG_SYNC_URL       & chr(34)   & objSMSubscriptionData.SyncURL         & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_WEB_SITE   & chr(34)   & objSMSubscriptionData.WebsiteURL          & chr(34) & SM_CLIENT_CONFIG_AFFID          & chr(34)   & l_AffId                               & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_BUILDID    & chr(34)   & objSMSubscriptionData.BuildID             & chr(34) & SM_CLIENT_CONFIG_LANG_CODE      & chr(34)   & objSMSubscriptionData.LangCode        & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_PKGID      & chr(34)   & objSMSubscriptionData.PkgID               & chr(34) & SM_CLIENT_CONFIG_PKGNM          & chr(34)   & strPkgName                            & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_RNWL_URL   & chr(34)   & objSMSubscriptionData.RenewalURL          & chr(34) & SM_CLIENT_CONFIG_REG_STATUS     & chr(34)   & l_RegistrationStatus                  & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_EMAIL      & chr(34)   & objSMSubscriptionData.Email               & chr(34) & SM_CLIENT_CONFIG_INSTALLED_DT   & chr(34)   & objSMSubscriptionData.InstalledDate   & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_LIC_QTY    & chr(34)   & objSMSubscriptionData.NumberOfLicenses    & chr(34) & SM_CLIENT_CONFIG_PROD_KEY       & chr(34)   & objSMSubscriptionData.ProductKey      & chr(34)
        strSubInfoXML = strSubInfoXML & SM_CLIENT_CONFIG_INSTALLED_DYS & chr(34)& objSMSubscriptionData.InstalledDays       & chr(34)
         
        strSubInfoXML = strSubInfoXML & "/>"
         
    End If  
    'MsgBox(strSubInfoXML)
    GetSubInfoXML   = strSubInfoXML
  End Function
  
  Function GetFilteringXML()
    Dim strMessageID, strNewMessageID, strActionValue, strMsgViewDt
    Dim strJoinTableNames, strColumnNames
    Dim strFilteringXML, strMsgXML
    Dim intNumOfRows, intResult, intCount
    Dim intMVCount
    '===Dim strMsgViewHistoryID, 
    ShowErrorMessage "filterxml"
    intNumOfRows               = 0
    intResult                  = SM_CLIENTDB_QUERY_EXEC_FAIL
    strColumnNames = "mv.MessageID, ml.ActionValue, mv.CreatedDate"
    strJoinTableNames      = "MessageViewHistory mv LEFT JOIN MessageLog ml ON mv.MessageViewHistoryID = ml.MessageViewHistoryID"    
    If Not(m_objClientDB Is Nothing) Then
        intResult = m_objClientDB.SelectRecords(strJoinTableNames, "count(*)", "")
        If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
            intNumOfRows = m_objClientDB.GetNumberOfRows()
        End If
        If CInt(intNumOfRows) > 0 Then
            intResult   = m_objClientDB.SelectRecords(strJoinTableNames, strColumnNames, "")
            If intResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
                strFilteringXML = "<FilterXML>"
                For intMVCount = 0 To intNumOfRows -1
                    '===strMsgViewHistoryID = m_objClientDB.GetData(0)
                    strMessageID    = m_objClientDB.GetData(0)
                    strActionValue = m_objClientDB.GetData(1)
                    strMsgViewDt    = m_objClientDB.GetData(2)
                    
                    If Len(Trim(strMsgXML)) = 0 Then
                        strMsgXML = "<Message id = " & chr(34) & strMessageID  & chr(34) &" MsgViewDt= " & chr(34) & strMsgViewDt & chr(34) & ">" 
                    Else
                        strMsgXML =  strMsgXML & "<Message id = " & chr(34) & strMessageID  & chr(34) &" MsgViewDt= " & chr(34) & strMsgViewDt & chr(34) & ">" 
                    End If
                    If Len(Trim(strActionValue)) > 0 Then
                        strMsgXML = strMsgXML & SetMessageLogValue(strActionValue) & "</Message>"
                    Else
                        strMsgXML = strMsgXML & "</Message>"
                    End If
                    
                    If (intMVCount + 1) < intNumOfRows Then
                        m_objClientDB.PointToNextRow()
                    End If
                Next
                strFilteringXML = strFilteringXML & strMsgXML & "</FilterXML>"
            End If
        End If
    End If
    'MsgBox ("GetFilteringXML " & strFilteringXML)
    GetFilteringXML = strFilteringXML
  End Function 
  
  Function SetMessageLogValue(strActionValue)
    Dim strMessageLogValue
    strMessageLogValue = "<MessageLog ActionValue = " & chr(34) & strActionValue & chr(34) &" />"
    SetMessageLogValue = strMessageLogValue
  End Function
  
  Function UpdateSyncFlag()
    Dim intUpdateResult
    Dim strUpdateTblName, strUpdateColumnName
    Dim strUpdateCondition, strColumnValues
    strUpdateColumnName = "SyncFlag"
    strColumnValues     = "1"
    strUpdateCondition  = "SyncFlag=0"
    
    If Not(m_objClientDB Is Nothing) Then
        strUpdateTblName= ""
        intUpdateResult = m_objClientDB.UpdateRecord(SM_MESSAGE_VIEW_HISTORY_TABLE_NAME,strUpdateColumnName, strColumnValues, strUpdateCondition)
        
        If intUpdateResult = SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL Then
            intUpdateResult = m_objClientDB.UpdateRecord(SM_MESSAGE_LOG_TABLE_NAME,strUpdateColumnName, strColumnValues, strUpdateCondition)
        End If
    End If
  End Function

Private Function GetCorrectDate(strDate)
    Dim strFinalDate,strDay, strMonth, strYear 
    If Len(Trim(strDate)) > 0 Then
        strYear     = Mid(strDate,1,4)
        strMonth    = Mid(strDate,5,2)
        strDay      = Mid(strDate,7,2)
        If (strMonth <> "") And (strDay <> "") And (strYear <> "") Then
            If Len(strMonth) <> 2 Then
                strMonth = "0" & strMonth
            End If
            If Len(strDay) <> 2 Then
                strDay = "0" & strDay
            End If
            strFinalDate  = strMonth & "/" & strDay & "/" & strYear
        Else
            strFinalDate    = ""
        End If
    Else
    strFinalDate    = ""
    End If

    GetCorrectDate = strFinalDate
End Function

Function InsertMsgSyncLogValues()
    Dim  objHdnMsgCnt
    Dim bInternetState
    Dim intMsgCount, intInsertResult, intActionTypeID
    Dim strTodaysDate, strServerUTCDtTime
    Dim strColumnValues, strDefaultResultTypeID
    Set objHdnMsgCnt    = Nothing
  
    strServerUTCDtTime = document.forms(0).hdnCurrentUTCTime.Value

    If Len(Trim(strServerUTCDtTime)) = 0 Then
        strTodaysDate = FormatDateTime(Now(),4) '=== 4 means (24 hour format) 
    Else
        strTodaysDate  = strServerUTCDtTime
    End If

    intInsertResult = SM_CLIENTDB_QUERY_EXEC_FAIL
    
    Set objHdnMsgCnt = document.getElementById("hdnMessageCount")
    If Not(objHdnMsgCnt is Nothing) Then
        If Len(Trim(objHdnMsgCnt.value)) > 0 Then
            intMsgCount = CInt(Trim(objHdnMsgCnt.value))
        Else
            intMsgCount = 0
        End If
    End If

    if intMsgCount > 0 Then
        intActionTypeID = 2
    Else
        intActionTypeID = 1
    End If

    If Not(m_objClientDB Is Nothing) Then 
        strColumnValues = "'" & SM_CLIENTDB_SYNC_MESSAGE_URL &"'," & intActionTypeID & "," & intMsgCount & ",'" & SM_DEFAULT_SYNC_FLAG_VALUE &"','" & strTodaysDate & "'" 
        intInsertResult = m_objClientDB.InsertRecord(SM_MESSAGE_SYNC_LOG_TABLE_NAME,SM_MESSAGE_SYNC_LOG_COLUMNS, strColumnValues)
    Else
        ShowErrorMessage("ClientDB Object is Empty")       
    End If
    InsertMsgSyncLogValues = intInsertResult
End Function


' Added for SM Optimization project 
 ' This method reads the McSubDB to get the installed date of the applications and gets the last installed application's install date
Function GetLastDayofInstall()
    Dim arrDaysAfterInstall
    Dim intMin, intIndex
    m_strDaysAfterInstall = m_strDaysAfterInstall
    
    arrDaysAfterInstall = Split(m_strDaysAfterInstall, ",")
    intMin = CInt(arrDaysAfterInstall(0))
   
    For intIndex = 1 to UBound(arrDaysAfterInstall)
        If intMin > CInt(arrDaysAfterInstall(intIndex)) Then
           intMin = CInt(arrDaysAfterInstall(intIndex))
        End If
    Next
    
    GetLastDayofInstall = intMin
End Function

Function GetUICntDetails()
   Dim objRegistryMgr, intRegObj
    Dim strUICntPath, strUICntArgs
    Dim intResult

    Set objRegistryMgr = New CSMRegistry

    If Not objRegistryMgr Is Nothing Then
        ' Get the Registry object to get the UIContainer EXE path and the command Args to be passed
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)

        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
            ' Get the UICntpath
            strUICntPath = objRegistryMgr.GetRegKeyValueName(SM_CLIENT_UICNTPATH_LOCATION, SM_CLIENT_UI_CONTAINER_PATH,True)

            strUICntPath = m_objFileSystem.GetShortPathName(strUICntPath)

            If m_IsOOBEV3Exists And m_OOBEV3Enabled = "1" Then ' If OOBE V3 get the EULA app location
                strUICntArgs = " " & m_objSystemInfo.GetMISPPath() & "oobe\mcocaw.dll"  'c:\Program Files\McAfee\MSC\OOBE\mcocaw.dll
            Else
                ' Get the ActWiz Args value
                ' Append SM=1 to track Omniture Tagging from app.mcafee.com
                strUICntArgs = " " & m_objSystemInfo.GetMISPPath() & "mcactwiz.dll sm=1 "
            End If

            document.forms(0).hdnUICntPath.value = strUICntPath
            document.forms(0).hdnUICntArgs.value = strUICntArgs

        End If
    End If
End Function
'End of Smart Messaging Optimization 

' =============================================================================================================
' Windows8-Specific Functions
' =============================================================================================================
Function IsUserSegmentationSupported()
    ' 11.6 (and higher) on Windows 8 supports the extended protection period
    IsUserSegmentationSupported = IsCoreApplicationSupportsExtendedProtection And IsOsSupportsExtendedProtection
End Function

Function IsCoreApplicationSupportsExtendedProtection()
     Dim dblCoreAppVersion : dblCoreAppVersion = CDbl(m_strAppVersion)

     ' 11.6 and higher supports an extended protection period (on Windows 8)
     IsCoreApplicationSupportsExtendedProtection = dblCoreAppVersion >= 11.6
End Function

Function IsOsSupportsExtendedProtection()
    ' Windows 8 (x86 and x64) and Windows 8.1 (x86 and x64) are currently the only oses that supports an extended protection period
    If m_osID = 12 Or m_osID = 13 Or m_osID = 14 Or m_osID = 15 Or m_osID = 16 Or m_osID = 17 Then
        IsOsSupportsExtendedProtection = True
    Else
        IsOsSupportsExtendedProtection = False
    End If
End Function

Function IsOsWin8OrGreater()
    Dim objSMSystemData
    Dim intResult
    Set objSMSystemData         = Nothing
    Set objSMSystemData         = New CSMSystemData
    IsOsWin8OrGreater = False
    If  Not(objSMSystemData Is Nothing) Then
                    intResult = objSMSystemData.InitializeSystemObject(m_objProviderEnum)
                    If intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then
                                m_osID = objSMSystemData.MachineOsId
                '12 & 13 are OS id for Windows 8
                '14 & 15 are for Windows 8.1
                If m_osID = 12 Or m_osID = 13 Or m_osID = 14 Or m_osID = 15 Or m_osID = 16 Or m_osID = 17 Then
                    IsOsWin8OrGreater = True
                End If
                    Else
                intResult = SM_SYSMGR_OBJECT_CREATION_FAILURE
                ShowErrorMessage "SyncMessage:Unable to Get OS Id in IsOsWin8OrGreater"      
        End If
    Else
        intResult = SM_SYSMGR_OBJECT_CREATION_FAILURE   
    End If
End Function

Function GetCohortRuleIds()
    Dim objRegistryMgr, intRegObj
	Dim strCohortId
	Set objRegistryMgr = New CSMRegistry

    If Not objRegistryMgr Is Nothing Then
        ' Get the Registry object to get the UIContainer EXE path and the command Args to be passed
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)

        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then
		 strCohortId = objRegistryMgr.GetRegKeyValueName(SM_CLIENT_COHORT_RULE_LOCATION, SM_CLIENT_COHORT_ID, False)
		End If
	End If
	GetCohortRuleIds = strCohortId
End Function
' =============================================================================================================
' Site Advisor Freemium-Specific Functions Start
' =============================================================================================================
Function GetSAFUICntDetails()
    Dim objRegistryMgr, intRegObj
    Dim strSAFUICntPath, strSAFUICntArgs
    Dim intResult
    Dim strSAFreemiumVersion, strSAFreemiumInstallFolder, strSAFreemiumSMOfferAccepted
    Dim intSaFreemiumApplicable
    Dim version
    Dim isVersionEligible
    Dim m_Result
    intSaFreemiumApplicable = "1"
    document.forms(0).hdnSAFUICntPath.value = ""
    document.forms(0).hdnSAFUICntArgs.value = ""

    Set objRegistryMgr = New CSMRegistry

    If Not objRegistryMgr Is Nothing Then
        ' Get the Registry object to get the UIContainer EXE path and the command Args to be passed
        intRegObj = objRegistryMgr.InitializeRegistryObject(m_objProviderEnum)

        If intRegObj = SM_REGISTRY_OBJECT_CREATION_SUCCESS Then

             'Read SM Offer Acceptance Status and Exit If already accepted
             strSAFreemiumSMOfferAccepted = objRegistryMgr.GetRegKeyValueName(SM_CLIENT_SAFREEMIUM_LOCATION, SM_CLIENT_SAFSMOFFER_STATUS, False)
             If Len(Trim(strSAFreemiumSMOfferAccepted)) > 0 AND strSAFreemiumSMOfferAccepted = "1" Then
                GetSAFUICntDetails = "1"
                Exit Function
             End If
            
            'Read Perform SA Freemium Version Validation > 3.6.4 
             strSAFreemiumVersion = objRegistryMgr.GetRegKeyValueName(SM_CLIENT_SAFREEMIUM_LOCATION, SM_CLIENT_SAF_VERSION, False)
             If Not Len(Trim(strSAFreemiumVersion)) > 0 Then
                GetSAFUICntDetails = "1"
                Exit Function
             End If
             
             isVersionEligible = IsSAVersionEligible(strSAFreemiumVersion)
             If (isVersionEligible = False)  Then
                GetSAFUICntDetails = "1"
                Exit Function
             End If
             
             'Read Perform SA Freemium EXE availability and populated hidden fields with location and argument values
            strSAFreemiumInstallFolder = objRegistryMgr.GetRegKeyValueName(SM_CLIENT_SAFREEMIUM_LOCATION, SM_CLIENT_SAINSTALL_FOLDER, False)
            If Not Len(Trim(strSAFreemiumInstallFolder)) > 0 Then
                GetSAFUICntDetails = "1"
                Exit Function
            End If 

            strSAFUICntArgs = " " & SM_CLIENT_SAFUICNTARGS
            strSAFUICntPath = strSAFreemiumInstallFolder & "\" & SM_CLIENT_SAFINSTALLER_FILE
            
            ' Validate SA Freemium file exists otherwise exit
            If Not m_objFileSystem.IsFile(strSAFUICntPath) Then
                GetSAFUICntDetails = "1"
                Exit Function
            End If
            
             ' Set Site Advisory Freemium Exe location and command parameters
            document.forms(0).hdnSAFUICntPath.value = strSAFUICntPath
            document.forms(0).hdnSAFUICntArgs.value = strSAFUICntArgs
            intSaFreemiumApplicable = "0"
            
        End If
    End If
    
    GetSAFUICntDetails = intSaFreemiumApplicable
    set objRegistryMgr = Nothing
End Function


Function IsSAVersionEligible(strSACurrentVersion)
    Dim isEligible
    Dim strSACurrentVersionArray
    isEligible = False
    strSACurrentVersionArray = Split(strSACurrentVersion , ".")
                
                'Check SA version greater than 3.6.4.000
    If Not IsNull(strSACurrentVersionArray) Then
                                If CInt(strSACurrentVersionArray(0)) > 3 Then
                                                isEligible = True
                                ElseIf CInt(strSACurrentVersionArray(0)) = 3 Then
                                                If CInt(Left(strSACurrentVersionArray(1),1)) > 6 Then
                                                                isEligible = True
                                                ElseIf CInt(Left(strSACurrentVersionArray(1),1)) = 6 Then
                                                                If CInt(Left(strSACurrentVersionArray(2),1)) >= 4 Then
                                                                                isEligible = True
                                                                Else
                                                                                isEligible = False
                                                                End If       
                                                Else
                                                                isEligible = False
                                                End If
                                Else
                                                isEligible = False
                                End If 
                End If

    IsSAVersionEligible = isEligible
End Function
' =============================================================================================================
' Site Advisor Freemium-Specific Functions End
' =============================================================================================================


'===========================================================================================
'                           Acer SKU swap Fix Start
'===========================================================================================

Function AcerSwapSKUForMLSToMIS()

                    Dim RegistryMgr
        Dim strAffId
        Dim strLanguage
        Dim strSwappedSku
        Dim RegistryStatus
        Dim isAcerSkuSwapped
        Dim strAffIdArray
        Set RegistryMgr = New CSMRegistry

                                ShowErrorMessage "Acer: Swap Affid call start"

            If Not RegistryMgr Is Nothing Then
            
                  RegistryStatus = RegistryMgr.InitializeRegistryObject(m_objProviderEnum)
                        
                    If RegistryStatus = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then

                     ShowErrorMessage "Acer: Reading the registry values"

                                   strAffId= RegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","affid",False)

                                   strLanguage= RegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams","lang",False)
                                                                                                                                                                                                 
                                                                                                                                                                                                                                                                Select Case Ucase(strLanguage)
                                                                    case "ZH-CN"  
                                                                        If strAffId="662-185" Then
                                                                            strSwappedSku="662-287"
                                                                        End If

                                                                    case "ZH-TW"  
                                                                        If strAffId="662-186" Then
                                                                            strSwappedSku="662-288"
                                                                       End If
             
                                                                End Select

                                                                                                                                                                                                                                                If (Len(Trim(strSwappedSku))>0) Then
                                                                 ShowErrorMessage "Acer: Swap Affid set call - Old Sku: " & strAffId & " New Sku:" & strSwappedSku & " Language:"+ strLanguage
                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\queryparams"),  CStr("affid"), CStr(strSwappedSku),False)
                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"),  CStr("affid"),  CStr(strSwappedSku),False)

                                                                                                                                                                                                                                                                'Update the packages
                                                                 Dim strFeatureList, strId
                                                                 Dim intMisPackageId,intMlsPackageId 
                                                                                                                                                                                                                                                                 
                                                                 intMisPackageId=232
                                                                 intMlsPackageId=431
                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                 ShowErrorMessage "Acer Swap: Package updates"  
                                                                                                                                                                                                                                                                 strFeatureList=RegistryMgr.GetRegKeyValueName("HKLM\SOFTWARE\McAfee\MSC\Packages\" & CStr(intMlsPackageId) ,"features",False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\"), CStr("activepackages"),  CStr(intMisPackageId),False)
                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\"), CStr("activeuipackages"),  CStr(intMisPackageId),False)
                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\"),  CStr("Packages"),  CStr(intMisPackageId),False)
                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\" & CStr(intMisPackageId)),  CStr("features"),  CStr(strFeatureList),False)
                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                 'Reading AppID informations
                                                                                                                                                                                                                                                                ShowErrorMessage "Acer Swap: AppID updates"  
                                                                                                                                                                                                                                                                 Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppID\mpf"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\mps"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\mqs"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\msad"), CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\msk"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\vso"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\MSC\Packages\AppId\vul"),  CStr("packageid"),  CStr(intMisPackageId),False)
                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                 ShowErrorMessage "Acer Swap: IsSwapped flag update"  
                                                                                                                                                                                                                                                     Call RegistryMgr.SetRegKeyValueCstr( CStr("HKLM\SOFTWARE\McAfee\PartnerCore"),  CStr("IsSwapped"),  CStr("1"),False)
                                                                                                                                                                                                                                                                
                                                            End If
                                                                                                                                                                                                                                                
                  
                         Else
                            RegistryStatus = SM_SYSMGR_OBJECT_CREATION_FAILURE 
                            ShowErrorMessage "Acer Swap: registry object creation failed"  

                    End If

            End If

            ShowErrorMessage "Acer: Swap Affid call end: Success"

End Function
  
 '===========================================================================================
'                                          Acer SKU swap Fix End
'===========================================================================================

' HTML Encodes a string
Function HTML_Encode(byVal strVal)
    If ((TypeName(strVal)="String") And (Not IsNull(strVal)) And (strVal<>"")) Then
      Dim tmp
      tmp = strVal
      tmp = Replace(tmp, chr(34), "&quot;")
      tmp = Replace(tmp, chr(39), "&apos;")
      tmp = Replace(tmp, chr(60), "&lt;")
      tmp = Replace(tmp, chr(62), "&gt;")
      tmp = Replace(tmp, chr(38), "&amp;")
    End If
      HTML_Encode = tmp
End Function

'Get the CSP Client Id
Function GetCSPClientId()
    Dim cl 
	On Error Resume Next
	set cl = CreateObject("McCSPClient")			

	If Err.Number <> 0 Then GetCSPClientId = "" : Exit Function
    
	Dim args: args = "{""application_id"": ""a053060c-3a34-11e4-8a01-005056b7244f"",""op"": ""get""}"
	Dim result: result = cl.GetAppInfo(args)
	Dim strDevice, strDeviceRemaining, arrDevice
	strDeviceRemaining = mid(result, instr(result,"deviceid"))
								
	strDevice = replace(mid(trim(strDeviceRemaining), 1, instr(trim(strDeviceRemaining),",")-1),"""", "")
								
	arrDevice = (split(strDevice, ":",-1))
	GetCSPClientId = arrDevice(1)
	    
End Function



Class CSMMSCHelper
    
    Private m_External, m_MscHelper
    
    Private Sub Class_Initialize
        Set m_External = window.external
        Set m_MscHelper = Nothing
    End Sub

    Public Function Initialize(isIptRequired)
        On Error Resume Next
        Dim intResult       
        intResult = 0        
        If m_MscHelper Is Nothing Then
            Set m_MscHelper = window.external.GetParam(PLATFORMDSP).GetProperty(SM_PROP_PROVIDER_MSCIPT,false)
        Else
            intResult = SM_OBJECT_CREATION_FAILED
            ShowErrorMessage("MscIPT:Unable to get the Provider object")
        End If

        If isIptRequired Then
            ' This code appears to have no upstream callers.
            ' If it is required, supply "True" to Initialize function for IPT.
            If Not IsObject(m_ObjMscIpt) Then
                Initialize = SM_OBJECT_CREATION_FAILED
                Exit Function
            End If

            If m_ObjMscIpt Is Nothing Then
                Initialize = SM_OBJECT_CREATION_FAILED
                Exit Function
            End If
        End If

        Initialize = SM_OBJECT_CREATION_SUCCESS
    End Function

    Public Function IsAutomaticUpdatesEnabled()
        Dim isUpdatesEnabled
        isUpdatesEnabled = false
        If Not (m_MscHelper is Nothing) Then
            isUpdatesEnabled = m_MscHelper.IsAutomaticUpdatesEnabled()
        End If
        IsAutomaticUpdatesEnabled = isUpdatesEnabled
    End Function

    Public Function DisplayMetroToast(strTitle, strDescription, strImagePath, intMetroTemplateType, objCallback)
        Dim intResult : intResult = -1
        If Not m_MscHelper Is Nothing Then
            Call m_MscHelper.DisplayMetroToast(strTitle, strDescription, strImagePath, intMetroTemplateType, objCallback)
        End If

        DisplayMetroToast = intResult
    End Function

    Public Function IsUserInMetroMode()
        Dim bResult : bResult = False
        Dim intMode : intMode = -1

        If Not m_MscHelper Is Nothing Then
            intMode = m_MscHelper.GetCurrentDisplayMode()
            If (intMode = SM_SUBMGR_METRODISPLAYMODE_METRO Or intMode = SM_SUBMGR_METRODISPLAYMODE_STARTUP) Then
                bResult = True
            End If
        End If

        IsUserInMetroMode = bResult
    End Function
        
    Private Sub Class_Terminate      
        Call CleanUp(m_External)
        Call CleanUp(m_MscHelper)        
    End Sub   

    Private Sub CleanUp(obj)
        If Not IsEmpty(obj) Then
            If Not obj Is Nothing Then
                Set obj = Nothing
            End If
        End If
    End Sub
        
End Class

Class CSMMetroHelper
    Private m_MetroDismissReason

    Private Sub Class_Initialize
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Function NotifyMetroToastAction(vStatus)
        ' This is asynchronous, so if we want to know how the toast was dismissed,
        ' we may have to set the result in a hidden DOM element for the Classic toast
        ' to access.
        m_MetroDismissReason = vStatus
    End Function
End Class


Class CSMMscIpt

	Private m_ObjMscIpt,actionUrl

	Private Sub Class_Initialize()
		Set m_ObjMscIpt = Nothing
	End Sub

	Private Sub Class_Terminate()
		If Not (m_ObjMscIpt is Nothing) Then
			Set m_ObjMscIpt = Nothing
		End If
	End Sub
	
	Public Function Initialize(objProvider)
      
        On Error Resume Next
        Dim intResult       
        intResult = 0
        If Not(objProvider Is Nothing) Then   
            If m_ObjMscIpt Is Nothing Then
                Set m_ObjMscIpt = window.external.GetParam(PLATFORMDSP).GetProperty(SM_PROP_PROVIDER_MSCIPT,false)
            End If
        Else
            intResult = SM_OBJECT_CREATION_FAILED
            ShowErrorMessage("MscIPT:Unable to get the Provider object")
        End If

        If Not IsObject(m_ObjMscIpt) Then
            Initialize = SM_OBJECT_CREATION_FAILED
            Exit Function
        End If

        If m_ObjMscIpt Is Nothing Then
            Initialize = SM_OBJECT_CREATION_FAILED
            Exit Function
        End If

        Initialize = SM_OBJECT_CREATION_SUCCESS

    End Function

    Public Function IsIPTEnabled()
 
         If Not IsObject(m_ObjMscIpt) Then
            IsIPTEnabled = False
            Exit Function
        End If

        If m_ObjMscIpt Is Nothing Then
            IsIPTEnabled = False
            Exit Function
        End If
                  
        IsIPTEnabled = m_ObjMscIpt.IsIPTEnabled()
                       
    End Function

    Public Sub LaunchInIPT(actionUrl, dwWebOverlayOption)
        Dim result
        If Not IsObject(m_ObjMscIpt) Then
            ShowErrorMessage("MscIPT:Unable to get the Provider object")            
            Exit Sub
        End If

        If m_ObjMscIpt Is Nothing Then
            ShowErrorMessage("MscIPT:Unable to get the Provider object")            
            Exit Sub
        End If

        Call m_ObjMscIpt.OpenUrl(actionUrl, 1, dwWebOverlayOption, "")

    End Sub

    Public Sub TriggerAppWSCall()

        Dim result
        If Not IsObject(m_ObjMscIpt) Then
            ShowErrorMessage("MscIPT:Unable to get the Provider object")            
            Exit Sub
        End If

        If m_ObjMscIpt Is Nothing Then
            ShowErrorMessage("MscIPT:Unable to get the Provider object")            
            Exit Sub
        End If

        Call m_ObjMscIpt.UpdateWebCustomizations()

    End Sub

End Class

