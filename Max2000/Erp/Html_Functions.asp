<OBJECT RUNAT=server PROGID=SubFunctions.Main id=FHTML></OBJECT>
<object id="Convert" PROGID="ConvertMisToWord.ChangMis" RUNAT="server"></object>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RsComp> </OBJECT>
<OBJECT RUNAT=server PROGID=NtvDB.Do id=RssDo></OBJECT>

<!-- #include file="../HebrewMeta.jv"-->


<%
    'ffffffffffffffff -- _ascii.exe
    ' //???? ????? ??????????? ??????????????- ??????????? ????????????? ?????????  
'================================
Dim CNm 
Dim CKod 
Dim CCity 
Dim CStreet
Dim CStreetNo
Dim CTel
Dim CTel2
Dim CFax
Dim COsekType
Dim LblTel
Dim LblTel2
Dim LblFax
Dim CNoMaam 
Dim COsek
Dim LblOsek 
Dim LblEmail
Dim CEmail
Dim LblPost
Dim PostBCompany
Dim CWebAddress
Dim SwCompanyFile
Dim SwCompanyFileSnif
Dim CompanyFileSnif
Dim SwComsign_Lk
Dim Idx_Invoce_Email
Dim SwLogoMail
Dim SwLogoPicMail
dim SwPrintLogoMail
dim msg200
dim msg201
dim msg202
dim msg203
dim msg204
dim msg205
dim msg206
dim msg207
dim msg208
dim SwUsemsg

SwKoteretSpk="0"
wrkC=Trim(Request("C"))

SwPrintToStore=Request("SwPrintToStore")
SwPrintFromPrintServer=Request("SwPrintFromPrintServer")
if (Request("SwIEWin9") ="1" or Request("SwPdf") ="1"  or SwPdf ="1" ) and IsNumeric(Trim(MaxLinePage)) then MaxLinePage=MaxLinePage-3

SentToEmail_Add=Replace(FHTML.getResponse_NoGeresh(Request("SentToEmail_Add")),"""","&quot;")

Private Function CompanyFile()	
    CNm = ""
    CKod = ""
    CCity = ""
    CStreet = ""
    CStreetNo = ""
    CTel = ""
    CTel2 = ""
    CReportHeader = ""
    CompanyFileSnif=Trim(Request("CurrSnif"))
    if not IsNumeric(Trim(CompanyFileSnif)) then CompanyFileSnif="0"
	SwCompanyFile=True
	

	
	set PG=Createobject("PG_SqlPostgres.Main")'<--
	
	if IsNumeric(Trim(DocType)) then 
		sql=" select   isnull(SwPrintKotFromSpk,0) as SwKoteretSpk " & _ 
			" from PrmDoc (nolock) " & _
		    " where Company = " & CurrCompany & _
		    " and Y= " & CurrYear & _
		    " and Type =  " & DocType
		sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
		RsComp.Open sql, Conn, 1, 1
		if not RsComp.Eof then 
			SwKoteretSpk=RsComp("SwKoteretSpk")
		end if
		RsComp.Close
	end if

	ssTbl=""
	select case DocType
	case "612","616", "216","212","416": ssTbl="MlayHzm"
	case "670", "671", "650", "651", "652": ssTbl="MlayDoc"
	end select
	
		SqlCompany = " select " & _
			" isnull(Company.Kod,0) as Kod , " & _
			" rtrim(Company.Nm)     as Nm , " & _
			" rtrim(ltrim(isnull(Company.eNm,'')))     as eNm , " & _
			" rtrim(isnull(MCity.Nm,''))   as City ," & _
			" rtrim(ltrim(isnull(MCity.eNm,'')))   as eCity ," & _
			" rtrim(isnull(Company.Street,'')) as Street , " & _
			" rtrim(ltrim(isnull(Company.eStreet,''))) as eStreet , " & _
			" isnull(convert(char ,Company.StreetNo),'')      as StreetNo , " & _
			" rtrim(isnull(Company.Tel,''))    as Tel, " & _
			" rtrim(isnull(Company.Tel2,'')) as Tel2 , " & _
			" rtrim(isnull(Company.Fax,'')) as Fax , " & _
			" isnull(Company.Osek,0) as Osek ,isnull(Company.OsekType,0) as OsekType,  isnull(Company.TikMaamMeochad,0) as TikMaamMeochad , " & _
		    " isnull(isnull(Snif.TikMaam_Snif,Company.TikMaam),0) as TikMaam, " & _
			" rtrim(isnull(Company.Email,'')) as Email ," & _
			" rtrim(isnull(Company.WebAddress,'')) as WebAddress ," & _
			" rtrim(ltrim(isnull(Company.ReportHeader,''))) as ReportHeader, " & _
			"  isnull(convert(char,Company.Mikod),'') as Mikod," & _
			" isnull(Company.CompanyPostBox,0) as PostBCompany ," & _
			" isnull(Company.NoMaam,0) as NoMaam " & _
			" from Company " & _
			" left join Max2000_Lib..City as MCity On Company.City = MCity.C " & _
		    " left join Snif on Snif.Company  = Company.C and  Snif.C = " & CompanyFileSnif & _	
			" where Company.C= " & CurrCompany   


    if SwKoteretSpk="2" and IsNumeric(Trim(CompanyFileSnif))  and CompanyFileSnif<>"0" then 
 		sql = " select " & _
			" isnull(Snif.Kod,0) as Kod , " & _
			" ltrim(rtrim(isnull(Company.Nm,''))) + ' - ' + ltrim(rtrim(isnull(Snif.Nm collate Hebrew_CI_AS,'')))    as Nm , " & _
			" rtrim(ltrim(isnull(Snif.Nm_Eng,'')))     as eNm , " & _
			" rtrim(isnull(MCity.Nm,''))   as City ," & _
			" rtrim(ltrim(isnull(MCity.eNm,'')))   as eCity ," & _
			" rtrim(isnull(Snif.Street,'')) as Street , " & _
			" isnull(convert(char ,Snif.StreetNo),'')      as StreetNo ,'' as  eStreet," & _
			" rtrim(isnull(Snif.Tel,''))    as Tel, " & _
			" '' as Tel2 , " & _
			" rtrim(isnull(Snif.Fax,'')) as Fax , " & _
			" isnull(Company.Osek,0) as Osek ,isnull(Company.OsekType,0) as OsekType,  isnull(Company.TikMaamMeochad,0) as TikMaamMeochad , " & _
		    " isnull(isnull(Snif.TikMaam_Snif,Company.TikMaam),0) as TikMaam, " & _
			" rtrim(isnull(Company.Email,'')) as Email ," & _
			" rtrim(isnull(Company.WebAddress,'')) as WebAddress ," & _
			" rtrim(ltrim(isnull(Company.ReportHeader,''))) as ReportHeader, " & _
			"  isnull(convert(char,Snif.Mikod),'') as Mikod," & _
			" 0 as PostBCompany ," & _
			" isnull(Snif.NoMaam_Snif,0) as NoMaam " & _
			" from Snif " & _
			" left join Max2000_Lib..City as MCity On Snif.City = MCity.C " & _
			" left join Company on Company.C=Snif.Company " & _
			" where Snif.C= " & CompanyFileSnif  
			SwCompanyFileSnif=true
   elseif SwKoteretSpk="1" and IsNumeric(Trim(Doc)) and ssTbl<>"" then
		sql=" select top 1 isnull(Idx.Kod,0) as Kod,rtrim(Idx.Nm) as Nm, rtrim(Idx.eNm) as eNm," & _
			" rtrim(isnull(MCity.Nm,''))   as City ," & _
			" rtrim(isnull(MCity.eNm,''))   as eCity ," & _
			" rtrim(isnull(Idx.Street,'')) as Street, " & _
			" isnull(convert(char ,Idx.Street_No),'')      as StreetNo  ,'' as eStreet, " & _
			" rtrim(isnull(Idx.Tel,'')) as Tel, " & _
			" '' as Tel2 , " & _
			" rtrim(isnull(Idx.Fax,'')) as Fax, " & _
			" rtrim(isnull(Idx.Email,'')) as Email ," & _
			" isnull(Idx.Mikod,'') as Mikod ," & _
			" isnull(Idx.Osek,0) as Osek ,isnull(Idx.OsekType,0) as OsekType, " & _
			" 0 as TikMaamMeochad, 0 as TikMaam, '' as ReportHeader," & _
			" rtrim(isnull(Idx.WebAddress,'')) as WebAddress,0 as NoMaam, " & _
			" rtrim(isnull(Idx.PostBox,'')) as PostBCompany " & _
			" from Idx " & _
			" inner join Prt on Prt.Spk=Idx.C " & _
			" left join " & ssTbl & "_Lines (nolock) as L ON L.Prt = Prt.C " & _
			" left join " & ssTbl & " (nolock) as D on D.C = L." & ssTbl  & _
			" left join Max2000_Lib..City as MCity On Idx.City = MCity.C " & _
	        " where D.Company = " & CurrCompany & _
			" and   D.Y= " & CurrYear & _
			" and   D.DocType= " & DocType  & _
			" and   D.Doc = " & Doc & _
			" order by L.Line "	
	else
		sql=SqlCompany
	end if
	sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
	RsComp.Open sql, Conn, 1, 1
	
	if RsComp.EOF then 
		SwCompanyFileSnif=false
		sql=SqlCompany
		RsComp.Close
		RsComp.Open sql, Conn, 1, 1
	end if
	
    If RsComp.RecordCount > 0 Then
		If ((RsComp("PostBCompany")) <> 0) Then
		    LblPost = "??.??."
			PostBCompany = RTrim(RsComp("PostBCompany"))
		End If

        CNm = RsComp("Nm")
        CStreet = Trim(RsComp("Street"))
        CKod = RsComp("Kod")
        CStreetNo = Trim(RsComp("StreetNo"))
        COsek = RsComp("Osek")
		if trim(COsek)<>"" then
			n=9-len(trim(COsek))
			if n<0 then n=0
			COsek=String(n,"0")+trim(COsek)
		end if
       COsekType = RsComp("OsekType")
        
        Select Case RsComp("OsekType")
        Case 1:
             
              if SwUsemsg = 1 then 
                LblOsek = msg205  
              else
                LblOsek = "??.??."
                 
              end if 
        Case 2:
               if SwUsemsg = 1 then 
                 LblOsek =  msg204  
               else
                LblOsek = "??.???."
                 
               end if
        Case 3: 
              if SwUsemsg = 1 then
                LblOsek = msg206 & ":"
              else
                LblOsek = "?????????? ???????????"
                
              end if
        case 4:
               if SwUsemsg = 1 then 
                  LblOsek =  msg207 & ":" 
               else
                 LblOsek = "???????''??."
                
               end if
        case 5,7: 
               if SwUsemsg = 1 then 
                 LblOsek = msg208 & ":"
               else
                LblOsek = "?????????? ?????????"
                 
               end if
        Case Else:
                if SwUsemsg = 1 then 
                  LblOsek = msg204 
                else
                 LblOsek = "??.???."
                  
                end if
        End Select
        if not isnull(RsComp("NoMaam")) then CNoMaam = Trim(RsComp("NoMaam"))
        
        CTikMaamMeochad = Trim(RsComp("TikMaamMeochad"))
        CTikMaam = Trim(RsComp("TikMaam"))
        
        CCity = Trim(RsComp("City"))
                
        If Trim(RsComp("Mikod")) <> ""   Then CCity = CCity & " " & Trim(RsComp("Mikod"))
        
        LblTel = ""
        If Len(RsComp("Tel")) > 0 Then
            
 
            if SwUsemsg = 1 then 
          
               LblTel = msg200 & ":"
            else
              LblTel = "?????????????:"
              
            end if    
            CTel = RTrim(RsComp("Tel"))
        End If
        
        LblTel2 = ""
        If Len(RsComp("Tel2")) > 0 Then
            if SwUsemsg = 1 then 
			  LblTel2 = msg201 & ":"
			else
			   LblTel2 = "????????????? ??????????:" 
			end if 
			CTel2 = RTrim(RsComp("Tel2")) 
        End If
        
        LblFax = ""
        If Len(RsComp("Fax")) > 0 Then
               if SwUsemsg = 1 then 
                  LblFax = msg202 & ":" 
              else
                 LblFax = "????????:"
              end  if 
             CFax = RTrim(RsComp("Fax"))    
        End If
        
        If RTrim(RsComp("Email")) = "" Then
			LblEmail = ""
			CEmail = ""
        Else
            if SwUsemsg = 1 then
               LblEmail = msg203 & ":"
            else
              LblEmail = "???????? ???????????????????:"  
            end if 
            CEmail = RTrim(RsComp("Email"))
        End If
                
        If RTrim(RsComp("WebAddress")) = "" Then
            CWebAddress = ""
        Else
            CWebAddress = RTrim(RsComp("WebAddress"))
        End If
        CReportHeader = Trim(RsComp("ReportHeader"))
    End If
    RsComp.Close
	if  CCity="0" then CCity=""
	if CNm="" then CNm="&nbsp;"
	set PG=nothing
End Function

Private Function CompanyFile_E()
    CNm = ""
    CKod = ""
    CCity = ""
    CCountry= ""
    CStreet = ""
    CStreetNo = ""
    CTel = ""
    CTel2 = ""
    CReportHeader = ""
    CompanyFileSnif=Trim(Request("CurrSnif"))
    if not IsNumeric(Trim(CompanyFileSnif)) then CompanyFileSnif="0"
    
    set PG=Createobject("PG_SqlPostgres.Main")'<--
    
    Sql = " select " & _
		" isnull(Company.Kod,0) as Kod , " & _
		" rtrim(Company.Nm)     as Nm , " & _
		" rtrim(ltrim(isnull(Company.eNm,'')))     as eNm , " & _
		" rtrim(isnull(MCity.Nm,''))   as City ," & _
		" rtrim(ltrim(isnull(MCity.eNm,'')))   as eCity ," & _
		" rtrim(isnull(Company.Street,'')) as Street , " & _
		" rtrim(ltrim(isnull(Company.eStreet,''))) as eStreet , " & _
        " rtrim(ltrim(isnull(MCountry.eNm,'')))   as eCountry ," & _
		" isnull(convert(char ,Company.StreetNo),'')      as StreetNo , " & _
		" isnull(State_City.Nm,'')      as State_City , " & _
		" rtrim(isnull(isnull(Company.eTel,Company.Tel),''))    as Tel, " & _
		" rtrim(isnull(Company.Tel2,'')) as Tel2 , " & _
		" rtrim(isnull(isnull(Company.eFax,Company.Fax),'')) as Fax , " & _
		" isnull(Company.Osek,0) as Osek ,isnull(Company.OsekType,0) as OsekType,  isnull(Company.TikMaamMeochad,0) as TikMaamMeochad , " & _
        " isnull(isnull(Snif.TikMaam_Snif,Company.TikMaam),0) as TikMaam, " & _
		" rtrim(isnull(Company.Email,'')) as Email ," & _
		" rtrim(isnull(Company.WebAddress,'')) as WebAddress ," & _
		" rtrim(ltrim(isnull(Company.ReportHeader,''))) as ReportHeader, " & _
		"  isnull(convert(char,Company.Mikod),'') as Mikod," & _
		" isnull(Company.CompanyPostBox,0) as PostBCompany, " & _
		" isnull(Company.NoMaam,0) as NoMaam " & _
		" from Company " & _
		" left join Max2000_Lib..City as MCity On Company.City = MCity.C " & _
        " left join Max2000_Lib..Country as MCountry On Company.Country = MCountry.C " & _
		" left join Max2000_Const..State_City State_City on State_City.C=Company.State " & _
		" left join Snif on Snif.Company  = Company.C and  Snif.C = " &  CompanyFileSnif & _
		" where Company.C= " & CurrCompany
	sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
    Rs.Open Sql, Conn, 1, 1
    If Rs.RecordCount > 0 Then
		If ((Rs("PostBCompany")) <> 0) Then
		    LblPost = "P.O.B:"
			PostBCompany = RTrim(Rs("PostBCompany"))
		End If

        CNm = Rs("eNm")
        CStreet = Trim(Rs("eStreet"))
        CKod = Rs("Kod")
        CStreetNo = Trim(Rs("StreetNo"))
		CCountry = Trim(Rs("eCountry"))
		State_City = Trim(Rs("State_City"))

        COsek = Rs("Osek")
		if trim(COsek)<>"" then
			n=9-len(trim(COsek))
			if n<0 then n=0
			COsek=String(n,"0")+trim(COsek)
		end if
		COsekType = Rs("OsekType")
		if not isnull(Rs("NoMaam")) then CNoMaam = Trim(Rs("NoMaam"))
		LblOsek="VAT No."
        CTikMaamMeochad = Trim(Rs("TikMaamMeochad"))
        CTikMaam = Trim(Rs("TikMaam"))
        CCity = Trim(Rs("eCity"))
                
        If Rs("Mikod") <> "" Then CCity = CCity & " " & Trim(Rs("Mikod"))
		if  CCity="0" then CCity=""
		if  CCountry="0" then CCountry=""
        if Trim(CCountry)<>"" then 
			CCity=CCity & " <BR>" &  CCountry 
			if State_City<>"" then CCity= CCity & ",&nbsp;" & State_City
		end if
        
        LblTel = ""
        If Len(Rs("Tel")) > 0 Then
            LblTel = "Tel:"
            CTel = RTrim(Rs("Tel"))
        End If
        
        LblFax = ""
        If Len(Rs("Fax")) > 0 Then
            LblFax = "Fax:"
            CFax = RTrim(Rs("Fax"))
        End If
        
        If RTrim(Rs("Email")) = "" Then
			LblEmail = ""
			CEmail = ""
        Else
            LblEmail = "E-Mail:"
            CEmail = RTrim(Rs("Email"))
        End If
                
        If RTrim(Rs("WebAddress")) = "" Then
            CWebAddress = ""
        Else
            CWebAddress = RTrim(Rs("WebAddress"))
        End If
        CReportHeader = Trim(Rs("ReportHeader"))
    End If
    Rs.Close
    set PG=nothing
End Function
'===================
 Function GetSwMursheUsr()
	SwLoMhrR=0
    sql=" SELECT  isnull(SwLoMhrR,0) as SwLoMhrR  FROM Users WHERE  C=" + cstr(UserCounter)			
	RS.Open sql,Conn
    If not Rs.Eof Then SwLoMhrR = Rs("SwLoMhrR")
    Rs.Close
    if SwLoMhrR<>0 and DocType>"250" and  DocType<"280" and DocType<>"250" and DocType<>"251" and PrintTofes<>"25" then 
		NoPrintMhr_Lines=True
		NoPrintMhr_Total=True
	end if
End Function  
Private Function GetPrmCompany()
    Sql = " select " & _
		  " isnull(SwGivunPrt,0) as SwGivunPrt," & _
		" isnull(SwWorkWithStoreLogisti,0) as SwWorkWithStoreLogisti  " & _
		" from Prm_Company " & _
		" where Prm_Company.Company= " & CurrCompany   
	Rs.Open Sql, Conn, 1, 1
    If Rs.RecordCount > 0 Then
        SwWorkWithStoreLogisti = (Rs("SwWorkWithStoreLogisti") = 1)
        SwGivunPrt = Rs("SwGivunPrt")
	end if
	Rs.Close
end function
Private Function GetPrmDoc(SwSnif)
	dim smMTbl
	SwIdxMeshek=0
	if not IsNumeric(Trim(DocType)) then DocType=0
	NoPrintMhr_Meshek = False
    SwNoLogo = False
    NoPrintStore = False
    SwPrintNoCopyNo = False
    SwNotPrintTopLine = False
    SwNotPrintTikun = False
    SwNoDisDoc = False
    SwPrintItra = False
    SwNotPrintSignLk = False
    SwLogoMail=false
    SwLogoPicMail=0
    SwPrintLogoPic = 0
    PrintJumpLinesBut=0
    SwCmtAriza=0
    SwMiun=0
    
    PrintTofes = 0
    SwNotPrintButLine = False
    SwPrintSnif = False
    sCurrYear=Y
	if not isnumeric(trim(Y)) then sCurrYear=CurrYear
	
	set PG=Createobject("PG_SqlPostgres.Main")'<--

		MDSnif="0"
		smMTbl=""
		SwMDSnif="0"
		select case DocType
		case "612","616", "216","212","416": smMTbl="MlayHzm" : SwMDSnif=1
		case "650", "651", "670", "671", "652", "470": smMTbl="MlayDoc": SwMDSnif=1
		case "628":  smMTbl="Inv"
		case "630":  smMTbl="InvKabala"
		case "680":  smMTbl="Kabala"
		end select
		if IsNumeric(Trim(Doc)) then
			if smMTbl<>"" then
				strSnif="0 as Snif"
				strDocType=""
				strSnifLeft=""
				if SwMDSnif="1" then
					strSnif="St.Snif as Snif"
					strSnifLeft=" left join Store St on St.C=D.Store "
					strDocType=" and D.DocType=  " & DocType
				end if
				sql =	" select " & strSnif & ",D.Idx as IdxC,isnull(Idx.SwIdxMeshek,0) as SwIdxMeshek, " & _
						" Idx_Email.Email as Idx_Invoce_Email " & _
						" FROM " & smMTbl & " (nolock) D " & _
						strSnifLeft & _
						" left join Idx on Idx.C=D.Idx " & _
						" left join Idx_Invoce_Email as Idx_Email on Idx_Email.Idx=Idx.C " & _
						" where D.Company = " & CStr(CurrCompany) & _
						" and D.Y= " & CStr(sCurrYear) & _
						" and D.Doc = " & CStr(Doc) & strDocType					
				sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
				Rs.Open sql, Conn, 1, 1
				If Rs.RecordCount > 0 Then 
					MDSnif = Rs("Snif") :IdxC=Rs("IdxC"): SwIdxMeshek=Rs("SwIdxMeshek")
					Idx_Invoce_Email=Rs("Idx_Invoce_Email")
					if not IsNumeric(Trim(MDSnif)) then MDSnif="0"
				end if
				Rs.Close
			end if
		end if
	sql = " select " & _
		" isnull(PrmDoc.SwPrintSnif,0) as SwPrintSnif ,isnull(PrmDoc.eSign,Sign)as eSign, isnull(PrmDoc.Sign,'')as Sign,  " & _
		" isnull(PrmDoc.SwPrintLogo,0) as SwPrintLogo , " & _
		" isnull(PrmDoc.SwLogoMail,0) as SwLogoMail , " & _
		" isnull(PrmDoc.SwLogoPicMail,0) as SwLogoPicMail , " & _
		" isnull(PrmDoc.SwPrintLogoMail ,0) as SwPrintLogoMail  , " & _
		" isnull(PrmDoc.SwEnglish,0) as SwEnglish, " & _
		" isnull(PrmDoc.PrintRemarks,'') as PrintRemarks ,  " & _
		" isnull(PrmDoc.PrintRemarks_Eng,'')as PrintRemarks_Eng, " & _
	    " isnull(PrmDoc.PrintRemarks1,'') as PrintRemarks1 ,  " & _
		" isnull(PrmDoc.PrintJumpLinesTop,0)as PrintJumpLinesTop, " & _
		" isnull(PrmDoc.PrintJumpLinesAfterTo,0)as PrintJumpLinesAfterTo, " & _
		" isnull(PrmDoc.PrintJumpLinesBut,0)as PrintJumpLinesBut, " & _
		" isnull(PrmDoc.NoPrintStore,0)as NoPrintStore, isnull(PrmDoc.SwNotPrintButLine,0) as SwNotPrintButLine , " & _
		" isnull(PrmDoc.SwPrintNoCopyNo,0)as SwPrintNoCopyNo, " & _
		" isnull(PrmDoc.SwNotPrintNumLk,0) as SwNotPrintNumLk ," & _
		" isnull(PrmDoc.NoPrintDatePeraon,0) as NoPrintDatePeraon,isnull(PrmDoc.NoPrintDateIt,0) as NoPrintDateIt, " & _
		" isnull(PrmDoc.NoPrintCar,0) as NoPrintCar, isnull(PrmDoc.SwCmtAriza,0) as SwCmtAriza, isnull(PrmDoc.NoPrintDriver,0) as NoPrintDriver, " & _
		" isnull(PrmDoc.PrintTofes,0) as PrintTofes  , " & _
		" isnull(PrmDoc.SwBarKod,0) as SwBarKod  , " & _
		" isnull(PrmDoc.NoPrintMhr,0)as NoPrintMhr, " & _
		" isnull(PrmDoc.SwMurshe,0)as SwMurshe, " & _
		" isnull(PrmDoc.SwNotPtintPhone,0) as SwNotPtintPhone ," & _
		" isnull(PrmDoc.NoDbr,0) as SwMishkalNeto, " & _
		" isnull(PrmDoc.SwPrintLogoPic,0) as SwPrintLogoPic, " & _
		" isnull(PrmDoc.SwNotPrintTopLine,0) as SwNotPrintTopLine, " & _
		" isnull(PrmDoc.SwPrintItra,0) as SwPrintItra ," & _
		" isnull(PrmDoc.SwPrtNm,0)as SwPrtNm, " & _
		" isnull(PrmDoc.SwNumN,2) as SwNumN, " & _
		" isnull(PrmDoc.SwContactMan,0) as SwContactMan ," & _
		" isnull(PrmDoc.SwNotPrintHzm,0) as SwNotPrintHzm, " & _
		" isnull(PrmDoc.SwNotPrintSign,0) as SwNotPrintSign ," & _
		" isnull(PrmDoc.SwIdxKupa,0) as SwProj, " & _        
		" isnull(PrmDoc.SwNotPrintTikun,0) as SwNotPrintTikun  , " & _
		" isnull(PrmDoc.SwNotPrintSignLk,0) as SwNotPrintSignLk  , " & _
		" isnull(PrmDoc.SwPrintIdxRemark,0) as SwPrintIdxRemark  , " & _
		" isnull(PrmDoc.SwBarKodDoc,0) as SwBarKodDoc , " & _
		" isnull(PrmDoc.SwPrtFullNm,0) as SwPrtFullNm, " & _
		" isnull(PrmDoc.SwPrintKotFromSpk,0) as SwPrintKotFromSpk, " & _
		" isnull(PrmDoc.SwRefA,0) as SwRefA, " & _
		" isnull(PrmDoc.SwPrm1,0)as SwPrm1, " & _
        " isnull(PrmDoc.SwSeifHiuv,0)as PrtGivun, " & _
		" isnull(PrmDoc.SwPrtSerial,0) as SwPrtSerial, " & _
		" isnull(PrmDoc.SwPrintNmEska,0) as SwPrintNmEska ," & _
		" isnull(PrmDoc.SwNoPrintSochen,0) as SwNoPrintSochen ," & _
		" isnull(PrmDoc.SwNotPrintMt,0) as SwNotPrintMt ," & _
		" isnull(PrmDoc.SwHzmOpen,0)as SwHzmOpen , " & _
		" isnull(PrmDoc.SwNotPrintMaaraz,0)as SwNotPrintMaaraz , " & _
        " isnull(PrmDoc.SwNotMustMakor,0) as SwNotMustMakor, " & _
        " isnull(PrmDoc.DocKot_Heb,'') as DocKot_Heb, " & _
        " isnull(PrmDoc.DocKot_Eng,'') as DocKot_Eng, " & _
        " isnull(PrmDoc.CmtIsh,0) as CmtIsh, " & _
        " isnull(PrmDoc.SwNotPrintEng_Nm,0) as SwNotPrintEng_Nm, " & _
        " isnull(PrmDoc.SwNotPrintRemark,0) as SwNotPrintRemark, " & _ 
        " isnull(PrmDoc.SwMustRemInLine,0) as SwMustRemInLine, " & _ 
        " isnull(PrmDoc.PrintCopies,0) as PrintCopies, " & _
        " isnull(PrmDoc.SwMailFax,0) as SwMailFax, " & _ 
        " isnull(PrmDoc.SwMiun,0) as SwMiun, " & _ 
        " isnull(PrmDoc.SwMiunShow,0) as SwMiunShow, " & _ 
        " isnull(PrmDoc.SwNotPrintBarkode,0) as SwNotPrintBarkode, " & _ 
        " isnull(PrmDoc.SwNoDisDoc,0) as SwNoDisDoc, " & _ 
        " isnull(PrmDoc.SwNotPrintReLMesofon,0) as SwNotPrintReLMesofon, " & _ 
        " isnull(PrmDoc.SwNotPrintNIS,0) as SwNotPrintNIS, " & _ 
        " isnull(PrmDoc.SwPrintMivza,0) as SwPrintMivza, " & _ 
        " isnull(PrmDoc.SwNotPrintDis,0) as SwNotPrintDis, " & _ 
        " isnull(M.C,0) as Meholel_TfasimC, " & _ 
		" isnull(PrmDoc.SwShowLastLine,0) as SwShowLastLine, " & _
		" isnull(PrmDoc.SwDifdufStatus,0) as SwDifdufStatus, " & _
		" isnull(PrmDoc.SwNotPrintTz,0) as SwNotPrintTz, " & _ 
		" isnull(PrmDoc.NotPrintLineWithoutMhr,0) as NotPrintLineWithoutMhr, " & _ 
		" isnull(PrmDoc_Nosaf.SwNotPrintKabBeginDocs,0) as SwNotPrintKabBeginDocs, " & _ 
		" isnull(PrmDoc_Nosaf.SwPrintBarKodRef,0) as SwPrintBarKodRef " & _ 
		" from PrmDoc (nolock) " & _
 		" left join Meholel_Tfasim M on M.N_Tofes=PrmDoc.PrintTofesGen and PrmDoc.Type=M.DocType " & _
		" LEFT JOIN PrmDoc_Nosaf  on PrmDoc_Nosaf.Company=" & cstr(CurrCompany) & " and PrmDoc_Nosaf.Y=" &  cstr(CurrYear) & " and PrmDoc.Type=PrmDoc_Nosaf.Type " & _
       " where PrmDoc.Company = " & CurrCompany & _
        " and PrmDoc.Y= " & CurrYear & _
        " and PrmDoc.Type =  " & DocType       
   sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
   if Request("SwDebud")="1" then 
		Response.Write sql
		Response.End 
   end if
   Rs.Open sql, Conn, 1, 1
    If Rs.RecordCount > 0 Then
        SwShowLastLine = Rs("SwShowLastLine")
        SwDifdufStatus = Rs("SwDifdufStatus")
        SwMurshe = Rs("SwMurshe")
		SwMiun = Rs("SwMiun")
		SwMiunShow = Rs("SwMiunShow")
		SwEnglish = Rs("SwEnglish")
       	NoPrintMhr=Rs("NoPrintMhr")
        SwNotMustMakor= Rs("SwNotMustMakor")
		SwNotPrintMaaraz= Rs("SwNotPrintMaaraz")
        PrintRemarks = Rs("PrintRemarks")
        PrintRemarks1 = Rs("PrintRemarks1")
        SwCmtAriza = Rs("SwCmtAriza")
        PrintRemarks_Eng = Rs("PrintRemarks_Eng")
        PrintJumpLinesTop = Rs("PrintJumpLinesTop")
        PrintJumpLinesAfterTo = Rs("PrintJumpLinesAfterTo")
        PrintJumpLinesBut = Rs("PrintJumpLinesBut")
        SwContactMan = Rs("SwContactMan")
        SwPrtFullNm = Rs("SwPrtFullNm")
		SwPrintNmEska=Rs("SwPrintNmEska")
		SwPrintKotFromSpk = Rs("SwPrintKotFromSpk")
		
        if PrintTofesLk<>"1" then
			NoPrintMhr_Lines = False
			NoPrintMhr_Total = False
			NoPrintMhr_Meshek = False
			If Rs("NoPrintMhr") <> 0 Then
			    If Rs("NoPrintMhr") = 2 Then NoPrintMhr_Lines = True
			    If Rs("NoPrintMhr") = 3 Then NoPrintMhr_Total = True
			    If Rs("NoPrintMhr") = 4 Then NoPrintMhr_Meshek = True
			    If Rs("NoPrintMhr") = 1 Then
			        NoPrintMhr_Total = True
			        NoPrintMhr_Lines = True
			    End If
			End If
        end if
        SwNotPrintEng_Nm = (Rs("SwNotPrintEng_Nm") = 1)
        SwNoDisDoc = (Rs("SwNoDisDoc") = 1)
        SwNotPrintTz = (Rs("SwNotPrintTz") = 1)
        NotPrintLineWithoutMhr = (Rs("NotPrintLineWithoutMhr") = 1)
        NoPrintDatePeraon = (Rs("NoPrintDatePeraon") = 1)
        NoPrintDateIt = (Rs("NoPrintDateIt") = 1)
        SwNotPrintRemark = (Rs("SwNotPrintRemark") = 1)
        SwMustRemInLine = (Rs("SwMustRemInLine") = 1)
        SwPrintMivza = (Rs("SwPrintMivza") = 1)
        SwNotPrintReLMesofon = Rs("SwNotPrintReLMesofon") 
        SwNotPrintNIS = Rs("SwNotPrintNIS") 
        NoPrintCar = (Rs("NoPrintCar") = 1)
        NoPrintDriver = (Rs("NoPrintDriver") = 1)
        SwPrintBarKodRef = (Rs("SwPrintBarKodRef") = 1)
        if cstr(swLkPrm)<> "true" then SwBarKod = Rs("SwBarKod")
        If Rs("SwPrintLogo") = 1 Then SwNoLogo = True
        If Rs("SwLogoMail") = 1 Then SwLogoMail = True
        If Rs("NoPrintStore") = 1 Then NoPrintStore = True
        If Rs("SwNotPrintButLine") = 1 Then SwNotPrintButLine = True
        If Rs("SwNotPrintTopLine") = 1 Then SwNotPrintTopLine = True       
        If Rs("SwPrintNoCopyNo") = 1 Then SwPrintNoCopyNo = True
        If Rs("SwNotPrintHzm") = 1 Then SwNotPrintHzm = True
        If Rs("SwPrintSnif")  Then SwPrintSnif = True
        If Rs("SwNotPrintNumLk") = 1 Then SwNotPrintNumLk = True
        If Rs("SwNotPrintDis") = 1 Then SwNotPrintDis = True
        If Rs("SwRefA") = 1 Then SwRefA = True

		SwNotPrintKabBeginDocs = Rs("SwNotPrintKabBeginDocs")
		SwLogoPicMail = Rs("SwLogoPicMail")
		SwPrintLogoMail  = (Rs("SwPrintLogoMail")=1)
		SwPrintLogoPic = Rs("SwPrintLogoPic")
 		SwHzmOpen = Rs("SwHzmOpen")
		PrintTofes = Rs("PrintTofes")
		Meholel_TfasimC = Rs("Meholel_TfasimC")
        SwNotPrintBarkode = (Rs("SwNotPrintBarkode") = 1)
        SwPrintItra = (Rs("SwPrintItra") = 1)
        SwMailFax = (Rs("SwMailFax") = 1)
        SwNumN = Rs("SwNumN")
        SwBarKodDoc = Rs("SwBarKodDoc")
        SwNotPrintSign = Rs("SwNotPrintSign")
        SwPrtNm = Rs("SwPrtNm") 
        SwMishkalNeto=Rs("SwMishkalNeto")
		SwProj= Rs("SwProj")
        SwNotPrintTikun = (Rs("SwNotPrintTikun") = 1)
        SwNotPrintSignLk = (Rs("SwNotPrintSignLk") = 1)
        SwPrintIdxRemark = (Rs("SwPrintIdxRemark") = 1)
        SwNotPtintPhone = Rs("SwNotPtintPhone")
 		SwPrtSerial= Rs("SwPrtSerial")
        SwNoPrintSochen= Rs("SwNoPrintSochen") 
        SwNotPrintMt= Rs("SwNotPrintMt") 
		SwPrm1 = Rs("SwPrm1")
        eSign = Rs("eSign")
        Sign = Rs("Sign")
		DocKot_Heb = Trim(Rs("DocKot_Heb"))
		DocKot_Eng = Trim(Rs("DocKot_Eng"))
        CmtIsh = Rs("CmtIsh")
        PrtGivun = Rs("PrtGivun")
		PrintCopies = Rs("PrintCopies")
    End If
    Rs.Close	

	If isnumeric(Trim(IdxC)) and IdxC<>"0" Then
		
	     sql = " SELECT isnull(isnull(I.Tofes,isnull(IG.Tofes,99)),99) as PrintTofes,I.SwSendAutoMail,Email,isnull(I.Remark,'') Remark,isnull(I.Remark_Eng,'') Remark_Eng, " & _
	             " isnull(I.NoPrintMhr,IG.NoPrintMhr) as NoPrintMhr,isnull(MIG.C,0) as Meholel_TfasimC, " & _
					"isnull(I.SwBarKod,IG.SwBarKod)as SwBarKod,isnull(SwPrintIdxRemark,0) as SwPrintIdxRemark " & _
	           " FROM Idx " & _
	           " left join Idx_PrintDocsPrm I  on Idx.C=I.Idx and I.DocType= " & DocType  & _
	           " left join IdxGrp_PrintDocsPrm IG on IG.DocType = " & DocType &  "  and IG.IdxGrp=Idx.Idx_Grp " & _
 			   " left join Meholel_Tfasim MIG on MIG.N_Tofes=IG.PrintTofesGen and MIG.DocType=IG.DocType " & _
	           " where Idx.C =" + CStr(IdxC) + " and  Not isnull(I.Idx,IG.IdxGrp) is null  "  
	    sql=PG.doSqL_POSTGRES(cstr(Request.ServerVariables("URL")),cstr(sql),cstr(Odbc),cstr(SwSQL))
	    Rs.Open sql, Conn, 1, 1
		swLkPrm = False
		If Not Rs.EOF then
			if  (  Rs("PrintTofes") <> 99  ) Then PrintTofes = Rs("PrintTofes") : Meholel_TfasimC=0
			PrintTofesLk=1
			if  DocType	="216" then	SwPrintIdxRemark=(Rs("SwPrintIdxRemark") = 1)
			NoPrintMhr_Total = False
			NoPrintMhr_Lines = False
			NoPrintMhr_Meshek = False
			If Rs("NoPrintMhr") = 2 Then NoPrintMhr_Lines = True
			If Rs("NoPrintMhr") = 3 Then NoPrintMhr_Total = True
			If Rs("NoPrintMhr") = 4 Then NoPrintMhr_Meshek = True
			If Rs("NoPrintMhr") = 1 Then
				NoPrintMhr_Total = True
				NoPrintMhr_Lines = True
			End If
			if	Trim(Rs("Remark"))<>"" then PrintRemarks=Trim(Rs("Remark"))
			if	Trim(Rs("Remark_Eng"))<>"" then PrintRemarks_Eng=Trim(Rs("Remark_Eng"))
			if Rs("SwBarKod") ="1" then SwBarKod = Rs("SwBarKod") 
			if Rs("Meholel_TfasimC")<>0  then Meholel_TfasimC=Rs("Meholel_TfasimC")
			'if DocType="670" then
				SwSendAutoMail=Rs("SwSendAutoMail")
				'if  SwSendAutoMail=1 and FromFrame="FrameDoc670CloseU" then  SwEmail="": EmailAdd="" : EmailUsr="" : SwPdf=""
			'end if
			swLkPrm = True
	    End If        
	    Rs.Close
	End If
 
	ssSnif=Trim(MDSnif)
	if IsNumeric(Trim(SnifC)) and Trim(SnifC)<>"0" then  ssSnif=Trim(SnifC)
	If  ssSnif<> "0" Then
		sql = " select isnull(NoPrintMhr,0) as NoPrintMhr, isnull(Tofes,99) as PrintTofes,isnull(RemarkSnif,'') RemarkSnif ,SwSendMail, " & _
				" isnull(NotPrintLineWithoutMhr,0) as NotPrintLineWithoutMhr  " & _
				" from PrmDoc_Snif " & _
				" where Snif=" & cstr(ssSnif)  + " and DocType= " & DocType 
	    Rs.Open sql, Conn, 1, 1
		If Not Rs.EOF Then
			if not NotPrintLineWithoutMhr then NotPrintLineWithoutMhr = (Rs("NotPrintLineWithoutMhr") = 1)
			
			if not swLkPrm then
				if  (  Rs("PrintTofes") <> 99  ) Then PrintTofes = Rs("PrintTofes") : Meholel_TfasimC=0
				if	Trim(Rs("RemarkSnif"))<>"" then PrintRemarks=Trim(Rs("RemarkSnif"))
				SwSendAutoMail=Rs("SwSendMail")
			
				if Rs("PrintTofes") <> 99 and Rs("PrintTofes") <> 0 or Rs("NoPrintMhr")<>"0" then
					NoPrintMhr_Total = False
					NoPrintMhr_Lines = False
					NoPrintMhr_Meshek = False
					If Rs("NoPrintMhr") = 2 Then NoPrintMhr_Lines = True
					If Rs("NoPrintMhr") = 3 Then NoPrintMhr_Total = True
					If Rs("NoPrintMhr") = 4 Then NoPrintMhr_Meshek = True
					If Rs("NoPrintMhr") = 1 Then
						NoPrintMhr_Total = True
						NoPrintMhr_Lines = True
					End If
				end if
			end if
			
		End If
		Rs.Close
	End If
	if SwIdxMeshek=0 and NoPrintMhr_Meshek then 
		NoPrintMhr_Total = True
		NoPrintMhr_Lines = True			
	end if
	if IsNumeric(UserCounter) then GetSwMursheUsr()
	if MDSnif<>"0" and MDSnif<>CompanyFileSnif and SwCompanyFile and SwPrintKotFromSpk="2" then 
		CompanyFileSnif=MDSnif
		Call CompanyFile
	end if
	if (DocType="628" or DocType="630" or DocType="650" or DocType="651" or DocType="680") and Idx_Invoce_Email<>""  then '
		SwComsign=GetSwComsign_Lk()
		if SwComsign then SwComsign_Lk=1
	end if
	set PG=nothing
End Function
Function GetSwComsign_Lk()
	SwComsign=false
	LkL=Request("ZorbaLk")
	if Trim(LkL)="" then LkL=Request("LkL")
    sql = " select Count(*) Cn from Max2000_BackOffice..DocPDF_Comsign_Lk " & _
         " where Lk=" & LkL & " and Company=" & CurrCompany
	Cn=RssDo.Main(cstr(sql),cstr("Max2000_BackOffice"),cstr(OdbcUserName),cstr(OdbcPassword),cstr(SwSQL))
	if not IsNumeric(Cn) then Cn="0"
	if Cn>0 then SwComsign=true
    GetSwComsign_Lk = SwComsign
End Function
function FieldLines( byval strS ,Length)
	strS=Replace(strS,"'","")
	'strS=Replace(strS,".","")
	strSp=split(strS,"<br>",-1,1)
	FieldLines=0
	Num_Line=0

	For x=0 to Ubound(strSp) 
		if len(strSp(x))>Length then
			Num_Line=fix(len(strSp(x))/Length)
			if Num_Line*Length<>len(strSp(x)) then Num_Line=Num_Line+1
			FieldLines=FieldLines+Num_Line
		else
			FieldLines=FieldLines+1
		end if
	next
end function
function TranchLines( strS,MaxLine,Length,byref nextDaf  )
	strSp=split(strS,"<br>",-1,1)
	CountLines=0
	Num_Line=0
	nextDaf=""
	x=0
	while (x<=Ubound(strSp) and  CountLines<MaxLine)
		if len(strSp(x))>Length then
			Num_Line=fix(len(strSp(x))/Length)
			if Num_Line*Length<>len(strSp(x)) then	Num_Line=Num_Line+1
			CountLines=CountLines+Num_Line
			if CountLines<MaxLine then 
				TranchLines=TranchLines+strSp(x)+"<br>"
			else
				nextDaf=nextDaf & strSp(x)+"<br>"
			end if
		else
			CountLines=CountLines+1
			TranchLines=TranchLines+strSp(x)+"<br>"
		end if
		x=x+1	
	wend
	while x<=Ubound(strSp) 
		nextDaf=nextDaf & strSp(x) & "<br>"
		x=x+1	
	wend
end function
function TranchLinesOpos( strS,MaxLine,Length  )
	TranchLinesOpos=""
	strS=replace(strS,chr(13),"\n")
	strS=replace(strS,chr(10)," ")
	strS=replace(strS,chr(34),"''")
	strSp=split(strS," ")
	strS=""
	CountLines=0
	x=0
	while (x<=Ubound(strSp) and  CountLines<MaxLine)
		if len(strS)+len(strSp(x))>Length or x=Ubound(strSp)  then 
			if strS="" then strS=strSp(x)
			TranchLinesOpos=TranchLinesOpos & "<R>" & strS &"\n"
			strS=""
			CountLines=CountLines+1
		end if	
		strS=strS+strSp(x) &" "
		x=x+1
		
	wend
end function
function ReplaceAll(str, str1,str2)
	str=trim(str)
	do while InStr(1,str,str1)>0 
		str=Replace(str,str1,str2)
	loop
	ReplaceAll=str
end function
Function Execute_Update(sql)
	on error resume next
	Execute_Update=false	
	Conn.Execute (sql)
	if Err.Number=0 then Execute_Update=true
End Function
Function Execute_Open(Rsss,sql)
	on error resume next
	Execute_Open=false	
	Rsss.Open sql,Conn
	if Err.Number=0 then Execute_Open=true
End Function
Function Loop_Update(sql,Times)
	if not IsNumeric(trim(Times)) or Times="0" then Times=2
	Times=cdbl(Times)
	a=false
	for i=0 to Times
		a=Execute_Update(sql)

		if a then exit for
	next	
End Function
function setResponsePOST(r) 
	r=Trim(r)
	r=Replace(r,"\n\r\n","@@@")
	r=Replace(r,"\r\n\r","@@@")
	r=Replace(r,"\n\r","@@@")
	r=Replace(r,"\r\n","@@@")
	r=Replace(r,"\n","@@@")
	r=Replace(r,"\r","@@@")
	r=Replace(r,"'","$$$")
	r=Replace(r,"""","|")
	setResponsePOST= r
end function
Private Function SetRemark(Str,MaxStr,MaxLines)
	Str=Replace(Str,Chr(13),"<br>")
	Str=Trim(Str)
	ilR= FieldLines(Str,MaxStr)
	if ilR>MaxLines then Str= Replace(Str,"<br>"," ")
	SetRemark=Str
End Function
function ConvertNumber(InNumber,Invert,Mtba)
	ConvertNumber = Convert.Main(CDbl(InNumber), CStr(Invert),cstr("1"), CStr(Mtba))
End function

if SwEmail<>"1" and SwFax<>"1" and SwPdf<>"1"  then%>
<Script language=javascript>
var wshshell2=new ActiveXObject("wscript.network");  
var sDefault="" 

function setPrintStore()
{
	var sPrinterStoreName= "Microsoft XPS Document Writer";
	if("<%=Odbc%>"!="Max2000_1002")
	{
		if( "<%=SwPrintToStore%>"=="1") sPrinterStoreName="Store";
		if( "<%=SwPrintToStore%>"=="2") sPrinterStoreName="Dalpak";
		if( "<%=SwPrintToStore%>"=="3") sPrinterStoreName="Vpsign vpad";
	}
	else
	{
		if( "<%=SwPrintToStore%>"=="2") sPrinterStoreName="Send To OneNote 2007";
		
	}
	repDefaltePrinter(sPrinterStoreName)
	//var sPrinterStoreName= "Send To OneNote 2007";
	//PrintDef=PrintHTML.A4Print_New(sPrinterStoreName);
	//idWBPrint.ExecWB (6,2);
	//PrintHTML.A4Print_New(PrintDef);
	window.close();
}
function funcPrintWindow()
{
	if ("<%=SwPrintFromPrintServer%>" == "1")
	{
	}
	else
	{
		if("<%=SwPrintToStore%>"=="1" || "<%=SwPrintToStore%>"=="2"|| "<%=SwPrintToStore%>"=="3") 	setPrintStore();
		else
		{
			if ("<%=SwPrinter%>"=="1") window.print();
			else idWBPrint.ExecWB (6,2) ; 
		}
	}
}
function repDefaltePrinter2( PrinterName )//???? ????? ??????????? ??????????????- ??????????? ????????????? ?????????  
{ 
	try{   
	//debugger;
		var wshshell=new ActiveXObject("wscript.shell"); 
	    var Printers = wshshell2.EnumPrinterConnections();   
	    
		sRegVal = 'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows\\Device'
	    sDefault = wshshell.RegRead(sRegVal).split(",")[0];
	    SwPrinter=-1;
	    for(i = 0;i<Printers.length;i++)
	    {   
	        if(Printers.Item(i+1)==PrinterName)  SwPrinter=i+1;             
	        i++;   
	    }  
	    if(SwPrinter!=-1)
	    { 
			ps = Printers.Item(SwPrinter);   
			wshshell2.SetDefaultPrinter(ps);   
				    
			idWBPrint.ExecWB(6,2); //print madbeka 
			
			alert("! ????????????? ??????????? ?????????????")
			
			window.setTimeout("wshshell2.SetDefaultPrinter(sDefault);",1000);
			window.close();
		}
	}   
	catch(err)
	{ 	
		wshshell2.SetDefaultPrinter(sDefault);			
	}   
}  

function OpenWin(sUrl)
{
	window.open(sUrl,"","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=1,height=1,left=0px,top=0px");
}
function repDefaltePrinter( PrinterName )
{   
        var printerArray = "";   
        try{   
            var wshshell=new ActiveXObject("wscript.shell");   
            username = wshshell.ExpandEnvironmentStrings("%username%");   
            userTemp = wshshell.ExpandEnvironmentStrings("%TEMP%");   
               
            var Printers = wshshell2.EnumPrinterConnections();   
               
			sRegVal = 'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows\\Device'
			sDefault = ""

		  sDefault = wshshell.RegRead(sRegVal).split(",")[0];
           var cn = wshshell2.ComputerName;   
            SwPrinter=-1;
            for(i = 0;i<Printers.length;i++){   
                 if(Printers.Item(i+1)==PrinterName)  SwPrinter=i+1;             
                i++;   
            }  
            if(i!=-1)
            { 
				ps =  Printers.Item(SwPrinter);   
				wshshell2.SetDefaultPrinter(ps);   
            }
            idWBPrint.ExecWB (6,2) ; 
            window.setTimeout("wshshell2.SetDefaultPrinter(sDefault);",100);
               
        }   
        catch(err){   
           // alert(err);   
        }   
}  
</Script>
<%end if%>
<Script language=vbscript> 
function vbTrim(s)	
	vbTrim=trim(cstr(s)) 
end function 
function vbDbl(nn)
	n=nn
	if not isnumeric(trim(cstr(n))) then n="0"
	vbDbl=cdbl(n)
end function
function vbReplace(s,s1,s2)	
	vbReplace=Replace(s,s1,s2)
end function 
</Script>
<Script language=javascript>
function ftrim(str) {
	if(str==null) str="";
    str = str.toString();
    var begin = 0;
    var end = str.length - 1;
    while (begin <= end && str.charCodeAt(begin) < 33) { ++begin; }
    while (end > begin && str.charCodeAt(end) < 33) { --end; }
    return str.substr(begin, end - begin + 1);
}
function LoadHtmlP(StrSendEmail,StrCounter,StrComsign_Lk)
{
		
	StrSendEmail=ftrim(StrSendEmail);
	StrCounter=vbDbl(StrCounter);
	StrComsign_Lk=vbDbl(StrComsign_Lk);
	if( StrSendEmail!="" && StrSendEmail!=SentToEmail_Add.value) SentToEmail_Add.value=StrSendEmail.replace(/&quot;/g,"\"");	
	if( StrCounter!="" && StrCounter!=wrkC.value) wrkC.value=StrCounter;	
	if( StrComsign_Lk!="" && StrComsign_Lk!=SwComsign_Lk.value) SwComsign_Lk.value=StrComsign_Lk;	
	
	
	if( document.getElementById("F1") != null)	
	{
		SwFind=document.getElementById("F1").SentToEmail_Add;
		if(SwFind==undefined) document.getElementById("F1").appendChild(SentToEmail_Add);
		document.getElementById("F1").appendChild(DocType);
		document.getElementById("F1").appendChild(Doc);
		document.getElementById("F1").appendChild(SwComsign_Lk);
		document.getElementById("F1").appendChild(sCurrYear);
		document.getElementById("F1").appendChild(wrkC);
	}
}


</Script>
<input width="1" height="1" id=SentToEmail_Add name=SentToEmail_Add style="display: none;" value="<%=trim(SentToEmail_Add)%>">
<input width="1" height="1" id=DocType   name=DocType style="display: none;" value="<%=trim(DocType)%>">
<input width="1" height="1" id=Doc name=Doc style="display: none;" value="<%=trim(Doc)%>">
<input width="1" height="1" id=SwComsign_Lk  name=SwComsign_Lk style="display: none;" value="<%=trim(SwComsign_Lk)%>">
<input width="1" height="1" id=sCurrYear  name=sCurrYear style="display: none;" value="<%=trim(CurrYear)%>">
<input width="1" height="1" id=wrkC  name=wrkC style="display: none;" value="<%=trim(wrkC)%>">
