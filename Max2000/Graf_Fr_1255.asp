<!-- #include file="include_Session.jv"-->

<HTML>
<HEAD>
<!-- #include file="HebrewMeta.jv"-->
<link rel="stylesheet" type="text/css" href="Max2000.css">
</HEAD>
<!-- #include file="inc_ChkSession_Func.jv"-->
<%
    ' טסט שחחחלללל
    ' 4

function getVfS(str,fieldNm)
	nm = "##"+fieldNm+":"
	s = ""
	j=InStr(1,str,nm)
	if (j > 0) then j1=InStr(j+1,str,"%^")
	if (j1 >0) then s=Mid(str,j+Len(nm),j1-j-Len(nm))
	
	getVfS=s
end function

ZorbaLk=Session("ZorbaLk")
UserCounter=Session("UserCounter")

Odbc="Max2000_"+trim(cstr(Session("ZorbaLk")))
if SwQ_T="1" then
	ZorbaLk=get_Request("ZorbaLk")
	UserCounter=get_Request("UserCounter")
	Odbc="Max2000_"+trim(cstr(get_Request("ZorbaLk")))
end if

str=sM.Main(cstr(ZorbaLk),cstr(UserCounter)) 
swDbAll=getVfS(str,"swDbAll")
swLogIns=getVfS(str,"swLogIns")
swFin=getVfS(str,"swAcc")
swErp=getVfS(str,"swErp")
swHlpD=getVfS(str,"swHlpD")
swHours=getVfS(str,"swHours")
swTl=getVfS(str,"swTl")
swBnk=getVfS(str,"swBnk")
swLk=getVfS(str,"swLk")
swSpk=getVfS(str,"swSpk")
swMlay=getVfS(str,"swMlay")
swSell=getVfS(str,"swSell")
swBuy=getVfS(str,"swBuy")
swIzur=getVfS(str,"swIzur")
swIvo=getVfS(str,"swIvo")
swTk=getVfS(str,"swTk")
swTz=getVfS(str,"swTz")
swAlv=getVfS(str,"swAlv")
swRech=getVfS(str,"swRech")
swCars=getVfS(str,"swCars")
swIvoIzu=getVfS(str,"swIvoIzu")
swSacar=getVfS(str,"swSacar")
swDivor=getVfS(str,"swDivor")
swBakara=getVfS(str,"swBakara")
swMznAndN=getVfS(str,"swMznAndN")
swMusdot=getVfS(str,"swMusdot")
swCompAsh=getVfS(str,"swCompAsh")
swMisrad=getVfS(str,"swMisrad")
swAfaza=getVfS(str,"swAfaza")
swTikshoret=getVfS(str,"swTikshoret")
swStoreNext=getVfS(str,"swStoreNext")
swKupa=getVfS(str,"swKupa")
swToto=getVfS(str,"swToto")
swBasisN=getVfS(str,"swBasisN")
swKatalog=getVfS(str,"swKatalog")
swNitPiz=getVfS(str,"swNitPiz")
swMComp=getVfS(str,"swMComp")
swNitEski=getVfS(str,"swNitEski")
swBithon=getVfS(str,"swBithon")
swMoadonLk=getVfS(str,"swMoadonLk")'swSell
'-----------------------------------------

c=0
t1=""
i_4="display:none"
i_5=""
if cstr(swFin)<>"1" then 
	t1="display:none"
	i_4="display:none"
	c=c+1
end if	
if cstr(swTikshoret)<>"1" then	i_5="display:none;"

t2=""
if cstr(swLk)<>"1" then
	t2="display:none;"
	c=c+1
end if	
t3=""
if cstr(swSpk)<>"1" then
	t3="display:none"
	c=c+1
end if	
t4=""
if cstr(swBnk)<>"1" then
	t4="display:none"
	c=c+1
end if	
t5=""
if cstr(swTk)<>"1" then
	t5="display:none;"
	c=c+1
end if	
t6=""
if cstr(swTz)<>"1" then
	t6="display:none;"
	c=c+1
end if	
t8=""
if cstr(swAlv)<>"1" then
	t8="display:none;"
	c=c+1
end if	
t10=""
if cstr(swMznAndN)<>"1" then
	t10="display:none;"
	c=c+1
end if	
t50=""
if cstr(swMusdot)<>"1" then
	t50="display:none;"
	c=c+1
end if
t51=""
if cstr(swCompAsh)<>"1" then
	t51="display:none;"
	c=c+1
end if
t52=""
if cstr(swNitEski)<>"1" then
	t52="display:none;"
	c=c+1
end if	
tBithon=""
if cstr(swBithon)<>"1" then
	tBithon="display:none;"
	c=c+1
end if	

i1=""
if (c=12) then i1="display:none;" 

tToto=""
if cstr(swToto)<>"1" then	tToto="display:none;"

t7=""
if cstr(swRech)<>"1" then	t7="display:none;"
t11=""
if cstr(swCars)<>"1" then	t11="display:none;"
 
i2=""		
if cstr(swSacar)<>"1" then	i2="display:none;"
c=0
t31=""
if cstr(swSell)<>"1" then
	t31="display:none;"
	c=c+1
end if
tNitPiz=""
if cstr(swNitPiz)<>"1" then
	tNitPiz="display:none;"
	c=c+1
end if
tMComp=""
if cstr(swMComp)<>"1" then
	tMComp="display:none;"
	c=c+1
end if
t32=""	
if cstr(swBuy)<>"1" then
	t32="display:none;"
	c=c+1
end if
t34=""
if cstr(swMlay)<>"1" then
	t34="display:none"
	c=c+1
end if	
tKupa=""
if cstr(swKupa)<>"1" then	
	tKupa="display:none;"
	c=c+1
end if	
tKatalog=""
if cstr(swKatalog)<>"1" then	
	tKatalog="display:none;"
	c=c+1
end if	

i3=""	
if (c=7) then i3="display:none;"  

t33=""	
if cstr(swIvo)<>"1" then
	t33="display:none;"
end if	

i4=""	
if cstr(swIzur)<>"1" then i4="display:none;"

i_6=""
if cstr(swStoreNext)<>"1" then	i_6="display:none;"

i5=""	
if cstr(swDivor)<>"1" then	i5="display:none;"	
i6=""	
if cstr(swHours)<>"1" then	i6="display:none;"
i7=""	
if cstr(swTl)<>"1" then	i7="display:none;"
i9=""	
if cstr(swLogIns)<>"1" then	i9="display:none;"	
i10=""	
if cstr(swDbAll)<>"1" then	i10="display:none;"
BasisN=""
if cstr(swBasisN)<>"1" then	BasisN="display:none;"

i11=""	
if cstr(swBakara)<>"1" then	i11="display:none;"
i_2=""	
if cstr(swMisrad)<>"1" then i_2="display:none;"  
i_3=""	
if cstr(swAfaza)<>"1" then	i_3="display:none;"	

t9=""
if cstr(swIvoIzu)<>"1" then	t9="display:none;"

t95=""
if cstr(swMoadonLk)<>"1" then t95="display:none;"

'--------------------------------------------

%>
<SCRIPT LANGUAGE=javascript>
<!--
var FrameName="FrameSystemsU";     
var Lk="<%=Request("Lk")%>";     
var UserCounter="<%=Request("UserCounter")%>";
function GoDbAll()	{startApp(90);}
function GoLogIns()	{startApp(80);}
function GoAcc()	{startApp(50);}
function GoMznAndN(){startApp(58);}
function GoHours()	{startApp(62);}
function GoTlunot()	{startApp(70); }
function GoBnk()	{startApp(53); }
function GoLk()		{startApp(51); }
function GoSpk()	{startApp(52);}
function GoTk()		{startApp(54);}
function GoBakara()	{startApp(30);}
function GoTz()		{startApp(55);}
function GoMlay()	{startApp(40);}
function GoSell()	{startApp(41);}
function GoBuy()	{startApp(42);}
function GoIzur()	{startApp(45);}
function GoIvo()	{startApp(43);}
function GoAlvaot()	{startApp(57);}
function GoCars()	{startApp(46);}
function GoRechush(){startApp(56);}
function GoIvoIzu()	{startApp(59);}
function GoDivor()	{startApp(65);}
function GoMusdot()	{startApp(64);}
function GoNitEsk()	{startApp(63);}
function GoCompAsh(){startApp(66);}
function GoMisrad() {startApp(200);}
function GoAfaza()	{startApp(72);}
function GoADoc()	{startApp(73);}
function GoTikshoret() {startApp(300);}
function GoStoreNext() {startApp(310);}
function GoToto() {startApp(400);}
function GoKupa() {startApp(47);}
function GoKat() {startApp(48);}
function GoBasisN() {startApp(91);}
function GoNitPiz() {startApp(49);}
function GoBithon() {startApp(500);}
function GoMComp() {startApp(510);}
function GoMoadonLk() {startApp(95);}
function GoHanut() {startApp(96);}


function startApp(ZorbaApplication)
{
	parent.onOK(ZorbaApplication);
}
function onMouseOut(a,b)
{
	document.all(a).style.color="#ff7f00";	
	document.all(b).style.color="black";
}
function onMouseOver(a,b)
{
	document.all(a).style.color="forestgreen";	
	document.all(b).style.color="green";
}

function onOut(a,b)
{
	document.all(a).style.color="#ff7f00";
	document.all(b).style.textDecoration='';	
	document.all(b).style.color="black";
}
function onOver(a,b)
{
	document.all(a).style.color="forestgreen";	
	document.all(b).style.textDecoration='underline';	
	document.all(b).style.color="green";
}

//-->
</SCRIPT>
<script LANGUAGE="vbscript">
function vbNow()
	vbNow=Cstr(Day(Date()))+"_"+Cstr(hour(Now()))+"_"+Cstr(Minute(Now()))+"_"+Cstr(Second(Now()))
end function
</SCRIPT>
<style>
	.ButA {text-decoration:underline;}
</style>
<BODY id=bdy onload="onLoad()" style="BACKGROUND-COLOR: transparent">
<center>
<table border=0  style="FONT-SIZE: 11px; WIDTH: 85%; FONT-FAMILY: tahoma; TEXT-ALIGN: right">
<%
SwSQL=session("SwSQL")
if SwQ_T="1" then SwSQL=get_Request("SwSQL")

Set BuildConn = CreateObject("Build_ConnString.Main")
connStr = BuildConn.bConnString(cstr(Odbc), cstr(SwSQL))
Set BuildConn = Nothing

Conn.Open connStr'cstr(Odbc),"sa",""

sql="SELECT Msg FROM LkMsg_User WHERE Usr="+cstr(UserCounter)+" and D<='" & F.ChangeMD(Date()) & "' and ((ToDate is null) or ToDate>='" & F.ChangeMD(Date()) & "') ORDER BY convert(datetime,D) DESC,C DESC "
Rs.Open sql,Conn,1,1
dMsgUser=0
if not Rs.EOF then 
%>
	<tr>
		<td  align=middle style="FONT-WEIGHT: bold; FONT-SIZE: 10px; COLOR: green" >לוח מודעות אישי</td>
	</tr>
	<tr>
		<td  bgcolor=white dir=rtl style="BORDER-RIGHT: green 1px solid; BORDER-TOP: green 1px solid; BORDER-LEFT: green 1px solid; BORDER-BOTTOM: green 1px solid" >
<%
end if
do while not Rs.EOF 
	str=""
	if dMsgUser>0 then str=str+"<hr  SIZE=1 >"
	msg=replace(Rs("Msg"),chr(13)+chr(10),"<br>")
	msg=replace(msg,chr(13),"<br>")
	msg=replace(msg,chr(32)+chr(32)," &nbsp;")
	str=str+msg
	Response.Write  str
	dMsgUser=dMsgUser+1
	Rs.MoveNext() 
loop
if dMsgUser>0 then Response.Write "</tr>"
Rs.Close()

Conn.Close()

Set BuildConn = CreateObject("Build_ConnString.Main")
connStr = BuildConn.bConnString("Max2000_BackOffice", "")
Set BuildConn = Nothing
 
Conn.Open connStr'"Max2000_BackOffice","sa",""
sql=" SELECT Msg FROM Msg_Lk WHERE Lk="+cstr(ZorbaLk)+" and  D<='" & F.ChangeMD(Date())  & "' and ( isnull(ToDate,'')='' or ToDate>='" & F.ChangeMD(Date()) & "') ORDER BY convert(datetime,D) DESC,C DESC "
Rs.Open sql,Conn,1,1
dMsgAll=0
if not Rs.EOF  then 
%>
	<tr >
		<td  align=middle        
    style="FONT-WEIGHT: bold; FONT-SIZE: 10px; COLOR: orange">לוח מודעות לארגון</td>
	</tr>
	<tr  >
		<td bgcolor=white  dir=rtl style="BORDER-RIGHT: orange 1px solid; BORDER-TOP: orange 1px solid; BORDER-LEFT: orange 1px solid; BORDER-BOTTOM: orange 1px solid"  >
<%
end if
do while not Rs.EOF 
	str=""
	if dMsgAll>0 then str=str+"<hr  SIZE=1 >"
	msg=replace(Rs("Msg"),chr(13)+chr(10),"<br>")
	msg=replace(msg,chr(13),"<br>")
	msg=replace(msg,chr(32)+chr(32)," &nbsp;")
	str=str+msg
	Response.Write  str
	dMsgAll=dMsgAll+1
	Rs.MoveNext() 
loop
if dMsgAll>0 then Response.Write "</tr>"
Rs.Close()

Conn.Close()
 
%></tr>		
</table>
<br>
<table  cellpadding=0 style="FONT-SIZE: 11px; WIDTH: 99%; FONT-FAMILY:  tahoma" cellspacing=0>
	<tr id=i_2  style="<%=i_2%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span >
					<span onclick=GoMisrad() id=s1 onmouseover="onOver('a28','s1')" onmouseout="onOut('a28','s1')"  dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >לקוחות,  עובדים, דיווח  שעות וניתוחי השקעה ורווחיות, 
					בקרה כוללת לתיקים, ניתוחים סטטיסטיים ומעקב</span>
					&nbsp;-&nbsp;<A class=ButA id=a28 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a28','s1'); " 
					 
					onmouseover="javascript:onMouseOver('a28','s1');" 
            onclick=GoMisrad()>ניהול משרד</A>
					<br></span>
			
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i1  style="<%=i1%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span id=t1  style="<%=t1%>"><span onclick="javascript:GoAcc();" id=s2 onmouseover="onOver('a1','s2')" onmouseout="onOut('a1','s2')"  style="CURSOR: hand" 
            dir=rtl 
           >חשבונות, פקודות יומן, התאמות כרטיסים פנימיות וחיצוניות, ספרים ושאילתות</span>
						&nbsp;-&nbsp;<A class=ButA id=a1 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout    
							   ="javascript:onMouseOut('a1','s2');" 
            onmouseover="javascript:onMouseOver('a1','s2');" 
            onclick="javascript:GoAcc();">פקודות וספרים</A> 
                       
					<br></span>
					<span id=t10  style="<%=t10%>"><span onclick="javascript:GoMznAndN();"  id=s3 onmouseover="onOver('a24','s3')" onmouseout="onOut('a24','s3')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           >מאזני בוחן ומאזני חברה</span>
						&nbsp;-&nbsp;<A class=ButA id=a24 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a24','s3'); " 
            onmouseover="javascript:onMouseOver('a24','s3')" 
            onclick=GoMznAndN()>מאזנים</A><br>
					</span>
					<span id=t52  style="<%=t52%>"><span onclick="javascript:GoNitEsk();"  id=s4 onmouseover="onOver('a27','s4')" onmouseout="onOut('a27','s4')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           >ניתוח הכנסות, הוצאות, רווח והפסד, גרפים, דוחות השוואה, ניתוחי התפלגות ומגמה</span>
						&nbsp;-&nbsp;<A class=ButA id=a27 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a27','s4'); " 
            onmouseover="javascript:onMouseOver('a27','s4')" 
            onclick=GoNitEsk()>ניתוח עסקי</A><br>
					</span>
					<span id=t2  style="<%=t2%>"><span onclick="javascript:GoLk();"  id=s5 onmouseover="onOver('a9','s5')" onmouseout="onOut('a9','s5')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           > קבלות, חשבוניות, שוברים, הפקדות, שיקים חוזרים, דוחות גביה, מכתבי פיגורים וניתוחים</span>
						&nbsp;-&nbsp;<A class=ButA id=a9 style="WIDTH: 95px;  CURSOR: hand;  COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout                    
			  ="javascript:onMouseOut('a9','s5');" onmouseover="javascript:onMouseOver('a9','s5');" onclick=GoLk()>לקוחות/קופה</A>
					<br></span>
					<span id=t3  style="<%=t3%>"><span        
            onclick="javascript:GoSpk();"  id=s6 onmouseover="onOver('a10','s6')" onmouseout="onOut('a10','s6')"      dir=rtl 
            style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >חשבוניות ספקים, הוראות תשלום,  מקדמות, פקודות תשלום, הדפסת שיקים, גיול חובות ומאגר היסטורי</span>
						&nbsp;-&nbsp;<A class=ButA id=a10 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a10','s6'); " onmouseover="javascript:onMouseOver('a10','s6');" onclick=GoSpk() >ספקים/תשלומים</A>
					<br></span>
					<span id=t50  style="<%=t50%>"><span        
            onclick="javascript:GoMusdot();"  id=s7 onmouseover="onOver('a25','s7')" onmouseout="onOut('a25','s7')" dir=rtl    style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול ודיווח מע"מ, ביטוח לאומי, מס הכנסה וניכוי במקור</span>
						&nbsp;-&nbsp;<A class=ButA id=a25 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a25','s7'); " onmouseover="javascript:onMouseOver('a25','s7');" onclick=GoMusdot() >מוסדות</A>
					<br></span>
					<span id=t4  style="<%=t4%>"><span onclick="javascript:GoBnk();"  id=s8 onmouseover="onOver('a8','s8')" onmouseout="onOut('a8','s8')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >יבוא וקליטת דפי בנק, ביצוע התאמות 
            אוטומטיות וידניות, דוחות ובקרת שלמות</span>
						&nbsp;-&nbsp;<A class=ButA id=a8 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout     
							="javascript:onMouseOut('a8','s8'); " 
            onmouseover="javascript:onMouseOver('a8','s8');" 
            onclick=GoBnk()>בנקים</A>
					<br></span>
					<span id=t51  style="<%=t51%>"><span onclick="javascript:GoCompAsh();"  id=s9 onmouseover="onOver('a26','s9')" onmouseout="onOut('a26','s9')"       
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           >חשבוניות עמלה, התאמה אוטומטית וידנית, דוחות ובקרת שלמות </span>
						&nbsp;-&nbsp;<A class=ButA id=a26 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a26','s9'); " onmouseover="javascript:onMouseOver('a26','s9');" onclick=GoCompAsh() >חברות אשראי</A>
					<br></span>
					<span id=t5  style="<%=t5%>"><span         
            onclick="javascript:GoTk();"  id=s10 onmouseover="onOver('a11','s10')" onmouseout="onOut('a11','s10')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >תיקצוב, מעקב תקציבי וניתוח תקציב מול ביצוע בפועל</span>
						&nbsp;-&nbsp;<A class=ButA id=a11 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a11','s10'); " 
            onmouseover="javascript:onMouseOver('a11','s10');" 
            onclick=GoTk()>תקציב</A>
					<br></span>
					<span id=t6  style="<%=t6%>"><span onclick="javascript:GoTz();"  id=s11 onmouseover="onOver('a12','s11')" onmouseout="onOut('a12','s11')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >תזרים מזומנים, בקרת אשראי, צפי 
            כולל תחזית גבייה ותשלומי ספקים</span>
						&nbsp;-&nbsp;<A class=ButA id=a12 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a12','s11'); " 
            onmouseover="javascript:onMouseOver('a12','s11');" 
            onclick=GoTz()>תחזית</A>
					<br></span>
					<span id=t8  style="<%=t8%>"><span onclick="javascript:GoAlvaot();"  id=s12 onmouseover="onOver('a19','s12')" onmouseout="onOut('a19','s12')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >רשימת הלוואות, חישוב פרעונות, 
            תחזית פרעונות, חיבור לתזרים מזומנים</span>
						&nbsp;-&nbsp;<A class=ButA id=a19 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a19','s12'); " 
            onmouseover="javascript:onMouseOver('a19','s12');" 
            onclick=GoAlvaot()>הלוואות</A>
					<br></span>
					<span id=tBiton  style="<%=tBiton%>"><span onclick="javascript:GoBithon();"  id=sBithon onmouseover="onOver('aBithon','sBithon')" onmouseout="onOut('aBithon','sBithon')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול ביטחונות וערבויות ספקים ולקוחות</span>
						&nbsp;-&nbsp;<A class=ButA id=aBithon style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aBithon','sBithon'); " 
            onmouseover="javascript:onMouseOver('aBithon','sBithon');" 
            onclick=GoBithon()>ביטחונות</A>
					<br></span>
					</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr  id=i_4  style="<%=i_4%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span >
					<span onclick="javascript:GoADoc();"  id=s15 onmouseover="onOver('a30','s15')" onmouseout="onOut('a30','s15')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" >ניהול ארכיון מסמכי מקור של המערכת הפיננסית</span>
					&nbsp;-&nbsp;<A class=ButA id=a30 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a30','s15'); " 
					onmouseover="javascript:onMouseOver('a30','s15');" onclick=GoADoc()>ארכיון מסמכים</A>
					<br></span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i3   style="<%=i3%>">
		<td >
		<table  border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right >
					<span id=t31  style="<%=t31%>" ><span  onclick="javascript:GoSell();"  id=s19 onmouseover="onOver('a13','s19')" onmouseout="onOut('a13','s19')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >הצעות מחיר, הזמנות, תעודות משלוח,   תעודות החזר, 
            חשבוניות, חשבוניות מרכזות, ניתוח ביצועים, דוחות 
            תרומה, השוואת תקופות וניתוח הנחות</span>
						&nbsp;-&nbsp;<A class=ButA id=a13 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout        
            ="javascript:onMouseOut('a13','s19'); " 
            onmouseover="javascript:onMouseOver('a13','s19');" 
            onclick=GoSell()>מכירות</A>
					<br></span>
					<span id=t32  style="<%=t32%>" ><span onclick="javascript:GoBuy();"  id=s20 onmouseover="onOver('a14','s20')" onmouseout="onOut('a14','s20')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           >הזמנות, יבואים,  תעודות כניסה, EDI,  תעודות החזר, חשבוניות, 
            חשבוניות מרכזות, מחירוני רכש והשוואות 
            מחירים, דוחות וניתוח רכישות 
            לפי פריטים/ספקים/תקופות</span>
			&nbsp;-&nbsp;<A class=ButA id=a14 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout          ="javascript:onMouseOut('a14','s20'); " 
      onmouseover="javascript:onMouseOver('a14','s20');" 
      onclick=GoBuy()>רכש</A>
					<br></span>
					<span id=t34  style="<%=t34%>" ><span onclick="javascript:GoMlay();"  id=s21 onmouseover="onOver('a16','s21')" onmouseout="onOut('a16','s21')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >הגדרת מחסנים, יתרות פתיחה, הזמנות פנימיות, תעודות 
            העברה, תעודות המרה, ספירות מלאי, ניתוח ושווי מלאי</span>
					&nbsp;-&nbsp;<A class=ButA id=a16 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a16','s21'); " 
      onmouseover="javascript:onMouseOver('a16','s21');" 
      onclick=GoMlay()>מלאי</A>
					</span>
					<span  style="<%=tKatalog%>" ><span onclick="javascript:GoKat();"  id=sKat onmouseover="onOver('aKat','sKat')" onmouseout="onOut('aKat','sKat')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >הגדרת פריטים, מחלקות, קבוצות, איפיונים, ניהול מחירוני קניה ומכירה, תימחור ודוחות השוואה</span>
					&nbsp;-&nbsp;<A class=ButA id=aKat style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aKat','sKat'); " 
      onmouseover="javascript:onMouseOver('aKat','sKat');" 
      onclick=GoKat()>קטלוג</A>
					</span>

					<span  style="<%=tNitPiz%>" ><span onclick="javascript:GoNitPiz();"  id=sNit onmouseover="onOver('aNit','sNit')" onmouseout="onOut('aNit','sNit')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניתוח ביצועי רכש, מכירות ומלאי, דוחות תרומה, השוואות, רווח והפסד, ניתוח הנחות, חקר התנהגות ושווי מלאי</span>
					&nbsp;-&nbsp;<A class=ButA id=aNit style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aNit','sNit'); " 
      onmouseover="javascript:onMouseOver('aNit','sNit');" 
      onclick=GoNitPiz()>ניתוח ביצועים</A>
					</span>
					<span  style="<%=tMComp%>" ><span onclick="javascript:GoMComp();"  id=sMComp onmouseover="onOver('aMComp','sMComp')" onmouseout="onOut('aMComp','sMComp')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניתוח ביצועי ארגון מרובה חברות בנות או זכיינים</span>
					&nbsp;-&nbsp;<A class=ButA id=aMComp style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aMComp','sMComp'); " 
					     onmouseover="javascript:onMouseOver('aMComp','sMComp');" 
      onclick=GoMComp()>רב חברתי</A>
					</span>

					<span  style="<%=tKupa%>"><span onclick="javascript:GoKupa();"  id=sKupa onmouseover="onOver('aKupa','sKupa')" onmouseout="onOut('aKupa','sKupa')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" >
		ממשק לקופות רושמות, עדכון פריטים ומחירונים, יבוא מכר ותקבולים, ניתוח פיתקיות שליליות, הנחות והחזרות פריטים
            </span>
					&nbsp;-&nbsp;<A class=ButA id=aKupa style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aKupa','sKupa'); " 
            onmouseover="javascript:onMouseOver('aKupa','sKupa');" 
            onclick=GoKupa()>קופות רושמות</A>
					</span>
					</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr    style="<%=t95%>"  >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span  style="<%=t95%>" ><span onclick="javascript:GoMoadonLk();"  id=s95 onmouseover="onOver('a95','s95')" onmouseout="onOut('a95','s95')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול מועדון לקוחות. צבירת מכירות\ניקוד , ניתוח ודיווח במאייל, פקס ו-SMS</span>
					&nbsp;-&nbsp;<A class=ButA id=a95 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a95','s95'); " 
            onmouseover="javascript:onMouseOver('a95','s95');" 
            onclick=GoMoadonLk()>מועדון לקוחות</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr    style="<%=t95%>"  >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span  style="<%=t95%>" ><span onclick="javascript:GoHanut();"  id=s96 onmouseover="onOver('a96','s96')" onmouseout="onOut('a96','s96')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול חנות וירטואלית. הגדרות, ניהול תוכן, מעקב וניתוחים סטטיסטיים של שימוש וביצועים</span>
					&nbsp;-&nbsp;<A class=ButA id=a96 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a96','s96'); " 
            onmouseover="javascript:onMouseOver('a96','s96');" 
            onclick=GoHanut()>חנות וירטואלית</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr  id=i4  style="<%=i4%>"  >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span id=i4  style="<%=i4%>" ><span onclick="javascript:GoIzur();"  id=s22 onmouseover="onOver('a17','s22')" onmouseout="onOut('a17','s22')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >תכנון והוראות ייצור, תחזית שימוש 
            בחומר גלם, דיווח  ייצור, גריעת מלאי אוטומטית, דוחות וניתוח ביצועים</span>
					&nbsp;-&nbsp;<A class=ButA id=a17 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a17','s22'); " 
            onmouseover="javascript:onMouseOver('a17','s22');" 
            onclick=GoIzur()>ייצור</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i6  style="<%=i6%>" >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span  onclick="javascript:GoHours();"  id=s17 onmouseover="onOver('a6','s17')" onmouseout="onOut('a6','s17')"        
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 
            78%; CURSOR: hand" 
            id=SPAN1>דיווח שעות עבודה לפי לקוח 
            ופרוייקט, מחירונים, חשבוניות ללקוחות בגין 
            שעות עבודה כולל חיובים, ניתוחי תרומה והשוואה</span>
				&nbsp;-&nbsp;<A class=ButA id=a6 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a6','s17'); "     
      onmouseover="javascript:onMouseOver('a6','s17');" 
      onclick=GoHours()>שעות עבודה</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i5   style="<%=i5%>">
		<td >
		<table  border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span onclick="javascript:GoDivor();"  id=s18 onmouseover="onOver('a23','s18')" onmouseout="onOut('a23','s18')"      
                 
            dir=rtl 
            style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >  ניהול משימות, מאגר לקוחות, 
            סיכומי פגישות, ביצוע דיוור, מעקב מתעניינים, ניתוח ודוחות סטטיסטים</span>
				&nbsp;-&nbsp;<A id=a23                   
      class=ButA onmouseout 
      ="javascript:onMouseOut('a23','s18') 
      &#13;&#10; &#13;&#10; "       
            onmouseover ="javascript:onMouseOver('a23','s18')" onclick       
           ="GoDivor()" style       
           ="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" 
           >שיווק ומשימות</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	
	<tr id=t7  style="<%=t7%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span onclick="javascript:GoRechush();"  id=s13 onmouseover="onOver('a18','s13')" onmouseout="onOut('a18','s13')"        dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול רכוש וחישובי פחת</span>
					&nbsp;-&nbsp;<A class=ButA id=a18 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a18','s13'); " 
					 
					onmouseover="javascript:onMouseOver('a18','s13');" 
            onclick=GoRechush()>רכוש קבוע</A>	<br>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr  style="<%=t11%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span id=t11  ><span onclick="javascript:GoCars();"  id=s14 onmouseover="onOver('a20','s14')" onmouseout="onOut('a20','s14')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
            id=SPAN3>  רשימת רכבים, מעקב 
						טיפולים,&nbsp; ביטוחים,&nbsp; רישוי, דוחות השוואה וניתוח שימוש</span>
					&nbsp;-&nbsp;<A class=ButA id=a20 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a20','s14'); " 
			 
            onmouseover="javascript:onMouseOver('a20','s14');" 
            onclick=GoCars()>רכבים</A>
					<br></span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr  id=i_5  style="<%=i_5%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span >
					<span onclick="javascript:GoTikshoret();"  id=s16 onmouseover="onOver('a31','s16')" onmouseout="onOut('a31','s16')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" >מעקב שיחות, רישום, דוחות השוואה ניתוח שימוש</span>
					&nbsp;-&nbsp;<A class=ButA id=a31 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a31','s16'); " 
					onmouseover="javascript:onMouseOver('a31','s16');" onclick=GoTikshoret()>תקשורת</A>
					<br></span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i_3  style="<%=i_3%>" >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span onclick="javascript:GoAfaza();"  id=s24 onmouseover="onOver('a29','s24')" onmouseout="onOut('a29','s24')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >הזמנות לאספקה, תאום מועדי אספקה, 
			שיבוץ, דיווח ביצוע, התחשבנות עם לקוחות,
			ניתוח ודוחות סטטיסטיים</span>
					&nbsp;-&nbsp;<A class=ButA id=a29 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a29','s24'); " 
            onmouseover="javascript:onMouseOver('a29','s24');" 
            onclick=GoAfaza()>הפצה</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr   style="<%=tToto%>"  >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span  ><span onclick="javascript:GoToto();"  id=sToto onmouseover="onOver('aToto','sToto')" onmouseout="onOut('aToto','sToto')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" >
		ניהול מועדון מנויים
            </span>
					&nbsp;-&nbsp;<A class=ButA id=aToto style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('aToto','sToto'); " 
            onmouseover="javascript:onMouseOver('aToto','sToto');" 
            onclick=GoToto()>טוטוסטרדה</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i7   style="<%=i7%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span  onclick="javascript:GoTlunot();"  id=s26 onmouseover="onOver('a7','s26')" onmouseout="onOut('a7','s26')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
            CURSOR: hand" 
            id=SPAN2>דיווח ופתיחת קריאת שרות, מעקב 
            טיפולים, ניתוח תלונות  וסטטיסטיקות לפי סוג תקלה, מיקום ומטפלים</span>
				&nbsp;-&nbsp;<A class=ButA id=a7 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a7','s26'); " onmouseover="javascript:onMouseOver('a7','s26');" onclick=GoTlunot()>שרות</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i11  style="<%=i11%>" >
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span ><span onclick="javascript:GoBakara();"  id=s25 onmouseover="onOver('a50','s25')" onmouseout="onOut('a50','s25')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >ניהול פיקוח עירוני</span>
				&nbsp;-&nbsp;<A class=ButA id=a50 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a50','s25'); " 
            onmouseover="javascript:onMouseOver('a50','s25');" 
            onclick=GoBakara()>חניתה</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr   style="<%=BasisN%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right  valign=top>
					<span ><span  onclick="javascript:GoBasisN();"  id=sBasisN onmouseover="onOver('aBasisN','sBasisN')" onmouseout="onOut('aBasisN','sBasisN')"      
                 
            dir=rtl 
            style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >מדדים, מטבעות ושערים, ישובים וטבלאות כלליות</span>
				&nbsp;-&nbsp;<A class=ButA id=aBasisN style="VERTICAL-ALIGN: top; WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout="javascript:onMouseOut('aBasisN','sBasisN'); " 
            onmouseover="javascript:onMouseOver('aBasisN','sBasisN')" 
            onclick=GoBasisN()>בסיס נתונים</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i10   style="<%=i10%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right  valign=top>
					<span ><span  onclick="javascript:GoDbAll();"  id=s27 onmouseover="onOver('a3','s27')" onmouseout="onOut('a3','s27')"      
                 
            dir=rtl 
            style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >הגדרת משתמשים, חברות, מבקרים, 
            בניית תפריטים אישיים וניווט השימוש 
            במערכת תוך דגש רב על הרשאות ומידור לפי הצורך</span>
				&nbsp;-&nbsp;<A class=ButA id=a3 style="VERTICAL-ALIGN: top; WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout="javascript:onMouseOut('a3','s27'); " 
            onmouseover="javascript:onMouseOver('a3','s27')" 
            onclick=GoDbAll()>הגדרות מערכת</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=i9   style="<%=i9%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right  valign=top>
					<span ><span onclick="javascript:GoLogIns();" id=s28 onmouseover="onOver('a2','s28')" onmouseout="onOut('a2','s28')"         
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
            CURSOR: hand" 
            id=SPAN1>כמרכיב משלים בפיתרון הכולל פותחו 
            הכלים המקנים יכולת מעקב וניתוח השימוש 
            במערכת, הן של המשתמשים באירגון עצמו והן של האורחים</span>
				&nbsp;-&nbsp;<A class=ButA id=a2 onmouseout="javascript:onMouseOut('a2','s28'); " onmouseover="javascript:onMouseOver('a2','s28');" onclick=GoLogIns() style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" >מעקב שימוש</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr   style="<%=i_6%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right valign=top>
					<span style="<%=i_6%>"><span onclick="javascript:GoStoreNext();"  id=s23 onmouseover="onOver('a32','s23')" onmouseout="onOut('a32','s23')"       dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; CURSOR: hand" 
           >מסרים אלקטרונים (E.D.I.), תעודת כניסה, תעודת החזר, קופות ותקבולים</span>
					&nbsp;-&nbsp;<A class=ButA id=a32 style="WIDTH: 95px; CURSOR: hand; COLOR:  #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a32','s23'); " 
            onmouseover="javascript:onMouseOver('a32','s23');" 
            onclick=GoStoreNext()>StoreNext</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr id=t9   style="<%=t9%>">
		<td >
		<table   border=0 width="100%" style="FONT-SIZE: 11px" cellSpacing =0  cellPadding=0>
			<tr >
				<td align=right  valign=top>
					<span ><span   onclick="javascript:GoIvoIzu();"  id=s29 onmouseover="onOver('a21','s29')" onmouseout="onOut('a21','s29')"      
            dir=rtl style="VERTICAL-ALIGN: bottom; WIDTH: 78%; 
           CURSOR: hand" 
           >משיכה והעברת  נתונים למערכות זרות</span>
				&nbsp;-&nbsp;<A class=ButA id=a21 style="WIDTH: 95px; CURSOR: hand; COLOR: #ff7f00; TEXT-ALIGN: left" onmouseout  ="javascript:onMouseOut('a21','s29'); " 
            onmouseover="javascript:onMouseOver('a21','s29')" 
            onclick=GoIvoIzu()>יבוא/ייצוא</A>
					</span>
				</td>
			</tr>
			<tr>
				<td ><hr class="HR" width="100%"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</center>
<OBJECT id=sDo PROGID="NtvDB.Do1" RUNAT="server"></OBJECT>
<OBJECT id=F PROGID="SubFunctions.Main" RUNAT="server"></OBJECT>
<OBJECT id=sM PROGID="NtvGate.Main" RUNAT="server"></OBJECT>
<OBJECT id=Conn PROGID="ADODB.Connection" RUNAT="server"></OBJECT>
<OBJECT id=Rs PROGID="ADODB.Recordset" RUNAT="server"></OBJECT>

</BODY>
</HTML>
<SCRIPT LANGUAGE=javascript>
<!--
function onLoad()
{ 
	//if (<%=sw%>==1) parent.clearTime();
	top.R.location.href="Code_Report.asp?d=<%=now()%>";
}
//-->
</SCRIPT>
