<%@ Language=VBScript %>
<!-- #include file="../../../HebrewMeta.jv"-->
<!-- #include file="../../../j-Document.jv"-->
<!-- #include file="../../../table_Const.jv"-->
<!-- #include file="../../../style_TopBar.jv"-->
<!-- #include file="../../../style_Input.jv"-->
<!-- #include file="../../../style_Button.jv"-->
<!-- #include file="../../../style_Order.jv"-->

<%
Odbc=Request("Odbc")
OdbcUserName=Request("OdbcUserName")
OdbcPassword=Request("OdbcPassword")
FrameName=Request("FrameName")
Company=Request("CurrCompany")
Y=Request("CurrYear")
Usr=Request("UserCounter")
Count=Request("Count")


'---------- new ----------
set Build_Asp=createobject("Build_ASP.Main_B")
path=Request.ServerVariables("URL")
str_HTML=Build_Asp.Main(cstr(FrameName),cstr(path),cstr(Odbc),cstr(UserCounter),"1","","576") 
set Build_Asp=nothing


%>

<LINK href="../../../Max2000.css" type=text/css rel=stylesheet>
<SCRIPT LANGUAGE=javascript>

var UpdateControls=new Array("delRec","editRec","PerutPrt");

var OrderControls=new Array("Daf","SfiraDate","C");
var OrderControlsCell=new Array(6,5,7);
var OrderControlsType=new Array("T","D","N");

var FrameName= <%=Request("FrameName")%>;
var FromFrame= "<%=Request("FromFrame")%>";
var OrderWeb = "SfiraTavimV";

var OrderObjects = new Array("wSnif","wFindDate");
var OrderObjectsType = new Array("R","D"); 

var RowsInTableMax=18;
var b = new Object();
var Order;

var UpdateFunction=new Array(SetBut,null);

function onLoad()
{
	b.TopBar = TopBar;
	b.Marker = null;
	b.FrameNm = FrameName;
	b.Title = document.title;
	b.SetTblKey = true;
	b.SetOrder = "Daf";
	//b.ButP=0;
	top.C.loadProgram(b);
	top.C.createCalender(FrameName,CalnDate,"wFindDate");
	
}	
function execTable(Mode)
{
	var wrkCounter=0;
	i=LineMark;
	if (i > -1) 
	{
		wrkCounter=top.C.getCellValue(tbl,i,7);
	}	
	top.S.runProgram("Kupa/Tavim/SfiraTavim/SfiraTavimU.asp?Counter="+wrkCounter+"&Mode='"+Mode+"'&FromFrame='"+FrameName+"'");
}


function SetBut()
{
	var cnt=0;
	var	i=LineMark;
	if (i <= -1) 
	{
		delRec.style.visibility="hidden"; 
		editRec.style.visibility="hidden"; 
	}
}

</SCRIPT>

<SCRIPT FOR=delRec EVENT="onclick()" >
	execTable('DELETE');
</SCRIPT>
<SCRIPT FOR=editRec EVENT="onclick()" >
	execTable('UPDATE');
</SCRIPT>
<SCRIPT FOR=newRec EVENT="onclick()" >
	execTable('ADD');
</SCRIPT>
<SCRIPT FOR=SfiraDate EVENT="onclick()" >
	top.C.setOrder('SfiraDate',FrameName);
</script>
<SCRIPT  FOR=Daf EVENT="onclick()" >
	top.C.setOrder('Daf',FrameName);
</script>
<SCRIPT  FOR=wSnif EVENT="onblur()" >
	top.C.setOrder(Order,FrameName);
</script>
<SCRIPT FOR=Lbl_Snif EVENT="onclick()" >
	top.C.AllTable(wSnif,Order,FrameName)
</script>

<SCRIPT  FOR=wFindDate EVENT="onblur()" >
	top.C.CheckDate(wFindDate);
	top.C.setOrder(Order,FrameName);
</SCRIPT>
<SCRIPT FOR=Lbl_wFindDate EVENT="onclick()" >
	top.C.AllTable(wFindDate,Order,FrameName);
</SCRIPT>

<SCRIPT FOR=SfiraDate EVENT="onclick()" >
	top.C.setOrder('SfiraDate',FrameName);
</script>
<SCRIPT  FOR=Daf EVENT="onclick()" >
	top.C.setOrder('Daf',FrameName);
</script>

<SCRIPT  FOR=PerutPrt EVENT="onclick()" >
 var wrkCounter=0;
	i=LineMark;	
	if (i > -1) 
	{
		wrkCounter=top.C.getCellValue(tbl,i,7);
		wrkDaf=top.C.getCellValue(tbl,i,6);
	}	
	top.S.runProgram("Kupa/Tavim/SfiraTavim/SfiraTavim_LinesV.asp?Counter="+wrkCounter+"&Daf="+wrkDaf+"&FromFrame='"+FrameName+"'");
</SCRIPT>


<HTML xmlns:TopBar xmlns:order>
<HEAD>
<TITLE>ספירת תווי קניה</TITLE>
</HEAD>
<BODY id=bdy onload="onLoad()" class="Iframe">
<TopBar:o id=TopBar class="TopBar" style="WIDTH: 580px; HEIGHT: 24px"></TopBar:o>&nbsp;
<%=str_HTML%>
<table  class="ButtonTable"  style="WIDTH: 578px; HEIGHT: 8px; TOP: 76px; LEFT: 4px" cellPadding=1 cellSpacing=0>
<tr >
	<td align="right" >
		<span id=CalnDate><input id=wFindDate dir=rtl style="LEFT: 86px; WIDTH: 70px; TOP: 49px" maxLength=10 tabIndex=3></span>
		<span  id=Lbl_wFindDate dir =ltr class="NmField_New">:תאריך</span>&nbsp;&nbsp;
		<input id=wSnif dir=rtl style="WIDTH: 111px; HEIGHT: 18px" maxLength=4 tabIndex=2 size=29>
		<span id=Lbl_Snif dir=ltr class="NmField_New">:סניף</span>&nbsp;&nbsp;
	</td> 
</tr>
</table>

<TABLE id=tbl  class="Table" style="WIDTH: 575px;   HEIGHT: 26px; TOP: 103px; LEFT: 4px">
<thead>
  <TR> 
	<TD id="id:ScmTavim;Type:NumN20;" class="Hc" width=100>סכום תווי קניה</TD> 
	<TD id="id:CmtTavim;" class="Hc" width=100>כמות תווי קניה</TD> 
    <td id="id:SnifNm;Type:Text;" class="Hc" width=170>שם סניף</td>
    <TD id="id:Snif;Type:Text;" class="Hc" width=45>סניף</TD> 
    <td id="id:SfiraTime;Type:TimeShort;" class="Hc" width=30>שעה</td>
    <td id="id:SfiraDate;Type:Date;" class="Hc" width=60><order:o id=SfiraDate style="WIDTH: 60px" >תאריך</order:o></td>
    <TD id="id:Daf;" class="Hc" width=70><order:o id=Daf style="WIDTH: 70px">דף</order:o></TD> 
    <TD id="id:C;" class="Hc" width=0 style="DISPLAY: none">Counter</TD>
  </TR>
</thead>   
<TBODY>
</TBODY>
</TABLE>

</BODY>
<OBJECT id=sDo RUNAT="server" PROGID="NtvDB.Do"></OBJECT>
</HTML>

