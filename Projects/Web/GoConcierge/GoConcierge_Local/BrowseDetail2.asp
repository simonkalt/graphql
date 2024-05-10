<%@ Language=VBScript %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "No-cache"
Response.Expires = -1
%>
<!--#include file = "Global.asp" ---> 
<%

Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))
dim cid, booFirstLoad
cid = remote.Session ("CompanyID")
booFirstLoad = true

'if vartype(remote.Session("ScreenHeight")) <> 8 then
'	remote.Session("ScreenHeight") = 600
'end if

If remote.Session("ScreenHeight") < 750 Then
	Max_Records		= 15
	ToolTipHeight	= 360
	pTop			= 2
	pLeft			= 60
	
	intNameWidth	= 185
	intAddressWidth	= 155
	intCityWidth	= 100
	intPhoneWidth	= 112
	intFaxWidth		= 0
	intMapWidth		= 22
	intMilesWidth	= 40
	intWebWidth		= 30
	intWeb2Width	= 40
	intStarsWidth	= 30
Else
	Max_Records		= 23 '22 
	ToolTipHeight	= 526
	pTop			= 2
	pLeft			= 140
	
	intNameWidth	= 210
	intAddressWidth	= 160
	intCityWidth	= 107
	intPhoneWidth	= 115
	intFaxWidth		= 108
	intMapWidth		= 22
	intMilesWidth	= 40
	intWebWidth		= 30
	intWeb2Width	= 40
	intStarsWidth	= 30
End If

dim RecordCount, LastRec, recnum
RecordCount = 0
LastRec = 0
recnum = 0

' just for live push...
if remote.Session("DefaultCategory") = "" then
	remote.Session("DefaultCategory") = 0
	remote.Session("DefaultBCK") = 1
	remote.Session("DefaultSortBy") = 1
	remote.Session("DefaultState") = 0
end if
'
%>
<HTML>
<HEAD>
<style type="text/css">
<!--
	.bord {border-bottom-width:1px;border-bottom-style:solid;border-bottom-color:#D8BFD8;border-right-width:1px;border-right-style:solid;border-right-color:#D8BFD8;cursor:arrow;height:21px; }
	.bordOT {background-color:lightgreen;border-bottom-width:1px;border-bottom-style:solid;border-bottom-color:#D8BFD8;border-right-width:1px;border-right-style:solid;border-right-color:#D8BFD8;cursor:arrow; }
	.normclass  { font-family : Tahoma; font-size:11px; border-color: black; border-style:1px;cursor:hand }
	TABLE	{ font-family: tahoma; font-size: 11px; }
	A:hover		{color:red}
	A:active	{color:blue}
	A:visited	{color:blue}
	
-->
</style>
<title>Detail</title>
</HEAD>
<BODY id="bdy" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" bgcolor="menu">
<input type=hidden id=txtID value="">
<div ID="tooltip" STYLE="left:0px;font-family: Helvetica; font-size: 8pt; position: absolute; z-index: 200; display: none; visibility: hidden; width:207px">
	<iframe height="<%=ToolTipHeight%>" width="640" frameborder="0" style="border-style: none; border-width: 2px;" src="LocationTooltip.asp" id="frameToolTip" scrolling="no"></iframe>
</div>
<iframe onload=sta() src="LoadingAppointment.asp" id=frameSTA style=display:none;visibility:hidden></iframe>

<script type="text/javascript" language="javascript1.2">

var lastColor = "";
var SelectedBGColor = "";
var SelectedID = -1;
var Red = "#ff9d9d"

function mo( t, i )		{ if(t) HoverOn(t, i);  }
function mout( t, i )	{ if(t) HoverOff(t, i); }
function mosel( t, i )	{ if(t) Selected(t, i); }


function wopen( s )
{
	h = screen.availHeight*.93
	w = screen.availWidth*.97
	<%
	sv = Request.ServerVariables("HTTP_USER_AGENT")
	if instr(1,sv,"MSIE 5.0") > 0 or instr(1,sv,"MSIE 5.1") > 0 then
		response.write "window.showModalDialog(s,'','dialogheight:'+h+'px;dialogwidth:'+w+'px');" & vbcrlf
	else
		response.write "window.open(s);" & vbcrlf
	end if
	%>
}

function HoverOn( t, i )
{
	try
	{
		if(document.all("tr"+i).style.backgroundColor != Red)
		{
			lastColor = document.all("tr"+i).style.backgroundColor
			document.all("tr"+i).style.backgroundColor = "#CFCFF5"
		}
	} 
	catch(x) {}
	finally {}
}

function HoverOff( t, i )
{	
	try
	{
		if(document.all("tr"+i).style.backgroundColor != Red)
			document.all("tr"+i).style.backgroundColor = lastColor;
	}
	catch(x) {}
	finally {}
}

function restoreRowColor()
{
	if(SelectedID != -1)
		document.all("tr"+SelectedID).style.backgroundColor = SelectedBGColor;
}


var clicked=false;
var ms=0;

function Selected( t, i )
{
	var str = t;
	SelectedBGColor = lastColor;

	restoreRowColor()
	/* if(SelectedID != 0)
		document.all("tr"+SelectedID).style.backgroundColor = SelectedBGColor; */
		
	document.all("tr"+i).style.backgroundColor = Red;
	SelectedID = i;
	document.all.txtID.value = t;
	
	parent.document.all("txtText").value = '';//document.all("td_2_"+i).innerText;
	
	
	str += "|" + document.all("tr"+i).childNodes(0).firstChild.firstChild.innerHTML;
	
	for(var j=1;j<6;j++)
	{
		str += "|" + document.all("tr"+i).childNodes(j).firstChild.innerHTML;
	}

		str += "|" + document.all("tr"+i).childNodes(document.all("tr"+i).childNodes.length-1).firstChild.innerHTML;
		
	parent.txtColumns = str;
	
	
	/*
	if (clicked)
		{
			clicked=false;
			showTextActivate(t);
		}
	else
		{
			clicked = true;
			var ms1 = new Date().getMilliseconds();
		}
		
			
	*/		
	
	/* good way to see the resulting html
	var tr = window.bdy.createTextRange();
	tr.expand("textedit");
	window.clipboardData.setData("Text",tr.htmlText);
	*/
}

function toolTip (t,i)
{
	showTextActivate(t);	
}


var ToolTipID = 0;

function showTextActivate(strField) { //, x, y) {
	
	
	window.frames("frameToolTip").document.all("toolTipID").value = strField;
	
	//var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	
	//alert("LocationToolTipDetail.asp?id=" + strField)
	//xmlHttp.open("POST", "LocationToolTipDetail.asp?id=" + strField, false)
	//xmlHttp.send()
	//alert(unescape(xmlHttp.responseText));
	//eval(unescape(xmlHttp.responseText));

	//xmlHttp = null
	document.all("frameSTA").src = "LocationToolTipDetail.asp?id=" + strField;
	document.frames.frameToolTip.document.all.tdLocationID.innerText = "Vendor ID: " + strField;
}	

var booTTLoaded = true;
function sta() { 
   if(booTTLoaded)
   	booTTLoaded = false;
   else
   	staContinue();
   	//alert(unescape(window.frames("frameSTA").document.body.innerHTML));
}

function staContinue() {

  eval(unescape(window.frames("frameSTA").document.body.innerHTML));

  var el = document.all.tooltip
  var frel = document.frames.frameToolTip.document.all.tbdyToolTip
  
  var booLastWasNote = false;
  
  el.style.pixelTop = <%=pTop%>;
  el.style.pixelLeft = <%=pLeft%>;
  var j = frel.rows.length;
  for(i=0;i<j;i++)
	frel.deleteRow(0);
	  
  var str = "", cnt = 0;
  
  a = aTT;
  strColor = a[1];
	  
	
  for(i=2; i<a.length; i++)
  {
	var b = Array();
	b = a[i].split("|");
	b[1] = b[1].replace(/\&lt;<TD>\&gt;/g,"<<TD>>").replace(/\&lt;<SQ>\&gt;/g,"<<SQ>>")
	if(b[1].length > 0)
		{
		oRow = frel.insertRow();
			if(b[1].indexOf("<<TD>>") == -1)
			{
				if(booLastWasNote)
					{
					if(cnt == 0)
						cnt = 1;
					else
						cnt = 0;
					} 
				if(cnt == 0)
					{
					if(strColor=="lightgreen" || strColor=="white")
						strColor = "#C2F5C2"
					oRow.style.backgroundColor = strColor; //"#C2F5C2";
					cnt++;
					}
				else
					{
					oRow.style.backgroundColor = "#FFFFE1";
					cnt = 0;
					}
				booLastWasNote = false;
			}
			else
			{
				if(cnt == 0)
					{
					if(strColor=="lightgreen" || strColor=="white")
						strColor = "#C2F5C2"
					oRow.style.backgroundColor = strColor; //"#C2F5C2";
					}
				else
					{
					oRow.style.backgroundColor = "#FFFFE1";
					}
				booLastWasNote = true;
			}
		oCell = oRow.insertCell();
		oCell.innerText = b[0]+":";
		oCell.vAlign = "top";
		oCell.width = "100px";
			
		oCell = oRow.insertCell();
		oCell.width = "516px";
		
		if (b[0].toUpperCase()=='WEBSITE' || b[0].toUpperCase()=='MENU') //  || b[0].toUpperCase()=='HOURS'
		{
				var tmp = b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"");
				tmp = '<a href="#" onClick="javascript:window.open(' + String.fromCharCode(39) + tmp + String.fromCharCode(39) + ')">' + 'Web' + '</a>';
				oCell.innerHTML = tmp;
		}
		else
			oCell.innerHTML = b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"");
			
			oCell.vAlign = "top"
			
			oCell = oRow.insertCell();
			
			if (b[0] != 'Name' && b[0] != 'Private Notes')
			{
				if (b[3]=='True') var strChecked = ' checked '
					else var strChecked = ' '
		
				var strHTML; // = '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
				strHTML = '<div style=width:18px;height:18px;overflow:hidden><input type="checkbox" onclick="dirty()" id="id' + b[2] + '"' + strChecked + '></div>';
			}
			else
			{
				var strHTML = '&nbsp;';
			
			}
			
			oCell.innerHTML = strHTML;
			
			
			oCell.vAlign = "middle";
			oCell.Align = "right";
			oCell.style.paddingTop = "0px";
			
		}
  }
	
  el.style.display = "inline";
  el.style.visibility = "visible";
	  
  //alert(document.frames.frameToolTip.document.all.divToolTip.scrollHeight+' - '+document.all.frameToolTip.height)
  if(document.frames.frameToolTip.document.all.divToolTip.scrollHeight < document.all.frameToolTip.height-48)
	document.frames.frameToolTip.document.all.pPrint.innerHTML = "Print&nbsp;";
  else
	document.frames.frameToolTip.document.all.pPrint.innerHTML = "Print&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";

  booToolTipOpen = true;
}


</script>
<%

set cn = server.CreateObject("adodb.connection")
cn.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")


 HighRes = (remote.Session ("AvailWidth") > 800)

 FieldArray = Array("LocationID","Alias","Name","Address","City","State","Zip","Phone","FaxNumber","Map","Miles","Menu","Website","Stars","Price","RecType","OTID","WebLinkType") ' "FaxNumber",

 SortArray = Array("Alias","City","Phone","Miles","Stars","Price")
 
 row = 1
 
 function StripAll(str)
		StripAll = Trim(Replace(Replace(Replace(Replace(Replace(Trim(str), "&", ""), ".", ""), "-", ""), Chr(39), ""), Chr(32), ""))
 End function

function oGetRecordset(cid, uid, maxRecs, page, dir, filter, sort, fieldArray)
	dim a, aFilter, aSort
	booFirstLoad = remote.session("FirstLoad")
	a = Array()
	aFilter = Split(Record_Filter,"|")
	aSort = Split(Sort,"|")
	
	if Record_Filter <> remote.session("rsLocFilter") then
		newCriteria = 1
		recnum = 1
	else
		newCriteria = 0
		recnum = remote.session("recnum")
		if dir = "p" then
			recnum = recnum-Max_Records
		elseif dir = "n" then
			recnum = recnum+Max_Records
			if remote.session("FirstLoad") = 1 then
				remote.session("FirstLoad") = 0
				newCriteria = 1
			end if
		else
			newCriteria = 1
			recnum = 1
		end if
	end if
	
	remote.session("recnum") = recnum
	TempFileName = Left("GCN_" & remote.session("CompanyID") & "_" & Trim(request.cookies("UserKey")),16)
	
	set r = server.CreateObject("adodb.recordset")
	

	if remote.session("FirstLoad") = "1" and aSort(0) = SortArray(remote.session("DefaultSortBy")-1) and aSort(1) = "Asc" and Request.QueryString("cat") = remote.Session("DefaultCategory") and Request.QueryString("scat") = 0 and aFilter(2) = "" and Len(KeyWordCriteria) = 0 and len(PointCriteria) = 0  and len(AreaCriteria) = 0  and remote.session("DefaultState") = Request.QueryString("state") then
	
		r.CursorLocation = 3
		r.Open "sp_GetLocationPage1 " & cid, cn
	else
		remote.session("FirstLoad") = 0
		set cmd = server.CreateObject("adodb.command")
		cmd.ActiveConnection = cn
     
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "sp_Vendors2"
		cmd.Parameters.Append cmd.CreateParameter("@CompanyID",adInteger,adParamInput,,cid)
		cmd.Parameters.Append cmd.CreateParameter("@AltOrder",adInteger,adParamInput,,recnum)
		cmd.Parameters.Append cmd.CreateParameter("@Rows",adInteger,adParamInput,,Max_Records)
		cmd.Parameters.Append cmd.CreateParameter("@FullCompanyName",adVarChar,adParamInput,50,left(aFilter(2),50))
		cmd.Parameters.Append cmd.CreateParameter("@BeginsOrContains",adChar,adParamInput,1,aFilter(3))
		cmd.Parameters.Append cmd.CreateParameter("@SortBy",adVarChar,adParamInput,24,aSort(0))
		cmd.Parameters.Append cmd.CreateParameter("@Order",adVarChar,adParamInput,10,aSort(1))
		cmd.Parameters.Append cmd.CreateParameter("@CategoryID",adInteger,adParamInput,,aFilter(0))
		cmd.Parameters.Append cmd.CreateParameter("@SubCategoryID",adInteger,adParamInput,,aFilter(1))
		cmd.Parameters.Append cmd.CreateParameter("@StateCriteria",adVarChar,adParamInput,2,aFilter(4))
		cmd.Parameters.Append cmd.CreateParameter("@NewCriteria",adBoolean,adParamInput,,newCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@TempTableName",adVarChar,adParamInput,16,TempFileName)
		cmd.Parameters.Append cmd.CreateParameter("@KeyWordCriteria",adVarChar,adParamInput,8000,KeyWordCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@PointsCriteria",adVarChar,adParamInput,8000,PointsCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@NameCriteria",adVarChar,adParamInput,8000,NameCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@AreaCriteria",adVarChar,adParamInput,8000,AreaCriteria)
	
		set r = cmd.Execute()
	end if
	
	if not r.EOF then
		if Max_Records < r.Fields("lastRec").Value then
			RecordCount	= Max_Records
		else
			RecordCount	= r.Fields("lastRec").Value
		end if
		redim a(Ubound(FieldArray)+1,RecordCount)
		LastRec = r.Fields("lastRec").Value
		BOF = (r.BOF and RecordCount=0)

		for i = 0 to Max_Records-1
			for j = 0 to Ubound(FieldArray)
				a(j,i) = r.Fields(FieldArray(j)).Value
			next
			r.MoveNext
			if r.EOF then
				RecordCount = i+1
				exit for
			end if
		next
		if i < Max_Records-1 or recnum+Max_Records > LastRec then
			EOF = true
		end if
	end if
	
	set cmd = nothing
	r.Close
	set r = nothing
	
	oGetRecordset = a
end function
 
function rs(field)
 
	If vartype(field) <> 8 Then 
		index = field 
	Else
		for index = LBound(FieldArray) to UBound(FieldArray)
			if Ucase(FieldArray(index)) = UCase(field) Then
				Exit For
			End If
		Next
	End If
 
	 rs = locations(index,row)
   
 end function

dim txtRecordset
dim txtTextField
dim txtTextIndex
dim txtFilter   
dim fcount
dim booEOF, EOF, BOF

booEOF = false

txtFilter    = Left(Trim(Request.QueryString("Filter")),4)
txtFilter2 = StripAll(txtFilter)
txtFullFilter    = Trim(Request.QueryString("Filter"))

txtCategory = Trim(Request.QueryString("cat"))
txtSubCategory = Trim(Request.QueryString("scat"))
txtSearchType = Trim(Request.QueryString("opt"))
txtState = Trim(Request.QueryString("state"))

Sort = Trim(Request.QueryString("sort"))
SortDir = Trim(Request.QueryString("dir"))

Sort = Sort & "|" & SortDir
page = Trim(Request.QueryString("page"))

'Response.Write txtCategory & ", " & txtSubCategory & ", " & txtSearchType & ", " & Sort & ", " & SortDir & ", " & Page

fcount = 0
forceCreate = false

	PointsCriteria = "0"
	
	
	if Cint(Request.QueryString("area")) > 0 Then 
		txtSearchType = 3
		AreaID = Cint(Request.QueryString("area"))
	Else
		AreaID = 0 
	End If


	select case page
		case "b"
			Record_Filter = remote.Session ("rsLocFilter") 
		    Sort = remote.Session ("rsLocSort") 
		    MoveDir = "p"
		    
		case "n"
			Record_Filter = remote.Session ("rsLocFilter") 
        	Sort = remote.Session ("rsLocSort") 
        	MoveDir = "n"
		case else
			Record_Filter = Record_Filter & txtCategory & "|" & txtSubCategory & "|"
			Record_Filter = Record_Filter & txtFullFilter & "|"
			If txtSearchType = 1 Then 
				boc = "b"
			End If
			
			If txtSearchType = 2 Then
				boc = "c"
			end if
			
			if Request.QueryString("area") > 0 Then 
				txtSearchType = 3
				AreaID = Request.QueryString("area")
			Else
				AreaID = 0 
			End If
			
			
			If txtSearchType = 3 Then 
				boc = "k"
				
				
				arrKeywords = Split(txtFullFilter,chr(32))
				
				set rsKey = Server.CreateObject ("Adodb.recordset")
				rsKey.CursorLocation = 3
				

				Points = 10
				KeyWordCriteria = ""
												

				NameCriteria = ""
				
				' On Error Resume Next
				
				for i = LBound(arrKeyWords) to UBound(arrKeyWords)

						NameCriteria = NameCriteria & " CharIndex (' " & replace(Trim(arrKeyWords(i)),"'","''") & " ', v.CompanyName) > 0 "
						
						For j = i to UBound(arrKeyWords)
						
						
							tmpKeyWord = ""
							for k = i to j
								tmpKeyWord = tmpKeyWord & replace(Trim(arrKeyWords(k)),"'","''''") & " "
							next
							
							If Len(tmpKeyword)  > 2 Then
								On Error Resume Next	' To take care of "ignroed words"
								rsKey.Open "sp_KeyWord_Find '" & tmpKeyWord & "'", cn
								
								if err = 0 then 's	
									If not rsKey.EOF and not rsKey.BOF Then
							
										Do while not rsKey.EOF 
											If len (KeyWordCriteria) > 0 Then
												KeyWordCriteria = KeyWordCriteria & " or "
											End If
											
											If len (PointsCriteria) > 0 Then
												PointsCriteria = PointsCriteria & " + "
											End If
							
												KeyWordCriteria = KeyWordCriteria & "charindex('|" & rsKey(0) & "|', keywords) > 0"
												PointsCriteria = PointsCriteria & "case charindex('|" & rsKey(0) & "|', lk.keywords) when 0 then 0 else " & (6-rsKey("ktype")) & " end "

												rsKey.MoveNext 
	
										Loop
												
									End If
									
									rsKey.Close ()
								end if
							End If
							
						next
					
				Next

				
				'Response.Write tmpKeyWord	& i & " " & j & " " & k & "<br>"
				if AreaID > 0 Then 			
						
						If  len (PointsCriteria) > 0 Then
							PointsCriteria = PointsCriteria & " + "
						End If
						
						AreaCriteria = "charindex('|a" & AreaID & "|', keywords) > 0"
						
						PointsCriteria = PointsCriteria & "case charindex('|a" & AreaID & "|', lk.keywords) when 0 then 0 else 7 end "
						
				End If
				
				if Len(PointsCriteria) = 0 Then
					PointsCriteria = "1"
				End If
				
				'Deb KeyWordCriteria & "<br>","1"
				'Deb PointsCriteria  & "<br>","2"
				'deb NameCriteria  & "<br>","3"
				
				'Response.End 
				
			end if
			
			Record_Filter = Record_Filter & boc & "|" & txtState
			remote.Session ("rsLocFilter") = Record_Filter
        	remote.Session ("rsLocSort") = Sort
	end select
	
	locations = oGetRecordset (cid, 0, Max_Records , 1, MoveDir, Record_Filter, Sort, FieldArray)' o.GetRecordset(32, "ilia", 20, 1, "", "", "", Array(1, 2, 3, 4, 5))

if page="" then
	oFilter    = Trim(Request.QueryString("Filter"))
	oFilter2 = StripAll(oFilter)
	remote.Session ("oFilter2") = oFilter2
Else
	oFilter2 = remote.Session ("oFilter2")
End If

'Response.Write boc = "b" and len(oFilter) > 0 'and len (oFilter2) > Len(txtFilter2) and Not EOF and page="")
If boc = "b" and len(oFilter) > 0 and len (oFilter2) > Len(txtFilter2) and Not EOF and page="" Then 
	i = 0
	
	Found = false
	Do while (i < RecordCount ) and not Found

		row = i
		
		If Ucase(Trim(oFilter2)) = Ucase(Left(StripAll(rs("Alias")),len(oFilter2))) Then
			Found = True
			'exit do
		End If 
		i = i + 1
		
		If i = RecordCount and not eof Then
		
			page="n"
			locations = oGetRecordset (cid, 0, Max_Records , 1, "n", Record_Filter, Sort, FieldArray)

			'If RecordCount > 0 Then
			'	RecordCount = UBound(locations,2) + 1
			'Else
			'	RecordCount = 0
			'End If
			'if EOF then
			'	RecordCount = 0
			'end if
			i = 0
		End If
	Loop
	
End If
	
if RecordCount > 0 then
	Response.Write ("<Center><TABLE border=0 cellspacing=0 cellpadding=1 style=""font-family:Tahoma;font-size:11px"">")
	'Response.Write "<TR style=""background-color:silver;border-style:solid""><TD style=""width:20"">&nbsp;&nbsp;&nbsp;&nbsp;</td><TD>Name</TD><tD>Address</TD><tD>City</td><td>Phone</td><td>Map</td><td>Miles</Td><td>Menu</td><td>Web</td><td>Stars</td><td>Price</td></tr>"
	i = 0

	Response.Write ("<TR style=""height:21px"">" & vbcrlf)
	Response.Write "<TD class=""bord"" style=""width:" & intNameWidth & "px;border-style:outset;border-width:1px;"">Name</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intAddressWidth & "px;border-style:outset;border-width:1px;"">Address</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intCityWidth & "px;border-style:outset;border-width:1px;"">City</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intPhoneWidth & "px;border-style:outset;border-width:1px;"">Phone</td>"
	if intFaxWidth > 0 then
		'Response.Write "<TD class=""bord"" style=""width:" & intFaxWidth & "px;border-style:outset;border-width:1px;"">Fax</td>"
		Response.Write "<TD class=""bord"" style=""width:" & intFaxWidth & ";border-style:outset;border-width:1px;"">Phone 2</td>"
	end if
	
	'If HighRes Then
	'	Response.Write "<TD class=""bord"" style=""width:70;border-style:outset;border-width:1px;"">Fax</td>"
	'End If
	
	'Response.Write "<TD class=""bord"" style=""width:22;border-style:outset;border-width:1px;"" align=center>Map</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intMilesWidth & "px;border-style:outset;border-width:1px;"" align=center>Miles</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intWebWidth & "px;border-style:outset;border-width:1px;"" align=center>Web</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intWeb2Width & "px;border-style:outset;border-width:1px;"" align=center>Web 2</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intStarsWidth & "px;border-style:outset;border-width:1px;"">Stars</td>"
	'Response.Write "<TD class=""bord"" style=""width:30;border-style:outset;border-width:1px;"">Price</td>"
	Response.Write "<TD class=""bord"" style=""width:20;border-style:outset;border-width:1px;"">OT</td>"
	Response.Write "</TR>"
	lastid = 0
	aCount = 0
	
	minusRec = 1
	If remote.session("FirstLoad") = 1 then
		Response.Write "<tr><td class=bord align=middle colspan=15><B>Frequently Used Vendors</B></td></tr>"
		minusRec = 2
	End If
	
	Do while  i <= RecordCount - minusRec
				
				row = i
				rslid = rs("LocationID")
				lastid = rslid

				if strBColor = "EDFAE7" then
					strBColor = "FFFFFF"
				else
					strBColor = "EDFAE7"
				end if
				
				If Ucase(Trim(oFilter2)) = Ucase(Left(StripAll(rs("Alias")),len(Trim(oFilter2)))) and Len(oFilter2) > 4 Then
					'Response.End 
						strBColor = "A0FAE7"
				End If
				
				
				raid = rslid
				cnt = i
				' mosel(" & raid & "," & cnt & ");parent.returnValues();
				'
				
				Response.Write "<span language=""javascript1.2"" ondblclick=""parent.returnValues()"" onmousedown=""mosel(" & raid & "," & cnt & ")"" onmouseout=""mout(" & raid & "," & cnt & ")"" onmouseover=""mo(" & raid & "," & cnt & ")"">"
				Response.Write (vbcrlf & "<TR bgcolor=#" & strBColor & " id=tr" & cnt & ">")

				' ID
				'Response.Write "<TD class=""bord"" style=""height:21;background-color:#317142""><Img width=15 height=15 name=""" & "n" & rslid & """ id=""" & "r" & i & """ onmousedown=""optCheck(this)"" src=""images/OptionUnChecked.jpg""></td>" & vbcrlf
				' Name
				'onmouseover=""setStatus('" & replace(Trim(rs("Name")),"'","<<sq>>") & "');return true;"" onmouseout=""window.status=''""
				Response.Write "<TD class=""bord""><div title=""" & Trim(rs("Alias")) & """ style=""overflow:hidden;width:" & intNameWidth-4 & ";height:15""><a href=""javascript:toolTip(" & raid & "," & cnt & ")"">" &  Trim(rs("Alias")) & "</a></div></td>"
				' Address
				Response.Write "<TD class=""bord"" nowrap><div title=""" & Trim(rs("Address")) & """ style=""overflow:hidden;width:" & intAddressWidth & "px;height:15"">" &  Trim(rs("Address")) & "</div></td>"
				'City
				Response.Write "<TD class=""bord"" nowrap><div style=""overflow:hidden;width:" & intCityWidth & ";height:15"">" &  Trim(rs("City")) & "</div></td>"
				
				'State does not show
				Response.Write "<TD style=""display:none"" class=""bord"" nowrap><div>" &  Trim(rs("state")) & "</div></td>"
				
				'Zip does not show
				Response.Write "<TD style=""display:none"" class=""bord"" nowrap><div>" &  Trim(rs("zip")) & "</div></td>"
				'Phone
				If Trim(rs("Phone")) <> "" Then Phone = rs("phone") Else Phone = " "
				Response.Write "<TD nowrap class=""bord""><div title=""" & Phone & """ style=""overflow:hidden;width:" & intPhoneWidth & "px;height:15"">" & Phone & "</div></td>"
			
				'Fax
				if intFaxWidth > 0 then
					If Trim(rs("FaxNumber")) <> "" Then FaxNumber = rs("FaxNumber") Else FaxNumber = " "
					Response.Write "<TD class=""bord""><div title=""" & trim(FaxNumber) & """ style=""overflow:hidden;width:" & intFaxWidth & ";height:15"">" &  FaxNumber & "</div></td>"
				end if
				
				'If HighRes Then
				'	If Trim(rs("FaxNumber")) <> "" Then FaxNumber = "(" & Left(rs("FaxNumber"),3) & ") " & Mid(rs("FaxNumber"),4,3) & "-" & Right(rs("FaxNumber"),4) Else FaxNumber = " "
				'	Response.Write "<TD class=""bord""><div style=""overflow:hidden;width=90;height:15"">" &  FaxNumber & "</div></td>"
				'End If
	
				' Map
				'If rs("Map") Then Map = "Map" Else Map = ""
				'Response.Write "<TD align=center class=""bord"">" & Map & "</td>"
				'Miles
				if n2z(rs("Miles")) = 0 then
					strMiles = ""
				else
					strMiles = formatnumber(rs("Miles"),1)
				end if
				Response.Write "<TD nowrap class=""bord"" align=right><div style=""overflow:hidden;width:" & intMilesWidth & ";height:15"">" & strMiles & "</div></td>"
				'' Menu
				'If len(trim(rs("Menu"))) > 0 Then
				'	Menu = "Menu"
				'else
				'	Menu = ""
				'end if
				
				'Website
				If Trim(rs("Website")) <> "" Then Website = "<a href=javascript:wopen('" & rs("Website") & "')>Web</a>" Else Website = " "
				Response.Write "<TD align=center class=""bord"">" &  Website & "</td>"
		
				if isnull(rs("Stars")) then
					str = ""
				else
					str = replace(rs("Stars")," ","")
				end if
		
				' WebLinkType
				If len(trim(rs("Menu"))) > 0 Then
					m = "<a href=javascript:wopen('" & rs("Menu") & "')>"
					if len(trim(rs("WebLinkType"))) > 0 then
						Menu = rs("WebLinkType")
					else
						Menu = "Menu"
					end if
					Menu = m & Menu & "</a>"
				else
					Menu = ""
				end if

				if str = "No" then
					strStars = "<img id=navSadFace name=navSadFace src=""images/Avoid.gif"">"
					str = ""
				else
					strStars = str
				end if

				Response.Write "<TD nowrap align=center class=""bord""><div style=""overflow:hidden;width:" & intStarsWidth & ";height:15"">" & Menu & "</div></td>"
				Response.Write "<TD align=""center"" title=""" & len(str) & " Stars" & """style=""color: #B16918"" class=""bord"">" &  strStars & " " & "</td>" 'Trim(rs("Stars"))

				str = ""
				str2 = ""
			
				'Response.Write "<TD class=""bord""><a title=""" & str2 & """>" & str & "</a></td>"
				
				
				if rs("OtID") > 0 Then
					Response.Write "<TD bgColor='red' align=middle class=""bord""><div style=""overflow:hidden;width=20;height:15;backgroundColor:red;"">OT</div></td>"
					Response.Write "<TD style=""display:none""><div>" & rs("OTid") & "</div></td>"
				Else
					Response.Write "<TD class=""bord""><div style=""overflow:hidden;width=20;height:15;"">&nbsp;</div></td>"
					Response.Write "<TD style=""display:none""><div></div></td>"
				End If


				Response.Write "</TR>"
				Response.Write "</span>"
				
				i = i + 1
		
	Loop	
	
	'booBack = (page <> "" and not BOF)
	booBack = (recnum > Max_Records)
	
	if RecordCount > 0 then
		if LastRec > Max_Records then
			Response.Write ("<TR bgcolor=#" & "EDFAE7>" & vbcrlf)

			Response.Write "<TD align=""center"" colspan=""11"" class=""bord"" style=""height:21px;width:144;font-weight:bold"">" & vbcrlf
			'Response.Write "<table cellspacing=""0"" style=""height:21px;width:360;font-weight:bold;font-size:11px;""><tr><td align=""center"" width=""80px"">"
			Response.Write "<table cellspacing=""0"" style=""height:21px;width:160;font-weight:bold;font-size:11px;""><tr><td align=""center"" width=""80px"">"
			if booBack then
				Response.Write "<span style=cursor:hand onclick=""parent.disableSearchFields();window.location='BrowseDetail2.asp?page=b'"" onmouseout=""this.style.borderWidth='0px'"" onmousedown=""this.style.borderStyle='inset'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:red;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_previous.gif><td>&nbsp;Prev&nbsp;&nbsp;</td></td></tr></table>"
				Response.Write "</span>" & vbcrlf
			else
				Response.Write "<span>" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:silver;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_previous.gif><td>&nbsp;Prev&nbsp;&nbsp;</td></td></tr></table>"
				Response.Write "</span>" & vbcrlf
			end if
			'
			Response.Write "</td><td valign=""middle"" align=""center"" width=""80px"">"
			if LastRec <= Max_Records or eof or recnum+Max_Records >= LastRec then
				Response.Write "<span>" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:silver;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;Next</td><td>&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_next.gif>&nbsp;&nbsp;</td></tr></table>"
				Response.Write "</span>" & vbcrlf
				'Response.Write "&nbsp;&nbsp;Next&nbsp;>>&nbsp;&nbsp;"
			else
				Response.Write "<span style=cursor:hand onclick=""parent.disableSearchFields();window.location='BrowseDetail2.asp?page=n'"" onmouseout=""this.style.borderWidth='0px'"" onmousedown=""this.style.borderStyle='inset'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:red;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;Next</td><td>&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_next.gif>&nbsp;&nbsp;</td></tr></table>"
				Response.Write "</span>" & vbcrlf
			end if
			'Response.Write "</td><td width=170px align=right>"
			'	Response.Write "<span onclick=""alert('test')"" onmouseout=""this.style.borderWidth='0px'"" onmousedown=""this.style.borderStyle='inset'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
			'	Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:red;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;Add New Vendor</td><td>&nbsp;<img id=nav_new name=nav_new border=0 src=images/nav_new.gif>&nbsp;&nbsp;</td></tr></table>"
			'	Response.Write "</span>" & vbcrlf
			Response.Write "</td></tr></table>"
			Response.Write "</td>"
			'WriteCols
		end if
	else
		if EOF and BOF then
			Response.Write "<TD colspan=11 class=""bord"" style=""color:red;width:160"">"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No Locations Found"
			Response.Write "</td>"
			'WriteCols
		end if
	end if

	Response.Write "</Center></table>"
else
	strScript = strScript & "document.all(""bdy"").background=""images/NoRecordsFound_Gray.jpg""" & vbcrlf
	strScript = strScript & "document.all(""bdy"").style.backgroundRepeat=""no-repeat""" & vbcrlf
	strScript = strScript & "document.all(""bdy"").style.backgroundPosition=""center""" & vbcrlf
end if

strScript = strScript & "parent.enableSearchFields();" & vbcrlf
strScript = strScript & "if(parent.intKeyCode == 9) parent.document.all(""Category"").focus(); else parent.document.all(""txtSearch"").focus();" & vbcrlf

Response.Write "<script>" & vbcrlf
Response.Write strScript
Response.Write "</script>" & vbcrlf

if (len(txtFilter) and RecordCount = 0) Then
	strScript = "parent.addNewVendor(true);" & vbcrlf
	Response.Write "<script>" & vbcrlf
	Response.Write strScript
	Response.Write "</script>" & vbcrlf
End If


	

function n2b( s )
	if len(trim(s)) > 0 then
		n2b = s
	else
		n2b = "&nbsp;"
	end if
end function

function n2z(var)
	dim retval
	if isnull(var) then
		retval = 0
	else
		retval = var
	end if
	n2z = retval
end function

function z2n(var)
	dim retval
	if var = 0 then
		retval = ""
	else
		retval = var
	end if
	z2n = retval
end function


Function AutoFormat(strString, strKey, strReplacer)
    Dim intKeyPos
		    
    If IsNull(strString) Then
        strString = ""
    End If
		    
    intKeyPos = InStr(1, strString, strKey)
    Do Until intKeyPos = 0
        strString = Mid(strString, 1, intKeyPos - 1) & strReplacer & Mid(strString, intKeyPos + Len(strKey))
        intKeyPos = InStr(intKeyPos + Len(strReplacer), strString, strKey)
    Loop
		    
    AutoFormat = strString
End Function

%>
<script language="javascript1.2" event="onload" for="window">
	parent.tr = document.body.createTextRange();
	parent.enableSearchFields();
</script>

</BODY>
</HTML>
<%
if booFirstLoad then
	Response.Flush
	url = Application("HomePage") & "/CheckRebuildPage1Cache.asp?cid=" & cid
	set xmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
	xmlhttp.open "GET", url, true
	xmlhttp.send ""
	set xmlhttp = nothing
end if
%>
