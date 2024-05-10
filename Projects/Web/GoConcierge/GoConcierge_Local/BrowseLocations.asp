<%@ Language=VBScript %>

<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!--#include file = "Global.asp" ---> 
<%
Set remote = Server.CreateObject ("UserClient.MySession")
remote.Init (Request.Cookies("UserKey"))
dim cid, booFirstLoad
cid = remote.Session ("CompanyID")

'---
'Response.Write ("--firstload:  " & remote.session("FirstLoad") & "--<br>")
'Response.Write ("--defaultstate:  " & remote.session("DefaultState") & "--<br>")      
'Response.Write ("--DefaultCategory:  " & remote.session("DefaultCategory") & "--<br>") 
'Response.Write ("--rsLocFilter:  " & remote.session("rsLocFilter") & "--<br>")
'Response.Write ("--page:  " & page & "--<br>")  
'---

If remote.Session("ScreenHeight") < 750 Then
	Max_Records		= 16 
	ToolTipHeight	= 378
	pTop			= 2
	pLeft			= 60

	intNameWidth	= 164
	intAddressWidth	= 140
	intCityWidth	= 88
	intPhoneWidth	= 117
	intFaxWidth		= 0
	intMapWidth		= 22
	intMilesWidth	= 40
	intWebWidth		= 30
	intWeb2Width	= 30
	intStarsWidth	= 30
Else
	Max_Records = 22 
	ToolTipHeight = 502
	pTop = 3
	pLeft			= 160

	intNameWidth	= 200
	intAddressWidth	= 180
	intCityWidth	= 107
	intPhoneWidth	= 120
	intFaxWidth		= 108
	intMapWidth		= 22
	intMilesWidth	= 40
	intWebWidth		= 30
	intWeb2Width	= 30
	intStarsWidth	= 30
End If

' just for live push...
'if remote.Session("DefaultCategory") = "" then
'	remote.Session("DefaultCategory") = 0
'	remote.Session("DefaultBCK") = 1
'	remote.Session("DefaultSortBy") = 1
'end if

set cn = server.CreateObject("adodb.connection")
cn.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")

'
%>
<HTML>
<HEAD>
<STYLE>
	<!-- #D8BFD8 -->
	.But {font-family:Tahoma;font-size:11px; }
	.bord {border-bottom-width:1pt;border-bottom-style:solid;border-bottom-color:#D8BFD8;border-right-width:1pt;border-right-style:solid;border-right-color:#D8BFD8;}
	A:hover		{color:red}
	A:active	{color:blue}
	A:visited	{color:blue}
</style>


<script type="text/javascript" language="javascript1.2" defer>

var lastColor = "";
var SelectedBGColor = "";
var SelectedID = 0;
var Red = "#ff9d9d"

function mo( t, i )		{ HoverOn(t, i);  }
function mout( t, i )	{ HoverOff(t, i); }
function mosel( t, i )	{ Selected(t, i); }

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



function Selected( t, i )
{
	var str = t;
	SelectedBGColor = lastColor;

	restoreRowColor()
		
	document.all("tr"+i).style.backgroundColor = Red;
	SelectedID = i;
	
	showTextActivate(t);
	
	/* good way to see the resulting html
	var tr = window.bdy.createTextRange();
	tr.expand("textedit");
	window.clipboardData.setData("Text",tr.htmlText);
	*/
}

function showTextActivate(strField) { //, x, y) {
	
	
	window.frames("frameToolTip").document.all("toolTipID").value = strField;
	
	//var xmlHttp = new ActiveXObject("Microsoft.XMLHTTP")
	
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
	// work around for weird returning of escaped less than/greater than signs
	b[1] = b[1].replace(/\&lt;<TD>\&gt;/g,"<<TD>>").replace(/\&lt;<SQ>\&gt;/g,"<<SQ>>")
	if(b[1].length > 0)
		{
		oRow = frel.insertRow();
			if(b[1].indexOf("<<td>>") == -1)
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
		
		if (b[0].toUpperCase()=='WEBSITE' || b[0].toUpperCase()=='MENU' && b[1].length > 0) //  || b[0].toUpperCase()=='HOURS'
		{
				var tmp = b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"").replace(' ','');
				tmp = '<a href="#" onClick="javascript:window.open(' + String.fromCharCode(39) + tmp + String.fromCharCode(39) + ')">' + 'Web' + '</a>';
				oCell.innerHTML = tmp;
		}
		
		else if ( b[1].indexOf("http://") ==0 || b[1].indexOf("https://") ==0) // if field begins w/ http: or https: make it a link
		{
			var tmp = b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"").replace(' ','');
				tmp = '<a href="#" onClick="javascript:window.open(' + String.fromCharCode(39) + tmp + String.fromCharCode(39) + ')">' + 'Web' + '</a>';
				oCell.innerHTML = tmp;
		}
		
		else
			if (b[0].toUpperCase()=='E-MAIL') //  || b[0].toUpperCase()=='HOURS'
			{
					var tmp = b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"");
					tmp = '<a onmousedown="window.event.cancelBubble=true;" hhref="#" href="javascript:parent.wopen(' + String.fromCharCode(39) + tmp + String.fromCharCode(39) + ')">' + b[1].replace(/<<SQ>>/g,"'").replace(/<<TD>>/g,"").replace(/mailto:/g,"") + '</a>';
					oCell.innerHTML = tmp;
			}
			//Menu = "<A onmousedown=""window.event.cancelBubble=true;"" href=""javascript:wopen('" & Trim(rs("menu")) & "')"">" & rs("WebLinkType") & "</a>"
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


var imgOptionUnChecked = new Image();
var imgOptionChecked = new Image();
//var aSelected = parent.document.all("txtSelected");
// hello

imgOptionUnChecked.src = "images/OptionUnChecked.jpg";
imgOptionChecked.src = "images/OptionChecked.jpg";



function optCheck (opt)
{
	var id, name, x, booAddRemove, alen;
	var idc;
	
	//name = opt.id;
	name = opt.name;
	id = name.slice(1);
	
	if(parent.document.all("txtSelected").value.indexOf(id) > -1)
		{
		idc = id+",";
		parent.document.all("txtSelected").value = parent.document.all("txtSelected").value.replace(idc,"");
		opt.src = imgOptionUnChecked.src;
		if(parent.mode == "Edit")
			fillPageOpts();
		}
	else
		{
		if(parent.mode == "Edit")
			parent.document.all("txtSelected").value = "";
		parent.document.all("txtSelected").value += id+","
		opt.src = imgOptionChecked.src;
		//alert(parent.document.all("txtSelected").value)
		if(parent.mode == "Edit")
			fillPageOpts();
		}
		
		
	//if # of vendors is selected is more than 30, don't allow them to print, email or view	
	selectedArray = parent.document.all("txtSelected").value.split(",");
	if (selectedArray.length > 11)
	{
		parent.document.all("cmdViewLocations").disabled = true;
		parent.document.all("cmdPrintLocations").disabled = true;
		parent.document.all("cmdEmailLocations").disabled = true;
	}
	else
	{
		parent.document.all("cmdViewLocations").disabled = false;
		parent.document.all("cmdPrintLocations").disabled = false;
		parent.document.all("cmdEmailLocations").disabled = false;
	}

}

function window_onload()
{
	fillPageOpts();
}

function fillPageOpts()
{
	 var images = document.all.tags("IMG");
	 var id, isChecked, rid, nrid, nid;
	 
	 for(var i=0; i<images.length; i++)
		{
			if(images[i].id.substr(0,3) != 'nav')
			{	
				rid = images[i].id;
				nrid = images[i].name;
				id = rid.slice(1);
				nid = nrid.slice(1);
				isChecked = parent.document.all("txtSelected").value.indexOf(nid);
				if(isChecked > -1)
					document.all(rid).src = imgOptionChecked.src;
				else
					document.all(rid).src = imgOptionUnChecked.src;
			}
		}
}

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
</SCRIPT>

<%
dim RecordCount, LastRec, recnum
RecordCount = 0
LastRec = 0
recnum = 0
' Response.Flush()


if Sort = "Miles" then
	'if SortDir <> "Desc" then
	'	Sort = "SortMiles"
	'end if
end if

 FieldArray = Array("LocationID","Alias","Name","Address","City","Phone","FaxNumber","Map","Miles","Menu","Website","Stars","Price","RecType","OTID","WebLinkType")
 SortArray = Array("Alias","City","Phone","Miles","Stars","Price")

 row = 1
 
 function StripAll(str)
		if isnull(str) then
			StripAll = ""
		else
			StripAll = Trim(Replace(Replace(Replace(Replace(Replace(Trim(str), "&", ""), ".", ""), "-", ""), Chr(39), ""), Chr(32), ""))
		end if
 End function
 
 GetRecordCount = ""

 
 
function oGetRecordset(cid, uid, maxRecs, page, dir, filter, sort, fieldArray)
	dim a, aFilter, aSort
	booFirstLoad = remote.session("FirstLoad")
	a = Array()
	aFilter = Split(Record_Filter,"|")
	'Response.Write record_filter
	'Response.End 
	
	aSort = Split(Sort,"|")
	
	'Response.Write Record_Filter & "=?" & remote.session("rsLocFilter") & "-" & Len(KeyWordCriteria) & "-" & len(PointCriteria) & "-" & dir & "<br><br>"
	if Record_Filter <> remote.session("rsLocFilter") or Len(KeyWordCriteria) > 0 or len(PointCriteria) > 0  then
		newCriteria = 1
		recnum = 1
		
	else
		newCriteria = 0
		recnum = remote.session("recnum")
		if dir = "p" then
			recnum = recnum-Max_Records
		elseif dir = "n" then
			if remote.session("FirstLoad") = 1 then
				remote.session("FirstLoad") = 0
				newCriteria = 1
			end if
			recnum = recnum+Max_Records
		else
			newCriteria = 1
			recnum = 1
		end if
	end if
	'Response.Write newCriteria
	remote.session("recnum") = recnum
		TempFileName = Left("GCN_" & remote.session("CompanyID") & "_" & Trim(request.cookies("UserKey")),16)

	set r = server.CreateObject("adodb.recordset")
	set cn2 = server.CreateObject("adodb.connection")
	cn2.Open Application("sqlInnSight_ConnectionString"), Application("sqlInnSight_RuntimeUsername"), Application("sqlInnSight_RuntimePassword")



'response.write "FirstLoad:  " & remote.session("FirstLoad")& "=?1<br>"
'response.write "DefaultSortBy:  " & aSort(0)& "=?" & SortArray(remote.session("DefaultSortBy")-1) & "<br>"
'response.write "aSort(1)=?ASC --> " & aSort(1)& "=?ASC<br>"
'response.write "Request.QueryString(""cat"")=?remote.Session(""DefaultCategory"") --> " & Request.QueryString("cat")& "=?" & remote.Session("DefaultCategory") & "<br>"
'response.write "Request.QueryString(""scat"")=?0 --> " & Request.QueryString("scat")& "=?0<br>"
'response.write "aFilter(2)=?"""" --> " & aFilter(2) & "=?""""<br>"
'response.write "len(KeyWordCriteria)=?0 --> " & Len(KeyWordCriteria) & "=?0<br>"
'response.write "len(PointCriteria)=?0 --> " & len(PointCriteria) & "=?0<br>"
'response.write "len(AreaCriteria)=?0 --> " & len(AreaCriteria) & "=?0<br>" 
'Response.Write "DefaultState=?State --> " & remote.session("DefaultState") & "=?" & Request.QueryString("state")	& "<br>"
	
	'Response.Write cint(Request.QueryString("cat")) = DefaultCategory 'remote.session("FirstLoad") & ", " & Request.QueryString("cat") & ", " & DefaultCategory & ", " & Request.QueryString("scat") & ", " & aFilter(2)
	if (remote.session("FirstLoad") = "1") and (aSort(0) = SortArray(remote.session("DefaultSortBy")-1)) and (aSort(1) = "Asc") and (Request.QueryString("cat") = remote.Session("DefaultCategory")) and (Request.QueryString("scat") = 0) and (aFilter(2) = "") and (Len(KeyWordCriteria) = 0) and len((PointCriteria) = 0) and (len(AreaCriteria) = 0) and (remote.session("DefaultState") = Request.QueryString("state") OR remote.session("DefaultState") = "") then
		'Response.Write ("<br>*****Running sp_GetLocationPage1*****<br><br>")
		r.CursorLocation = 3
		r.Open "sp_GetLocationPage1 " & cid, cn2
		
		set rscount = cn2.Execute ("Select count(*) from tblCompanyLocation where companyID=" & remote.Session ("CompanyID"))
		strCount = rscount(0)
		
		remote.Session ("GetRecordCount") = strCount
		
	else
		Response.Write ("<br>*****Running sp_Vendors2*****<br><br>")
		remote.session("FirstLoad") = "0"
		set cmd = server.CreateObject("adodb.command")
		
		cmd.ActiveConnection = cn2
		
		catid = aFilter(0)
		scatid = aFilter(1)
		dstate = aFilter(4)
		
		'Response.Write "<br>dstate:  " & dstate & "<br>"
		if catid = "" then
			catid = 0
		else
			catid = cint(catid)
		end if
		
		if scatid = "" then
			scatid = 0
		else
			scatid = cint(scatid)
		end if
		
		if dstate = "" then
			dstate = ""
		end if
		'Response.Write "<br>*****sp_vendors_2****<br>"
		'Response.Write ("catid:  " & catid & "--<br>")  
		
		cmd.CommandType = adCmdStoredProc
		cmd.CommandText = "sp_Vendors2"
		cmd.Parameters.Append cmd.CreateParameter("@CompanyID",adInteger,adParamInput,,cid)
		cmd.Parameters.Append cmd.CreateParameter("@AltOrder",adInteger,adParamInput,,recnum)
		cmd.Parameters.Append cmd.CreateParameter("@Rows",adInteger,adParamInput,,Max_Records)
		cmd.Parameters.Append cmd.CreateParameter("@FullCompanyName",adVarChar,adParamInput,50,left(aFilter(2),50))
		cmd.Parameters.Append cmd.CreateParameter("@BeginsOrContains",adChar,adParamInput,1,aFilter(3))
		cmd.Parameters.Append cmd.CreateParameter("@SortBy",adVarChar,adParamInput,24,aSort(0))
		cmd.Parameters.Append cmd.CreateParameter("@Order",adVarChar,adParamInput,10,aSort(1))
		cmd.Parameters.Append cmd.CreateParameter("@CategoryID",adInteger,adParamInput,,catid)
		cmd.Parameters.Append cmd.CreateParameter("@SubCategoryID",adInteger,adParamInput,,scatid)
		cmd.Parameters.Append cmd.CreateParameter("@NewCriteria",adBoolean,adParamInput,,newCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@TempTableName",adVarChar,adParamInput,16,TempFileName)
		cmd.Parameters.Append cmd.CreateParameter("@KeyWordCriteria",adVarChar,adParamInput,8000,KeyWordCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@PointsCriteria",adVarChar,adParamInput,8000,PointsCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@NameCriteria",adVarChar,adParamInput,8000,NameCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@AreaCriteria",adVarChar,adParamInput,8000,AreaCriteria)
		cmd.Parameters.Append cmd.CreateParameter("@StateCriteria",adVarChar,adParamInput,2,dstate)
	
		set r = cmd.Execute()
		
		if dir <> "p" and dir <> "n" Then
			set rscount = cn2.Execute ("Select count(*) from LocationCache.innsight_user." & TempFileName ) 
			strCount = rscount(0)
			remote.Session ("GetRecordCount") = strCount
		End If
		
		
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
	set cn = nothing
	
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
dim booEOF
dim EOF, BOF, booBack	

booEOF = false

rawSearchText = Request.QueryString("sText")


if Len (rawSearchText) > 4 Then
	st = lcase(Replace(Ucase(rawSearchText),"THE ",""))
Else
	st = rawSearchText
End If


txtFilter  = Left(Trim(st),4)
txtFilter2 = StripAll(txtFilter)
txtFullFilter = Trim(st)

txtCategory = Trim(Request.QueryString("cat"))
txtSubCategory = Trim(Request.QueryString("scat"))
txtSearchType = Trim(Request.QueryString("opt"))
txtState = Trim(Request.QueryString("state"))
if txtSearchType = "" then
	txtSearchType = 1
end if

Sort = Trim(Request.QueryString("sort"))
SortDir = Trim(Request.QueryString("dir"))

Sort = Sort & "|" & SortDir

page = Trim(Request.QueryString("page"))

'Response.Write "<br>" & txtCategory & ", " & txtSubCategory & ", " & txtSearchType & ", " & Sort & ", " & SortDir & ", " & Page & "<br>"

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
			
			'Response.Write "Record_Filter: " & Record_Filter & "<br>"


			If txtSearchType = 1 Then 
				boc = "b"
			End If
			If txtSearchType = 2 Then 
				boc = "c"
			end if
			
			If txtSearchType = 3 Then 
				boc = "k"
				
				
				arrKeywords = Split(txtFullFilter,chr(32))
				
				set rsKey = Server.CreateObject ("Adodb.recordset")
				rsKey.CursorLocation = 3
				

				Points = 10
				KeyWordCriteria = ""
												

				NameCriteria = ""
				
				for i = LBound(arrKeyWords) to UBound(arrKeyWords)

						If len (NameCriteria) > 0 Then
							NameCriteria = NameCriteria & " or "
						End If
				
					
						NameCriteria = NameCriteria & " CharIndex (' " & replace(Trim(arrKeyWords(i)),"'","''") & " ', v.CompanyName) > 0 "
						
						For j = i to UBound(arrKeyWords)
						
						
							tmpKeyWord = ""
							for k = i to j
								tmpKeyWord = tmpKeyWord & replace(Trim(arrKeyWords(k)),"'","''''") & " "
							next
							
							If Len(tmpKeyword)  > 2 Then
								On Error Resume Next	' To take care of "ignroed words"
								rsKey.Open "sp_KeyWord_Find '" & tmpKeyWord & "'", cn
								'Response.Write err
								'Response.End
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
								end if 's
							End If
							
				next
					
				Next
				
				'Response.Write tmpKeyWord	& i & " " & j & " " & k & "<br>"
				if AreaID > 0 Then 			
						
						If  len (PointsCriteria) > 0 Then
							PointsCriteria = PointsCriteria & " + "
						End If
						
						AreaCriteria = "charindex('|a" & AreaID & "|', keywords) > 0"
						
									
						' KeyWordCriteria = KeyWordCriteria & "charindex('|a" & AreaID & "|', keywords) > 0"
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
	oFilter    = Trim(Request.QueryString("sText"))
	oFilter2 = StripAll(oFilter)
	remote.Session ("oFilter2") = oFilter2
Else
	oFilter2 = remote.Session ("oFilter2")
End If


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

%>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function bdy_ondragstart() {
	window.event.returnValue = false;
	return false;
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=bdy EVENT=ondragstart>
<!--
 bdy_ondragstart()
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript1.2" onload="return window_onload();" id="bdy" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" style="background-color: silver">
<div ID="tooltip" STYLE="left:0px;font-family: Helvetica; font-size: 8pt; position: absolute; z-index: 200; display: none; visibility: hidden; width:207px">
	<iframe height="<%=ToolTipHeight%>" width="640" frameborder="0" style="border-style: none; border-width: 1px;" src="LocationTooltip.asp?select=no&Task=yes&bl=true" id="frameToolTip" scrolling="no"></iframe>
</div>
<iframe onload=sta() src="LoadingAppointment.asp" id=frameSTA style=display:none;visibility:hidden></iframe>
<%
if RecordCount > 0 then

	Response.Write ("<Center><TABLE width=100% border=0 cellspacing=0 cellpadding=1 style=""font-family:Tahoma;font-size:11px"">")
	'Response.Write "<TR style=""background-color:silver;border-style:solid""><TD style=""width:20"">&nbsp;&nbsp;&nbsp;&nbsp;</td><TD>Name</TD><tD>Address</TD><tD>City</td><td>Phone</td><td>Map</td><td>Miles</Td><td>Menu</td><td>Web</td><td>Stars</td><td>Price</td></tr>"

	Response.Write vbCRLF & vbCRLF & "<script language=javascript>" & vbCRLF
	'Response.Write "if(parent.scr.rec != " & Max_Records & ")" & vbCRLF
	'Response.Write "parent.vso.listrows = " & Max_Records & ";" & vbCRLF
	mmrec = Recordcount
	Response.Write "if (parent.scr) if(parent.scr.max != " & mmrec & " || parent.scr.rec !=" & Max_Records & ")" & vbCRLF
	Response.Write "parent.scr.init(" & mmrec & "," & Max_Records & ");" & vbCRLF


	Response.Write "</script>" & vbCRLF & vbCRLF

	i = 0

	'Response.Write remote.Session("loc_Recordset").Bookmark


	Response.Write ("<TR>" & vbcrlf)
	Response.Write "<TD class=""bord"" style=""width:20;border-style:outset;border-width:1px;"">&nbsp;</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intNameWidth & ";border-style:outset;border-width:1px;"">Name</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intAddressWidth & ";border-style:outset;border-width:1px;"">Address</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intCityWidth & ";border-style:outset;border-width:1px;"">City</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intPhoneWidth & ";border-style:outset;border-width:1px;"">Phone</td>"
	if intFaxWidth > 0 then
		Response.Write "<TD class=""bord"" style=""width:" & intFaxWidth & ";border-style:outset;border-width:1px;"">Phone 2</td>"
	end if
	Response.Write "<TD class=""bord"" style=""width:" & intMapWidth & ";border-style:outset;border-width:1px;"" align=center>Map</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intMilesWidth & ";border-style:outset;border-width:1px;"" align=center>Miles</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intWebWidth & ";border-style:outset;border-width:1px;"" align=center>Web</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intWeb2Width & ";border-style:outset;border-width:1px;"" align=center>Web2</td>"
	Response.Write "<TD class=""bord"" style=""width:" & intStarsWidth & ";border-style:outset;border-width:1px;"">Stars</td>"
	'Response.Write "<TD class=""bord"" style=""width:30;border-style:outset;border-width:1px;"">Price</td>"
	Response.Write "</TR>"
	lastid = 0
	aCount = 0
	
	strFoundLen = 0
	
	Do while  i <= RecordCount - 1
		row = i
		rslid = rs("LocationID")
		'if lastid <> rslid then
		lastid = rslid

		if strBColor = "EDFAE7" then
			strBColor = "FFFFFF"
		else
			strBColor = "EDFAE7"
		end if
				
		'Response.Write Ucase(Trim(oFilter2))
		If Len(oFilter2) > 4 Then 
			If Ucase(Trim(oFilter2)) = Ucase(Left(StripAll(rs("Alias")),len(Trim(oFilter2)))) Then
				strBColor = "A0FAE7"
			Else
				
					For thisI = 1 to len(Trim(oFilter2))
						If Ucase(Trim(Left(oFilter2,thisI))) = Ucase(Left(StripAll(rs("Alias")),thisI)) and thisI > strFoundLen Then
							'strBColor = "A0FAE7"	
							strFoundLen = thisI
							strFoundID = "tr" & i
							'AlreadyFound = true
						End If
					next
				
			End If
		End If
				
				
		raid = rslid
		cnt = i
				
		Response.Write (vbcrlf & "<TR bgcolor=#" & strBColor & " id=tr" & cnt & ">")
				
		
		'Response.Write (vbcrlf & "<TR bgcolor=#" & strBColor & ">")

		' ID
		Response.Write "<TD class=""bord"" style=""height:21;background-color:#317142""><Img width=15 height=15 name=""" & "n" & rslid & """ id=""" & "r" & i & """ onmousedown=""optCheck(this)"" src=""images/OptionUnChecked.jpg""></td>" & vbcrlf
		' Name
		Response.Write "<span language=""javascript1.2"" ondblclick=""mosel(" & raid & "," & cnt & ");"" onmousedown=""mosel(" & raid & "," & cnt & ")"" onmouseout=""mout(" & raid & "," & cnt & ")"" onmouseover=""mo(" & raid & "," & cnt & ")"">"
				
		'onmouseover=""setStatus('" & replace(Trim(rs("Name")),"'","<<sq>>") & "');return true;"" onmouseout=""window.status=''""
		rsa = Trim(rs("Alias"))
		Response.Write "<TD nowrap class=""bord""><div title=""" & rsa & """ style=""overflow:hidden;width=" & intNameWidth-4 & ";height:15""><a onmousedown=""window.event.cancelBubble=true;"" id=""" & "a" & rslid & """  href=""javascript:parent.viewOneOff(" & rslid & ")"">" & rsa & "</a></div></td>"
		' Address
		rsad = Trim(rs("Address"))
		Response.Write "<TD nowrap class=""bord""><div title=""" & rsad & """ style=""overflow:hidden;width=" & intAddressWidth & ";height:15"">" & rsad & "</a></div></td>"
		'City
		Response.Write "<TD class=""bord"" nowrap><div style=""overflow:hidden;width=" & intCityWidth & ";height:15"">" &  Trim(rs("City")) & "</a></div></td>"
		'Phone
		If Trim(rs("Phone")) <> "" Then Phone = rs("phone") Else Phone = " "
		Response.Write "<TD nowrap class=""bord""><div title=""" & Phone & """ style=""overflow:hidden;width=" & intPhoneWidth & ";height:15"">" &  Phone & "</div></td>"

		'Fax
		if intFaxWidth > 0 then
			If Trim(rs("FaxNumber")) <> "" Then Fax = rs("FaxNumber") Else Fax = " "
			Response.Write "<TD nowrap class=""bord""><div title=""" & trim(Fax) & """ style=""overflow:hidden;width=" & intFaxWidth & ";height:15"">" &  Fax & "</div></td>"
		end if
		
		' Map
		If rs("Map") Then Map = "Map" Else Map = ""
		Response.Write "<TD align=center class=""bord"">" & Map & "</td>"
		'Miles
		if n2z(rs("Miles")) = 0 then
			strMiles = ""
		else
			strMiles = formatnumber(rs("Miles"),1)
		end if
		Response.Write "<TD class=""bord"" align=right>" & strMiles & "</td>"
		'' Menu
		'If len(trim(rs("Menu"))) > 0 Then
		'	'Menu = "<A target=""x"" href=""http://" & Trim(rs("menu")) & """>Menu</a>"
		'	Menu = "<A onmousedown=""window.event.cancelBubble=true;"" href=""javascript:wopen('http://" & Trim(rs("menu")) & "')"">Menu</a>"
		'else
		'	Menu = ""
		'end if
		
		' WebLinkType
		
		'Website
		If Trim(rs("Website")) <> "" Then Website = "<A onmousedown=""window.event.cancelBubble=true;"" onclick=""parent.document.all.sText.focus();"" href=""javascript:wopen('" & Trim(rs("websiteprotocol")) & Trim(rs("website")) & "')"">Web</a>" Else Website = " "
		Response.Write "<TD align=center class=""bord"">" &  Website & "</td>"
		
		If len(trim(rs("Menu"))) > 0 Then
			'Menu = "<A target=""x"" href=""http://" & Trim(rs("menu")) & """>Menu</a>"
			if len(trim(rs("WebLinkType"))) > 0 then
				Menu = "<A onmousedown=""window.event.cancelBubble=true;"" href=""javascript:wopen('" & Trim(rs("menu")) & "')"">" & rs("WebLinkType") & "</a>"
			else
				Menu = "<A onmousedown=""window.event.cancelBubble=true;"" href=""javascript:wopen('" & Trim(rs("menu")) & "')"">Menu</a>"
			end if
		else
			Menu = ""
		end if
		Response.Write "<TD nowrap align=center class=""bord""><div style=""overflow:hidden;width=40;height:15"">" & Menu & "</div></td>"
		
		
		'' Stars
		if isnull(rs("Stars")) then
			str = ""
		else
			str = replace(rs("Stars")," ","")
		end if
		if str = "No" then
			strStars = "<img id=navSadFace name=navSadFace src=""images/Avoid.gif"">"
			str = ""
		else
			strStars = str
		end if
		
		Response.Write "<TD align=""center"" title=""" & len(str) & " Stars" & """ style=""color: #B16918"" class=""bord"">" &  strStars & " </td>" 'Trim(rs("Stars"))
		'' Price
		str = ""
		str2 = ""
		'Response.Write "<TD class=""bord""><a title=""" & str2 & """>" & str & "</a></td>"

		Response.Write "</TR>"
		Response.Write "</span>"
		i = i + 1
	Loop	
	
	if strFoundID <> "" Then
		Response.Write "<script>document.all(""" & strFoundID & """).style.backgroundColor=""#A0FAE7""</script>"
	End If
	
	
	
	booBack = (recnum > Max_Records)

	if RecordCount > 0 then
		if LastRec > Max_Records then
		
			
			if (recnum + Max_Records - 1) > CLng(remote.Session ("GetRecordCount")) Then
				strFooter = "Records " & recnum & "-" & remote.Session ("GetRecordCount") & " of " & remote.Session ("GetRecordCount") & "&nbsp;&nbsp;&nbsp;"			
			Else
				strFooter = "Records " & recnum & "-" & Cstr(recnum + Max_Records - 1) & " of " & remote.Session ("GetRecordCount") & "&nbsp;&nbsp;&nbsp;"
			End If
			
			Response.Write ("<TR bgcolor=#" & "EDFAE7>" & vbcrlf)

			Response.Write "<TD align=left colspan=6 class=bord style=""border-right-style:none;"">" & vbcrlf
			'Response.Write "<table cellspacing=""0"" style=""height:21px;width:360;font-weight:bold;font-size:11px;""><tr><td align=""center"" width=""80px"">"
			Response.Write "<table cellspacing=0 style=height:21px;width:160;font-weight:bold;font-size:11px;><tr><td align=center width=80px>"
			if booBack then
				Response.Write "<span style=cursor:hand onclick=""parent.disableSearchFields();window.location='browselocations.asp?page=b'"" onmouseout=""this.style.borderWidth='0px'"" onmousedown=""this.style.borderStyle='inset'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:red;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_previous.gif><td>&nbsp;Prev&nbsp;&nbsp;</td></td></tr></table>"
				Response.Write "</span>" & vbcrlf
				'Response.Write "<span onmouseout=""this.style.borderWidth='0px'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
				'Response.Write "<a style=""color:red;text-decoration:none"" href=browselocations.asp?page=b>&nbsp;&nbsp;<< Prev&nbsp;&nbsp;</a>"
				'Response.Write "</span>" & vbcrlf
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
				Response.Write "<span style=cursor:hand onclick=""parent.disableSearchFields();window.location='browselocations.asp?page=n'"" onmouseout=""this.style.borderWidth='0px'"" onmousedown=""this.style.borderStyle='inset'"" onmouseover=""this.style.borderStyle='outset';this.style.borderWidth='1px'"">" & vbcrlf
				Response.Write "<table cellspacing=""1"" cellpadding=""0"" style=""color:red;font-weight:bold;font-size:11px;""><tr><td>&nbsp;&nbsp;Next</td><td>&nbsp;<img id=nav_next name=nav_next border=0 src=images/nav_next.gif>&nbsp;&nbsp;</td></tr></table>"
				Response.Write "</span>" & vbcrlf
			end if
			Response.Write "</td></tr></table>"
			Response.Write "</td>"
			Response.Write "<td align=right colspan=5 class=bord ""border-left-style:none;style=height:21px;width:144""><span style=""font-family:Tahoma;font-size:11px;color:gray;ffont-weight:bold"">" & strFooter & "</span></td>"
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
	strScript = strScript & "parent.document.all(""txtEOF"").value = 'false';" & vbcrlf
else
	strScript = strScript & "document.all(""bdy"").background=""images/NoRecordsFound_Gray.jpg""" & vbcrlf
	strScript = strScript & "document.all(""bdy"").style.backgroundRepeat=""no-repeat""" & vbcrlf
	strScript = strScript & "document.all(""bdy"").style.backgroundPosition=""center""" & vbcrlf
	strScript = strScript & "parent.document.all(""txtEOF"").value = 'true';" & vbcrlf
end if

strScript = strScript & "parent.enableSearchFields();" & vbcrlf
strScript = strScript & "if(parent.intKeyCode == 9) parent.document.all(""Category"").focus(); else parent.document.all(""sText"").focus();" & vbcrlf

Response.Write "<script>" & vbcrlf
Response.Write strScript
Response.Write "</script>" & vbcrlf



sub WriteCols()
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "<TD class=""bord""></td>"
	Response.Write "</TR>"
end sub

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

'if booFirstLoad then
'	Response.Flush
'	url = Application("HomePage") & "/CheckRebuildPage1Cache.asp?cid=" & cid
'	set xmlhttp = server.CreateObject("Microsoft.XMLHTTP") 
'	xmlhttp.open "GET", url, true
'	xmlhttp.send ""
'	set xmlhttp = nothing
'end if
%>

</BODY>
</HTML>

