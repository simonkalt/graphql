<%
Response.CacheControl = "No-Cache"
Response.AddHeader "Pragma", "No-cache"
Response.Expires = -1
%>
<script language="javascript1.2">
	var strChanged, trl, bl, moveLeft, moveRight, vv, oContainer, glbBGColor = 'white', booUSAMask;
	
	function CreatePhoneField( strID, strFontStyle, strControlHeight, strControlWidth, intTabIndex, oCont, strBGColor, booUSA )
		{
		var i, j, k, fontsize, wthree, wfour;
		var s = new String();
		var a = strFontStyle.split(" ");
		var s = "";

		oContainer = oCont
		glbBGColor = strBGColor
		
		for(i=0; i<a.length; i++)
			s += a[i];
		
		strFontStyle = s;
		
		i = strFontStyle.indexOf("font-size:")+10;
		j = strFontStyle.indexOf(";",i)-1
		if(j==0)
			k = strFontStyle.slice(i,j);
		else
			k = strFontStyle.slice(i);
		
		fontsize = parseInt(k);
		xpfontsize = fontsize;
		wthree = (fontsize*2)-3;
		wfour = Math.round(wthree*1.33);
			
		vv = strFontStyle;
		strFontStyle = vv.substring(0,i)+xpfontsize+"px";
		var strTabIndex
		if(intTabIndex)
			strTabIndex = "tabindex="+intTabIndex
		else
			strTabIndex = ""
		
		if(!glbBGColor)
			glbBGColor = 'white';
			
		var strTableStyle		= "<table language=javascript1.2 id=tblInner bgcolor="+glbBGColor+" style="+strFontStyle+" cellspacing=0 cellpadding=0 border=0>";
		var strGrayBorder		= "<table cellpadding=0 cellspacing=0 width=100% style=border-style:inset;border-color:silver;border-width:1px>"
		
		if(booUSA==null || booUSA == true)
		{
			var strAreaCodeStyle	= "<td><P style="+strFontStyle+">&nbsp;(</p></td><td><p><input language=javascript1.2 "+strTabIndex+" id=pm_AreaCode_"+strID+" name=pm_AreaCode_"+strID+" type=text maxlength=3 minlength=3 onpaste=pastePhone('"+strID+"') onfocus=settrl(this) onkeydown=saveVal(this) onkeypress=validateNum() onblur=validateLen(3,this) onkeyup=TabOn(3,this) style="+strFontStyle+";width:"+wthree+"px;border-style:none;border-width:0px;height:"+strControlHeight+";padding:0;background-color:"+glbBGColor+"></p></td><td><P style="+strFontStyle+">)</P></td>";
			var strPrefixStyle		= "<td><P style="+strFontStyle+">&nbsp;</td><td><p><input language=javascript1.2 id=pm_Prefix_"+strID+" name=pm_Prefix_"+strID+" type=text minlength=3 maxlength=3 onpaste=pastePhone('"+strID+"') onfocus=settrl(this) onkeydown=saveVal(this) onkeypress=validateNum() onblur=validateLen(3,this) onkeyup=TabOn(3,this) style="+strFontStyle+";width:"+wthree+"px;border-style:none;border-width:0px;height:"+strControlHeight+";padding:0;background-color:"+glbBGColor+"></P></td>";
			var strSuffixStyle		= "<td><P style="+strFontStyle+">-</p></td><td><p><input language=javascript1.2 id=pm_Suffix_"+strID+" name=pm_Suffix_"+strID+" type=text minlength=4 maxlength=4 onpaste=pastePhone('"+strID+"') onfocus=settrl(this) onkeydown=saveVal(this) onkeypress=validateNum() onblur=validateLen(4,this) onkeyup=TabOn(4,this) style="+strFontStyle+";width:"+wfour+"px;border-style:none;border-width:0px;height:"+strControlHeight+";padding:0;background-color:"+glbBGColor+"></P></td>";
			booUSAMask = true;
		}
		else
		{
			var strAreaCodeStyle	= "";
			var strPrefixStyle		= "<td><P style="+strFontStyle+">&nbsp;</td><td><p><input type=text id=pm_Prefix_"+strID+" name=pm_Prefix_"+strID+" style="+strFontStyle+";width:"+(wthree+wthree+wthree+wthree+wfour)+"px;border-style:none;border-width:0px;height:"+strControlHeight+";padding:0;background-color:"+glbBGColor+"></P></td>";
			var strSuffixStyle		= "";
			booUSAMask = false;
		}		

		var strControl			= "<table id=tblControl style="+strFontStyle+";border-style:none;border-color:silver;width:"+strControlWidth+";height:"+strControlHeight+" cellspacing=0 cellpadding=0 border=1 bgcolor="+glbBGColor+">";
		var strDiv				= "<div id=pm_div_"+strID+">"

		var str = "";
		str += strDiv;
		str += "<input type=hidden id="+strID+" name="+strID+">";
		str += strControl;
		str += "<tr>";
		str += "<td>";

		str += strGrayBorder;
		str += "<tr>";
		str += "<td>";

		str += strTableStyle;
		str += "<tr>";
		
		str += strAreaCodeStyle;
		str += strPrefixStyle;
		str += strSuffixStyle;
		str += "</tr>";
		str += "</table>";
		
		str += "</td>";
		str += "</tr>";
		str += "</table>";

		str += "</td>";
		str += "</tr>";
		str += "</table>";
		str += "</div>";
		
		addTag(str);
		
		return true;
		}
	

	function addTag( str )
		{
			if(oContainer)
				oContainer.innerHTML = str;
			else
				document.write(str);
		}
		
		
	function TabOn( intLen, field )
		{
		var i;
		var fkey;
		var fname;
		var ac, bs;
		var p;
		var s;
		var tr;
		var fieldname = field.name;
		
		i = fieldname.lastIndexOf("_");
		fkey = fieldname.slice(0,i);
		fname = fieldname.slice(i+1);
		
		//alert(window.event.keyCode);
		
		switch(fkey)
			{
			case "pm_AreaCode":
				{
				p = "pm_Prefix_"+fname
				if(window.event.keyCode < 37 || window.event.keyCode > 40)
					{
					if(window.event.keyCode != 9 && window.event.keyCode != 16)
						{
						if(field.value.length == intLen)
							{
							document.all(p).focus();
							tr = document.all(p).createTextRange();
							if(tr.expand("textedit"))
								tr.select();
							}
						}
					}
				else
					{
					if(moveRight)
						document.all(p).focus();
					}
					break;
				}
			case "pm_Prefix":
				{
				ac = "pm_AreaCode_"+fname
				s  = "pm_Suffix_"+fname;
				p  = "pm_Prefix_"+fname

				if(window.event.keyCode < 37 || window.event.keyCode > 40)
					{
					if(window.event.keyCode != 9 && window.event.keyCode != 16)
						{
						if(field.value.length == intLen)
							{
							document.all(s).focus();
							tr = document.all(s).createTextRange();
							if(tr.expand("textedit"))
								tr.select();
							}
						}
					if(strChanged == field.value)
						{
						if(window.event.keyCode == 8)  //backspace
							{
							tr = document.all(ac).createTextRange();
							document.all(ac).focus();
							tr.collapse();
							if(tr.expand("textedit"))
								{
								tr.collapse(false);
								tr.select();
								}
							}
						}
					}
				else
					{
					if(moveLeft)
						{
						tr = document.all(ac).createTextRange();
						document.all(ac).focus();
						tr.collapse();
						if(tr.expand("textedit"))
							{
							tr.collapse(false);
							tr.select();
							}
						}
					}
					if(moveRight)
						document.all(s).focus();
				break;
				}
			case "pm_Suffix":
				{
				s = "pm_Prefix_"+fname;

				if(window.event.keyCode < 37 || window.event.keyCode > 40)
					{
					if(window.event.keyCode == 8)  //backspace
						{
						if(strChanged == field.value)
							{
							tr = document.all(s).createTextRange();
							document.all(s).focus();
							tr.collapse();
							if(tr.expand("textedit"))
								{
								tr.collapse(false);
								tr.select();
								}
							}
						}
					}
				else
					{
					if(moveLeft)
						{
						tr = document.all(s).createTextRange();
						document.all(s).focus();
						tr.collapse();
						if(tr.expand("textedit"))
							{
							tr.collapse(false);
							tr.select();
							}
						}
					}
				}
				break;
			}
			var a = "pm_AreaCode_"+fname;
			var b = "pm_Prefix_"+fname;
			var c = "pm_Suffix_"+fname;
			
			document.all(fname).value = document.all(a).value + document.all(b).value + document.all(c).value;
		}


	function validateLen( intLen, field )
		{
		var booRetVal = true;
		//if(field.value.length > 0 && field.value.length < intLen)
		//	{
		//	alert("You must enter "+intLen+" numbers here.");
		//	booRetVal = false;
		//	field.focus();
		//	}
		return(booRetVal);
		}

	function validateNum()
		{
			if(window.event.keyCode < 48 || window.event.keyCode > 57)
				window.event.keyCode = 0;
		}

		
	function saveVal( field )
		{
		var intKeyCode;
		var tr;

		if(window.event.keyCode == 38)
			window.event.keyCode = 37;
		if(window.event.keyCode == 40)
			window.event.keyCode = 39;
		
		intKeyCode = window.event.keyCode;
		
		tr = document.selection.createRange();
		moveLeft = false;
		moveRight = false;

		switch( intKeyCode )
			{
			case 8: // backspace
				{
					strChanged = field.value;
					break;
				}
			case 37:  // Left
				{
					if(tr.compareEndPoints("StartToStart", trl)==0)
						{
						//alert("Beginning of Range");
						moveLeft = true;
						}
					break;
				}
			case 39:  // Right
				{
					if(tr.compareEndPoints("EndToEnd", trl)==0)
						//alert("End of Range");
						moveRight = true;
					break;
				}
			}
		}	

	function pastePhone( field )
		{
		window.event.returnValue = false;
		var s = window.clipboardData.getData("Text");
		var str = s.replace(/\-/g,"").replace(/\(/g,"").replace(/\)/g,"").replace(/\ /g,"").replace(/\./g,"").replace(/[/]/g,"");
		str = str.substr(0,10);
		FillPhone( field, str );
		return (false);
		}
		
	function FillPhone( field, str )
		{
			var ac, p, s;
			
			ac = "pm_AreaCode_" + field;
			p  = "pm_Prefix_"   + field;
			s  = "pm_Suffix_"   + field;

			if(booUSAMask)
			{
				if(str)	   //check if defined (passed null)
				{
					document.all(ac).value = str.slice(0,3);
					document.all(p).value = str.slice(3,6);
					document.all(s).value = str.slice(6);
				}
				else
				{
					document.all(ac).value = '';
					document.all(p).value = '';
					document.all(s).value = '';
				}
			}
			else
			{
				if(str)
					document.all(p).value = str;
			}

			document.all(field).value = str;
		}
		
	function settrl( field )
		{
			trl = field.createTextRange();
			bl = trl.boundingLeft;
		}		
		
	function isXP()
		{
			var hua = "<%=Request.ServerVariables("HTTP_USER_AGENT")%>";
			var booXP = ((hua.indexOf("Windows NT 5.1") > -1) || (hua.indexOf("Windows XP") > -1));
			return (booXP)
		}
		
	function pm_enabled( fieldname, boo )
		{
			var ac, p, s
			ac = "pm_AreaCode_"+fieldname;
			p  = "pm_Prefix_"+fieldname;
			s  = "pm_Suffix_"+fieldname;
			document.all(ac).disabled = !boo;
			document.all(p).disabled  = !boo;
			document.all(s).disabled  = !boo;

			if(boo == true)
				{
					document.all("tblInner").style.backgroundColor = glbBGColor;
					document.all("tblInner").style.color   = "black";
					document.all("tblControl").style.backgroundColor = glbBGColor;
					document.all(ac).style.backgroundColor = glbBGColor;
					document.all(p).style.backgroundColor  = glbBGColor;
					document.all(s).style.backgroundColor  = glbBGColor;
				}
			else
				{
					document.all("tblInner").style.backgroundColor   = "silver";
					document.all("tblInner").style.color   = "gray";
					document.all("tblControl").style.backgroundColor = "silver";
					document.all(ac).style.backgroundColor = "silver";
					document.all(p).style.backgroundColor  = "silver";
					document.all(s).style.backgroundColor  = "silver";
				}
		}
</script>
