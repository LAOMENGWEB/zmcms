<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=mc%> >> <%=ym%></title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/css/default.css" />
<link rel="stylesheet" href="<%=template%>wap/awesome/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.css" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.default.css" /> 
<!--#include file="../../../wap/include/js.asp"-->

</head>
<body>
<div class="zmcont">
<!--#include file="top.asp"-->

<header>
<div class="top mtop">
<div class="mreturn"><a href="javascript:history.go(-1);"></a></div>
<div class="mtitle"><%=tw(190)%></div>
<div class="rightnav"></div>
</div>
</header>


<%Quarters = request.QueryString("Quarters")
%>
<SCRIPT language=javaScript>
function CheckJob()
{
	if (document.form1.Quarters.value.length < 2 || document.form1.Quarters.value.length > 30){

	layer.open({
    content: '<%=tw(191)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Quarters.focus();
		return false;
	}
	if (document.form1.Name.value.length < 2 || document.form1.Name.value.length > 16){
	layer.open({
    content: '<%=tw(192)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Name.focus();
		return false;
	}
	if (document.form1.Birthday.value.length!=10){
		layer.open({
    content: '<%=tw(193)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Birthday.focus();
		return false;
	}
	if (document.form1.Stature.value.length != 3){
		layer.open({
    content: '<%=tw(194)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Stature.focus();
		return false;
	}
	if (document.form1.Residence.value.length < 4 ||document.form1.Residence.value.length > 30 ){
		layer.open({
    content: '<%=tw(195)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Residence.focus();
		return false;
	}
	if (document.form1.Edulevel.value.length < 20 ){
		layer.open({
    content: '<%=tw(196)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Edulevel.focus();
		return false;
	}
	if (document.form1.Experience.value.length < 20 ){
				layer.open({
    content: '<%=tw(197)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Experience.focus();
		return false;
	}
	if (document.form1.Phone.value.length < 11){
			layer.open({
    content: '<%=tw(198)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
		document.form1.Phone.focus();
		return false;
	}

    if(document.form1.Email.value.length!=0)
     {
       if (document.form1.Email.value.charAt(0)=="." ||        
            document.form1.Email.value.charAt(0)=="@"||       
            document.form1.Email.value.indexOf('@', 0) == -1 || 
            document.form1.Email.value.indexOf('.', 0) == -1 || 
            document.form1.Email.value.lastIndexOf("@")==document.form1.Email.value.length-1 || 
            document.form1.Email.value.lastIndexOf(".")==document.form1.Email.value.length-1)
        {
         		layer.open({
    content: '<%=tw(199)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
         document.form1.Email.focus();
         return false;
         }
      }
    else
     {
      		layer.open({
    content: '<%=tw(200)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
    });
      document.form1.Email.focus();
      return false;
      }
 }
</SCRIPT>
<div class="mobile_form ">
<ul>
<form name="form1" method="post" onSubmit="return CheckJob()" action="<%=sitepath%>inc/mobiletjjl/?MemberID=0&action=add&m=<%=request.QueryString("m")%>&language=<%=language%>">


<li class="input required">

<div class="xbox"><input name="Quarters" type="text"  class='input-text'  value="<%=Quarters%>"/>
  </div>
</li>

<li class="input">

<div class="xbox"><input  name="Name"  type="text"  class='input-text' placeholder="<%=tw(201)%>"/> </div>
</li>


<li class="radio">
<span class="name"><%=tw(202)%></span>

<div class="xbox"><label><input  name="sex"  type="radio"   vaule="<%=tw(203)%>"  checked="checked"/> <%=tw(203)%></label> </div>
<div class="xbox"><label><input  name="sex"  type="radio"   vaule="<%=tw(204)%>"/> <%=tw(204)%></label></div>

</li>
  

<li class="select">
<span class="name"><%=tw(206)%>：</span>
      <select id="Marry"   name="Marry" >
          <option value="<%=tw(207)%>"  selected><%=tw(207)%></option>
          <option  value="<%=tw(208)%>"><%=tw(208)%></option>
      </select>
	  </li>


  
  <li class="input">
<span class='info'><%=tw(209)%>（<%=tw(210)%>）</span>
<div class="xbox"><input  name="Birthday"  type="text"  class='input-text' placeholder="<%=tw(211)%>"/> </div>
</li>
  
<li class="input">
<span class='info'>(<%=tw(212)%>)</span>
<div class="xbox"><input  name="Stature"  type="text"  class='input-text' placeholder="<%=tw(213)%>"/></div>
</li>
  
 <li class="input">
<div class="xbox"><input  name="School"  type="text"  class='input-text' placeholder="<%=tw(214)%>"/></div>
</li> 

 <li class="input">
<div class="xbox"><input  name="Studydegree"  type="text"  class='input-text' placeholder="<%=tw(215)%>"/></div>
</li> 

 <li class="input">
<div class="xbox"><input  name="Specialty"  type="text"  class='input-text' placeholder="<%=tw(216)%>"/></div>
</li> 
 <li class="input">
<div class="xbox"><input  name="Gradyear"  type="text"  class='input-text' placeholder="<%=tw(217)%>"/></div>
</li> 
 <li class="input">
<div class="xbox"><input  name="Residence"  type="text"  class='input-text' placeholder="<%=tw(218)%>"/></div>
</li> 
<li class="input required">
<div class="xbox"><textarea name='Edulevel' class='textarea-text' placeholder="<%=tw(219)%>"></textarea></div>
</li>

<li class="input required">
<div class="xbox"><textarea name='Experience' class='textarea-text' placeholder="<%=tw(220)%>"></textarea></div>
</li>

   <li class="input">
   <span class='info'>(<%=tw(221)%>)</span>
<div class="xbox"><input  name="Phone"  type="text"  class='input-text' placeholder="<%=tw(222)%>"/></div>
</li> 
 <li class="input">
<div class="xbox"><input  name="Email"  type="text"  class='input-text' placeholder="<%=tw(223)%>"/></div>
</li> 
  
 <li class="input">
<div class="xbox"><input  name="Add"  type="text"  class='input-text' placeholder="<%=tw(224)%>"/></div>
</li> 
  <li class="input">
<div class="xbox"><input  name="Postcode"  type="text"  class='input-text' placeholder="<%=tw(225)%>"/></div>
</li> 
 
<div class="input-submit"><input type='submit' name='Submit' value='<%=tw(226)%>' /></div>
</form>
</ul>
</div>
<!--#include file="bottom.asp"--> 
</div>
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>


