<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Enter News</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body bgcolor="#000000" link="#CCCCCC" onLoad=document.form1.title.focus();>
<form name="form1" method="post" action="savenews.asp">
  <div align="center">
    <table width="731" height="452" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr> 
        <td width="107" height="47" valign="top">
		<%
dim cn
dim rs
dim strsql
dim str
set cn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
str="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("happen.mdb") & ";Persist Security Info=False" 
cn.open str
strsql="select * from check"
rs.open strsql,cn , 3, 3
rs.movefirst
if rs("user") <> "admin" then
%>
          <script language="javascript">
		window.location='news.asp';
	    </script> 
          <%
end if



cn.close
set cn=nothing
%>
 
                  </td>
        <td width="345">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="102"></td>
      </tr>
      <tr> 
        <td rowspan="4">&nbsp;</td>
        <td height="75" align="center" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC"> 
          <div align="center"> 
            <input name="title" type="text" tabindex="1" dir="rtl" lang="fa" size="40">
          </div></td>
        <td width="177" align="left" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC">
        <font face="Tahoma"><span lang="fa">عنوان</span></font></td>
        <td rowspan="4">&nbsp;</td>
      </tr>
      <tr> 
        <td height="95" align="center" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC"> 
          <div align="center"> 
            <input name="userdate" type="text" tabindex="2" dir="rtl" lang="fa" size="20">
          </div></td>
        <td align="left" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC">
        <font face="Tahoma"><span lang="fa">تاريخ امروز</span></font></td>
      </tr>
      <tr> 
        <td height="153" align="center" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC"> 
          <div align="center"> 
            <textarea name="new" cols="30" rows="7" tabindex="3" dir="rtl" lang="fa"></textarea>
          </div></td>
        <td align="left" valign="middle" bordercolor="#FFFF99" bgcolor="#6699CC">
        <font face="Tahoma"><span lang="fa">متن خبر</span></font></td>
      </tr>
      <tr> 
        <td height="44" valign="top" bgcolor="#6699CC"><div align="center"> 
            <input type="submit" name="Submit" tabindex="4" value="ارسال" style="font-family: Tahoma">
          </div></td>
        <td valign="top" bgcolor="#6699CC"><!--DWLayoutEmptyCell-->&nbsp;</td>
      </tr>
      <tr> 
        <td height="38">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
  </div>
</form>
</body>
</html>