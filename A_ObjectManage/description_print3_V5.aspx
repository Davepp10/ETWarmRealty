<%@ Page Language="VB" AutoEventWireup="false" CodeFile="description_print3_V5.aspx.vb" Inherits="description_print3_V4" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
<title>東森房仲網站總管</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<script type="text/javascript" src="js/png.js"></script>
<script type="text/javascript" src="js/pngfix.js"></script>
<script type="text/javascript" src="../js/calendar.js" ></script> 
<link href="../css/index.css" rel="stylesheet" type="text/css">
<link href="../css/inpage.css" rel="stylesheet" type="text/css">
<link href="../css/popup.css" rel="stylesheet" type="text/css">
<link href="../css/tab.css" rel="stylesheet" type="text/css">

<script type="text/JavaScript">
    var myI, myW, myH
    function ResizeIframe(i) {
        i.height = 10;
        i.width = 10;
        var b = i.contentWindow.document.body;
        myI = i;
        myW = b.scrollWidth;
        myH = b.scrollHeight;
        setTimeout("ResizeIframe2(myI,myW,myH)", 100);
    }
    function ResizeIframe2(i, w, h) {
        i.height = h;
        i.width = w;
    }

    function checknum2() {
        return true;
    }
</script>
<script type="text/javascript">
    var GB_ROOT_DIR = "./greybox/";         
</script>       

<script type="text/javascript">
    <!-- #include file="greybox/AJS.js" -->
    <!-- #include file="greybox/AJS_fx.js" -->
    <!-- #include file="greybox/gb_scripts2.js" -->
</script>

<style>
    .fixTitle { BACKGROUND: navy; COLOR: white; POSITION: relative; ; TOP: expression(this.offsetParent.scrollTop) }
    .scorllDataGrid { OVERFLOW-Y: scroll; HEIGHT: 200px ;WIDTH:100% }
</style>

<link href="greybox/gb_styles.css" rel="stylesheet" type="text/css" media="all" />

<script type="text/javascript">
    function test() {
        alert('111');
    }

    GB_myShow = function (caption, url, /* optional */height, width, callback_fn) {
        var options = {
            caption: caption,
            height: height || 510,
            width: width || 610,
            fullscreen: false,
            show_loading: false,
            callback_fn: callback_fn
        }
        var win = new GB_Window(options);
        return win.show(url);
    }
</script>
<STYLE type=text/css>.SELECT_BOX {
	BORDER-RIGHT: #333333 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #333333 1px solid; COLOR: #666666; BORDER-BOTTOM: #333333 1px solid; FONT-FAMILY: "新細明體"; BACKGROUND-COLOR: #ffffff
}
</STYLE>

<style type="text/css">
<!--
.SELECT_BOX {BORDER-RIGHT: #333333 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #333333 1px solid; COLOR: #666666; BORDER-BOTTOM: #333333 1px solid; FONT-FAMILY: "新細明體"; BACKGROUND-COLOR: #ffffff
}
-->
</style>



</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<form id="form1" runat=server>
<!-- ImageReady Slices (index.psd) -->

<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20%" align="left"><span style="font-size:16px; color:#f96f00; font-weight:bold; line-height:30px;"><img src="../images/dot_05.png" width="16" height="16" align="absmiddle" /><asp:Label ID="Label1" runat="server" Text="不動產說明書列印"></asp:Label></span></td>
  </tr>
  <tr>
    <td><table class="search" width="100%" bgcolor="#e7e7e7" border="0" cellspacing="1" cellpadding="1">
      <tr>
        <td width="25%" align="left" bgcolor="#f7f7f7" id="td1">&nbsp;</td>
        <td align="left" bgcolor="#FFFFFF" id="td1">※請依需求勾選方塊，按確定送出，即可列印所選取項目。<asp:CheckBox 
                                  ID="CheckBox9" runat="server" Text="WORD版" onclick="thesame_onclick()" 
                                Checked="True" Visible="False"/>
                        <asp:Label ID="Label2" runat="server" Visible="False"></asp:Label>
                        <asp:Label ID="src" runat="server"></asp:Label>
                        <asp:Label ID="stree1" runat="server" Visible="False"></asp:Label>
                        <asp:Label ID="stree2" runat="server" Visible="False"></asp:Label>
                          </td>
        </tr>
      <tr>
        <td width="20%" align="left" bgcolor="#f7f7f7" id="td1"><strong>成交行情</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"><asp:DropDownList ID="DropDownList1" 
                runat="server"  OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" AutoPostBack="True">
                            <asp:ListItem>依範圍</asp:ListItem>
                            <asp:ListItem>依路名</asp:ListItem>
                        </asp:DropDownList>
                            </td>
        </tr>
          <%If show = 1 Then%> 
      <tr >
        <td align="left" bgcolor="#f7f7f7" id="td1">&nbsp;</td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> 
        路(地)段名1：<asp:DropDownList ID="ddl路名" runat="server" 
                CssClass="inputbox" Visible="False"></asp:DropDownList>
            <asp:TextBox ID="TextBox26" runat="server"></asp:TextBox>&nbsp;<asp:Image ID="Image11" 
                runat="server" ImageUrl="../images/s.png" 
                      
                      ToolTip="自行輸入路名，段的部份請輸入國字，例:忠孝東路四段" ImageAlign="AbsMiddle" />
            </br>
        路(地)段名2：<asp:DropDownList ID="ddl路名2" runat="server" CssClass="inputbox" 
                Visible="False" ></asp:DropDownList>
            <asp:TextBox ID="TextBox27" runat="server"></asp:TextBox>
            </br>
        路(地)段名3：<asp:DropDownList ID="ddl路名3" runat="server" CssClass="inputbox" 
                Visible="False" ></asp:DropDownList>
            <asp:TextBox ID="TextBox28" runat="server"></asp:TextBox>
          </td>
      </tr>
         <%End If%> 
         <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>建物型態</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"><asp:DropDownList ID="DropDownList3" 
                runat="server"  
                 AutoPostBack="False" ></asp:DropDownList></td>
      </tr>
      <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>金額</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> 
            <asp:TextBox ID="TextBox1" runat="server" CssClass="inputbox" Width="72px"></asp:TextBox>
            至 
            <asp:TextBox ID="TextBox3" runat="server" CssClass="inputbox" Width="72px"></asp:TextBox>
            萬</td>
      </tr>
      <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>平方公尺</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> <asp:TextBox ID="TextBox4" runat="server" CssClass="inputbox" Width="72px"></asp:TextBox>
                        至 
                        <asp:TextBox ID="TextBox5" runat="server" CssClass="inputbox" Width="72px"></asp:TextBox>
                                    平方公尺&nbsp;&nbsp;&nbsp;</td>
      </tr>
      <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>成交日期(實價登錄版)</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> <asp:TextBox ID="TextBox22" runat="server" CssClass="inputbox" Width="72px" MaxLength="7"></asp:TextBox>
                                            <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="TextBox22"
                                                Display="Dynamic" ErrorMessage="成交起始日期須為半型數字7碼民國年，如0980101" SetFocusOnError="True"
                                                ValidationExpression="[0-9]{7}">*</asp:RegularExpressionValidator>
                                    <input id="date1" runat="server" name="date1" onclick="javascript:ShowCalendar(form1.TextBox22);Menu_OP(Calendar)"
                                                            type="button" value="..." size="20" />
                                                        至
                                    <asp:TextBox ID="TextBox25" runat="server" CssClass="inputbox" Width="72px" MaxLength="7"></asp:TextBox>
                                            <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" ControlToValidate="TextBox25"
                                                Display="Dynamic" ErrorMessage="成交結束日期須為7碼民國年，如0980101" ValidationExpression="[0-9]{7}">*</asp:RegularExpressionValidator>
                                    <input id="date2" runat="server" name="date2" onclick="javascript:ShowCalendar(form1.TextBox25);Menu_OP(Calendar)"
                                                            type="button" value="..." size="20" />
            <asp:Image ID="Image13" 
                runat="server" ImageUrl="../images/s.png" 
                      
                      ToolTip="預設:近三個月" ImageAlign="AbsMiddle" />
            &nbsp; <asp:TextBox ID="TextBox2" runat="server" CssClass="inputbox" Width="72px" MaxLength="7" Visible="False"></asp:TextBox>
                                            <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" ControlToValidate="TextBox22"
                                                Display="Dynamic" ErrorMessage="成交起始日期須為半型數字7碼民國年，如0980101" SetFocusOnError="True"
                                                ValidationExpression="[0-9]{7}">*</asp:RegularExpressionValidator>
                                    <input id="date3" runat="server" name="date3" onclick="javascript:ShowCalendar(form1.TextBox2);Menu_OP(Calendar)"
                                                            type="button" value="..." size="20" visible="False" />&nbsp;
                                    <asp:TextBox ID="TextBox6" runat="server" CssClass="inputbox" Width="72px" MaxLength="7" Visible="False"></asp:TextBox>
                                            <asp:RegularExpressionValidator ID="RegularExpressionValidator4" runat="server" ControlToValidate="TextBox25"
                                                Display="Dynamic" ErrorMessage="成交結束日期須為7碼民國年，如0980101" ValidationExpression="[0-9]{7}" Visible="False">*</asp:RegularExpressionValidator>
                                    <input id="date4" runat="server" name="date4" onclick="javascript:ShowCalendar(form1.TextBox6);Menu_OP(Calendar)"
                                                            type="button" value="..." size="20" visible="False" />
            <asp:Image ID="Image12" 
                runat="server" ImageUrl="../images/s.png" 
                      
                      ToolTip="預設:近六個月，但不含近一個月的成交物件" ImageAlign="AbsMiddle" Visible="False" />
            </td>
      </tr>
     <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>生活機能</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> 
                              <asp:CheckBoxList ID="CheckBoxList1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" Enabled="false" >
                            <asp:ListItem Value="1" Selected="True">公(私)有市場</asp:ListItem>
                            <asp:ListItem Value="2" Selected="True">超級市場</asp:ListItem>
                            <asp:ListItem Value="3" Selected="True">學校</asp:ListItem>
                            <asp:ListItem Value="4" Selected="True">警察局(分駐所、派出所)</asp:ListItem>
                            <asp:ListItem Value="5" Selected="True">行政機關</asp:ListItem>
                                  <asp:ListItem Value="6" Selected="True">體育場</asp:ListItem>
                                  <asp:ListItem Value="7" Selected="True">醫院</asp:ListItem>
                        </asp:CheckBoxList>
                            </td>
      </tr>
     <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>鄰避設施</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> 
                              <asp:Label ID="Label3" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                              <asp:CheckBoxList ID="CheckBoxList2" runat="server" RepeatColumns="5" 
                                  RepeatDirection="Horizontal" Enabled="False"  >
                                  <asp:ListItem addattr="7," Selected="True" Value="7,">飛機場</asp:ListItem>
                                  <asp:ListItem addattr="8," Selected="True" Value="8,">台電變電所用地</asp:ListItem>
                                  <asp:ListItem addattr="20," Selected="True" Value="20,">寺廟</asp:ListItem>
                                  <asp:ListItem addattr="9," Selected="True" Value="9,">地面高壓電塔(線)</asp:ListItem>
                                  <asp:ListItem addattr="10," Selected="True" Value="10,">殯儀館</asp:ListItem>
                                  <asp:ListItem addattr="11," Selected="True" Value="11,">公墓</asp:ListItem>
                                  <asp:ListItem addattr="12," Selected="True" Value="12,">火化場</asp:ListItem>
                                  <asp:ListItem addattr="13," Selected="True" Value="13,">骨灰(骸)存放設施</asp:ListItem>
                                  <asp:ListItem addattr="14," Selected="True" Value="14,">垃圾場(掩埋場、焚化廠)</asp:ListItem>
                                  <asp:ListItem addattr="16," Selected="True" Value="16,">加(氣)油站</asp:ListItem>
                                  <asp:ListItem addattr="17," Selected="True" Value="17,">瓦斯行(場)</asp:ListItem>
                                  <asp:ListItem addattr="18," Selected="True" Value="18,">葬儀社</asp:ListItem>
                              </asp:CheckBoxList>
                            </td>
      </tr>
      <tr>
        <td align="left" bgcolor="#f7f7f7" id="td1"><strong>列印條件</strong></td>
        <td align="left" bgcolor="#FFFFFF" id="td1"> 
            <asp:CheckBox ID="CheckBox8" 
                runat="server" Text="物件顯示到路名"
                          Width="100%" Checked="True" Font-Bold="True" ForeColor="Red" />
                            <asp:CheckBox ID="CheckBox4" 
                runat="server" Text="所有權人只顯示姓氏"
                          Width="100%" Font-Bold="True" ForeColor="Red" />
                          <asp:CheckBox ID="CheckBox1" 
                runat="server" Text="所有權人住址部分不顯示"
                          Width="100%" Checked="True" Font-Bold="True" ForeColor="Red" 
                AutoPostBack="True" Visible="False" />
                          <asp:CheckBox ID="CheckBox2" 
                runat="server" Text="顯示完整地址"
                          Width="100%" Checked="False" Font-Bold="True" ForeColor="Red" 
                Visible="False" />
                           <asp:CheckBox ID="CheckBox3" 
                runat="server" Text="顯示環境介紹(若無勾選，僅顯示訴求重點)"
                          Width="100%" Checked="False" Font-Bold="True" ForeColor="Red" 
                Visible="false" />
                          
            <asp:CheckBox ID="CheckBox5" runat="server" Text="一併產生物調表" Width="100%" Checked="False" Font-Bold="True" ForeColor="Red"  />
                          </td>
      </tr>
      </table></td>
  </tr>
  <tr>
    <td height="70" align="center" colspan="2">
                  <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../images/search_bt_25.gif" />
      </td>
  </tr>
  </table>






<!-- End ImageReady Slices -->
</form>
</body>
</html>
