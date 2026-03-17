<%@ Page Language="VB" AutoEventWireup="false" CodeFile="customer_mediation.aspx.vb" Inherits="B_CustomerManage_customer_mediation" MaintainScrollPositionOnPostback="True" Debug="true"%>
<%@ Register src="../usercontrol/main_menu.ascx" tagname="main_menu" tagprefix="uc1" %>
<%@ Register src="../usercontrol/reveserd.ascx" tagname="reveserd" tagprefix="uc2" %>
<%@ Register assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" namespace="System.Web.UI" tagprefix="asp" %>
<%@ Register src="../usercontrol/customer_left_menu.ascx" tagname="customer_left_menu" tagprefix="uc3" %>
<%@ Register src="../usercontrol/Customer_detail.ascx" tagname="Customer_detail" tagprefix="uc4" %>
<%@ Register src="../usercontrol/customer_top_funtion.ascx" tagname="customer_top_funtion" tagprefix="uc5" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>超級房仲家管理系統</title>
<script src="../jquery/jquery-1.8.3.min.js" type="text/javascript"></script>
<script type="text/javascript" src="../js/calendar.js"></script>
<link href="../css/index.css" rel="stylesheet" type="text/css"/>
<link href="../css/inpage.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab1.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab2.css" rel="stylesheet" type="text/css"/>
<link href="../css/gridview.css" rel="stylesheet" type="text/css">
<link href="../css/left_table.css" rel="stylesheet" type="text/css"/>
<script language="javascript" type="text/javascript">
    function CheckFunc() {
        msg = "";
        if ($('#<%=offerid.ClientID %>').val() == "") {
            msg += "請輸入編號\n";
        }
        if ($('#<%=objectname.ClientID %>').val() == "") {
            msg += "請選擇物件\n";
        }
        if ($('#<%=offerdateb.ClientID %>').val() == "") {
            msg += "請輸入有效日期起始日\n";
        }
        if ($('#<%=offerdatee.ClientID %>').val() == "") {
            msg += "請輸入有效日期截止日\n";
        }
        if (($('#<%=offerprice.ClientID %>').val() == "") && ($('#<%=offertick.ClientID %>').val() == "")) {
            msg += "請輸入斡旋保證金現金或票據金額\n";
        }

        if (msg != "") {
            alert(msg);
            return false;
        }
        else {
            return true;
        }
    }
</script>
    <!--/全選或全取消/-->
<script type="text/javascript">
    function Check(parentChk, ChildId) {
        var oElements = document.getElementsByTagName("INPUT");
        var bIsChecked = parentChk.checked;
        for (i = 0; i < oElements.length; i++) {
            if (IsCheckBox(oElements[i]) && IsMatch(oElements[i].id, ChildId)) {
                oElements[i].checked = bIsChecked;
            }
        }
    }
    function IsMatch(id, ChildId) {
        //var sPattern = '^GridView1.*' + ChildId + '$';
        var sPattern = '^GridView1.*' + ChildId + '_';
        var oRegExp = new RegExp(sPattern);
        if (oRegExp.exec(id))
            return true;
        else
            return false;
    }
    function IsCheckBox(chk) {
        if (chk.type == 'checkbox') return true;
        else return false;
    }

    //限填數字
    function JHshNumberText() {
        if (!(((window.event.keyCode >= 48) && (window.event.keyCode <= 57)) || (window.event.keyCode == 46))) {
            window.event.keyCode = 0; alert("僅限填數字");
        }
    }

    //限填數字格局
    function JHshNumberText1() {
        if (!((window.event.keyCode >= 48) && (window.event.keyCode <= 57))) {
            window.event.keyCode = 0; alert("僅限填數字");
        }
    }

</script> 
</head>

<body>
<div id="wrapper"><!--wrapper版頭 --> 
    <uc1:main_menu ID="main_menu1" runat="server" />
<form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
<!--container -------------------------------------------------------------------------------------------------------------->
<div id="inpage_container"><!--inpage_container版頭 -->
<div id="inpage_container_main"><!--inpage_container_main版頭 -->
<div id="uniquename99" style="display:block; width: 141px;" runat="server">    
<div id="inpage_main_bar"><!--inpage_main_bar版頭 -->
<div id="inpage_left_bar"><img src="../images/manage_a.jpg" width="126" height="107" /></div><div id="inpage_bar"><img src="../images/manage_bar_09.jpg" width="99" height="59" /></div>
<div id="inpage_route"><img src="../images/home_icon.jpg" width="14" height="11" />Home > 客戶管理 > 客戶斡旋</div>
</div><!--inpage_main_bar版尾 -->
</div>
<div id="inpage_main_content"><!--inpage_main_content版頭 -->

<!--content_head -------------------------------------------------------------------------------------------------------------->
<div id="content_head"><table width="100%" height="26" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div id="tab_01">
    <ul>      
        <uc5:customer_top_funtion ID="customer_top_funtion1" runat="server" />
      </ul>
    </div></td>
  </tr>
</table>
</div>
<!--content_main -------------------------------------------------------------------------------------------------------------->
<div id="content_main" align="center">
<!--content_left -------------------------------------------------------------------------------------------------------------->
<div class="content_left" align="center">
 <uc3:customer_left_menu ID="customer_left_menu" runat="server" />
</div>
<div class="content_left_foot"></div>

<!--content_right -------------------------------------------------------------------------------------------------------------->
<div class="content_right">
<div class="content_right_head"></div>
<div class="content_right_main">
<uc4:Customer_detail ID="Customer_detail1" runat="server" />
&nbsp;<table width="98%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      <td align="left">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="left"><div id="inpage_content_title">客戶斡旋</div></td>
            <td align="right"><asp:Label ID="Label3" runat="server" Visible="false">0</asp:Label>
                 <a href="#"><asp:Label 
               ID="savamaster" runat="server" Text="no" Visible="false"></asp:Label>
           </a><span style="font-size:15px; color:#3e3e3e">&nbsp;<asp:ImageButton 
            ID="save" runat="server" 
                     ImageUrl="../images/manage_bt_39.gif" Visible="True" 
            ImageAlign="Middle" OnClientClick="javascript:return CheckFunc(this);" />&nbsp;<asp:ImageButton ID="del" runat="server" 
                     ImageUrl="../images/shopkeeping_del.jpg" Visible="True" 
            ImageAlign="Middle" />&nbsp;<asp:ImageButton ID="Clear" runat="server" ImageUrl="../images/shopkeeping_edit_04.jpg" align="Middle" OnClientClick="Clear();" />
        </span></td>
          </tr>
      </table></td>
    </tr>
  <tr>
    <td align="center"><table class="search" width="100%" bgcolor="#e7e7e7" border="0" cellspacing="1" cellpadding="1">
          <tr>
            <td width="210" align="left" bgcolor="#f7f7f7" id="td1"><strong>編號</strong></td>
            <td colspan="3" align="left" bgcolor="#FFFFFF" id="td1">
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                    <ContentTemplate>
                        <asp:Label ID="sid" runat="server" Text="sid" Visible="true"></asp:Label>
                        <asp:DropDownList ID="offerid" runat="server" AutoPostBack="True" OnSelectedIndexChanged="offerid_SelectedIndexChanged" BackColor="#FFF79E" Width="120px"></asp:DropDownList>
                        <asp:HiddenField ID="hf_offerid" runat="server" />
                        &nbsp;<asp:Label ID="Label464" runat="server" style="color: #FF0000; font-weight: 700; background-color: #FFCC00;" Text="消費者是否願意提供個資" Visible="False"></asp:Label>
                        <asp:RadioButton ID="RadioButton1" runat="server" GroupName="A" style="background-color: #FFCC00" Text="同意" Visible="False" />
                        <asp:RadioButton ID="RadioButton2" runat="server" GroupName="A" style="background-color: #FFCC00" Text="不同意" Visible="False" />
                    </ContentTemplate>
                </asp:UpdatePanel>
              </td>
            </tr>
          <tr>
            <td width="210" align="left" bgcolor="#f7f7f7" id="td1"><strong>物件編號(案名)</strong></td>
            <td colspan="3" align="left" bgcolor="#FFFFFF" id="td1">
               <asp:TextBox ID="objectname" runat="server" ReadOnly="True" Width="215px" BackColor="#FFF79E"></asp:TextBox>        
        <asp:ImageButton ID="Button1" runat="server" CommandName="select" ImageUrl="../images/select_bt_01.gif" />
              </td>
            </tr>
          <tr>
            <td align="left" bgcolor="#f7f7f7" id="td1"><strong>有效期間</strong></td>
            <td align="left" bgcolor="#FFFFFF" id="td1">
               <input id="offerdateb" runat="server" name="see" class="inputbox" size="10" 
            maxlength="7" readonly="readonly" style="background-color: #FFF79E"/>
<input name="ch1" id="sofferdateb" runat="server" 
            onclick="javascript:ShowCalendar(form.offerdateb);Menu_OP(Calendar)" type="button" 
            value="..." />至<input id="offerdatee" runat="server" name="nextsee" 
              class="inputbox" size="10" 
            maxlength="7" readonly="readonly" style="background-color: #FFF79E"/><input type="button" name="date2" 
            id="sofferdatee" runat="server" value="..." 
            onclick="javascript:ShowCalendar(form.offerdatee);Menu_OP(Calendar)"/>           
              </td>
            <td width="145" align="left" bgcolor="#f7f7f7" id="td1"><strong>買方出價(萬)</strong></td>
            <td align="left" bgcolor="#FFFFFF" id="td1"><asp:TextBox ID="price" runat="server" Width="87px" 
              Height="17px" onKeyPress="javascript:JHshNumberText()"></asp:TextBox>
              </td>
            </tr>
          <tr>
            <td align="left" bgcolor="#f7f7f7" id="td1"><strong>斡旋保證金(萬)</strong></td>
            <td align="left" bgcolor="#FFFFFF" id="td1">
                 現金<asp:TextBox 
              ID="offerprice" runat="server" Width="71px" BackColor="#FFF79E"></asp:TextBox>
        票據<asp:TextBox 
              ID="offertick" runat="server" Width="71px" BackColor="#FFF79E"></asp:TextBox>
                 （需輸入其中任一項）</td>
            <td width="145" align="left" bgcolor="#f7f7f7" id="td1">&nbsp;</td>
            <td align="left" bgcolor="#FFFFFF" id="td1">&nbsp;</td>
            </tr>
            </table>
  </td>
  </tr>
</table>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
        ConnectionString="<%$ ConnectionStrings:EGOUPLOADConnectionString %>" 
        ProviderName="<%$ ConnectionStrings:EGOUPLOADConnectionString.ProviderName %>"></asp:SqlDataSource>    
    <a href="#"><asp:Label 
               ID="Label0" runat="server" Text="no"></asp:Label>
           </a>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" PageSize="15" BorderStyle="None"
              CellPadding="3" Width="100%"  AllowPaging="true" AllowSorting="True" CssClass="GridViewStyle">                
                    <RowStyle CssClass="row" />
                    <AlternatingRowStyle CssClass="alt-row" /> 
              <Columns>               
                  <asp:TemplateField HeaderText="刪除">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                      </EditItemTemplate>
                      <HeaderTemplate>
                                    <input id="chkAll" name="chkAll" onclick="Check(this,'chkSelect1')" type="checkbox" />
                      </HeaderTemplate>
                      <ItemTemplate>
                          &nbsp;<asp:CheckBox ID="chkSelect1" runat="server" />
                          <asp:Label ID="Label10" runat="server" Text='<%# Bind("要約編號") %>' Visible="false"></asp:Label>                         
                      </ItemTemplate>
                      <HeaderStyle CssClass="GridViewHeaderStyle" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />                       
                  </asp:TemplateField>
                  <asp:TemplateField>
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox5" runat="server"></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:ImageButton ID="Button5" runat="server" CommandName="edits" ImageUrl="../images/shopkeeping_edit_01.jpg"/>
                      </ItemTemplate> 
                      <HeaderStyle CssClass="GridViewHeaderStyle" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />                    
                  </asp:TemplateField>
                  <asp:BoundField DataField="要約編號" HeaderText="編號" SortExpression="要約.要約編號"> 
                       <HeaderStyle CssClass="GridViewHeaderStyle" />
                    　　<ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="center"  />                 
                   </asp:BoundField>                                                       
                  <asp:TemplateField HeaderText="物件編號(案名)" SortExpression="要約.物件編號">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox21" runat="server" Text='<%# Bind("物件編號") %>'></asp:TextBox>（<asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("建築名稱") %>'></asp:TextBox>）
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label1" runat="server" Text='<%# Bind("物件編號") %>'></asp:Label>（<asp:Label ID="Label11" runat="server" Text='<%# Bind("建築名稱") %>'></asp:Label>）
                      </ItemTemplate>  
                      <HeaderStyle CssClass="GridViewHeaderStyle" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />                      
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="有效期間" SortExpression="要約起">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("要約起") %>'></asp:TextBox>至<asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("要約訖") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label2" runat="server" Text='<%# Bind("要約起") %>'></asp:Label>至<asp:Label ID="Label5" runat="server" Text='<%# Bind("要約訖") %>'></asp:Label>
                      </ItemTemplate>
                      <HeaderStyle CssClass="GridViewHeaderStyle" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />                       
                  </asp:TemplateField>
                   <asp:BoundField DataField="要約金" HeaderText="買方出價(萬)" SortExpression="要約金"/>                                                    
                   <asp:TemplateField HeaderText="斡旋保證金(萬)" SortExpression="斡旋金">
                      <EditItemTemplate>
                          現金：<asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("斡旋金") %>'></asp:TextBox>　票據：<asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("票據面額") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                         現金： <asp:Label ID="Label31" runat="server" Text='<%# Bind("斡旋金") %>'></asp:Label>　票據：<asp:Label ID="Label6" runat="server" Text='<%# Bind("票據面額") %>'></asp:Label>
                      </ItemTemplate>
                      <HeaderStyle CssClass="GridViewHeaderStyle" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />                       
                  </asp:TemplateField>
              </Columns>               
          </asp:GridView></td>
  </tr>
</table>
 
</div>
<div class="content_right_foot"></div>
</div>
</div><!--content_main版尾 -->

<!--content_foot -------------------------------------------------------------------------------------------------------------->
<div id="content_foot"></div>

</div><!--inpage_main_content版尾 -->
</div><!--inpage_container_main版尾 -->
</div><!--inpage_container版尾 -->

<!--footer -------------------------------------------------------------------------------------------------------------->
</form>
<script type="text/javascript" src="../js/colorbox/jquery.colorbox-min.js"></script>
<link rel="stylesheet" type="text/css" href="../js/colorbox/colorbox.css" media="screen" />
<script type="text/javascript">
    function GB_showCenter(v, url, height, width) {
        $.colorbox({
            'width': width,
            'height': height,
            'iframe': true,
            'href': url,
            'fixed': true
        });
    }
</script>
<uc2:reveserd ID="reveserd1" runat="server" />
</div><!--最外層wrapper版尾 -->
    
</body>
</html>