<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Paramatersetting.aspx.vb" MaintainScrollPositionOnPostback="true" Inherits="System_Paramatersetting" %>
<%@ Register src="../usercontrol/main_menu.ascx" tagname="main_menu" tagprefix="uc1" %>
<%@ Register src="../usercontrol/reveserd.ascx" tagname="reveserd" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>超級房仲家管理系統</title>
<script type="text/javascript" src="js/png.js"></script>
<script type="text/javascript" src="js/pngfix.js"></script>
<link href="../css/index.css" rel="stylesheet" type="text/css"/>
<link href="../css/inpage.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab1.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab2.css" rel="stylesheet" type="text/css"/>
<link href="../css/gridview.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="wrapper"><!--wrapper版頭 --> 
    <uc1:main_menu ID="main_menu1" runat="server" />
<form id="form1" runat="server">
<!--container -------------------------------------------------------------------------------------------------------------->
<div id="inpage_container"><!--inpage_container版頭 -->
<div id="inpage_container_main"><!--inpage_container_main版頭 -->
<div id="inpage_main_bar"><!--inpage_main_bar版頭 -->
<div id="inpage_left_bar"><img src="../images/system_a.jpg" width="126" height="107" /></div><div id="inpage_bar"><img src="../images/system_bar_02.jpg" width="99" height="59" /></div><div id="inpage_route"><img src="../images/home_icon.jpg" width="14" height="11" /><a href="#">Home</a> > <a href="#">系統設定</a> > <a href="#">參數設定</a></div>
</div><!--inpage_main_bar版尾 -->
<div id="inpage_main_content"><!--inpage_main_content版頭 -->

<!--content_head -------------------------------------------------------------------------------------------------------------->
<div id="content_head"><table width="100%" height="26" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div id="tab_01">
    <ul>
      <li><a href="#">參數設定</a></li>
      </ul>
    </div></td>
  </tr>
</table>
</div>
<!--content_main -------------------------------------------------------------------------------------------------------------->
<div id="content_main" align="center">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="25">&nbsp;</td>
  </tr>
  <tr>
    <td align="left">
        <asp:DropDownList ID="store" runat="server" AutoPostBack="True" Height="23px" AppendDataBoundItems="True">             
          </asp:DropDownList>
    </td>
  </tr>
    <tr>
    <td align="center" height="70">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
            Width="100%" BorderStyle="None" GridLines="none" CssClass="GridViewStyle">
            <Columns>
                <asp:TemplateField HeaderText="編號">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("Sort") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="sort" runat="server" Text='<%# Bind("Sort") %>'></asp:Label>
                        <asp:Label ID="ParamaterID" runat="server"  Text='<%# Bind("ParamaterID") %>' Visible="false"></asp:Label>
                        <asp:Label ID="Num" runat="server" Text='<%# Bind("Num") %>' Visible="false"></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle CssClass="GridViewHeaderStyle" Width="5%" />
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="使用與否">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("UseState") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="UseState" runat="server" />
                    </ItemTemplate>
                    <HeaderStyle CssClass="GridViewHeaderStyle" Width="7%"/>
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="參數功能">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("ParamaterName") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("ParamaterName") %>'></asp:Label>
                        <asp:Button ID="btnSyncICBStoreObjects" runat="server" Text="同步全店物件"
                            CommandName="SyncICBStoreObjects" Visible="False"
                            OnClientClick="return confirm('確定要同步該店所有物件至聯房通嗎？');" />
                    </ItemTemplate>
                    <HeaderStyle CssClass="GridViewHeaderStyle" Width="15%"/>
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="left"  />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="參數值">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("Value") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="Value" runat="server" Text='<%# Bind("Value") %>'></asp:TextBox>
                        <asp:RadioButtonList ID="ValueRadio" runat="server" RepeatDirection="Horizontal" Visible="False">
                            <asp:ListItem Value="1">是</asp:ListItem>
                            <asp:ListItem Value="0">否</asp:ListItem>
                        </asp:RadioButtonList>
                    </ItemTemplate>
                    <HeaderStyle CssClass="GridViewHeaderStyle" Width="15%"/>
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center"  />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="參數說明">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" 
                            Text='<%# Bind("ParamaterContent") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("ParamaterContent") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle CssClass="GridViewHeaderStyle" Width="58%"/>
                    <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="left"  />
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
      </td>
  </tr>
    <tr>
    <td align="center" height="70"><asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../images/search_bt_06.gif" />        
      </td>
  </tr>
</table>

</div><!--content_main版尾 -->

<!--content_foot -------------------------------------------------------------------------------------------------------------->
<div id="content_foot"></div>

</div><!--inpage_main_content版尾 -->

</div><!--inpage_container_main版尾 -->
</div><!--inpage_container版尾 -->

<!--footer -------------------------------------------------------------------------------------------------------------->
</form>
<uc2:reveserd ID="reveserd1" runat="server" />
</div><!--最外層wrapper版尾 -->
    
</body>
</html>
