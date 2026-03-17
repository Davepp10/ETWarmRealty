<%@ Page Language="VB" AutoEventWireup="false" CodeFile="vacancy_search.aspx.vb" Inherits="G_WebService_vacancy_search"    %>
<%@ Register src="../usercontrol/main_menu.ascx" tagname="main_menu" tagprefix="uc1" %>
<%@ Register src="../usercontrol/reveserd.ascx" tagname="reveserd" tagprefix="uc2" %>

<!DOCTYPE html>
<html>
<head >
<title ></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<%--<script type="text/javascript" src="js/png.js"></script>--%>
<%--<script type="text/javascript" src="js/pngfix.js"></script>--%>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.4/themes/smoothness/jquery-ui.css" />  
<link href="../css/index.css" rel="stylesheet" type="text/css"/>
<link href="../css/inpage.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab1.css" rel="stylesheet" type="text/css"/>
<link href="../css/tab2.css" rel="stylesheet" type="text/css"/>
<link href="../css/gridview.css" rel="stylesheet" type="text/css">
<!-- self use js -->
<style >
.button2_style{
	width:92px;
	height:23px;
	border:none;
	background:url(../images/webservice_bt_14.gif) no-repeat;
	vertical-align:middle;
}
    .search td:first-child {text-align :left ; font-weight:bold ;padding-left:4px;}
    .search td:nth-child(2)  { background-color : #FFFFFF; padding-left:4px;}
    </style>
<!-- self use js -->
</head>
<body>
<div id="wrapper"><!--wrapper版頭 --> 
<uc1:main_menu ID="main_menu1" runat="server" />
<form id="form1" runat="server">
<asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

<!--container -------------------------------------------------------------------------------------------------------------->
<div id="inpage_container"><!--inpage_container版頭 -->
<div id="inpage_container_main"><!--inpage_container_main版頭 -->
<div id="inpage_main_bar"><!--inpage_main_bar版頭 -->
<div id="inpage_left_bar"><img src="../images/shopkeeping_a.jpg" width="126" height="107" /></div><div id="inpage_bar"><img src="../images/webservice_bar_08.jpg" width="151" height="59" /></div><div id="inpage_route"><img src="../images/home_icon.jpg" width="14" height="11" /><a href="#">Home</a> > <a href="#">店務人員</a> > <a href="#">商城線上徵才</a></div>
</div><!--inpage_main_bar版尾 -->
<!--第一個內容 -------------------------------------------------------------------------------------------------------------->
<div id="inpage_main_content"><!--inpage_main_content版頭 -->
<!--content_head -------------------------------------------------------------------------------------------------------------->
<div id="content_head"><table width="100%" height="26" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div id="tab_01">
    <ul>
            <li id="current"><a href="vacancy_add.aspx">刊登徵才資料</a></li> 
            <li><a href="vacancy_search.aspx" >查詢徵才資料</a></li> 
            <asp:Label ID="Label1" runat="server"></asp:Label>       
      </ul>
    </div></td>
  </tr>
</table>
</div>
    
<!--content_main -------------------------------------------------------------------------------------------------------------->
<div id="content_main" align="center">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
    <!-- don't remove this tr td tag -->
  <tr>
    <td height="25">&nbsp;</td>
  </tr>
    <!-- don't remove this tr td tag -->
    <tr>
        <td>
            <table class="search" width="100%" bgcolor="#e7e7e7" border="0" cellspacing="1" cellpadding="1">
                <%--<tr>

                    <td align="left" bgcolor="#FFFFFF" colspan="2"><span class="auto-style3"><strong>[</strong><a href="https://superwebnew.etwarm.com.tw/files/線上徵才官網.jpg" target="_blank"><strong style="color:red">官網顯示位置</strong></a><strong>]&nbsp; [</strong><a href="https://superwebnew.etwarm.com.tw/files/線上徵才商城.jpg" target="_blank" class="auto-style2"><strong style="color:red">商城顯示位置</strong></a></span><span class="auto-style1">]</span></td>
                </tr>--%>
                <tr>
                    <td align="left" bgcolor="#FFFFFF" colspan="2">
                        [<a href="https://img.etwarm.com.tw/files/線上徵才官網.jpg" target="_blank"><strong style="color:red">官網顯示位置</strong></a>]&nbsp;
                        [<a href="https://img.etwarm.com.tw/files/線上徵才商城.jpg" target="_blank"><strong style="color:red">商城顯示位置</strong></a>]
                    </td>
                </tr>
                <tr>

                    <td width="105" align="left" bgcolor="#f7f7f7">加盟店</td>
                    <td align="left" bgcolor="#FFFFFF">
                        <asp:DropDownList ID="DropDownList1" runat="server"   ></asp:DropDownList>
                    </td>
                </tr>
                <tr>

                    <td width="105" align="left" bgcolor="#f7f7f7">職位名稱</td>
                    <td align="left" bgcolor="#FFFFFF">
                        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="left" bgcolor="#f7f7f7">職位有效期限</td>
                    <td align="left" bgcolor="#FFFFFF">
                        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                        ~<asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
                          <script type="text/javascript"  >
                              $(function () {
                                  $("#<%=TextBox2.ClientID%>").datepicker({
                                      dateFormat: 'yy-mm-dd',
                                      changeMonth: true,
                                      changeYear: true,
                                      minDate: 0,
                                      onClose: function (selectedDate) {
                                          $("#<%=TextBox3.ClientID%>").datepicker("option", "minDate", selectedDate);
                                        }
                                  });
                                  $("#<%=TextBox3.ClientID%>").datepicker({
                                      dateFormat: 'yy-mm-dd',
                                      changeMonth: true,
                                      changeYear: true,
                                      maxDate: '+2y',
                                      onClose: function (selectedDate) {
                                          $("#<%=TextBox2.ClientID%>").datepicker("option", "maxDate", selectedDate);
                                        }

                                    });

                              });
	                         </script> 
                              <a id="link_todetail" href="#employdetail"  ></a>
                    </td>
                </tr>
               
            </table>
        </td>
    </tr>
    <tr>
        <td style ="text-align :center " height="70">
                <asp:ImageButton ID="Button1" runat="server" ImageUrl ="~/images/search_bt_03.gif" />
        </td>
        
    </tr>
    
</table>
    <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
        <td style="color:#3e3e3e">
             共計<asp:Label ID="Label3" runat="server" ForeColor="#ff0000" Font-Bold="true" Text="0"></asp:Label>&nbsp;筆
        </td>
    </tr>
    <tr>
        <td>
              
           
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="GridViewStyle" BorderStyle="None">
                <Columns>
                    <asp:BoundField HeaderText="職稱" DataField="vacancy_name" />
                    <asp:TemplateField HeaderText="職務說明" ItemStyle-HorizontalAlign="Left">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("vacancy_content")%>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Eval("vacancy_content").ToString().Replace(Environment.NewLine, "<br />")%>'></asp:Label>
                            
                        </ItemTemplate>
                        <ItemStyle Height="35px"  />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="建立日">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server"></asp:Label>
                        </ItemTemplate>
                        <ItemStyle Width="100px" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="有效日">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server"></asp:Label>
                        </ItemTemplate>
                        <ItemStyle Width="100px" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="應徵數">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label3" runat="server"></asp:Label>
                        </ItemTemplate>
                        <ItemStyle Width="70px" HorizontalAlign="Center" VerticalAlign="Middle"  />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="編輯">
                        <ItemTemplate>
                          
                            <a href="#" onclick='<%# String.Format("callFancy({0})", Eval("vacancy_oid"))%>'><img src="../images/shopkeeping_edit_01.jpg" width="59" height="23" align="absmiddle" /></a>
                            <asp:ImageButton ID="Btn_del" runat="server" CausesValidation="False" CommandName="Dele" ImageUrl ="../images/shopkeeping_edit_02.jpg" CommandArgument='<%# Eval("vacancy_oid")%>' OnClientClick ="if ( !confirm('確定刪除?')) return false;" ImageAlign="AbsMiddle" />
                           
                            <asp:Button ID="Button2" runat="server" CommandName="seedetail" CommandArgument='<%# Eval("vacancy_oid")%>' CssClass="button2_style"  />
                        </ItemTemplate>  
                        <ItemStyle Width="220px" HorizontalAlign="left" VerticalAlign="Middle" />
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle BackColor="#F96F00" BorderStyle="Solid" BorderColor="#C95A00" BorderWidth="1px" ForeColor="White" Height="30px"  HorizontalAlign="Center" />
                <RowStyle BorderColor="#C95A00" BorderStyle="Solid" BorderWidth="1px" HorizontalAlign="Center"   Height="30px" />
            </asp:GridView>
            
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ProviderName ="MySql.Data.MySqlClient" ConnectionString="server=124.9.10.168;UID=PW74AND112;PWD=ha6945AAd37@@;database=etwarm;CHARSET=utf8;allow zero datetime=true" >
            </asp:SqlDataSource>
  </td>
  </tr>
        </table>
</div><!--content_main版尾 -->

<!--content_foot -------------------------------------------------------------------------------------------------------------->
<div id="content_foot"></div>

</div><!--inpage_main_content版尾 -->


<div id="inpage_main_content"><!--inpage_main_content版頭 -->
<!--content_head -------------------------------------------------------------------------------------------------------------->
<div id="content_head"><table width="100%" height="26" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div id="tab_01">
    <ul>
      <li><a href="#">目前應徵的人員</a></li>
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
    <td align="center">
    
       <div id="employdetail">
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="GridViewStyle" BorderStyle="None">
            <Columns>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <%--<asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="false" CommandName="Del" Text="刪除" CommandArgument='<%# Eval("vacancy_bio_oid")%>' ></asp:LinkButton>--%>
                        <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Dele" ImageUrl ="../images/shopkeeping_edit_02.jpg" CommandArgument='<%# Eval("vacancy_bio_oid")%>' OnClientClick ="if ( !confirm('確定刪除?')) return false;" ImageAlign="AbsMiddle" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="投遞時間">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Eval("vacancy_bio_maketime").ToString%>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("vacancy_bio_maketime").ToString%>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="應徵者姓名" DataField="vacancy_bio_name"  />
                <asp:BoundField HeaderText="聯絡電話"  DataField="vacancy_bio_mobile" />
                <asp:BoundField HeaderText="E-MAIL"  DataField="vacancy_bio_email" />
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                         
                        <input type="button" value="詳細資料" onclick="<%# String.Format("callFancy2({0},{1},{2});", Eval("vacancy_bio_oid"), Eval("id"), IIf(IsDBNull(Eval("store_report")), "''", "'" & Eval("store_report") & "'"))%>"  />
                       
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
                 <HeaderStyle BackColor="#F96F00" BorderStyle="Solid" BorderColor="#C95A00" BorderWidth="1px" ForeColor="White" Height="30px"  HorizontalAlign="Center" />
                <RowStyle BorderColor="#C95A00" BorderStyle="Solid" BorderWidth="1px" HorizontalAlign="Center" Height="30px" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ProviderName ="MySql.Data.MySqlClient" ConnectionString="server=124.9.10.168;UID=PW74AND112;PWD=ha6945AAd37@@;database=etwarm;CHARSET=utf8;allow zero datetime=true" >
        </asp:SqlDataSource>

        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="GridViewStyle" BorderStyle="None">
            <Columns>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="false" CommandName="Del" Text="刪除" CommandArgument='<%# Eval("id")%>' OnClientClick="if ( !confirm('確定刪除?')) return false;"></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="Label5" runat="server" Text='<%# Eval("ID").ToString%>'></asp:Label>
                            <asp:Label ID="Label6" runat="server" Text='<%# Eval("store_report").ToString%>'></asp:Label>
                        </ItemTemplate>
                        <ItemStyle Width="100px" />
                    </asp:TemplateField>
                <asp:TemplateField HeaderText="投遞時間">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Eval("create_at").ToString%>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("create_at").ToString%>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField HeaderText="應徵者姓名" DataField="name"  />
                <asp:BoundField HeaderText="聯絡電話"  DataField="mobile" />
                <asp:BoundField HeaderText="E-MAIL"  DataField="email" />
                <asp:TemplateField ShowHeader="False">
                    <%--<ItemTemplate>
                        <input type="button" value="詳細資料" onclick="<%# String.Format("callFancy3({0},{1});", Eval("id"),IIf(IsDBNull(Eval("store_report")), "''", "'" & Eval("store_report") & "'"))%>"  />
                    </ItemTemplate>--%>
                    <ItemTemplate>
                        <asp:HyperLink ID="HyperLink1" runat="server" CssClass="inpage_content_font_08"><img src="../images/land_bt_01.gif"/></asp:HyperLink>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
                 <HeaderStyle BackColor="#F96F00" BorderStyle="Solid" BorderColor="#C95A00" BorderWidth="1px" ForeColor="White" Height="30px"  HorizontalAlign="Center" />
                <RowStyle BorderColor="#C95A00" BorderStyle="Solid" BorderWidth="1px" HorizontalAlign="Center" Height="30px" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ProviderName ="MySql.Data.MySqlClient" ConnectionString="server=124.9.10.168;UID=PW74AND112;PWD=ha6945AAd37@@;database=etwarm;CHARSET=utf8;allow zero datetime=true" >
        </asp:SqlDataSource>
    </div>
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
  
<!-- funcy box use only -->
 <div id="hidden_clicker" style="display:none;"><a id="hiddenclicker" href="#" >Hidden Clicker</a></div>
<!-- funcy box use only --> 
<script src="//code.jquery.com/ui/1.10.4/jquery-ui.js"></script>

         <!-- funcybox -->
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
   <script type="text/javascript">
     

         //呼叫fancybox
         function callFancy(my_href) {
             var hrefnew = 'vacancy/vacancy_edit.aspx?id=' + my_href;
             $.colorbox({
                 href: hrefnew,
                 width: '70%',
                 height: '60%',
                 iframe: true
             });
             return;
         }

         function callFancy2(para1,para2,para3) {
            
             $.colorbox({
                 href: 'vacancy/vacancy_biodetail.aspx?oid=' + para1 + '&id=' + para2 + '&dlr=' + para3,
                 width: '60%',
                 height: '50%',
                 iframe:true
             });
             return;
            
         }
         function callFancy3(para1, para2) {

             $.colorbox({
                 href: 'vacancy/vacancy_biodetailNew.aspx?id=' + para1 + '&dlr=' + para2,
                 width: '60%',
                 height: '50%',
                 iframe: true
             });
             return;

         }
       //useless
         $(function () {
             $(".goEmploy").click(function () {
                        
             });
         });


         $(document).ready(function () {

             /*Button Onclick*/
             $("#hiddenclicker").colorbox({
                 'width': '80%',
                 'autoScale': true,
                 'transitionIn': 'none',
                 'transitionOut': 'none',
                 'type': 'iframe',
                 'onClosed': function () {
                     parent.location.reload(true);
                 }
             });

         });
    </script>
    
</form>
<uc2:reveserd ID="reveserd1" runat="server" />
</div>
</body>
</html>
