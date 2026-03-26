<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Obj_Add_V4.aspx.vb" Inherits="Obj_Add_V4" MaintainScrollPositionOnPostback="True" EnableEventValidation="false" Debug="true" ValidateRequest="false" %>

<%@ Register Src="../usercontrol/main_menu.ascx" TagName="main_menu" TagPrefix="uc1" %>
<%@ Register Src="../usercontrol/reveserd.ascx" TagName="reveserd" TagPrefix="uc2" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>超級房仲家管理系統</title>

    <script type="text/javascript" src="../js/png.js"></script>
    <script type="text/javascript" src="../js/pngfix.js"></script>
    <script type="text/javascript" src="../js/calendar.js"></script>
    <link href="../css/index.css" rel="stylesheet" type="text/css" />
    <link href="../css/inpage.css" rel="stylesheet" type="text/css" />
    <link href="../css/tab1.css" rel="stylesheet" type="text/css" />
    <link href="../css/tab2.css" rel="stylesheet" type="text/css" />
    <link href="../css/tab_02.css" rel="stylesheet" type="text/css">
    <link href="../css/gridview_p.css" rel="stylesheet" type="text/css">
    <script type="text/javascript">
        function mark(key, input) {
            if (key.value == "輸入關鍵字即可") {
                input.value = "";
            }
            if (key.value == "請填寫地號") {
                input.value = "";
            }
            if (key.value == "請填寫建號") {
                input.value = "";
            }
            if (key.value == "網路案名最多為40個字") {
                input.value = "";
            }
            if (key.value == "一般案名最多為15個字") {
                input.value = "";
            }
            if (key.value == "特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。") {
                input.value = "";
            }
        }
    </script>



    <script type="text/javascript">
        //物件型態
        function ClientValidate(source, arguments) {
            if (document.getElementById('DropDownList3').value == "請選擇") {
                arguments.IsValid = false;
            } else if (document.getElementById('DropDownList3').value == "") {
                arguments.IsValid = false;
            } else {
                arguments.IsValid = true;
            };

        }
        //使用分區
        function ClientValidate1(source, arguments) {
            if (document.getElementById('DropDownList16').value == "請選擇") {
                arguments.IsValid = false;
            } else if (document.getElementById('DropDownList16').value == "") {
                arguments.IsValid = false;
            } else {
                arguments.IsValid = true;
            };

        }
    </script>


    <!--/全選或全取消/-->
    <script type="text/javascript">
        function Check(parentChk, ChildId, GridId) {
            var oElements = document.getElementsByTagName("INPUT");
            var bIsChecked = parentChk.checked;
            for (i = 0; i < oElements.length; i++) {
                if (IsCheckBox(oElements[i]) && IsMatch(oElements[i].id, ChildId, GridId)) {
                    oElements[i].checked = bIsChecked;
                }
            }
        }
        function IsMatch(id, ChildId, GridId) {
            var sPattern = GridId + '_' + ChildId;  //'^GridView1.*' + ChildId + '$'; 
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

    <script type="text/javascript">

        function displayOwnershipTable(option) {

            var DropDownList2 = document.getElementById("DropDownList2");
            var type = DropDownList2.options[DropDownList2.selectedIndex].value;

            if (option == "0") {

                if (type == "主建物") {
                    ownership.style.display = "block"
                } else {
                    ownership.style.display = "none"
                }

            } else if (option == "1") {

                var DDL_level2 = document.getElementById("DDL_level2");
                var item = DDL_level2.options[DDL_level2.selectedIndex].value;

                if (type == "主建物" && item != "主建物','附屬物") {
                    ownership.style.display = "none"
                } else {
                    ownership.style.display = "block"
                }

            } else if (option == "2") {

                var DropDownList1 = document.getElementById("DropDownList1");
                var item = DropDownList1.options[DropDownList1.selectedIndex].value;
                if (type == "土地面積" && item != "土地面積") {
                    ownership.style.display = "none"
                } else {
                    ownership.style.display = "block"
                }

            }
        }

    </script>

    <script type="text/javascript">
        var GB_ROOT_DIR = "./greybox/";
    </script>
    <%--<script type="text/javascript" src="greybox/AJS.js"></script>
<script type="text/javascript" src="greybox/AJS_fx.js"></script>
<script type="text/javascript" src="greybox/gb_scripts2.js"></script>--%>


    <link href="greybox/gb_styles.css" rel="stylesheet" type="text/css" media="all" />



    <script type="text/javascript">

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



    </script>
    <style type="text/css">
        #loading {
            display: none;
            width: 100%;
            height: 100%;
            position: fixed;
            z-index: 99;
        }

            #loading #loading_bg {
                width: 100%;
                height: 100%;
                background: #FFF;
                filter: alpha(opacity=50);
                _filter: alpha(opacity=50); /*IE6*/
                -moz-opacity: 0.5; /*fireFox*/
                opacity: 0.5;
            }

            #loading #loading_box {
                border: 0;
                position: absolute;
                top: 50%;
                left: 50%;
                margin-left: -50px;
                margin-top: -50px;
            }

        .auto-style1 {
            width: 156px;
        }
    </style>
</head>

<body>
    <div id='loading'>
        <div id="loading_bg"></div>
        <div id="loading_box">
            <div>
                <img src="../images/loading.gif" width="100" height="100" alt="loading" />
            </div>
        </div>
    </div>
    <div id="wrapper">
        <!--wrapper版頭 -->
        <uc1:main_menu ID="main_menu1" runat="server" />
        <form id="form1" runat="server">
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <asp:HiddenField ID="hidCurrentTab" runat="server" />
            <!--container -------------------------------------------------------------------------------------------------------------->
            <div id="inpage_container">
                <!--inpage_container版頭 -->
                <div id="inpage_container_main">
                    <!--inpage_container_main版頭 -->
                    <div id="uniquename99" style="display: block; width: 141px;" runat="server">
                        <div id="inpage_main_bar">
                            <!--inpage_main_bar版頭 -->
                            <div id="inpage_left_bar">
                                <img src="../images/objectmanage_a.jpg" width="126" height="107" />
                            </div>
                            <div id="inpage_bar">
                                <img src="../images/objectmanage_bar_14.jpg" width="188" height="59" />
                            </div>
                            <div id="inpage_route">
                                <img src="../images/home_icon.jpg" width="14" height="11" /><a href="#">Home</a> > <a href="#">物件管理</a> > <a href="#">物件資料新增</a>
                            </div>
                        </div>
                        <!--inpage_main_bar版尾 -->
                    </div>
                    <div id="inpage_main_content">
                        <!--inpage_main_content版頭 -->

                        <!--content_head -------------------------------------------------------------------------------------------------------------->
                        <div id="content_head">
                            <table width="100%" height="26" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>
                                        <div id="tab_01">
                                            <ul>
                                                <%-- <li><a href="Obj_Add.aspx?state=add&src=NOW">物件新增</a></li>
      <li id="current"><a href="Object_Search.aspx">物件查詢及維護</a></li>--%>
                                                <asp:Label ID="li2" runat="server" Text=""></asp:Label>
                                                <!--<li   id="current"><a href="Obj_PriceChange.aspx?oid=10928BAA13513&sid=A0928&src=NOW&cls=Sale">契變</a></li>-->
                                            </ul>
                                        </div>
                                    </td>
                                </tr>
                            </table>
                        </div>

                        <!--content_main -------------------------------------------------------------------------------------------------------------->
                        <div id="content_main" align="center">


                            <table width="98%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td height="25">
                                        <input id="Hidden1" type="hidden" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left">
                                        <div id="tabsI">

                                            <ul id="myul">

                                                <li id="current"><a href="#View1"><span>建物基本資料</span></a></li>
                                                <li id="current_A"><a href="#View2"><span>謄本資料</span></a></li>
                                                <li id="current_B"><a href="#View3"><span>生活周邊</span></a></li>
                                                <li id="current_C"><a href="#View4"><span>附件及重要交易條件</span></a></li>
                                                <li><a href="#View5"><span>空白文件及範本</span></a></li>
                                            </ul>
                                            <script type="text/javascript">
                                                $("#current_A").css({ display: "none" });
                                                $("#current_B").css({ display: "none" });
                                                $("#current_C").css({ display: "none" });
                                            </script>
                                        </div>
                                    </td>
                                </tr>

                                <tr>
                                    <td height="4" bgcolor="#f69d47">
                                        <img src="../images/space.gif" width="5" height="4" /></td>
                                </tr>

                                <tr>
                                    <td height="6" align="center" style="background: url(../images/content_m_bg.png) left top repeat-x">
                                        <table width="99%" border="0" cellspacing="0" cellpadding="0">
                                            <tr>

                                                <td height="10">
                                                    <img src="../images/space.gif" width="5" height="5" /></td>
                                            </tr>
                                            <tr>
                                                <td align="center">

                                                    <!--<div id="View1" style="display:block" runat="server">-->

                                                    <div id="View1">
                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <%If copy = 1 Then%>
                                                            <tr>
                                                                <td id="td1" width="105" bgcolor="#f7f7f7"><strong>複製類型</strong></td>
                                                                <td id="td1" colspan="3" bgcolor="#FFFFFF">
                                                                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatColumns="3"
                                                                        RepeatDirection="Horizontal">
                                                                        <asp:ListItem Value="normal">一般</asp:ListItem>
                                                                        <asp:ListItem Value="flow">同法人或同體係流通</asp:ListItem>
                                                                        <asp:ListItem Value="many">一約多屋</asp:ListItem>
                                                                    </asp:RadioButtonList>
                                                                </td>
                                                            </tr>
                                                            <%End if%><%If objcls = 1 Then%>
                                                            <tr>
                                                                <td id="td1" bgcolor="#f7f7f7" colspan="4">
                                                                    <asp:Label ID="Label21" runat="server" Text="謄本資料、生活周邊、附件及重要交易條件需等物件存檔後，才會顯示" Style="font-weight: 700; color: #FF0000; font-size: large"></asp:Label>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td id="td1" width="105" bgcolor="#f7f7f7"><strong>物件類型</strong></td>
                                                                <td id="td1" colspan="3" bgcolor="#FFFFFF">
                                                                    <asp:RadioButtonList ID="RadioButtonList2" runat="server" RepeatColumns="3"
                                                                        RepeatDirection="Horizontal" AutoPostBack="True">
                                                                        <asp:ListItem Selected="True">售件</asp:ListItem>
                                                                        <asp:ListItem>租件</asp:ListItem>
                                                                        <asp:ListItem Enabled="False">潛件</asp:ListItem>
                                                                    </asp:RadioButtonList>
                                                                </td>
                                                            </tr>
                                                            <%End if%>
                                                            <tr>
                                                                <td id="td1" width="105" bgcolor="#f7f7f7"><strong>物件網址</strong></td>
                                                                <td id="td1" colspan="3" bgcolor="#FFFFFF">
                                                                    <asp:HyperLink ID="HyperLink1" runat="server" Target="_blank"></asp:HyperLink>

                                                                </td>
                                                            </tr>
                                                            <%--<asp:UpdatePanel ID="UpdatePanel11" runat="server" UpdateMode="Conditional">
            <ContentTemplate>--%>
                                                            <tr>
                                                                <td id="td1" width="105" bgcolor="#f7f7f7"><strong>刊登狀態</strong></td>
                                                                <td id="td1" colspan="3" bgcolor="#FFFFFF">
                                                                    <asp:UpdatePanel ID="UpdatePanel5" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:CheckBox ID="CheckBox1" runat="server" Checked="True" Text="網頁刊登" />
                                                                            <asp:CheckBox ID="CheckBox2" runat="server" Text="暫停銷售" Visible="False" />
                                                                            <asp:CheckBox ID="CheckBox5" runat="server" Text="強銷物件" />
                                                                            <asp:CheckBox ID="CheckBox100" runat="server" Text="合約終止" Visible="False" />
                                                                            <asp:CheckBox ID="CheckBox101" runat="server" Text="聯賣上架" Visible="False" />
                                                                            <asp:Label ID="Label465" runat="server"></asp:Label>
                                                                            <asp:Image ID="Image10" runat="server" ImageUrl="../images/s.png"
                                                                                ToolTip="強銷物件僅會顯示在房仲家後台首頁的區塊內，供各店自行運用。"
                                                                                ImageAlign="Middle" />
                                                                            <span style="color: #ff0033">
                                                                                <asp:Label ID="Label27" runat="server" Visible="False">庫存物件須於簽訂委託銷售契約書後，方可進行銷售行為!!</asp:Label>
                                                                            </span>
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger ControlID="ddl契約類別" />
                                                                        </Triggers>
                                                                    </asp:UpdatePanel>
                                                                    <span></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">
                                                                    <strong>物件編號</strong><asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator4" runat="server"
                                                                        ControlToValidate="TextBox2" ErrorMessage="請輸入物件編號"
                                                                        ForeColor="Red" SetFocusOnError="True">*</asp:RequiredFieldValidator>
                                                                </td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">

                                                                    <table bgcolor="#FFFFFF" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                                        <tr>
                                                                            <td id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel6" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:DropDownList ID="ddl契約類別" runat="server" AutoPostBack="True"
                                                                                            BackColor="#FFF79E"
                                                                                            OnSelectedIndexChanged="ddl契約類別_SelectedIndexChanged">
                                                                                            <asp:ListItem>一般</asp:ListItem>
                                                                                            <asp:ListItem>專任</asp:ListItem>
                                                                                            <asp:ListItem>流通</asp:ListItem>
                                                                                            <asp:ListItem>同意書</asp:ListItem>
                                                                                            <asp:ListItem>庫存</asp:ListItem>
                                                                                        </asp:DropDownList>
                                                                                        <asp:TextBox ID="TextBox2" runat="server" AutoPostBack="True"
                                                                                            BackColor="#FFF79E" MaxLength="8"
                                                                                            OnTextChanged="TextBox2_TextChanged" Width="83px"></asp:TextBox>&nbsp;<asp:Image ID="Image2" runat="server" ImageUrl="../images/s.png"
                                                                                                ToolTip="輸入8碼合約書編號(表單編號A開頭，請選擇一般；表單編號B開頭，請選擇專任；表單編號C開頭，請選擇同意書；若為物件用途為土地類且為專任約，請在舊表單前加上LAA不足5碼數字請在前面補0) "
                                                                                                ImageAlign="Middle" />
                                                                                        &nbsp;<asp:Label ID="Label28" runat="server" CssClass="inpage_page_font_01" Text="可用表單"></asp:Label>
                                                                                        &nbsp;<asp:DropDownList ID="DropDownList4" runat="server" AppendDataBoundItems="True"
                                                                                            AutoPostBack="True">
                                                                                        </asp:DropDownList>
                                                                                        &nbsp;<asp:Label ID="Label464" runat="server" Style="color: #FF0000; font-weight: 700; background-color: #FFCC00;" Text="消費者是否願意提供個資" Visible="False"></asp:Label>
                                                                                        <asp:RadioButton ID="RadioButton1" runat="server" GroupName="A" Style="background-color: #FFCC00" Text="同意" Visible="False" />
                                                                                        <asp:RadioButton ID="RadioButton2" runat="server" GroupName="A" Style="background-color: #FFCC00" Text="不同意" Visible="False" />
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="store" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>

                                                                            </td>
                                                                            <td align="right">
                                                                                <asp:ImageButton ID="ImageButton16" runat="server" CausesValidation="False"
                                                                                    ImageAlign="Middle" ImageUrl="../images/objectmanage_bt_20.gif" Visible="False" />&nbsp;&nbsp;</td>
                                                                        </tr>
                                                                    </table>



                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>網路案名</td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input4" runat="server" maxlength="40" name="input4"
                                                                        onmousedown="mark(this,input4)" value="網路案名最多為40個字" style="width: 600px" /><asp:Label
                                                                            ID="Label57" runat="server" Visible="False"></asp:Label>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>一般案名</strong></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <input id="Text15" runat="server" maxlength="15" name="input5"
                                                                        onmousedown="mark(this,Text15)" size="15" value="一般案名最多為15個字"
                                                                        style="width: 250px" />&nbsp;<asp:Image ID="Image3" runat="server" ImageUrl="../images/s.png"
                                                                            ToolTip="使用在張貼卡及型錄列印或其他需要簡短的案名" ImageAlign="Middle" /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">
                                                                    <strong>售價</strong><asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator3" runat="server"
                                                                        ControlToValidate="TextBox12" ErrorMessage="請輸入刊登售價"
                                                                        ForeColor="Red">*</asp:RequiredFieldValidator>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox12" runat="server" BackColor="#FFF79E" MaxLength="9"
                                                                        Width="128px" onKeyPress="javascript:JHshNumberText()" type="number" step="any"></asp:TextBox>
                                                                    &nbsp;萬<asp:Label
                                                                        ID="Label462" runat="server" Visible="False"></asp:Label>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1" class="auto-style2"><strong>車位售價</strong> </td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel25" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <input id="input55" runat="server" name="password323222" maxlength="10" />萬
                                           
                                                    <asp:CheckBox ID="CheckBox3" runat="server" AutoPostBack="True" OnCheckedChanged="CheckBox3_CheckedChanged"
                                                        Text="含於開價中" ForeColor="Red" />
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">

                                                                    <strong>物件型態</strong>
                                                                    <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="ClientValidate" ControlToValidate="DropDownList3"
                                                                        ErrorMessage="請選擇物件型態" ForeColor="Red">*</asp:CustomValidator></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList3"
                                                                        runat="server"
                                                                        BackColor="#FFF79E" AutoPostBack="True" OnSelectedIndexChanged="DropDownList3_SelectedIndexChanged">
                                                                    </asp:DropDownList>&nbsp;<asp:Image ID="Image4"
                                                                        runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="類別選擇「土地、建地」等土地類別或「透天、別墅」會在列表中顯示該物件地坪" ImageAlign="Middle" /></td>
                                                                <td bgcolor="#f7f7f7" id="td1" class="auto-style1"><strong>物件主要用途</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="DropDownList19" runat="server" AutoPostBack="True"
                                                                                OnSelectedIndexChanged="DropDownList19_SelectedIndexChanged"
                                                                                AppendDataBoundItems="True">
                                                                            </asp:DropDownList>
                                                                            <asp:TextBox ID="TextBox4" runat="server" Visible="False" Width="73px"
                                                                                MaxLength="100"></asp:TextBox>

                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <!-- 20151001 移除 -->
                                                            <%--<tr >
        <td bgcolor="#f7f7f7" id="td1">
            
            <strong>(主要)土地使用分區</strong><asp:CustomValidator ID="CustomValidator2" runat="server" 
                ClientValidationFunction="ClientValidate1" ControlToValidate="DropDownList16" 
                ErrorMessage="請選擇土地使用分區" ForeColor="Red" >*</asp:CustomValidator></td>
        <td bgcolor="#FFFFFF" id="td1">

            <asp:UpdatePanel ID="UpdatePanel31" runat="server">
            <ContentTemplate>
             <asp:DropDownList ID="DropDownList16" runat="server" 
                        AppendDataBoundItems="True" AutoPostBack="True" BackColor="#FFF79E" 
                        OnSelectedIndexChanged="DropDownList16_SelectedIndexChanged"
                       >
                    </asp:DropDownList>
                    <asp:DropDownList ID="DropDownList17" runat="server" AutoPostBack="True" 
                        visible="False">
                    </asp:DropDownList>&nbsp;<asp:Image ID="Image5" runat="server" ImageUrl="../images/s.png" 
                    ToolTip="使用分區請參照使用執照 " ImageAlign="Middle" />
            </ContentTemplate>

           
            </asp:UpdatePanel>
</td>
        <td bgcolor="#f7f7f7" id="td1"><strong>(其它)土地使用分區</strong></td>
        <td bgcolor="#FFFFFF" id="td1">
            <asp:UpdatePanel ID="UpdatePanel34" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
                    <asp:DropDownList ID="DropDownList11" runat="server" 
                        AppendDataBoundItems="True" AutoPostBack="True" 
                        OnSelectedIndexChanged="DropDownList11_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:DropDownList ID="DropDownList12" runat="server" AutoPostBack="True" 
                        visible="False">
                    </asp:DropDownList>&nbsp;
                    <asp:ImageButton ID="ImageButton18" runat="server" ToolTip="加入" 
                        ImageUrl="~/images/add_plus.png" ImageAlign="AbsMiddle" 
                        CausesValidation="False" />
                    <br>
                    <asp:TextBox ID="TextBox253" runat="server" MaxLength="500" 
                        Text="" TextMode="MultiLine" Width="250px"></asp:TextBox>&nbsp;<asp:Image ID="Image11" 
                runat="server" ImageUrl="../images/s.png" 
                      
                      ToolTip="如有一種以上的土地使用分區，第二種之後的可從此加入下方選項，內容範例:行水區,其他使用區,工二," ImageAlign="AbsMiddle" />
                </ContentTemplate>
            </asp:UpdatePanel>
          </td>
        </tr>--%>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">
                                                                    <strong>縣市鄉鎮</strong><asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator2" runat="server"
                                                                        ControlToValidate="TB_AreaCode" ErrorMessage="請選擇鄉鎮市區"
                                                                        ForeColor="Red">*</asp:RequiredFieldValidator></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox ID="TB_AreaCode" runat="server" AutoPostBack="True"
                                                                                CssClass="textfield_form_04" Width="50px" BackColor="#FFF79E"></asp:TextBox>
                                                                            <asp:DropDownList ID="DDL_County" runat="server" AppendDataBoundItems="True"
                                                                                AutoPostBack="True" BackColor="#FFF79E">
                                                                                <asp:ListItem>選擇縣市</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:DropDownList ID="DDL_Area" runat="server" AppendDataBoundItems="True"
                                                                                AutoPostBack="True" BackColor="#FFF79E">
                                                                                <asp:ListItem>選擇鄉鎮市區</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">

                                                                    <strong>地址</strong><asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator5" runat="server"
                                                                        ControlToValidate="add3" ErrorMessage="請輸入地址"
                                                                        ForeColor="Red">*</asp:RequiredFieldValidator></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel35" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox ID="add1" runat="server" MaxLength="18" Width="70px"
                                                                                Height="17px"></asp:TextBox>
                                                                            <asp:DropDownList ID="zone3" name="zone3" runat="server">
                                                                                <asp:ListItem>村</asp:ListItem>
                                                                                <asp:ListItem>里</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:TextBox ID="add2" runat="server" MaxLength="2" Width="31px"
                                                                                Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label63" runat="server" Text="鄰" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add3" runat="server" MaxLength="10" Width="70px" Height="17px"
                                                                                BackColor="#FFF79E"></asp:TextBox>
                                                                            <asp:DropDownList ID="address20" name="address20" runat="server"></asp:DropDownList>
                                                                            <asp:TextBox ID="add4" runat="server" MaxLength="2" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label64" runat="server" Text="段" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add5" runat="server" MaxLength="10" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label65" runat="server" Text="巷" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add6" runat="server" MaxLength="10" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label66" runat="server" Text="弄" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add7" runat="server" MaxLength="10" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label67" runat="server" Text="號之" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add8" runat="server" MaxLength="4" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:TextBox ID="add9" runat="server" MaxLength="50" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label68" runat="server" Text="樓之" Visible="True"></asp:Label>
                                                                            <asp:TextBox ID="add10" runat="server" MaxLength="4" Width="31px" Height="17px"></asp:TextBox>
                                                                            <asp:Label ID="Label58" runat="server" Text="棟別" Visible="False"></asp:Label>
                                                                            <asp:TextBox ID="add11" runat="server" Height="17px" MaxLength="18"
                                                                                Width="70px" Visible="False"></asp:TextBox>
                                                                            &nbsp;<strong> [<asp:LinkButton ID="LinkButton5" runat="server" Font-Bold="True" Font-Underline="True" ForeColor="Red" OnClientClick="window.open('https://easymap.land.moi.gov.tw/R02/Index#'); return false;">地籍圖資網路便民服務系統</asp:LinkButton>
                                                                            </strong>]
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger ControlID="DropDownList3"></asp:AsyncPostBackTrigger>
                                                                        </Triggers>
                                                                    </asp:UpdatePanel>

                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">

                                                                    <strong>委託期間</strong><asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator1" runat="server"
                                                                        ControlToValidate="date2" ErrorMessage="請輸入委託起始日期"
                                                                        ForeColor="Red">*</asp:RequiredFieldValidator>
                                                                    <asp:RequiredFieldValidator
                                                                        ID="RequiredFieldValidator6" runat="server"
                                                                        ControlToValidate="date3" ErrorMessage="請輸入委託截止日期"
                                                                        ForeColor="Red">*</asp:RequiredFieldValidator></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="Date2" runat="server" Width="75px" MaxLength="7"
                                                                        BackColor="#FFF79E" ToolTip="ex:1030101"></asp:TextBox>
                                                                    <input id="seeday0" runat="server" name="ch2"
                                                                        onclick="javascript: ShowCalendar(form1.Date2); Menu_OP(Calendar)" type="button"
                                                                        value="..." />
                                                                    至
            <asp:TextBox ID="Date3" runat="server" Width="75px" MaxLength="7"
                BackColor="#FFF79E" ToolTip="ex:1030101"></asp:TextBox>
                                                                    <input id="seeday1" runat="server" name="ch3"
                                                                        onclick="javascript: ShowCalendar(form1.Date3); Menu_OP(Calendar)" type="button"
                                                                        value="..." /><asp:ImageButton ID="ImageButton20" runat="server" CausesValidation="False"
                                                                            ImageAlign="Middle" ImageUrl="../images/audit_apply.gif" Visible="false" /></td>
                                                                <td bgcolor="#f7f7f7" id="td1" class="auto-style1"><strong>公設比</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox41" runat="server" MaxLength="6"
                                                                        Width="75px" onKeyPress="javascript:JHshNumberText()"></asp:TextBox>&nbsp;%</td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>銷售狀態</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList21" runat="server"
                                                                        AppendDataBoundItems="True">
                                                                        <asp:ListItem>熱賣中</asp:ListItem>
                                                                        <asp:ListItem>已停售</asp:ListItem>
                                                                        <asp:ListItem>已過期</asp:ListItem>
                                                                        <asp:ListItem>已成交</asp:ListItem>
                                                                        <asp:ListItem>已解約</asp:ListItem>
                                                                        <asp:ListItem>已終止</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1" class="auto-style1"><strong>帶看方式</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList20" runat="server"
                                                                        AppendDataBoundItems="True">
                                                                        <asp:ListItem>洽經紀人</asp:ListItem>
                                                                        <asp:ListItem>洽營業員</asp:ListItem>
                                                                        <asp:ListItem>洽管理員</asp:ListItem>
                                                                        <asp:ListItem>鑰匙</asp:ListItem>
                                                                        <asp:ListItem>密碼鎖</asp:ListItem>
                                                                        <asp:ListItem>需帶看</asp:ListItem>
                                                                        <asp:ListItem>約看</asp:ListItem>
                                                                        <asp:ListItem>自由</asp:ListItem>
                                                                        <asp:ListItem>自住</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>店代號</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="store" runat="server" AutoPostBack="true" Height="23px" OnSelectedIndexChanged="store_SelectedIndexChanged"></asp:DropDownList>
                                                                    &nbsp;<asp:CheckBox ID="all_people" runat="server"
                                                                        Text=" 顯示所有人員" AutoPostBack="True" /></td>
                                                                <td bgcolor="#f7f7f7" id="td1" class="auto-style1"><strong>輸入者</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:Label ID="Label34" runat="server"></asp:Label>
                                                                    <asp:Label ID="Label33" runat="server" Visible="False"></asp:Label>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>承辦人</strong></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel7" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            承辦人1
          <asp:DropDownList ID="sale1"
              runat="server" AppendDataBoundItems="True">
          </asp:DropDownList>
                                                                            &nbsp;承辦人2
          <asp:DropDownList ID="sale2" runat="server" AppendDataBoundItems="True">
          </asp:DropDownList>
                                                                            &nbsp;承辦人3
          <asp:DropDownList ID="sale3" runat="server" AppendDataBoundItems="True">
          </asp:DropDownList>
                                                                            <asp:Label ID="Label59" runat="server" Visible="False"></asp:Label>
                                                                            <asp:Label ID="Label60" runat="server" Visible="False"></asp:Label>
                                                                            <asp:Label ID="Label61" runat="server" Visible="False"></asp:Label>
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger ControlID="store" />
                                                                            <%--<asp:AsyncPostBackTrigger ControlID="store" EventName="SelectedIndexChanged" />--%>
                                                                        </Triggers>
                                                                    </asp:UpdatePanel>

                                                                </td>
                                                            </tr>


                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td19"><strong>格局</strong> </td>
                                                                <td bgcolor="#FFFFFF" id="td20">
                                                                    <asp:UpdatePanel ID="UpdatePanel4" runat="server" UpdateMode="Conditional">

                                                                        <ContentTemplate>
                                                                            <asp:CheckBox ID="C1" runat="server" AutoPostBack="True"
                                                                                OnCheckedChanged="C1_CheckedChanged" Text=" 開放" />&nbsp;<asp:TextBox ID="TextBox13" runat="server"
                                                                                    Width="25px" onKeyPress="javascript:JHshNumberText1()"></asp:TextBox>
                                                                            &nbsp;<asp:Label ID="Label29" runat="server" Text="房"></asp:Label>
                                                                            &nbsp;<asp:TextBox ID="TextBox14" runat="server" Width="25px" onKeyPress="javascript:JHshNumberText1()"></asp:TextBox>
                                                                            &nbsp;<asp:Label ID="Label30" runat="server" Text="廳"></asp:Label>
                                                                            &nbsp;<asp:TextBox ID="TextBox15" runat="server" Width="35px" onKeyPress="javascript:JHshNumberText()" MaxLength="3"></asp:TextBox>
                                                                            &nbsp;<asp:Label ID="Label31" runat="server" Text="衛"></asp:Label>
                                                                            &nbsp;<asp:TextBox ID="TextBox16" runat="server" Width="25px" onKeyPress="javascript:JHshNumberText1()"></asp:TextBox>
                                                                            &nbsp;<asp:Label ID="Label32" runat="server" Text="室"></asp:Label>&nbsp;<asp:Image ID="Image8" runat="server" ImageUrl="../images/s.png"
                                                                                ToolTip="格局的[房]請輸入整數，如有A+B房的狀況，請將+B房輸入至[室]" ImageAlign="Middle" />
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td21" class="auto-style1"><strong>座向</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td22">
                                                                    <asp:DropDownList ID="DropDownList22" runat="server">
                                                                        <asp:ListItem>暫未調查</asp:ListItem>
                                                                        <asp:ListItem>座北朝南</asp:ListItem>
                                                                        <asp:ListItem>座南朝北</asp:ListItem>
                                                                        <asp:ListItem>座東朝西</asp:ListItem>
                                                                        <asp:ListItem>座西朝東</asp:ListItem>
                                                                        <asp:ListItem>座東南朝西北</asp:ListItem>
                                                                        <asp:ListItem>座西南朝東北</asp:ListItem>
                                                                        <asp:ListItem>座東北朝西南</asp:ListItem>
                                                                        <asp:ListItem>座西北朝東南</asp:ListItem>
                                                                        <asp:ListItem>其它</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <%--</ContentTemplate>
            </asp:UpdatePanel>--%>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td23"><strong>樓層</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td24">地上 
            <asp:TextBox
                ID="TextBox88" runat="server"
                Width="25px" MaxLength="3" BackColor="#FFF79E"></asp:TextBox>
                                                                    &nbsp;地下 
            <asp:TextBox
                ID="TextBox89" runat="server"
                Width="25px" MaxLength="3" BackColor="#FFF79E"></asp:TextBox>
                                                                    &nbsp;所在 
            <asp:TextBox
                ID="TextBox90" runat="server"
                Width="100px" MaxLength="10" BackColor="#FFF79E"></asp:TextBox>
                                                                    <asp:Image ID="Image21" runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="當物件型態為透天，整棟販售時，所在樓層直接輸入 全" ImageAlign="Middle" />
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td25" class="auto-style1"><strong>每層戶數</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td26">
                                                                    <asp:TextBox
                                                                        ID="TextBox91" runat="server" type="number"
                                                                        Width="25px" MaxLength="2" BackColor="#FFF79E" min="0"></asp:TextBox>
                                                                    &nbsp;戶 
            <asp:TextBox
                ID="TextBox92" runat="server" type="number"
                Width="25px" MaxLength="2" BackColor="#FFF79E" min="0"></asp:TextBox>
                                                                    &nbsp;部電梯<asp:Image ID="Image22" runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="當無電梯時，直接輸入0" ImageAlign="Middle" /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td27"><strong>完工年月</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td28">
                                                                    <input id="Text2" runat="server"
                                                                        name="Text2" maxlength="5" style="width: 75px; background-color: #FFF79E;" onkeypress="javascript:JHshNumberText()" />
                                                                    <input id="ch1" runat="server" name="ch1"
                                                                        onclick="javascript: ShowCalendar(form1.Text2); Menu_OP(Calendar)" type="button"
                                                                        value="..." />&nbsp;<asp:Image ID="Image7" runat="server" ImageUrl="../images/s.png"
                                                                            ToolTip="輸入格式為yyymm，共5碼，例:09901，如無完工年月，則輸入00000" ImageAlign="Middle" /></td>
                                                                <td bgcolor="#f7f7f7" id="td29" class="auto-style1"><strong>登記日期</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td30">
                                                                    <input id="Text11" runat="server"
                                                                        name="Text2" size="10" maxlength="7" style="width: 75px" onkeypress="javascript:JHshNumberText()" />
                                                                    <input name="ch1" id="Button3" runat="server" onclick="javascript: ShowCalendar(form1.Text11); Menu_OP(Calendar)"
                                                                        type="button" value="..." /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td31"><strong>管理費</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td32">
                                                                    <asp:UpdatePanel ID="UpdatePanel22" runat="server">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True"
                                                                                OnSelectedIndexChanged="DropDownList5_SelectedIndexChanged"
                                                                                Width="50px">
                                                                                <asp:ListItem></asp:ListItem>
                                                                                <asp:ListItem>無</asp:ListItem>
                                                                                <asp:ListItem>坪</asp:ListItem>
                                                                                <asp:ListItem>月</asp:ListItem>
                                                                                <asp:ListItem>季</asp:ListItem>
                                                                                <asp:ListItem>年</asp:ListItem>
                                                                                <asp:ListItem>半年</asp:ListItem>
                                                                                <asp:ListItem>雙月</asp:ListItem>
                                                                                <asp:ListItem>未知</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:TextBox ID="TextBox36" runat="server" MaxLength="5"
                                                                                Width="80px" onKeyPress="javascript:JHshNumberText()"></asp:TextBox>
                                                                            <asp:Label ID="Label45" runat="server" Text="元"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td33" class="auto-style1"><strong>管理費繳交方式</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td34">
                                                                    <asp:TextBox ID="TextBox266" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                                                    <asp:Image ID="Image17" runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="例:月繳、半年繳、季繳、年繳或其他方式請敘明之" ImageAlign="Middle" /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td31"><strong>臨路寬</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td32">
                                                                    <asp:TextBox ID="TextBox245" runat="server"
                                                                        Width="25px"></asp:TextBox>
                                                                    米</td>
                                                                <td bgcolor="#f7f7f7" id="td33" class="auto-style1"><strong>面寬</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td34">
                                                                    <asp:TextBox ID="TextBox39" runat="server"
                                                                        Width="25px"></asp:TextBox>
                                                                    米</td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td31"><strong>縱深</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td32">
                                                                    <asp:TextBox ID="TextBox40" runat="server"
                                                                        Width="25px"></asp:TextBox>
                                                                    米</td>
                                                                <td bgcolor="#f7f7f7" id="td33" class="auto-style1"><strong>磁扣配對</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td34">
                                                                    <asp:TextBox ID="TextBox267" runat="server"></asp:TextBox>
                                                                    <asp:Image ID="Image19" runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="請將磁扣輕觸讀卡機上感應內碼" ImageAlign="Middle" /></td>
                                                            </tr>

                                                        </table>

                                                        <div style="text-align: left">
                                                            <img src="../images/dot_05.png" width="16" height="16" align="absmiddle" /><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">其他
      
                                                        </div>
                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td39">物件標的</td>
                                                                <td bgcolor="#f7f7f7" id="td40">現況</td>
                                                                <td bgcolor="#f7f7f7" id="td41">交屋情況</td>
                                                                <td bgcolor="#f7f7f7" id="td42">商談交屋情況</td>
                                                                <td bgcolor="#f7f7f7" id="td43">中庭花園</td>
                                                                <td bgcolor="#f7f7f7" id="td44">&nbsp;
                                                                </td>

                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#FFFFFF" id="td45">
                                                                    <asp:DropDownList ID="DropDownList38" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>成屋</asp:ListItem>
                                                                        <asp:ListItem>預售屋</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td46">
                                                                    <asp:DropDownList ID="DropDownList39" runat="server" AppendDataBoundItems="False">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>空房</asp:ListItem>
                                                                        <asp:ListItem>空屋</asp:ListItem>
                                                                        <asp:ListItem>自用</asp:ListItem>
                                                                        <asp:ListItem>出租</asp:ListItem>
                                                                        <asp:ListItem>施工中</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td47">
                                                                    <asp:DropDownList ID="DropDownList40" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>立即</asp:ListItem>
                                                                        <asp:ListItem>商談</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td48">
                                                                    <input id="input89" runat="server" maxlength="10"
                                                                        name="objectnumber323372226" size="20" visible="true" /></td>
                                                                <td bgcolor="#FFFFFF" id="td49">
                                                                    <asp:DropDownList ID="DropDownList41" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>無</asp:ListItem>
                                                                        <asp:ListItem>有</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td50">
                                                                    <input id="input90" runat="server" maxlength="12"
                                                                        name="objectnumber323372227" size="20" visible="true" /></td>

                                                            </tr>

                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td51">警衛管理</td>
                                                                <td bgcolor="#f7f7f7" id="td52"></td>
                                                                <td bgcolor="#f7f7f7" id="td53">外牆外飾</td>
                                                                <td bgcolor="#f7f7f7" id="td54">其他外牆外飾</td>
                                                                <td bgcolor="#f7f7f7" id="td55">地板</td>
                                                                <td bgcolor="#f7f7f7" id="td56">其他地板</td>

                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#FFFFFF" id="td57">
                                                                    <asp:DropDownList ID="DropDownList42" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>無</asp:ListItem>
                                                                        <asp:ListItem>有</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td58">
                                                                    <input id="input91" runat="server" maxlength="12"
                                                                        name="objectnumber32337222" size="20" visible="true" /></td>
                                                                <td bgcolor="#FFFFFF" id="td59">
                                                                    <asp:DropDownList ID="DropDownList43" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td60">
                                                                    <input id="input92" runat="server" maxlength="10"
                                                                        name="objectnumber323372225" size="20" visible="true" /></td>
                                                                <td bgcolor="#FFFFFF" id="td61">
                                                                    <asp:DropDownList ID="DropDownList44" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td62">
                                                                    <input id="input93" runat="server" maxlength="10"
                                                                        name="objectnumber323372228" size="20" visible="true" /></td>

                                                            </tr>

                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td63">自來水</td>
                                                                <td bgcolor="#f7f7f7" id="td64">未安裝原因</td>
                                                                <td bgcolor="#f7f7f7" id="td65">電力系統</td>
                                                                <td bgcolor="#f7f7f7" id="td66">有無獨立電錶</td>
                                                                <td bgcolor="#f7f7f7" id="td67">室內建材</td>
                                                                <td bgcolor="#f7f7f7" id="td68">隔間材料</td>

                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#FFFFFF" id="td69">
                                                                    <asp:DropDownList ID="DropDownList45" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>已安裝</asp:ListItem>
                                                                        <asp:ListItem>未安裝</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td70">
                                                                    <input id="input94" runat="server" maxlength="20"
                                                                        name="objectnumber323372222" size="20" /></td>
                                                                <td bgcolor="#FFFFFF" id="td71">
                                                                    <asp:DropDownList ID="DropDownList46" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>已安裝</asp:ListItem>
                                                                        <asp:ListItem>未安裝</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td72">
                                                                    <asp:DropDownList ID="DropDownList47" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>無</asp:ListItem>
                                                                        <asp:ListItem>有</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td73">
                                                                    <asp:UpdatePanel ID="UpdatePanel32" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                                <asp:DropDownList ID="DropDownList48" runat="server" AutoPostBack="True"
                                                                                    OnSelectedIndexChanged="DropDownList48_SelectedIndexChanged">
                                                                                </asp:DropDownList>
                                                                                <asp:TextBox ID="TextBox243" runat="server" Visible="False" Width="67px"></asp:TextBox>
                                                                            </span>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td74">
                                                                    <asp:UpdatePanel ID="UpdatePanel33" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                                <asp:DropDownList ID="DropDownList49" runat="server" AutoPostBack="True"
                                                                                    OnSelectedIndexChanged="DropDownList49_SelectedIndexChanged">
                                                                                </asp:DropDownList>
                                                                                <asp:TextBox ID="TextBox244" runat="server" Visible="False" Width="67px"></asp:TextBox>
                                                                            </span>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>

                                                            </tr>

                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td75">電話系統</td>
                                                                <td bgcolor="#f7f7f7" id="td76">未安裝原因</td>
                                                                <td bgcolor="#f7f7f7" id="td77">瓦斯系統</td>
                                                                <td bgcolor="#f7f7f7" id="td78">未安裝原因</td>
                                                                <td bgcolor="#f7f7f7" id="td79">主要建材&nbsp;<asp:Image ID="Image9" runat="server" ImageUrl="../images/s.png"
                                                                    ToolTip="原[建築結構]，因文字敘述較不恰當，故修正為[主要建材]" ImageAlign="Middle" />
                                                                    <select id="build_structure_temp">
                                                                        <option value="0">-請選擇預設主要建材-</option>
                                                                        <option value="RC(鋼筋混凝土)">RC(鋼筋混凝土)</option>
                                                                        <option value="SRC(鋼骨鋼筋混凝土)">SRC(鋼骨鋼筋混凝土)</option>
                                                                        <option value="SS(鋼骨構造)">SS(鋼骨構造)</option>
                                                                        <option value="磚造/加強磚造">磚造/加強磚造</option>
                                                                    </select>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td80" rowspan="2">&nbsp;</td>

                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#FFFFFF" id="td81">
                                                                    <asp:DropDownList ID="DropDownList50" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem Value="已安裝">有線路</asp:ListItem>
                                                                        <asp:ListItem Value="未安裝">無線路</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td82">
                                                                    <input id="input95" runat="server" maxlength="20"
                                                                        name="objectnumber323372223" size="20" /></td>
                                                                <td bgcolor="#FFFFFF" id="td83">
                                                                    <asp:DropDownList ID="DropDownList51" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>已安裝</asp:ListItem>
                                                                        <asp:ListItem>未安裝</asp:ListItem>
                                                                        <asp:ListItem>天然氣</asp:ListItem>
                                                                        <asp:ListItem>桶裝</asp:ListItem>
                                                                        <asp:ListItem>電熱器</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td84">
                                                                    <input id="input96" runat="server" maxlength="20"
                                                                        name="objectnumber323372224" size="20" /></td>
                                                                <td bgcolor="#FFFFFF" id="td85">
                                                                    <input id="input97" runat="server" maxlength="20"
                                                                        name="objectnumber323372229" size="20" /></td> 
                                                            </tr>
                                                             
                                                        </table>
                                                    </div>
                                                     
                                                    <div id="View3">
                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>商圈資訊</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox96" runat="server"
                                                                        Width="152px"></asp:TextBox>
                                                                    <input id="Button1" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=A', 500, 900)"
                                                                        type="button" value="..." />&nbsp;<asp:Image ID="Image12" runat="server" ImageUrl="../images/s.png"
                                                                            ToolTip="最多選擇8個，超過會儲存失敗!!" ImageAlign="Middle" />&nbsp;<asp:ImageButton ID="ImageButton7"
                                                                                runat="server" ImageUrl="~/images/shopkeeping_add.jpg"
                                                                                ImageAlign="Middle" target="_blank" CausesValidation="False" />
                                                                    <asp:TextBox ID="TextBox250" runat="server"
                                                                        Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>公園綠地</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox97" runat="server"
                                                                        Width="166px"></asp:TextBox>
                                                                    <input id="Button2" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=B', 500, 900)"
                                                                        type="button" value="..." />&nbsp;<asp:Image ID="Image13" runat="server" ImageUrl="../images/s.png"
                                                                            ToolTip="最多選擇8個，超過會儲存失敗!!" ImageAlign="Middle" />&nbsp;<asp:ImageButton ID="ImageButton8"
                                                                                runat="server" ImageUrl="~/images/shopkeeping_add.jpg"
                                                                                ImageAlign="Middle" CausesValidation="False" />
                                                                    <asp:TextBox ID="TextBox251" runat="server"
                                                                        Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近捷運</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel26" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="DropDownList8" runat="server"
                                                                                AutoPostBack="True" OnSelectedIndexChanged="DropDownList8_SelectedIndexChanged">
                                                                                <asp:ListItem>選擇路線</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:DropDownList ID="DropDownList9" runat="server">
                                                                                <asp:ListItem>選擇站名</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger ControlID="DDL_County" />
                                                                        </Triggers>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近公車站牌</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input60" runat="server" maxlength="20"
                                                                        name="password32325" /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近國小</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox98" runat="server"
                                                                        Width="200px"></asp:TextBox>
                                                                    <input id="Button4" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=D', 500, 900)"
                                                                        type="button" value="..." /><asp:TextBox ID="TextBox246" runat="server"
                                                                            Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近國中</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox99" runat="server"
                                                                        Width="200px"></asp:TextBox>
                                                                    <input id="Button6" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=E', 500, 900)"
                                                                        type="button" value="..." /><asp:TextBox ID="TextBox247" runat="server"
                                                                            Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近高中</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox100" runat="server"
                                                                        Width="200px"></asp:TextBox>
                                                                    <input id="Button7" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=F', 500, 900)"
                                                                        type="button" value="..." /><asp:TextBox ID="TextBox248" runat="server"
                                                                            Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鄰近大專院校</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox101" runat="server"
                                                                        Width="200px"></asp:TextBox>
                                                                    <input id="Button8" runat="server"
                                                                        onclick="GB_showCenter('功能選單', '../A_ObjectManage/School_Data.aspx?choose_type=G', 500, 900)"
                                                                        type="button" value="..." /><asp:TextBox ID="TextBox249" runat="server"
                                                                            Width="0px" BorderStyle="None" ForeColor="White"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>社區大樓</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel29" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="ddl社區大樓" runat="server">
                                                                            </asp:DropDownList>
                                                                            &nbsp;關鍵字查詢:<asp:TextBox ID="TextBox252" runat="server" MaxLength="15" Width="100px" AutoPostBack="True"></asp:TextBox>
                                                                            <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                                <asp:ImageButton ID="ImageButton22" runat="server" CausesValidation="False" ImageAlign="Middle" ImageUrl="~/images/shopkeeping_add.jpg" target="_blank" />
                                                                            </span>
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger ControlID="DDL_Area" />
                                                                            <asp:AsyncPostBackTrigger ControlID="TB_AreaCode" />
                                                                        </Triggers>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>鑰匙編號</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input66" runat="server"
                                                                        name="address32242" size="10" maxlength="10" /></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel10" runat="server">
                                                                        <ContentTemplate>
                                                                            <asp:Label ID="Label466" runat="server" Text="社區養寵物"></asp:Label>
                                                                            <asp:Image ID="Image20" runat="server" ImageUrl="../images/s.png" ToolTip="請了解該物件是否可飼養寵物(貓、狗)。" ImageAlign="Middle" />
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel45" runat="server">
                                                                        <ContentTemplate>
                                                                            <asp:RadioButton ID="RadioButton3" runat="server" AutoPostBack="True" Checked="True" GroupName="B" Text="否" />
                                                                            <asp:RadioButton ID="RadioButton4" runat="server" AutoPostBack="True" GroupName="B" Text="是" />
                                                                            <asp:CheckBox ID="CheckBox102" runat="server" Text="貓" />
                                                                            <asp:CheckBox ID="CheckBox103" runat="server" Text="狗" />
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1">&nbsp;</td>
                                                                <td bgcolor="#FFFFFF" id="td1">&nbsp;</td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>訴求重點 <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                    <asp:Button ID="Button14" runat="server" Text="AI生成" Visible="False" />
                                                                    &nbsp; <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                        <asp:Button ID="Button15" runat="server" Text="載入物件特色" Visible="False" />
                                                                </strong></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox102" runat="server" Height="150px"
                                                                        onclick="mark(this,TextBox102)" TextMode="MultiLine" Width="80%" placeholder="特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。"
                                                                        onfocus="if(this.value=='特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。'){this.value='';}"
                                                                        onblur="if(this.value==''){this.value='特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。';}"
                                                                        value="特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>備註</strong></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input79" runat="server" class="inputbox" maxlength="100"
                                                                        name="password332" size="80" />&nbsp;<asp:Image ID="Image6" runat="server" ImageUrl="../images/s.png"
                                                                            ToolTip="備註僅供後台使用，並不會在前台網站上呈現" ImageAlign="Middle" /></td>
                                                            </tr>


                                                        </table>

                                                    </div>

                                                    <%--<div id="View6">
     <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
      <tr>
        <td bgcolor="#f7f7f7" id="td1"><span style="line-height:30px; font-size:15px; color:#f96f00;font-weight:bold;"><asp:ImageButton ID="ImageButton11" 
                runat="server" ImageUrl="~/images/shopkeeping_add.jpg" 
                ImageAlign="Middle" target="_blank" CausesValidation="False" Visible="False" />
            </td>
        </tr>  
      </table>
      
      </div>--%>

                                                    <div id="View4" align="left">
                                                        <img src="../images/dot_05.png" width="16" height="16" align="absmiddle" /><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">說明書封面
     <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
         <tr>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong>物件條碼</strong></td>
             <td bgcolor="#FFFFFF" id="td1" colspan="3">
                 <asp:Image ID="Image1" runat="server" Visible="False" />
             </td>

         </tr>
         <tr>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong>報告單位</strong></td>
             <td bgcolor="#FFFFFF" id="td1">
                 <asp:UpdatePanel ID="UpdatePanel30" runat="server" UpdateMode="Conditional">
                     <ContentTemplate>
                         <input id="input2" runat="server" class="inputbox" maxlength="20"
                             name="object_name" size="30" />
                     </ContentTemplate>
                     <Triggers>
                         <asp:AsyncPostBackTrigger ControlID="store" />
                     </Triggers>
                 </asp:UpdatePanel>
             </td>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong>報告日期</strong></td>
             <td bgcolor="#FFFFFF" id="td1" colspan="2">
                 <asp:TextBox ID="Date5" runat="server" Width="75px" MaxLength="7"
                     ToolTip="ex:1030101"></asp:TextBox>
                 <input id="seeday3" runat="server" name="ch5"
                     onclick="javascript: ShowCalendar(form1.Date5); Menu_OP(Calendar)" type="button"
                     value="..." /></td>

         </tr>
         <tr>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong></strong></td>
             <td bgcolor="#FFFFFF" id="td1" colspan="3">本報告之不動產物全係依據 
             <input id="input102" runat="server" maxlength="8"
                 name="address33" style="width: 50px" />
                 市(縣)
             <input id="input5" runat="server" maxlength="8"
                 name="address332" style="width: 75px" />
                 地政事務所
            <asp:TextBox ID="Date7" runat="server" Width="75px" MaxLength="7"
                ToolTip="ex:1030101"></asp:TextBox>
                 &nbsp;<input id="date6" runat="server" name="Submit"
                     onclick="javascript: ShowCalendar(form1.Date7); Menu_OP(Calendar)" type="button"
                     value="..." />
                 日核發之謄本繕寫</td>

         </tr>

         <tr>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong>內容</strong></td>
             <td bgcolor="#FFFFFF" id="td1" colspan="3">
                 <asp:CheckBoxList ID="CheckBoxList1" runat="server" RepeatColumns="3"
                     RepeatDirection="Horizontal" Width="100%">
                     <asp:ListItem>產權調查表</asp:ListItem>
                     <asp:ListItem>物件個案調查表</asp:ListItem>
                     <asp:ListItem>物件個案照片</asp:ListItem>
                     <asp:ListItem>重要交易條件</asp:ListItem>
                     <asp:ListItem>成交行情參考表</asp:ListItem>
                     <asp:ListItem>鄰近重要設施參考表</asp:ListItem>
                     <asp:ListItem>其他說明</asp:ListItem>
                 </asp:CheckBoxList>
             </td>


         </tr>

         <tr>
             <td bgcolor="#f7f7f7" id="td1">
                 <strong>附件</strong></td>
             <td bgcolor="#FFFFFF" id="td1" colspan="3">
                 <asp:CheckBoxList ID="CheckBoxList2" runat="server" RepeatColumns="3"
                     RepeatDirection="Horizontal" Width="85%">
                     <asp:ListItem>土地權狀(影本)</asp:ListItem>
                     <asp:ListItem>建物權狀(影本)</asp:ListItem>
                     <asp:ListItem>標的現況說明書</asp:ListItem>
                     <asp:ListItem>土地謄本</asp:ListItem>
                     <asp:ListItem>建物謄本</asp:ListItem>
                     <asp:ListItem>委託銷售契約書</asp:ListItem>

                     <asp:ListItem>土地地籍圖</asp:ListItem>
                     <asp:ListItem>建物勘測成果圖</asp:ListItem>

                     <asp:ListItem>住戶規約</asp:ListItem>

                     <asp:ListItem>土地相關位置略圖</asp:ListItem>
                     <asp:ListItem>建物相關位置略圖</asp:ListItem>

                     <asp:ListItem>停車位位置圖</asp:ListItem>

                     <asp:ListItem>土地分管協議</asp:ListItem>
                     <asp:ListItem>建物分管協議</asp:ListItem>
                     <asp:ListItem>樑柱顯見裂痕照片</asp:ListItem>
                     <asp:ListItem>土地分區種類證明</asp:ListItem>
                     <asp:ListItem>房屋稅籍相關證明</asp:ListItem>
                     <asp:ListItem>其他</asp:ListItem>

                     <asp:ListItem>土地增值稅概算表</asp:ListItem>

                     <asp:ListItem>使用執照(影本)</asp:ListItem>
                 </asp:CheckBoxList>
             </td>


         </tr>

     </table>
                                                            <img src="../images/dot_05.png" width="16" height="16" align="absmiddle" /><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">重要交易條件</span>
                                                            <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1">
                                                                        <strong>付款方式</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1" colspan="7">&nbsp;
                                                                    </td>

                                                                </tr>
                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td87" rowspan="2">
                                                                        <strong>1.第一期</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td88">
                                                                        <asp:TextBox ID="TextBox262" runat="server" Width="75px"></asp:TextBox>&nbsp;萬元 
            
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td89" rowspan="2">
                                                                        <strong>2.第二期</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td90">
                                                                        <asp:TextBox ID="TextBox263" runat="server" Width="75px"></asp:TextBox>
                                                                        &nbsp;萬元
           
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td91" rowspan="2">
                                                                        <strong>3.第三期</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td92">
                                                                        <asp:TextBox
                                                                            ID="TextBox264" runat="server" Width="75px"></asp:TextBox>
                                                                        &nbsp;萬元
            
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td93" rowspan="2">
                                                                        <strong>4.第四期</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td94">
                                                                        <asp:TextBox
                                                                            ID="TextBox265" runat="server" Width="75px"></asp:TextBox>
                                                                        &nbsp;萬元
           
                                                                    </td>

                                                                </tr>


                                                                <tr>
                                                                    <td bgcolor="#FFFFFF" id="td95">
                                                                        <asp:TextBox ID="TextBox258" runat="server" Width="75px" onchange="checkPeriodMoney();"></asp:TextBox>
                                                                        %</td>
                                                                    <td bgcolor="#FFFFFF" id="td96">
                                                                        <asp:TextBox ID="TextBox259" runat="server" Width="75px" onchange="checkPeriodMoney();"></asp:TextBox>
                                                                        %</td>
                                                                    <td bgcolor="#FFFFFF" id="td97">
                                                                        <asp:TextBox ID="TextBox260" runat="server" Width="75px" onchange="checkPeriodMoney();"></asp:TextBox>
                                                                        %</td>
                                                                    <td bgcolor="#FFFFFF" id="td98">
                                                                        <asp:TextBox ID="TextBox261" runat="server" Width="75px" onchange="checkPeriodMoney();"></asp:TextBox>
                                                                        %</td>
                                                                </tr>




                                                            </table>
                                                            </br>
           <%--<asp:AsyncPostBackTrigger ControlID="store" EventName="SelectedIndexChanged" />--%>      <%-- <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
             <tr>
        <td bgcolor="#f7f7f7" id="td1">
           <strong>應納稅額</strong></td>
        <td bgcolor="#FFFFFF" id="td1" colspan="5">
                <asp:ImageButton ID="ImageButton11" runat="server" 
                    ImageUrl="~/images/objectmanage_bt_21.gif" CausesValidation="False" />
                 </td>



        </tr>
         <tr>
        <td bgcolor="#f7f7f7" id="td1">
            <strong>土地增值稅約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <strong>自用</strong>
            <input id="input103" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2236" size="12" /> 元<br />
            一般
            <input id="input104" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber323372230" size="12" /> 元</td>
        <td bgcolor="#f7f7f7" id="td1">
              <strong>契稅約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" > <input id="input105" runat="server" class="inputbox" maxlength="6" 
                name="objectnumber2238" size="12" /> 元</td>
         <td bgcolor="#f7f7f7" id="td1">
               <strong>地價稅約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" ><input id="input106" runat="server" 
                class="inputbox" maxlength="10" 
                name="objectnumber2236" size="12" /> 元
           </td>


        </tr>
        <tr>
        <td bgcolor="#f7f7f7" id="td1">
            房屋稅<strong>約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input107" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2236" size="12" /> 元</td>
        <td bgcolor="#f7f7f7" id="td1">
              工程受益費<strong>約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input108" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2237" size="12" /> 元</td>
         <td bgcolor="#f7f7f7" id="td1">
               所有權移轉代辦費<strong>約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input109" runat="server" class="inputbox" maxlength="6" 
                name="objectnumber2238" size="12" /> 元</td>


        </tr>
        <tr>
        <td bgcolor="#f7f7f7" id="td1">
            <strong>登記規費約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input110" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2236" size="12" /> 元</td>
        <td bgcolor="#f7f7f7" id="td1">
              公(監)<strong>證費約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input111" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2237" size="12" /> 元</td>
         <td bgcolor="#f7f7f7" id="td1">
               印花稅<strong>約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input112" runat="server" class="inputbox" maxlength="6" 
                name="objectnumber2238" size="12" /> 元</td>


        </tr>
        <tr>
        <td bgcolor="#f7f7f7" id="td1">
            <strong>水電費約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input113" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2236" size="12" /> 元</td>
        <td bgcolor="#f7f7f7" id="td1">
              <strong>瓦斯費約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
           <input id="input116" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2237" size="12" /> 元</td>
         <td bgcolor="#f7f7f7" id="td1">
               <strong>電話費約</strong></td>
        <td bgcolor="#FFFFFF" id="td1" >
            <input id="input115" runat="server" class="inputbox" maxlength="6" 
                name="objectnumber2238" size="12" /> 元</td>


        </tr>
        <tr>
        <td bgcolor="#f7f7f7" id="td1"><strong>管理費約<span style="line-height:30px; font-size:15px; color:#f96f00;font-weight:bold;"><asp:Image 
                ID="Image16" runat="server" ImageUrl="../images/s.png" 
                    ToolTip="大樓管理費及車位管理費合計" ImageAlign="Middle" /></strong></td>
        <td bgcolor="#FFFFFF" id="td1" > <input id="input114" runat="server" class="inputbox" maxlength="10" 
                name="objectnumber2237" size="12" /> 元</td>
        <td bgcolor="#f7f7f7" id="td1"></td>
        <td bgcolor="#FFFFFF" id="td1" ></td>
         <td bgcolor="#f7f7f7" id="td1">
               &nbsp;</td>
        <td bgcolor="#FFFFFF" id="td1" >
            &nbsp;</td>


        </tr>
         </table>--%>


                                                            <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">

                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>增值稅</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList52" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>契稅</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList53" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>地價稅</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList54" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>

                                                                </tr>
                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>房屋稅</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList55" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>工程受益費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList56" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem Selected="True">另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>所有權移轉代辦費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList57" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>

                                                                </tr>
                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>登記規費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList58" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>公(監)證費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList59" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem Selected="True">另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>印花稅</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList60" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>

                                                                </tr>

                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>水電瓦斯電話費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList61" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>鑑界費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList13" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList></td>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>履保費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList63" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                </tr>

                                                                <tr>
                                                                    <td bgcolor="#f7f7f7" id="td1"><strong>管理費</strong></td>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <asp:DropDownList ID="DropDownList62" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem>買方負擔</asp:ListItem>
                                                                            <asp:ListItem Selected="True">賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList></td>
                                                                    <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                        <td bgcolor="#f7f7f7" id="td1"><strong>實價登錄費</strong></td>
                                                                    </span>
                                                                    <td bgcolor="#FFFFFF" id="td1"><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                        <asp:DropDownList ID="DropDownList70" runat="server" CssClass="inputbox">
                                                                            <asp:ListItem></asp:ListItem>
                                                                            <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                            <asp:ListItem>賣方負擔</asp:ListItem>
                                                                            <asp:ListItem>雙方各半</asp:ListItem>
                                                                            <asp:ListItem>另以契約約定</asp:ListItem>
                                                                            <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                        <td bgcolor="#f7f7f7" id="td1"><strong>代書費</strong></td>
                                                                    </span>
                                                                    <td bgcolor="#FFFFFF" id="td1">
                                                                        <span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                            <asp:DropDownList ID="DropDownList71" runat="server" CssClass="inputbox">
                                                                                <asp:ListItem></asp:ListItem>
                                                                                <asp:ListItem Selected="True">買方負擔</asp:ListItem>
                                                                                <asp:ListItem>賣方負擔</asp:ListItem>
                                                                                <asp:ListItem>雙方各半</asp:ListItem>
                                                                                <asp:ListItem>另以契約約定</asp:ListItem>
                                                                                <asp:ListItem>依交屋日分算</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                    </td>
                                                                </tr>
                                                            </table>

                                                            <br />
                                                            <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                                <tr>
                                                                    <td>※參考:<a href="fine.pdf" target="_blank">罰則違反不動產經紀業管理條例[pdf]</a></td>
                                                                </tr>
                                                            </table>
                                                    </div>

                                                    <div id="View5" align="left">
                                                        <asp:Literal ID="content" runat="server" Text=""></asp:Literal>

                                                    </div>

                                                    <div id="View2">
                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <tr>
                                                                <%if show_new = "1" %>
                                                                <td bgcolor="#FFFFFF" id="td1" colspan="4"><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold"><a href="#">
                                                                    <img id="ImageButton23" align="absmiddle" height="23" src="../images/objectmanage_bt_22.gif" width="92" /></a>&nbsp;
                                                                </span><span><a href="#">
                                                                    <img id="ImageButton21" align="absmiddle" height="23" src="../images/objectmanage_bt_22.gif" width="92" /></a>
                                                                    <asp:ImageButton ID="ImageButton11" runat="server"
                                                                        ImageUrl="~/images/objectmanage_bt_21.gif" CausesValidation="False" />
                                                                </span></td>
                                                            </tr>
                                                            <%end if%>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(A)主建物</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel12" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox ID="TextBox5" runat="server" AutoPostBack="True" Width="75px"
                                                                                OnTextChanged="TextBox5_TextChanged" MaxLength="9" BorderStyle="None"
                                                                                ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label3" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox6" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label7" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(B)附屬建物</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel13" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox7_TextChanged" ID="TextBox7" runat="server"
                                                                                AutoPostBack="True" Width="75px" MaxLength="9" BorderStyle="None"
                                                                                ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label15" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox8" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label1" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(C)公共設施</strong>
                                                                    &nbsp;&nbsp;&nbsp;<input id="Checkbox27" type="checkbox" name="Checkbox27" onchange="DisablePublicCar();" runat="server" />
                                                                    <label for="Checkbox27">含公設車位坪數</label>
                                                                    <input id="hid_PublicCar" type="hidden" runat="server" value="N" />
                                                                    <asp:Image ID="Image18"
                                                                        runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="賣方確有車位產權, 但謄本中無標註車位持分, 致無法計算時方可勾選" ImageAlign="AbsMiddle" />
                                                                </td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel14" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox9_TextChanged" ID="TextBox9" runat="server"
                                                                                AutoPostBack="True" Width="75px" MaxLength="9" BorderStyle="None"
                                                                                ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label2" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox10" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label8" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(D)地下室</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel15" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox19_TextChanged" ID="TextBox19"
                                                                                runat="server" AutoPostBack="True" Width="75px" MaxLength="9"
                                                                                BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label5" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox20" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label9" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1" class="tdPublicCar"><span id="txtPublicCar" style="font-weight: bold">
                                                                    <asp:Label ID="Label77" runat="server" Text="(E)車位面積"></asp:Label></span><asp:Image ID="Image5"
                                                                        runat="server" ImageUrl="../images/s.png"
                                                                        ToolTip="要判別停車空間是屬於何種類型，無法從登記機關建物登記簿謄本或建物測量成果圖分辦，而應從使用執照所附竣工平面圖上所標示之位置判別，是法定停車空間、自行增設停車空間或是獎勵增設停車空間，另竣工平面圖上以顏色繪出，法定停車空間以黃色標線繪出、自行增設停車空間以橙色標線繪出，獎勵增設停車空間以紅色標線繪出。" ImageAlign="AbsMiddle" /></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1" class="tdPublicCar">
                                                                    <asp:UpdatePanel ID="UpdatePanel16" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox21_TextChanged" ID="TextBox21"
                                                                                runat="server" AutoPostBack="True" Width="75px" MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label16" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox23" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label6" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>
                                                                    <asp:Label ID="Label78" runat="server" Text="(F)產權車位" Visible="false"></asp:Label></strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel17" runat="server" UpdateMode="Conditional" Visible="false">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox26_TextChanged" ID="TextBox26"
                                                                                runat="server" Width="75px" AutoPostBack="True" MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label17" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox27" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label13" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(G)土地坪數</strong></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel18" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox30_TextChanged" ID="TextBox30"
                                                                                runat="server" Width="75px" AutoPostBack="True" MaxLength="9"
                                                                                BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label10" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox31" runat="server" Width="75px"
                                                                                MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label19" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>(H)總坪數(H=A+B+C+D+E)</strong><%--<asp:RequiredFieldValidator 
                ID="RequiredFieldValidator7" runat="server" 
                        ControlToValidate="TextBox28" ErrorMessage="請輸入總坪數平方公尺" 
                ForeColor="Red">*</asp:RequiredFieldValidator>--%></td>
                                                                <td colspan="3" bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel19" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox28_TextChanged" ID="TextBox28"
                                                                                runat="server" Width="75px" MaxLength="9" AutoPostBack="True"
                                                                                BackColor="#FFF79E" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label18" runat="server" Text="平方公尺"></asp:Label>
                                                                            <asp:TextBox ID="TextBox29" runat="server" Width="75px"
                                                                                MaxLength="9" AutoPostBack="True" BackColor="#FFF79E" BorderStyle="None"
                                                                                ForeColor="#0084D7"></asp:TextBox>
                                                                            <asp:Label ID="Label14" runat="server" Text="坪"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>庭院坪數</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel20" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox32_TextChanged" ID="TextBox32"
                                                                                runat="server" Width="75px" MaxLength="9"
                                                                                AutoPostBack="True" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            平方公尺<asp:TextBox ID="TextBox33" runat="server"
                                                                                Width="75px" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            坪 
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>增建</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:RadioButtonList ID="RadioButtonList3" runat="server"
                                                                        RepeatDirection="Horizontal">
                                                                        <asp:ListItem Value="Y">有</asp:ListItem>
                                                                        <asp:ListItem Selected="True" Value="N">無</asp:ListItem>
                                                                    </asp:RadioButtonList>
                                                                    <asp:UpdatePanel ID="UpdatePanel21" runat="server" UpdateMode="Conditional"
                                                                        Visible="False">
                                                                        <ContentTemplate>
                                                                            <asp:TextBox OnTextChanged="TextBox34_TextChanged" ID="TextBox34"
                                                                                runat="server" Width="75px" AutoPostBack="True" BorderStyle="None"
                                                                                ForeColor="#0084D7"></asp:TextBox>
                                                                            平方公尺<asp:TextBox ID="TextBox35" runat="server"
                                                                                CssClass="inputbox" Width="75px" MaxLength="9" BorderStyle="None" ForeColor="#0084D7"></asp:TextBox>
                                                                            坪      
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                            <tr>

                                                                <td bgcolor="#f7f7f7" id="td1"><strong>土地標示</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox17" runat="server" Width="250px"
                                                                        onmousedown="mark(this,TextBox17)" TextMode="MultiLine" MaxLength="500">請填寫地號</asp:TextBox>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>建物標示</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox18" runat="server" Width="250px"
                                                                        onmousedown="mark(this,TextBox18)" TextMode="MultiLine" MaxLength="500">請填寫建號</asp:TextBox>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        </br>
      </br>
                                                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                            <tr>
                                                                <td>
                                                                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                                        <tr>
                                                                            <td align="left">
                                                                                <div id="inpage_content_title">標示部</div>
                                                                            </td>
                                                                            <td align="right">
                                                                                <font style="color: red">*提醒您新增面積細項前先按下方確定新增按鈕存檔,確定有物件編號後再行新增</font>
                                                                                <asp:ImageButton ID="ImageButton17" runat="server"
                                                                                    ImageUrl="~/images/objectmanage_bt_21.gif" CausesValidation="False" Visible="False" />

                                                                                <asp:ImageButton ID="ImageButton5" runat="server"
                                                                                    ImageUrl="~/images/shopkeeping_add.jpg" CausesValidation="False" />
                                                                                <asp:ImageButton ID="ImageButton3" runat="server"
                                                                                    ImageUrl="~/images/shopkeeping_edit_03.jpg" CausesValidation="False" />
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">

                                                                        <tr>
                                                                            <td align="center" bgcolor="#f7f7f7">
                                                                                <strong>類別名稱</strong></td>
                                                                            <td align="center" bgcolor="#f7f7f7">
                                                                                <strong>項目名稱</strong></td>
                                                                            <td align="center" bgcolor="#f7f7f7">
                                                                                <strong>總面積(㎡)</strong></td>
                                                                            <td align="center" bgcolor="#f7f7f7" colspan="4">
                                                                                <strong>總面積(坪)</strong></td>



                                                                        </tr>
                                                                        <tr>
                                                                            <td align="left" bgcolor="#FFFFFF">
                                                                                <asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True">
                                                                                </asp:DropDownList>
                                                                                <asp:Label ID="Label4" runat="server" Visible="False">0</asp:Label>

                                                                                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional" RenderMode="Inline">
                                                                                    <ContentTemplate>
                                                                                        <%--<asp:CheckBox ID="CheckBox28" runat="server" Text="是否為公設" Visible="false"    />--%>
                                                                                        <asp:DropDownList ID="DDL_level2" runat="server" AutoPostBack="True" Visible="false" OnSelectedIndexChanged="DDL_level2_SelectedIndexChanged">
                                                                                            <asp:ListItem Value="主建物','附屬物">層次</asp:ListItem>
                                                                                            <asp:ListItem Value="附屬物">附屬建物</asp:ListItem>
                                                                                            <asp:ListItem>共有部分</asp:ListItem>
                                                                                        </asp:DropDownList>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />

                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                                <asp:Label ID="Label76" runat="server" Text="類別名稱為：土地面積、主建物、車位面積(含公設內、產權獨立)者不可修改，欲更改請刪除重建。" Visible="False" ForeColor="#CC3300" Font-Size="Small"></asp:Label>

                                                                            </td>
                                                                            <td align="center" bgcolor="#FFFFFF">
                                                                                <asp:UpdatePanel ID="UpdatePanel8" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>


                                                                                        <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True">
                                                                                        </asp:DropDownList>
                                                                                        <asp:TextBox ID="TextBox11" runat="server" Visible="False" Width="50px"></asp:TextBox>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DDL_level2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td align="center" bgcolor="#FFFFFF">
                                                                                <asp:TextBox ID="TextBox77" runat="server" AutoPostBack="True" Width="50px"></asp:TextBox>
                                                                            </td>
                                                                            <td align="center" bgcolor="#FFFFFF" colspan="4">
                                                                                <asp:UpdatePanel ID="UpdatePanel9" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:TextBox ID="TextBox73" runat="server" Width="50px"></asp:TextBox>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="TextBox77" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>


                                                                            </td>


                                                                        </tr>
                                                                        <tr>
                                                                            <td align="center" bgcolor="#f7f7f7" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel37" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:Label ID="Label69" runat="server" Font-Bold="True">建號</asp:Label>
                                                                                        &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="CheckBox98" Text="本建號為公設" runat="server" onclick="CheckCheckboxes(98);" Visible="False"></asp:CheckBox>
                                                                                        &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="CheckBox99" Text="本建號為車位" runat="server" onclick="CheckCheckboxes(99);" Visible="False"></asp:CheckBox>



                                                                                        <asp:TextBox ID="TextBox22" runat="server" Visible="false"></asp:TextBox>
                                                                                        <asp:TextBox ID="TextBox24" runat="server" Visible="false"></asp:TextBox>
                                                                                        <asp:Label runat="server" ID="lb_細項流水" Visible="false"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td colspan="6" align="left" bgcolor="#FFFFFF" style="padding-left: 4px">
                                                                                <asp:TextBox ID="TextBox25" runat="server" Width="90%" MaxLength="30"></asp:TextBox>&nbsp;
                    <asp:Image ID="Image14" runat="server" ImageUrl="../images/s.png"
                        ToolTip="字數上限為30個字!!" ImageAlign="AbsMiddle" />
                                                                            </td>

                                                                        </tr>

                                                                        <tr>
                                                                            <td bgcolor="#f7f7f7" align="center" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel38" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:Label ID="Label70" runat="server" Font-Bold="True"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DDL_level2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td bgcolor="#FFFFFF" colspan="3" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel36" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:TextBox ID="TextBox1" runat="server" AutoPostBack="True" ToolTip="分母" placeholder="分母" Visible="false"></asp:TextBox>
                                                                                        <asp:Label ID="Label26" runat="server" Text="分之" Visible="false"></asp:Label>
                                                                                        <asp:TextBox ID="TextBox3" runat="server" AutoPostBack="True" ToolTip="分子" placeholder="分子" Visible="false"></asp:TextBox><br />
                                                                                        <asp:DropDownList ID="DropDownList69" runat="server" Visible="False">
                                                                                            <asp:ListItem>請選擇</asp:ListItem>
                                                                                            <asp:ListItem>都市</asp:ListItem>
                                                                                            <asp:ListItem>非都市</asp:ListItem>
                                                                                        </asp:DropDownList>
                                                                                        <asp:DropDownList ID="DropDownList65" runat="server"
                                                                                            AppendDataBoundItems="True" AutoPostBack="True"
                                                                                            OnSelectedIndexChanged="DropDownList65_SelectedIndexChanged"
                                                                                            Visible="False">
                                                                                        </asp:DropDownList>
                                                                                        <asp:DropDownList ID="DropDownList66" runat="server"
                                                                                            Visible="False">
                                                                                        </asp:DropDownList>
                                                                                        <asp:TextBox ID="TextBox254" runat="server" MaxLength="10" Visible="False"></asp:TextBox>
                                                                                        <asp:Image ID="Image15" runat="server" ImageAlign="AbsMiddle"
                                                                                            ImageUrl="../images/s.png" ToolTip="字數上限為10個字!!" Visible="False" />
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DDL_level2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>

                                                                            </td>
                                                                            <td bgcolor="#f7f7f7" align="center" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel39" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:Label ID="Label71" runat="server"
                                                                                            Font-Bold="True"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td bgcolor="#FFFFFF" colspan="2" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel40" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:TextBox ID="Date8" runat="server" MaxLength="7" ToolTip="ex:1030101"
                                                                                            Width="75px" Visible="False"></asp:TextBox>
                                                                                        <input id="seeday4" runat="server" name="ch6"
                                                                                            onclick="javascript: ShowCalendar(form1.Date8); Menu_OP(Calendar)" type="button"
                                                                                            value="..." visible="False" />
                                                                                        <asp:TextBox ID="TextBox255" runat="server" Width="75px" Visible="False"></asp:TextBox>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>

                                                                        </tr>






                                                                        <tr>
                                                                            <td bgcolor="#f7f7f7" align="center" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel41" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:Label ID="Label72" runat="server" Font-Bold="True"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DDL_level2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td bgcolor="#FFFFFF" colspan="3" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel42" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:TextBox ID="TextBox44" runat="server" AutoPostBack="True" ToolTip="分母" placeholder="分母" Visible="false"></asp:TextBox>
                                                                                        <asp:Label ID="Label46" runat="server" Text="分之" Visible="false"></asp:Label>
                                                                                        <asp:TextBox ID="TextBox45" runat="server" AutoPostBack="True" ToolTip="分子" placeholder="分子" Visible="false"></asp:TextBox>


                                                                                        <asp:TextBox ID="TextBox256" runat="server" MaxLength="10" Width="50px"
                                                                                            Visible="False"></asp:TextBox>
                                                                                        <asp:Label ID="Label74" runat="server" Text="%" Visible="False"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DDL_level2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>

                                                                            </td>
                                                                            <td bgcolor="#f7f7f7" align="center" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel43" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:Label ID="Label73" runat="server"
                                                                                            Font-Bold="True"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                            <td bgcolor="#FFFFFF" colspan="2" id="td1">
                                                                                <asp:UpdatePanel ID="UpdatePanel44" runat="server" UpdateMode="Conditional">
                                                                                    <ContentTemplate>
                                                                                        <asp:TextBox ID="TextBox257" runat="server"
                                                                                            MaxLength="10" Width="50px" Visible="False"></asp:TextBox>
                                                                                        <asp:Label ID="Label75" runat="server" Text="%" Visible="False"></asp:Label>
                                                                                    </ContentTemplate>
                                                                                    <Triggers>
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList1" />
                                                                                        <asp:AsyncPostBackTrigger ControlID="DropDownList2" />
                                                                                    </Triggers>
                                                                                </asp:UpdatePanel>
                                                                            </td>

                                                                        </tr>






                                                                        <tr>
                                                                            <td align="left">
                                                                                <div style="width: 200px; height: 31px; background: url(../images/title_bg.png); padding-top: 4px; font-size: 15px; color: #FFFFFF; font-weight: bold; text-align: center">所有權部</div>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td colspan="7">
                                                                                <!-- HTML Table -->
                                                                                <asp:Literal ID="editTable" runat="server">
               <table cellspacing="0" rules="all" border="1" id="Table_right" style="width:100%;border-collapse:collapse;"  >
               
               <tr>
                   <td colspan="7" >權利人部分
                         
                   </td>
               </tr>
		<tr>
			<th scope="col">&nbsp;</th>
            <th scope="col">姓名</th>
            <th scope="col">所有權部權利範圍</th>
            <th scope="col">面積</th>
            <th scope="col">出售權利種類</th>
            <th scope="col">所有權種類</th>
            <th scope="col"></th> 
		</tr>
        <tr align="center">
			<td style="width:30px;"><span id="tb_num">1</span></td>
            <td style="width:55px;">
                 <input type="text" id="tb_Textbox50" style="width:50px;" />
            </td>
            <td align="left"  >
                權利範圍.<input type="text" id="tb_Numerator_1_0" title="分母" style="width:50px;" />&nbsp;分之 &nbsp;
                  <input type="text" id="tb_Denominator_1_0" title="分子" style="width:50px;" onblur='calculatorarea(this.id);'   />
                                  <br />
                出售權利範圍.<input type="text"  id="tb_Numerator_2_0" title="分母" style="width:50px;" />&nbsp;分之 &nbsp;
                    <input type="text" id="tb_Denominator_2_0" title="分子" style="width:50px;" onblur='calculatorarea2(this.id);' />
                              
            </td>
            <td>
               持有面積: <span id="tb_Label21_1_0"></span><br />
               出售面積: <span id="tb_Label21_2_0"></span>
            </td>
            <td>
                <select  id="tb_DropDownList70_0">
				    <option selected="selected" value="所有權">所有權</option>
				    <option value="地上權">地上權</option>
				    <option value="典權">典權</option>
				    <option value="使用權">使用權</option>

			&nbsp;&nbsp;&nbsp; </select>
                           </td>
            <td>
                <select  id="tb_DropDownList71_0">
				    <option value="單獨所有">單獨所有</option>
				    <option selected="selected" value="分別共有">分別共有</option>
				    <option value="公同共有">公同共有</option>

			    </select>
                               
           </td>
            <td>
                <input id="Button11" type="button" value="儲存" onclick="SaveRow(this)" />
                <input id="Button12" type="button" value="刪除" onclick="DelRow(this)" />
                <br /><span id="tempcode"  style="font-size :3px ;color:Gray; line-height :6px; "  ></span>
            </td>
		</tr>
       
	</table>
	           
                                                                                </asp:Literal>




                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td colspan="7">
                                                                                <input id="Button10" type="button" value="新增一筆權利人資料" onclick="AddOneRowTb()" />
                                                                            </td>
                                                                        </tr>



                                                                    </table>

                                                                    <asp:Label ID="Label12" runat="server" Visible="False"></asp:Label>
                                                                    <asp:Label ID="Label11" runat="server" Visible="False"></asp:Label>
                                                                </td>
                                                            </tr>
                                                        </table>

                                                        </br>
        
       <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">

           <tr>
               <td bgcolor="#f7f7f7" id="td1" colspan="8">
                   <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
                       BorderStyle="None" GridLines="None" CssClass="GridViewStyle"
                       CellPadding="3" Width="100%">
                       <PagerStyle Height="8px" HorizontalAlign="Center" BorderStyle="None" />
                       <RowStyle CssClass="row" />
                       <Columns>
                           <asp:TemplateField HeaderText="刪除">

                               <HeaderTemplate>
                                   <input id="chkAll" name="chkAll" onclick="Check(this, 'chkSelect1', 'GridView1')" type="checkbox" />
                                   刪除
                               </HeaderTemplate>
                               <ItemTemplate>
                                   &nbsp;<asp:CheckBox ID="chkSelect1" runat="server" />
                                   <asp:Label ID="Label35" runat="server" Text='<%# Bind("物件編號") %>'
                                       Visible="False"></asp:Label>
                                   <asp:Label ID="Label36" runat="server" Text='<%# Bind("流水號") %>'
                                       Visible="False"></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField>
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox79" runat="server"></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Button ID="Button5" runat="server" CommandName="edits" Text="編輯"
                                       CausesValidation="False" />
                                   <br />
                                   <asp:Button ID="Button6" runat="server" CausesValidation="False"
                                       CommandName="check" Text="調查" />
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="類別" ItemStyle-HorizontalAlign="Left">

                               <ItemTemplate>
                                   <asp:Label ID="Label37" runat="server" Text='<%# Bind("類別") %>'></asp:Label>
                                   <asp:Label ID="Label24" runat="server" Text='<%# Bind("DL_level2_selectindex")%>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>



                           <asp:TemplateField HeaderText="項目">
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox81" runat="server" Text='<%# Bind("項目名稱") %>'></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Label ID="Label38" runat="server" Text='<%# Bind("項目名稱") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="建號/地號">
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox82" runat="server" Text='<%# Bind("建號") %>'></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Label ID="Label39" runat="server" Text='<%# Bind("建號") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="總面積(㎡)">
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox83" runat="server" Text='<%# Bind("總面積平方公尺") %>'></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Label ID="Label40" runat="server" Text='<%# Bind("總面積平方公尺") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="總面積(坪)">

                               <ItemTemplate>
                                   <asp:Label ID="Label41" runat="server" Text='<%# Bind("總面積坪") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="權利範圍">

                               <ItemTemplate>
                                   <asp:Label ID="Label42" runat="server" Text='<%# Bind("權利範圍") %>'></asp:Label>
                                   <br />
                                   <asp:Label ID="Label20" runat="server" Text='<%# Bind("權利範圍2") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="持有面積(㎡/坪)">
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox86" runat="server" Text='<%# Bind("實際持有平方公尺") %>'></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Label ID="Label43" runat="server" Text='<%# Bind("實際持有平方公尺") %>'></asp:Label>/<asp:Label ID="Label44" runat="server" Text='<%# Bind("實際持有坪") %>'></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="出售面積(㎡/坪)">
                               <EditItemTemplate>
                                   <asp:TextBox ID="TextBox87" runat="server" Text=""></asp:TextBox>
                               </EditItemTemplate>
                               <ItemTemplate>
                                   <asp:Label ID="Label88" runat="server" Text=""></asp:Label>/<asp:Label ID="Label89" runat="server" Text=""></asp:Label>
                               </ItemTemplate>
                               <HeaderStyle CssClass="GridViewHeaderStyle" HorizontalAlign="Center" />
                               <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                           </asp:TemplateField>

                           <asp:TemplateField Visible="false">
                               <ItemTemplate>
                                   <asp:Label ID="Label98" runat="server" Text='<%# Bind("是否為公設") %>'></asp:Label>
                                   <asp:Label ID="Label99" runat="server" Text='<%# Bind("是否為車位") %>'></asp:Label>
                               </ItemTemplate>
                           </asp:TemplateField>
                       </Columns>

                   </asp:GridView>
               </td>
           </tr>
           <tr>
               <td bgcolor="#f7f7f7" id="td1" colspan="8" align="right">
                   <asp:ImageButton ID="ImageButton2" runat="server"
                       ImageUrl="~/images/shopkeeping_del.jpg" CausesValidation="False"
                       Visible="False" />
               </td>
           </tr>

       </table>

                                                        <!-- start here -->
                                                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                            <tr>
                                                                <td>
                                                                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                                        <tr>
                                                                            <td align="left">
                                                                                <div id="inpage_content_title">他項權利部</div>
                                                                            </td>
                                                                            <td align="right"><span style="line-height: 30px; font-size: 15px; color: #f96f00; font-weight: bold;">
                                                                                <asp:ImageButton ID="ImageButton4" runat="server"
                                                                                    ImageUrl="~/images/shopkeeping_add.jpg" CausesValidation="False" />
                                                                                <asp:ImageButton ID="ImageButton9" runat="server"
                                                                                    ImageUrl="~/images/shopkeeping_edit_03.jpg" CausesValidation="False" /></td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">



                                                                        <tr>
                                                                            <td id="td2" align="left" bgcolor="#f7f7f7" colspan="7">
                                                                                <asp:CheckBox ID="CheckBox24" runat="server" Text="建物與土地他項權利部同" />
                                                                                &nbsp; &nbsp;
                    <asp:CheckBox ID="CheckBox25" runat="server" Text="其它如下" Visible="False" />
                                                                            </td>
                                                                        </tr>



                                                                        <tr>
                                                                            <td id="td3" align="center" bgcolor="#f7f7f7">
                                                                                <strong>他項權利類別</strong></td>
                                                                            <td id="td4" align="center" bgcolor="#f7f7f7">
                                                                                <strong>他項權利種類</strong></td>
                                                                            <td id="td5" align="center" bgcolor="#f7f7f7">
                                                                                <strong>順位</strong></td>
                                                                            <td id="td6" align="center" bgcolor="#f7f7f7">
                                                                                <strong>登記日期</strong></td>
                                                                            <td id="td7" align="center" bgcolor="#f7f7f7">
                                                                                <strong>設定性質及設定金額</strong></td>
                                                                            <td id="td8" align="center" bgcolor="#f7f7f7">
                                                                                <strong>設定權利人</strong></td>
                                                                            <%--<asp:CheckBox ID="CheckBox28" runat="server" Text="是否為公設" Visible="false"    />--%>
                                                                        </tr>
                                                                        <tr>
                                                                            <td id="td10" align="center" bgcolor="#FFFFFF">
                                                                                <asp:DropDownList ID="DropDownList37" runat="server">
                                                                                    <asp:ListItem></asp:ListItem>
                                                                                    <asp:ListItem>土地</asp:ListItem>
                                                                                    <asp:ListItem>建物</asp:ListItem>
                                                                                </asp:DropDownList>
                                                                                <asp:Label ID="Label56" runat="server" Visible="False">0</asp:Label>
                                                                            </td>
                                                                            <td id="td11" align="center" bgcolor="#FFFFFF">
                                                                                <asp:DropDownList ID="DropDownList36" runat="server">
                                                                                    <asp:ListItem></asp:ListItem>
                                                                                    <asp:ListItem>地上權</asp:ListItem>
                                                                                    <asp:ListItem>抵押權</asp:ListItem>
                                                                                    <asp:ListItem>最高限額抵押權</asp:ListItem>
                                                                                    <asp:ListItem>普通抵押權</asp:ListItem>
                                                                                    <asp:ListItem>不動產役權</asp:ListItem>
                                                                                    <asp:ListItem>農育權</asp:ListItem>
                                                                                    <asp:ListItem>典權</asp:ListItem>
                                                                                    <asp:ListItem>永佃權</asp:ListItem>
                                                                                    <asp:ListItem>耕作權</asp:ListItem>
                                                                                </asp:DropDownList>
                                                                            </td>
                                                                            <td id="td12" align="center" bgcolor="#FFFFFF">
                                                                                <input id="input44" runat="server" maxlength="2" name="address_82102"
                                                                                    style="width: 50px" /></td>
                                                                            <td id="td13" align="center" bgcolor="#FFFFFF">
                                                                                <asp:TextBox ID="Date4" runat="server" Width="75px" MaxLength="7" ToolTip="ex:1030101"></asp:TextBox>
                                                                                <input id="seeday2" runat="server" name="ch4"
                                                                                    onclick="javascript: ShowCalendar(form1.Date4); Menu_OP(Calendar)" type="button"
                                                                                    value="..." /></td>
                                                                            <td id="td14" align="center" bgcolor="#FFFFFF">新台幣 
                    <input id="input46" runat="server" maxlength="10" name="objectnumber3233472"
                        style="width: 75px" />
                                                                                萬元</td>
                                                                            <td id="td15" align="center" bgcolor="#FFFFFF">
                                                                                <input id="input47" runat="server" maxlength="20"
                                                                                    name="objectnumber3233822" size="12" />
                                                                            </td>

                                                                            <%-- <td  id="td9" align="center" bgcolor="#f7f7f7">
                    管理人</td>    --%>
                                                                        </tr>
                                                                        <tr>
                                                                            <td colspan="7">處理方式</td>
                                                                        </tr>

                                                                        <tr>
                                                                            <td colspan="7" bgcolor="#FFFFFF">
                                                                                <asp:CheckBoxList ID="CheckBoxList3" runat="server" RepeatLayout="Flow">
                                                                                    <asp:ListItem Value="6">由委託人於簽妥買賣契約後負責清償及塗銷，或由買賣雙方協議之</asp:ListItem>
                                                                                    <asp:ListItem Value="1">由買方向金融機構辦理貸款撥款清償並塗銷</asp:ListItem>
                                                                                    <asp:ListItem Value="2">由委託人於交付交屋款前清償並塗銷</asp:ListItem>
                                                                                    <asp:ListItem Value="3">由買方承受原債權及其抵押權</asp:ListItem>
                                                                                    <asp:ListItem Value="4">由買方清償並塗銷</asp:ListItem>
                                                                                    <asp:ListItem Value="5">其他</asp:ListItem>
                                                                                </asp:CheckBoxList><asp:TextBox ID="TextBox38" runat="server" placeholder="其他請填此"></asp:TextBox>

                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </table>

                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td17" colspan="6">
                                                                    <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False"
                                                                        BorderStyle="None" GridLines="none" CssClass="GridViewStyle"
                                                                        CellPadding="3" Width="100%">
                                                                        <PagerStyle Height="8px" HorizontalAlign="Center" BorderStyle="None" />
                                                                        <RowStyle CssClass="row" />
                                                                        <Columns>
                                                                            <asp:TemplateField HeaderText="刪除">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox233" runat="server"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <HeaderTemplate>
                                                                                    <input id="chkAll0" name="chkAll" onclick="Check(this, 'chkSelect2', 'GridView2')" type="checkbox" />
                                                                                    刪除
                                                                                </HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    &nbsp;<asp:CheckBox ID="chkSelect2" runat="server" />
                                                                                    <asp:Label ID="Label48" runat="server" Text='<%# Bind("物件編號") %>'
                                                                                        Visible="False"></asp:Label>
                                                                                    <asp:Label ID="Label49" runat="server" Text='<%# Bind("Num") %>'
                                                                                        Visible="False"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox234" runat="server"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Button ID="Button9" runat="server" CommandName="edits" Text="編輯"
                                                                                        CausesValidation="False" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="他項權利類別">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox235" runat="server"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label50" runat="server" Text='<%# Bind("權利類別") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="他項權利種類">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox236" runat="server" Text='<%# Bind("權利種類") %>'></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label51" runat="server" Text='<%# Bind("權利種類") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="順位">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox237" runat="server" Text='<%# Bind("順位") %>'></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label52" runat="server" Text='<%# Bind("順位") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="登記日期">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox238" runat="server" Text='<%# Bind("登記日期") %>'></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label53" runat="server" Text='<%# Bind("登記日期") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Center" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="設定性質及設定金額">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox239" runat="server" Text='<%# Bind("設定") %>'></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label54" runat="server" Text='<%# Bind("設定") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Right" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField HeaderText="設定權利人">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="TextBox240" runat="server" Text='<%# Bind("設定權利人") %>'></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="Label55" runat="server" Text='<%# Bind("設定權利人") %>'></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle CssClass="GridViewHeaderStyle" />
                                                                                <ItemStyle CssClass="GridViewItemStyle" HorizontalAlign="Left" />
                                                                            </asp:TemplateField>
                                                                        </Columns>
                                                                    </asp:GridView>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td18" colspan="6" align="right">
                                                                    <asp:ImageButton ID="ImageButton10" runat="server"
                                                                        ImageUrl="~/images/shopkeeping_del.jpg" CausesValidation="False"
                                                                        Visible="False" />
                                                                </td>
                                                            </tr>

                                                        </table>
                                                        <!-- end here -->
                                                    </div>

                                                    <%--UI介面優化--%>
                                                    <div id="View9" style="display: none">
                                                        <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位類別</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel23" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="ddl車位類別" runat="server" AutoPostBack="True">
                                                                            </asp:DropDownList>
                                                                            <asp:TextBox ID="TextBox37" runat="server" MaxLength="5"
                                                                                Width="80px"></asp:TextBox>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>進出口為</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList23" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>坡道</asp:ListItem>
                                                                        <asp:ListItem>機械升降</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位租售</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList6" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>租</asp:ListItem>
                                                                        <asp:ListItem>售</asp:ListItem>
                                                                    </asp:DropDownList>&nbsp;</td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位價格</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1"></td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位數量</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input53" runat="server"
                                                                        name="password3233" maxlength="20" onkeypress="javascript:JHshNumberText()"
                                                                        style="width: 35px" /></td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位號碼</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <input id="input42" runat="server"
                                                                        name="objectnumber223" size="12" maxlength="100" />&nbsp;號</td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位位置</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList ID="DropDownList64" runat="server">
                                                                        <asp:ListItem>地下</asp:ListItem>
                                                                        <asp:ListItem>地上</asp:ListItem>
                                                                    </asp:DropDownList>
                                                                    &nbsp;<input id="input43" runat="server"
                                                                        name="objectnumber323" size="12" maxlength="100" />&nbsp;樓</td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位說明</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox93" runat="server" MaxLength="40"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位管理費</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:UpdatePanel ID="UpdatePanel24" runat="server" UpdateMode="Conditional">
                                                                        <ContentTemplate>
                                                                            <asp:DropDownList ID="DropDownList25" runat="server">
                                                                                <asp:ListItem></asp:ListItem>
                                                                                <asp:ListItem>管理費</asp:ListItem>
                                                                                <asp:ListItem>清潔費</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:DropDownList ID="DropDownList24" runat="server" AutoPostBack="True"
                                                                                Width="50px">
                                                                                <asp:ListItem></asp:ListItem>
                                                                                <asp:ListItem>無</asp:ListItem>
                                                                                <asp:ListItem>坪</asp:ListItem>
                                                                                <asp:ListItem>月</asp:ListItem>
                                                                                <asp:ListItem>季</asp:ListItem>
                                                                                <asp:ListItem>年</asp:ListItem>
                                                                                <asp:ListItem>半年</asp:ListItem>
                                                                                <asp:ListItem>雙月</asp:ListItem>
                                                                                <asp:ListItem>未知</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                            <asp:TextBox ID="TextBox94" runat="server" MaxLength="5" Width="80px" onKeyPress="javascript:JHshNumberText()"></asp:TextBox>
                                                                            <asp:Label ID="Label461" runat="server" Text="元"></asp:Label>
                                                                        </ContentTemplate>
                                                                    </asp:UpdatePanel>
                                                                </td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>車位性質</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">

                                                                    <asp:DropDownList ID="DropDownList67" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>法定停車位</asp:ListItem>
                                                                        <asp:ListItem>自行增設停車位</asp:ListItem>
                                                                        <asp:ListItem>獎勵增設停車位</asp:ListItem>
                                                                        <asp:ListItem>無法辨識</asp:ListItem>
                                                                    </asp:DropDownList>

                                                                </td>
                                                            </tr>

                                                            <tr>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>有無辦單獨區分所有建物登記</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:DropDownList CssClass="inputbox" ID="DropDownList7" runat="server">
                                                                        <asp:ListItem></asp:ListItem>
                                                                        <asp:ListItem>無</asp:ListItem>
                                                                        <asp:ListItem>有</asp:ListItem>
                                                                    </asp:DropDownList></td>
                                                                <td bgcolor="#f7f7f7" id="td1"><strong>使用約定方式</strong></td>
                                                                <td bgcolor="#FFFFFF" id="td1">
                                                                    <asp:TextBox ID="TextBox95" runat="server"></asp:TextBox></td>
                                                            </tr>

                                                        </table>

                                                    </div>


                                                </td>
                                            </tr>
                                            <tr>
                                                <td align="center">&nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="70" align="center">
                                        <input type="hidden" id="hidtempcode" runat="server" value="" />
                                        <%--   <td  id="td16" align="center" bgcolor="#FFFFFF">
                    <input id="input122" runat="server" maxlength="20" 
                        name="objectnumber323372235" size="12" /></td>--%>

                                        <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../images/search_bt_22.gif" />

                                        <asp:ValidationSummary ID="ValidationSummary1" runat="server"
                                            EnableTheming="True" Height="1px" ShowMessageBox="True"
                                            ShowSummary="False" Width="143px" />
                                        <asp:ImageButton ID="ImageButton12" runat="server"
                                            ImageUrl="../images/tax_bt_01.gif" />
                                        <asp:ImageButton ID="ImageButton13" runat="server"
                                            ImageUrl="../images/objectmanage_bt_17.gif" />
                                        <asp:ImageButton ID="ImageButton19" runat="server"
                                            ImageUrl="../images/search_bt_07.gif" OnClientClick="if ( !confirm('是否真的要刪除該筆物件及所有相關資料?')) return false;" />
                                        <asp:CheckBox ID="CheckBox26" runat="server" Text="一併複製房地產說明書、屋主資料、照片...等資料"
                                            Checked="True" />
                                        <asp:Label ID="src" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Label22" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Label25" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="movie_h" runat="server" Visible="False"></asp:Label>
                                        <%--<asp:hiddenfield id="hidtempcode" runat="server"></asp:hiddenfield>--%>
                                        <asp:Label ID="Label23" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="ez_code" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Label62" runat="server" Visible="False"></asp:Label>

                                        <%--<input id="是否一併複制其他東西" runat="server"  type="hidden" />--%>
                                        <asp:Label ID="NotShow" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="OverDay" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="LongOrShort" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="ShowMsg" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="ShowDt" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="ShowDt_Start" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Float" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Land_FileNo" runat="server" Visible="False"></asp:Label>
                                        <asp:Label ID="Data_Source" runat="server" Visible="False"></asp:Label>
                                        <%---------------------------------參數用---------------------------------%>
                                    </td>
                                </tr>
                            </table>
                            <!--content_foot -------------------------------------------------------------------------------------------------------------->
                            <div id="content_foot"></div>


                        </div>
                        <!--inpage_main_content版尾 -->



                    </div>
                    <!--inpage_container_main版尾 -->
                </div>
                <!--inpage_container版尾 -->

                <!--footer -------------------------------------------------------------------------------------------------------------->
        <script src="../js/jquery-ui-1.8.9.custom.min.js" type="text/javascript"></script>
        <script type="text/javascript" src="js/obj_add_v4.js?_=2"></script>
        </form>

    </div>
    <!--最外層wrapper版尾 -->
    <uc2:reveserd ID="reveserd1" runat="server" />

    <%--<script type="text/javascript">
         //$("#current").css({ display: "none" });
         $("#current li").slice(3).hide();
    </script>--%>

    <!--判斷目前在哪個TABS-->
    <script type="text/javascript">
        $(document).ready(function () {
            var tabIndex = $("#<%= hidCurrentTab.ClientID %>").val();
            $("#content_main").tabs({
                select: function (event, ui) {
                    $("#<%= hidCurrentTab.ClientID %>").val(ui.index);

                }

                , selected: tabIndex

            });

            changeliback();
        });

        //判斷選取在哪個細項功能
        function changeliback() {

            $("#myul li").each(function (index, value) {

                // var selectedTab = $("#tabsI").tabs('option', 'active');

                if (index == $("#hidCurrentTab").val()) {

                    $(this).attr("id", "current");

                } else {

                    $(this).attr("id", "");

                }

            });

        }

        //<!--判斷目前在哪個TABS-->

        $(document).ready(function () {
            if ($("#Checkbox27").is(':checked')) {
                $('.tdPublicCar').find(':input').prop("disabled", true);
                $('.tdPublicCar').find(':input').val("");
                $('#txtPublicCar').css("color", "LightGray");
                $("#hid_PublicCar").val("Y");
            }

            $("#DropDownList2").change(function () {
                var str = "";
                $("#DropDownList2 option:selected").each(function () {
                    str += $(this).text() + " ";
                });

                if (str.indexOf('公設內') >= 0 && $("#Checkbox27").is(':checked')) {
                    $("#TextBox77").val('0');
                    $("#TextBox73").val('0');
                    $("#TextBox1").val('1');
                    $("#TextBox3").val('1');
                    $("#TextBox44").val('1');
                    $("#TextBox45").val('1');
                    $("#TextBox22").val('0');
                    $("#TextBox24").val('0');

                    $("#TextBox77").prop("disabled", true);
                    $("#TextBox73").prop("disabled", true);
                    $("#TextBox1").prop("disabled", true);
                    $("#TextBox3").prop("disabled", true);
                    $("#TextBox44").prop("disabled", true);
                    $("#TextBox45").prop("disabled", true);
                    $("#TextBox22").prop("disabled", true);
                    $("#TextBox24").prop("disabled", true);
                } else {
                    $("#TextBox77").prop("disabled", false);
                    $("#TextBox73").prop("disabled", false);
                    $("#TextBox1").prop("disabled", false);
                    $("#TextBox3").prop("disabled", false);
                    $("#TextBox44").prop("disabled", false);
                    $("#TextBox45").prop("disabled", false);
                    $("#TextBox22").prop("disabled", false);
                    $("#TextBox24").prop("disabled", false);
                }

            });

            $("#content_main").tabs({

                select: function (event, ui) {

                    // Do stuff here

                    $("#<%= hidCurrentTab.ClientID %>").val(ui.index);

                    changeliback();

                }

            });
        });
        function DisablePublicCar() {

            if ($("#Checkbox27").is(':checked')) {
                //tdPublicCar disable
                $('.tdPublicCar').find(':input').prop("disabled", true);
                $('.tdPublicCar').find(':input').val("");
                $('#txtPublicCar').css("color", "LightGray");

                $("#hid_PublicCar").val("Y");

            } else {
                $('.tdPublicCar').find(':input').prop("disabled", false);
                $('#txtPublicCar').css("color", "");

                $("#hid_PublicCar").val("N");

                //$("#DropDownList2").append('<option value="車位面積(公設內)" >車位面積(公設內)</option>');
            }

        }

        function GB_showCenter(v, url, height, width) {
            $.colorbox({
                'width': width,
                'height': height,
                'iframe': true,
                'href': url,
                'fixed': true
            });
        }

        function AddOneRowTb() {
            var rowCount = $('#Table_right tr').length - 1;
            //name
            var txt_name = "<input type='text' id='tb_Textbox50_" + rowCount + "' style='width:50px;' />"
            //權利範圍 textbox
            var txtbox_right = " 權利範圍.<input type='text' id='tb_Numerator_1_" + rowCount + "' title='分母' style='width:50px;' /> &nbsp;分之 &nbsp;<input type=text' id='tb_Denominator_1_" + rowCount + "' title='分子' style='width:50px;' onblur='calculatorarea(this.id);' /></br>出售權利範圍.<input type='text'  id='tb_Numerator_2_" + rowCount + "' title='分母' style='width:50px;' /> &nbsp;分之 &nbsp;<input type='text' id='tb_Denominator_2_" + rowCount + "' title='分子' style='width:50px;' onblur='calculatorarea2(this.id);' />"
            //持有面積 lb
            var lb_holder = "持有面積:<span id='tb_Label21_1_" + rowCount + "'></span><br />出售面積:<span id='tb_Label21_2_" + rowCount + "'></span>"
            //出售權利種類 select 
            var select1 = "<select id='tb_DropDownList70_" + rowCount + "'><option selected='selected' value='所有權'>所有權</option><option value='地上權'>地上權</option><option value='典權'>典權</option><option value='使用權'>使用權</option></select>"
            //add and del button
            var bons = "<td><input id='Button11" + rowCount + "' type='button' value='儲存' onclick='SaveRow(this)'   /><input id='Button12" + rowCount + "' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode" + rowCount + "'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span></td>"

            //所有權種類 select
            var select2 = " <select  id='tb_DropDownList71_" + rowCount + "'><option value='單獨所有'>單獨所有</option><option selected='selected' value='分別共有'>分別共有</option><option value='公同共有'>公同共有</option></select>"
            $('#Table_right tr:last').after('<tr align="center"><td><span id="tb_num' + rowCount + '">' + rowCount + '</span></td><td>' + txt_name + '</td><td align=left>' + txtbox_right + '</td><td>' + lb_holder + '</td><td>' + select1 + '</td><td>' + select2 + '</td>' + bons + '</tr>');
        }
        function SaveRow(n) {
            $("#loading").show();
            var sid = $("#store").val();

            var contractType = $("#ddl契約類別").val();
            var contractCode;
            if (contractType == "專任") { contractCode = '1' }
            if (contractType == "一般") { contractCode = '6' }
            if (contractType == "同意書") { contractCode = '7' }
            if (contractType == "流通") { contractCode = '5' }
            if (contractType == "庫存") { contractCode = '9' }

            //alert(sid.substring(1, 5));
            var oid = contractCode + sid.substring(1, 5) + $("#TextBox2").val();
            if (oid.length < 13 || sid.length == 0) { alert("請先輸入正確的物件編號!!"); $("#loading").hide(); return; }

            var row = n.parentNode.parentNode;
            var rowNumber = $(row).find('td:first').text();
            //所有權人姓名
            var rowName = $(row).find('td').eq(1).find("input").val();
            if (rowName == "") { $("#loading").hide(); return false }
            //權利範圍
            var right1 = $(row).find('td').eq(2).find("input[id*=tb_Numerator_1]").val();
            if (right1.length == 0) { right1 = '0' }

            var right2 = $(row).find('td').eq(2).find("input[id*=tb_Denominator_1]").val();
            if (right2.length == 0) { right2 = '0' }
            //出售權利範圍
            var right3 = $(row).find('td').eq(2).find("input[id*=tb_Numerator_2]").val();
            var right4 = $(row).find('td').eq(2).find("input[id*=tb_Denominator_2]").val();
            if (right3.length == 0) { right3 = '0' }
            if (right4.length == 0) { right4 = '0' }
            //持有面積
            var area1 = $(row).find('td').eq(3).find("span[id*=tb_Label21_1_]").text();
            var ar_ar1 = area1.split(",");
            if (ar_ar1.length > 1) {
                area1 = ar_ar1[0].replace("㎡", "");
            } else {
                area1 = '0';
            }

            //出售面積
            var area2 = $(row).find('td').eq(3).find("span[id*=tb_Label21_2_]").text();
            var ar_ar2 = area2.split(",");

            if (ar_ar2.length > 1) {
                area2 = ar_ar2[0].replace("㎡", "");
            } else {
                area2 = '0';
            }


            //權利種類
            var righttype = $(row).find('td').eq(4).find("select option:selected").text();
            //所有權種類
            var ownerrighttype = $(row).find('td').eq(5).find("select option:selected").text();

            var st = 'rowNum:' + rowNumber + ',rowName:' + rowName + ', rg1:' + right1 + ', rg2:' + right2 + ', rg3:' + right3 + ', rg4:' + right4 + ', ar1:' + area1 + ', ar2:' + area2 + ', righttype1:' + righttype + ', righttype2:' + ownerrighttype + ', oid:' + oid + ', sid:' + sid;

            $.ajax({
                url: "Obj_add_ownerlist_add.ashx",
                cache: false,
                data: { rowNum: rowNumber, rowName: rowName, rg1: right1, rg2: right2, rg3: right3, rg4: right4, ar1: area1, ar2: area2, righttype1: righttype, righttype2: ownerrighttype, oid: oid, sid: sid },
                dataType: "json",
                success: function (data) {
                    if (data.AddRusult == false) {
                        $(row).find('td').eq(6).find("span[id*=tempcode]").text(data.Tempcode);
                        $("#loading").hide();
                        return false;
                    } else {
                        $("#loading").hide();
                        //alert(rowName + " 新增成功!(" + data.ResuMsg + ")");
                        alert(rowName + " 新增成功!");
                        $(row).find('td').eq(6).find("input[id*=Button11]").replaceWith("<img src='image/tick_64.png' style='width:30px;height:30px' />");
                        $(row).find('td').eq(6).find("span[id*=tempcode]").text(data.Tempcode);
                        var tmpstr = $("#hidtempcode").val() + ',' + data.Tempcode;
                        $("#hidtempcode").val(tmpstr);

                    }
                },
                error: function (request, status, error) { $("#loading").hide(); alert("失敗，請洽管理員"); return false; }
            });

            //$.getJSON('Obj_add_ownerlist_add.ashx', { rowNum: rowNumber, rowName: rowName, rg1: right1, rg2: right2, rg3: right3, rg4: right4, ar1: area1, ar2: area2, righttype1: righttype, righttype2: ownerrighttype, oid: oid, sid: sid }, function(data) {
            //    if (data.AddRusult == false) {
            //        $(row).find('td').eq(6).find("span[id*=tempcode]").text(data.Tempcode);
            //        $("#loading").hide();
            //        return false;
            //    } else {
            //        $("#loading").hide();
            //        alert(rowName + " 新增成功!");
            //        $(row).find('td').eq(6).find("input[id*=Button11]").replaceWith("<img src='image/tick_64.png' style='width:30px;height:30px' />");
            //        $(row).find('td').eq(6).find("span[id*=tempcode]").text(data.Tempcode);
            //        var tmpstr = $("#hidtempcode").val() + ',' + data.Tempcode;
            //        $("#hidtempcode").val(tmpstr);

            //    }

            //}).fail(function() { $("#loading").hide(); alert("失敗，請洽管理員"); return false; });

        }

        function DelRow(n) {
            $("#loading").show();
            var sid = getUrlParameter('sid');
            var oid = getUrlParameter('oid');
            var row = n.parentNode.parentNode;
            var rowNumber = $(row).find('td:first').text();
            var tmpcode = $(row).find('td').eq(6).find("span[id*=tempcode]").text();

            $.getJSON('Obj_add_ownerlist_del.ashx', { rowNum: rowNumber, oid: oid, sid: sid, tmpcode: tmpcode }, function (data) {
                if (data.AddRusult == false) {
                    $("#loading").hide();
                    return false;
                } else {
                    //remove tmpcode in  hidtempcode
                    var temstr = $("#hidtempcode").val();
                    temstr = temstr.replace(',,', '');
                    temstr = temstr.replace(tmpcode, '');

                    document.getElementById('hidtempcode').value = temstr;

                    $("#loading").hide();
                    //if (rowNumber == '1') { return false; }
                    $(row).remove();
                }

            }).fail(function () { $("#loading").hide(); alert("失敗，請洽管理員"); return false; });
        }
        function calculatorarea(idname) {
            var idrow = idname.split("_");
            idrow = idrow[idrow.length - 1];
            //計算標示部權利範圍
            if ($('#DropDownList2 option:selected').val() == "主建物") {
                var TX1 = $("#TextBox1").val();
                var TX3 = $("#TextBox3").val();
            }

            //取得最後一碼判斷row
            var rowLabel = "#tb_Label21_1_" + idrow;

            //分子 idname 分母
            var tb_Denominator = $("#" + idname).val();


            var tb_Numerator = "#tb_Numerator_1_" + idrow;
            var tb_Numerator_value = $(tb_Numerator).val();

            //總面積(坪)
            var TotalPin = $("#TextBox77").val();
            //$(rowLabel).text(tb_Denominator + ',' + tb_Numerator_value + ',' + TotalPin);
            if (tb_Denominator.length > 0 && tb_Numerator_value.length > 0 && TotalPin.length > 0) {


                var ReturnVal = (parseFloat(TotalPin) * parseFloat(tb_Denominator)) / parseFloat(tb_Numerator_value);
                if (TX1) {
                    if (TX1.length > 0 && TX3.length > 0) {
                        ReturnVal = (ReturnVal * parseFloat(TX3)) / parseFloat(TX1);
                    }
                }
                $(rowLabel).text(ReturnVal + '㎡, ' + ReturnVal * 0.3025 + '坪');
            }
        }

        function calculatorarea2(idname) {
            var idrow = idname.split("_");
            idrow = idrow[idrow.length - 1];



            var TX1, TX3, TX44, TX45;
            //計算標示部權利範圍
            if ($('#DropDownList2 option:selected').val() == "主建物") {
                TX1 = $("#TextBox1").val();
                TX3 = $("#TextBox3").val();

                TX44 = $("#TextBox44").val();
                TX45 = $("#TextBox45").val();

            }
            //取得最後一碼判斷row
            var rowLabel = "#tb_Label21_2_" + idrow;

            //分子 idname 分母
            var tb_Denominator = $("#" + idname).val();


            var tb_Numerator = "#tb_Numerator_2_" + idrow;
            var tb_Numerator_value = $(tb_Numerator).val();

            //總面積(坪)
            var TotalPin = $("#TextBox77").val();
            //$(rowLabel).text(tb_Denominator + ',' + tb_Numerator_value + ',' + TotalPin);
            if (tb_Denominator.length > 0 && tb_Numerator_value.length > 0 && TotalPin.length > 0) {
                var ReturnVal = (parseFloat(TotalPin) * parseFloat(tb_Denominator)) / parseFloat(tb_Numerator_value);

                if (TX1) {
                    if (TX1.length > 0 && TX3.length > 0) {
                        ReturnVal = (ReturnVal * parseFloat(TX3)) / parseFloat(TX1);

                    }
                }
                if (TX44) {
                    if (TX44.length > 0 && TX45.length > 0) {
                        ReturnVal = (ReturnVal * parseFloat(TX45)) / parseFloat(TX44);
                    }
                }

                ////回傳值應為持有面積乘上 出售面積
                ////持有面積
                //var rowLabelholder = "#tb_Label21_1_" + idrow;
                //var area1 = $(rowLabelholder).text();
                //var ar_ar1 = area1.split(",");
                //if (ar_ar1.length > 1) {
                //    area1 = ar_ar1[0].replace("㎡", "");
                //} else {
                //    area1 = '0';
                //}
                //ReturnVal = (parseFloat(area1) * parseFloat(tb_Denominator)) / parseFloat(tb_Numerator_value)

                $(rowLabel).text(ReturnVal + '㎡, ' + ReturnVal * 0.3025 + '坪');
            }
        }
        //計算四期款
        function checkPeriodMoney() {
            var TextBox12 = document.getElementById('<%=TextBox12.ClientID  %>');

            if (TextBox12.value == '') {
                alert("請先輸入售價！")
                return;
            }
            var TotalPrice = parseFloat(TextBox12.value); //總價

            var TextBox258 = document.getElementById('<%=TextBox258.ClientID  %>'); //第一期百分比
            var period1 = parseInt(TextBox258.value);
            var TextBox259 = document.getElementById('<%=TextBox259.ClientID  %>'); //第二期百分比
            var period2 = parseInt(TextBox259.value);
            var TextBox260 = document.getElementById('<%=TextBox260.ClientID  %>'); //第三期百分比
            var period3 = parseInt(TextBox260.value);
            var TextBox261 = document.getElementById('<%=TextBox261.ClientID  %>'); //第四期百分比
            var period4 = parseInt(TextBox261.value);

            if (!isNaN(period1)) {
                $("#TextBox262").val(TotalPrice * (parseInt(period1) / 100));
            }
            if (!isNaN(period2)) {
                $("#TextBox263").val(TotalPrice * (parseInt(period2) / 100));
            }
            if (!isNaN(period3)) {
                $("#TextBox264").val(TotalPrice * (parseInt(period3) / 100));
            }
            if (!isNaN(period4)) {
                $("#TextBox265").val(TotalPrice * (parseInt(period4) / 100));
            }
        }

        function checkchar(ar) {
            var ty = ar.value.match(/[\ufee0-\uffdf]/g);
            if (ty) {
                alert("訴求重點不允許全形字");
                return false;
            }
        }

        // 本建號為公設、本建號為車位兩個 CheckBox 只能選一個(因為 RadioButton 不能取消選取所以用 CheckBox )
        function CheckCheckboxes(n) {

            if (n == 98) {
                if (document.getElementById("CheckBox98").checked == true) {
                    document.getElementById("CheckBox99").checked = false
                }
            } else if (n == 99) {
                if (document.getElementById("CheckBox99").checked == true) {
                    document.getElementById("CheckBox98").checked = false
                }
            }

        }

        document.addEventListener('DOMContentLoaded', function () {
            const sel = document.getElementById('build_structure_temp');
            const input97 = document.getElementById('input97');

            sel.addEventListener('change', function () {
                const v = sel.value; // 例如 "SRC(鋼骨鋼筋混凝土)"

                if (v !== "0") {
                    input97.value = v;
                } else {
                    // 選回「請選擇」時要怎麼處理：
                    // 1) 清空
                    input97.value = "";
                    // 2) 或者不動（把上面這行註解掉即可）
                }
            });
        });

    </script>


</body>
</html>
