<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Build_DataV2.aspx.vb" Inherits="A_ObjectManage_Build_Data" MaintainScrollPositionOnPostback="True" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>超級房仲家管理系統</title>
    <script type="text/javascript" src="../js/png.js"></script>
    <script type="text/javascript" src="../js/pngfix.js"></script>
    <script type="text/javascript" src="../js/calendar.js"></script>
    <link href="../css/index.css" rel="stylesheet" type="text/css" />
    <link href="../css/inpage.css" rel="stylesheet" type="text/css" />
    <link href="../css/popup.css" rel="stylesheet" type="text/css">
    <link href="../css/tab1.css" rel="stylesheet" type="text/css" />
    <link href="../css/tab2.css" rel="stylesheet" type="text/css" />
    <link href="../css/tab_02.css" rel="stylesheet" type="text/css">

    <style type="text/css">
        .auto-style1 {
            width: 21px;
        }

        .auto-style2 {
            color: #FF0000;
            font-size: large;
        }

        .auto-style3 {
            width: 21px;
            height: 75px;
        }

        .auto-style4 {
            height: 75px;
        }
    </style>
    <style type="text/css">
        div#tabs-1 > table:nth-child(1) td:nth-child(1) {
            width: 211px !important;
        }

        input, select {
            margin-left: 6px !important;
        }

            input[type=number], input[type=text] {
                width: 80px;
            }

        .search td {
            padding: 6px 4px 6px 6px;
            text-align: left;
        }

        textarea {
            width: 300px !important;
            margin-bottom: -10px;
        }

        td#td59 input {
            width: 88px;
        }
    </style>
</head>
<body>
    <div id="wrapper">
        <!--最外層wrapper版頭 -->
        <form id="form2" runat="server">
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <div id="popup">
                <!--popup版頭 -->
                <div id="popup_main">
                    <!--popup_main版頭 -->
                    <div id="popup_main_bar">
                        <!--popup_main_bar版頭 -->
                        <div id="popup_left_bar">
                            <img src="../images/place_data_bar.jpg" height="59" />
                        </div>
                        <div id="popup_route">
                            <img src="../images/home_icon.jpg" width="14" height="11" /><a href="#">Home</a> > <a href="#">委賣物件新增</a> > <a href="#">產調_建物</a>
                        </div>
                    </div>
                    <!--popup_main_bar版尾 -->
                    <div id="popup_main_content">
                        <!--popup_main_content版頭 -->
                        <div id="popup_content">
                            <!--popup_content版頭 -->

                            <div id="tabs">
                                <ul>
                                    <li><a href="#tabs-1">經紀人員調查細項</a></li>
                                    <li><a href="#tabs-2">屋主現況說明</a></li>
                                </ul>
                                <div id="tabs-1">
                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">

                                        <tr>
                                            <td id="td1" width="175" bgcolor="#f7f7f7">&nbsp;</td>
                                            <td id="td1" bgcolor="#FFFFFF">
                                                <asp:HyperLink ID="HyperLink1" runat="server" ForeColor="Red"
                                                    NavigateUrl="~/A_ObjectManage/應記載事項-使用管制函詢單位.xls" Target="_blank">各項應記載事項-使用管制函詢單位查詢</asp:HyperLink>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td id="td1" width="175" bgcolor="#f7f7f7"><strong>建物建號</strong></td>
                                            <td id="td1" bgcolor="#FFFFFF">
                                                <asp:Label ID="Label5" runat="server"></asp:Label>
                                                <asp:Label ID="sid" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="oid" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="NUM" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="usid" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="uoid" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="Label12" runat="server" Visible="False"></asp:Label>
                                                <asp:Label ID="Label11" runat="server" Visible="False"></asp:Label>
                                            </td>
                                        </tr>

                                    </table>
                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">1</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>出售之建物是否為共有</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel11" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList58" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder19" runat="server" Visible="false">
                                                            <asp:RadioButtonList ID="RadioButtonList4" runat="server">
                                                                <asp:ListItem Value="無">無分管協議</asp:ListItem>
                                                                <asp:ListItem Value="有">有分管協議</asp:ListItem>
                                                            </asp:RadioButtonList>
                                                            &nbsp;<asp:Label ID="Label8" runat="server" Font-Bold="True" Text="說明"></asp:Label>
                                                            &nbsp;<asp:TextBox ID="TextBox68" runat="server" Width="500px"
                                                                Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>

                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">2</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>獎勵容積開放空間供公共使用情形</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">

                                                <asp:UpdatePanel ID="UpdatePanel6" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>

                                                        <asp:DropDownList ID="DropDownList36" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label17" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox236" runat="server" Width="500px" Visible="False"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">3</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>目前作住宅使用之建物是否位屬工業區或不得作住宅使用之商業區或其他分區(若有，其合法性)</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel27" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList10" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label37" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox245" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">4</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>使用執照有無備註之注意事項</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">

                                                <asp:UpdatePanel ID="UpdatePanel17" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>

                                                        <asp:DropDownList ID="DropDownList59" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label36" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox241" runat="server" Width="500px" Visible="False" Height="50px"
                                                            TextMode="MultiLine" MaxLength="500"></asp:TextBox>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style3">5</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7" class="auto-style4"><strong>禁建情事(應敘明位置、約略面積、列管情形)</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF" class="auto-style4">

                                                <asp:UpdatePanel ID="UpdatePanel35" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList57" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>

                                                        &nbsp;<asp:Label ID="Label44" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox21" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td76" bgcolor="#f7f7f7" class="auto-style1">6</td>
                                            <td id="td77" width="175" bgcolor="#f7f7f7" class="auto-style4"><strong>本棟建物有無依法設置之中繼幫浦機械室或水箱</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">

                                                <asp:UpdatePanel ID="UpdatePanel44" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList61" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>

                                                        &nbsp;<asp:Label ID="Label45" runat="server" Font-Bold="True" Text="說明(應敘明樓層)"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox249" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td bgcolor="#f7f7f7" class="auto-style1">7</td>
                                            <td width="175" bgcolor="#f7f7f7" class="auto-style4"><strong>有無設置太陽光電發電設備</strong></td>
                                            <td bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel45" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList62" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="LabelSolarPos" runat="server" Font-Bold="True" Text="設置位置" Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox250" runat="server" Visible="False" Width="500px" MaxLength="200"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td bgcolor="#f7f7f7" class="auto-style1">8</td>
                                            <td width="175" bgcolor="#f7f7f7" class="auto-style4"><strong>有無取得建築能效標示</strong></td>
                                            <td bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel46" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList63" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder86" runat="server" Visible="False">能效等級：<asp:DropDownList ID="DropDownList64" runat="server" Width="180px">
                                                            <asp:ListItem Text="-請選擇能效等級-" Value="" Selected="True"></asp:ListItem>
                                                            <asp:ListItem Text="1+" Value="1+"></asp:ListItem>
                                                            <asp:ListItem Text="1" Value="1"></asp:ListItem>
                                                            <asp:ListItem Text="2" Value="2"></asp:ListItem>
                                                            <asp:ListItem Text="3" Value="3"></asp:ListItem>
                                                            <asp:ListItem Text="4" Value="4"></asp:ListItem>
                                                            <asp:ListItem Text="5" Value="5"></asp:ListItem>
                                                            <asp:ListItem Text="6" Value="6"></asp:ListItem>
                                                            <asp:ListItem Text="7" Value="7"></asp:ListItem>
                                                        </asp:DropDownList>&nbsp;&nbsp;1(最高) ~ 7(最低) 1+代表「近零碳建築」<br />
                                                            有效期間：<asp:TextBox ID="TextBox253" runat="server" Width="220px" MaxLength="50"></asp:TextBox>
                                                        </asp:PlaceHolder>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>

                                    </table>
                                    <div class="text-center">
                                        <span class="auto-style2"><strong>個案權利調查</strong></span>
                                    </div>
                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">1</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>他項權利</strong></td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:Label ID="Label46" runat="server" Font-Bold="True" Text="同謄本的他項權利部"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">2</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>限制登記</strong><asp:Image
                                                ID="Image2" runat="server" ImageUrl="../images/s.png"
                                                ToolTip="包括：預告登記、查封、假扣押、假處分及其他禁止處分之登記，詳如附登記謄本，若有，應敘明。"
                                                ImageAlign="Middle" />
                                            </td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel28" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList41" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>查封</asp:ListItem>
                                                            <asp:ListItem>假扣押</asp:ListItem>
                                                            <asp:ListItem>假處分</asp:ListItem>
                                                            <asp:ListItem>預告登記</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label38" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox246" runat="server" Height="50px" TextMode="MultiLine"
                                                            Visible="False" Width="500px" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">3</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>信託登記</strong><asp:Image
                                                ID="Image3" runat="server" ImageUrl="../images/s.png"
                                                ToolTip="若有，應敘明信託契約之主要條款內容。（依登記謄本及信託專簿記載)"
                                                ImageAlign="Middle" />
                                            </td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel29" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList29" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label39" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox247" runat="server" Height="50px" TextMode="MultiLine"
                                                            Visible="False" Width="500px" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">4</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7"><strong>其他事項</strong><asp:Image
                                                ID="Image4" runat="server" ImageUrl="../images/s.png"
                                                ToolTip="如：依民事訴訟法第二百五十四條規定及其他相關之註記等。"
                                                ImageAlign="Middle" />
                                            </td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel30" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label40" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox248" runat="server" Height="50px" TextMode="MultiLine"
                                                            Visible="False" Width="500px" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <div id="tabs-2">
                                    <table bgcolor="#e7e7e7" width="100%" border="0" cellpadding="1" cellspacing="1" class="search">
                                        <tr>

                                            <td id="td2" colspan="2" width="175" bgcolor="#f7f7f7">&nbsp;<asp:CheckBox ID="house" Text="全棟" runat="server"></asp:CheckBox></td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:Label runat="server" Text="&amp;nbsp;※勾選全棟時請將其它樓層情況於此頁面說明欄位敘明。" ForeColor="#CC0000"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">&nbsp;</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7">1-7項</td>
                                            <td id="td3" bgcolor="#FFFFFF">本欄位調查已於土地調查輸入</td>
                                        </tr>
                                        <tr>
                                            <td id="td2" bgcolor="#f7f7f7" class="auto-style1">8</td>
                                            <td id="td2" width="175" bgcolor="#f7f7f7">本棟建物頂樓平台有無依法設置之行動電話基地台設施</td>
                                            <td id="td3" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox1" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td4" bgcolor="#f7f7f7" class="auto-style1">9</td>
                                            <td id="td4" width="175" bgcolor="#f7f7f7">出售之建物衛生下水道工程有無完成</td>
                                            <td id="td5" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList3" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:DropDownList ID="DropDownList4" runat="server">
                                                            <asp:ListItem>主管機關尚未通知施工日期</asp:ListItem>
                                                            <asp:ListItem>其他</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label2" runat="server" Font-Bold="True" Text="說明"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox2" runat="server" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td6" bgcolor="#f7f7f7" class="auto-style1">10</td>
                                            <td id="td6" width="175" bgcolor="#f7f7f7">有無規約以外特殊使用及其限制</td>
                                            <td id="td7" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel3" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;共用部分有分管協議
                        <asp:PlaceHolder ID="PlaceHolder7" runat="server" Visible="false">,敘明:
                           
                            &nbsp;<asp:TextBox ID="TextBox3" runat="server" Width="500px"
                                Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>



                        </asp:PlaceHolder>
                                                        <br />

                                                        <asp:DropDownList ID="DropDownList60" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;使用專有部分有限制
                        <asp:PlaceHolder ID="PlaceHolder77" runat="server" Visible="false">,敘明:

                            &nbsp;<asp:TextBox ID="TextBox4" runat="server" Width="500px"
                                Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td8" bgcolor="#f7f7f7" class="auto-style1">11</td>
                                            <td id="td8" width="175" bgcolor="#f7f7f7">社區有無管理維護公司</td>
                                            <td id="td9" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel4" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList6" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label6" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox5" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td10" bgcolor="#f7f7f7" class="auto-style1">12</td>
                                            <td id="td10" width="175" bgcolor="#f7f7f7">住戶規約內容是否有約定共用</td>
                                            <td id="td11" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel5" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList8" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder8" runat="server" Visible="false">&nbsp;範圍:<asp:TextBox ID="TextBox6" runat="server" MaxLength="100"></asp:TextBox>
                                                            &nbsp;使用方式及文件:
                            <asp:TextBox ID="TextBox7" runat="server" MaxLength="100"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td12" bgcolor="#f7f7f7" class="auto-style1">13</td>
                                            <td id="td12" width="175" bgcolor="#f7f7f7">住戶規約內容是否有約定專用</td>
                                            <td id="td13" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel7" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList9" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder9" runat="server" Visible="false">範圍:
                            &nbsp;<asp:TextBox ID="TextBox8" runat="server" MaxLength="100"></asp:TextBox>
                                                            使用方式及文件:
                            &nbsp;
                            <asp:TextBox ID="TextBox9" runat="server" MaxLength="100"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td14" bgcolor="#f7f7f7" class="auto-style1">14</td>
                                            <td id="td14" width="175" bgcolor="#f7f7f7">社區之公共基金之數額、提撥及運用方式</td>
                                            <td id="td15" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel8" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList11" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder5" runat="server" Visible="false">數額:
                          
                            &nbsp;
                            <asp:TextBox ID="TextBox10" runat="server" MaxLength="100"></asp:TextBox>
                                                            萬元&nbsp;提撥方式:
                            <asp:DropDownList ID="DropDownList12" runat="server">
                                <asp:ListItem>管理費</asp:ListItem>
                                <asp:ListItem>其他</asp:ListItem>
                            </asp:DropDownList>
                                                            &nbsp;<asp:TextBox ID="TextBox11" runat="server" MaxLength="100" placeholder="若為其他請填寫"></asp:TextBox>
                                                            運用方式:
                           
                            <asp:DropDownList ID="DropDownList13" runat="server">
                                <asp:ListItem>管委會決議</asp:ListItem>
                                <asp:ListItem>其他</asp:ListItem>
                            </asp:DropDownList>
                                                            &nbsp;<asp:TextBox ID="TextBox12" runat="server" MaxLength="100" placeholder="若為其他請填寫"></asp:TextBox>

                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td16" bgcolor="#f7f7f7" class="auto-style1">15</td>
                                            <td id="td16" width="175" bgcolor="#f7f7f7">社區之管理費或使用費數額及繳交方式</td>
                                            <td id="td17" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel9" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList14" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder6" runat="server" Visible="false">
                                                            <asp:DropDownList ID="DropDownList15" runat="server">
                                                                <asp:ListItem>管理費</asp:ListItem>
                                                                <asp:ListItem>使用費</asp:ListItem>
                                                            </asp:DropDownList>
                                                            &nbsp; &nbsp;數額:
                            <asp:TextBox ID="TextBox14" runat="server" MaxLength="100" onkeydown="onlyNum(event)"></asp:TextBox>元
                            &nbsp; &nbsp;繳交方式:
                           
                            &nbsp;<asp:TextBox ID="TextBox15" runat="server" MaxLength="100"></asp:TextBox>
                                                        </asp:PlaceHolder>



                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td18" bgcolor="#f7f7f7" class="auto-style1">16</td>
                                            <td id="td18" width="175" bgcolor="#f7f7f7">社區有無有管理組織及管理方式</td>
                                            <td id="td19" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel10" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList16" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>



                                                        <asp:DropDownList ID="DropDownList17" runat="server" Visible="false">
                                                            <asp:ListItem>管委會</asp:ListItem>
                                                            <asp:ListItem>管理負責人</asp:ListItem>
                                                            <asp:ListItem>管理服務人</asp:ListItem>
                                                            <asp:ListItem>其他</asp:ListItem>
                                                        </asp:DropDownList>

                                                        <asp:TextBox ID="TextBox13" runat="server" MaxLength="100" placeholder="若為其他請填寫" Visible="false"></asp:TextBox>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td20" bgcolor="#f7f7f7" class="auto-style1">17</td>
                                            <td id="td20" width="175" bgcolor="#f7f7f7">住戶規約內容有無使用手冊</td>
                                            <td id="td21" bgcolor="#FFFFFF">
                                                <asp:DropDownList ID="DropDownList18" runat="server">
                                                    <asp:ListItem>無</asp:ListItem>
                                                    <asp:ListItem>有</asp:ListItem>
                                                </asp:DropDownList>
                                                (若有請檢附手冊)
                         
                   
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td22" bgcolor="#f7f7f7" class="auto-style1">18 </td>
                                            <td id="td22" width="175" bgcolor="#f7f7f7">有無電梯設備</td>
                                            <td id="td23" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel12" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList7" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:Label ID="Label7" runat="server" Font-Bold="True" Text="電梯設備有無張貼有效合格認證標章" Visible="False"></asp:Label>

                                                        <asp:DropDownList ID="DropDownList19" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label16" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox16" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td24" bgcolor="#f7f7f7" class="auto-style1">19</td>
                                            <td id="td24" width="175" bgcolor="#f7f7f7">建物有無出租情形</td>
                                            <td id="td25" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel42" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList53" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder1" runat="server" Visible="false">出租範圍: 
                            <asp:DropDownList ID="DropDownList55" runat="server">
                                <asp:ListItem>建物</asp:ListItem>
                                <asp:ListItem>車位</asp:ListItem>
                                <asp:ListItem>其他</asp:ListItem>
                            </asp:DropDownList>
                                                            &nbsp;
                            <asp:TextBox ID="TextBox57" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                            <br />
                                                            <asp:RadioButton ID="RadioButton6" runat="server" Text="不定期租約，租金" GroupName="rent_gp" />
                                                            <asp:TextBox ID="TextBox55" runat="server" Type="number" step="any">0</asp:TextBox>
                                                            萬/月<br />
                                                            <asp:RadioButton ID="RadioButton7" runat="server" Text="定期租約(須檢附租約)，租賃契約有無公證：" GroupName="rent_gp" />
                                                            <asp:RadioButtonList ID="RadioButtonList3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                                <asp:ListItem>有</asp:ListItem>
                                                                <asp:ListItem>無</asp:ListItem>
                                                            </asp:RadioButtonList><br />
                                                            &nbsp;&nbsp; 租金:<asp:TextBox ID="TextBox59" runat="server" Type="number" step="any">0</asp:TextBox>萬/月，押租保證金<asp:TextBox ID="TextBox60" runat="server" Type="number" step="any">0</asp:TextBox>萬
                                租期:<asp:TextBox ID="TextBox61" runat="server" type="date"></asp:TextBox>~<asp:TextBox ID="TextBox62" runat="server" type="date"></asp:TextBox><br />
                                                            <asp:RadioButton ID="RadioButton8" runat="server" Text="租賃之權利義務隨同移轉" GroupName="rent_gp2" /><br />
                                                            <asp:RadioButton ID="RadioButton9" runat="server" Text="屋主終止租約騰空交屋" GroupName="rent_gp2" /><br />
                                                            <asp:RadioButton ID="RadioButton10" runat="server" Text="其他" GroupName="rent_gp2" /><asp:TextBox ID="TextBox58" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td26" bgcolor="#f7f7f7" class="auto-style1">20</td>
                                            <td id="td26" width="175" bgcolor="#f7f7f7">建物有無有出借情形</td>
                                            <td id="td27" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel43" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList54" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder2" runat="server" Visible="false">出借範圍: 
                            <asp:DropDownList ID="DropDownList56" runat="server">
                                <asp:ListItem>建物</asp:ListItem>
                                <asp:ListItem>車位</asp:ListItem>
                                <asp:ListItem>其他</asp:ListItem>
                            </asp:DropDownList>
                                                            &nbsp;
                            <asp:TextBox ID="TextBox56" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                            <br />
                                                            <asp:RadioButton ID="RadioButton11" runat="server" Text="無書面約定" GroupName="barrow_gp" /><br />
                                                            <asp:RadioButton ID="RadioButton12" runat="server" Text="有書面約定" GroupName="barrow_gp" /><br />
                                                            需檢附借用人姓名:
                                                            <asp:TextBox ID="TextBox63" runat="server"></asp:TextBox><br />
                                                            借用期間:
                                                            <asp:TextBox ID="TextBox64" runat="server" type="date"></asp:TextBox>~<asp:TextBox ID="TextBox65" runat="server" type="date"></asp:TextBox><br />
                                                            返還條件:
                                                            <asp:TextBox ID="TextBox66" runat="server"></asp:TextBox>
                                                        </asp:PlaceHolder>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td28" bgcolor="#f7f7f7" class="auto-style1">21</td>
                                            <td id="td28" width="175" bgcolor="#f7f7f7">建物有無占用情形</td>
                                            <td id="td29" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel13" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList20" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder13" runat="server" Visible="false">
                                                            <asp:DropDownList ID="DropDownList21" runat="server">
                                                                <asp:ListItem>建物占用他人土地</asp:ListItem>
                                                                <asp:ListItem>建物被他人占用</asp:ListItem>
                                                                <asp:ListItem>其他</asp:ListItem>
                                                            </asp:DropDownList>
                                                            <asp:TextBox ID="TextBox67" runat="server" placeholder="若為其他請填寫"></asp:TextBox><br />
                                                            &nbsp;<asp:Label ID="Label20" runat="server" Font-Bold="True" Text="說明"></asp:Label>
                                                            &nbsp;<asp:TextBox ID="TextBox17" runat="server" Width="500px"
                                                                Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td30" bgcolor="#f7f7f7" class="auto-style3">22</td>
                                            <td id="td30" width="175" bgcolor="#f7f7f7" class="auto-style4">本標的物及公設內有無消防設施</td>
                                            <td id="td31" bgcolor="#FFFFFF" class="auto-style4">
                                                <asp:UpdatePanel ID="UpdatePanel14" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList22" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>

                                                        <asp:CheckBoxList ID="CheckBoxList1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" Visible="false">
                                                            <asp:ListItem Value="緊急照明燈">緊急照明燈 </asp:ListItem>
                                                            <asp:ListItem Value="滅火器">滅火器 </asp:ListItem>
                                                            <asp:ListItem Value="灑水設施">灑水設施 </asp:ListItem>
                                                            <asp:ListItem Value="緩降機">緩降機 </asp:ListItem>
                                                            <asp:ListItem Value="其他">其他 </asp:ListItem>
                                                        </asp:CheckBoxList>
                                                        <asp:TextBox ID="TextBox18" runat="server" placeholder="若為其他請填寫" Visible="false"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td32" bgcolor="#f7f7f7" class="auto-style1">23</td>
                                            <td id="td32" width="175" bgcolor="#f7f7f7">本標的物及公設內有無無障礙設施</td>
                                            <td id="td33" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel15" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList23" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder10" runat="server" Visible="false">&nbsp;設施項目:
                                                            <asp:TextBox ID="TextBox19" runat="server"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td34" bgcolor="#f7f7f7" class="auto-style1">24</td>
                                            <td id="td34" width="175" bgcolor="#f7f7f7">房屋有無施作夾層</td>
                                            <td id="td35" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel16" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList24" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder11" runat="server" Visible="false">&nbsp;權狀面積:
                            <asp:TextBox ID="TextBox20" runat="server" Type="number" step="any">0</asp:TextBox>
                                                            m², 無建物所有權狀面積約
                                                            <asp:TextBox ID="TextBox22" runat="server" Type="number" step="any">0</asp:TextBox>m², 其他:
                                                            <asp:TextBox ID="TextBox23" runat="server" MaxLength="15"></asp:TextBox>

                                                        </asp:PlaceHolder>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td36" bgcolor="#f7f7f7" class="auto-style1">25</td>
                                            <td id="td36" width="175" bgcolor="#f7f7f7">本戶有無供水使用</td>
                                            <td id="td37" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel18" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList25" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder12" runat="server" Visible="false">
                                                            <asp:DropDownList ID="DropDownList26" runat="server">
                                                                <asp:ListItem>自來水</asp:ListItem>
                                                                <asp:ListItem>地下水</asp:ListItem>
                                                            </asp:DropDownList>
                                                            &nbsp;供水是否正常:
                           <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                               <asp:ListItem Selected="True">是</asp:ListItem>
                               <asp:ListItem>否</asp:ListItem>
                           </asp:RadioButtonList>
                                                            &nbsp;說明<asp:TextBox ID="TextBox24" runat="server" Width="500px"
                                                                Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td38" bgcolor="#f7f7f7" class="auto-style1">26</td>
                                            <td id="td38" width="175" bgcolor="#f7f7f7">本戶有無獨立之電表供電</td>
                                            <td id="td39" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel19" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList27" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label21" runat="server" Font-Bold="True" Text="若無請敘明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox25" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td40" bgcolor="#f7f7f7" class="auto-style1">27</td>
                                            <td id="td40" width="175" bgcolor="#f7f7f7">本戶瓦斯供應情形</td>
                                            <td id="td41" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel20" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList28" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:RadioButtonList ID="RadioButtonList2" runat="server" RepeatDirection="Vertical" RepeatLayout="Flow" Visible="false">
                                                            <asp:ListItem Selected="True">正常使用天然瓦斯</asp:ListItem>
                                                            <asp:ListItem>屋內有天然瓦斯管線，尚未裝表</asp:ListItem>
                                                            <asp:ListItem>僅大樓有天然瓦斯管線</asp:ListItem>
                                                            <asp:ListItem>屋內有天然瓦斯管線，已拆表</asp:ListItem>
                                                            <asp:ListItem>使用桶裝瓦斯</asp:ListItem>
                                                            <asp:ListItem>其他</asp:ListItem>
                                                        </asp:RadioButtonList>
                                                        <asp:TextBox ID="TextBox26" runat="server" placeholder="若為其他請填寫" Visible="false"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td42" bgcolor="#f7f7f7" class="auto-style1">28</td>
                                            <td id="td42" width="175" bgcolor="#f7f7f7">水、電管線於產權持有期間有無更新</td>
                                            <td id="td43" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel21" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList30" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder14" runat="server" Visible="false">&nbsp;水管更新日期:
                            <asp:TextBox ID="TextBox27" runat="server"></asp:TextBox>

                                                            &nbsp;電線更新日期: 
                            <asp:TextBox ID="TextBox28" runat="server"></asp:TextBox>
                                                        </asp:PlaceHolder>


                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td44" bgcolor="#f7f7f7" class="auto-style1">29</td>
                                            <td id="td44" width="175" bgcolor="#f7f7f7">有無積欠應繳費用情形？ (水費、電費、瓦斯費、管理費或其他費用) </td>
                                            <td id="td45" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel22" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList31" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>

                                                        &nbsp;<asp:Label ID="Label3" runat="server" Text="項目與金額:" Visible="true"></asp:Label>
                                                        <asp:TextBox ID="TextBox29" runat="server" Visible="true"></asp:TextBox>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td46" bgcolor="#f7f7f7" class="auto-style1">30</td>
                                            <td id="td46" width="175" bgcolor="#f7f7f7">所有權持有期間有無居住</td>
                                            <td id="td47" bgcolor="#FFFFFF">
                                                <asp:DropDownList ID="DropDownList32" runat="server">
                                                    <asp:ListItem>無</asp:ListItem>
                                                    <asp:ListItem>有</asp:ListItem>
                                                </asp:DropDownList>

                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td48" bgcolor="#f7f7f7" class="auto-style1">31</td>
                                            <td id="td48" width="175" bgcolor="#f7f7f7">社區有無公共設施重大修繕決議？(所有權人另須付費)</td>
                                            <td id="td49" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel24" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList33" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder15" runat="server" Visible="false">&nbsp;決議內容:
                            <asp:TextBox ID="TextBox30" runat="server"></asp:TextBox>
                                                            &nbsp;須負擔金額: 
                            <asp:TextBox ID="TextBox31" runat="server"></asp:TextBox>
                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td50" bgcolor="#f7f7f7" class="auto-style1">32</td>
                                            <td id="td50" width="175" bgcolor="#f7f7f7">混凝土中水溶性氯離子含量檢測情事</td>
                                            <td id="td51" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel25" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList34" runat="server">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        &nbsp;若無,未檢測原因:
                             <asp:DropDownList ID="DropDownList35" runat="server">
                                 <asp:ListItem>屋主未檢測</asp:ListItem>
                                 <asp:ListItem>其他</asp:ListItem>
                             </asp:DropDownList>
                                                        <asp:TextBox ID="TextBox32" runat="server" placeholder="若為其他請填寫"></asp:TextBox>

                                                        <br />
                                                        &nbsp;若有,附檢測結果（依建築日期區分標準，民國87年6月25日以前為0.6kg/m3以下；民國87年6月25日以後修正為0.3kg/m3以下。）  
                            
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td52" bgcolor="#f7f7f7" class="auto-style1">33</td>
                                            <td id="td52" width="175" bgcolor="#f7f7f7">輻射檢測情事</td>
                                            <td id="td53" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel26" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList37" runat="server" AutoPostBack="True">
                                                            <asp:ListItem Selected="True">無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>

                                                        &nbsp;<asp:Label ID="Label9" runat="server" Text="若無,未檢測原因:"></asp:Label>
                                                        <asp:DropDownList ID="DropDownList38" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>屋主未檢測</asp:ListItem>
                                                            <asp:ListItem>其他</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:TextBox ID="TextBox33" runat="server" placeholder="若為其他請填寫"></asp:TextBox>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td54" bgcolor="#f7f7f7" class="auto-style1">34</td>
                                            <td id="td54" width="175" bgcolor="#f7f7f7">建物有無曾經發生火災或其他天然災害或人為破壞造成建築物損害及修繕情形</td>
                                            <td id="td55" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel31" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList39" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label22" runat="server" Font-Bold="True" Text="損害及修繕情形:"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox34" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td56" bgcolor="#f7f7f7" class="auto-style1">35</td>
                                            <td id="td56" width="175" bgcolor="#f7f7f7">建物目前有無因地震被建管單位公告列為危險建築情形</td>
                                            <td id="td57" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel32" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList40" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label23" runat="server" Font-Bold="True" Text="列管時間及危險等級:"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox35" runat="server" Visible="False" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td58" bgcolor="#f7f7f7" class="auto-style1">36</td>
                                            <td id="td58" width="175" bgcolor="#f7f7f7">樑、柱部分有無顯見間隙裂痕</td>
                                            <td id="td59" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel33" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList42" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <asp:PlaceHolder ID="PlaceHolder16" runat="server" Visible="false">其位置:
                            &nbsp;<asp:TextBox ID="TextBox36" runat="server"></asp:TextBox>
                                                            裂痕長度:
                            &nbsp;
                            &nbsp;<asp:TextBox ID="TextBox37" runat="server"></asp:TextBox>
                                                            間隙寬度:
                            &nbsp;
                            &nbsp;<asp:TextBox ID="TextBox38" runat="server"></asp:TextBox>

                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td60" bgcolor="#f7f7f7" class="auto-style1">37</td>
                                            <td id="td60" width="175" bgcolor="#f7f7f7">房屋鋼筋有無裸露</td>
                                            <td id="td61" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel34" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList43" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label27" runat="server" Font-Bold="True" Text="若有,其位置:"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox39" runat="server" Visible="False"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td62" bgcolor="#f7f7f7" class="auto-style1">38</td>
                                            <td id="td62" width="175" bgcolor="#f7f7f7">本建物(專有部分)有無發生兇殺、自殺、一氧化碳中毒或其他非自然死亡之情形</td>
                                            <td id="td63" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel36" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList44" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder17" runat="server" Visible="false">
                                                            <asp:RadioButton ID="RadioButton1" runat="server" Text="產權持有期間:" GroupName="suicide_gp" /><asp:TextBox ID="TextBox40" runat="server"></asp:TextBox><br />
                                                            <asp:RadioButton ID="RadioButton2" runat="server" Text="產權持有期間前,賣方:" GroupName="suicide_gp" /><asp:TextBox ID="TextBox41" runat="server"></asp:TextBox><br />
                                                            <asp:RadioButton ID="RadioButton3" runat="server" Text="確認無發生過" GroupName="suicide_gp" /><br />
                                                            <asp:RadioButton ID="RadioButton4" runat="server" Text="知道曾發生過" GroupName="suicide_gp" /><br />
                                                            <asp:RadioButton ID="RadioButton5" runat="server" Text="其他:" GroupName="suicide_gp" /><asp:TextBox ID="TextBox42" runat="server" placeholder="若為其他請填寫"></asp:TextBox><br />

                                                        </asp:PlaceHolder>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td64" bgcolor="#f7f7f7" class="auto-style1">39</td>
                                            <td id="td64" width="175" bgcolor="#f7f7f7">有無滲漏水狀況</td>
                                            <td id="td65" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel37" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList45" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder4" runat="server" Visible="false">滲漏水位置: 
                            <asp:CheckBoxList ID="CheckBoxList2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem Value="屋頂">屋頂 </asp:ListItem>
                                <asp:ListItem Value="外牆">外牆 </asp:ListItem>
                                <asp:ListItem Value="窗框">窗框 </asp:ListItem>
                                <asp:ListItem Value="客廳">客廳 </asp:ListItem>
                                <asp:ListItem Value="臥室">臥室 </asp:ListItem>
                                <asp:ListItem Value="廚房">廚房 </asp:ListItem>
                                <asp:ListItem Value="浴室">浴室 </asp:ListItem>
                                <asp:ListItem Value="前陽台">前陽台 </asp:ListItem>
                                <asp:ListItem Value="後陽台">後陽台 </asp:ListItem>
                                <asp:ListItem Value="冷熱水管">冷熱水管 </asp:ListItem>
                                <asp:ListItem Value="其他">其他 </asp:ListItem>
                            </asp:CheckBoxList><asp:TextBox ID="TextBox43" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                            <br />
                                                            處理方式:
                             <asp:DropDownList ID="DropDownList46" runat="server" AutoPostBack="True">
                                 <asp:ListItem>賣方修繕後交屋</asp:ListItem>
                                 <asp:ListItem>買方自行修繕</asp:ListItem>
                                 <asp:ListItem>減價</asp:ListItem>
                                 <asp:ListItem>其他</asp:ListItem>
                             </asp:DropDownList>
                                                            &nbsp;
                            <asp:TextBox ID="TextBox44" runat="server" placeholder="若為其他請填寫"></asp:TextBox>

                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td66" bgcolor="#f7f7f7" class="auto-style1">40</td>
                                            <td id="td66" width="175" bgcolor="#f7f7f7">建物有無違建、增建情事？ 其位置、約略面積及建管機關列管情形？ （包括未登記之改建、增建、違建及建商二次施工部分） </td>
                                            <td id="td67" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel38" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList47" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder3" runat="server" Visible="false">違建、增建位置: 
                            <br />
                                                            <asp:CheckBoxList ID="CheckBoxList3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="5">
                                                                <asp:ListItem Value="地下室">地下室 </asp:ListItem>
                                                                <asp:ListItem Value="一樓空地">一樓空地 </asp:ListItem>
                                                                <asp:ListItem Value="夾層">夾層 </asp:ListItem>
                                                                <asp:ListItem Value="頂樓">頂樓 </asp:ListItem>
                                                                <asp:ListItem Value="防火巷">防火巷 </asp:ListItem>
                                                                <asp:ListItem Value="格局變更">格局變更 </asp:ListItem>
                                                                <asp:ListItem Value="陽台外推">陽台外推 </asp:ListItem>
                                                                <asp:ListItem Value="陽台加建">陽台加建 </asp:ListItem>
                                                                <asp:ListItem Value="露台">露台 </asp:ListItem>
                                                                <asp:ListItem Value="平台外推">平台外推 </asp:ListItem>
                                                                <asp:ListItem Value="其他">其他 </asp:ListItem>
                                                            </asp:CheckBoxList><asp:TextBox ID="TextBox45" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                            <br />
                                                            約略面積:
                                                            <asp:TextBox ID="TextBox47" runat="server" onKeyPress="javascript:JHshNumberText()" type="number" step="any"></asp:TextBox>㎡ 
                            <br />
                                                            列管情形:
                             <asp:DropDownList ID="DropDownList48" runat="server">
                                 <asp:ListItem>無</asp:ListItem>
                                 <asp:ListItem>不知</asp:ListItem>
                                 <asp:ListItem>有</asp:ListItem>
                                 <asp:ListItem>其他</asp:ListItem>
                             </asp:DropDownList>&nbsp;
                            <asp:TextBox ID="TextBox46" runat="server" placeholder="若為其他請填寫"></asp:TextBox>
                                                            <br />
                                                            賣方保證有權處分違建、增建物及保證將違建、增建物隨同主建物一併移轉買方，絕無異議。

                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td68" bgcolor="#f7f7f7" class="auto-style1">41</td>
                                            <td id="td68" width="175" bgcolor="#f7f7f7">建物排水系統有無正常</td>
                                            <td id="td69" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel39" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList49" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem Selected="True">有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:Label ID="Label4" runat="server" Text="若不正常:" Visible="false"></asp:Label>
                                                        <asp:DropDownList ID="DropDownList50" runat="server" Visible="false">
                                                            <asp:ListItem>交屋前修復</asp:ListItem>
                                                            <asp:ListItem>現況交屋賣方不負修繕責任</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="td70" bgcolor="#f7f7f7" class="auto-style1">42</td>
                                            <td id="td70" width="175" bgcolor="#f7f7f7">是否有附屬設備？ (未記載者，以不動產委託銷售契約簽定時之現況為準)</td>
                                            <td id="td71" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel40" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList51" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <br />
                                                        <asp:PlaceHolder ID="PlaceHolder18" runat="server" Visible="false">
                                                            <asp:CheckBox ID="CheckBox1" runat="server" Text="固定物現況" />
                                                            <asp:CheckBox ID="CheckBox2" runat="server" Text="燈飾" />
                                                            <asp:CheckBox ID="CheckBox3" runat="server" Text="梳妝台" />
                                                            <asp:CheckBox ID="CheckBox4" runat="server" Text="沙發" /><asp:TextBox ID="TextBox48" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>組
                                 <asp:CheckBox ID="CheckBox5" runat="server" Text="電視機" /><asp:TextBox ID="TextBox49" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>台
                                 <asp:CheckBox ID="CheckBox6" runat="server" Text="窗簾" />
                                                            <asp:CheckBox ID="CheckBox7" runat="server" Text="衣櫃" />
                                                            <asp:CheckBox ID="CheckBox8" runat="server" Text="壁櫥" />
                                                            <asp:CheckBox ID="CheckBox9" runat="server" Text="冰箱" /><asp:TextBox ID="TextBox50" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>台
                                 <asp:CheckBox ID="CheckBox10" runat="server" Text="冷氣機" /><asp:TextBox ID="TextBox51" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>台
                                 <asp:CheckBox ID="CheckBox11" runat="server" Text="流理台" />
                                                            <asp:CheckBox ID="CheckBox12" runat="server" Text="熱水器" />
                                                            <asp:CheckBox ID="CheckBox13" runat="server" Text="抽油煙機" />
                                                            <asp:CheckBox ID="CheckBox14" runat="server" Text="洗衣機" /><asp:TextBox ID="TextBox52" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>台
                                <asp:CheckBox ID="CheckBox15" runat="server" Text="乾衣機" /><asp:TextBox ID="TextBox53" runat="server" Width="30" MaxLength="2" Type="number">0</asp:TextBox>台
                                <asp:CheckBox ID="CheckBox16" runat="server" Text="瓦斯爐" />
                                                            <asp:CheckBox ID="CheckBox17" runat="server" Text="天然瓦斯" />
                                                            <asp:CheckBox ID="CheckBox18" runat="server" Text="瓦斯度數表" />
                                                            <asp:CheckBox ID="CheckBox19" runat="server" Text="其他" />

                                                        </asp:PlaceHolder>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td id="td74" bgcolor="#f7f7f7" class="auto-style1">43</td>
                                            <td id="td74" width="175" bgcolor="#f7f7f7">有無併同出售之車位</td>
                                            <td id="td75" bgcolor="#FFFFFF">本欄位調查已於 產調_車位 輸入</td>
                                        </tr>

                                        <tr>
                                            <td id="td74" bgcolor="#f7f7f7" class="auto-style1">44</td>
                                            <td id="td74" width="175" bgcolor="#f7f7f7">本棟建物有無依法設置之中繼幫浦機械室或水箱</td>
                                            <td id="td75" bgcolor="#FFFFFF">本欄位調查已於 產調_建物-經紀人員調查細項-第六項 輸入</td>
                                        </tr>

                                        <tr>
                                            <td id="td74" bgcolor="#f7f7f7" class="auto-style1">45</td>
                                            <td id="td74" width="175" bgcolor="#f7f7f7">其他重要事項</td>
                                            <td id="td75" bgcolor="#FFFFFF">
                                                <asp:UpdatePanel ID="UpdatePanel41" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:DropDownList ID="DropDownList52" runat="server" AutoPostBack="True">
                                                            <asp:ListItem>無</asp:ListItem>
                                                            <asp:ListItem>有</asp:ListItem>
                                                        </asp:DropDownList>
                                                        &nbsp;<asp:Label ID="Label28" runat="server" Font-Bold="True" Text="說明"
                                                            Visible="False"></asp:Label>
                                                        &nbsp;<asp:TextBox ID="TextBox54" runat="server" Visible="False" Width="500px"
                                                            Height="50px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </div>

                            <br />
                            <div style="text-align: center">
                                <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../images/search_bt_22.gif" OnClientClick="return validateEnergyFields();"/>

                                <asp:ImageButton ID="ImageButton12" runat="server" OnClientClick="return validateEnergyFields();"
                                    ImageUrl="../images/tax_bt_01.gif" />
                                <asp:ImageButton ID="ImageButton19" runat="server"
                                    ImageUrl="../images/search_bt_07.gif" OnClientClick="if ( !confirm('是否真的要刪除該筆物件及所有相關資料?')) return false;" />

                            </div>


                        </div>
                        <!--popup_content版尾 -->
                    </div>
                    <!--popup_main_content版尾 -->
                </div>
                <!--popup_main版尾 -->
            </div>
            <!--popup版尾 -->
        </form>
    </div>
    <!--最外層wrapper版尾 -->
    <script src="../js/jquery-1.8.3.min.js"></script>
    <script src="../js/jquery-ui-1.10.4.custom.js" type="text/javascript"></script>
    <%--<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.4/themes/smoothness/jquery-ui.css" />--%>
    <link rel="stylesheet" href="css/jquery-ui-1.10.4.css" />
    <script type="text/javascript">
        function validateEnergyFields() {
            var solar = $("#<%= DropDownList62.ClientID %>").val();
            var solarPos = $("#<%= TextBox250.ClientID %>").val().trim();
            var eff = $("#<%= DropDownList63.ClientID %>").val();
            var grade = $("#<%= DropDownList64.ClientID %>").val();
            var period = $("#<%= TextBox253.ClientID %>").val().trim();

            var msgs = [];
            if (solar === "有" && solarPos === "") msgs.push("「設置位置」為必填");
            if (eff === "有") {
                if (grade === "") msgs.push("「能效等級」為必填");
                if (period === "") msgs.push("「有效期間」為必填");
            }

            if (msgs.length > 0) {
                alert(msgs.join("\n"));
                return false;
            }
            return true;
        }


        function bindDatePickers() {
            $("#TextBox253").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy-mm-dd'
            });
        }

        $(function () {
            bindDatePickers();
            if (typeof (Sys) !== "undefined" && Sys.WebForms) {
                Sys.WebForms.PageRequestManager.getInstance().add_endRequest(function () {
                    bindDatePickers();
                });
            }
        });

        $(function () {
            $("#tabs").tabs();

            $("#TextBox27").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd'
            });
            $("#TextBox28").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd'
            });

        });

        //限填數字
        function JHshNumberText() {
            if (!(((window.event.keyCode >= 48) && (window.event.keyCode <= 57)) || (window.event.keyCode == 46))) {
                window.event.keyCode = 0; alert("僅限填數字");
            }
        }

        function onlyNum(event) {

            // 允許的鍵：數字鍵 (0-9)、退格 (Backspace)、刪除 (Delete)、方向鍵 (←↑↓→)、Tab、Enter
            if (!/[\d]/.test(event.key) && !["Backspace", "Tab", "ArrowLeft", "ArrowRight", "Delete", "Enter"].includes(event.key)) {
                event.preventDefault(); // 阻止輸入
            }

            //// 允許的鍵：數字鍵 (0-9)、退格 (Backspace)、刪除 (Delete)、方向鍵 (←↑↓→)、Tab、Enter、小數點 (.)、負號 (-)
            //if (!/[\d.\-]/.test(event.key) && !["Backspace", "Tab", "ArrowLeft", "ArrowRight", "Delete", "Enter"].includes(event.key)) {
            //    event.preventDefault(); // 阻止輸入
            //}

            //// 限制小數點只能輸入一次
            //if (event.key === "." && event.target.value.includes(".")) {
            //    event.preventDefault();
            //}

            //// 限制負號 (-) 只能輸入在最前面
            //if (event.key === "-" && event.target.value.length > 0) {
            //    event.preventDefault();
            //}
        }
    </script>
</body>
</html>
