Imports XmlToDoc.XmlToDoc
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Web.HttpPostedFile
Imports System.Diagnostics
Imports clspowerset
Imports System.Drawing
Imports System.Security.Cryptography
Imports System.Text
Imports ICSharpCode.SharpZipLib
Imports ICSharpCode.SharpZipLib.Zip

Partial Class description_print3_V4
    Inherits System.Web.UI.Page

    Public conn As SqlConnection
    Public cmd As SqlCommand
    Public ds As DataSet
    Public adpt As SqlDataAdapter
    Dim i As Integer
    Dim sql As String
    Dim table1, table2 As DataTable
    Dim temp As String = "" '連結地址
    Dim sqlstr As String = ""
    '連線字串
    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    Dim myobj As New clspowerset
    Dim checkvalue As New checkvpn

    Public show As Integer = 1
    Const EraEnvCount As Integer = 15 '周邊設施每頁數量

    Dim TempDictforZip As String = Server.MapPath("~/Filesystem/2015/") '多檔案時, 產生資料夾給壓縮用
    Dim RandomDictName As String '隨機亂數

    Dim sysdate As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.Cookies("webfly_empno") Is Nothing Then
            Response.Redirect("../indexnew/login3.aspx")
        End If

        If Not Page.IsCrossPagePostBack Then
            If Trim(Request("src")) = "OLD" Then
                src.Text = "委賣物件過期資料表"
            ElseIf Trim(Request("src")) = "NOW" Then
                src.Text = "委賣物件資料表"
            End If
        End If

        If Not Page.IsCrossPagePostBack Then
            If checkvalue.checkvpn(Request.Cookies("store_id").Value) = "Y" Or checkvalue.checkaround(Request.Cookies("store_id").Value) = "Y" Then
                CheckBoxList2.Visible = True
                Label3.Text = ""
            Else
                CheckBoxList2.Visible = False
                Label3.Text = "因貴店未申請安裝加密盒，故無法在不安全的網路環境下傳輸相關的鄰避設施資訊!!"
            End If

            conn = New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            Dim storeid As String = Request("sid")
            Dim oid As String = Request("oid")

            myobj.power_object(Request.Cookies("webfly_empno").Value)

            If Not Page.IsPostBack Then
                '物件型態(類別)
                sql = "select * from 資料_物件類別 With(NoLock) where 店代號 = 'A0001'"
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                DropDownList3.Items.Clear()
                DropDownList3.Items.Add("請選擇")
                i = 0
                For i = 0 To table1.Rows.Count - 1
                    DropDownList3.Items.Add(table1.Rows(i)("實價全名").ToString.Trim)
                    DropDownList3.Items(i + 1).Value = Trim(table1.Rows(i)("名稱").ToString.Trim)
                Next

                '帶入物件型態(類別)
                sql = "select 物件類別 from " & src.Text & " With(NoLock) where 店代號 = '" & Request("sid") & "' and 物件編號='" & Request("oid") & "'"
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                If table1.Rows.Count > 0 Then
                    If Not IsDBNull(table1.Rows(0)("物件類別").ToString.Trim) Then
                        DropDownList3.SelectedValue = table1.Rows(0)("物件類別").ToString.Trim
                    End If
                End If

                '東森
                Dim endday_東 As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)
                Dim startday_東 = (DateAdd("m", -6, (Left(endday_東, 3) + 1911) & "/" & (Mid(endday_東, 4, 2)) & "/" + Right(endday_東, 2)))

                Dim startday_東STR As String = Right("000" & (Year(startday_東) - 1911), 3) & Right("00" & Month(startday_東), 2) & Right("00" & Day(startday_東), 2)

                TextBox2.Text = startday_東STR
                TextBox6.Text = endday_東

                '內政部
                Dim endday As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)
                Dim startday = (DateAdd("m", -3, (Left(endday, 3) + 1911) & "/" & (Mid(endday, 4, 2)) & "/" + Right(endday, 2)))
                Dim startday_STR As String = Right("000" & (Year(startday) - 1911), 3) & Right("00" & Month(startday), 2) & Right("00" & Day(startday), 2)

                TextBox22.Text = startday_STR
                TextBox25.Text = endday
            End If

            conn.Close()
            conn.Dispose()
        End If

        Call power()

        If Not IsPostBack Then
            If Trim(DropDownList3.SelectedItem.Text) = "土地" Then
                show = 1
                DropDownList1.SelectedValue = "依路名"
                DropDownList1.Enabled = False
            Else
                show = 0
                DropDownList1.Enabled = True
            End If
            Call ddl產生路名()
        End If
    End Sub

    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        If CheckBox5.Checked = True And IsPostBack Then
            '產生亂數資料夾
            RandomDictName = CreateRand(8)
            StartGenPostCard6()
        End If
        Call open_crt()
    End Sub

    '物調表4
    Protected Sub StartGenPostCard6()
        Dim Contract_No As String = Request.QueryString("oid")
        Dim sid As String = Request.QueryString("sid")

        Dim addrField As String = "完整地址"
        Dim ds1 As DataSet
        Dim adpt1 As SqlDataAdapter
        Dim checkvalue As New checkleave

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            Dim sql As String = ""
            Dim sql2 As String = ""
            sql2 = "SELECT 張貼卡案名 "
            sql2 = sql2 & " From " & src.Text & " Where 物件編號 = '" & Contract_No & "' and 店代號 = '" & sid & "'"  '20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
            Dim Cmd2 As New SqlCommand(sql2, conn)
            Dim myDRr2 As SqlDataReader = Cmd2.ExecuteReader

            'If Request("oid") = "61319AAE32716" Then
            '    Response.Write(sql2)
            '    Response.End()
            'End If

            myDRr2.Read()
            If IsDBNull(myDRr2("張貼卡案名")) Then
                sql = "Select  A.建築名稱,A.臨路寬,A.面寬,A.縱深 ,B.bs_officer As 店長,'(' + B.bs_oarea + ')' + B.bs_otel as 連絡電話1,C.man_name as 員工姓名, D.man_name AS 行政, H.現況, A.帶看方式,isnull(P.man_name,'') AS 營業員代號1,A.委託截止日,H.建築用途 AS 產用途 "
            Else
                If Trim(myDRr2("張貼卡案名")) <> "" Then
                    sql = "Select  A.張貼卡案名,A.建築名稱,A.臨路寬,A.面寬,A.縱深 ,B.bs_officer As 店長,'(' + B.bs_oarea + ')' + B.bs_otel as 連絡電話1,C.man_name as 員工姓名, D.man_name AS 行政, H.現況, A.帶看方式, isnull(P.man_name,'') AS 營業員代號1, A.營業員代號2,A.委託截止日,H.建築用途 AS 產用途 "
                Else
                    sql = "Select  A.建築名稱,A.臨路寬,A.面寬,A.縱深 ,B.bs_officer As 店長,'(' + B.bs_oarea + ')' + B.bs_otel as 連絡電話1,C.man_name as 員工姓名, D.man_name AS 行政, H.現況, A.帶看方式, isnull(P.man_name,'') AS 營業員代號1, A.營業員代號2,A.委託截止日,H.建築用途 AS 產用途 "
                End If
            End If
            myDRr2.Close()

            sql &= ",* "
            sql = sql & " FROM   ((((((((" & src.Text & " A LEFT JOIN hsbsmg B ON A.店代號=B.bs_dept) " '20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
            sql = sql & " LEFT JOIN psman C ON A.經紀人代號=C.man_emp_no) "
            sql = sql & " LEFT JOIN psman P ON A.營業員代號1=P.man_emp_no) "
            sql = sql & " LEFT JOIN psman D ON A.輸入者=D.man_emp_no) "
            sql = sql & " LEFT JOIN 委賣物件相片 E ON A.物件編號=E.物件編號 and a.店代號 = e.店代號) "
            sql = sql & " LEFT JOIN 委賣地圖 F ON A.物件編號=F.物件編號 and a.店代號 = f.店代號) "
            sql = sql & " LEFT JOIN 委賣隔局圖 G ON A.物件編號=G.物件編號 and a.店代號 = g.店代號 ) "
            sql = sql & " LEFT JOIN 委賣_房地產說明書 H ON A.物件編號=H.物件編號 and a.店代號 = h.店代號) "
            sql = sql & " Where A.物件編號 = '" & Contract_No & "' and A.店代號 = '" & sid & "'"

            'If Request("oid") = "61319AAE32716" Then
            '    Response.Write(sql)
            '    Response.End()
            'End If

            Dim Cmd As New SqlCommand(sql, conn)
            Dim myDRr As SqlDataReader = Cmd.ExecuteReader
            Dim RptDT As String
            Dim postcardName As String
            Dim objID As String
            Dim buildName As String
            Dim SalePrice As String
            Dim CityArea As String
            Dim objAddr As String
            Dim H720 As String
            Dim bulidPin As String
            Dim AttachPin As String
            Dim HousePin As String
            Dim PublicPin As String
            Dim TotPin As String
            Dim HouseStyle As String
            Dim ManaPrice As String
            Dim seatDirect As String
            Dim NearWidth As String
            Dim NearAWidth As String
            Dim NearBWidth As String
            Dim GroundPin As String
            Dim PubStallPin As String
            Dim AddBuildPin As String
            Dim HomeLiftNum As String
            Dim UpFloorNum As String
            Dim DownFloorNum As String
            Dim loanBank As String
            Dim setMoney As String
            Dim loan2Bank As String
            Dim set2Money As String
            Dim CompleteDT As String
            Dim HouseOld As String
            Dim SiteFloor As String
            Dim bulidType As String
            Dim NowState As String
            Dim cartype As String = ""
            Dim carex As String
            Dim bringLookType As String
            Dim PriSchool As String
            Dim SecSchool As String
            Dim parkname As String
            Dim markname As String
            Dim Environment As String
            Dim KeyNo As String
            Dim objDescrption As String
            Dim AllPin As String
            Dim StoreBoss As String
            Dim AdminiName As String
            Dim Promoter As String
            Dim usetype As String
            Dim Huse As String

            While myDRr.Read
                '1000325小豪新增-相片資料夾區分為過期(expired).有效(available)
                'If myDRr("委託截止日") >= (Format(Year(Today.AddDays(-7)) - 1911, "00#") & Format(Month(Today.AddDays(-7)), "0#") & Format(Day(Today.AddDays(-7)), "0#")) Then
                Me.Label3.Text = "available"
                'Else
                '    Me.Label3.Text = "expired"
                'End If

                RptDT = Right("0" & (Now.Year - 1911), 3) & "年" & Now.ToString("MM月dd日")
                postcardName = Replace(Server.HtmlEncode(myDRr("張貼卡案名").ToString.Trim), "&", "＆")
                objID = Mid(myDRr("物件編號").ToString.Trim, 6)
                'Promoter = Server.HtmlEncode(myDRr("員工姓名").ToString.Trim) & "　" & Server.HtmlEncode(myDRr("營業員代號1").ToString.Trim)

                buildName = Replace(Server.HtmlEncode(myDRr("建築名稱").ToString.Trim), "&", "＆")

                'AllPin = Convert.ToDouble("0" & myDRr("總坪數").ToString.Trim)
                SalePrice = Convert.ToDouble("0" & myDRr("刊登售價").ToString.Trim)
                CityArea = Server.HtmlEncode(myDRr("縣市").ToString.Trim) & Server.HtmlEncode(myDRr("鄉鎮市區").ToString.Trim)
                objAddr = Server.HtmlEncode(myDRr(addrField).ToString.Trim) '地址
                If myDRr("movie_720").ToString.Trim = "Y" Then
                    H720 = "■有拍攝720度"
                Else
                    H720 = "□有拍攝720度"
                End If
                bulidPin = Convert.ToDouble("0" & myDRr("主建物").ToString.Trim) & "坪"
                AttachPin = Convert.ToDouble("0" & myDRr("附屬建物").ToString.Trim) & "坪"
                PublicPin = Convert.ToDouble("0" & myDRr("公共設施").ToString.Trim) & "坪"
                'TotPin = Convert.ToDouble(Val(myDRr("主建物").ToString.Trim) + Val(myDRr("附屬建物").ToString.Trim) + Val(myDRr("公共設施").ToString.Trim)) & "坪"
                TotPin = Convert.ToDouble(Val(myDRr("總坪數").ToString.Trim)) & "坪"

                ' -------- 格局 ----------------------
                If myDRr.Item("房").ToString.Trim = -1 Then
                    HouseStyle = "開放格局"
                Else
                    If Convert.ToDouble("0" & myDRr.Item("房").ToString.Trim) <> 0 Or Convert.ToDouble("0" & myDRr.Item("廳").ToString.Trim) <> 0 Or Convert.ToDouble("0" & myDRr.Item("衛").ToString.Trim) <> 0 Or Convert.ToDouble("0" & myDRr.Item("室").ToString.Trim) <> 0 Then
                        HouseStyle = Convert.ToDouble("0" & myDRr.Item("房").ToString.Trim) & "房" & Convert.ToDouble("0" & myDRr.Item("廳").ToString.Trim) & "廳" & Convert.ToDouble("0" & myDRr.Item("衛").ToString.Trim) & "衛" & Convert.ToDouble("0" & myDRr.Item("室").ToString.Trim) & "室"
                    End If
                End If

                ' --------- 管理費 ----------------------
                If Convert.ToDouble("0" & myDRr("管理費").ToString.Trim) <= 0 Then
                    ManaPrice = myDRr("管理費1").ToString.Trim
                Else
                    ManaPrice = Convert.ToDouble("0" & myDRr("管理費").ToString.Trim) & "元"

                    If Server.HtmlEncode(myDRr("管理費單位").ToString.Trim) <> "" Then
                        ManaPrice &= " / " & Server.HtmlEncode(myDRr("管理費單位").ToString.Trim)
                    End If
                End If

                seatDirect = Server.HtmlEncode(myDRr("座向").ToString.Trim)

                If myDRr("臨路寬").ToString.Trim <> "" Then
                    NearWidth = Server.HtmlEncode(myDRr("臨路寬").ToString.Trim) & "米"
                End If
                If myDRr("面寬").ToString.Trim <> "" Then
                    NearAWidth = Server.HtmlEncode(myDRr("面寬").ToString.Trim) & "米"
                End If
                If myDRr("縱深").ToString.Trim <> "" Then
                    NearBWidth = Server.HtmlEncode(myDRr("縱深").ToString.Trim) & "米"
                End If

                GroundPin = Convert.ToDouble("0" & myDRr("土地坪數").ToString.Trim) & "坪"
                '使用分區

                HousePin = Convert.ToDouble(Val(myDRr("主建物").ToString.Trim) + Val(myDRr("附屬建物").ToString.Trim)) & "坪" '室內坪數

                'AddBuildPin = Convert.ToDouble("0" & myDRr("加蓋坪數").ToString.Trim) & "坪"

                Dim 增建 As String = ""

                '1040415新增有無增建
                If IsDBNull(myDRr("增建")) = False Then
                    If myDRr("增建") = "Y" Then
                        增建 = "有"
                    ElseIf myDRr("增建") = "N" Then
                        增建 = "無"
                    End If
                Else '若為NULL值，改判段加蓋平方公尺是否有值
                    If IsDBNull(myDRr("加蓋平方公尺")) = False Then
                        If myDRr("加蓋平方公尺") <> 0 Then
                            增建 = "有"
                        Else
                            增建 = "無"
                        End If
                    Else
                        增建 = "無"
                    End If
                End If
                AddBuildPin = 增建

                ' --------- 每層戶數+每層電梯數 ----------------------
                Dim tmp每層戶數 As String = ""
                Dim tmp電梯數 As String = ""

                If Server.HtmlEncode(myDRr("每層戶數").ToString.Trim) <> "" Then
                    tmp每層戶數 = Server.HtmlEncode(myDRr("每層戶數").ToString.Trim) & "戶"
                End If

                If Server.HtmlEncode(myDRr("每層電梯數").ToString.Trim) <> "" Then
                    tmp電梯數 = "共用" & Server.HtmlEncode(myDRr("每層電梯數").ToString.Trim) & "部電梯"
                End If

                HomeLiftNum = tmp每層戶數 & tmp電梯數

                ' ------ 地上樓層 ------------------------------
                If myDRr.Item("地上層數").ToString.Trim <> "" Then
                    UpFloorNum = myDRr.Item("地上層數").ToString.Trim & "樓"
                Else
                    UpFloorNum = ""
                End If

                ' ------ 地下樓層 ------------------------------
                If myDRr.Item("地下層數").ToString.Trim <> "" Then
                    DownFloorNum = myDRr.Item("地下層數").ToString.Trim & "樓"
                Else
                    DownFloorNum = ""
                End If

                loanBank = Server.HtmlEncode(myDRr("貸款銀行").ToString.Trim)
                setMoney = Convert.ToDouble("0" & myDRr("貸款金額").ToString.Trim) & "萬"
                loan2Bank = Server.HtmlEncode(myDRr("貸款銀行2").ToString.Trim)
                set2Money = Convert.ToDouble("0" & myDRr("貸款金額2").ToString.Trim) & "萬"

                Dim adpt_他項權利 As SqlDataAdapter
                Dim ds_他項權利 As DataSet
                Dim ii As Integer = 0
                If Server.HtmlEncode(myDRr("貸款銀行").ToString.Trim) = "" Then
                    '讀入新版他項權利
                    Dim conn_他項權利 = New SqlConnection(EGOUPLOADSqlConnStr)

                    conn_他項權利.Open()

                    sql = "Select * From 物件他項權利細項 Where 物件編號 = '" & Contract_No & "' and 店代號='" & sid & "'  order by 權利類別 desc ,順位 asc "

                    adpt_他項權利 = New SqlDataAdapter(sql, conn_他項權利)
                    ds_他項權利 = New DataSet()

                    adpt_他項權利.Fill(ds_他項權利, "細項內容")
                    Dim tb_細項內容 As DataTable = ds_他項權利.Tables("細項內容")

                    If tb_細項內容.Rows.Count > 0 Then
                        '判斷有無與土地他項權利部相同
                        If Not IsDBNull(myDRr("與土地他項權利部相同")) Then
                            If myDRr("與土地他項權利部相同").ToString = 1 Then '相同
                                For ii = 0 To tb_細項內容.Rows.Count - 1
                                    If tb_細項內容.Rows(ii)("權利類別") = "土地" Then
                                        If Not IsDBNull(tb_細項內容.Rows(ii)("設定")) Then
                                            If tb_細項內容.Rows(ii)("設定") <> "" Then
                                                Select Case ii
                                                    Case 0
                                                        loanBank = tb_細項內容.Rows(ii)("設定權利人")
                                                        setMoney = tb_細項內容.Rows(ii)("設定") & "萬"
                                                    Case 1
                                                        loan2Bank = tb_細項內容.Rows(ii)("設定權利人")
                                                        set2Money = tb_細項內容.Rows(ii)("設定") & "萬"
                                                End Select
                                            End If
                                        End If
                                    End If
                                Next

                                '若無值，再撈取權力類別為建物的
                                If loanBank = "" Then
                                    ii = 0
                                    For ii = 0 To tb_細項內容.Rows.Count - 1
                                        If tb_細項內容.Rows(ii)("權利類別") = "建物" Then
                                            If Not IsDBNull(tb_細項內容.Rows(ii)("設定")) Then
                                                If tb_細項內容.Rows(ii)("設定") <> "" Then
                                                    Select Case ii
                                                        Case 0
                                                            loanBank = tb_細項內容.Rows(ii)("設定權利人")
                                                            setMoney = tb_細項內容.Rows(ii)("設定") & "萬"
                                                        Case 1
                                                            loan2Bank = tb_細項內容.Rows(ii)("設定權利人")
                                                            set2Money = tb_細項內容.Rows(ii)("設定") & "萬"
                                                    End Select
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            Else '不同
                                For ii = 0 To tb_細項內容.Rows.Count - 1
                                    If Not IsDBNull(tb_細項內容.Rows(ii)("設定")) Then
                                        If tb_細項內容.Rows(ii)("設定") <> "" Then
                                            Select Case ii
                                                Case 0
                                                    loanBank = tb_細項內容.Rows(ii)("設定權利人")
                                                    setMoney = tb_細項內容.Rows(ii)("設定") & "萬"
                                                Case 1
                                                    loan2Bank = tb_細項內容.Rows(ii)("設定權利人")
                                                    set2Money = tb_細項內容.Rows(ii)("設定") & "萬"
                                            End Select
                                        End If
                                    End If
                                Next
                            End If
                        Else '無選取
                            For ii = 0 To tb_細項內容.Rows.Count - 1
                                If Not IsDBNull(tb_細項內容.Rows(ii)("設定")) Then
                                    If tb_細項內容.Rows(ii)("設定") <> "" Then
                                        Select Case ii
                                            Case 0
                                                loanBank = tb_細項內容.Rows(ii)("設定權利人")
                                                setMoney = tb_細項內容.Rows(ii)("設定") & "萬"
                                            Case 1
                                                loan2Bank = tb_細項內容.Rows(ii)("設定權利人")
                                                set2Money = tb_細項內容.Rows(ii)("設定") & "萬"
                                        End Select
                                    End If
                                End If
                            Next
                        End If
                    End If
                    conn_他項權利.Close()
                    conn_他項權利.Dispose()
                End If

                ' -------- 竣工日期、屋齡 ----------------------
                If myDRr("竣工日期").ToString.Trim <> "" Then
                    If IsNumeric(myDRr("竣工日期").ToString) Then
                        CompleteDT = Mid(myDRr("竣工日期").ToString.Trim, 1, 3) & "年" & Mid(myDRr("竣工日期").ToString.Trim, 4, 2) & "月"
                        HouseOld = (Now.Year - 1911) - Convert.ToInt32(Mid(myDRr("竣工日期").ToString.Trim, 1, 3)) & "年"
                    Else
                        CompleteDT = ""
                        HouseOld = ""
                    End If
                End If

                ' ------ 所在樓層 ------------------------------
                If myDRr.Item("銷售樓層").ToString.Trim <> "" Or myDRr.Item("銷售樓層").ToString IsNot DBNull.Value Then
                    SiteFloor = myDRr.Item("銷售樓層").ToString.Trim & "樓"
                Else
                    SiteFloor = ""
                End If

                bulidType = Server.HtmlEncode(myDRr("建築結構").ToString.Trim)

                '------- 現況 -----------------------------
                NowState = "□空房□自用□出租"
                If myDRr.Item("現況").ToString.Trim <> "" Then
                    NowState = NowState.Replace("□" & myDRr.Item("現況").ToString.Trim, "■" & myDRr.Item("現況").ToString.Trim)
                End If

                '------- 車位形式 -----------------------------
                '1041021增加佩
                Dim i2 As Integer = 0
                Dim conn_車位 = New SqlConnection(EGOUPLOADSqlConnStr)
                conn_車位.Open()
                Dim sql1 As String = "Select * From 產調_車位 Where 物件編號='" & Request("oid") & "' and 店代號='" & Request("sid") & "' order by 流水號 "
                adpt1 = New SqlDataAdapter(sql1, conn_車位)
                ds1 = New DataSet()
                adpt1.Fill(ds1, "細項內容_車位")
                Dim tb_細項內容_車位 As DataTable = ds1.Tables("細項內容_車位")
                Dim 無_車位, 機械_車位, 平面_車位, 機械和平面_車位, 坡道平面_車位, 坡道機械_車位, 母子車位_車位, 其他_車位 As String
                Dim 無_車位說明, 機械_車位說明, 平面_車位說明, 機械和平面_車位說明, 坡道平面_車位說明, 坡道機械_車位說明, 母子車位_車位說明, 其他_車位說明 As String
                Dim 車位_編號 As String = ""

                If tb_細項內容_車位.Rows.Count > 0 Then
                    For i2 = 0 To tb_細項內容_車位.Rows.Count - 1
                        ' "無", "機械", "平面", "機械和平面", "坡道平面", "坡道機械", "母子車位"
                        If tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "無" Then
                            無_車位 += 1
                            無_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "機械" Then
                            機械_車位 += 1
                            機械_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "平面" Then
                            平面_車位 += 1
                            平面_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "機械和平面" Then
                            機械和平面_車位 += 1
                            機械和平面_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "坡道平面" Then
                            坡道平面_車位 += 1
                            坡道平面_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "坡道機械" Then
                            坡道機械_車位 += 1
                            坡道機械_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        ElseIf tb_細項內容_車位.Rows(i2)("車位類別").ToString.Trim = "母子車位" Then
                            母子車位_車位 += 1
                            母子車位_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        Else
                            其他_車位 += 1
                            其他_車位說明 += tb_細項內容_車位.Rows(i2)("車位說明").ToString.Trim
                        End If

                        If tb_細項內容_車位.Rows(i2)("車位號碼").ToString.Trim <> "" Then
                            車位_編號 &= tb_細項內容_車位.Rows(i2)("車位號碼").ToString.Trim & "、"
                        End If
                    Next
                    If 無_車位 > 0 Then
                        cartype = "無(" & 無_車位 & ") "
                        If 無_車位說明 <> "" Then
                            cartype &= "說明:" & 無_車位說明 & "、"
                        End If
                    End If

                    If 機械_車位 > 0 Then
                        cartype &= "機械(" & 機械_車位 & ") "
                        If 機械_車位說明 <> "" Then
                            cartype &= "說明:" & 機械_車位說明 & "、"
                        End If
                    End If

                    If 平面_車位 <> 0 Then
                        cartype &= "平面(" & 平面_車位 & ") "
                        If 平面_車位說明 <> "" Then
                            cartype &= "說明:" & 平面_車位說明 & "、"
                        End If
                    End If

                    If 機械和平面_車位 <> 0 Then
                        cartype &= "機械和平面(" & 機械和平面_車位 & ") "
                        If 機械和平面_車位說明 <> "" Then
                            cartype &= "說明:" & 機械和平面_車位說明 & "、"
                        End If
                    End If
                    If 坡道平面_車位 <> 0 Then
                        cartype &= "坡道平面(" & 坡道平面_車位 & ") "
                        If 坡道平面_車位說明 <> "" Then
                            cartype &= "說明:" & 坡道平面_車位說明 & "、"
                        End If
                    End If
                    If 坡道機械_車位 <> 0 Then
                        cartype &= "坡道機械(" & 坡道機械_車位 & ") "
                        If 坡道機械_車位說明 <> "" Then
                            cartype &= "說明:" & 坡道機械_車位說明 & "、"
                        End If
                    End If
                    If 母子車位_車位 <> 0 Then
                        cartype &= "母子車位(" & 母子車位_車位 & ") "
                        If 母子車位_車位說明 <> "" Then
                            cartype &= "說明:" & 母子車位_車位說明 & "、"
                        End If
                    End If
                    If 其他_車位 <> 0 Then
                        cartype &= "其他車位(" & 其他_車位 & ") "
                        If 其他_車位說明 <> "" Then
                            cartype &= "說明:" & 其他_車位說明 & "、"
                        End If
                    End If
                    carex = 車位_編號
                Else
                    If myDRr("車位").ToString.Trim = "坡道平面" Then
                        cartype = "■坡平□坡機□機平□機機"
                    ElseIf myDRr("車位").ToString.Trim = "坡道機械" Then
                        cartype = "□坡平■坡機□機平□機機"
                    ElseIf myDRr("車位").ToString.Trim = "機械和平面" Then
                        cartype = "□坡平□坡機■機平□機機"
                    ElseIf myDRr("車位").ToString.Trim = "機械" Then
                        cartype = "□坡平□坡機□機平■機機"
                    Else
                        cartype = "□坡平□坡機□機平□機機"
                    End If
                    cartype &= " 高度     CM"
                End If


                Dim conn_面積細項 = New SqlConnection(EGOUPLOADSqlConnStr)
                conn_面積細項.Open()
                Dim sql3 As String = "Select * From 委賣物件資料表_面積細項 Where 物件編號='" & Request("oid") & "' and 店代號='" & Request("sid") & "' and (類別 IN ('公共設施', '主建物', '車位面積(公設內)','車位面積(產權獨立)')) AND (項目名稱 LIKE '%車%') order by 流水號 "
                adpt1 = New SqlDataAdapter(sql3, conn_面積細項)
                ds1 = New DataSet()
                adpt1.Fill(ds1, "細項內容_車位面積細項")
                Dim car_total As Double = 0
                Dim car_total_p As Double = 0
                Dim ii2 As Integer = 0
                Dim ii3 As Integer = 0
                Dim tb_細項內容_面積細項 As DataTable = ds1.Tables("細項內容_車位面積細項")
                If tb_細項內容_面積細項.Rows.Count > 0 Then
                    For ii2 = 0 To tb_細項內容_面積細項.Rows.Count - 1
                        sql3 = " Select * From 委賣物件資料表_細項所有權人 Where 物件編號='" & tb_細項內容_面積細項.Rows(ii2)("物件編號") & "' and 店代號='" & tb_細項內容_面積細項.Rows(ii2)("店代號") & "' and 細項流水號='" & tb_細項內容_面積細項.Rows(ii2)("流水號") & "' "
                        adpt1 = New SqlDataAdapter(sql3, conn_面積細項)
                        ds1 = New DataSet()
                        adpt1.Fill(ds1, "細項內容_車位面積所有權人")
                        Dim tb_細項內容_車位面積所有權人 As DataTable = ds1.Tables("細項內容_車位面積所有權人")
                        If tb_細項內容_車位面積所有權人.Rows.Count > 0 Then
                            For ii3 = 0 To tb_細項內容_車位面積所有權人.Rows.Count - 1
                                If Not IsDBNull(tb_細項內容_車位面積所有權人.Rows(ii3)("出售面積")) And tb_細項內容_車位面積所有權人.Rows(ii3)("出售面積").ToString <> "0.0000" And tb_細項內容_車位面積所有權人.Rows(ii3)("出售面積").ToString <> "" Then
                                    'car_total += tb_細項內容_車位面積所有權人.Rows(ii2)("出售面積")
                                    car_total_p += tb_細項內容_車位面積所有權人.Rows(ii3)("出售面積") * 0.3025
                                Else
                                    car_total_p += tb_細項內容_車位面積所有權人.Rows(ii3)("持有面積") * 0.3025
                                End If
                            Next
                        Else
                            'For ii2 = 0 To tb_細項內容_面積細項.Rows.Count - 1
                            car_total += tb_細項內容_面積細項.Rows(ii2)("實際持有平方公尺")
                            car_total_p += tb_細項內容_面積細項.Rows(ii2)("實際持有坪")
                            'Next
                        End If
                    Next
                    'For ii2 = 0 To tb_細項內容_面積細項.Rows.Count - 1
                    '    car_total += tb_細項內容_面積細項.Rows(ii2)("實際持有平方公尺")
                    '    car_total_p += tb_細項內容_面積細項.Rows(ii2)("實際持有坪")
                    'Next
                    PubStallPin = Convert.ToDouble("0" & car_total_p) & "坪(共" & tb_細項內容_面積細項.Rows.Count & "個)"
                Else
                    ' 2016.04.25 Modify by Finch 增加判斷 "0" & NULL
                    If myDRr("公設內車位坪數").ToString.Trim <> "" And myDRr("公設內車位坪數").ToString.Trim <> "0" And IsDBNull(myDRr("公設內車位坪數")) = False Then
                        PubStallPin = Convert.ToDouble("0" & myDRr("公設內車位坪數").ToString.Trim) & "坪"
                    ElseIf myDRr("車位坪數").ToString.Trim <> "" And myDRr("車位坪數").ToString.Trim <> "0" And IsDBNull(myDRr("車位坪數")) = False Then
                        PubStallPin = Convert.ToDouble("0" & myDRr("車位坪數").ToString.Trim) & "坪(獨)"
                    End If
                End If

                '------- 帶看方式 -----------------------------
                'If myDRr.Item("帶看方式").ToString.Trim = "洽經紀人" Then
                '    bringLookType = "■洽經紀人　□鑰匙"
                'ElseIf myDRr.Item("帶看方式").ToString.Trim = "鑰匙" Then
                '    bringLookType = "□洽經紀人　■鑰匙"
                'End If
                bringLookType = myDRr.Item("帶看方式").ToString.Trim

                PriSchool = Server.HtmlEncode(myDRr("國小1").ToString.Trim)
                SecSchool = Server.HtmlEncode(myDRr("國中1").ToString.Trim)

                Dim i, s, j, k, m As Integer
                Dim adpt As SqlDataAdapter
                Dim ds As DataSet
                Dim table4 As DataTable
                Dim f1 As String
                If IsDBNull(myDRr("公園代號")) = False Then
                    If myDRr("公園代號") <> "" Then
                        '為了舊資料相容,要再加上.
                        s = InStr(myDRr("公園代號"), ".")
                        If s = 0 Then
                            f1 = myDRr("公園代號") + "."
                        Else
                            f1 = myDRr("公園代號")
                        End If
                        Dim item1 As Array = Split(f1, ".")
                        Dim temp As String = ""
                        For i = 0 To item1.Length - 1
                            If item1(i) <> "" Then
                                temp &= "'" & item1(i) & "',"
                                k += 1
                            End If
                        Next i
                        Dim cc As String = ""
                        cc = Mid(temp, 1, Len(temp) - 1)

                        Using conn1 As New SqlConnection(EGOUPLOADSqlConnStr)
                            conn1.Open()
                            sql = "Select 公園名稱 FROM 公園資料表 where 公園代號 in (" & cc & ") and 核准否='1'"
                            adpt = New SqlDataAdapter(sql, conn1)
                            ds = New DataSet()
                            adpt.Fill(ds, "table4")
                            table4 = ds.Tables("table4")
                            If table4.Rows.Count > 0 Then
                                i = 0
                                For i = 0 To table4.Rows.Count - 1
                                    If table4.Rows(i)("公園名稱") <> "" And IsDBNull(table4.Rows(i)("公園名稱")) = False Then
                                        parkname &= table4.Rows(i)("公園名稱") & "."
                                    End If
                                Next i
                            End If
                        End Using
                    End If
                End If

                Dim f2 As String
                If IsDBNull(myDRr("商圈代號")) = False Then
                    If myDRr("商圈代號") <> "" Then
                        '為了舊資料相容,要再加上.
                        j = InStr(myDRr("商圈代號"), ".")
                        If j = 0 Then
                            f2 = myDRr("商圈代號") + "."
                        Else
                            f2 = myDRr("商圈代號")
                        End If
                        Dim item1 As Array = Split(f2, ".")
                        Dim temp As String = ""
                        For i = 0 To item1.Length - 1
                            If item1(i) <> "" Then
                                temp &= "'" & item1(i) & "',"
                                m += 1
                            End If
                        Next i
                        Dim cc As String = ""
                        cc = Mid(temp, 1, Len(temp) - 1)

                        Using conn1 As New SqlConnection(EGOUPLOADSqlConnStr)
                            sql = "Select 商圈名稱 FROM 精華生活圈資料表 where 商圈代號 in (" & cc & ") and 核准否='1'"
                            adpt = New SqlDataAdapter(sql, conn1)
                            ds = New DataSet()
                            adpt.Fill(ds, "table4")
                            table4 = ds.Tables("table4")
                            If table4.Rows.Count > 0 Then
                                i = 0
                                For i = 0 To table4.Rows.Count - 1
                                    If table4.Rows(i)("商圈名稱") <> "" And IsDBNull(table4.Rows(i)("商圈名稱")) = False Then
                                        markname &= table4.Rows(i)("商圈名稱") & "."
                                    End If
                                Next i
                            End If
                        End Using
                    End If
                End If

                Environment = PriSchool & SecSchool & parkname & markname

                KeyNo = Server.HtmlEncode(myDRr("鑰匙編號").ToString.Trim)

                objDescrption = Server.HtmlEncode(myDRr("訴求重點").ToString.Trim)

                '----------當物件用途為土地時,總坪數即為土地坪數----------------
                If (Convert.ToDouble("0" & myDRr.Item("土地坪數")) > 0) And InStr(Server.HtmlEncode(myDRr("物件類別").ToString.Trim), "地") <> 0 Then
                    AllPin = Convert.ToDouble(myDRr.Item("土地坪數"))
                End If

                StoreBoss = Server.HtmlEncode(myDRr("店長").ToString.Trim)
                AdminiName = Server.HtmlEncode(myDRr("行政").ToString.Trim)
                Promoter = Server.HtmlEncode(myDRr("員工姓名").ToString.Trim) & " " & Server.HtmlEncode(myDRr("營業員代號1").ToString.Trim) & " " & Server.HtmlEncode(IIf(checkvalue.name(myDRr("營業員代號2").ToString.Trim) = "error", "", checkvalue.name(myDRr("營業員代號2").ToString.Trim)))
                usetype = Server.HtmlEncode(myDRr("物件用途").ToString.Trim)
                Huse = Server.HtmlEncode(myDRr("產用途").ToString.Trim)
            End While

            Dim Xml2Doc As New XmlToDoc.XmlToDoc
            Dim sFile As String = Xml2Doc.MyGetFullTextFile(Server.MapPath("\reports\word\Word物調表6New.xml"), enuStandardCodePages.SCP_CP_UTF8)
            sFile = sFile.Replace("RptDT", RptDT)
            sFile = sFile.Replace("postcardName", Server.HtmlEncode(postcardName))
            sFile = sFile.Replace("objID", objID)
            sFile = sFile.Replace("Promoter", Promoter)
            sFile = sFile.Replace("buildName", Server.HtmlEncode(buildName))
            sFile = sFile.Replace("GroundPin", GroundPin)
            sFile = sFile.Replace("TotPin", TotPin)
            sFile = sFile.Replace("SalePrice", SalePrice)
            sFile = sFile.Replace("CityArea", CityArea)
            sFile = sFile.Replace("objAddr", Server.HtmlEncode(objAddr))
            sFile = sFile.Replace("H720", H720)
            sFile = sFile.Replace("bulidPin", bulidPin)
            sFile = sFile.Replace("AttachPin", AttachPin)
            sFile = sFile.Replace("PublicPin", PublicPin)
            sFile = sFile.Replace("TotPin", TotPin)
            sFile = sFile.Replace("HouseStyle", HouseStyle)
            sFile = sFile.Replace("ManaPrice", ManaPrice)
            sFile = sFile.Replace("seatDirect", seatDirect)
            sFile = sFile.Replace("NearWidth", NearWidth)
            sFile = sFile.Replace("NearAWidth", NearAWidth)
            sFile = sFile.Replace("NearBWidth", NearBWidth)
            sFile = sFile.Replace("GroundPin", GroundPin)
            sFile = sFile.Replace("HousePin", HousePin)
            sFile = sFile.Replace("PubStallPin", PubStallPin)
            sFile = sFile.Replace("AddBuildPin", AddBuildPin)
            sFile = sFile.Replace("HomeLiftNum", HomeLiftNum)
            sFile = sFile.Replace("UpFloorNum", UpFloorNum)
            sFile = sFile.Replace("DownFloorNum", DownFloorNum)
            sFile = sFile.Replace("loanBank", loanBank)
            sFile = sFile.Replace("setMoney", setMoney)
            sFile = sFile.Replace("loan2Bank", loan2Bank)
            sFile = sFile.Replace("set2Money", set2Money)
            sFile = sFile.Replace("CompleteDT ", CompleteDT)
            sFile = sFile.Replace("HouseOld", HouseOld)
            sFile = sFile.Replace("SiteFloor", SiteFloor)
            sFile = sFile.Replace("bulidType", bulidType)
            sFile = sFile.Replace("bulidPin", bulidPin)
            sFile = sFile.Replace("NowState", NowState)
            sFile = sFile.Replace("cartypes", Server.HtmlEncode(cartype))
            sFile = sFile.Replace("carex", carex)
            sFile = sFile.Replace("Environment", Server.HtmlEncode(Environment))
            sFile = sFile.Replace("bringLookType", bringLookType)
            sFile = sFile.Replace("KeyNo", KeyNo)
            sFile = sFile.Replace("objDescrption", Server.HtmlEncode(objDescrption))
            sFile = sFile.Replace("CompleteDT", CompleteDT)
            'sFile = sFile.Replace("cartype", cartype)
            sFile = sFile.Replace("StoreBoss", StoreBoss)
            sFile = sFile.Replace("AdminiName", AdminiName)
            sFile = sFile.Replace("Promoter", Promoter)
            sFile = sFile.Replace("usetype", usetype)
            sFile = sFile.Replace("Huse", Huse)
            Dim paths As String = ""
            Dim newpaths As String = Server.MapPath("~/images/newobject.jpg")
            Dim Img_byte() As Byte
            Dim ImgObj As Object


            Dim newWidth As Double
            Dim newHeight As Double

            ' 地圖部份 ============================================================

            '判別網路上有沒有地圖

            'If Request("oid") = "61319AAE32716" Then
            '    'Response.Write(sql)
            '    'Response.End()
            'Else

            'End If

            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "m.jpg"

            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "m.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg")

                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        'If image.Width > 1024 Or image.Height > 715 Then '480,360
                        newWidth = 266.25
                        newHeight = 266.25
                        'If image.Width > image.Height Then
                        '    newWidth = 189
                        '    newHeight = image.Height / (image.Width / 189)
                        'Else
                        '    newHeight = 253.5
                        '    newWidth = image.Width / (image.Height / 253.5)
                        'End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()

                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "m.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic1", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠地圖寬", newWidth & "pt")
            sFile = sFile.Replace("≠地圖高", newHeight & "pt")

            ' 隔局圖部份 ============================================================

            '判別網路上有沒有隔局圖
            newpaths = Server.MapPath("~/images/newobject.jpg")
            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "g.jpg"

            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "g.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg")

                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        'If image.Width > 1024 Or image.Height > 715 Then '480,360
                        If (image.Width / image.Height) <= 1 Then
                            newHeight = 267.75
                            newWidth = image.Width / (image.Height / 267.75)
                        Else
                            newWidth = 267.75
                            newHeight = image.Height / (image.Width / 267.75)
                        End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()
                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "g.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic2", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠格局寬", newWidth & "pt")
            sFile = sFile.Replace("≠格局高", newHeight & "pt")

            ' 照片1部份 ============================================================

            '判別網路上有沒有照片1
            newpaths = Server.MapPath("~/images/newobject.jpg")
            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "a.jpg"
            newWidth = 261
            newHeight = 182
            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "a.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg")

                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        If (image.Width / image.Height) <= 1.4375 Then
                            newHeight = 182
                            newWidth = image.Width / (image.Height / 182)
                        Else
                            newWidth = 261
                            newHeight = image.Height / (image.Width / 261)
                        End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()

                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "a.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic3", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠寬1", newWidth & "pt")
            sFile = sFile.Replace("≠高1", newHeight & "pt")

            ' 照片2部份 ============================================================

            '判別網路上有沒有照片2
            newpaths = Server.MapPath("~/images/newobject.jpg")
            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "b.jpg"
            newWidth = 261
            newHeight = 182
            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "b.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg")

                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        If (image.Width / image.Height) <= 1.4375 Then
                            newHeight = 182
                            newWidth = image.Width / (image.Height / 182)
                        Else
                            newWidth = 261
                            newHeight = image.Height / (image.Width / 261)
                        End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()
                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "b.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic4", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠寬2", newWidth & "pt")
            sFile = sFile.Replace("≠高2", newHeight & "pt")
            ' 照片3部份 ============================================================

            '判別網路上有沒有照片3
            newpaths = Server.MapPath("~/images/newobject.jpg")
            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "c.jpg"
            newWidth = 261
            newHeight = 182
            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "c.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg")
                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        If (image.Width / image.Height) <= 1.4375 Then
                            newHeight = 182
                            newWidth = image.Width / (image.Height / 182)
                        Else
                            newWidth = 261
                            newHeight = image.Height / (image.Width / 261)
                        End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()
                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "c.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic5", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠寬3", newWidth & "pt")
            sFile = sFile.Replace("≠高3", newHeight & "pt")

            '判別網路上有沒有照片4
            newpaths = Server.MapPath("~/images/newobject.jpg")
            paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label3.Text & "/" & Contract_No & "d.jpg"
            newWidth = 261
            newHeight = 182
            Try
                Dim requests As WebRequest = HttpWebRequest.Create(paths)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists(Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg") Then
                            File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg")
                        End If

                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label3.Text & "/" & Contract_No & "d.jpg", Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg")

                        Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg")
                        Dim sFileName As String = Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg"
                        Dim sFileStream As FileStream = File.OpenRead(sFileName)
                        Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                        If (image.Width / image.Height) <= 1.4375 Then
                            newHeight = 182
                            newWidth = image.Width / (image.Height / 182)
                        Else
                            newWidth = 261
                            newHeight = image.Height / (image.Width / 261)
                        End If
                        image.Dispose()
                        sFileStream.Close()
                        sFileStream.Dispose()
                        '20110412刪除UPFILE裡的物件
                        File.Delete(Server.MapPath("~/webfly/Upfile/") & Contract_No & "d.jpg")
                    End If
                End Using
            Catch ex As WebException
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
                'End If
            End Try

            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("objPic6", Xml2Doc.MyBase64Encode(ImgObj))
            sFile = sFile.Replace("≠寬4", newWidth & "pt")
            sFile = sFile.Replace("≠高4", newHeight & "pt")

            '品牌
            If InStr(Request("sid").ToUpper, "S") = 0 Then
                Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/images/ObjInquire_etwarm.png"))
            Else
                Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/images/ObjInquire_life.png"))
            End If
            '將圖檔轉成Base64Code後置換
            ImgObj = Img_byte
            sFile = sFile.Replace("logo", Xml2Doc.MyBase64Encode(ImgObj))

            '下載檔案
            Dim arrBytes = Xml2Doc.StringToBytes(sFile, enuStandardCodePages.SCP_CP_UTF8)

            Dim randpath As String = TempDictforZip & RandomDictName
            If Directory.Exists(randpath) = False Then
                Directory.CreateDirectory(TempDictforZip & RandomDictName)
            End If

            '存檔
            File.WriteAllBytes(TempDictforZip & RandomDictName & "\" & Contract_No & "_物調表.doc", arrBytes)

            myDRr.Close()
            myDRr = Nothing

            Cmd.Dispose()
            Cmd = Nothing
        End Using
    End Sub

    Sub ddl產生路名()
        Dim table1 As DataTable
        Dim table2 As DataTable
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        '物件地址
        sql = "Select 郵遞區號,縣市, 鄉鎮市區 From " & src.Text
        sql &= "  With(NoLock) Where 物件編號 = '" & Request("oid") & "' "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")

        '郵遞區號
        sql = "Select 郵遞區號 From 資料_鄉鎮市區  With(NoLock) "
        sql &= "Where 縣市名 = '" & table2.Rows(0)("縣市") & "' "
        sql &= "And 鄉鎮市區 = '" & table2.Rows(0)("鄉鎮市區") & "'"
        stree1.Text = table2.Rows(0)("縣市")
        stree2.Text = table2.Rows(0)("鄉鎮市區")
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        If table1.Rows.Count <> 0 Then
            'txt郵遞區號.Text = table1.Rows(0)("郵遞區號")
            '產生路名
            sql = "SELECT 名稱 from 資料_道路  With(NoLock) WHERE  郵遞區號='" & table1.Rows(0)("郵遞區號") & "' order by 名稱"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            Dim zipcode As DataTable = ds.Tables("table1")
            With ddl路名
                .DataSource = zipcode.DefaultView
                .DataTextField = "名稱"
                .DataValueField = "名稱"
                .DataBind()
            End With
            ddl路名.Items.Insert(0, New ListItem("請選擇", "請選擇"))
            With ddl路名2
                .DataSource = zipcode.DefaultView
                .DataTextField = "名稱"
                .DataValueField = "名稱"
                .DataBind()
            End With
            ddl路名2.Items.Insert(0, New ListItem("請選擇", "請選擇"))
            With ddl路名3
                .DataSource = zipcode.DefaultView
                .DataTextField = "名稱"
                .DataValueField = "名稱"
                .DataBind()
            End With
            ddl路名3.Items.Insert(0, New ListItem("請選擇", "請選擇"))
        End If

        conn.Close()
        conn = Nothing
    End Sub

    Sub open_crt()
        'Dim Fname As String
        Dim sid As String = Request("sid")

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        Dim Contract_No As String = Request("oid")

        Dim sql As String

        sql = "SELECT ROUND(a.主建物,2, 1) as 主建物,ROUND(a.附屬建物,2, 1) as 附屬建物,ROUND(a.公共設施,2, 1) as 公共設施,ROUND(a.公設內車位坪數,2, 1) as 公設內車位坪數,ROUND(a.地下室,2, 1) as 地下室,ROUND(a.總坪數,2, 1) as 總坪數,ROUND(a.車位坪數,2, 1) as 車位坪數,a.物件用途 as 使用管制內容,a.其他使用分區,b.其他外牆外飾 as 其他外牆外飾,b.其他地板 AS 其他地板,b.其他項目 AS 其他項目,c.bs_dept_brief AS 店名,c.bs_comp_name AS 法人, c.bs_oarea + '-' + c.bs_fax AS 傳真, "
        sql &= " c.bs_email AS 電子郵件,c.bs_email2 as 電子郵件2,c.bs_email_flag,d.縣市名 + d.鄉鎮市區 + c.bs_c_adds as 店_完整地址,c.bs_oarea + '-' + c.bs_otel as 連絡電話, "
        sql &= " a.縣市 + a.鄉鎮市區 + a.完整地址 as 物件_完整地址,a.縣市 + a.鄉鎮市區 + a.部份地址 as 物件_部份地址, e.縣市 + e.鄉鎮市區 + e.完整地址 as 客戶_完整地址,"
        sql &= " g.man_name as 經紀人,h.man_name as 營業員1,i.man_name as 營業員2, "
        sql &= " b.法定建蔽率 + '%' as 法定建蔽率,b.法定容積率 + '%' as  法定容積率,b.契稅 as 房地產契稅,b.管理費 as 房地產管理費,b.建築用途 as 物件用途,a.物件用途 as true物件用途,a.物件類別, "
        sql &= " a.房 as 房,a.廳 as 廳,ROUND(a.衛, 1, 1) as 衛,a.室 as 室,b.冷氣台 as 冷氣台,b.冰箱台 as 冰箱台,洗衣機台,乾衣機台,飲水機,b.地板,* "
        sql &= "FROM " & src.Text & " a With(NoLock) "  '20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
        sql &= "inner JOIN 委賣_房地產說明書 b With(NoLock) ON a.物件編號 = b.物件編號 and a.店代號 = b.店代號 "
        sql &= "Left JOIN HSBSMG c With(NoLock) ON a.店代號 = c.bs_dept "
        sql &= "Left JOIN 資料_鄉鎮市區 d With(NoLock) ON d.郵遞區號 = c.bs_c_area "
        sql &= "Left Join 委賣屋主資料表 e With(NoLock) ON e.物件編號 = a.物件編號 and e.店代號 = b.店代號  "
        sql &= "Left JOIN psman g With(NoLock) ON g.man_emp_no = a.經紀人代號 "
        sql &= "Left JOIN psman h With(NoLock) ON h.man_emp_no = a.營業員代號1 "
        sql &= "Left JOIN psman i With(NoLock) ON i.man_emp_no = a.營業員代號2 "
        sql &= "Left JOIN 委賣地圖 j With(NoLock) ON a.物件編號 = j.物件編號 and a.店代號 = j.店代號 "
        sql &= "Where a.物件編號 = '" & Contract_No & "' and a.店代號= '" & Request("sid") & "'"

        'If Contract_No = "10619BAA52141" Then
        '    Response.Write(sql)
        '    Exit Sub
        'End If
        'If Request.Cookies("webfly_empno").Value.ToLower = "92H" Then
        '    Response.Write(sql)
        '    Response.End()
        'End If

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet
        adpt.Fill(ds, "new_store")
        Dim t2 As DataTable = ds.Tables("new_store")

        '判斷有無填寫不動產說明書(舊版完全未填寫，會出錯)
        If t2.Rows.Count = 0 Then
            Dim Script As String = ""
            Script += "alert('查無相關不動產說明書內容，煩請先補足相關內容再行列印');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)
            Exit Sub
        End If

        '1000325小豪新增-相片資料夾區分為過期(expired).有效(available)
        'If t2.Rows(0).Item("委託截止日") >= (Format(Year(Today) - 1911, "00#") & Format(Month(Today), "0#") & Format(Day(Today), "0#")) Then
        Me.Label2.Text = "available"
        'Else
        '    Me.Label2.Text = "expired"
        'End If

        Dim Cmd1, Cmd2, msg1, msg2

        '共用
        '[START============================================封面+目錄=============================================START]
        '頁首
        Dim 頁首編號及案名 As String = ""
        '封面
        Dim 委託編號 As String = ""
        Dim 封面案名 As String = ""
        Dim 說明書類別 As String = ""
        Dim 報告日期 As String = "" ' 預設先不用
        Dim 店名 As String = ""
        Dim 法人 As String = ""
        Dim 店完整地址 As String = ""
        Dim 連絡電話 As String = ""
        Dim 傳真 As String = ""
        Dim 電子郵件 As String = ""
        Dim 承辦人 As String = ""
        '目錄-內容
        Dim 產權調查 As String = ""
        Dim 物件個案調查 As String = ""
        Dim 照片說明 As String = ""
        Dim 重要交易條件 As String = ""
        Dim 成交行情參考表 As String = ""
        Dim 重要設施參考圖 As String = ""
        Dim 其他說明 As String = ""
        '目錄-附件
        '第一列
        Dim 土地權狀影本 As String = ""
        Dim 建物權狀影本 As String = ""
        Dim 房地產標的現況說明書 As String = ""
        '第二列
        Dim 土地謄本 As String = ""
        Dim 建物謄本 As String = ""
        Dim 預售屋買賣契約書 As String = ""
        '第三列
        Dim 土地地籍圖 As String = ""
        Dim 建物測量成果圖 As String = ""
        Dim 住戶規約 As String = ""
        '第四列
        Dim 土地相關位置略圖 As String = ""
        Dim 建物相關位置略圖 As String = ""
        Dim 停車位位置圖 As String = ""
        '第五列
        Dim 土地分管協議 As String = ""
        Dim 建物分管協議 As String = ""
        Dim 樑柱顯見裂痕照片 As String = ""
        '第六列
        Dim 分區使用說明 As String = ""
        Dim 房屋稅單 As String = ""
        Dim 封面其他 As String = ""
        '第七列
        Dim 增值稅概算 As String = ""
        Dim 使用執照 As String = ""

        Dim sales2, sales3 As String
        If Not IsDBNull(t2.Rows(0)("營業員1")) Then
            sales2 = "," & t2.Rows(0)("營業員1").ToString
        Else
            sales2 = ""
        End If
        If Not IsDBNull(t2.Rows(0)("營業員2")) Then
            sales3 = "," & t2.Rows(0)("營業員2").ToString
        Else
            sales3 = ""
        End If

        承辦人 = t2.Rows(0)("經紀人").ToString & sales2 & sales3

        委託編號 = t2.Rows(0)("物件編號").ToString

        '封面案名
        If IsDBNull(t2.Rows(0)("張貼卡案名").ToString) Or t2.Rows(0)("張貼卡案名").ToString.Trim = "" Then
            封面案名 = Server.HtmlEncode(t2.Rows(0)("建築名稱").ToString)
        Else
            封面案名 = Server.HtmlEncode(t2.Rows(0)("張貼卡案名").ToString)
        End If

        '頁首編號及案名
        頁首編號及案名 = 封面案名 & "  " & Mid(t2.Rows(0)("物件編號").ToString, 6)

        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then
            說明書類別 = "成屋不動產說明書"
        Else
            說明書類別 = "土地不動產說明書"
        End If

        店名 = t2.Rows(0)("店名").ToString
        法人 = t2.Rows(0)("法人").ToString
        店完整地址 = t2.Rows(0)("店_完整地址").ToString
        連絡電話 = t2.Rows(0)("連絡電話").ToString
        傳真 = t2.Rows(0)("傳真").ToString
        'If Contract_No = "60985AAD36025" Then
        '    Response.Write("_" & 傳真 & "_")
        '    Exit Sub
        'End If
        If t2.Rows(0)("bs_email_flag") = 1 Then
            電子郵件 = t2.Rows(0)("電子郵件2").ToString
        Else
            電子郵件 = t2.Rows(0)("電子郵件").ToString
        End If

        '產權調查
        If IsDBNull(t2.Rows(0)("產權調查")) = True Then
            產權調查 = "□產權調查表"
        Else
            If t2.Rows(0)("產權調查").ToString = "" Then
                產權調查 = "□產權調查表"
            Else
                If t2.Rows(0)("產權調查") = 0 Then
                    產權調查 = "□產權調查表"
                Else
                    產權調查 = "█產權調查表"
                End If
            End If
        End If

        '物件個案調查
        If IsDBNull(t2.Rows(0)("物件個案調查")) = True Then
            物件個案調查 = "□物件個案調查表"
        Else
            If t2.Rows(0)("物件個案調查").ToString = "" Then
                物件個案調查 = "□物件個案調查表"
            Else
                If t2.Rows(0)("物件個案調查") = 0 Then
                    物件個案調查 = "□物件個案調查表"
                Else
                    物件個案調查 = "█物件個案調查表"
                End If
            End If
        End If

        '照片說明
        If IsDBNull(t2.Rows(0)("照片說明")) = True Then
            照片說明 = "□物件個案照片"
        Else
            If t2.Rows(0)("照片說明").ToString = "" Then
                照片說明 = "□物件個案照片"
            Else
                If t2.Rows(0)("照片說明") = 0 Then
                    照片說明 = "□物件個案照片"
                Else
                    照片說明 = "█物件個案照片"
                End If
            End If
        End If

        '重要交易條件
        If IsDBNull(t2.Rows(0)("重要交易條件")) = True Then
            重要交易條件 = "□重要交易條件"
        Else
            If t2.Rows(0)("重要交易條件").ToString = "" Then
                重要交易條件 = "□重要交易條件"
            Else
                If t2.Rows(0)("重要交易條件") = 0 Then
                    重要交易條件 = "□重要交易條件"
                Else
                    重要交易條件 = "█重要交易條件"
                End If
            End If
        End If

        '成交行情參考表
        If IsDBNull(t2.Rows(0)("成交行情")) = True Then
            成交行情參考表 = "□成交行情參考表"
        Else
            If t2.Rows(0)("成交行情").ToString = "" Then
                成交行情參考表 = "□成交行情參考表"
            Else
                If t2.Rows(0)("成交行情") = 0 Then
                    成交行情參考表 = "□成交行情參考表"
                Else
                    成交行情參考表 = "█成交行情參考表"
                End If
            End If
        End If

        '重要設施參考表
        If IsDBNull(t2.Rows(0)("重要設施")) = True Then
            重要設施參考圖 = "□鄰近重要設施參考表"
        Else
            If t2.Rows(0)("重要設施").ToString = "" Then
                重要設施參考圖 = "□鄰近重要設施參考表"
            Else
                If t2.Rows(0)("重要設施") = 0 Then
                    重要設施參考圖 = "□鄰近重要設施參考表"
                Else
                    重要設施參考圖 = "█鄰近重要設施參考表"
                End If
            End If
        End If

        '其他說明
        If IsDBNull(t2.Rows(0)("其他說明")) = True Then
            其他說明 = "□其他說明"
        Else
            If t2.Rows(0)("其他說明").ToString = "" Then
                其他說明 = "□其他說明"
            Else
                If t2.Rows(0)("其他說明") = 0 Then
                    其他說明 = "□其他說明"
                Else
                    其他說明 = "█其他說明"
                End If
            End If
        End If

        '第一列------------------------------------------------------------------------------------
        '土地權狀影本
        If IsDBNull(t2.Rows(0)("土地權狀影本")) = True Then
            土地權狀影本 = "□土地權狀(影本)"
        Else
            If t2.Rows(0)("土地權狀影本").ToString = "" Then
                土地權狀影本 = "□土地權狀(影本)"
            Else
                If t2.Rows(0)("土地權狀影本") = 0 Then
                    土地權狀影本 = "□土地權狀(影本)"
                Else
                    土地權狀影本 = "█土地權狀(影本)"
                End If
            End If
        End If

        '建物權狀影本
        If IsDBNull(t2.Rows(0)("建物權狀影本")) = True Then
            建物權狀影本 = "□建物權狀(影本)"
        Else
            If t2.Rows(0)("建物權狀影本").ToString = "" Then
                建物權狀影本 = "□建物權狀(影本)"
            Else
                If t2.Rows(0)("建物權狀影本") = 0 Then
                    建物權狀影本 = "□建物權狀(影本)"
                Else
                    建物權狀影本 = "█建物權狀(影本)"
                End If
            End If
        End If

        '標的現況說明書
        If IsDBNull(t2.Rows(0)("房地產標的現況說明書")) = True Then
            房地產標的現況說明書 = "□標的現況說明書"
        Else
            If t2.Rows(0)("房地產標的現況說明書").ToString = "" Then
                房地產標的現況說明書 = "□標的現況說明書"
            Else
                If t2.Rows(0)("房地產標的現況說明書") = 0 Then
                    房地產標的現況說明書 = "□標的現況說明書"
                Else
                    房地產標的現況說明書 = "█標的現況說明書"
                End If
            End If
        End If

        '第二列------------------------------------------------------------------------------------
        '土地謄本
        If IsDBNull(t2.Rows(0)("土地謄本")) = True Then
            土地謄本 = "□土地謄本"
        Else
            If t2.Rows(0)("土地謄本").ToString = "" Then
                土地謄本 = "□土地謄本"
            Else
                If t2.Rows(0)("土地謄本") = 0 Then
                    土地謄本 = "□土地謄本"
                Else
                    土地謄本 = "█土地謄本"
                End If
            End If
        End If

        '建物謄本
        If IsDBNull(t2.Rows(0)("建物謄本")) = True Then
            建物謄本 = "□建物謄本"
        Else
            If t2.Rows(0)("建物謄本").ToString = "" Then
                建物謄本 = "□建物謄本"
            Else
                If t2.Rows(0)("建物謄本") = 0 Then
                    建物謄本 = "□建物謄本"
                Else
                    建物謄本 = "█建物謄本"
                End If
            End If
        End If

        '預售屋買賣契約書
        If IsDBNull(t2.Rows(0)("預售買賣契約書")) = True Then
            預售屋買賣契約書 = "□委託銷售契約書"
        Else
            If t2.Rows(0)("預售買賣契約書").ToString = "" Then
                預售屋買賣契約書 = "□委託銷售契約書"
            Else
                If t2.Rows(0)("預售買賣契約書") = 0 Then
                    預售屋買賣契約書 = "□委託銷售契約書"
                Else
                    預售屋買賣契約書 = "█委託銷售契約書"
                End If
            End If
        End If

        '第三列------------------------------------------------------------------------------------
        '土地地籍圖
        If IsDBNull(t2.Rows(0)("地籍圖")) = True Then
            土地地籍圖 = "□土地地籍圖"
        Else
            If t2.Rows(0)("地籍圖").ToString = "" Then
                土地地籍圖 = "□土地地籍圖"
            Else
                If t2.Rows(0)("地籍圖") = 0 Then
                    土地地籍圖 = "□土地地籍圖"
                Else
                    土地地籍圖 = "█土地地籍圖"
                End If
            End If
        End If

        '建物勘測成果圖
        If IsDBNull(t2.Rows(0)("建物勘測成果圖")) = True Then
            建物測量成果圖 = "□建物測量成果圖"
        Else
            If t2.Rows(0)("建物勘測成果圖").ToString = "" Then
                建物測量成果圖 = "□建物測量成果圖"
            Else
                If t2.Rows(0)("建物勘測成果圖") = 0 Then
                    建物測量成果圖 = "□建物測量成果圖"
                Else
                    建物測量成果圖 = "█建物測量成果圖"
                End If
            End If
        End If

        '住戶規約
        If IsDBNull(t2.Rows(0)("住戶規約")) = True Then
            住戶規約 = "□住戶規約"
        Else
            If t2.Rows(0)("住戶規約").ToString = "" Then
                住戶規約 = "□住戶規約"
            Else
                If t2.Rows(0)("住戶規約") = 0 Then
                    住戶規約 = "□住戶規約"
                Else
                    住戶規約 = "█住戶規約"
                End If
            End If
        End If

        '第四列------------------------------------------------------------------------------------
        '土地相關位置略圖
        If IsDBNull(t2.Rows(0)("土地相關位置略圖")) = True Then
            土地相關位置略圖 = "□土地相關位置略圖"
        Else
            If t2.Rows(0)("土地相關位置略圖").ToString = "" Then
                土地相關位置略圖 = "□土地相關位置略圖"
            Else
                If t2.Rows(0)("土地相關位置略圖") = 0 Then
                    土地相關位置略圖 = "□土地相關位置略圖"
                Else
                    土地相關位置略圖 = "█土地相關位置略圖"
                End If
            End If
        End If

        '建物相關位置略圖
        If IsDBNull(t2.Rows(0)("建物相關位置略圖")) = True Then
            建物相關位置略圖 = "□建物相關位置略圖"
        Else
            If t2.Rows(0)("建物相關位置略圖").ToString = "" Then
                建物相關位置略圖 = "□建物相關位置略圖"
            Else
                If t2.Rows(0)("建物相關位置略圖") = 0 Then
                    建物相關位置略圖 = "□建物相關位置略圖"
                Else
                    建物相關位置略圖 = "█建物相關位置略圖"
                End If
            End If
        End If

        '停車位位置圖
        If IsDBNull(t2.Rows(0)("停車位位置圖")) = True Then
            停車位位置圖 = "□停車位位置圖"
        Else
            If t2.Rows(0)("停車位位置圖").ToString = "" Then
                停車位位置圖 = "□停車位位置圖"
            Else
                If t2.Rows(0)("停車位位置圖") = 0 Then
                    停車位位置圖 = "□停車位位置圖"
                Else
                    停車位位置圖 = "█停車位位置圖"
                End If
            End If
        End If

        '第伍列------------------------------------------------------------------------------------
        '土地分管協議
        If IsDBNull(t2.Rows(0)("土地分管協議")) = True Then
            土地分管協議 = "□土地分管協議"
        Else
            If t2.Rows(0)("土地分管協議").ToString = "" Then
                土地分管協議 = "□土地分管協議"
            Else
                If t2.Rows(0)("土地分管協議") = 0 Then
                    土地分管協議 = "□土地分管協議"
                Else
                    土地分管協議 = "█土地分管協議"
                End If
            End If
        End If

        '建物分管協議
        If IsDBNull(t2.Rows(0)("建物分管協議")) = True Then
            建物分管協議 = "□建物分管協議"
        Else
            If t2.Rows(0)("建物分管協議").ToString = "" Then
                建物分管協議 = "□建物分管協議"
            Else
                If t2.Rows(0)("建物分管協議") = 0 Then
                    建物分管協議 = "□建物分管協議"
                Else
                    建物分管協議 = "█建物分管協議"
                End If
            End If
        End If

        '樑柱顯見裂痕照片
        If IsDBNull(t2.Rows(0)("樑柱顯見裂痕照片")) = True Then
            樑柱顯見裂痕照片 = "□樑柱顯見裂痕照片"
        Else
            If t2.Rows(0)("樑柱顯見裂痕照片").ToString = "" Then
                樑柱顯見裂痕照片 = "□樑柱顯見裂痕照片"
            Else
                If t2.Rows(0)("樑柱顯見裂痕照片") = 0 Then
                    樑柱顯見裂痕照片 = "□樑柱顯見裂痕照片"
                Else
                    樑柱顯見裂痕照片 = "█樑柱顯見裂痕照片"
                End If
            End If
        End If

        '第六列------------------------------------------------------------------------------------
        '分區使用說明
        If IsDBNull(t2.Rows(0)("分區使用證明")) = True Then
            分區使用說明 = "□土地分區種類說明"
        Else
            If t2.Rows(0)("分區使用證明").ToString = "" Then
                分區使用說明 = "□土地分區種類說明"
            Else
                If t2.Rows(0)("分區使用證明") = 0 Then
                    分區使用說明 = "□土地分區種類說明"
                Else
                    分區使用說明 = "█土地分區種類說明"
                End If
            End If
        End If

        '房屋稅籍相關證明
        If IsDBNull(t2.Rows(0)("房屋稅單")) = True Then
            房屋稅單 = "□房屋稅籍相關證明"
        Else
            If t2.Rows(0)("房屋稅單").ToString = "" Then
                房屋稅單 = "□房屋稅籍相關證明"
            Else
                If t2.Rows(0)("房屋稅單") = 0 Then
                    房屋稅單 = "□房屋稅籍相關證明"
                Else
                    房屋稅單 = "█房屋稅籍相關證明"
                End If
            End If
        End If

        '封面其他
        If IsDBNull(t2.Rows(0)("其他")) = True Then
            封面其他 = "□其他"
        Else
            If t2.Rows(0)("其他").ToString = "" Then
                封面其他 = "□其他"
            Else
                If t2.Rows(0)("其他") = 0 Then
                    封面其他 = "□其他"
                Else
                    封面其他 = "█其他"
                End If
            End If
        End If

        '第七列------------------------------------------------------------------------------------
        '增值稅概算
        If IsDBNull(t2.Rows(0)("增值稅概算")) = True Then
            增值稅概算 = "□土地增值稅概算表"
        Else
            If t2.Rows(0)("增值稅概算").ToString = "" Then
                增值稅概算 = "□土地增值稅概算表"
            Else
                If t2.Rows(0)("增值稅概算") = 0 Then
                    增值稅概算 = "□土地增值稅概算表"
                Else
                    增值稅概算 = "█土地增值稅概算表"
                End If
            End If
        End If

        '使用執照
        If IsDBNull(t2.Rows(0)("使用執照")) = True Then
            使用執照 = "□使用執照(影本)"
        Else
            If t2.Rows(0)("使用執照").ToString = "" Then
                使用執照 = "□使用執照(影本)"
            Else
                If t2.Rows(0)("使用執照") = 0 Then
                    使用執照 = "□使用執照(影本)"
                Else
                    使用執照 = "█使用執照(影本)"
                End If
            End If
        End If

        '[END============================================封面+目錄=============================================END]

        '產權調查-建物標示說明P1參數
        Dim 建物門牌 As String = ""
        Dim 建物所在樓層 As String = ""
        Dim 建築用途 As String = ""
        Dim 建築結構 As String = ""
        Dim 建築完成日期 As String = ""
        Dim 屋齡 As String = ""
        Dim 權利種類說明 As String = ""
        Dim 所有權人姓名 As String = ""
        Dim 格局 As String = ""
        Dim 建物座落 As String = ""
        '產權調查-土地標示說明P1參數
        'Dim 土地筆數 As String = ""
        Dim 土地權利種類 As String = ""
        Dim 土地開發方式限制或其他負擔 As String = ""
        '共用
        Dim 委託價格 As String = ""

        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then '------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-建物標示說明P1=============================================START]
            If CheckBox8.Checked = True Then
                建物門牌 = t2.Rows(0)("物件_部份地址").ToString
            Else
                建物門牌 = t2.Rows(0)("物件_完整地址").ToString
            End If

            If Right(t2.Rows(0)("刊登售價").ToString, 5) = ".0000" Then
                委託價格 = Int(t2.Rows(0)("刊登售價").ToString) & " 萬元整"
            Else
                委託價格 = t2.Rows(0)("刊登售價").ToString & " 萬元整"
            End If

            建物所在樓層 = "地上共" & t2.Rows(0)("地上層數").ToString & "層"
            If Not IsDBNull(t2.Rows(0)("物件類別")) Then
                If Trim(t2.Rows(0)("物件類別").ToString) = "透天" And Trim(t2.Rows(0)("銷售樓層").ToString) = "全" Then
                    '建物所在樓層 += "，本案建物在" & t2.Rows(0)("所在樓層").ToString & "層"
                    建物所在樓層 += "，本案整棟出售"
                Else
                    '建物所在樓層 += "，本案建物在" & t2.Rows(0)("所在樓層").ToString & "層"
                    建物所在樓層 += "，本案建物在" & Trim(t2.Rows(0)("銷售樓層").ToString) & "層"
                End If
            End If
            建築用途 = t2.Rows(0)("物件用途").ToString

            建築結構 = t2.Rows(0)("建築結構").ToString

            If Not IsDBNull(t2.Rows(0)("竣工日期")) Then
                If t2.Rows(0)("竣工日期").ToString.Trim <> "" Then
                    建築完成日期 = "民國" & Left(t2.Rows(0)("竣工日期").ToString.Trim, 3) & "年" & Right(t2.Rows(0)("竣工日期").ToString.Trim, 2) & "月"
                End If
            Else
                建築完成日期 = ""
            End If

            屋齡 = IIf(t2.Rows(0)("竣工日期").ToString <> "", Val(Right("000" & Year(Now) - 1911, 3)) - Val(Left(t2.Rows(0)("竣工日期").ToString, 3)), " ")

            If 建築完成日期 <> "" Then
                建築完成日期 &= "，屋齡約" & 屋齡 & "年"
            Else
                建築完成日期 &= "屋齡約" & 屋齡 & "年"
            End If

            '權利種類說明 = "所有權"

            '建物所有權人姓名 改抓 委賣物件資料表_細項所有權人 20160414 by nick
            Dim sqlstr As String = ""
            sqlstr = "SELECT a.所有權人,a.權利範圍_分子,a.權利範圍_分母 ,a.權利總類, a.所有權種類,a.權利人流水號 FROM 委賣物件資料表_細項所有權人 a JOIN 委賣物件資料表_面積細項 b ON a.物件編號 = b.物件編號 AND a.店代號 = b.店代號 AND a.細項流水號 = b.流水號 and a.店代號=b.店代號 WHERE b.類別 = '主建物' and b.DL_level2_selectindex = '0' and  (a.物件編號 = '" & Contract_No & "') and a.店代號 = '" & Request("sid") & "' "
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                Using cmd As New SqlCommand(sqlstr, conn)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader())
                    '所有權人姓名
                    Dim columns As String() = {"所有權人", "權利範圍_分子", "權利範圍_分母", "權利總類", "所有權種類", "權利人流水號"}
                    Dim dt_name As DataTable = dt.DefaultView.ToTable(True, columns)
                    Dim i As Integer = 0
                    For Each dr As DataRow In dt_name.Rows
                        '是否顯示完整客戶姓名
                        If Not IsDBNull(dr("所有權人")) Then
                            If CheckBox4.Checked = True Then
                                If i = 0 Then
                                    所有權人姓名 = Mid(Replace(dr("所有權人"), "\", ""), 1, 1) '& "先生/小姐"
                                    '加入所有權持份
                                Else
                                    所有權人姓名 &= "、" & Mid(Replace(dr("所有權人"), "\", ""), 1, 1) '& "先生/小姐"
                                End If
                            Else
                                If i = 0 Then
                                    所有權人姓名 = Replace(dr("所有權人"), "\", "") '& "先生/小姐"
                                Else
                                    所有權人姓名 &= "、" & Replace(dr("所有權人"), "\", "") '& "先生/小姐"
                                End If
                            End If
                            '加入所有權持份
                            所有權人姓名 &= "," & dr("權利總類") & "持分" & dr("權利範圍_分子") & "/" & dr("權利範圍_分母") & "<w:br/>"
                        End If
                        If Not IsDBNull(dr("所有權種類")) Then
                            權利種類說明 = dr("所有權種類")
                        End If
                        i = i + 1
                    Next
                End Using
            End Using

            'Dim sqlstr As String = ""
            'sqlstr = "Select 客戶姓名,縣市,鄉鎮市區,完整地址 from 委賣屋主資料表  With(NoLock) where 物件編號='" & Contract_No & "' and 店代號 ='" & Request("sid") & "'"
            'sqlstr += " union all "
            'sqlstr += "Select 客戶姓名,縣市,鄉鎮市區,完整地址 from 委賣代理人資料表  With(NoLock) where 物件編號='" & Contract_No & "' and 店代號 ='" & Request("sid") & "' and 同為屋主='Y'"

            'Dim table1 As DataTable

            'Dim i As Integer = 0
            'adpt = New SqlDataAdapter(sqlstr, conn)
            'ds = New DataSet()
            'adpt.Fill(ds, "table1")
            'table1 = ds.Tables("table1")

            'If table1.Rows.Count > 0 Then
            '    For i = 0 To table1.Rows.Count - 1

            '        '是否顯示完整客戶姓名
            '        If Not IsDBNull(table1.Rows(i)("客戶姓名")) Then
            '            If CheckBox4.Checked = True Then
            '                If i = 0 Then
            '                    所有權人姓名 = Mid(Replace(table1.Rows(i)("客戶姓名"), "\", ""), 1, 1) & "先生/小姐"
            '                Else
            '                    所有權人姓名 &= "、" & Mid(Replace(table1.Rows(i)("客戶姓名"), "\", ""), 1, 1) & "先生/小姐"
            '                End If

            '            Else
            '                If i = 0 Then
            '                    所有權人姓名 = Replace(table1.Rows(i)("客戶姓名"), "\", "") & "先生/小姐"
            '                Else
            '                    所有權人姓名 &= "、" & Replace(table1.Rows(i)("客戶姓名"), "\", "") & "先生/小姐"
            '                End If

            '            End If
            '        End If
            '    Next
            'Else
            '    所有權人姓名 = ""
            'End If

            '型態
            Dim 物件型態 As String = ""

            If Not IsDBNull(t2.Rows(0)("物件類別")) Then
                Select Case Trim(t2.Rows(0)("物件類別").ToString)
                    Case "公寓", "透天", "店面", "商辦", "華廈", "大樓", "套房"
                        物件型態 = "區分所有建物，"
                    Case "工廠", "廠辦", "農舍", "倉庫"
                        物件型態 = "其他特殊建物，"
                    Case Else
                End Select
                Select Case Trim(t2.Rows(0)("物件類別").ToString)
                    Case "大樓"
                        物件型態 &= "住宅大樓(11層含以上有電梯)"
                    Case "公寓"
                        物件型態 &= "公寓(5樓含以下)"
                    Case "套房"
                        物件型態 &= "套房(1房1廳1衛)"
                    Case "華廈"
                        物件型態 &= "華廈(10層含以下有電梯)"
                    Case "商辦"
                        物件型態 &= "辦公商業大樓"
                    Case "店面"
                        物件型態 &= "店面(店鋪)"
                    Case "透天"
                        物件型態 &= "透天厝"
                    Case Else
                        物件型態 &= t2.Rows(0)("物件類別").ToString
                End Select
            End If

            '格局
            If t2.Rows(0)("房").ToString = "-1" Then
                格局 = "開放格局"
            Else
                If Convert.ToDouble("0" & t2.Rows(0)("房").ToString) > 0 Then
                    格局 &= t2.Rows(0)("房").ToString & "房"
                End If

                If Convert.ToDouble("0" & t2.Rows(0)("廳").ToString) > 0 Then
                    格局 &= t2.Rows(0)("廳").ToString & "廳"
                End If

                If Convert.ToDouble("0" & t2.Rows(0)("衛").ToString) > 0 Then
                    格局 &= t2.Rows(0)("衛").ToString & "衛"
                End If

                If Convert.ToDouble("0" & t2.Rows(0)("室").ToString) > 0 Then
                    格局 &= t2.Rows(0)("室").ToString & "室"
                End If

            End If

            If 物件型態 <> "" Then
                格局 = 物件型態 & "，" & 格局
            End If

            If Not IsDBNull(t2.Rows(0)("土地標示")) Then
                建物座落 = t2.Rows(0)("土地標示").ToString
            End If

            '[END============================================產權調查-建物標示說明P1=============================================END]

        Else

            '[START============================================產權調查-土地標示說明P1=============================================START]
            '委託價格
            If Right(t2.Rows(0)("刊登售價").ToString, 5) = ".0000" Then
                委託價格 = Int(t2.Rows(0)("刊登售價").ToString) & " 萬元整"
            Else
                委託價格 = t2.Rows(0)("刊登售價").ToString & " 萬元整"
            End If

            '土地權利種類
            Dim sqlstr As String = ""
            sqlstr = "SELECT top 1 a.所有權人,a.權利範圍_分子,a.權利範圍_分母 ,a.權利總類, a.所有權種類,a.權利人流水號 FROM 委賣物件資料表_細項所有權人 a JOIN 委賣物件資料表_面積細項 b ON a.物件編號 = b.物件編號 AND a.店代號 = b.店代號 AND a.細項流水號 = b.流水號 and a.店代號=b.店代號 WHERE b.類別 = '土地面積' and b.DL_level2_selectindex = '0' and  (a.物件編號 = '" & Contract_No & "') and a.店代號 = '" & Request("sid") & "' order by b.流水號 "
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                Using cmd As New SqlCommand(sqlstr, conn)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader())

                    '所有權人姓名
                    Dim columns As String() = {"所有權人", "權利範圍_分子", "權利範圍_分母", "權利總類", "所有權種類", "權利人流水號"}
                    Dim dt_name As DataTable = dt.DefaultView.ToTable(True, columns)
                    For Each dr As DataRow In dt_name.Rows
                        If Not IsDBNull(dr("所有權種類")) Then
                            土地權利種類 = dr("所有權種類")
                        End If
                    Next
                End Using
            End Using
            '土地權利種類 = t2.Rows(0)("所有權型態為").ToString

            '土地開發方式限制或其他負擔
            If Not IsDBNull(t2.Rows(0)("開發限制方式")) Then
                If t2.Rows(0)("開發限制方式").ToString = "其它" Then
                    土地開發方式限制或其他負擔 = "其它：" & t2.Rows(0)("開發限制方式").ToString
                Else
                    土地開發方式限制或其他負擔 = t2.Rows(0)("開發限制方式").ToString
                End If
            Else
                土地開發方式限制或其他負擔 = ""
            End If
            '[END============================================產權調查-土地標示說明P1=============================================END]
        End If '---------------------------------------------------------------------------------------------------------------------------判斷成屋還是土地用

        '產權調查-基地調查參數
        Dim 基地座落 As String = ""
        Dim 土地使用分區 As String = ""
        Dim 法定建蔽率 As String = ""
        Dim 法定容積率 As String = ""
        Dim 基地權利種類 As String = ""
        Dim 基地所有權人 As String = ""
        Dim 基地開發方式限制或其他負擔 As String = ""
        Dim 建物基地座落 As String = ""
        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then '-------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-基地調查=============================================START]
            '所有權人姓名 改抓 委賣物件資料表_細項所有權人 20160414 by nick
            Dim sqlstr As String = ""
            sqlstr = "SELECT a.所有權人,a.所有權種類,權利總類,法定建蔽率 + '%' as 法定建蔽率,法定容積率 + '%' as 法定容積率,建號,b.使用分區 FROM 委賣物件資料表_細項所有權人 a RIGHT JOIN 委賣物件資料表_面積細項 b ON a.物件編號 = b.物件編號 AND a.細項流水號 = b.流水號 and a.店代號=b.店代號 WHERE b.類別 = '土地面積' and (b.物件編號 = '" & Contract_No & "') and b.店代號 = '" & Request("sid") & "' "
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                Using cmd As New SqlCommand(sqlstr, conn)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader())

                    '權利種類
                    Dim dt_rights As DataTable = dt.DefaultView.ToTable(True, "權利總類")
                    Dim dt_people As DataTable = dt.DefaultView.ToTable(True, "所有權人")
                    Dim dt_owner As DataTable = dt.DefaultView.ToTable(True, "所有權種類")
                    For Each dr As DataRow In dt_owner.Rows
                        If Not IsDBNull(dr("所有權種類")) Then
                            基地權利種類 = dr("所有權種類")
                        End If
                    Next
                    '所有權人
                    For Each dr As DataRow In dt_people.Rows
                        '是否顯示完整客戶姓名
                        If Not IsDBNull(dr("所有權人")) Then
                            If CheckBox4.Checked = True Then
                                If i = 0 Then
                                    基地所有權人 = Mid(Replace(dr("所有權人"), "\", ""), 1, 1) '& "先生/小姐"
                                    '加入所有權持份
                                Else
                                    基地所有權人 &= "、" & Mid(Replace(dr("所有權人"), "\", ""), 1, 1) '& "先生/小姐"
                                End If
                            Else
                                If i = 0 Then
                                    基地所有權人 = Replace(dr("所有權人"), "\", "") '& "先生/小姐"
                                Else
                                    基地所有權人 &= "、" & Replace(dr("所有權人"), "\", "") '& "先生/小姐"
                                End If
                            End If
                        End If
                        i = i + 1
                    Next
                    For Each dr As DataRow In dt.Rows
                        If Not IsDBNull(dr("使用分區")) Then
                            If dr("使用分區").ToString.Trim <> "" Then
                                土地使用分區 &= dr("使用分區") & "(地號" & dr("建號") & ") "
                            End If
                        End If
                        '土地使用分區 = dr("使用分區").ToString
                        If Not IsDBNull(dr("法定建蔽率")) Then
                            If dr("法定建蔽率") <> "%" Then
                                法定建蔽率 &= dr("法定建蔽率") & "(地號" & dr("建號") & ") "
                            End If
                        End If
                        If Not IsDBNull(dr("法定容積率")) Then
                            If dr("法定容積率") <> "%" Then
                                法定容積率 &= dr("法定容積率") & "(地號" & dr("建號") & ") "
                            End If
                        End If
                    Next
                End Using
            End Using

            If 法定建蔽率 = "" Then
                'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                法定建蔽率 = "本案已開發建築，若買方欲增建、改建時，仍須依都市計畫法、區域計畫法及其他相關法令之非都市或都市土地使用分區管制規定之最新建蔽率規定為準"
                'Else
                '    法定建蔽率 = "尚未調查"
                'End If
            End If
            If 法定容積率 = "" Then
                'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                法定容積率 = "本案已開發建築，若買方欲增建、改建時，仍須依都市計畫法、區域計畫法及其他相關法令之非都市或都市土地使用分區管制規定之最新容積率規定為準"
                'Else
                '    法定容積率 = "尚未調查"
                'End If
                '法定容積率 = "尚未調查"
            End If

            基地座落 = ""

            '土地使用分區 = t2.Rows(0)("true物件用途").ToString

            If Not IsDBNull(t2.Rows(0)("其他使用分區")) Then
                If Trim(t2.Rows(0)("其他使用分區").ToString) <> "" Then
                    土地使用分區 &= "," & t2.Rows(0)("其他使用分區").ToString
                End If
            End If

            'If Not IsDBNull(t2.Rows(0)("法定建蔽率")) Then
            '    If t2.Rows(0)("法定建蔽率").ToString.Trim <> "" Then
            '        If t2.Rows(0)("法定建蔽率").ToString.Trim = "%" Then
            '            法定建蔽率 = "尚未調查"
            '        Else
            '            法定建蔽率 = t2.Rows(0)("法定建蔽率").ToString
            '        End If
            '    Else
            '        法定建蔽率 = "尚未調查"
            '    End If
            'Else
            '    法定建蔽率 = "尚未調查"
            'End If


            'If Not IsDBNull(t2.Rows(0)("法定容積率")) Then
            '    If t2.Rows(0)("法定容積率").ToString.Trim <> "" Then
            '        If t2.Rows(0)("法定容積率").ToString.Trim = "%" Then
            '            法定容積率 = "尚未調查"
            '        Else
            '            法定容積率 = t2.Rows(0)("法定容積率").ToString
            '        End If
            '    Else
            '        法定容積率 = "尚未調查"
            '    End If

            'Else
            '    法定容積率 = "尚未調查"
            'End If

            '基地權利種類 = t2.Rows(0)("所有權型態為").ToString


            '基地開發方式限制或其他負擔 
            If Not IsDBNull(t2.Rows(0)("開發限制方式")) Then
                If t2.Rows(0)("開發限制方式").ToString = "其它" Then
                    基地開發方式限制或其他負擔 = "其它：" & t2.Rows(0)("開發限制方式").ToString
                Else
                    基地開發方式限制或其他負擔 = t2.Rows(0)("開發限制方式").ToString
                End If
            Else
                基地開發方式限制或其他負擔 = ""
            End If

            If Not IsDBNull(t2.Rows(0)("土地標示")) Then
                建物基地座落 = t2.Rows(0)("土地標示").ToString
            End If
            '[END============================================產權調查-基地調查=============================================END]
        End If '---------------------------------------------------------------------------------------------------------------------------判斷成屋還是土地用

        '共用
        '[START============================================重要交易條件=============================================START]
        Dim 交易價金 As String = ""
        Dim 第一期款 As String = ""
        Dim 第二期款 As String = ""
        Dim 第三期款 As String = ""
        Dim 第四期款 As String = ""
        Dim 一期 As String = ""
        Dim 二期 As String = ""
        Dim 三期 As String = ""
        Dim 四期 As String = ""
        Dim 土增稅賣 As String = ""
        Dim 土增稅買 As String = ""
        Dim 土增稅雙 As String = ""
        Dim 土增稅交 As String = ""
        Dim 土增稅約 As String = ""
        Dim 土增稅金 As String = ""
        Dim 地價稅賣 As String = ""
        Dim 地價稅買 As String = ""
        Dim 地價稅雙 As String = ""
        Dim 地價稅交 As String = ""
        Dim 地價稅約 As String = ""
        Dim 地價稅金 As String = ""
        Dim 房屋稅賣 As String = ""
        Dim 房屋稅買 As String = ""
        Dim 房屋稅雙 As String = ""
        Dim 房屋稅交 As String = ""
        Dim 房屋稅約 As String = ""
        Dim 房屋稅金 As String = ""
        Dim 契稅賣 As String = ""
        Dim 契稅買 As String = ""
        Dim 契稅雙 As String = ""
        Dim 契稅交 As String = ""
        Dim 契稅約 As String = ""
        Dim 契稅金 As String = ""
        Dim 印花稅賣 As String = ""
        Dim 印花稅買 As String = ""
        Dim 印花稅雙 As String = ""
        Dim 印花稅交 As String = ""
        Dim 印花稅約 As String = ""
        Dim 印花稅金 As String = ""
        Dim 代書費賣 As String = ""
        Dim 代書費買 As String = ""
        Dim 代書費雙 As String = ""
        Dim 代書費交 As String = ""
        Dim 代書費約 As String = ""
        Dim 代書費金 As String = ""
        Dim 登記費賣 As String = ""
        Dim 登記費買 As String = ""
        Dim 登記費雙 As String = ""
        Dim 登記費交 As String = ""
        Dim 登記費約 As String = ""
        Dim 登記費金 As String = ""
        Dim 公證費賣 As String = ""
        Dim 公證費買 As String = ""
        Dim 公證費雙 As String = ""
        Dim 公證費交 As String = ""
        Dim 公證費約 As String = ""
        Dim 公證費金 As String = ""
        Dim 水電費賣 As String = ""
        Dim 水電費買 As String = ""
        Dim 水電費雙 As String = ""
        Dim 水電費交 As String = ""
        Dim 水電費約 As String = ""
        Dim 水電費金 As String = ""
        Dim 瓦斯費賣 As String = ""
        Dim 瓦斯費買 As String = ""
        Dim 瓦斯費雙 As String = ""
        Dim 瓦斯費交 As String = ""
        Dim 瓦斯費約 As String = ""
        Dim 瓦斯費金 As String = ""
        Dim 管理費賣 As String = ""
        Dim 管理費買 As String = ""
        Dim 管理費雙 As String = ""
        Dim 管理費交 As String = ""
        Dim 管理費約 As String = ""
        Dim 管理費金 As String = ""
        Dim 電話費賣 As String = ""
        Dim 電話費買 As String = ""
        Dim 電話費雙 As String = ""
        Dim 電話費交 As String = ""
        Dim 電話費約 As String = ""
        Dim 電話費金 As String = ""
        Dim 工程費賣 As String = ""
        Dim 工程費買 As String = ""
        Dim 工程費雙 As String = ""
        Dim 工程費交 As String = ""
        Dim 工程費約 As String = ""
        Dim 工程費金 As String = ""
        Dim 房地合一賣 As String = ""
        Dim 房地合一買 As String = ""
        Dim 房地合一雙 As String = ""
        Dim 房地合一交 As String = ""
        Dim 房地合一約 As String = ""
        Dim 房地合一金 As String = ""
        'Dim 奢侈稅賣 As String = ""
        'Dim 奢侈稅買 As String = ""
        'Dim 奢侈稅雙 As String = ""
        'Dim 奢侈稅交 As String = ""
        'Dim 奢侈稅約 As String = ""
        'Dim 奢侈稅金 As String = ""

        Dim 實價登錄賣 As String = ""
        Dim 實價登錄買 As String = ""
        Dim 實價登錄雙 As String = ""
        Dim 實價登錄交 As String = ""
        Dim 實價登錄約 As String = ""
        Dim 實價登錄金 As String = ""
        Dim 代書費New賣 As String = ""
        Dim 代書費New買 As String = ""
        Dim 代書費New雙 As String = ""
        Dim 代書費New交 As String = ""
        Dim 代書費New約 As String = ""
        Dim 代書費New金 As String = ""

        If Right(t2.Rows(0)("刊登售價").ToString, 5) = ".0000" Then
            交易價金 = Int(t2.Rows(0)("刊登售價").ToString) & " 萬元整"
        Else
            交易價金 = t2.Rows(0)("刊登售價").ToString & " 萬元整"
        End If

        '第一期金額
        If Not IsDBNull(t2.Rows(0)("第一期金額")) Then
            If Trim(t2.Rows(0)("第一期金額").ToString) <> "" Then
                If Right(t2.Rows(0)("第一期金額").ToString, 5) = ".0000" Then
                    第一期款 = Int(t2.Rows(0)("第一期金額").ToString)
                ElseIf t2.Rows(0)("第一期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第一期金額").ToString, 3) = "000" Then
                    第一期款 = Left(t2.Rows(0)("第一期金額").ToString, Len(t2.Rows(0)("第一期金額").ToString) - 3)
                ElseIf t2.Rows(0)("第一期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第一期金額").ToString, 2) = "00" Then
                    第一期款 = Left(t2.Rows(0)("第一期金額").ToString, Len(t2.Rows(0)("第一期金額").ToString) - 2)
                ElseIf t2.Rows(0)("第一期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第一期金額").ToString, 1) = "0" Then
                    第一期款 = Left(t2.Rows(0)("第一期金額").ToString, Len(t2.Rows(0)("第一期金額").ToString) - 1)
                Else
                    第一期款 = t2.Rows(0)("第一期金額").ToString
                End If
            End If
        End If
        '一期(簽約金)
        If Not IsDBNull(t2.Rows(0)("簽約金")) Then
            一期 = t2.Rows(0)("簽約金").ToString & "%"
        End If

        '第二期金額
        If Not IsDBNull(t2.Rows(0)("第二期金額")) Then
            If Trim(t2.Rows(0)("第二期金額").ToString) <> "" Then
                If Right(t2.Rows(0)("第二期金額").ToString, 5) = ".0000" Then
                    第二期款 = Int(t2.Rows(0)("第二期金額").ToString)
                ElseIf t2.Rows(0)("第二期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第二期金額").ToString, 3) = "000" Then
                    第二期款 = Left(t2.Rows(0)("第二期金額").ToString, Len(t2.Rows(0)("第二期金額").ToString) - 3)
                ElseIf t2.Rows(0)("第二期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第二期金額").ToString, 2) = "00" Then
                    第二期款 = Left(t2.Rows(0)("第二期金額").ToString, Len(t2.Rows(0)("第二期金額").ToString) - 2)
                ElseIf t2.Rows(0)("第二期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第二期金額").ToString, 1) = "0" Then
                    第二期款 = Left(t2.Rows(0)("第二期金額").ToString, Len(t2.Rows(0)("第二期金額").ToString) - 1)
                Else
                    第二期款 = t2.Rows(0)("第二期金額").ToString
                End If
            End If
        End If
        '二期(備證款)
        If Not IsDBNull(t2.Rows(0)("備証款")) Then
            二期 = t2.Rows(0)("備証款").ToString & "%"
        End If

        '第三期金額
        If Not IsDBNull(t2.Rows(0)("第三期金額")) Then
            If Trim(t2.Rows(0)("第三期金額").ToString) <> "" Then
                If Right(t2.Rows(0)("第三期金額").ToString, 5) = ".0000" Then
                    第三期款 = Int(t2.Rows(0)("第三期金額").ToString)
                ElseIf t2.Rows(0)("第三期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第三期金額").ToString, 3) = "000" Then
                    第三期款 = Left(t2.Rows(0)("第三期金額").ToString, Len(t2.Rows(0)("第三期金額").ToString) - 3)
                ElseIf t2.Rows(0)("第三期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第三期金額").ToString, 2) = "00" Then
                    第三期款 = Left(t2.Rows(0)("第三期金額").ToString, Len(t2.Rows(0)("第三期金額").ToString) - 2)
                ElseIf t2.Rows(0)("第三期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第三期金額").ToString, 1) = "0" Then
                    第三期款 = Left(t2.Rows(0)("第三期金額").ToString, Len(t2.Rows(0)("第三期金額").ToString) - 1)
                Else
                    第三期款 = t2.Rows(0)("第三期金額").ToString
                End If
            End If
        End If
        '三期(完稅款)
        If Not IsDBNull(t2.Rows(0)("完稅款")) Then
            三期 = t2.Rows(0)("完稅款").ToString & "%"
        End If

        '第四期金額
        If Not IsDBNull(t2.Rows(0)("第四期金額")) Then
            If Trim(t2.Rows(0)("第四期金額").ToString) <> "" Then
                If Right(t2.Rows(0)("第四期金額").ToString, 5) = ".0000" Then
                    第四期款 = Int(t2.Rows(0)("第四期金額").ToString)
                ElseIf t2.Rows(0)("第四期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第四期金額").ToString, 3) = "000" Then
                    第四期款 = Left(t2.Rows(0)("第四期金額").ToString, Len(t2.Rows(0)("第四期金額").ToString) - 3)
                ElseIf t2.Rows(0)("第四期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第四期金額").ToString, 2) = "00" Then
                    第四期款 = Left(t2.Rows(0)("第四期金額").ToString, Len(t2.Rows(0)("第四期金額").ToString) - 2)
                ElseIf t2.Rows(0)("第四期金額").ToString.LastIndexOf(".") > 0 And Right(t2.Rows(0)("第四期金額").ToString, 1) = "0" Then
                    第四期款 = Left(t2.Rows(0)("第四期金額").ToString, Len(t2.Rows(0)("第四期金額").ToString) - 1)
                Else
                    第四期款 = t2.Rows(0)("第四期金額").ToString
                End If
            End If
        End If
        '四期(尾款)
        If Not IsDBNull(t2.Rows(0)("尾款")) Then
            四期 = t2.Rows(0)("尾款").ToString & "%"
        End If

        '增值稅  
        If Not IsDBNull(t2.Rows(0)("增值稅")) Then
            Select Case t2.Rows(0)("增值稅").ToString
                Case "買方負擔"
                    土增稅買 = "◎"
                Case "賣方負擔"
                    土增稅賣 = "◎"
                Case "雙方各半"
                    土增稅雙 = "◎"
                Case "另以契約約定"
                    土增稅約 = "◎"
                Case "依交屋日分算"
                    土增稅交 = "◎"
            End Select
        End If

        '增值稅金
        土增稅金 = "依核定稅單為準"

        '地價稅  
        If Not IsDBNull(t2.Rows(0)("地價稅")) Then
            Select Case t2.Rows(0)("地價稅").ToString
                Case "買方負擔"
                    地價稅買 = "◎"
                Case "賣方負擔"
                    地價稅賣 = "◎"
                Case "雙方各半"
                    地價稅雙 = "◎"
                Case "另以契約約定"
                    地價稅約 = "◎"
                Case "依交屋日分算"
                    地價稅交 = "◎"
            End Select
        End If

        '地價稅金  
        If Not IsDBNull(t2.Rows(0)("地價稅約")) Then
            If Trim(t2.Rows(0)("地價稅約").ToString) <> "" Then
                If Right(t2.Rows(0)("地價稅約").ToString, 5) = ".0000" Then
                    地價稅金 = Int(t2.Rows(0)("地價稅約").ToString) & "元整"
                Else
                    地價稅金 = t2.Rows(0)("地價稅約").ToString & "元整"
                End If
            Else
                地價稅金 = "依核定稅單為準"
            End If
        Else
            地價稅金 = "依核定稅單為準"
        End If

        '房屋稅  
        If Not IsDBNull(t2.Rows(0)("房屋稅")) Then
            Select Case t2.Rows(0)("房屋稅").ToString
                Case "買方負擔"
                    房屋稅買 = "◎"
                Case "賣方負擔"
                    房屋稅賣 = "◎"
                Case "雙方各半"
                    房屋稅雙 = "◎"
                Case "另以契約約定"
                    房屋稅約 = "◎"
                Case "依交屋日分算"
                    房屋稅交 = "◎"
            End Select
        End If

        '房屋稅金  
        If Not IsDBNull(t2.Rows(0)("房屋稅約")) Then
            If Trim(t2.Rows(0)("房屋稅約").ToString) <> "" Then
                If Right(t2.Rows(0)("房屋稅約").ToString, 5) = ".0000" Then
                    房屋稅金 = Int(t2.Rows(0)("房屋稅約").ToString) & "元整"
                Else
                    房屋稅金 = t2.Rows(0)("房屋稅約").ToString & "元整"
                End If
            Else
                房屋稅金 = "依核定稅單為準"
            End If
        Else
            房屋稅金 = "依核定稅單為準"
        End If

        '契稅  
        If Not IsDBNull(t2.Rows(0)("房地產契稅")) Then
            Select Case t2.Rows(0)("房地產契稅").ToString
                Case "買方負擔"
                    契稅買 = "◎"
                Case "賣方負擔"
                    契稅賣 = "◎"
                Case "雙方各半"
                    契稅雙 = "◎"
                Case "另以契約約定"
                    契稅約 = "◎"
                Case "依交屋日分算"
                    契稅交 = "◎"
            End Select
        End If

        '契稅金  
        If Not IsDBNull(t2.Rows(0)("契稅約")) Then
            '    If Trim(t2.Rows(0)("契稅約").ToString) <> "" Then
            '        If Right(t2.Rows(0)("契稅約").ToString, 5) = ".0000" Then
            '            契稅金 = Int(t2.Rows(0)("契稅約").ToString) & "元整"
            '        Else
            '            契稅金 = t2.Rows(0)("契稅約").ToString & "元整"
            '        End If
            '    Else
            '        契稅金 = "依核定稅單為準"
            '    End If
            'Else
            契稅金 = "依核定稅單為準"
        End If

        '印花稅  
        If Not IsDBNull(t2.Rows(0)("印花稅")) Then
            Select Case t2.Rows(0)("印花稅").ToString
                Case "買方負擔"
                    印花稅買 = "◎"
                Case "賣方負擔"
                    印花稅賣 = "◎"
                Case "雙方各半"
                    印花稅雙 = "◎"
                Case "另以契約約定"
                    印花稅約 = "◎"
                Case "依交屋日分算"
                    印花稅交 = "◎"
            End Select
        End If

        '印花稅金  
        If Not IsDBNull(t2.Rows(0)("印花稅約")) Then
            If Trim(t2.Rows(0)("印花稅約").ToString) <> "" Then
                If Right(t2.Rows(0)("印花稅約").ToString, 5) = ".0000" Then
                    印花稅金 = Int(t2.Rows(0)("印花稅約").ToString) & "元整"
                Else
                    印花稅金 = t2.Rows(0)("印花稅約").ToString & "元整"
                End If
            Else
                印花稅金 = "依核定稅單為準"
            End If
        Else
            印花稅金 = "依核定稅單為準"
        End If

        '代書費  
        If Not IsDBNull(t2.Rows(0)("代書費")) Then
            Select Case t2.Rows(0)("代書費").ToString
                Case "買方負擔"
                    代書費買 = "◎"
                Case "賣方負擔"
                    代書費賣 = "◎"
                Case "雙方各半"
                    代書費雙 = "◎"
                Case "另以契約約定"
                    代書費約 = "◎"
                Case "依交屋日分算"
                    代書費交 = "◎"
            End Select
        End If

        '代書費金  
        'If Not IsDBNull(t2.Rows(0)("代書費約")) Then
        '    If Trim(t2.Rows(0)("代書費約").ToString) <> "" Then
        '        If Right(t2.Rows(0)("代書費約").ToString, 5) = ".0000" Then
        '            代書費金 = Int(t2.Rows(0)("代書費約").ToString) & "元整"
        '        Else
        '            代書費金 = t2.Rows(0)("代書費約").ToString & "元整"
        '        End If
        '    End If
        'End If

        '登記規費  
        If Not IsDBNull(t2.Rows(0)("登記規費")) Then
            Select Case t2.Rows(0)("登記規費").ToString
                Case "買方負擔"
                    登記費買 = "◎"
                Case "賣方負擔"
                    登記費賣 = "◎"
                Case "雙方各半"
                    登記費雙 = "◎"
                Case "另以契約約定"
                    登記費約 = "◎"
                Case "依交屋日分算"
                    登記費交 = "◎"
            End Select
        End If

        '登記費金  
        If Not IsDBNull(t2.Rows(0)("登記規費約")) Then
            If Trim(t2.Rows(0)("登記規費約").ToString) <> "" Then
                If Right(t2.Rows(0)("登記規費約").ToString, 5) = ".0000" Then
                    登記費金 = Int(t2.Rows(0)("登記規費約").ToString) & "元整"
                Else
                    登記費金 = t2.Rows(0)("登記規費約").ToString & "元整"
                End If
            End If
        End If

        '公(監)證費  
        If Not IsDBNull(t2.Rows(0)("公證費")) Then
            Select Case t2.Rows(0)("公證費").ToString
                Case "買方負擔"
                    公證費買 = "◎"
                Case "賣方負擔"
                    公證費賣 = "◎"
                Case "雙方各半"
                    公證費雙 = "◎"
                Case "另以契約約定"
                    公證費約 = "◎"
                Case "依交屋日分算"
                    公證費交 = "◎"
            End Select
        End If

        '公證費金  
        If Not IsDBNull(t2.Rows(0)("公證費約")) Then
            If Trim(t2.Rows(0)("公證費約").ToString) <> "" Then
                If Right(t2.Rows(0)("公證費約").ToString, 5) = ".0000" Then
                    公證費金 = Int(t2.Rows(0)("公證費約").ToString) & "元整"
                Else
                    公證費金 = t2.Rows(0)("公證費約").ToString & "元整"
                End If
            End If
        End If

        '水電費  
        If Not IsDBNull(t2.Rows(0)("水電費")) Then
            Select Case t2.Rows(0)("水電費").ToString
                Case "買方負擔"
                    水電費買 = "◎"
                Case "賣方負擔"
                    水電費賣 = "◎"
                Case "雙方各半"
                    水電費雙 = "◎"
                Case "另以契約約定"
                    水電費約 = "◎"
                Case "依交屋日分算"
                    水電費交 = "◎"
            End Select
        End If

        '水電費金  
        'If Not IsDBNull(t2.Rows(0)("水電費約")) Then
        '    If Trim(t2.Rows(0)("水電費約").ToString) <> "" Then
        '        If Right(t2.Rows(0)("水電費約").ToString, 5) = ".0000" Then
        '            水電費金 = Int(t2.Rows(0)("水電費約").ToString) & "元整"
        '        Else
        '            水電費金 = t2.Rows(0)("水電費約").ToString & "元整"
        '        End If
        '    End If
        'End If

        '瓦斯費  
        If Not IsDBNull(t2.Rows(0)("瓦斯費")) Then
            Select Case t2.Rows(0)("瓦斯費").ToString
                Case "買方負擔"
                    瓦斯費買 = "◎"
                Case "賣方負擔"
                    瓦斯費賣 = "◎"
                Case "雙方各半"
                    瓦斯費雙 = "◎"
                Case "另以契約約定"
                    瓦斯費約 = "◎"
                Case "依交屋日分算"
                    瓦斯費交 = "◎"
            End Select
        End If

        '瓦斯費金  
        If Not IsDBNull(t2.Rows(0)("瓦斯費約")) Then
            If Trim(t2.Rows(0)("瓦斯費約").ToString) <> "" Then
                If Right(t2.Rows(0)("瓦斯費約").ToString, 5) = ".0000" Then
                    瓦斯費金 = Int(t2.Rows(0)("瓦斯費約").ToString) & "元整"
                Else
                    瓦斯費金 = t2.Rows(0)("瓦斯費約").ToString & "元整"
                End If
            End If
        End If

        '管理費  
        If Not IsDBNull(t2.Rows(0)("房地產管理費")) Then
            Select Case t2.Rows(0)("房地產管理費").ToString
                Case "買方負擔"
                    管理費買 = "◎"
                Case "賣方負擔"
                    管理費賣 = "◎"
                Case "雙方各半"
                    管理費雙 = "◎"
                Case "另以契約約定"
                    管理費約 = "◎"
                Case "依交屋日分算"
                    管理費交 = "◎"
            End Select
        End If

        '管理費金  
        If Not IsDBNull(t2.Rows(0)("管理費約")) Then
            If Trim(t2.Rows(0)("管理費約").ToString) <> "" Then
                If Right(t2.Rows(0)("管理費約").ToString, 5) = ".0000" Then
                    管理費金 = Int(t2.Rows(0)("管理費約").ToString) & "元整"
                Else
                    管理費金 = t2.Rows(0)("管理費約").ToString & "元整"
                End If
            End If
        End If

        '電話費 
        If Not IsDBNull(t2.Rows(0)("電話費")) Then
            Select Case t2.Rows(0)("電話費").ToString
                Case "買方負擔"
                    電話費買 = "◎"
                Case "賣方負擔"
                    電話費賣 = "◎"
                Case "雙方各半"
                    電話費雙 = "◎"
                Case "另以契約約定"
                    電話費約 = "◎"
                Case "依交屋日分算"
                    電話費交 = "◎"
            End Select
        End If

        '電話費金  
        If Not IsDBNull(t2.Rows(0)("電話費約")) Then
            If Trim(t2.Rows(0)("電話費約").ToString) <> "" Then
                If Right(t2.Rows(0)("電話費約").ToString, 5) = ".0000" Then
                    電話費金 = Int(t2.Rows(0)("電話費約").ToString) & "元整"
                Else
                    電話費金 = t2.Rows(0)("電話費約").ToString & "元整"
                End If
            End If
        End If

        '工程受益費  
        If Not IsDBNull(t2.Rows(0)("工程受益費")) Then
            Select Case t2.Rows(0)("工程受益費").ToString
                Case "買方負擔"
                    工程費買 = "◎"
                Case "賣方負擔"
                    工程費賣 = "◎"
                Case "雙方各半"
                    工程費雙 = "◎"
                Case "另以契約約定"
                    工程費約 = "◎"
                Case "依交屋日分算"
                    工程費交 = "◎"
            End Select
        End If

        '工程費金  
        If Not IsDBNull(t2.Rows(0)("工程受益費約")) Then
            If Trim(t2.Rows(0)("工程受益費約").ToString) <> "" Then
                If Right(t2.Rows(0)("工程受益費約").ToString, 5) = ".0000" Then
                    工程費金 = Int(t2.Rows(0)("工程受益費約").ToString) & "元整"
                Else
                    工程費金 = t2.Rows(0)("工程受益費約").ToString & "元整"
                End If
            Else
                工程費金 = "依核定稅單為準"
            End If
        Else
            工程費金 = "依核定稅單為準"
        End If

        '房地合一稅
        房地合一賣 = "◎"
        房地合一金 = "依政府規定"

        '實價登錄  
        If Not IsDBNull(t2.Rows(0)("實價登錄費")) Then
            Select Case t2.Rows(0)("實價登錄費").ToString
                Case "買方負擔"
                    實價登錄買 = "◎"
                Case "賣方負擔"
                    實價登錄賣 = "◎"
                Case "雙方各半"
                    實價登錄雙 = "◎"
                Case "另以契約約定"
                    實價登錄約 = "◎"
                Case "依交屋日分算"
                    實價登錄交 = "◎"
            End Select
        End If
        '實價登錄
        實價登錄金 = "依政府規定"

        '實價登錄  
        If Not IsDBNull(t2.Rows(0)("代書費New")) Then
            Select Case t2.Rows(0)("代書費New").ToString
                Case "買方負擔"
                    代書費New買 = "◎"
                Case "賣方負擔"
                    代書費New賣 = "◎"
                Case "雙方各半"
                    代書費New雙 = "◎"
                Case "另以契約約定"
                    代書費New約 = "◎"
                Case "依交屋日分算"
                    代書費New交 = "◎"
            End Select
        End If
        '實價登錄
        代書費New金 = "依政府規定"

        '20160222 移除此段 移到 產調_建物 內
        ''附贈設備參數
        Dim 買方設備 As String = "移至個案目前管理與使用狀況內"
        Dim 買方設備1 As String = ""
        Dim 買方設備2 As String = ""
        Dim 買方設備3 As String = ""

        '共用
        '[START============================================照片說明=============================================START]
        ''以下為Cry照片說明.rpt 中的參數
        Dim 環境介紹 As String = ""
        ' 環境介紹 
        If t2.Rows(0)("訴求重點").ToString = "" Then
            環境介紹 = ""
        Else
            環境介紹 = "訴求重點︰" & Chr(13) + Chr(10) & Server.HtmlEncode(t2.Rows(0)("訴求重點"))
            'If Contract_No = "60692AAE29288" Then
            '    環境介紹 = "訴求重點︰" & Chr(13) + Chr(10) & Server.HtmlEncode(t2.Rows(0)("訴求重點"))
            'End If
            '環境介紹 = "訴求重點︰" & Chr(13) + Chr(10) & t2.Rows(0)("訴求重點")
        End If

        '判斷有無勾選環境介紹
        If CheckBox3.Checked = True Then
            If IsDBNull(t2.Rows(0)("環境介紹")) = False Then
                If t2.Rows(0)("環境介紹") <> "" Then
                    環境介紹 &= Chr(13) + Chr(10)
                    環境介紹 &= Chr(13) + Chr(10) + "環境介紹︰" & Chr(13) + Chr(10)
                    環境介紹 &= Replace(t2.Rows(0)("環境介紹").ToString, "<br>", Chr(13) + Chr(10))
                End If
            End If
        End If

        '20150921 加入提醒
        環境介紹 &= "環境介紹:請參閱鄰近重要設施參考圖暨位置略圖"
        '[END============================================照片說明=============================================END]

        Dim Xml2Doc As New XmlToDoc.XmlToDoc

        Dim sFile As String = ""
        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then '成屋
            'If Request.Cookies("webfly_empno").Value.ToUpper = "H2L" Then
            '    sFile = Xml2Doc.MyGetFullTextFile(Server.MapPath("..\reports\New_Description_V5.xml"), enuStandardCodePages.SCP_CP_UTF8)
            'Else
            '    sFile = Xml2Doc.MyGetFullTextFile(Server.MapPath("..\reports\New_Description_V4.xml"), enuStandardCodePages.SCP_CP_UTF8)
            'End If
            sFile = Xml2Doc.MyGetFullTextFile(Server.MapPath("..\reports\New_Description_V4.xml"), enuStandardCodePages.SCP_CP_UTF8)
        Else '素地
            sFile = Xml2Doc.MyGetFullTextFile(Server.MapPath("..\reports\New_Description_Land_V3.xml"), enuStandardCodePages.SCP_CP_UTF8)
        End If

        'If Request("oid") = "60692AAE34137" Then

        'Else
        '頁首
        sFile = sFile.Replace("≠頁首編號及案名", 頁首編號及案名)
        '封面
        sFile = sFile.Replace("≠委託編號", 委託編號)
        sFile = sFile.Replace("≠封面案名", 封面案名)
        sFile = sFile.Replace("≠說明書類別", 說明書類別)
        sFile = sFile.Replace("≠報告日期", 報告日期) '預設先不用,=""

        sFile = sFile.Replace("≠店名", 店名)
        sFile = sFile.Replace("≠法人", 法人)

        sFile = sFile.Replace("≠店完整地址", 店完整地址)
        sFile = sFile.Replace("≠連絡電話", 連絡電話)
        sFile = sFile.Replace("≠傳真", 傳真)
        sFile = sFile.Replace("≠電子郵件", 電子郵件)
        sFile = sFile.Replace("≠承辦人", 承辦人)

        '目錄-內容
        sFile = sFile.Replace("≠產權調查", 產權調查)
        sFile = sFile.Replace("≠物件個案調查", 物件個案調查)
        sFile = sFile.Replace("≠照片說明", Server.HtmlEncode(照片說明))
        sFile = sFile.Replace("≠重要交易條件", 重要交易條件)
        sFile = sFile.Replace("≠成交行情參考表", 成交行情參考表)
        sFile = sFile.Replace("≠重要設施參考圖", 重要設施參考圖)
        sFile = sFile.Replace("≠其他說明", Server.HtmlEncode(其他說明))
        '目錄-附件
        sFile = sFile.Replace("≠土地權狀影本", 土地權狀影本)
        sFile = sFile.Replace("≠建物權狀影本", 建物權狀影本)
        sFile = sFile.Replace("≠房地產標的現況說明書", 房地產標的現況說明書)

        sFile = sFile.Replace("≠土地謄本", 土地謄本)
        sFile = sFile.Replace("≠建物謄本", 建物謄本)
        sFile = sFile.Replace("≠預售屋買賣契約書", 預售屋買賣契約書)

        sFile = sFile.Replace("≠土地地籍圖", 土地地籍圖)
        sFile = sFile.Replace("≠建物勘測成果圖", 建物測量成果圖)
        sFile = sFile.Replace("≠住戶規約", 住戶規約)

        sFile = sFile.Replace("≠土地相關位置略圖", 土地相關位置略圖)
        sFile = sFile.Replace("≠建物相關位置略圖", 建物相關位置略圖)
        sFile = sFile.Replace("≠停車位位置圖", 停車位位置圖)

        sFile = sFile.Replace("≠土地分管協議", 土地分管協議)
        sFile = sFile.Replace("≠建物分管協議", 建物分管協議)
        sFile = sFile.Replace("≠樑柱顯見裂痕照片", 樑柱顯見裂痕照片)

        sFile = sFile.Replace("≠分區使用說明", 分區使用說明)
        sFile = sFile.Replace("≠房屋稅單", 房屋稅單)
        sFile = sFile.Replace("≠封面其他", 封面其他)

        sFile = sFile.Replace("≠增值稅概算", 增值稅概算)
        sFile = sFile.Replace("≠使用執照", 使用執照)
        'End If



        'If Request("oid") = "60692AAE34137" Then

        'Else
        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then '-------判斷成屋還是土地用
            '產調-建物P1
            sFile = sFile.Replace("≠建物門牌", 建物門牌)
            sFile = sFile.Replace("≠建物所在樓層", 建物所在樓層)
            sFile = sFile.Replace("≠建築用途", 建築用途)
            sFile = sFile.Replace("≠建築結構", 建築結構)
            sFile = sFile.Replace("≠建築完成日期", 建築完成日期)
            sFile = sFile.Replace("≠權利種類說明", 權利種類說明)
            sFile = sFile.Replace("≠所有權人姓名", 所有權人姓名)
            sFile = sFile.Replace("≠格局", 格局)
            sFile = sFile.Replace("≠建物座落", 建物座落)
        Else
            '產調-土地
            sFile = sFile.Replace("≠土地權利種類", 土地權利種類)
            sFile = sFile.Replace("≠土地開發方式限制或其他負擔", 土地開發方式限制或其他負擔)
        End If '-------判斷成屋還是土地用
        '共用
        sFile = sFile.Replace("≠委託價格", 委託價格)
        'End If

        '共用
        '產調增建
        '[START============================================產權調查-建物標示說明P2-增建=============================================START]
        '有無增建
        Dim YorN As String = ""
        If IsDBNull(t2.Rows(0)("增建")) = False Then
            YorN = Trim(t2.Rows(0)("增建").ToString)
        Else '若為NULL值，改判段加蓋平方公尺是否有值
            If IsDBNull(t2.Rows(0)("加蓋平方公尺")) = False Then
                If t2.Rows(0)("加蓋平方公尺") <> 0 Then
                    YorN = "Y"
                Else
                    YorN = "N"
                End If
            Else
                YorN = "N"
            End If
        End If

        If YorN = "Y" Then
            '增建產權說明
            Dim myXML_Add As String = ""
            Dim myText_Add As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_Add As New StreamReader(Server.MapPath("..\reports\重複增建產權說明_V2.txt"))
            myText_Add = srText_Add.ReadToEnd()
            srText_Add.Close()

            Dim 最後要替代掉的字串_增建產權說明 As New StringBuilder()

            '先把讀出的xml複製一份，接著開始改
            Dim tempdata As String = myText_Add

            '改完加到最後的字串裡面  
            最後要替代掉的字串_增建產權說明.Append(tempdata)

            sFile = sFile.Replace("!重複增建產權說明", 最後要替代掉的字串_增建產權說明.ToString())
        End If
        '[END============================================產權調查-建物標示說明P2-增建=============================================END]

        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------判斷成屋還是土地用
            '20151005 依照車位數產生 一筆車位一張產調
            sqlstr = "Select B.建號 as 車建號,* From 產調_車位 A  With(NoLock) Left Join 委賣物件資料表_面積細項 B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號  Where A.物件編號 = '" & Contract_No & "' and A.店代號='" & Request("sid") & "'  order by A.流水號 "
            'Response.Write(sqlstr)
            'Response.End()
            Dim 最後要替代掉的字串_停車位產權說明 As New StringBuilder()

            Using conn_car As New SqlConnection(EGOUPLOADSqlConnStr)
                conn_car.Open()
                Using cmd As New SqlCommand(sqlstr, conn_car)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader)

                    For Each dr As DataRow In dt.Rows
                        '*******************
                        '產調-車位
                        '[START============================================產權調查-停車位產權說明=============================================START]
                        Dim 車位產權說明 As String = ""
                        Dim 車位位置 As String = ""
                        Dim 車位價格 As String = ""
                        Dim 車位性質 As String = ""
                        Dim 車位出租或占用情形 As String = ""
                        Dim 車位略圖 As String = ""
                        Dim 車位其他說明事項 As String = ""
                        Dim 車位管理費 As String = ""
                        '**** 車位細項部份 ****
                        Dim 車獨立產權 As String = ""
                        Dim 車權利種類 As String = ""
                        Dim 車建號 As String = ""
                        Dim 車型式 As String = ""
                        Dim 車編號 As String = ""
                        Dim 車長 As String = ""
                        Dim 車寬 As String = ""
                        Dim 車高 As String = ""
                        Dim 車重 As String = ""

                        'If dr("類別").ToString.IndexOf("車位面積") >= 0 Then

                        '停車位產權說明
                        Dim myXML_Car As String = ""
                        Dim myText_Car As String = ""

                        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
                        Dim srText_Car As New StreamReader(Server.MapPath("..\reports\重複停車位產權說明_V3.txt"))
                        myText_Car = srText_Car.ReadToEnd()
                        srText_Car.Close()


                        '車位產權說明
                        'If (Not String.IsNullOrEmpty(t2.Rows(0)("公設內車位坪數").ToString)) = True And (String.IsNullOrEmpty(t2.Rows(0)("車位坪數").ToString)) = True Then
                        '    車位產權說明 = "公設內,無產權獨立"

                        'ElseIf (Not String.IsNullOrEmpty(t2.Rows(0)("車位坪數").ToString)) = True And (String.IsNullOrEmpty(t2.Rows(0)("公設內車位坪數").ToString)) = True Then
                        '    車位產權說明 = "產權獨立"

                        'ElseIf (((Not String.IsNullOrEmpty(t2.Rows(0)("公設內車位坪數").ToString)) = True) And ((Not String.IsNullOrEmpty(t2.Rows(0)("車位坪數").ToString)) = True)) Then
                        '    車位產權說明 = "公設產權"
                        'End If

                        '2016.06.20 by Finch
                        'If dr("項目名稱").ToString = "車位面積(公設內)" Then
                        '    '分有產權和無產權面積
                        '    If dr("總面積平方公尺") = 0 Then
                        '        車位產權說明 = "含於公設內且無法計算面積"
                        '    Else
                        '        車位產權說明 = "公設產權"
                        '    End If
                        'Else
                        '    車位產權說明 = "產權獨立"
                        'End If
                        車位產權說明 = "如下所示"

                        '車位位置
                        '1040417新增位置地上地下	nvarchar	2
                        If IsDBNull(dr("車位位置_上下")) = False Then
                            車位位置 = dr("車位位置_上下") & " "
                        Else '若為NULL值
                            車位位置 = "地下 "
                        End If

                        If IsDBNull(dr("車位位置_樓層")) = False Then
                            車位位置 &= dr("車位位置_樓層") & "　樓"
                        Else
                            車位位置 &= "　樓"
                        End If


                        '價格或租金
                        If IsDBNull(dr("車位獨立售價")) = False Then
                            車位價格 = dr("車位獨立售價").ToString & "萬元"
                        Else
                            車位價格 = ""
                        End If
                        'If IsDBNull(t2.Rows(0)("車位價格")) = False Then
                        '    If t2.Rows(0)("車位價格").ToString = "" Then
                        '        車位價格 = ""
                        '    ElseIf t2.Rows(0)("車位價格").ToString = "含於開價中" Then
                        '        車位價格 = "合併於總價內"
                        '    Else
                        '        車位價格 = t2.Rows(0)("車位價格").ToString
                        '    End If
                        'Else
                        '    車位價格 = ""
                        'End If

                        If IsDBNull(dr("車位性質")) = False Then
                            車位性質 = dr("車位性質") '& "(詳見委託銷售契約書)"
                            If 車位性質 = "無法辨識" Then
                                車位性質 = "無法辨識其為法定停車位、自行增設停車位或獎勵增設停車位"
                            End If
                        End If

                        '車位管理費
                        '車位管理費_類別, 車位管理費_期間, 車位管理費_金額
                        If Not IsDBNull(dr("車位管理費_類別")) Then
                            Dim period As String = IIf(IsDBNull(dr("車位管理費_期間")), "", dr("車位管理費_期間"))
                            Dim fee As String = IIf(IsDBNull(dr("車位管理費_金額")), "", dr("車位管理費_金額"))
                            車位管理費 = dr("車位管理費_類別") & fee & "/" & period
                        End If

                        車位出租或占用情形 = IIf(IsDBNull(dr("出租或占用情形")), "", dr("出租或占用情形"))

                        '車位說明
                        車位其他說明事項 = IIf(IsDBNull(dr("車位說明")), "", dr("車位說明"))

                        '******* 車位細項 ******
                        '車獨立產權
                        If Not IsDBNull(dr("獨立產權")) Then
                            車獨立產權 = dr("獨立產權").ToString
                        End If
                        '車權利種類
                        If Not IsDBNull(dr("權利種類")) Then
                            車權利種類 = dr("權利種類").ToString
                        End If
                        '車建號
                        If Not IsDBNull(dr("車建號")) Then
                            車建號 = dr("車建號").ToString
                        End If
                        '車型式
                        If Not IsDBNull(dr("車位類別")) Then
                            車型式 = dr("車位類別").ToString
                        End If
                        '車編號
                        If Not IsDBNull(dr("車位號碼")) Then
                            車編號 = dr("車位號碼").ToString
                        End If
                        '車長
                        If Not IsDBNull(dr("車位_長")) Then
                            車長 = dr("車位_長").ToString
                        End If
                        '車寬
                        If Not IsDBNull(dr("車位_寬")) Then
                            車寬 = dr("車位_寬").ToString
                        End If
                        '車高
                        If Not IsDBNull(dr("車位_高")) Then
                            車高 = dr("車位_高").ToString
                        End If
                        '車重
                        If Not IsDBNull(dr("車位_承重")) Then
                            車重 = dr("車位_承重").ToString
                        End If
                        '******* 車位細項 ******

                        'If Request("oid") = "60692AAE34137" Then

                        'Else
                        '先把讀出的xml複製一份，接著開始改
                        Dim tempdata As String = myText_Car
                        tempdata = tempdata.Replace("≠車位產權說明", Server.HtmlEncode(車位產權說明))
                        tempdata = tempdata.Replace("≠車位位置", Server.HtmlEncode(車位位置))
                        tempdata = tempdata.Replace("≠車位價格", Server.HtmlEncode(車位價格))
                        tempdata = tempdata.Replace("≠車位性質", Server.HtmlEncode(車位性質))
                        tempdata = tempdata.Replace("≠車位出租或占用情形", Server.HtmlEncode(車位出租或占用情形))
                        tempdata = tempdata.Replace("≠車位管理費", Server.HtmlEncode(車位管理費))
                        tempdata = tempdata.Replace("≠車位其他說明事項", Server.HtmlEncode(車位其他說明事項))
                        '******* 車位細項 ******
                        tempdata = tempdata.Replace("≠車獨立產權", Server.HtmlEncode(車獨立產權))
                        tempdata = tempdata.Replace("≠車權利種類", Server.HtmlEncode(車權利種類))
                        tempdata = tempdata.Replace("≠車建號", Server.HtmlEncode(車建號))
                        tempdata = tempdata.Replace("≠車型式", Server.HtmlEncode(車型式))
                        tempdata = tempdata.Replace("≠車編號", Server.HtmlEncode(車編號))
                        tempdata = tempdata.Replace("≠車長", Server.HtmlEncode(車長))
                        tempdata = tempdata.Replace("≠車寬", Server.HtmlEncode(車寬))
                        tempdata = tempdata.Replace("≠車高", Server.HtmlEncode(車高))
                        tempdata = tempdata.Replace("≠車重", Server.HtmlEncode(車重))
                        '******* 車位細項 ******

                        '車位位置略圖
                        If Not IsDBNull(dr("車位照片")) Then
                            車位略圖 = "https://img.etwarm.com.tw/" & Request("sid") & "/available/" & dr("車位照片")
                            '使用 System.IO.File.Exists(車位略圖) 無法判別到是否有照片，故改寫法
                            Try
                                Dim imgbyts As Byte() = New System.Net.WebClient().DownloadData(車位略圖)
                                Dim xmlcarparkimg As String = "<w:br/><w:pict><w:binData w:name=""wordml://" & dr("車位照片") & """ xml:space=""preserve"">" & Xml2Doc.MyBase64Encode(imgbyts) & "</w:binData><v:shape id=""_x0000_" & dr("車位照片") & """ type=""#_x0000_t75"" style=""width:8.8cm;height:6.6cm""><v:imagedata src=""wordml://" & dr("車位照片") & """ o:title=""a0001""/></v:shape></w:pict>"
                                tempdata = tempdata.Replace("≠objCarPark", xmlcarparkimg)
                            Catch ex As Exception
                                tempdata = tempdata.Replace("≠objCarPark", "")
                            End Try
                        Else
                            tempdata = tempdata.Replace("≠objCarPark", "")
                        End If

                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_停車位產權說明.Append(tempdata)
                        'End If


                        'End If

                        '[END============================================產權調查-停車位產權說明=============================================END]
                    Next
                End Using
            End Using
            sFile = sFile.Replace("!重複停車位產權說明", 最後要替代掉的字串_停車位產權說明.ToString())
        End If  '-------判斷成屋還是土地用

        '共用
        '重要設施
        '[START============================================重要設施=============================================START]

        If Not IsDBNull(t2.Rows(0)("經度")) And Not IsDBNull(t2.Rows(0)("緯度")) Then
            '********讀取重要設施內容
            Dim store As String = Trim(t2.Rows(0)("店代號").ToString) '"A0001"  '加盟店
            Dim lat As String = Trim(t2.Rows(0)("經度").ToString) '"25.0448370467" '緯度
            Dim lng As String = Trim(t2.Rows(0)("緯度").ToString) '"121.524685661" '經度 
            Dim life1 As String = ""

            For i As Integer = 0 To CheckBoxList1.Items.Count - 1
                If CheckBoxList1.Items(i).Selected = True Then
                    life1 &= CheckBoxList1.Items(i).Value & ","
                End If
            Next


            If CheckBoxList2.Visible = True Then
                For i As Integer = 0 To CheckBoxList2.Items.Count - 1
                    If CheckBoxList2.Items(i).Selected = True Then
                        life1 &= CheckBoxList2.Items(i).Value & ","
                    End If
                Next
            End If

            Dim types As String = ""
            Dim str2 As String()
            str2 = life1.Split(",")
            i = 0
            For i As Integer = 0 To str2.Length - 1
                If str2(i) = "1" Then '公有市場
                    types &= "'H3201000','H3202000','H3203000',"
                ElseIf str2(i) = "2" Then '超級市場
                    types &= "'H0400001','H0400002','H0400003','H0400004','H0400005','H0400006','H0400007','H0400008','H0400009','H0400010','H0400011','H0400012','H0400013',"
                ElseIf str2(i) = "3" Then '學校
                    types &= "'B0200000','B0302000','B0303000','B0301000','B0400000','B0500000','B0800000',"
                ElseIf str2(i) = "4" Then '警察局(分駐所、派出所)
                    types &= "'A0302000',"
                ElseIf str2(i) = "5" Then '行政機關
                    types &= "'A0100000','A0201000','A0202000','A0203000','A0204000','A0205000','A0206000','A0207000','A0208000','A0210000','A0211000','A0212000','A0213000','A0215000',"
                ElseIf str2(i) = "19" Then '體育場
                    types &= "'G0404000',"
                ElseIf str2(i) = "6" Then '醫院
                    types &= "'C0100000','C0600000','C0800000','C1100000','C1200000',"
                ElseIf str2(i) = "7" Then '飛機場 
                    types &= "'K0400000',"
                ElseIf str2(i) = "8" Then '台電變電所用地 
                    types &= "'X0101000','X0102000','X0200000',"
                ElseIf str2(i) = "9" Then '地面高壓電塔(線) 
                    types &= "'X0300000',"
                ElseIf str2(i) = "20" Then '寺廟
                    types &= "'X0800000',"
                ElseIf str2(i) = "10" Then '殯儀館 
                    types &= "'L0504000',"
                ElseIf str2(i) = "11" Then '公墓
                    types &= "'L0501000',"
                ElseIf str2(i) = "12" Then '火化場
                    types &= "'L0502000',"
                ElseIf str2(i) = "13" Then '骨灰(骸)存放設施
                    types &= "'L0503000',"
                ElseIf str2(i) = "14" Then '垃圾場(掩埋場、焚化廠)
                    types &= "'X0500000','X0600000',"
                ElseIf str2(i) = "16" Then '加(氣)油站"
                    types &= "'J0101000','J0102000','J0103000','X0900000',"
                ElseIf str2(i) = "17" Then '瓦斯行(場)
                    types &= "'X1000000','X1100000',"
                ElseIf str2(i) = "18" Then '葬儀社
                    types &= "'L0505000',"
                End If
            Next

            sql = "SELECT   Code_type, name, Mangt_Add, latitude, longitude, Code, ( 6378.137 * acos( cos( radians(" & lat & ") ) * cos( radians( Latitude ) ) * cos( radians( Longitude ) - radians(" & lng & ") ) + sin( radians(" & lat & ") ) * sin( radians( Latitude ) ) ) ) AS distance  "
            sql &= " FROM life_funtion  where  "
            sql &= " (Code in (" & Mid(types, 1, Len(types) - 1) & ")) "
            sql &= " or (Code='H3100000' and (Subname like '奧斯卡%' or Subname like '東森%')) "
            sql &= " GROUP BY Code_type, name, Mangt_Add, latitude, longitude, Code"
            'sql &= " HAVING ( 6378.137 * acos( cos( radians(" & lat & ") ) * cos( radians( Latitude ) ) * cos( radians( Longitude ) - radians(" & lng & ") ) + sin( radians(" & lat & ") ) * sin( radians( Latitude ) ) ) ) < (0.3)"
            sql &= " ORDER BY distance ASC"

            'If Contract_No = "60692AAE41203" Then
            '    Response.Write(sql)
            '    Exit Sub
            'End If

            'If Request.Cookies("webfly_empno").Value.ToLower <> "92H" Then
            '    Response.Write(sql)
            '    Response.End()
            'End If

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "鄰近重要設施")
            Dim table_重要設施 As DataTable = ds.Tables("鄰近重要設施")
            'Dim dt As DataTable = table_重要設施

            Dim dt1 As DataTable
            Dim row As DataRow
            Dim ds1 As New DataSet

            '虛擬表格_統計表
            ds1.Tables.Add("重要設施")
            dt1 = ds1.Tables("重要設施")
            dt1.Columns.Add("id", GetType(System.String))
            dt1.Columns.Add("Code_type", GetType(System.String))
            dt1.Columns.Add("iconName", GetType(System.String))
            dt1.Columns.Add("Name", GetType(System.String))
            dt1.Columns.Add("Mangt_Add", GetType(System.String))
            dt1.Columns.Add("Longitude", GetType(System.String))
            dt1.Columns.Add("Latitude", GetType(System.String))

            If table_重要設施.Rows.Count > 0 Then

                For i As Integer = 0 To table_重要設施.Rows.Count - 1
                    If table_重要設施.Rows(i)("distance") <= 0.3 Then
                        row = dt1.NewRow

                        row("id") = i + 1

                        If table_重要設施.Rows(i)("Code") = "H3201000" Or table_重要設施.Rows(i)("Code") = "H3202000" Or table_重要設施.Rows(i)("Code") = "H3203000" Then '類別
                            row("Code_type") = "公有市場"
                            row("iconName") = "imp1.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "H0400001" Or table_重要設施.Rows(i)("Code") = "H0400002" Or table_重要設施.Rows(i)("Code") = "H0400003" Or table_重要設施.Rows(i)("Code") = "H0400004" Or table_重要設施.Rows(i)("Code") = "H0400005" Or table_重要設施.Rows(i)("Code") = "H0400006" Or table_重要設施.Rows(i)("Code") = "H0400007" Or table_重要設施.Rows(i)("Code") = "H0400008" Or table_重要設施.Rows(i)("Code") = "H0400009" Or table_重要設施.Rows(i)("Code") = "H0400010" Or table_重要設施.Rows(i)("Code") = "H0400011" Or table_重要設施.Rows(i)("Code") = "H0400012" Or table_重要設施.Rows(i)("Code") = "H0400013" Then
                            row("Code_type") = "超級市場"
                            row("iconName") = "imp2.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "B0200000" Or table_重要設施.Rows(i)("Code") = "B0302000" Or table_重要設施.Rows(i)("Code") = "B0303000" Or table_重要設施.Rows(i)("Code") = "B0301000" Or table_重要設施.Rows(i)("Code") = "B0400000" Or table_重要設施.Rows(i)("Code") = "B0500000" Or table_重要設施.Rows(i)("Code") = "B0800000" Then
                            row("Code_type") = "學校"
                            row("iconName") = "imp3.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "A0302000" Then
                            row("Code_type") = "警察局(分駐所、派出所)"
                            row("iconName") = "imp4.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "A0100000" Or table_重要設施.Rows(i)("Code") = "A0201000" Or table_重要設施.Rows(i)("Code") = "A0202000" Or table_重要設施.Rows(i)("Code") = "A0203000" Or table_重要設施.Rows(i)("Code") = "A0204000" Or table_重要設施.Rows(i)("Code") = "A0205000" Or table_重要設施.Rows(i)("Code") = "A0206000" Or table_重要設施.Rows(i)("Code") = "A0207000" Or table_重要設施.Rows(i)("Code") = "A0208000" Or table_重要設施.Rows(i)("Code") = "A0210000" Or table_重要設施.Rows(i)("Code") = "A0211000" Or table_重要設施.Rows(i)("Code") = "A0212000" Or table_重要設施.Rows(i)("Code") = "A0213000" Or table_重要設施.Rows(i)("Code") = "A0215000" Then
                            row("Code_type") = "行政機關"
                            row("iconName") = "imp5.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "G0404000" Then
                            row("Code_type") = "體育場"
                            row("iconName") = "imp19.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "C0100000" Or table_重要設施.Rows(i)("Code") = "C0600000" Or table_重要設施.Rows(i)("Code") = "C0800000" Or table_重要設施.Rows(i)("Code") = "C1100000" Or table_重要設施.Rows(i)("Code") = "C1200000" Then
                            row("Code_type") = "醫院"
                            row("iconName") = "imp6.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "K0400000" Then
                            row("Code_type") = "飛機場"
                            row("iconName") = "imp7.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "X0101000" Or table_重要設施.Rows(i)("Code") = "X0102000" Or table_重要設施.Rows(i)("Code") = "X0200000" Then
                            row("Code_type") = "台電變電所用地"
                            row("iconName") = "imp8.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "X0300000" Then
                            row("Code_type") = "地面高壓電塔(線)"
                            row("iconName") = "imp9.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "X0800000" Then
                            row("Code_type") = "寺廟"
                            row("iconName") = "imp20.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "L0504000" Then
                            row("Code_type") = "殯儀館"
                            row("iconName") = "imp10.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "L0501000" Then
                            row("Code_type") = "公墓"
                            row("iconName") = "imp11.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "L0502000" Then
                            row("Code_type") = "火化場"
                            row("iconName") = "imp12.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "L0503000" Then
                            row("Code_type") = "骨灰(骸)存放設施"
                            row("iconName") = "imp13.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "X0500000" Or table_重要設施.Rows(i)("Code") = "X0600000" Then
                            row("Code_type") = "垃圾場(掩埋場、焚化廠)"
                            row("iconName") = "imp14.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "J0101000" Or table_重要設施.Rows(i)("Code") = "J0102000" Or table_重要設施.Rows(i)("Code") = "J0103000" Or table_重要設施.Rows(i)("Code") = "X0900000" Then
                            row("Code_type") = "加(氣)油站"
                            row("iconName") = "imp16.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "X1000000" Or table_重要設施.Rows(i)("Code") = "X1100000" Then
                            row("Code_type") = "瓦斯行(場)"
                            row("iconName") = "imp17.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "L0505000" Then
                            row("Code_type") = "葬儀社"
                            row("iconName") = "imp18.png"
                        ElseIf table_重要設施.Rows(i)("Code") = "H3100000" Then
                            row("Code_type") = "逛街購物"
                            row("iconName") = "imp1.png"
                        End If

                        row("Name") = table_重要設施.Rows(i)("Name")
                        row("Mangt_Add") = table_重要設施.Rows(i)("Mangt_Add")
                        row("Longitude") = table_重要設施.Rows(i)("Longitude")
                        row("Latitude") = table_重要設施.Rows(i)("Latitude")

                        dt1.Rows.Add(row)
                    End If
                Next
                '重要設施列表dt1完成
            End If  '****End table_重要設施.Rows.Count

            Dim location As String = ""
            Dim modcount As Decimal = 0 '循環次數
            modcount = dt1.Rows.Count / EraEnvCount
            If modcount < 1 Then modcount = 1 '最小值為1
            Dim getmod As Integer = dt1.Rows.Count Mod EraEnvCount

            If (dt1.Rows.Count Mod EraEnvCount = 0) And dt1.Rows.Count <> 0 Then
                modcount = modcount - 1
                If modcount <= 1 Then
                    modcount = 1
                End If
            End If

            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            '    Response.Write(modcount & "," & dt1.Rows.Count)
            '    Exit Sub
            'End If
            'Response.Write(modcount & "," & dt1.Rows.Count)
            'Response.End()
            '重要設施
            Dim myXML_Import As String = ""
            Dim myText_Import As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_Import As New StreamReader(Server.MapPath("..\reports\鄰近重要設施參考圖_V3.txt"))
            myText_Import = srText_Import.ReadToEnd()
            srText_Import.Close()
            Dim 最後要替代掉的字串_鄰近重要設施參考圖 As New StringBuilder()

            '鄰近重要設施
            Dim myXML_鄰近重要設施 As String = ""
            Dim myText_鄰近重要設施 As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_鄰近重要設施 As New StreamReader(Server.MapPath("..\reports\重複重要設施內容_V2.txt"))
            myText_鄰近重要設施 = srText_鄰近重要設施.ReadToEnd()
            srText_鄰近重要設施.Close()

            Dim EnglishChar() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"}
            For k As Integer = 0 To modcount - 1  '如果有40筆 就會跑兩次 50筆跑3次
                '先把讀出的xml複製一份，接著開始改
                Dim tempdata_Import As String = myText_Import

                '********* 產生圖片 *****
                '判別地址定位是否能訂到

                Dim startcount, endcount As Integer
                startcount = EraEnvCount * k
                endcount = startcount + (EraEnvCount - 1)
                If endcount > dt1.Rows.Count Then
                    endcount = dt1.Rows.Count - 1
                End If
                'Response.Write(startcount & ",," & endcount)
                'Response.End()
                For oo As Integer = startcount To endcount Step 1  '跑重要設施的細項
                    If oo >= dt1.Rows.Count Then
                        Exit For
                    End If
                    'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
                    location &= "%7Clabel:" & EnglishChar(oo - startcount) & "," & dt1.Rows(oo)("Latitude") & "," & dt1.Rows(oo)("Longitude")
                    'Else
                    'location &= "&markers=color:blue%7Clabel:" & EnglishChar(oo - startcount) & "%7C" & dt1.Rows(oo)("Latitude") & "," & dt1.Rows(oo)("Longitude")
                    'End If
                Next
                'Dim imgpaths As String = "https://maps.google.com/maps/api/staticmap?language=zh-tw&center=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&markers=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&zoom=16&size=700x300&sensor=false" & location
                Dim imgpaths As String = ""

                ''If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then

                ''Else
                ''利用php抓取enc值, php為威呈撰寫
                ''傳遞給與經度、緯度、半徑
                'Dim str_url As String
                'str_url = "https://superweb2.etwarm.com.tw/staticmap.php?lat=" & t2.Rows(0)("經度") & "&lng=" & t2.Rows(0)("緯度") & "&radius=0.3"
                'Dim objHTTP = CreateObject("microsoft.XMLHTTP")
                'objHTTP.open("get", str_url, False)
                'objHTTP.setRequestHeader("CONTENT-Type", "text/html; charset=big5")
                'objHTTP.send()
                'Dim returnCODE As String = objHTTP.responseText
                'Dim encString As String = returnCODE
                ''End If

                'Dim lat = t2.Rows(0)("經度")
                'Dim lng = t2.Rows(0)("緯度")
                'Dim mapCentLat = lat + 0.002
                'Dim mapCentLng = lng - 0.011
                Dim mapCentLat = t2.Rows(0)("經度")
                Dim mapCentLng = t2.Rows(0)("緯度")
                Dim mapW = 700
                Dim mapH = 300
                Dim zoom = 16

                Dim circRadius = 0.8
                Dim circRadiusThick = 1
                Dim circFill = "FF0000"
                Dim circFillOpacity = 30
                Dim circBorder = "FF0000"

                Dim src As String = ""

                'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
                src = "https://map.etwarm.com.tw/RiChiMap/getImage.ashx?"
                src += "center=" & mapCentLat & "," & mapCentLng & "&"
                src += "zoom=" & zoom & "&"
                src += "size=" & mapW & "x" & mapH & "&"
                src += "format=jpg&markers=" & mapCentLat & "," & mapCentLng & ""
                If location <> "" Then src += Left(location, Len(location) - 1)
                src += "&circle=radius:300"
                imgpaths = src

                sql = " insert into 靜態地圖產生 "
                sql += " (店代號,員工代號,網址,產生程式,產生日期) "
                sql += " select '" & Request.QueryString("sid").ToString & "','" & Request.Cookies("webfly_empno").Value.ToString & "', "
                sql += " '" & imgpaths & "','委賣物件-不動產說明書',GETDATE() "
                Using conn1 As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn1.Open()
                    Using cmd3 As New SqlCommand(sql, conn1)
                        Try
                            cmd3.ExecuteNonQuery()
                        Catch ex As Exception
                            'Response.Write(ex.ToString)
                            'Response.End()
                            Throw ex
                        End Try
                    End Using
                End Using

                Try
                    '將圖檔轉成Base64Code後置換
                    Dim imgbyts As Byte() = New System.Net.WebClient().DownloadData(imgpaths)

                    Dim xmlEraStr As String = "<w:pict><w:binData w:name=""wordml://mapera" & k & ".png"" xml:space=""preserve"">" & Xml2Doc.MyBase64Encode(imgbyts) & "</w:binData><v:shape id=""_x0000_" & k & """ type=""#_x0000_t75"" style=""width:480pt;height:225pt""><v:imagedata src=""wordml://mapera" & k & ".png"" /></v:shape></w:pict>"

                    tempdata_Import = tempdata_Import.Replace("≠ObjPicImport", xmlEraStr)
                Catch ex As Exception
                    tempdata_Import = tempdata_Import.Replace("≠ObjPicImport", "")
                End Try

                location = ""

                '[START============================================鄰近重要設施=============================================START]
                Dim 周邊編號 As String = ""
                Dim 周邊種類 As String = ""
                Dim 周邊名稱 As String = ""
                Dim 周邊地址 As String = ""


                Dim 最後要替代掉的字串_重複鄰近重要設施列表 As New StringBuilder()
                For m As Integer = startcount To endcount Step 1 '跑重要設施的細項
                    If m >= dt1.Rows.Count Then
                        Exit For
                    End If

                    If dt1.Rows.Count > 0 Then
                        '周邊編號
                        周邊編號 = EnglishChar(m - startcount)

                        '周邊種類
                        If Not IsDBNull(dt1.Rows(m)("Code_type")) Then
                            周邊種類 = dt1.Rows(m)("Code_type").ToString
                        End If
                        '周邊名稱
                        If Not IsDBNull(dt1.Rows(m)("Name")) Then
                            周邊名稱 = dt1.Rows(m)("Name").ToString
                        End If
                        '周邊地址
                        If Not IsDBNull(dt1.Rows(m)("Mangt_Add")) Then
                            周邊地址 = dt1.Rows(m)("Mangt_Add").ToString
                        End If

                        '先把讀出的xml複製一份，接著開始改
                        Dim tempdata As String = myText_鄰近重要設施
                        tempdata = tempdata.Replace("≠周邊編號", 周邊編號)
                        tempdata = tempdata.Replace("≠周邊種類", 周邊種類)
                        tempdata = tempdata.Replace("≠周邊名稱", Server.HtmlEncode(周邊名稱))
                        tempdata = tempdata.Replace("≠周邊地址", Server.HtmlEncode(周邊地址))

                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_重複鄰近重要設施列表.Append(tempdata)
                    Else
                        Dim tempdata As String = myText_鄰近重要設施
                        tempdata = tempdata.Replace("≠周邊編號", "")
                        tempdata = tempdata.Replace("≠周邊種類", "")
                        tempdata = tempdata.Replace("≠周邊名稱", "")
                        tempdata = tempdata.Replace("≠周邊地址", "")

                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_重複鄰近重要設施列表.Append(tempdata)
                    End If
                    '[END============================================鄰近重要設施=============================================END]
                Next
                tempdata_Import = tempdata_Import.Replace("!重複重要設施內容", 最後要替代掉的字串_重複鄰近重要設施列表.ToString())
                '改完加到最後的字串裡面  
                最後要替代掉的字串_鄰近重要設施參考圖.Append(tempdata_Import)
            Next

            'If Request("oid") = "60692AAE34137" Then

            'Else
            sFile = sFile.Replace("!鄰近重要設施參考圖", 最後要替代掉的字串_鄰近重要設施參考圖.ToString())
            'End If

            '******** 判斷是否有其他頁
        End If

        '[END============================================重要設施=============================================END]

        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------判斷成屋還是土地用
            '產調-基地P1
            sFile = sFile.Replace("≠基地座落", 基地座落)
            sFile = sFile.Replace("≠土地使用分區", 土地使用分區)
            sFile = sFile.Replace("≠法定建蔽率", 法定建蔽率)
            sFile = sFile.Replace("≠法定容積率", 法定容積率)
            sFile = sFile.Replace("≠基地權利種類", 基地權利種類)
            sFile = sFile.Replace("≠基地所有權人", 基地所有權人)
            sFile = sFile.Replace("≠基地開發方式限制或其他負擔", 基地開發方式限制或其他負擔)
            sFile = sFile.Replace("≠建物基地座落", 建物基地座落)
        End If  '-------判斷成屋還是土地用


        'If Request("oid") = "60692AAE34137" Then

        'Else
        '共用
        '[START============================================重要交易條件=============================================START]
        '重要交易條件
        sFile = sFile.Replace("≠交易價金", 交易價金)
        sFile = sFile.Replace("≠第一期款", 第一期款)
        sFile = sFile.Replace("≠第二期款", 第二期款)
        sFile = sFile.Replace("≠第三期款", 第三期款)
        sFile = sFile.Replace("≠第四期款", 第四期款)
        sFile = sFile.Replace("≠一期", 一期)
        sFile = sFile.Replace("≠二期", 二期)
        sFile = sFile.Replace("≠三期", 三期)
        sFile = sFile.Replace("≠四期", 四期)
        sFile = sFile.Replace("≠土增稅賣", 土增稅賣)
        sFile = sFile.Replace("≠土增稅買", 土增稅買)
        sFile = sFile.Replace("≠土增稅雙", 土增稅雙)
        sFile = sFile.Replace("≠土增稅交", 土增稅交)
        sFile = sFile.Replace("≠土增稅約", 土增稅約)
        sFile = sFile.Replace("≠土增稅金", 土增稅金)

        sFile = sFile.Replace("≠地價稅賣", 地價稅賣)
        sFile = sFile.Replace("≠地價稅買", 地價稅買)
        sFile = sFile.Replace("≠地價稅雙", 地價稅雙)
        sFile = sFile.Replace("≠地價稅交", 地價稅交)
        sFile = sFile.Replace("≠地價稅約", 地價稅約)
        sFile = sFile.Replace("≠地價稅金", 地價稅金)

        sFile = sFile.Replace("≠房屋稅賣", 房屋稅賣)
        sFile = sFile.Replace("≠房屋稅買", 房屋稅買)
        sFile = sFile.Replace("≠房屋稅雙", 房屋稅雙)
        sFile = sFile.Replace("≠房屋稅交", 房屋稅交)
        sFile = sFile.Replace("≠房屋稅約", 房屋稅約)
        sFile = sFile.Replace("≠房屋稅金", 房屋稅金)

        sFile = sFile.Replace("≠契稅賣", 契稅賣)
        sFile = sFile.Replace("≠契稅買", 契稅買)
        sFile = sFile.Replace("≠契稅雙", 契稅雙)
        sFile = sFile.Replace("≠契稅交", 契稅交)
        sFile = sFile.Replace("≠契稅約", 契稅約)
        sFile = sFile.Replace("≠契稅金", 契稅金)

        sFile = sFile.Replace("≠印花稅賣", 印花稅賣)
        sFile = sFile.Replace("≠印花稅買", 印花稅買)
        sFile = sFile.Replace("≠印花稅雙", 印花稅雙)
        sFile = sFile.Replace("≠印花稅交", 印花稅交)
        sFile = sFile.Replace("≠印花稅約", 印花稅約)
        sFile = sFile.Replace("≠印花稅金", 印花稅金)

        sFile = sFile.Replace("≠代書費賣", 代書費賣)
        sFile = sFile.Replace("≠代書費買", 代書費買)
        sFile = sFile.Replace("≠代書費雙", 代書費雙)
        sFile = sFile.Replace("≠代書費交", 代書費交)
        sFile = sFile.Replace("≠代書費約", 代書費約)
        sFile = sFile.Replace("≠代書費金", 代書費金)

        sFile = sFile.Replace("≠登記費賣", 登記費賣)
        sFile = sFile.Replace("≠登記費買", 登記費買)
        sFile = sFile.Replace("≠登記費雙", 登記費雙)
        sFile = sFile.Replace("≠登記費交", 登記費交)
        sFile = sFile.Replace("≠登記費約", 登記費約)
        sFile = sFile.Replace("≠登記費金", 登記費金)

        sFile = sFile.Replace("≠公證費賣", 公證費賣)
        sFile = sFile.Replace("≠公證費買", 公證費買)
        sFile = sFile.Replace("≠公證費雙", 公證費雙)
        sFile = sFile.Replace("≠公證費交", 公證費交)
        sFile = sFile.Replace("≠公證費約", 公證費約)
        sFile = sFile.Replace("≠公證費金", 公證費金)

        sFile = sFile.Replace("≠水電費賣", 水電費賣)
        sFile = sFile.Replace("≠水電費買", 水電費買)
        sFile = sFile.Replace("≠水電費雙", 水電費雙)
        sFile = sFile.Replace("≠水電費交", 水電費交)
        sFile = sFile.Replace("≠水電費約", 水電費約)
        sFile = sFile.Replace("≠水電費金", 水電費金)

        sFile = sFile.Replace("≠瓦斯費賣", 瓦斯費賣)
        sFile = sFile.Replace("≠瓦斯費買", 瓦斯費買)
        sFile = sFile.Replace("≠瓦斯費雙", 瓦斯費雙)
        sFile = sFile.Replace("≠瓦斯費交", 瓦斯費交)
        sFile = sFile.Replace("≠瓦斯費約", 瓦斯費約)
        sFile = sFile.Replace("≠瓦斯費金", 瓦斯費金)

        sFile = sFile.Replace("≠管理費賣", 管理費賣)
        sFile = sFile.Replace("≠管理費買", 管理費買)
        sFile = sFile.Replace("≠管理費雙", 管理費雙)
        sFile = sFile.Replace("≠管理費交", 管理費交)
        sFile = sFile.Replace("≠管理費約", 管理費約)
        sFile = sFile.Replace("≠管理費金", 管理費金)

        sFile = sFile.Replace("≠電話費賣", 電話費賣)
        sFile = sFile.Replace("≠電話費買", 電話費買)
        sFile = sFile.Replace("≠電話費雙", 電話費雙)
        sFile = sFile.Replace("≠電話費交", 電話費交)
        sFile = sFile.Replace("≠電話費約", 電話費約)
        sFile = sFile.Replace("≠電話費金", 電話費金)

        sFile = sFile.Replace("≠工程費賣", 工程費賣)
        sFile = sFile.Replace("≠工程費買", 工程費買)
        sFile = sFile.Replace("≠工程費雙", 工程費雙)
        sFile = sFile.Replace("≠工程費交", 工程費交)
        sFile = sFile.Replace("≠工程費約", 工程費約)
        sFile = sFile.Replace("≠工程費金", 工程費金)

        'sFile = sFile.Replace("≠奢侈稅賣", 奢侈稅賣)
        'sFile = sFile.Replace("≠奢侈稅買", 奢侈稅買)
        'sFile = sFile.Replace("≠奢侈稅雙", 奢侈稅雙)
        'sFile = sFile.Replace("≠奢侈稅交", 奢侈稅交)
        'sFile = sFile.Replace("≠奢侈稅約", 奢侈稅約)
        'sFile = sFile.Replace("≠奢侈稅金", 奢侈稅金)

        sFile = sFile.Replace("≠房地合一賣", 房地合一賣)
        sFile = sFile.Replace("≠房地合一買", 房地合一買)
        sFile = sFile.Replace("≠房地合一雙", 房地合一雙)
        sFile = sFile.Replace("≠房地合一交", 房地合一交)
        sFile = sFile.Replace("≠房地合一約", 房地合一約)
        sFile = sFile.Replace("≠房地合一金", 房地合一金)

        sFile = sFile.Replace("≠實價登錄賣", 實價登錄賣)
        sFile = sFile.Replace("≠實價登錄買", 實價登錄買)
        sFile = sFile.Replace("≠實價登錄雙", 實價登錄雙)
        sFile = sFile.Replace("≠實價登錄交", 實價登錄交)
        sFile = sFile.Replace("≠實價登錄約", 實價登錄約)
        sFile = sFile.Replace("≠實價登錄金", 實價登錄金)
        sFile = sFile.Replace("≠代書費New賣", 代書費New賣)
        sFile = sFile.Replace("≠代書費New買", 代書費New買)
        sFile = sFile.Replace("≠代書費New雙", 代書費New雙)
        sFile = sFile.Replace("≠代書費New交", 代書費New交)
        sFile = sFile.Replace("≠代書費New約", 代書費New約)
        sFile = sFile.Replace("≠代書費New金", 代書費New金)
        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '附贈設備
            sFile = sFile.Replace("≠買方設備0", 買方設備)
            sFile = sFile.Replace("≠買方設備1", 買方設備1)
            sFile = sFile.Replace("≠買方設備2", 買方設備2)
            sFile = sFile.Replace("≠買方設備3", 買方設備3)
        End If  '-------------------------------------------------------------------------判斷成屋還是土地用
        '照片說明
        'If Request("oid") = "60692AAE34074" Then

        'Else
        sFile = sFile.Replace("≠環境介紹", Server.HtmlEncode(環境介紹))
        'End If

        'End If

        '[END============================================重要交易條件=============================================END]




        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-建物標示說明P1=============================================START]
            '建物面積細項
            Dim myXML_House As String = ""
            Dim myText_House As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_House As New StreamReader(Server.MapPath("..\reports\重複建物面積細項_V3.txt")) '專屬print_V4用
            myText_House = srText_House.ReadToEnd()
            srText_House.Close()

            '建物標示說明 - 層次面積
            'sqlstr = "Select *,(權利範圍1分子+'/'+權利範圍1分母) as 權利範圍,(權利範圍2分子+'/'+權利範圍2分母) as 權利範圍2 From 委賣物件資料表_面積細項  With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別<>'土地面積'   order by 流水號 "
            sqlstr = "Select (  a.權利範圍_分子  +'/'+ a.權利範圍_分母  ) as 權利範圍,( a.出售權利範圍_分子 +'/'+ a.出售權利範圍_分母  ) as 權利範圍2, (權利範圍1分子  +'/'+ 權利範圍1分母  ) as 權利範圍1_old, (權利範圍2分子  +'/'+ 權利範圍2分母  ) as 權利範圍2_old, * "
            sqlstr &= " From 委賣物件資料表_面積細項 b With(NoLock) "
            sqlstr &= " LEFT JOIN 委賣物件資料表_細項所有權人 a ON a.物件編號 = b.物件編號 AND a.店代號 = b.店代號 AND a.細項流水號 = b.流水號  "
            sqlstr &= " Where b.物件編號 = '" & Contract_No & "' and b.店代號='" & Request("sid") & "' and b.類別 <> '土地面積'  order by b.流水號"

            Dim table4 As DataTable
            Dim 建物建號 As String = ""
            Dim 建物型態 As String = ""
            Dim 建物用途 As String = ""
            Dim 建物總面積 As String = ""
            Dim 建物_權利範圍 As String = ""
            Dim 建物持份平方 As String = ""
            Dim 建物持份坪 As String = ""
            Dim 建物平方和 As Decimal = 0
            Dim 建物坪和 As Decimal = 0


            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table4")
            table4 = ds.Tables("table4")
            Dim 最後要替代掉的字串_建物面積細項 As New StringBuilder()

            If table4.Rows.Count > 0 Then
                For i = 0 To table4.Rows.Count - 1

                    '建物建號
                    If Not IsDBNull(table4.Rows(i)("建號")) Then
                        建物建號 = table4.Rows(i)("建號").ToString
                    Else
                        建物建號 = ""
                    End If


                    '建物型態
                    If Not IsDBNull(table4.Rows(i)("類別")) Then
                        建物型態 = table4.Rows(i)("類別").ToString
                        If 建物型態 = "主建物" Then
                            If table4.Rows(i)("DL_level2_selectindex").ToString = "0" Then
                                建物型態 = "主建物"
                            End If
                            If table4.Rows(i)("DL_level2_selectindex").ToString = "1" Then
                                建物型態 = "附屬建物"
                            End If
                            If table4.Rows(i)("DL_level2_selectindex").ToString = "2" Then
                                建物型態 = "共有部分"
                            End If
                            If table4.Rows(i)("是否為車位").ToString = "Y" Then
                                建物型態 += "(產權獨立車位)"
                            End If
                        End If
                    Else
                        建物型態 = ""
                    End If

                    '20151021 如果車位建地號為空值則不顯示 by nick
                    If 建物建號 = "" And 建物型態.LastIndexOf("車位") >= 0 Then
                        Continue For
                    End If

                    '建物用途
                    If Not IsDBNull(table4.Rows(i)("項目名稱")) Then
                        建物用途 = table4.Rows(i)("項目名稱").ToString
                    Else
                        建物用途 = ""
                    End If

                    If Not IsDBNull(table4.Rows(i)("總面積平方公尺")) Then
                        If Right(table4.Rows(i)("總面積平方公尺").ToString, 5) = ".0000" Then
                            建物總面積 = Int(table4.Rows(i)("總面積平方公尺").ToString)
                        ElseIf table4.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("總面積平方公尺").ToString, 3) = "000" Then
                            建物總面積 = Left(table4.Rows(i)("總面積平方公尺").ToString, Len(table4.Rows(i)("總面積平方公尺").ToString) - 3)
                        ElseIf table4.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("總面積平方公尺").ToString, 2) = "00" Then
                            建物總面積 = Left(table4.Rows(i)("總面積平方公尺").ToString, Len(table4.Rows(i)("總面積平方公尺").ToString) - 2)
                        ElseIf table4.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("總面積平方公尺").ToString, 1) = "0" Then
                            建物總面積 = Left(table4.Rows(i)("總面積平方公尺").ToString, Len(table4.Rows(i)("總面積平方公尺").ToString) - 1)
                        Else
                            建物總面積 = table4.Rows(i)("總面積平方公尺").ToString
                        End If
                    Else
                        建物總面積 = "0"
                    End If



                    '建物權利範圍
                    Dim 權利範圍1分子 As String = ""
                    Dim 權利範圍1分母 As String = ""
                    Dim 權利範圍2分子 As String = ""
                    Dim 權利範圍2分母 As String = ""

                    'If Not IsDBNull(table4.Rows(i)("權利範圍")) Then
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍")
                    'ElseIf IsDBNull(table4.Rows(i)("權利範圍1_old")) Then
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍1_old")
                    'End If

                    'If Request("oid") = "10602837819" Then
                    建物_權利範圍 = ""
                    Try
                        If Not IsDBNull(table4.Rows(i)("權利範圍2_old")) Then
                            If table4.Rows(i)("權利範圍2_old") <> "1/1" Then
                                If Not IsDBNull(table4.Rows(i)("權利範圍2")) Then
                                    權利範圍1分子 = CInt(table4.Rows(i)("出售權利範圍_分子")) * CInt(table4.Rows(i)("權利範圍2分子"))
                                    權利範圍1分母 = CInt(table4.Rows(i)("出售權利範圍_分母")) * CInt(table4.Rows(i)("權利範圍2分母"))
                                Else
                                    '權利範圍1分子 = CInt(table4.Rows(i)("權利範圍_分子")) * CInt(table4.Rows(i)("權利範圍2分子"))
                                    '權利範圍1分母 = CInt(table4.Rows(i)("權利範圍_分母")) * CInt(table4.Rows(i)("權利範圍2分母"))
                                    If Not IsDBNull(table4.Rows(i)("權利範圍_分母")) Then
                                        權利範圍1分子 = CInt(table4.Rows(i)("權利範圍_分子")) * CInt(table4.Rows(i)("權利範圍2分子"))
                                        權利範圍1分母 = CInt(table4.Rows(i)("權利範圍_分母")) * CInt(table4.Rows(i)("權利範圍2分母"))
                                    Else
                                        權利範圍1分子 = CInt(table4.Rows(i)("權利範圍2分子"))
                                        權利範圍1分母 = CInt(table4.Rows(i)("權利範圍2分母"))
                                    End If

                                End If
                                建物_權利範圍 = 權利範圍1分子 & "/" & 權利範圍1分母
                            End If
                        End If
                        If 建物_權利範圍 = "" Then
                            If Not IsDBNull(table4.Rows(i)("權利範圍1_old")) Then
                                If Not IsDBNull(table4.Rows(i)("權利範圍2")) Then
                                    權利範圍1分子 = CInt(table4.Rows(i)("出售權利範圍_分子")) * CInt(table4.Rows(i)("權利範圍1分子"))
                                    權利範圍1分母 = CInt(table4.Rows(i)("出售權利範圍_分母")) * CInt(table4.Rows(i)("權利範圍1分母"))
                                Else
                                    If Not IsDBNull(table4.Rows(i)("權利範圍_分母")) Then
                                        權利範圍1分子 = CInt(table4.Rows(i)("權利範圍_分子")) * CInt(table4.Rows(i)("權利範圍1分子"))
                                        權利範圍1分母 = CInt(table4.Rows(i)("權利範圍_分母")) * CInt(table4.Rows(i)("權利範圍1分母"))
                                    Else
                                        權利範圍1分子 = CInt(table4.Rows(i)("權利範圍1分子"))
                                        權利範圍1分母 = CInt(table4.Rows(i)("權利範圍1分母"))
                                    End If
                                End If
                                建物_權利範圍 = 權利範圍1分子 & "/" & 權利範圍1分母
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                    'Else
                    '    'If Not IsDBNull(table4.Rows(i)("權利範圍")) Then
                    '    '    建物_權利範圍 = table4.Rows(i)("權利範圍")
                    '    'ElseIf IsDBNull(table4.Rows(i)("權利範圍1_old")) Then
                    '    '    建物_權利範圍 = table4.Rows(i)("權利範圍1_old")
                    '    'End If
                    'End If


                    'If Not IsDBNull(table4.Rows(i)("權利範圍")) Then
                    '    If table4.Rows(i)("權利範圍") = "1/1" Then
                    '        建物_權利範圍 = table4.Rows(i)("權利範圍1_old")
                    '    End If
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍")
                    'ElseIf IsDBNull(table4.Rows(i)("權利範圍1_old")) Then
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍1_old")
                    'End If

                    ''權利範圍1分子
                    'If Not IsDBNull(table4.Rows(i)("權利範圍1分子")) Then
                    '    If table4.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                    '        權利範圍1分子 = "1"
                    '    Else
                    '        權利範圍1分子 = table4.Rows(i)("權利範圍1分子").ToString
                    '    End If
                    'Else
                    '    權利範圍1分子 = "1"
                    'End If

                    ''權利範圍1分母
                    'If Not IsDBNull(table4.Rows(i)("權利範圍1分母")) Then
                    '    If table4.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                    '        權利範圍1分母 = "1"
                    '    Else
                    '        權利範圍1分母 = table4.Rows(i)("權利範圍1分母").ToString
                    '    End If
                    'Else
                    '    權利範圍1分母 = "1"
                    'End If

                    ''權利範圍2分子
                    'If Not IsDBNull(table4.Rows(i)("權利範圍2分子")) Then
                    '    If table4.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                    '        權利範圍2分子 = "1"
                    '    Else
                    '        權利範圍2分子 = table4.Rows(i)("權利範圍2分子").ToString
                    '    End If
                    'Else
                    '    權利範圍2分子 = "1"
                    'End If

                    ''權利範圍2分母
                    'If Not IsDBNull(table4.Rows(i)("權利範圍2分母")) Then
                    '    If table4.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                    '        權利範圍2分母 = "1"
                    '    Else
                    '        權利範圍2分母 = table4.Rows(i)("權利範圍2分母").ToString
                    '    End If
                    'Else
                    '    權利範圍2分母 = "1"
                    'End If

                    '建物_權利範圍 = CType(CType(權利範圍1分子, Decimal) * CType(權利範圍2分子, Decimal), String) & "/" & CType(CType(權利範圍1分母, Decimal) * CType(權利範圍2分母, Decimal), String)

                    ' 2016.06.08 by Finch 如果"委賣物件資料表_細項所有權人"權利範圍沒有值，改從"委賣物件資料表_面積細項"裡抓舊版輸入的值
                    'If Not IsDBNull(table4.Rows(i)("權利範圍")) Then
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍")
                    'ElseIf IsDBNull(table4.Rows(i)("權利範圍1_old")) Then
                    '    建物_權利範圍 = table4.Rows(i)("權利範圍1_old")
                    'End If

                    ''建物持份平方
                    'If Not IsDBNull(table4.Rows(i)("實際持有平方公尺")) Then

                    '    If 建物型態 <> "庭院坪數" And 建物型態 <> "增建" And 建物型態 <> "車位面積(產權獨立)" Then
                    '        '總和
                    '        建物平方和 += table4.Rows(i)("實際持有平方公尺")
                    '    End If

                    '    If Right(table4.Rows(i)("實際持有平方公尺").ToString, 5) = ".0000" Then
                    '        建物持份平方 = Int(table4.Rows(i)("實際持有平方公尺").ToString)
                    '    ElseIf table4.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有平方公尺").ToString, 3) = "000" Then
                    '        建物持份平方 = Left(table4.Rows(i)("實際持有平方公尺").ToString, Len(table4.Rows(i)("實際持有平方公尺").ToString) - 3)
                    '    ElseIf table4.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有平方公尺").ToString, 2) = "00" Then
                    '        建物持份平方 = Left(table4.Rows(i)("實際持有平方公尺").ToString, Len(table4.Rows(i)("實際持有平方公尺").ToString) - 2)
                    '    ElseIf table4.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有平方公尺").ToString, 1) = "0" Then
                    '        建物持份平方 = Left(table4.Rows(i)("實際持有平方公尺").ToString, Len(table4.Rows(i)("實際持有平方公尺").ToString) - 1)
                    '    Else
                    '        建物持份平方 = table4.Rows(i)("實際持有平方公尺").ToString
                    '    End If

                    'Else
                    '    建物持份平方 = ""
                    'End If

                    ''建物持份坪
                    'If Not IsDBNull(table4.Rows(i)("實際持有坪")) Then
                    '    If 建物型態 <> "庭院坪數" And 建物型態 <> "增建" And 建物型態 <> "車位面積(產權獨立)" Then
                    '        '總和
                    '        建物坪和 += table4.Rows(i)("實際持有坪")
                    '    End If

                    '    If Right(table4.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                    '        建物持份坪 = Int(table4.Rows(i)("實際持有坪").ToString)
                    '    ElseIf table4.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                    '        建物持份坪 = Left(table4.Rows(i)("實際持有坪").ToString, Len(table4.Rows(i)("實際持有坪").ToString) - 3)
                    '    ElseIf table4.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                    '        建物持份坪 = Left(table4.Rows(i)("實際持有坪").ToString, Len(table4.Rows(i)("實際持有坪").ToString) - 2)
                    '    ElseIf table4.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table4.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                    '        建物持份坪 = Left(table4.Rows(i)("實際持有坪").ToString, Len(table4.Rows(i)("實際持有坪").ToString) - 1)
                    '    Else
                    '        建物持份坪 = table4.Rows(i)("實際持有坪").ToString
                    '    End If
                    'Else
                    '    建物持份坪 = ""
                    'End If

                    If Not IsDBNull(table4.Rows(i)("持有面積")) Then
                        '總和
                        If 建物型態 <> "庭院坪數" And 建物型態 <> "增建" And 建物型態 <> "車位面積(產權獨立)" Then
                            'If Request("oid") = "106021039416-01" Then
                            If Not IsDBNull(table4.Rows(i)("出售面積")) Then
                                If CDec(table4.Rows(i)("出售面積")) > 0 Then
                                    建物平方和 += CDec(table4.Rows(i)("出售面積"))
                                Else
                                    建物平方和 += CDec(table4.Rows(i)("持有面積"))
                                End If
                            Else
                                建物平方和 += CDec(table4.Rows(i)("持有面積"))
                            End If
                            'Else
                            '    建物平方和 += CDec(table4.Rows(i)("持有面積"))
                            '    '建物持份平方 = table4.Rows(i)("持有面積")
                            'End If
                        End If
                        'If Request("oid") = "106021039416-01" Then
                        If Not IsDBNull(table4.Rows(i)("出售面積")) Then
                            If CDec(table4.Rows(i)("出售面積")) > 0 Then
                                建物持份平方 = table4.Rows(i)("出售面積")

                                建物持份坪 = Format(CDec(建物持份平方) * 0.3025, "0.000000")
                                建物坪和 = Math.Round(CDec(建物平方和) * 0.3025, 4)
                            Else
                                建物持份平方 = table4.Rows(i)("持有面積")

                                建物持份坪 = Format(CDec(建物持份平方) * 0.3025, "0.000000")
                                建物坪和 = Math.Round(CDec(建物平方和) * 0.3025, 4)
                            End If
                        Else
                            建物持份平方 = table4.Rows(i)("持有面積")

                            建物持份坪 = Format(CDec(建物持份平方) * 0.3025, "0.000000")
                            建物坪和 = Math.Round(CDec(建物平方和) * 0.3025, 4)
                        End If
                        'Else
                        '    建物持份平方 = table4.Rows(i)("持有面積")

                        '    建物持份坪 = CDec(建物持份平方) * 0.3025
                        '    建物坪和 = Math.Round(CDec(建物平方和) * 0.3025, 4)
                        'End If
                    ElseIf Not IsDBNull(table4.Rows(i)("實際持有平方公尺")) Then

                        If 建物型態 <> "庭院坪數" And 建物型態 <> "增建" And 建物型態 <> "車位面積(產權獨立)" Then
                            建物平方和 += CDec(table4.Rows(i)("實際持有平方公尺"))
                            '建物持份平方 = table4.Rows(i)("實際持有平方公尺")

                            建物坪和 += CDec(table4.Rows(i)("實際持有坪"))
                        End If
                        建物持份平方 = table4.Rows(i)("實際持有平方公尺")
                        建物持份坪 = CDec(table4.Rows(i)("實際持有坪"))
                    Else
                        建物持份平方 = 0
                        建物持份坪 = 0
                        建物坪和 = 0
                    End If

                    '建物持份坪 = CDec(建物持份平方) * 0.3025
                    '建物坪和 = Math.Round(CDec(建物平方和) * 0.3025, 4)

                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_House
                    tempdata = tempdata.Replace("≠建物建號", 建物建號)
                    tempdata = tempdata.Replace("≠建物型態", 建物型態)
                    tempdata = tempdata.Replace("≠建物用途", 建物用途)
                    tempdata = tempdata.Replace("≠建物總面積", IIf(建物型態 = "增建", "約", "") & 建物總面積)
                    tempdata = tempdata.Replace("≠建物權利範圍", 建物_權利範圍)
                    tempdata = tempdata.Replace("≠建物持份平方", IIf(建物型態 = "增建", "約", "") & 建物持份平方)
                    tempdata = tempdata.Replace("≠建物持份坪", IIf(建物型態 = "增建", "約", "") & 建物持份坪)
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_建物面積細項.Append(tempdata)
                Next
            Else
                Dim tempdata As String = myText_House
                tempdata = tempdata.Replace("≠建物建號", "")
                tempdata = tempdata.Replace("≠建物型態", "")
                tempdata = tempdata.Replace("≠建物用途", "")
                tempdata = tempdata.Replace("≠建物總面積", "")
                tempdata = tempdata.Replace("≠建物權利範圍", "")
                tempdata = tempdata.Replace("≠建物持份平方", "")
                tempdata = tempdata.Replace("≠建物持份坪", "")

                '改完加到最後的字串裡面  
                最後要替代掉的字串_建物面積細項.Append(tempdata)
            End If
            sFile = sFile.Replace("!重複建物面積細項", 最後要替代掉的字串_建物面積細項.ToString())

            '建物細項小計
            '建物平方公尺
            'If 建物平方和 <> "" Then
            If Right(建物平方和.ToString, 5) = ".0000" Then
                建物平方和 = Int(建物平方和.ToString)
            ElseIf 建物平方和.ToString.LastIndexOf(".") > 0 And Right(建物平方和.ToString, 3) = "000" Then
                建物平方和 = Left(建物平方和.ToString, Len(建物平方和.ToString) - 3)
            ElseIf 建物平方和.ToString.LastIndexOf(".") > 0 And Right(建物平方和.ToString, 2) = "00" Then
                建物平方和 = Left(建物平方和.ToString, Len(建物平方和.ToString) - 2)
            ElseIf 建物平方和.ToString.LastIndexOf(".") > 0 And Right(建物平方和.ToString, 1) = "0" Then
                建物平方和 = Left(建物平方和.ToString, Len(建物平方和.ToString) - 1)
            Else
                建物平方和 = 建物平方和.ToString
            End If
            'End If

            '建物坪
            'If 建物坪和 <> "" Then
            If Right(建物坪和.ToString, 5) = ".0000" Then
                建物坪和 = Int(建物坪和.ToString)
            ElseIf 建物坪和.ToString.LastIndexOf(".") > 0 And Right(建物坪和.ToString, 3) = "000" Then
                建物坪和 = Left(建物坪和.ToString, Len(建物坪和.ToString) - 3)
            ElseIf 建物坪和.ToString.LastIndexOf(".") > 0 And Right(建物坪和.ToString, 2) = "00" Then
                建物坪和 = Left(建物坪和.ToString, Len(建物坪和.ToString) - 2)
            ElseIf 建物坪和.ToString.LastIndexOf(".") > 0 And Right(建物坪和.ToString, 1) = "0" Then
                建物坪和 = Left(建物坪和.ToString, Len(建物坪和.ToString) - 1)
            Else
                建物坪和 = 建物坪和.ToString
            End If
            'End If

            sFile = sFile.Replace("≠建物平方和", 建物平方和)
            sFile = sFile.Replace("≠建物坪和", 建物坪和)

            '建物權利人
            Dim myXML_HOwner As String = ""
            Dim myText_HOwner As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_HOwner As New StreamReader(Server.MapPath("..\reports\重複建物權利人_V2.txt"))
            myText_HOwner = srText_HOwner.ReadToEnd()
            srText_HOwner.Close()

            sqlstr = ""
            sqlstr = "Select * From 物件他項權利細項  With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 權利類別='建物'  order by 順位 asc,權利類別 desc  "

            Dim table5 As DataTable
            Dim 建物權利種類 As String = ""
            Dim 建物順位 As String = ""
            Dim 建物登記日期 As String = ""
            Dim 建物設定性質設定金額 As String = ""
            Dim 建物權利人 As String = ""
            Dim 建物管理人 As String = ""
            Dim 建物處理方式 As String = ""

            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table5")
            table5 = ds.Tables("table5")
            Dim 最後要替代掉的字串_建物權利人 As New StringBuilder()

            If t2.Rows(0)("與土地他項權利部相同").ToString = "1" Then
                '判斷建物他項無值時，才執行下段
                If table5.Rows.Count = 0 Then
                    '建物權利種類
                    建物權利種類 = "與基地他項權利部同"
                    '建物登記日期
                    建物登記日期 = "-------"
                    '建物設定性質設定金額
                    建物設定性質設定金額 = "---------"
                    '建物權利人
                    建物權利人 = "---------------------"
                    '建物管理人
                    建物管理人 = "-------"


                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_HOwner

                    tempdata = tempdata.Replace("≠建物權利種類", 建物權利種類)
                    tempdata = tempdata.Replace("≠建登記日期", 建物登記日期)
                    tempdata = tempdata.Replace("≠建設定性質設定金額", 建物設定性質設定金額)
                    tempdata = tempdata.Replace("≠建權利人", 建物權利人)
                    tempdata = tempdata.Replace("≠建管理人", 建物管理人)
                    tempdata = tempdata.Replace("≠這裡放處理方式", "與基地他項權利部同")
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_建物權利人.Append(tempdata)
                Else
                    For i = 0 To table5.Rows.Count - 1
                        建物處理方式 = ""
                        If Not IsDBNull(table5.Rows(i)("處理方式1")) Then
                            If table5.Rows(i)("處理方式1") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式1") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式2")) Then
                            If table5.Rows(i)("處理方式2") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式2") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式3")) Then
                            If table5.Rows(i)("處理方式3") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式3") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式4")) Then
                            If table5.Rows(i)("處理方式4") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式4") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式5")) Then
                            If table5.Rows(i)("處理方式5") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式5") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式6")) Then
                            If table5.Rows(i)("處理方式6") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式6") & ","
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("其他說明")) Then
                            If table5.Rows(i)("其他說明") <> "" Then
                                建物處理方式 &= table5.Rows(i)("其他說明")
                            End If
                        End If
                        If Right(建物處理方式, 1) = "," Then
                            建物處理方式 = 建物處理方式.Substring(0, 建物處理方式.Length - 1)
                        End If

                        '建物權利種類
                        If Not IsDBNull(table5.Rows(i)("權利種類")) Then
                            建物權利種類 = table5.Rows(i)("權利種類")
                        Else
                            建物權利種類 = ""
                        End If

                        '建物登記日期
                        If Not IsDBNull(table5.Rows(i)("登記日期")) Then
                            建物登記日期 = table5.Rows(i)("登記日期").ToString
                        Else
                            建物登記日期 = ""
                        End If

                        '建物設定性質設定金額
                        If Not IsDBNull(table5.Rows(i)("設定")) Then
                            建物設定性質設定金額 = table5.Rows(i)("設定").ToString & "萬"
                        Else
                            建物設定性質設定金額 = ""
                        End If

                        '建物權利人
                        If Not IsDBNull(table5.Rows(i)("設定權利人")) Then
                            建物權利人 = table5.Rows(i)("設定權利人").ToString
                        Else
                            建物權利人 = ""
                        End If

                        '建物管理人
                        If Not IsDBNull(table5.Rows(i)("管理人")) Then
                            建物管理人 = table5.Rows(i)("管理人").ToString
                        Else
                            建物管理人 = ""
                        End If


                        '先把讀出的xml複製一份，接著開始改
                        Dim tempdata As String = myText_HOwner

                        tempdata = tempdata.Replace("≠建物權利種類", 建物權利種類)
                        tempdata = tempdata.Replace("≠建登記日期", 建物登記日期)
                        tempdata = tempdata.Replace("≠建設定性質設定金額", 建物設定性質設定金額)
                        tempdata = tempdata.Replace("≠建權利人", 建物權利人)
                        tempdata = tempdata.Replace("≠建管理人", 建物管理人)
                        tempdata = tempdata.Replace("≠這裡放處理方式", 建物處理方式)
                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_建物權利人.Append(tempdata)
                    Next
                End If

                If t2.Rows(0)("其他如下").ToString = "1" Then
                    If table5.Rows.Count > 0 Then
                        '判斷是否勾選,建物與土地權利部同
                        For i As Integer = 0 To table5.Rows.Count - 1
                            建物處理方式 = ""
                            If Not IsDBNull(table5.Rows(i)("處理方式1")) Then
                                If table5.Rows(i)("處理方式1") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式1") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("處理方式2")) Then
                                If table5.Rows(i)("處理方式2") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式2") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("處理方式3")) Then
                                If table5.Rows(i)("處理方式3") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式3") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("處理方式4")) Then
                                If table5.Rows(i)("處理方式4") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式4") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("處理方式5")) Then
                                If table5.Rows(i)("處理方式5") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式5") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("處理方式6")) Then
                                If table5.Rows(i)("處理方式6") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("處理方式6") & ",‎"
                                End If
                            End If
                            If Not IsDBNull(table5.Rows(i)("其他說明")) Then
                                If table5.Rows(i)("其他說明") <> "" Then
                                    建物處理方式 &= table5.Rows(i)("其他說明")
                                End If
                            End If
                            If Right(建物處理方式, 1) = "," Then
                                建物處理方式 = 建物處理方式.Substring(0, 建物處理方式.Length - 1)
                            End If

                            '建物權利種類
                            If Not IsDBNull(table5.Rows(i)("權利種類")) Then
                                建物權利種類 = table5.Rows(i)("權利種類")
                            Else
                                建物權利種類 = ""
                            End If

                            '建物登記日期
                            If Not IsDBNull(table5.Rows(i)("登記日期")) Then
                                建物登記日期 = table5.Rows(i)("登記日期").ToString
                            Else
                                建物登記日期 = ""
                            End If

                            '建物設定性質設定金額
                            If Not IsDBNull(table5.Rows(i)("設定")) Then
                                建物設定性質設定金額 = table5.Rows(i)("設定").ToString & "萬"
                            Else
                                建物設定性質設定金額 = ""
                            End If

                            '建物權利人
                            If Not IsDBNull(table5.Rows(i)("設定權利人")) Then
                                建物權利人 = table5.Rows(i)("設定權利人").ToString
                            Else
                                建物權利人 = ""
                            End If

                            '建物管理人
                            If Not IsDBNull(table5.Rows(i)("管理人")) Then
                                建物管理人 = table5.Rows(i)("管理人").ToString
                            Else
                                建物管理人 = ""
                            End If


                            '先把讀出的xml複製一份，接著開始改
                            Dim tempdata As String = myText_HOwner

                            tempdata = tempdata.Replace("≠建物權利種類", 建物權利種類)
                            tempdata = tempdata.Replace("≠建登記日期", 建物登記日期)
                            tempdata = tempdata.Replace("≠建設定性質設定金額", 建物設定性質設定金額)
                            tempdata = tempdata.Replace("≠建權利人", 建物權利人)
                            tempdata = tempdata.Replace("≠建管理人", 建物管理人)
                            tempdata = tempdata.Replace("≠這裡放處理方式", 建物處理方式)
                            '改完加到最後的字串裡面  
                            最後要替代掉的字串_建物權利人.Append(tempdata)
                        Next
                    End If
                End If
            Else
                If table5.Rows.Count > 0 Then
                    '判斷是否勾選,建物與土地權利部同

                    For i As Integer = 0 To table5.Rows.Count - 1
                        建物處理方式 = ""
                        If Not IsDBNull(table5.Rows(i)("處理方式1")) Then
                            If table5.Rows(i)("處理方式1") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式1") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式2")) Then
                            If table5.Rows(i)("處理方式2") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式2") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式3")) Then
                            If table5.Rows(i)("處理方式3") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式3") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式4")) Then
                            If table5.Rows(i)("處理方式4") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式4") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式5")) Then
                            If table5.Rows(i)("處理方式5") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式5") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("處理方式6")) Then
                            If table5.Rows(i)("處理方式6") <> "" Then
                                建物處理方式 &= table5.Rows(i)("處理方式6") & ",‎"
                            End If
                        End If
                        If Not IsDBNull(table5.Rows(i)("其他說明")) Then
                            If table5.Rows(i)("其他說明") <> "" Then
                                建物處理方式 &= table5.Rows(i)("其他說明")
                            End If
                        End If
                        If Right(建物處理方式, 1) = "," Then
                            建物處理方式 = 建物處理方式.Substring(0, 建物處理方式.Length - 1)
                        End If

                        '建物權利種類
                        If Not IsDBNull(table5.Rows(i)("權利種類")) Then
                            建物權利種類 = table5.Rows(i)("權利種類")
                        Else
                            建物權利種類 = ""
                        End If

                        '建物登記日期
                        If Not IsDBNull(table5.Rows(i)("登記日期")) Then
                            建物登記日期 = table5.Rows(i)("登記日期").ToString
                        Else
                            建物登記日期 = ""
                        End If

                        '建物設定性質設定金額
                        If Not IsDBNull(table5.Rows(i)("設定")) Then
                            建物設定性質設定金額 = table5.Rows(i)("設定").ToString & "萬"
                        Else
                            建物設定性質設定金額 = ""
                        End If

                        '建物權利人
                        If Not IsDBNull(table5.Rows(i)("設定權利人")) Then
                            建物權利人 = table5.Rows(i)("設定權利人").ToString
                        Else
                            建物權利人 = ""
                        End If

                        '建物管理人
                        If Not IsDBNull(table5.Rows(i)("管理人")) Then
                            建物管理人 = table5.Rows(i)("管理人").ToString
                        Else
                            建物管理人 = ""
                        End If


                        '先把讀出的xml複製一份，接著開始改
                        Dim tempdata As String = myText_HOwner

                        tempdata = tempdata.Replace("≠建物權利種類", 建物權利種類)
                        tempdata = tempdata.Replace("≠建登記日期", 建物登記日期)
                        tempdata = tempdata.Replace("≠建設定性質設定金額", 建物設定性質設定金額)
                        tempdata = tempdata.Replace("≠建權利人", 建物權利人)
                        tempdata = tempdata.Replace("≠建管理人", 建物管理人)
                        tempdata = tempdata.Replace("≠這裡放處理方式", 建物處理方式)
                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_建物權利人.Append(tempdata)
                    Next
                Else
                    Dim tempdata As String = myText_HOwner

                    tempdata = tempdata.Replace("≠建物權利種類", "")
                    tempdata = tempdata.Replace("≠建登記日期", "")
                    tempdata = tempdata.Replace("≠建設定性質設定金額", "")
                    tempdata = tempdata.Replace("≠建權利人", "")
                    tempdata = tempdata.Replace("≠建管理人", "")
                    tempdata = tempdata.Replace("≠這裡放處理方式", "")
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_建物權利人.Append(tempdata)
                End If
            End If

            sFile = sFile.Replace("!重複建物權利人", 最後要替代掉的字串_建物權利人.ToString())
            '[END============================================產權調查-建物標示說明P1=============================================END]
        End If  '-------------------------------------------------------------------------判斷成屋還是土地用

        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-建物目前管理狀況P1-P3=============================================START]
            '壹.
            Dim 建號層次說明 As String = ""
            Dim 建物限制登記 As String = ""
            Dim 說明建物限制登記 As String = ""
            Dim 建物信託登記 As String = ""
            Dim 說明建物信託登記 As String = ""
            Dim 建物其他權利 As String = ""
            Dim 說明建物其他權利 As String = ""
            '貳.
            Dim 建物是否共有 As String = ""
            Dim 建物有無分管協議 As String = ""
            Dim 有無分管協議說明 As String = ""
            Dim 有無專有部分範圍 As String = ""
            Dim 專有部分範圍 As String = ""
            Dim 有無共有部分範圍 As String = ""
            Dim 共有部分範圍 As String = ""
            Dim 專有約定共用 As String = ""
            Dim 共有約定專用 As String = ""
            Dim 公共基金數額 As String = ""
            Dim 公共基金提撥方式 As String = ""
            Dim 公共基金運用方式 As String = ""
            Dim 建物管理費 As String = ""
            Dim 管理費繳交方式 As String = ""
            Dim 管理組織及方式 As String = ""
            Dim 有無管理公司 As String = ""
            Dim 管理公司 As String = ""
            Dim 管理手冊 As String = ""
            Dim 說明管理手冊 As String = ""
            Dim 獎勵容積 As String = ""
            Dim 說明獎勵容積 As String = ""
            Dim 電梯設備 As String = ""
            Dim 張貼合格標章 As String = ""
            Dim 說明張貼合格標章 As String = "張貼合格標章"
            Dim 頂樓基地台 As String = "無"
            Dim 基地台說明 As String = ""
            '叁.
            Dim 出租狀況 As String = ""
            Dim s說明出租狀況 As String = ""
            Dim 標題租金 As String = ""
            Dim 標題租期 As String = ""
            Dim 標題租約公證 As String = ""
            Dim 標題說明出租狀況 As String = ""
            Dim 標題租期保證金 As String = ""
            Dim 租期保證金 As String = ""
            Dim 租金 As String = ""
            Dim 租期起 As String = ""
            Dim 租期止 As String = ""
            Dim 租約公證 As String = ""
            Dim 說明出租狀況 As String = ""
            Dim 出借狀況 As String = ""
            Dim 說明出借狀況 As String = ""
            Dim 占用情形 As String = ""
            Dim 標題占用他人建物土地 As String = ""
            Dim 標題被他人占用建物土地 As String = ""
            Dim 標題佔用情形其他 As String = ""
            Dim 占用他人 As String = ""
            Dim 被他人占用 As String = ""
            Dim 佔用情形其他 As String = ""
            Dim 消防設備 As String = ""
            Dim 說明消防設備 As String = ""
            Dim 無障礙設施 As String = ""
            Dim 說明無障礙設施 As String = ""
            Dim 有無夾層 As String = ""
            Dim 夾層面積_0 As String = ""
            Dim 合法登記之夾層_0 As String = ""
            Dim 夾層面積_1 As String = ""
            Dim 合法登記之夾層_1 As String = ""
            Dim 夾層面積_2 As String = ""
            Dim 合法登記之夾層_2 As String = ""
            Dim 獨立供水 As String = ""
            Dim 說明獨立供水 As String = ""
            Dim 獨立電表 As String = ""
            Dim 說明獨立電表 As String = ""
            Dim 天然瓦斯 As String = ""
            Dim 說明天然瓦斯 As String = ""
            Dim 說明2天然瓦斯 As String = ""
            Dim 管線更新 As String = ""
            Dim 管線更新說明 As String = ""
            Dim 積欠費用 As String = ""
            Dim 說明積欠費用 As String = ""
            Dim 屬工業區或其他分區 As String = ""
            Dim 說明屬工業區或其他分區 As String = ""
            Dim 持有期間有無居住 As String = ""
            Dim 使照注意事項 As String = ""
            Dim 說明使照注意事項 As String = ""
            Dim 有無公共設施重大修繕 As String = ""
            Dim 說明有無公共設施重大修繕 As String = ""
            Dim 衛生下水道 As String = ""
            Dim 說明衛生下水道 As String = ""
            Dim 隨附設備 As String = ""
            Dim 說明隨附設備 As String = ""
            '20160625
            Dim 是否有車位 As String = ""
            Dim 說明是否有車位 As String = ""
            '肆
            Dim 氯離子含量 As String = ""
            Dim 氯離子檢測結果 As String = ""
            Dim 檢測日期 As String = ""
            Dim 檢測值 As String = ""
            Dim 有無輻射檢測 As String = ""
            Dim 輻射檢測結果 As String = ""
            Dim 輻射檢測值 As String = ""
            Dim 發生火災 As String = ""
            Dim 標題修繕情形 As String = ""
            Dim 修繕情形 As String = ""
            Dim 發生地震 As String = ""
            Dim 地震等級 As String = ""
            Dim 地震說明 As String = ""
            Dim 樑柱有無裂痕 As String = ""
            Dim 說明樑柱有無裂痕 As String = ""
            Dim 建物鋼筋裸露 As String = ""
            Dim 鋼筋裸露說明 As String = ""
            Dim 有無兇殺情形 As String = ""
            Dim 說明有無兇殺情形 As String = ""
            Dim 非持有有無兇殺情形 As String = ""
            Dim 非持有說明有無兇殺情形 As String = ""
            Dim 有無漏水 As String = ""
            Dim 說明有無漏水 As String = ""
            Dim 有無禁建情事 As String = ""
            Dim 說明有無禁建情事 As String = ""
            Dim 中繼幫浦 As String = ""
            Dim 說明中繼幫浦 As String = ""

            Dim 太陽光電發電設備 As String = "無"
            Dim 說明太陽光電發電設備 As String = ""
            Dim 建築能效標示 As String = "無"
            Dim 說明建築能效標示 As String = ""

            Dim 違增建使用權 As String = ""
            Dim 說明違增建使用權 As String = ""
            Dim 排水系統 As String = ""
            Dim 說明排水系統 As String = ""
            Dim 其他重要事項 As String = ""
            Dim 說明其他重要事項 As String = ""

            '建物目前管理狀況
            Dim myXML_BuildManage As String = ""
            Dim myText_BuildManage As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_BuildManage As New StreamReader(Server.MapPath("..\reports\重複建物管理狀況_V30801.txt"))
            myText_BuildManage = srText_BuildManage.ReadToEnd()
            srText_BuildManage.Close()

            sqlstr = ""
            sqlstr = "Select * From 產調_建物 A  With(NoLock) Left Join 委賣物件資料表_面積細項 B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號  Where A.物件編號 = '" & Contract_No & "' and A.店代號='" & Request("sid") & "'  order by A.流水號 "

            Dim table9 As DataTable

            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table9")
            table9 = ds.Tables("table9")

            Dim 最後要替代掉的字串_建物管理狀況 As New StringBuilder()

            If table9.Rows.Count > 0 Then
                For i = 0 To table9.Rows.Count - 1
                    '建號層次說明
                    建號層次說明 = "<w:br/>‎"
                    If Not IsDBNull(table9.Rows(i)("類別")) Then
                        建號層次說明 &= table9.Rows(i)("類別").ToString
                    End If
                    If Not IsDBNull(table9.Rows(i)("全棟")) Then
                        If table9.Rows(i)("全棟") = "N" Then
                            If Not IsDBNull(table9.Rows(i)("項目名稱")) Then
                                建號層次說明 &= "-" & table9.Rows(i)("項目名稱").ToString
                            End If
                            If Not IsDBNull(table9.Rows(i)("建號")) Then
                                If table9.Rows(i)("建號").ToString.Trim <> "" Then
                                    建號層次說明 &= "(" & table9.Rows(i)("建號").ToString & ")"
                                End If
                            End If
                        End If
                    Else
                        If Not IsDBNull(table9.Rows(i)("項目名稱")) Then
                            建號層次說明 &= "-" & table9.Rows(i)("項目名稱").ToString
                        End If
                        If Not IsDBNull(table9.Rows(i)("建號")) Then
                            If table9.Rows(i)("建號").ToString.Trim <> "" Then
                                建號層次說明 &= "(" & table9.Rows(i)("建號").ToString & ")"
                            End If
                        End If
                    End If

                    '壹
                    '限制登記
                    If Not IsDBNull(table9.Rows(i)("限制登記")) Then
                        建物限制登記 = table9.Rows(i)("限制登記").ToString
                    End If
                    '說明限制登記
                    If Not IsDBNull(table9.Rows(i)("限制登記_說明")) Then
                        說明建物限制登記 = table9.Rows(i)("限制登記_說明").ToString
                    End If
                    '信託登記
                    If Not IsDBNull(table9.Rows(i)("信託登記")) Then
                        建物信託登記 = table9.Rows(i)("信託登記").ToString
                    End If
                    '說明信託登記
                    If Not IsDBNull(table9.Rows(i)("信託登記_說明")) Then
                        說明建物信託登記 = table9.Rows(i)("信託登記_說明").ToString
                    End If
                    '其他權利
                    If Not IsDBNull(table9.Rows(i)("其他權利")) Then
                        建物其他權利 = table9.Rows(i)("其他權利").ToString
                    End If
                    '說明其他權利
                    If Not IsDBNull(table9.Rows(i)("其他權利_說明")) Then
                        說明建物其他權利 = table9.Rows(i)("其他權利_說明").ToString
                    End If

                    '貳
                    '是否共有
                    If Not IsDBNull(table9.Rows(i)("是否共有")) Then
                        建物是否共有 = table9.Rows(i)("是否共有").ToString

                    End If
                    '有無分管協議登記
                    If Not IsDBNull(table9.Rows(i)("有無分管協議登記")) Then
                        建物有無分管協議 = table9.Rows(i)("有無分管協議登記").ToString
                        '有無分管協議說明
                        If table9.Rows(i)("有無分管協議登記_說明").ToString.Length > 0 Then
                            建物有無分管協議 &= "," & table9.Rows(i)("有無分管協議登記_說明").ToString
                        End If
                    End If

                    '有無專有部分範圍
                    If Not IsDBNull(table9.Rows(i)("專有部分範圍_有無")) Then
                        If table9.Rows(i)("專有部分範圍_有無").ToString <> "" Then
                            有無共有部分範圍 = table9.Rows(i)("專有部分範圍_有無").ToString
                        Else
                            有無共有部分範圍 = "無"
                        End If
                    Else
                        有無共有部分範圍 = "無"
                    End If
                    '專有部分範圍
                    If 有無共有部分範圍 <> "無" Then
                        If Not IsDBNull(table9.Rows(i)("專有部分範圍")) Then
                            共有部分範圍 = table9.Rows(i)("專有部分範圍").ToString
                        End If
                    End If

                    '有無共有部分範圍
                    If Not IsDBNull(table9.Rows(i)("共有部分範圍_有無")) Then
                        If table9.Rows(i)("共有部分範圍_有無").ToString <> "" Then
                            有無專有部分範圍 = table9.Rows(i)("共有部分範圍_有無").ToString
                        Else
                            有無專有部分範圍 = "無"
                        End If
                    Else
                        有無專有部分範圍 = "無"
                    End If
                    '共有部分範圍
                    If 有無專有部分範圍 <> "無" Then
                        If Not IsDBNull(table9.Rows(i)("共有部分範圍")) Then
                            專有部分範圍 = table9.Rows(i)("共有部分範圍").ToString
                        End If
                    End If

                    '專有約定共用
                    If Not IsDBNull(table9.Rows(i)("專有約定共用")) Then
                        專有約定共用 = table9.Rows(i)("專有約定共用").ToString

                        If table9.Rows(i)("專有約定共用").ToString = "有" Then
                            If Not IsDBNull(table9.Rows(i)("專有約定共用之範圍")) Then
                                If Trim(table9.Rows(i)("專有約定共用之範圍").ToString) <> "" Then
                                    專有約定共用 &= "，範圍為：" & table9.Rows(i)("專有約定共用之範圍").ToString
                                End If
                            End If


                            If Not IsDBNull(table9.Rows(i)("專有約定共用之使用方式")) Then
                                If Trim(table9.Rows(i)("專有約定共用之使用方式").ToString) <> "" Then
                                    專有約定共用 &= "，使用方式為：" & table9.Rows(i)("專有約定共用之使用方式").ToString
                                End If
                            End If

                        End If
                    Else
                        專有約定共用 = "無"
                    End If
                    '共有約定專用
                    If Not IsDBNull(table9.Rows(i)("共有約定專用")) Then
                        共有約定專用 = table9.Rows(i)("共有約定專用").ToString

                        If table9.Rows(i)("共有約定專用").ToString = "有" Then
                            If Not IsDBNull(table9.Rows(i)("共有約定專用之範圍")) Then
                                If Trim(table9.Rows(i)("共有約定專用之範圍").ToString) <> "" Then
                                    共有約定專用 &= "，範圍為：" & table9.Rows(i)("共有約定專用之範圍").ToString
                                End If
                            End If


                            If Not IsDBNull(table9.Rows(i)("共有約定專用之使用方式")) Then
                                If Trim(table9.Rows(i)("共有約定專用之使用方式").ToString) <> "" Then
                                    共有約定專用 &= "，使用方式為：" & table9.Rows(i)("共有約定專用之使用方式").ToString
                                End If
                            End If

                        End If
                    Else
                        共有約定專用 = "無"
                    End If
                    '公共基金數額
                    If Not IsDBNull(table9.Rows(i)("公共基金數額")) Then
                        公共基金數額 = table9.Rows(i)("公共基金數額").ToString & "萬"
                    End If
                    '公共基金提撥方式
                    If Not IsDBNull(table9.Rows(i)("公共基金提撥方式")) Then
                        公共基金提撥方式 = table9.Rows(i)("公共基金提撥方式").ToString
                    End If
                    '公共基金運用方式
                    If Not IsDBNull(table9.Rows(i)("公共基金運用方式")) Then
                        公共基金運用方式 = table9.Rows(i)("公共基金運用方式").ToString
                    End If

                    '管理費
                    If Not IsDBNull(table9.Rows(0)("管理費使用")) Then
                        If table9.Rows(0)("管理費使用") = "有" Then
                            建物管理費 = table9.Rows(0)("管理費或使用費")
                            If Right(table9.Rows(0)("管理費").ToString, 5) = ".0000" Then
                                建物管理費 &= "," & Int(table9.Rows(0)("管理費").ToString) & "元"
                            Else
                                建物管理費 &= "," & table9.Rows(0)("管理費").ToString & "元"
                            End If
                        Else
                            建物管理費 = "無"
                        End If


                        ''管理費繳交方式-抓委賣物件資料表
                        'If Not IsDBNull(t2.Rows(0)("管理費單位")) Then
                        '    If t2.Rows(0)("管理費單位").ToString <> "" And t2.Rows(0)("管理費單位").ToString <> "無" And t2.Rows(0)("管理費單位").ToString <> "未知" Then
                        '        建物管理費 &= "/" & t2.Rows(0)("管理費單位").ToString
                        '    End If
                        'End If
                    Else
                        建物管理費 = "無"
                    End If

                    '管理費繳交方式
                    If Not IsDBNull(table9.Rows(0)("管理費繳交方式")) Then
                        If table9.Rows(0)("管理費繳交方式").ToString.Trim <> "" Then
                            管理費繳交方式 = table9.Rows(0)("管理費繳交方式").ToString
                        Else
                            管理費繳交方式 = "無"
                        End If
                    Else
                        管理費繳交方式 = "無"
                    End If


                    '管理組織及方式
                    If Not IsDBNull(table9.Rows(i)("管理組織及方式")) Then
                        If Right(table9.Rows(i)("管理組織及方式").ToString, 1) = "," Then
                            管理組織及方式 = Mid(table9.Rows(i)("管理組織及方式").ToString, 1, Len(table9.Rows(i)("管理組織及方式").ToString) - 1)
                        Else
                            管理組織及方式 = table9.Rows(i)("管理組織及方式").ToString
                        End If
                    End If
                    '有無管理公司
                    If Not IsDBNull(table9.Rows(i)("管理公司_有無")) Then
                        If table9.Rows(i)("管理公司_有無").ToString <> "" Then
                            有無管理公司 = table9.Rows(i)("管理公司_有無").ToString
                        Else
                            有無管理公司 = "無"
                        End If
                    Else
                        有無管理公司 = "無"
                    End If
                    '管理公司
                    If Not IsDBNull(table9.Rows(i)("管理公司")) Then
                        管理公司 = table9.Rows(i)("管理公司").ToString
                    End If
                    '管理手冊
                    If Not IsDBNull(table9.Rows(i)("住戶規約使用手冊")) Then
                        管理手冊 = table9.Rows(i)("住戶規約使用手冊").ToString
                    End If
                    '管理手冊說明
                    If Not IsDBNull(table9.Rows(i)("管理手冊_說明")) Then
                        說明管理手冊 = "," & table9.Rows(i)("管理手冊_說明").ToString
                    End If
                    '獎勵容積開放空間提供公共使用
                    If Not IsDBNull(table9.Rows(i)("獎勵容積開放空間提供公共使用")) Then
                        獎勵容積 = table9.Rows(i)("獎勵容積開放空間提供公共使用").ToString
                    End If
                    '獎勵容積開放空間提供公共使用_說明
                    If Not IsDBNull(table9.Rows(i)("獎勵容積開放空間提供公共使用_說明")) Then
                        說明獎勵容積 = table9.Rows(i)("獎勵容積開放空間提供公共使用_說明").ToString
                    End If
                    '電梯設備
                    If Not IsDBNull(table9.Rows(i)("電梯設備")) Then
                        電梯設備 = table9.Rows(i)("電梯設備").ToString
                    End If
                    '張貼合格標章
                    If 電梯設備 = "有" Then
                        If Not IsDBNull(table9.Rows(i)("張貼合格標章")) Then
                            張貼合格標章 = table9.Rows(i)("張貼合格標章").ToString

                            If 張貼合格標章 = "無" Then
                                If Not IsDBNull(table9.Rows(i)("張貼合格標章_說明")) Then
                                    If table9.Rows(i)("張貼合格標章_說明").trim() <> "" Then
                                        張貼合格標章 &= "，" & table9.Rows(i)("張貼合格標章_說明")
                                    End If
                                End If
                            End If
                        End If
                    Else '無時不顯示
                        張貼合格標章 = ""
                        說明張貼合格標章 = ""
                    End If
                    '基地台
                    If Not IsDBNull(table9.Rows(i)("頂樓基地台")) Then
                        If table9.Rows(i)("頂樓基地台").ToString = "有" Then
                            頂樓基地台 = "有"
                            If Not IsDBNull(table9.Rows(i)("頂樓基地台_說明")) Then
                                If table9.Rows(i)("頂樓基地台_說明").trim() <> "" Then
                                    基地台說明 = table9.Rows(i)("頂樓基地台_說明")
                                End If
                            End If

                        Else
                            頂樓基地台 = "無"
                        End If
                    Else
                        頂樓基地台 = "無"
                    End If

                    '叁
                    '出租狀況
                    If Not IsDBNull(table9.Rows(i)("出租狀況")) Then
                        If table9.Rows(i)("出租狀況").ToString <> "" Then

                            If table9.Rows(i)("出租狀況").ToString = "無" Then
                                出租狀況 = "無"
                                標題租金 = ""
                                標題租期 = ""
                                標題租約公證 = ""
                                租金 = ""
                                租期起 = ""
                                租期止 = ""
                                租約公證 = ""
                                標題租期保證金 = ""
                            Else
                                出租狀況 = "有"
                                標題說明出租狀況 = "範圍"
                                If Not IsDBNull(table9.Rows(i)("出租範圍")) Then
                                    If table9.Rows(i)("出租範圍").ToString <> "" Then
                                        標題說明出租狀況 &= ":" & table9.Rows(i)("出租範圍").ToString
                                    End If
                                End If
                                If Not IsDBNull(table9.Rows(i)("押租保證金")) Then
                                    If table9.Rows(i)("押租保證金").ToString <> "" Then
                                        '租期保證金 = CDec(table9.Rows(i)("押租保證金")).ToString("0.0000") & " 萬元"

                                        If Right(table9.Rows(i)("押租保證金").ToString, 5) = ".0000" Then
                                            租期保證金 = Int(table9.Rows(i)("押租保證金").ToString)
                                        ElseIf table9.Rows(i)("押租保證金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("押租保證金").ToString, 3) = "000" Then
                                            租期保證金 = Left(table9.Rows(i)("押租保證金").ToString, Len(table9.Rows(i)("押租保證金").ToString) - 3)
                                        ElseIf table9.Rows(i)("押租保證金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("押租保證金").ToString, 2) = "00" Then
                                            租期保證金 = Left(table9.Rows(i)("押租保證金").ToString, Len(table9.Rows(i)("押租保證金").ToString) - 2)
                                        ElseIf table9.Rows(i)("押租保證金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("押租保證金").ToString, 1) = "0" Then
                                            租期保證金 = Left(table9.Rows(i)("押租保證金").ToString, Len(table9.Rows(i)("押租保證金").ToString) - 1)
                                        Else
                                            租期保證金 = table9.Rows(i)("押租保證金").ToString
                                        End If
                                        租期保證金 &= " 萬元"
                                    End If
                                End If
                                標題租金 = "租金"
                                標題租期 = "租期"
                                標題租約公證 = "租約公證"
                                標題租期保證金 = "租期保證金"
                                '租金
                                If Not IsDBNull(table9.Rows(i)("租金")) Then
                                    If table9.Rows(i)("租金").ToString <> "" Then
                                        '租金 = CDec(table9.Rows(i)("租金")).ToString("0.0000") & " 萬元"

                                        If Right(table9.Rows(i)("租金").ToString, 5) = ".0000" Then
                                            租金 = Int(table9.Rows(i)("租金").ToString)
                                        ElseIf table9.Rows(i)("租金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("租金").ToString, 3) = "000" Then
                                            租金 = Left(table9.Rows(i)("租金").ToString, Len(table9.Rows(i)("租金").ToString) - 3)
                                        ElseIf table9.Rows(i)("租金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("租金").ToString, 2) = "00" Then
                                            租金 = Left(table9.Rows(i)("租金").ToString, Len(table9.Rows(i)("租金").ToString) - 2)
                                        ElseIf table9.Rows(i)("租金").ToString.LastIndexOf(".") > 0 And Right(table9.Rows(i)("租金").ToString, 1) = "0" Then
                                            租金 = Left(table9.Rows(i)("租金").ToString, Len(table9.Rows(i)("租金").ToString) - 1)
                                        Else
                                            租金 = table9.Rows(i)("租金").ToString
                                        End If
                                        租金 &= " 萬元"

                                        'If Right(table9.Rows(0)("租金").ToString, 5) = ".0000" Then
                                        '    租金 = Int(table9.Rows(0)("租金").ToString) & " 萬元"
                                        'Else

                                        '    租金 = String.Format("{0:#.##}", CDec(table9.Rows(0)("租金"))) & " 萬元"
                                        'End If
                                    End If
                                End If




                                If Not IsDBNull(table9.Rows(i)("出租情況類型")) And table9.Rows(i)("出租情況類型").ToString.Trim <> "" Then
                                    'If table9.Rows(i)("出租情況類型").ToString <> "" Then
                                    '    標題租期 &= "(" & table9.Rows(i)("出租情況類型").ToString & ")"

                                    'End If
                                    Dim 出租情況類型 As Array = Split(table9.Rows(i)("出租情況類型").ToString, ";")
                                    For j = 0 To 出租情況類型.Length - 1
                                        If 出租情況類型(j).ToString = "不定期租約" Or 出租情況類型(j).ToString = "定期租約" Then
                                            標題租期 &= "(" & 出租情況類型(j).ToString & ")"
                                        ElseIf 出租情況類型(j).ToString = "租賃之權利義務隨同移轉" Or 出租情況類型(j).ToString = "屋主終止租約騰空交屋" Then
                                            s說明出租狀況 = "說明:" & 出租情況類型(j).ToString
                                        ElseIf 出租情況類型(j).ToString = "其他" Then
                                            s說明出租狀況 = "說明:" & table9.Rows(i)("出租狀況_說明").ToString
                                        End If
                                    Next

                                End If
                                '租期
                                If Not IsDBNull(table9.Rows(i)("租期起")) Then
                                    If table9.Rows(i)("租期起").ToString <> "" Then
                                        租期起 = table9.Rows(i)("租期起").ToString & "起"
                                    End If
                                End If
                                If Not IsDBNull(table9.Rows(i)("租期迄")) Then
                                    If table9.Rows(i)("租期迄").ToString <> "" Then
                                        租期止 = table9.Rows(i)("租期迄").ToString & "止"
                                    End If
                                End If
                                '租約公證
                                If Not IsDBNull(table9.Rows(i)("租約是否公證")) Then
                                    If table9.Rows(i)("租約是否公證").ToString <> "" Then
                                        租約公證 = table9.Rows(i)("租約是否公證").ToString
                                    End If
                                End If


                            End If
                        Else
                            出租狀況 = "無"
                        End If
                    Else
                        出租狀況 = "無"
                    End If



                    '出租狀況_說明
                    If Not IsDBNull(table9.Rows(i)("出租狀況_說明")) Then
                        If table9.Rows(i)("出租狀況_說明").ToString <> "" Then
                            '標題說明出租狀況 = "說明"
                            說明出租狀況 = table9.Rows(i)("出租狀況_說明").ToString
                        End If
                    End If


                    '出借狀況
                    If Not IsDBNull(table9.Rows(i)("出借狀況")) Then
                        If table9.Rows(i)("出借狀況").ToString <> "" Then
                            出借狀況 = table9.Rows(i)("出借狀況").ToString

                            If table9.Rows(i)("出借狀況").ToString <> "無" Then
                                說明出借狀況 = "範圍:" & table9.Rows(i)("出借範圍") & table9.Rows(i)("出借範圍備註") & ","
                                '出借書面約定與否
                                If Not IsDBNull(table9.Rows(i)("出借書面約定與否")) Then
                                    If table9.Rows(i)("出借書面約定與否").ToString <> "" Then
                                        說明出借狀況 &= table9.Rows(i)("出借書面約定與否").ToString
                                    End If
                                End If
                                '出借狀況_說明
                                If Not IsDBNull(table9.Rows(i)("出借狀況_說明")) Then
                                    If table9.Rows(i)("出借狀況_說明").ToString <> "" Then
                                        說明出借狀況 &= table9.Rows(i)("出借狀況_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            出借狀況 = "無"
                        End If
                    Else
                        出借狀況 = "無"
                    End If

                    '占用情形
                    If Not IsDBNull(table9.Rows(i)("佔用情形")) Then
                        If table9.Rows(i)("佔用情形").ToString <> "" Then
                            占用情形 = table9.Rows(i)("佔用情形").ToString

                            If table9.Rows(i)("佔用情形").ToString = "無" Then
                                標題占用他人建物土地 = ""
                                標題被他人占用建物土地 = ""
                                標題佔用情形其他 = ""

                                占用他人 = ""
                                被他人占用 = ""
                                佔用情形其他 = ""
                            Else
                                標題占用他人建物土地 = "有無占用他人土地"
                                標題被他人占用建物土地 = "有無被他人占用建物"
                                標題佔用情形其他 = "有無其他占用"

                                '占用他人建物土地
                                If Not IsDBNull(table9.Rows(i)("佔用他人建物土地")) Then
                                    If table9.Rows(i)("佔用他人建物土地").ToString <> "" Then
                                        占用他人 = table9.Rows(i)("佔用他人建物土地").ToString
                                    End If
                                    If table9.Rows(i)("佔用他人建物土地_說明").ToString <> "" Then
                                        占用他人 &= "," & table9.Rows(i)("佔用他人建物土地_說明").ToString
                                    End If
                                End If
                                '被他人占用建物土地
                                If Not IsDBNull(table9.Rows(i)("被他人佔用建物土地")) Then
                                    If table9.Rows(i)("被他人佔用建物土地").ToString <> "" Then
                                        被他人占用 = table9.Rows(i)("被他人佔用建物土地").ToString
                                    End If
                                    If table9.Rows(i)("被他人佔用建物土地_說明").ToString <> "" Then
                                        被他人占用 &= "," & table9.Rows(i)("被他人佔用建物土地_說明").ToString
                                    End If
                                End If

                                '佔用情形其他
                                If Not IsDBNull(table9.Rows(i)("佔用情形其他")) Then
                                    If table9.Rows(i)("佔用情形其他").ToString <> "" Then
                                        佔用情形其他 = table9.Rows(i)("佔用情形其他").ToString
                                    End If
                                    If table9.Rows(i)("佔用情形其他_說明").ToString <> "" Then
                                        佔用情形其他 &= "," & table9.Rows(i)("佔用情形其他_說明").ToString
                                    End If
                                End If
                            End If
                        Else
                            占用情形 = "無"
                        End If
                    Else
                        占用情形 = "無"
                    End If
                    '消防設備
                    If Not IsDBNull(table9.Rows(i)("消防設備")) Then
                        If table9.Rows(i)("消防設備").ToString <> "" Then
                            消防設備 = table9.Rows(i)("消防設備").ToString

                            If table9.Rows(i)("消防設備").ToString <> "無" Then
                                '消防設備_說明
                                If Not IsDBNull(table9.Rows(i)("消防設備_說明")) Then
                                    If table9.Rows(i)("消防設備_說明").ToString <> "" Then
                                        說明消防設備 = table9.Rows(i)("消防設備_說明").ToString.Replace(",,", "")
                                    End If
                                End If

                            End If
                        Else
                            消防設備 = "無"
                        End If
                    Else
                        消防設備 = "無"
                    End If
                    '無障礙設施
                    If Not IsDBNull(table9.Rows(i)("無障礙設施")) Then
                        If table9.Rows(i)("無障礙設施").ToString <> "" Then
                            無障礙設施 = table9.Rows(i)("無障礙設施").ToString

                            If table9.Rows(i)("無障礙設施").ToString <> "無" Then
                                '無障礙設施_說明
                                If Not IsDBNull(table9.Rows(i)("無障礙設施_說明")) Then
                                    If table9.Rows(i)("無障礙設施_說明").ToString <> "" Then
                                        說明無障礙設施 = table9.Rows(i)("無障礙設施_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            無障礙設施 = "無"
                        End If
                    Else
                        無障礙設施 = "無"
                    End If
                    '夾層
                    If Not IsDBNull(table9.Rows(i)("夾層")) Then
                        If table9.Rows(i)("夾層").ToString <> "" Then
                            有無夾層 = table9.Rows(i)("夾層").ToString

                            If table9.Rows(i)("夾層").ToString <> "無" Then
                                '夾層面積
                                If Not IsDBNull(table9.Rows(i)("夾層面積")) Then
                                    If Right(table9.Rows(i)("夾層面積").ToString, 5) = ".0000" Then
                                        夾層面積_0 = "約 " & Int(table9.Rows(i)("夾層面積").ToString) & "  m²"
                                    Else
                                        夾層面積_0 = "約 " & table9.Rows(i)("夾層面積").ToString & "  m²"
                                    End If
                                End If
                                ''合法登記之夾層
                                合法登記之夾層_0 = "權狀面積"
                                'If Not IsDBNull(table9.Rows(i)("是否合法登記之夾層")) Then
                                '    If table9.Rows(i)("是否合法登記之夾層").ToString = "有" Then
                                '        合法登記之夾層_0 = "有建物所有權狀"
                                '    ElseIf table9.Rows(i)("是否合法登記之夾層").ToString = "無" Then
                                '        合法登記之夾層_0 = "無建物所有權狀"
                                '    Else
                                '        合法登記之夾層_0 = "其他夾層"
                                '    End If
                                'End If

                                '夾層面積1
                                If Not IsDBNull(table9.Rows(i)("夾層面積1")) Then
                                    If Right(table9.Rows(i)("夾層面積1").ToString, 5) = ".0000" Then
                                        夾層面積_1 = "約 " & Int(table9.Rows(i)("夾層面積1").ToString) & " m²"
                                    Else
                                        夾層面積_1 = "約 " & table9.Rows(i)("夾層面積1").ToString & " m²"
                                    End If
                                End If
                                ''合法登記之夾層1
                                合法登記之夾層_1 = "無所有權狀面積"
                                'If Not IsDBNull(table9.Rows(i)("是否合法登記之夾層1")) Then
                                '    If table9.Rows(i)("是否合法登記之夾層1").ToString = "有" Then
                                '        合法登記之夾層_1 = "有建物所有權狀"
                                '    ElseIf table9.Rows(i)("是否合法登記之夾層1").ToString = "無" Then
                                '        合法登記之夾層_1 = "無建物所有權狀"
                                '    Else
                                '        合法登記之夾層_1 = "其他夾層"
                                '    End If
                                'End If
                                '夾層面積1
                                If Not IsDBNull(table9.Rows(i)("夾層面積2")) Then
                                    If Right(table9.Rows(i)("夾層面積2").ToString, 5) = ".0000" Then
                                        夾層面積_2 = "約 " & Int(table9.Rows(i)("夾層面積2").ToString) & " m²"
                                    Else
                                        夾層面積_2 = "約 " & table9.Rows(i)("夾層面積2").ToString & " m²"
                                    End If
                                End If
                                ''合法登記之夾層2
                                合法登記之夾層_2 = "其他面積"
                                If Not IsDBNull(table9.Rows(i)("夾層其他")) Then
                                    夾層面積_2 = "約 " & table9.Rows(i)("夾層其他").ToString & " m²"
                                End If
                                'If Not IsDBNull(table9.Rows(i)("是否合法登記之夾層2")) Then
                                '    If table9.Rows(i)("是否合法登記之夾層2").ToString = "有" Then
                                '        合法登記之夾層_2 = "有建物所有權狀"
                                '    ElseIf table9.Rows(i)("是否合法登記之夾層2").ToString = "無" Then
                                '        合法登記之夾層_2 = "無建物所有權狀"
                                '    Else
                                '        合法登記之夾層_2 = "其他夾層"
                                '    End If
                                'End If
                            End If
                        Else
                            有無夾層 = "無"
                        End If
                    Else
                        有無夾層 = "無"
                    End If
                    '獨立供水
                    If Not IsDBNull(table9.Rows(i)("獨立供水")) Then
                        If table9.Rows(i)("獨立供水").ToString <> "" Then
                            獨立供水 = table9.Rows(i)("獨立供水").ToString

                            If table9.Rows(i)("獨立供水").ToString <> "無" Then


                                '供水類型
                                If Not IsDBNull(table9.Rows(i)("供水類型")) Then
                                    If table9.Rows(i)("供水類型").ToString <> "" Then
                                        說明獨立供水 = table9.Rows(i)("供水類型").ToString
                                    End If
                                End If
                                '供水是否正常
                                If Not IsDBNull(table9.Rows(i)("供水是否正常")) Then
                                    If table9.Rows(i)("供水是否正常").ToString = "是" Then
                                        說明獨立供水 &= ",供水正常"
                                    End If
                                End If
                                '獨立供水_說明
                                If Not IsDBNull(table9.Rows(i)("獨立供水_說明")) Then
                                    If table9.Rows(i)("獨立供水_說明").ToString <> "" Then

                                        說明獨立供水 &= "," & table9.Rows(i)("獨立供水_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            獨立供水 = "無"
                        End If
                    Else
                        獨立供水 = "無"
                    End If
                    '獨立電表
                    If Not IsDBNull(table9.Rows(i)("獨立電表")) Then
                        If table9.Rows(i)("獨立電表").ToString <> "" Then
                            獨立電表 = table9.Rows(i)("獨立電表").ToString
                        Else
                            獨立電表 = "無"
                        End If
                    Else
                        獨立電表 = "無"
                    End If
                    '獨立電表_說明
                    If Not IsDBNull(table9.Rows(i)("獨立電表_說明")) Then
                        If table9.Rows(i)("獨立電表_說明").ToString <> "" Then
                            說明獨立電表 = table9.Rows(i)("獨立電表_說明").ToString
                        Else
                            說明獨立電表 = ""
                        End If
                    End If

                    '天然瓦斯
                    If Not IsDBNull(table9.Rows(i)("天然瓦斯")) Then
                        If table9.Rows(i)("天然瓦斯").ToString <> "" Then
                            天然瓦斯 = table9.Rows(i)("天然瓦斯").ToString
                        Else
                            天然瓦斯 = "無"
                        End If
                    Else
                        天然瓦斯 = "無"
                    End If
                    '天然瓦斯_說明
                    If 天然瓦斯 = "有" Then
                        If Not IsDBNull(table9.Rows(i)("天然瓦斯_說明")) Then
                            If table9.Rows(i)("天然瓦斯_說明").ToString <> "" Then
                                說明天然瓦斯 = table9.Rows(i)("天然瓦斯_說明").ToString
                            End If
                        End If
                        If Not IsDBNull(table9.Rows(i)("天然瓦斯_說明2")) Then
                            說明2天然瓦斯 = table9.Rows(i)("天然瓦斯_說明2").ToString
                        End If
                    End If


                    '管線更新說明
                    管線更新說明 = table9.Rows(i)("水電管線是否更新_說明").ToString
                    '管線更新
                    If Not IsDBNull(table9.Rows(i)("水電管線是否更新")) Then
                        If table9.Rows(i)("水電管線是否更新").ToString <> "" Then
                            管線更新 = table9.Rows(i)("水電管線是否更新").ToString
                            If table9.Rows(i)("水電管線是否更新").ToString = "有" Then
                                If Not IsDBNull(table9.Rows(i)("水管更新日期")) Then
                                    管線更新說明 &= ",水管更新日:" & table9.Rows(i)("水管更新日期")
                                End If
                                If Not IsDBNull(table9.Rows(i)("電線更新日期")) Then
                                    管線更新說明 &= ",電線更新日:" & table9.Rows(i)("電線更新日期")
                                End If
                            End If
                        Else
                            管線更新 = "無"
                        End If
                    Else
                        管線更新 = "無"
                    End If


                    '積欠費用
                    If Not IsDBNull(table9.Rows(i)("積欠應繳費用")) Then
                        If table9.Rows(i)("積欠應繳費用").ToString <> "" Then
                            積欠費用 = table9.Rows(i)("積欠應繳費用").ToString

                            If table9.Rows(i)("積欠應繳費用").ToString <> "無" Then
                                '積欠費用_說明
                                If Not IsDBNull(table9.Rows(i)("積欠應繳費用_說明")) Then
                                    If table9.Rows(i)("積欠應繳費用_說明").ToString <> "" Then
                                        說明積欠費用 = table9.Rows(i)("積欠應繳費用_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            積欠費用 = "無"
                        End If
                    Else
                        積欠費用 = "無"
                    End If
                    '屬工業區或其他分區
                    If Not IsDBNull(table9.Rows(i)("屬工業區或其他分區")) Then
                        If table9.Rows(i)("屬工業區或其他分區").ToString <> "" Then
                            屬工業區或其他分區 = table9.Rows(i)("屬工業區或其他分區").ToString

                            If table9.Rows(i)("屬工業區或其他分區").ToString <> "無" Then
                                '屬工業區或其他分區_說明
                                If Not IsDBNull(table9.Rows(i)("屬工業區或其他分區_說明")) Then
                                    If table9.Rows(i)("屬工業區或其他分區_說明").ToString <> "" Then
                                        說明屬工業區或其他分區 = table9.Rows(i)("屬工業區或其他分區_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            屬工業區或其他分區 = "無"
                        End If
                    Else
                        屬工業區或其他分區 = "無"
                    End If
                    '持有期間有無居住
                    If Not IsDBNull(table9.Rows(i)("持有期間有無居住")) Then
                        If table9.Rows(i)("持有期間有無居住").ToString <> "" Then
                            持有期間有無居住 = table9.Rows(i)("持有期間有無居住").ToString
                        Else
                            持有期間有無居住 = "無"
                        End If
                    Else
                        持有期間有無居住 = "無"
                    End If
                    '使照注意事項
                    If Not IsDBNull(table9.Rows(i)("使用執照有無備註之注意事項")) Then
                        If table9.Rows(i)("使用執照有無備註之注意事項").ToString <> "" Then
                            使照注意事項 = table9.Rows(i)("使用執照有無備註之注意事項").ToString

                            If table9.Rows(i)("使用執照有無備註之注意事項").ToString <> "無" Then
                                '使照注意事項_說明
                                If Not IsDBNull(table9.Rows(i)("使用執照有無備註之注意事項_說明")) Then
                                    If table9.Rows(i)("使用執照有無備註之注意事項_說明").ToString <> "" Then
                                        說明使照注意事項 = table9.Rows(i)("使用執照有無備註之注意事項_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            使照注意事項 = "無"
                        End If
                    Else
                        使照注意事項 = "無"
                    End If
                    '有無公共設施重大修繕
                    If Not IsDBNull(table9.Rows(i)("有無公共設施重大修繕")) Then
                        If table9.Rows(i)("有無公共設施重大修繕").ToString <> "" Then
                            有無公共設施重大修繕 = table9.Rows(i)("有無公共設施重大修繕").ToString

                            If table9.Rows(i)("有無公共設施重大修繕").ToString <> "無" Then
                                '說明有無公共設施重大修繕_說明
                                If Not IsDBNull(table9.Rows(i)("有無公共設施重大修繕_說明")) Then
                                    If table9.Rows(i)("有無公共設施重大修繕_說明").ToString <> "" Then
                                        說明有無公共設施重大修繕 = table9.Rows(i)("有無公共設施重大修繕_說明").ToString
                                    End If
                                    If table9.Rows(i)("有無公共設施重大修繕_金額").ToString <> "" Then
                                        說明有無公共設施重大修繕 &= ",修繕金額:" & table9.Rows(i)("有無公共設施重大修繕_金額").ToString
                                    End If
                                End If

                            End If
                        Else
                            有無公共設施重大修繕 = "無"
                        End If
                    Else
                        有無公共設施重大修繕 = "無"
                    End If
                    '衛生下水道工程
                    If Not IsDBNull(table9.Rows(i)("衛生下水道工程")) Then
                        If table9.Rows(i)("衛生下水道工程").ToString <> "" Then
                            衛生下水道 = table9.Rows(i)("衛生下水道工程").ToString

                            If table9.Rows(i)("衛生下水道工程").ToString = "無" Then
                                '衛生下水道工程_選項
                                If Not IsDBNull(table9.Rows(i)("衛生下水道工程_選項")) Then
                                    If table9.Rows(i)("衛生下水道工程_選項").ToString <> "" Then
                                        說明衛生下水道 = table9.Rows(i)("衛生下水道工程_選項").ToString
                                    End If
                                End If
                                '說明衛生下水道
                                If Not IsDBNull(table9.Rows(i)("衛生下水道工程_說明")) Then
                                    If table9.Rows(i)("衛生下水道工程_說明").ToString <> "" Then
                                        說明衛生下水道 &= "," & table9.Rows(i)("衛生下水道工程_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            衛生下水道 = table9.Rows(i)("衛生下水道工程").ToString
                        End If
                    Else
                        衛生下水道 = table9.Rows(i)("衛生下水道工程").ToString
                    End If
                    '附屬物部分
                    If Not IsDBNull(table9.Rows(i)("隨附設備有無")) Then
                        If table9.Rows(i)("隨附設備有無").ToString <> "" Then
                            隨附設備 = table9.Rows(i)("隨附設備有無").ToString

                            If table9.Rows(i)("隨附設備有無").ToString.Trim = "有" Then
                                '隨附設備
                                If Not IsDBNull(table9.Rows(i)("隨附設備")) Then
                                    If table9.Rows(i)("隨附設備").ToString <> "" And table9.Rows(i)("隨附設備").ToString.Length > 1 Then
                                        '先將設備轉成 array 固定物現況;燈飾;冰箱;冷氣機;
                                        Dim equAr() As String = table9.Rows(i)("隨附設備").ToString.Split(";")
                                        Dim ShowStr As String = ""
                                        '如果是,[沙發數],[電視數] ,[冰箱數],[冷氣數],[洗衣機數],[乾衣機數] 要加入數目
                                        For Each equstr As String In equAr
                                            If equstr = "沙發" Then
                                                equstr = "沙發" & IIf(IsDBNull(table9.Rows(i)("沙發數")), "0", table9.Rows(i)("沙發數")) & "組"
                                            ElseIf equstr = "電視機" Then
                                                equstr = "電視機" & IIf(IsDBNull(table9.Rows(i)("電視數")), "0", table9.Rows(i)("電視數")) & "台"
                                            ElseIf equstr = "冰箱" Then
                                                equstr = "冰箱" & IIf(IsDBNull(table9.Rows(i)("冰箱數")), "0", table9.Rows(i)("冰箱數")) & "台"
                                            ElseIf equstr = "冷氣機" Then
                                                equstr = "冷氣機" & IIf(IsDBNull(table9.Rows(i)("冷氣數")), "0", table9.Rows(i)("冷氣數")) & "台"
                                            ElseIf equstr = "洗衣機" Then
                                                equstr = "洗衣機" & IIf(IsDBNull(table9.Rows(i)("洗衣機數")), "0", table9.Rows(i)("洗衣機數")) & "台"
                                            ElseIf equstr = "乾衣機" Then
                                                equstr = "乾衣機" & IIf(IsDBNull(table9.Rows(i)("乾衣機數")), "0", table9.Rows(i)("乾衣機數")) & "台"
                                            End If
                                            ShowStr += equstr & ";"
                                        Next

                                        說明隨附設備 = ShowStr
                                    End If
                                End If


                            End If
                        Else
                            隨附設備 = "無"
                        End If
                    Else
                        隨附設備 = "無"
                    End If

                    '是否有車位

                    If table9.Rows(i)("類別").ToString.IndexOf("車位") >= 0 Or table9.Rows(i)("項目名稱").ToString.IndexOf("車位") >= 0 Then
                        是否有車位 = "有"
                        說明是否有車位 = "有車位，請參閱車位產權調查項目"
                    Else
                        是否有車位 = "無"
                        說明是否有車位 = "無車位"
                    End If
                    sqlstr = "select 車位 from 委賣物件資料表 where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "'"
                    Using conn_checkcarpark As New SqlConnection(EGOUPLOADSqlConnStr)
                        conn_checkcarpark.Open()
                        Using cmd As New SqlCommand(sqlstr, conn)
                            Dim dt As New DataTable
                            dt.Load(cmd.ExecuteReader)
                            If dt.Rows.Count > 0 Then
                                If dt.Rows(0)(0) <> "" And dt.Rows(0)(0) <> "無" And dt.Rows(0)(0) <> "請選擇" Then
                                    是否有車位 = "有"
                                    說明是否有車位 = "有車位，請參閱車位產權調查項目"
                                End If
                            End If
                        End Using
                    End Using


                    '肆
                    '氯離子含量
                    氯離子含量 = table9.Rows(i)("混凝土中氯離子含量").ToString
                    氯離子檢測結果 = table9.Rows(i)("混凝土中氯離子含量_說明").ToString
                    If 氯離子含量 = "有" Then
                        氯離子檢測結果 = "(附檢測結果)"
                    End If

                    '有無輻射檢測
                    有無輻射檢測 = table9.Rows(i)("輻射檢測").ToString
                    輻射檢測結果 = table9.Rows(i)("輻射檢測_說明").ToString
                    If 有無輻射檢測 = "有" Then
                        輻射檢測結果 = "(附檢測結果)"
                    End If

                    'If Not IsDBNull(table9.Rows(i)("混凝土中氯離子含量")) Then
                    '    If table9.Rows(i)("混凝土中氯離子含量").ToString <> "" Then
                    '        有無輻射檢測 = table9.Rows(i)("混凝土中氯離子含量").ToString

                    '        If table9.Rows(i)("混凝土中氯離子含量").ToString = "有" Then
                    '            輻射檢測結果 = "詳如檢測報告"
                    '            輻射檢測值 = ""
                    '        Else
                    '            輻射檢測結果 = ""


                    '            '輻射檢測值-輻射檢測_說明
                    '            If Not IsDBNull(table9.Rows(i)("輻射檢測_說明")) Then
                    '                If table9.Rows(i)("輻射檢測_說明").ToString <> "" Then
                    '                    輻射檢測值 = table9.Rows(0)("輻射檢測_說明").ToString
                    '                End If
                    '            End If

                    '        End If
                    '    Else
                    '        有無輻射檢測 = "有"
                    '    End If
                    'Else
                    '    有無輻射檢測 = "有"
                    'End If
                    '發生火災
                    If Not IsDBNull(table9.Rows(i)("曾發生火災或其他災害")) Then
                        If table9.Rows(i)("曾發生火災或其他災害").ToString <> "" Then
                            發生火災 = table9.Rows(i)("曾發生火災或其他災害").ToString

                            If table9.Rows(i)("曾發生火災或其他災害").ToString = "無" Then
                                標題修繕情形 = ""
                                修繕情形 = ""
                            Else
                                標題修繕情形 = "修繕情形"


                                '修繕情形
                                If Not IsDBNull(table9.Rows(i)("曾發生火災或其他災害_說明")) Then
                                    If table9.Rows(i)("曾發生火災或其他災害_說明").ToString <> "" Then
                                        修繕情形 = table9.Rows(i)("曾發生火災或其他災害_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            發生火災 = "無"
                        End If
                    Else
                        發生火災 = "無"
                    End If
                    '發生地震
                    If Not IsDBNull(table9.Rows(i)("因地震被公告為危險建築")) Then
                        If table9.Rows(i)("因地震被公告為危險建築").ToString <> "" Then
                            發生地震 = table9.Rows(i)("因地震被公告為危險建築").ToString

                            If table9.Rows(i)("因地震被公告為危險建築").ToString = "無" Then
                                地震等級 = ""
                                地震說明 = ""
                            Else
                                地震等級 = ""

                                '地震說明
                                If Not IsDBNull(table9.Rows(i)("因地震被公告為危險建築_說明")) Then
                                    If table9.Rows(i)("因地震被公告為危險建築_說明").ToString <> "" Then
                                        地震說明 = table9.Rows(i)("因地震被公告為危險建築_說明").ToString
                                    End If
                                End If

                            End If
                        Else
                            發生地震 = "無"
                        End If
                    Else
                        發生地震 = "無"
                    End If
                    '樑柱有無裂痕
                    If Not IsDBNull(table9.Rows(i)("樑柱部分是否有顯見裂痕")) Then
                        If table9.Rows(i)("樑柱部分是否有顯見裂痕").ToString <> "" Then
                            樑柱有無裂痕 = table9.Rows(i)("樑柱部分是否有顯見裂痕").ToString

                            If table9.Rows(i)("樑柱部分是否有顯見裂痕").ToString <> "無" Then
                                If Not IsDBNull(table9.Rows(i)("樑柱部分是否有顯見裂痕_說明")) Then
                                    說明樑柱有無裂痕 = table9.Rows(i)("樑柱部分是否有顯見裂痕_說明")
                                End If
                                If Not IsDBNull(table9.Rows(i)("裂痕長度")) Then
                                    說明樑柱有無裂痕 &= " ,裂痕長度:" & table9.Rows(i)("裂痕長度")
                                End If
                                If Not IsDBNull(table9.Rows(i)("間隙寬度")) Then
                                    說明樑柱有無裂痕 &= " ,間隙寬度:" & table9.Rows(i)("間隙寬度")
                                End If
                            End If
                        Else
                            樑柱有無裂痕 = "無"
                        End If
                    Else
                        樑柱有無裂痕 = "無"
                    End If

                    '建物鋼筋裸露
                    If Not IsDBNull(table9.Rows(i)("建物鋼筋裸露")) Then
                        If table9.Rows(i)("建物鋼筋裸露").ToString <> "" Then
                            建物鋼筋裸露 = table9.Rows(i)("建物鋼筋裸露").ToString

                            If table9.Rows(i)("建物鋼筋裸露").ToString <> "無" Then
                                '鋼筋裸露說明
                                If Not IsDBNull(table9.Rows(i)("建物鋼筋裸露_說明")) Then
                                    If table9.Rows(i)("建物鋼筋裸露_說明").ToString <> "" Then
                                        鋼筋裸露說明 = "其位置:" & table9.Rows(i)("建物鋼筋裸露_說明").ToString
                                    End If
                                End If
                            End If
                        Else
                            建物鋼筋裸露 = "無"
                        End If
                    Else
                        建物鋼筋裸露 = "無"
                    End If
                    '----------------------------- 下面兩個要整合
                    '有無兇殺情形
                    If Not IsDBNull(table9.Rows(i)("是否為兇宅")) Then
                        If table9.Rows(i)("是否為兇宅").ToString <> "" Then
                            有無兇殺情形 = table9.Rows(i)("是否為兇宅").ToString
                            '發生期間
                            If 有無兇殺情形 = "有" And table9.Rows(i)("兇宅發生期間").ToString.Length > 0 Then
                                說明有無兇殺情形 = "發生期間:" & table9.Rows(i)("兇宅發生期間").ToString
                            End If
                        Else
                            有無兇殺情形 = "無"
                        End If
                    Else
                        有無兇殺情形 = "無"
                    End If

                    '說明有無兇殺情形
                    If Not IsDBNull(table9.Rows(i)("是否為兇宅_說明")) Then

                        If table9.Rows(i)("是否為兇宅_說明").ToString <> "" Then
                            說明有無兇殺情形 &= ",說明:" & table9.Rows(i)("是否為兇宅_說明").ToString
                        End If
                    End If

                    ''非持有有無兇殺情形
                    'If Not IsDBNull(table9.Rows(i)("非持有是否為兇宅")) Then
                    '    If table9.Rows(i)("非持有是否為兇宅").ToString <> "" Then
                    '        非持有有無兇殺情形 = table9.Rows(i)("非持有是否為兇宅").ToString
                    '    Else
                    '        非持有有無兇殺情形 = "無"
                    '    End If
                    'Else
                    '    非持有有無兇殺情形 = "無"
                    'End If
                    ''非持有說明有無兇殺情形
                    'If Not IsDBNull(table9.Rows(i)("非持有是否為兇宅_說明")) Then
                    '    If table9.Rows(i)("非持有是否為兇宅_說明").ToString <> "" Then
                    '        非持有說明有無兇殺情形 = table9.Rows(i)("非持有是否為兇宅_說明").ToString
                    '    End If
                    'End If
                    '-----------------------
                    '有無漏水
                    If Not IsDBNull(table9.Rows(i)("滲漏水狀態")) Then
                        If table9.Rows(i)("滲漏水狀態").ToString <> "" Then
                            有無漏水 = table9.Rows(i)("滲漏水狀態").ToString

                            If table9.Rows(i)("滲漏水狀態").ToString <> "無" Then
                                '說明有無漏水
                                If Not IsDBNull(table9.Rows(i)("滲漏水狀態_說明")) Then
                                    If table9.Rows(i)("滲漏水狀態_說明").ToString <> "" Then
                                        說明有無漏水 = table9.Rows(i)("滲漏水狀態_說明").ToString
                                    End If
                                    If table9.Rows(i)("滲漏水狀態_處理").ToString <> "" Then
                                        說明有無漏水 &= " ,處理:" & table9.Rows(i)("滲漏水狀態_處理").ToString
                                    End If
                                End If

                            End If
                        Else
                            有無漏水 = "無"
                        End If
                    Else
                        有無漏水 = "無"
                    End If
                    '有無禁建情事
                    If Not IsDBNull(table9.Rows(i)("有無禁建情事")) Then
                        If table9.Rows(i)("有無禁建情事").ToString <> "" Then
                            有無禁建情事 = table9.Rows(i)("有無禁建情事").ToString

                            If table9.Rows(i)("有無禁建情事").ToString <> "無" Then
                                '說明有無禁建情事
                                If Not IsDBNull(table9.Rows(i)("有無禁建情事_說明")) Then
                                    If table9.Rows(i)("有無禁建情事_說明").ToString <> "" Then
                                        說明有無禁建情事 = table9.Rows(i)("有無禁建情事_說明").ToString
                                    End If
                                End If
                            End If
                        Else
                            有無禁建情事 = "無"
                        End If
                    Else
                        有無禁建情事 = "無"
                    End If

                    If Not IsDBNull(table9.Rows(i)("中繼幫浦")) Then
                        If table9.Rows(i)("中繼幫浦").ToString <> "" Then
                            中繼幫浦 = table9.Rows(i)("中繼幫浦").ToString

                            If table9.Rows(i)("中繼幫浦").ToString <> "無" Then
                                '說明中繼幫浦
                                If Not IsDBNull(table9.Rows(i)("中繼幫浦_說明")) Then
                                    If table9.Rows(i)("中繼幫浦_說明").ToString <> "" Then
                                        說明中繼幫浦 = table9.Rows(i)("中繼幫浦_說明").ToString
                                    End If
                                End If
                            End If
                        Else
                            中繼幫浦 = "無"
                        End If
                    Else
                        中繼幫浦 = "無"
                    End If

                    If Not IsDBNull(table9.Rows(i)("太陽光電發電設備")) AndAlso table9.Rows(i)("太陽光電發電設備").ToString.Trim <> "" Then
                        太陽光電發電設備 = table9.Rows(i)("太陽光電發電設備").ToString.Trim
                    Else
                        太陽光電發電設備 = "無"
                    End If

                    If 太陽光電發電設備 = "有" AndAlso Not IsDBNull(table9.Rows(i)("太陽光電發電設備_設置位置")) Then
                        Dim solarPos As String = table9.Rows(i)("太陽光電發電設備_設置位置").ToString.Trim
                        說明太陽光電發電設備 = If(solarPos = "", "", "位置：" & solarPos)
                    Else
                        說明太陽光電發電設備 = ""
                    End If

                    If Not IsDBNull(table9.Rows(i)("建築能效標示")) AndAlso table9.Rows(i)("建築能效標示").ToString.Trim <> "" Then
                        建築能效標示 = table9.Rows(i)("建築能效標示").ToString.Trim
                    Else
                        建築能效標示 = "無"
                    End If

                    If 建築能效標示 = "有" Then
                        Dim 等級 As String = If(IsDBNull(table9.Rows(i)("建築能效標示_能效等級")), "", table9.Rows(i)("建築能效標示_能效等級").ToString.Trim)
                        Dim 有效期間 As String = If(IsDBNull(table9.Rows(i)("建築能效標示_證書效期")), "", table9.Rows(i)("建築能效標示_證書效期").ToString.Trim)
                        說明建築能效標示 = "能效等級：" & 等級 & "<w:br/>有效期間：" & 有效期間 & "<w:br/>自核發日起算5年內有效"
                    Else
                        說明建築能效標示 = ""
                    End If

                    '違增建使用權
                    If Not IsDBNull(table9.Rows(i)("違增建使用權")) Then
                        If table9.Rows(i)("違增建使用權").ToString <> "" Then
                            違增建使用權 = table9.Rows(i)("違增建使用權").ToString
                        Else
                            違增建使用權 = "無"
                        End If
                    Else
                        違增建使用權 = "無"
                    End If

                    '違增建使用權_說明
                    If Not IsDBNull(table9.Rows(i)("違增建使用權_說明")) Then
                        If table9.Rows(i)("違增建使用權_說明").ToString <> "" Then
                            說明違增建使用權 = table9.Rows(i)("違增建使用權_說明").ToString
                        End If
                        If table9.Rows(i)("違增建_面積").ToString <> "" Then
                            說明違增建使用權 &= ",違增建面積:約" & table9.Rows(i)("違增建_面積").ToString & "㎡"
                        End If
                        If table9.Rows(i)("違增建列管情形").ToString <> "" Then
                            說明違增建使用權 &= ",列管情形:" & table9.Rows(i)("違增建列管情形").ToString
                        End If
                        If table9.Rows(i)("違增建列管情形_說明").ToString <> "" Then
                            說明違增建使用權 &= "," & table9.Rows(i)("違增建列管情形_說明").ToString
                        End If
                    End If

                    '排水系統
                    If Not IsDBNull(table9.Rows(i)("排水系統")) Then
                        If table9.Rows(i)("排水系統").ToString = "有" Then
                            排水系統 = "有"


                        Else
                            排水系統 = "無"
                            '排水系統_說明 ps. 不正常的情況才需要說明
                            If Not IsDBNull(table9.Rows(i)("排水系統_說明")) Then
                                If table9.Rows(i)("排水系統_說明").ToString <> "" Then
                                    說明排水系統 = table9.Rows(i)("排水系統_說明").ToString
                                End If
                            End If
                        End If
                    Else
                        排水系統 = "無"
                    End If



                    '其他重要事項
                    If Not IsDBNull(table9.Rows(i)("其他重要事項")) Then
                        If table9.Rows(i)("其他重要事項").ToString <> "" Then
                            其他重要事項 = table9.Rows(i)("其他重要事項").ToString
                        Else
                            其他重要事項 = "無"
                        End If
                    Else
                        其他重要事項 = "無"
                    End If

                    '其他重要事項_說明
                    If Not IsDBNull(table9.Rows(i)("其他重要事項_說明")) Then
                        If table9.Rows(i)("其他重要事項_說明").ToString <> "" Then
                            說明其他重要事項 = table9.Rows(i)("其他重要事項_說明").ToString
                        End If
                    End If

                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_BuildManage
                    tempdata = tempdata.Replace("≠建號層次說明", 建號層次說明)
                    tempdata = tempdata.Replace("≠建物限制登記", 建物限制登記)
                    tempdata = tempdata.Replace("≠說明建物限制登記", 說明建物限制登記)
                    tempdata = tempdata.Replace("≠建物信託登記", 建物信託登記)
                    tempdata = tempdata.Replace("≠說明建物信託登記", 說明建物信託登記)
                    tempdata = tempdata.Replace("≠建物其他權利", 建物其他權利)
                    tempdata = tempdata.Replace("≠說明建物其他權利", 說明建物其他權利)


                    tempdata = tempdata.Replace("≠建物是否共有", 建物是否共有)
                    tempdata = tempdata.Replace("≠建物有無分管協議", 建物有無分管協議)
                    tempdata = tempdata.Replace("≠專有部分範圍", 有無專有部分範圍 & "," & 專有部分範圍)
                    tempdata = tempdata.Replace("≠共有部分範圍", 有無共有部分範圍 & "," & 共有部分範圍)
                    tempdata = tempdata.Replace("≠專有約定共用", 專有約定共用)
                    tempdata = tempdata.Replace("≠共有約定專用", 共有約定專用)
                    tempdata = tempdata.Replace("≠公共基金數額", 公共基金數額)
                    tempdata = tempdata.Replace("≠公共基金提撥方式", 公共基金提撥方式)
                    tempdata = tempdata.Replace("≠公共基金運用方式", 公共基金運用方式)
                    tempdata = tempdata.Replace("≠建物管理費", 建物管理費)
                    tempdata = tempdata.Replace("≠管理費繳交方式", 管理費繳交方式)
                    tempdata = tempdata.Replace("≠管理組織及方式", 管理組織及方式)
                    tempdata = tempdata.Replace("≠管理公司", 有無管理公司 & "," & 管理公司)
                    tempdata = tempdata.Replace("≠管理手冊", 管理手冊)
                    tempdata = tempdata.Replace("≠說明管理手冊", 說明管理手冊)
                    tempdata = tempdata.Replace("≠獎勵容積", 獎勵容積)
                    tempdata = tempdata.Replace("≠說明獎勵容積", 說明獎勵容積)
                    tempdata = tempdata.Replace("≠電梯設備", 電梯設備)
                    tempdata = tempdata.Replace("≠說明張貼合格標章", 說明張貼合格標章)
                    tempdata = tempdata.Replace("≠張貼合格標章", 張貼合格標章)

                    tempdata = tempdata.Replace("≠行動電話基地台", 頂樓基地台)
                    tempdata = tempdata.Replace("≠基地台說明", 基地台說明)
                    tempdata = tempdata.Replace("≠出租狀況", 出租狀況)
                    tempdata = tempdata.Replace("≠標題租金", 標題租金)
                    tempdata = tempdata.Replace("≠標題租期保證金", 標題租期保證金)
                    tempdata = tempdata.Replace("≠標題租期", 標題租期)
                    tempdata = tempdata.Replace("≠標題說明出租狀況", 標題說明出租狀況)
                    tempdata = tempdata.Replace("≠標題租約公證", 標題租約公證)
                    tempdata = tempdata.Replace("≠說明出租狀況", 說明出租狀況)
                    tempdata = tempdata.Replace("≠說明出借狀況", 說明出借狀況)
                    tempdata = tempdata.Replace("≠出借狀況", 出借狀況)
                    tempdata = tempdata.Replace("≠新標題說明出租", s說明出租狀況)
                    tempdata = tempdata.Replace("≠新說明出租", "")
                    tempdata = tempdata.Replace("≠租期保證金額", 租期保證金)
                    tempdata = tempdata.Replace("≠租金", 租金)
                    tempdata = tempdata.Replace("≠租期起", 租期起)
                    tempdata = tempdata.Replace("≠租期止", 租期止)
                    tempdata = tempdata.Replace("≠租約公證", 租約公證)
                    tempdata = tempdata.Replace("≠占用情形", 占用情形)
                    tempdata = tempdata.Replace("≠標題占用他人建物土地", 標題占用他人建物土地)
                    tempdata = tempdata.Replace("≠標題被他人占用建物土地", 標題被他人占用建物土地)
                    tempdata = tempdata.Replace("≠標題佔用情形其他", 標題佔用情形其他)
                    tempdata = tempdata.Replace("≠占用他人", 占用他人)
                    tempdata = tempdata.Replace("≠被他人占用", 被他人占用)
                    tempdata = tempdata.Replace("≠其他佔用情形", 佔用情形其他)
                    tempdata = tempdata.Replace("≠消防設備", 消防設備)
                    tempdata = tempdata.Replace("≠說明消防設備", 說明消防設備)
                    tempdata = tempdata.Replace("≠無障礙設施", 無障礙設施)
                    tempdata = tempdata.Replace("≠說明無障礙設施", 說明無障礙設施)
                    tempdata = tempdata.Replace("≠有無夾層", 有無夾層)
                    tempdata = tempdata.Replace("≠夾層面積_0", 夾層面積_0)
                    tempdata = tempdata.Replace("≠合法登記之夾層_0", 合法登記之夾層_0)
                    tempdata = tempdata.Replace("≠夾層面積_1", 夾層面積_1)
                    tempdata = tempdata.Replace("≠合法登記之夾層_1", 合法登記之夾層_1)
                    tempdata = tempdata.Replace("≠夾層面積_2", 夾層面積_2)
                    tempdata = tempdata.Replace("≠合法登記之夾層_2", 合法登記之夾層_2)
                    tempdata = tempdata.Replace("≠獨立供水", 獨立供水)
                    tempdata = tempdata.Replace("≠說明獨立供水", 說明獨立供水)
                    tempdata = tempdata.Replace("≠獨立電表", 獨立電表)
                    tempdata = tempdata.Replace("≠說明獨立電表", 說明獨立電表)
                    tempdata = tempdata.Replace("≠天然瓦斯", 天然瓦斯)
                    tempdata = tempdata.Replace("≠說明天然瓦斯", 說明天然瓦斯)
                    tempdata = tempdata.Replace("≠說明2天然瓦斯", 說明2天然瓦斯)
                    tempdata = tempdata.Replace("≠管線更新", 管線更新)
                    tempdata = tempdata.Replace("≠說明管線更新", 管線更新說明)
                    tempdata = tempdata.Replace("≠積欠費用", 積欠費用)
                    tempdata = tempdata.Replace("≠說明積欠費用", 說明積欠費用)
                    tempdata = tempdata.Replace("≠屬工業區或其他分區", 屬工業區或其他分區)
                    tempdata = tempdata.Replace("≠說明屬工業區或其他分區", 說明屬工業區或其他分區)
                    tempdata = tempdata.Replace("≠持有期間有無居住", 持有期間有無居住)
                    tempdata = tempdata.Replace("≠使照注意事項", 使照注意事項)
                    tempdata = tempdata.Replace("≠說明使照注意事項", 說明使照注意事項)
                    tempdata = tempdata.Replace("≠有無公共設施重大修繕", 有無公共設施重大修繕)
                    tempdata = tempdata.Replace("≠說明有無公共設施重大修繕", 說明有無公共設施重大修繕)
                    tempdata = tempdata.Replace("≠下水道工程", 衛生下水道)
                    tempdata = tempdata.Replace("≠說明下水道工程", 說明衛生下水道)
                    tempdata = tempdata.Replace("≠附隨買賣設備", 隨附設備)
                    tempdata = tempdata.Replace("≠說明附隨買賣設備", 說明隨附設備)
                    tempdata = tempdata.Replace("≠說明是否有車位", 說明是否有車位) '20160625 add by nick
                    tempdata = tempdata.Replace("≠是否有車位", 是否有車位)
                    tempdata = tempdata.Replace("≠氯離子含量", 氯離子含量)
                    tempdata = tempdata.Replace("≠氯離子檢測結果", 氯離子檢測結果)
                    tempdata = tempdata.Replace("≠檢測日期", 檢測日期)
                    tempdata = tempdata.Replace("≠檢測值", 檢測值) '用不到
                    tempdata = tempdata.Replace("≠有無輻射檢測", 有無輻射檢測)
                    tempdata = tempdata.Replace("≠輻射檢測結果", 輻射檢測結果)
                    tempdata = tempdata.Replace("≠輻射檢測值", 輻射檢測值)
                    tempdata = tempdata.Replace("≠發生火災", 發生火災)
                    tempdata = tempdata.Replace("≠標題修繕情形", 標題修繕情形)
                    tempdata = tempdata.Replace("≠修繕情形", 修繕情形)
                    tempdata = tempdata.Replace("≠發生地震", 發生地震)
                    tempdata = tempdata.Replace("≠地震等級", 地震等級) '用不到
                    tempdata = tempdata.Replace("≠地震說明", 地震說明)
                    tempdata = tempdata.Replace("≠樑柱有無裂痕", 樑柱有無裂痕)
                    tempdata = tempdata.Replace("≠說明樑柱有無裂痕", 說明樑柱有無裂痕)
                    tempdata = tempdata.Replace("≠建物鋼筋裸露", 建物鋼筋裸露)
                    tempdata = tempdata.Replace("≠鋼筋裸露說明", 鋼筋裸露說明)
                    tempdata = tempdata.Replace("≠有無兇殺情形", 有無兇殺情形)
                    tempdata = tempdata.Replace("≠說明有無兇殺情形", 說明有無兇殺情形)
                    'tempdata = tempdata.Replace("≠非持有有無兇殺情形", 非持有有無兇殺情形)
                    'tempdata = tempdata.Replace("≠非持有說明有無兇殺情形", 非持有說明有無兇殺情形)

                    tempdata = tempdata.Replace("≠有無漏水", 有無漏水)
                    tempdata = tempdata.Replace("≠說明有無漏水", 說明有無漏水)
                    tempdata = tempdata.Replace("≠有無禁建情事", 有無禁建情事)
                    tempdata = tempdata.Replace("≠說明有無禁建情事", 說明有無禁建情事)
                    tempdata = tempdata.Replace("≠是否有中繼幫浦", 中繼幫浦)
                    tempdata = tempdata.Replace("≠中繼幫浦說明", 說明中繼幫浦)

                    tempdata = tempdata.Replace("≠太陽光電發電設備說明", 說明太陽光電發電設備)
                    tempdata = tempdata.Replace("≠太陽光電發電設備", 太陽光電發電設備)
                    tempdata = tempdata.Replace("≠建築能效標示說明", 說明建築能效標示)
                    tempdata = tempdata.Replace("≠建築能效標示", 建築能效標示)

                    tempdata = tempdata.Replace("≠說明違增建使用權", 說明違增建使用權)
                    tempdata = tempdata.Replace("≠違增建使用權", 違增建使用權)
                    tempdata = tempdata.Replace("≠說明排水系統", 說明排水系統)
                    tempdata = tempdata.Replace("≠排水系統", 排水系統)

                    tempdata = tempdata.Replace("≠其他重要事項", 其他重要事項)
                    tempdata = tempdata.Replace("≠說明其他重要事項", 說明其他重要事項)

                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_建物管理狀況.Append(tempdata)
                Next
            Else
                Dim tempdata As String = myText_BuildManage
                tempdata = tempdata.Replace("≠建號層次說明", "")
                tempdata = tempdata.Replace("≠建物限制登記", "")
                tempdata = tempdata.Replace("≠說明建物限制登記", "")
                tempdata = tempdata.Replace("≠建物信託登記", "")
                tempdata = tempdata.Replace("≠說明建物信託登記", "")
                tempdata = tempdata.Replace("≠建物其他權利", "")
                tempdata = tempdata.Replace("≠說明建物其他權利", "")

                tempdata = tempdata.Replace("≠新標題說明出租", "")
                tempdata = tempdata.Replace("≠新說明出租", "")

                tempdata = tempdata.Replace("≠說明張貼合格標章", "")
                tempdata = tempdata.Replace("≠建物是否共有", "")
                tempdata = tempdata.Replace("≠建物有無分管協議", "")
                tempdata = tempdata.Replace("≠有無專有部分範圍", "")
                tempdata = tempdata.Replace("≠專有部分範圍", "")
                tempdata = tempdata.Replace("≠有無共有部分範圍", "")
                tempdata = tempdata.Replace("≠共有部分範圍", "")
                tempdata = tempdata.Replace("≠專有約定共用", "")
                tempdata = tempdata.Replace("≠共有約定專用", "")
                tempdata = tempdata.Replace("≠公共基金數額", "")
                tempdata = tempdata.Replace("≠公共基金提撥方式", "")
                tempdata = tempdata.Replace("≠公共基金運用方式", "")
                tempdata = tempdata.Replace("≠建物管理費", "")
                tempdata = tempdata.Replace("≠管理費繳交方式", "")
                tempdata = tempdata.Replace("≠管理組織及方式", "")
                tempdata = tempdata.Replace("≠管理公司", "")
                tempdata = tempdata.Replace("≠管理手冊", "")
                tempdata = tempdata.Replace("≠說明管理手冊", "")
                tempdata = tempdata.Replace("≠獎勵容積", "")
                tempdata = tempdata.Replace("≠說明獎勵容積", "")
                tempdata = tempdata.Replace("≠電梯設備", "")
                tempdata = tempdata.Replace("≠張貼合格標章", "")
                tempdata = tempdata.Replace("≠張貼合格標章說明", "")
                tempdata = tempdata.Replace("≠行動電話基地台", "")
                tempdata = tempdata.Replace("≠基地台說明", "")
                tempdata = tempdata.Replace("≠出租狀況", "")
                tempdata = tempdata.Replace("≠標題租金", "")
                tempdata = tempdata.Replace("≠標題租期", "")
                tempdata = tempdata.Replace("≠標題租約公證", "")
                tempdata = tempdata.Replace("≠租期保證金額", "")
                tempdata = tempdata.Replace("≠標題說明出租狀況", "")
                tempdata = tempdata.Replace("≠說明出租狀況", "")
                tempdata = tempdata.Replace("≠出借狀況", "")
                tempdata = tempdata.Replace("≠說明出借狀況", "")
                tempdata = tempdata.Replace("≠租金", "")
                tempdata = tempdata.Replace("≠租期起", "")
                tempdata = tempdata.Replace("≠租期止", "")
                tempdata = tempdata.Replace("≠租約公證", "")
                tempdata = tempdata.Replace("≠占用情形", "")
                tempdata = tempdata.Replace("≠標題占用他人建物土地", "")
                tempdata = tempdata.Replace("≠標題被他人占用建物土地", "")
                tempdata = tempdata.Replace("≠標題佔用情形其他", "")
                tempdata = tempdata.Replace("≠占用他人", "")
                tempdata = tempdata.Replace("≠被他人占用", "")
                tempdata = tempdata.Replace("≠其他佔用情形", "")
                tempdata = tempdata.Replace("≠消防設備", "")
                tempdata = tempdata.Replace("≠說明消防設備", "")
                tempdata = tempdata.Replace("≠無障礙設施", "")
                tempdata = tempdata.Replace("≠說明無障礙設施", "")
                tempdata = tempdata.Replace("≠有無夾層", "")
                tempdata = tempdata.Replace("≠夾層面積_0", "")
                tempdata = tempdata.Replace("≠合法登記之夾層_0", "")
                tempdata = tempdata.Replace("≠夾層面積_1", "")
                tempdata = tempdata.Replace("≠合法登記之夾層_1", "")
                tempdata = tempdata.Replace("≠夾層面積_2", "")
                tempdata = tempdata.Replace("≠合法登記之夾層_2", "")
                tempdata = tempdata.Replace("≠獨立供水", "")
                tempdata = tempdata.Replace("≠說明獨立供水", "")
                tempdata = tempdata.Replace("≠獨立電表", "")
                tempdata = tempdata.Replace("≠說明獨立電表", "")
                tempdata = tempdata.Replace("≠天然瓦斯", "")
                tempdata = tempdata.Replace("≠說明天然瓦斯", "")
                tempdata = tempdata.Replace("≠說明2天然瓦斯", "")
                tempdata = tempdata.Replace("≠管線更新", "")
                tempdata = tempdata.Replace("≠說明管線更新", "")
                tempdata = tempdata.Replace("≠積欠費用", "")
                tempdata = tempdata.Replace("≠說明積欠費用", "")
                tempdata = tempdata.Replace("≠屬工業區或其他分區", "")
                tempdata = tempdata.Replace("≠說明屬工業區或其他分區", "")
                tempdata = tempdata.Replace("≠持有期間有無居住", "")
                tempdata = tempdata.Replace("≠使照注意事項", "")
                tempdata = tempdata.Replace("≠說明使照注意事項", "")
                tempdata = tempdata.Replace("≠有無公共設施重大修繕", "")
                tempdata = tempdata.Replace("≠說明有無公共設施重大修繕", "")
                tempdata = tempdata.Replace("≠下水道工程", "")
                tempdata = tempdata.Replace("≠說明下水道工程", "")
                tempdata = tempdata.Replace("≠附隨買賣設備", "")
                tempdata = tempdata.Replace("≠說明附隨買賣設備", "")
                tempdata = tempdata.Replace("≠說明是否有車位", "") '20160625 add by nick
                tempdata = tempdata.Replace("≠是否有車位", "")
                tempdata = tempdata.Replace("≠氯離子含量", "")
                tempdata = tempdata.Replace("≠氯離子檢測結果", "")
                tempdata = tempdata.Replace("≠檢測日期", "")
                tempdata = tempdata.Replace("≠檢測值", "") '用不到
                tempdata = tempdata.Replace("≠有無輻射檢測", "")
                tempdata = tempdata.Replace("≠輻射檢測結果", "")
                tempdata = tempdata.Replace("≠輻射檢測值", "")
                tempdata = tempdata.Replace("≠發生火災", "")
                tempdata = tempdata.Replace("≠標題修繕情形", "")
                tempdata = tempdata.Replace("≠修繕情形", "")
                tempdata = tempdata.Replace("≠發生地震", "")
                tempdata = tempdata.Replace("≠地震等級", "") '用不到
                tempdata = tempdata.Replace("≠地震說明", "")
                tempdata = tempdata.Replace("≠樑柱有無裂痕", "")
                tempdata = tempdata.Replace("≠說明樑柱有無裂痕", "")
                tempdata = tempdata.Replace("≠建物鋼筋裸露", "")
                tempdata = tempdata.Replace("≠鋼筋裸露說明", "")
                tempdata = tempdata.Replace("≠有無兇殺情形", "")
                tempdata = tempdata.Replace("≠說明有無兇殺情形", "")
                'tempdata = tempdata.Replace("≠非持有有無兇殺情形", "")
                'tempdata = tempdata.Replace("≠非持有說明有無兇殺情形", "")
                tempdata = tempdata.Replace("≠有無漏水", "")
                tempdata = tempdata.Replace("≠說明有無漏水", "")
                tempdata = tempdata.Replace("≠有無禁建情事", "")
                tempdata = tempdata.Replace("≠說明有無禁建情事", "")
                tempdata = tempdata.Replace("≠是否有中繼幫浦", "")
                tempdata = tempdata.Replace("≠中繼幫浦說明", "")

                tempdata = tempdata.Replace("≠太陽光電發電設備說明", "")
                tempdata = tempdata.Replace("≠太陽光電發電設備", "無")
                tempdata = tempdata.Replace("≠建築能效標示說明", "")
                tempdata = tempdata.Replace("≠建築能效標示", "無")

                tempdata = tempdata.Replace("≠違增建使用權", "")
                tempdata = tempdata.Replace("≠說明違增建使用權", "")
                tempdata = tempdata.Replace("≠排水系統", "")
                tempdata = tempdata.Replace("≠說明排水系統", "")
                tempdata = tempdata.Replace("≠其他重要事項", "")
                tempdata = tempdata.Replace("≠說明其他重要事項", "")



                '改完加到最後的字串裡面  
                最後要替代掉的字串_建物管理狀況.Append(tempdata)
            End If

            sFile = sFile.Replace("!重複建物管理狀況", 最後要替代掉的字串_建物管理狀況.ToString())

            '[END============================================產權調查-建物目前管理狀況P1-P3=============================================END]


        Else '土地
            '[START============================================產權調查-土地目前管理狀況P1-P3=============================================START]
            Dim 土地建號說明 As String = ""
            '壹.
            Dim 土地限制登記 As String = ""
            Dim 說明土地限制登記 As String = ""
            Dim 土地信託登記 As String = ""
            Dim 說明土地信託登記 As String = ""
            Dim 土地其他權利 As String = ""
            Dim 說明土地其他權利 As String = ""
            '貳.
            Dim 土地是否依慣例使用 As String = ""
            Dim 說明土地是否依慣例使用 As String = ""
            Dim 分管協議 As String = ""
            Dim 說明分管協議 As String = ""
            Dim 出租出借 As String = ""
            Dim 說明出租出借 As String = ""
            Dim 無權占用 As String = ""
            Dim 說明無權占用 As String = ""
            Dim 公眾通行私有道路 As String = ""
            Dim 說明公眾通行私有道路 As String = ""
            Dim 地籍圖重測 As String = ""
            Dim 標題主管機關已公告辦理 As String = ""
            Dim 主管機關已公告辦理 As String = ""
            Dim 越界建築 As String = ""
            Dim 說明越界建築 As String = ""
            Dim 公告徵收 As String = ""
            Dim 說明公告徵收範圍 As String = ""
            Dim 公共基礎設施 As String = ""
            Dim 說明公共基礎設施 As String = ""
            '叁.
            Dim 禁限建地區 As String = ""
            Dim 說明禁限建地區 As String = ""
            Dim 不得興建農舍 As String = ""
            Dim 說明不得興建農舍 As String = ""
            Dim 山坡地範圍 As String = ""
            Dim 說明山坡地範圍 As String = ""
            Dim 水土保持區 As String = ""
            Dim 說明水土保持區 As String = ""
            Dim 河川區域 As String = ""
            Dim 說明河川區域 As String = ""
            Dim 排水設施 As String = ""
            Dim 說明排水設施 As String = ""
            Dim 國家公園 As String = ""
            Dim 說明國家公園 As String = ""
            Dim 水質保護區 As String = ""
            Dim 說明水質保護區 As String = ""
            Dim 水量保護區 As String = ""
            Dim 說明水量保護區 As String = ""
            Dim 地下水汙染場址 As String = ""
            Dim 說明地下水汙染場址 As String = ""


            '土地目前管理狀況
            Dim myXML_PlaceManage As String = ""
            Dim myText_PlaceManage As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_PlaceManage As New StreamReader(Server.MapPath("..\reports\重複土地管理狀況_V2.txt"))
            myText_PlaceManage = srText_PlaceManage.ReadToEnd()
            srText_PlaceManage.Close()

            sqlstr = ""
            sqlstr = "Select * From 產調_土地 A  With(NoLock) Left Join 委賣物件資料表_面積細項 B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號  Where A.物件編號 = '" & Contract_No & "' and A.店代號='" & Request("sid") & "'  order by A.流水號 "


            Dim table9 As DataTable



            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table9")
            table9 = ds.Tables("table9")

            Dim 最後要替代掉的字串_土地管理狀況 As New StringBuilder()

            If table9.Rows.Count > 0 Then
                For i = 0 To table9.Rows.Count - 1
                    土地建號說明 = "<w:br/>‎"
                    If Not IsDBNull(table9.Rows(i)("建號")) Then
                        If table9.Rows(i)("建號").ToString.Trim <> "" Then
                            土地建號說明 &= "(" & table9.Rows(i)("建號").ToString & ")"
                        End If
                    End If
                    '壹
                    '限制登記
                    If Not IsDBNull(table9.Rows(i)("限制登記")) Then
                        土地限制登記 = table9.Rows(i)("限制登記").ToString
                    End If
                    '說明限制登記
                    If Not IsDBNull(table9.Rows(i)("限制登記_說明")) Then
                        說明土地限制登記 = table9.Rows(i)("限制登記_說明").ToString
                    End If
                    '信託登記
                    If Not IsDBNull(table9.Rows(i)("信託登記")) Then
                        土地信託登記 = table9.Rows(i)("信託登記").ToString
                    End If
                    '說明信託登記
                    If Not IsDBNull(table9.Rows(i)("信託登記_說明")) Then
                        說明土地信託登記 = table9.Rows(i)("信託登記_說明").ToString
                    End If
                    '其他權利
                    If Not IsDBNull(table9.Rows(i)("其他權利")) Then
                        土地其他權利 = table9.Rows(i)("其他權利").ToString
                    End If
                    '說明其他權利
                    If Not IsDBNull(table9.Rows(i)("其他權利_說明")) Then
                        說明土地其他權利 = table9.Rows(i)("其他權利_說明").ToString
                    End If

                    '貳
                    '是否依慣例使用
                    If Not IsDBNull(table9.Rows(i)("是否依慣例使用")) Then
                        土地是否依慣例使用 = table9.Rows(i)("是否依慣例使用").ToString
                    End If
                    '說明是否依慣例使用
                    If Not IsDBNull(table9.Rows(i)("是否依慣例使用_說明")) Then
                        說明土地是否依慣例使用 = table9.Rows(i)("是否依慣例使用_說明").ToString
                    End If
                    '分管協議
                    If Not IsDBNull(table9.Rows(i)("共有土地有無分管協議")) Then
                        分管協議 = table9.Rows(i)("共有土地有無分管協議").ToString
                    End If
                    '說明分管協議
                    If Not IsDBNull(table9.Rows(i)("共有土地有無分管協議_說明")) Then
                        說明分管協議 = table9.Rows(i)("共有土地有無分管協議_說明").ToString
                    End If
                    '出租出借
                    If Not IsDBNull(table9.Rows(i)("有無出租或出借")) Then
                        出租出借 = table9.Rows(i)("有無出租或出借").ToString
                    End If
                    '說明出租出借
                    If Not IsDBNull(table9.Rows(i)("有無出租或出借_說明")) Then
                        說明出租出借 = table9.Rows(i)("有無出租或出借_說明").ToString
                    End If
                    '無權占用
                    If Not IsDBNull(table9.Rows(i)("有無被他人無權占用")) Then
                        無權占用 = table9.Rows(i)("有無被他人無權占用").ToString
                    End If
                    '說明無權占用
                    If Not IsDBNull(table9.Rows(i)("有無被他人無權占用_說明")) Then
                        說明無權占用 = table9.Rows(i)("有無被他人無權占用_說明").ToString
                    End If
                    '公眾通行私有道路
                    If Not IsDBNull(table9.Rows(i)("供公眾通行之私有道路")) Then
                        公眾通行私有道路 = table9.Rows(i)("供公眾通行之私有道路").ToString

                        If table9.Rows(i)("供公眾通行之私有道路").ToString = "有" Then
                            If Not IsDBNull(table9.Rows(i)("供公眾通行之私有道路_說明")) Then
                                If Trim(table9.Rows(i)("供公眾通行之私有道路_說明").ToString) <> "" Then
                                    說明公眾通行私有道路 = table9.Rows(i)("供公眾通行之私有道路_說明").ToString
                                End If
                            End If

                        End If
                    Else
                        公眾通行私有道路 = "無"
                    End If
                    '地籍圖重測
                    If Not IsDBNull(table9.Rows(i)("辦理地籍圖重測")) Then
                        If table9.Rows(i)("辦理地籍圖重測").ToString <> "" Then
                            地籍圖重測 = table9.Rows(i)("辦理地籍圖重測").ToString

                            If table9.Rows(i)("辦理地籍圖重測").ToString = "有" Then
                                標題主管機關已公告辦理 = ""
                                主管機關已公告辦理 = ""

                            Else
                                標題主管機關已公告辦理 = "主管機關已公告辦理"
                                '主管機關已公告辦理
                                If Not IsDBNull(table9.Rows(i)("主管機關已公告辦理")) Then
                                    If table9.Rows(i)("主管機關已公告辦理").ToString <> "" Then
                                        主管機關已公告辦理 = table9.Rows(0)("主管機關已公告辦理").ToString
                                    End If
                                End If
                            End If
                        Else
                            地籍圖重測 = "有"
                        End If
                    Else
                        地籍圖重測 = "有"
                    End If
                    '越界建築
                    If Not IsDBNull(table9.Rows(i)("是否越界建築")) Then
                        越界建築 = table9.Rows(i)("是否越界建築").ToString
                    End If
                    '說明越界建築
                    If Not IsDBNull(table9.Rows(i)("是否越界建築_說明")) Then
                        說明越界建築 = table9.Rows(i)("是否越界建築_說明").ToString
                    End If
                    '公告徵收
                    If Not IsDBNull(table9.Rows(i)("公告徵收")) Then
                        公告徵收 = table9.Rows(i)("公告徵收").ToString

                        If table9.Rows(i)("公告徵收").ToString = "有" Then
                            If Not IsDBNull(table9.Rows(i)("公告徵收範圍")) Then
                                If Trim(table9.Rows(i)("公告徵收範圍").ToString) <> "" Then
                                    說明公告徵收範圍 = table9.Rows(i)("公告徵收範圍").ToString
                                End If
                            End If

                        End If
                    Else
                        公告徵收 = "無"
                    End If
                    '公共基礎建設
                    If Not IsDBNull(table9.Rows(i)("有無公共基礎設施")) Then
                        公共基礎設施 = table9.Rows(i)("有無公共基礎設施").ToString

                        If table9.Rows(i)("有無公共基礎設施").ToString = "有" Then
                            If Not IsDBNull(table9.Rows(i)("有無公共基礎設施_說明")) Then
                                If Trim(table9.Rows(i)("有無公共基礎設施_說明").ToString) <> "" Then
                                    說明公共基礎設施 = table9.Rows(i)("有無公共基礎設施_說明").ToString
                                End If
                            End If

                        End If
                    Else
                        公共基礎設施 = "無"
                    End If

                    '叁
                    '禁限建地區
                    If Not IsDBNull(table9.Rows(i)("禁限建地區")) Then
                        禁限建地區 = table9.Rows(i)("禁限建地區").ToString
                    End If
                    '說明禁限建地區
                    If Not IsDBNull(table9.Rows(i)("禁限建地區_說明")) Then
                        說明禁限建地區 = table9.Rows(i)("禁限建地區_說明").ToString
                    End If
                    '不得興建農舍
                    If Not IsDBNull(table9.Rows(i)("不得興建農舍")) Then
                        不得興建農舍 = table9.Rows(i)("不得興建農舍").ToString
                    End If
                    '說明不得興建農舍
                    If Not IsDBNull(table9.Rows(i)("不得興建農舍_說明")) Then
                        說明不得興建農舍 = table9.Rows(i)("不得興建農舍_說明").ToString
                    End If
                    '山坡地範圍
                    If Not IsDBNull(table9.Rows(i)("山坡地範圍")) Then
                        山坡地範圍 = table9.Rows(i)("山坡地範圍").ToString
                    End If
                    '說明山坡地範圍
                    If Not IsDBNull(table9.Rows(i)("山坡地範圍_說明")) Then
                        說明山坡地範圍 = table9.Rows(i)("山坡地範圍_說明").ToString
                    End If
                    '水土保持區
                    If Not IsDBNull(table9.Rows(i)("水土保持區")) Then
                        水土保持區 = table9.Rows(i)("水土保持區").ToString
                    End If
                    '說明水土保持區
                    If Not IsDBNull(table9.Rows(i)("水土保持區_說明")) Then
                        說明水土保持區 = table9.Rows(i)("水土保持區_說明").ToString
                    End If
                    '河川區域
                    If Not IsDBNull(table9.Rows(i)("河川區域")) Then
                        河川區域 = table9.Rows(i)("河川區域").ToString
                    End If
                    '說明河川區域
                    If Not IsDBNull(table9.Rows(i)("河川區域_說明")) Then
                        說明河川區域 = table9.Rows(i)("河川區域_說明").ToString
                    End If
                    '排水設施
                    If Not IsDBNull(table9.Rows(i)("排水設施")) Then
                        排水設施 = table9.Rows(i)("排水設施").ToString
                    End If
                    '說明排水設施
                    If Not IsDBNull(table9.Rows(i)("排水設施_說明")) Then
                        說明排水設施 = table9.Rows(i)("排水設施_說明").ToString
                    End If
                    '國家公園
                    If Not IsDBNull(table9.Rows(i)("國家公園")) Then
                        國家公園 = table9.Rows(i)("國家公園").ToString
                    End If
                    '說明國家公園
                    If Not IsDBNull(table9.Rows(i)("國家公園_說明")) Then
                        說明國家公園 = table9.Rows(i)("國家公園_說明").ToString
                    End If
                    '水質保護區
                    If Not IsDBNull(table9.Rows(i)("水質保護區")) Then
                        水質保護區 = table9.Rows(i)("水質保護區").ToString
                    End If
                    '說明水質保護區
                    If Not IsDBNull(table9.Rows(i)("水質保護區_說明")) Then
                        說明水質保護區 = table9.Rows(i)("水質保護區_說明").ToString
                    End If
                    '水量保護區
                    If Not IsDBNull(table9.Rows(i)("水量保護區")) Then
                        水量保護區 = table9.Rows(i)("水量保護區").ToString
                    End If
                    '說明水量保護區
                    If Not IsDBNull(table9.Rows(i)("水量保護區_說明")) Then
                        說明水量保護區 = table9.Rows(i)("水量保護區_說明").ToString
                    End If
                    '地下水汙染場址
                    If Not IsDBNull(table9.Rows(i)("地下水汙染場址")) Then
                        地下水汙染場址 = table9.Rows(i)("地下水汙染場址").ToString
                    End If
                    '說明地下水汙染場址
                    If Not IsDBNull(table9.Rows(i)("地下水汙染場址_說明")) Then
                        說明地下水汙染場址 = table9.Rows(i)("地下水汙染場址_說明").ToString
                    End If


                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_PlaceManage
                    tempdata = tempdata.Replace("≠土地建號說明", 土地建號說明)
                    tempdata = tempdata.Replace("≠土地限制登記", 土地限制登記)
                    tempdata = tempdata.Replace("≠說明土地限制登記", 說明土地限制登記)
                    tempdata = tempdata.Replace("≠土地信託登記", 土地信託登記)
                    tempdata = tempdata.Replace("≠說明土地信託登記", 說明土地信託登記)
                    tempdata = tempdata.Replace("≠土地其他權利", 土地其他權利)
                    tempdata = tempdata.Replace("≠說明土地其他權利", 說明土地其他權利)

                    tempdata = tempdata.Replace("≠土地是否依慣例使用", 土地是否依慣例使用)
                    tempdata = tempdata.Replace("≠說明土地是否依慣例使用", 說明土地是否依慣例使用)
                    tempdata = tempdata.Replace("≠分管協議", 分管協議)
                    tempdata = tempdata.Replace("≠說明分管協議", 說明分管協議)
                    tempdata = tempdata.Replace("≠出租出借", 出租出借)
                    tempdata = tempdata.Replace("≠說明出租出借", 說明出租出借)
                    tempdata = tempdata.Replace("≠無權占用", 無權占用)
                    tempdata = tempdata.Replace("≠說明無權占用", 說明無權占用)
                    tempdata = tempdata.Replace("≠公眾通行私有道路", 公眾通行私有道路)
                    tempdata = tempdata.Replace("≠說明公眾通行私有道路", 說明公眾通行私有道路)
                    tempdata = tempdata.Replace("≠地籍圖重測", 地籍圖重測)
                    tempdata = tempdata.Replace("≠標題主管機關已公告辦理", 標題主管機關已公告辦理)
                    tempdata = tempdata.Replace("≠主管機關已公告辦理", 主管機關已公告辦理)
                    tempdata = tempdata.Replace("≠越界建築", 越界建築)
                    tempdata = tempdata.Replace("≠說明越界建築", 說明越界建築)
                    tempdata = tempdata.Replace("≠公告徵收", 公告徵收)
                    tempdata = tempdata.Replace("≠說明公告徵收範圍", 說明公告徵收範圍)
                    tempdata = tempdata.Replace("≠公共基礎設施", 公共基礎設施)
                    tempdata = tempdata.Replace("≠說明公共基礎設施", 說明公共基礎設施)

                    tempdata = tempdata.Replace("≠禁限建地區", 禁限建地區)
                    tempdata = tempdata.Replace("≠說明禁限建地區", 說明禁限建地區)
                    tempdata = tempdata.Replace("≠不得興建農舍", 不得興建農舍)
                    tempdata = tempdata.Replace("≠說明不得興建農舍", 說明不得興建農舍)
                    tempdata = tempdata.Replace("≠山坡地範圍", 山坡地範圍)
                    tempdata = tempdata.Replace("≠說明山坡地範圍", 說明山坡地範圍)
                    tempdata = tempdata.Replace("≠水土保持區", 水土保持區)
                    tempdata = tempdata.Replace("≠說明水土保持區", 說明水土保持區)
                    tempdata = tempdata.Replace("≠河川區域", 河川區域)
                    tempdata = tempdata.Replace("≠說明河川區域", 說明河川區域)
                    tempdata = tempdata.Replace("≠排水設施", 排水設施)
                    tempdata = tempdata.Replace("≠說明排水設施", 說明排水設施)
                    tempdata = tempdata.Replace("≠國家公園", 國家公園)
                    tempdata = tempdata.Replace("≠說明國家公園", 說明國家公園)
                    tempdata = tempdata.Replace("≠水質保護區", 水質保護區)
                    tempdata = tempdata.Replace("≠說明水質保護區", 說明水質保護區)
                    tempdata = tempdata.Replace("≠水量保護區", 水量保護區)
                    tempdata = tempdata.Replace("≠說明水量保護區", 說明水量保護區)
                    tempdata = tempdata.Replace("≠地下水汙染場址", 地下水汙染場址)
                    tempdata = tempdata.Replace("≠說明地下水汙染場址", 說明地下水汙染場址)
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_土地管理狀況.Append(tempdata)
                Next
            Else
                Dim tempdata As String = myText_PlaceManage
                tempdata = tempdata.Replace("≠土地建號說明", "")
                tempdata = tempdata.Replace("≠土地限制登記", "")
                tempdata = tempdata.Replace("≠說明土地限制登記", "")
                tempdata = tempdata.Replace("≠土地信託登記", "")
                tempdata = tempdata.Replace("≠說明土地信託登記", "")
                tempdata = tempdata.Replace("≠土地其他權利", "")
                tempdata = tempdata.Replace("≠說明土地其他權利", "")

                tempdata = tempdata.Replace("≠土地是否依慣例使用", "")
                tempdata = tempdata.Replace("≠說明土地是否依慣例使用", "")
                tempdata = tempdata.Replace("≠分管協議", "")
                tempdata = tempdata.Replace("≠說明分管協議", "")
                tempdata = tempdata.Replace("≠出租出借", "")
                tempdata = tempdata.Replace("≠說明出租出借", "")
                tempdata = tempdata.Replace("≠無權占用", "")
                tempdata = tempdata.Replace("≠說明無權占用", "")
                tempdata = tempdata.Replace("≠公眾通行私有道路", "")
                tempdata = tempdata.Replace("≠說明公眾通行私有道路", "")
                tempdata = tempdata.Replace("≠地籍圖重測", "")
                tempdata = tempdata.Replace("≠標題主管機關已公告辦理", "")
                tempdata = tempdata.Replace("≠主管機關已公告辦理", "")
                tempdata = tempdata.Replace("≠越界建築", "")
                tempdata = tempdata.Replace("≠說明越界建築", "")
                tempdata = tempdata.Replace("≠公告徵收", "")
                tempdata = tempdata.Replace("≠說明公告徵收範圍", "")
                tempdata = tempdata.Replace("≠公共基礎設施", "")
                tempdata = tempdata.Replace("≠說明公共基礎設施", "")

                tempdata = tempdata.Replace("≠禁限建地區", "")
                tempdata = tempdata.Replace("≠說明禁限建地區", "")
                tempdata = tempdata.Replace("≠不得興建農舍", "")
                tempdata = tempdata.Replace("≠說明不得興建農舍", "")
                tempdata = tempdata.Replace("≠山坡地範圍", "")
                tempdata = tempdata.Replace("≠說明山坡地範圍", "")
                tempdata = tempdata.Replace("≠水土保持區", "")
                tempdata = tempdata.Replace("≠說明水土保持區", "")
                tempdata = tempdata.Replace("≠河川區域", "")
                tempdata = tempdata.Replace("≠說明河川區域", "")
                tempdata = tempdata.Replace("≠排水設施", "")
                tempdata = tempdata.Replace("≠說明排水設施", "")
                tempdata = tempdata.Replace("≠國家公園", "")
                tempdata = tempdata.Replace("≠說明國家公園", "")
                tempdata = tempdata.Replace("≠水質保護區", "")
                tempdata = tempdata.Replace("≠說明水質保護區", "")
                tempdata = tempdata.Replace("≠水量保護區", "")
                tempdata = tempdata.Replace("≠說明水量保護區", "")
                tempdata = tempdata.Replace("≠地下水汙染場址", "")
                tempdata = tempdata.Replace("≠說明地下水汙染場址", "")



                '改完加到最後的字串裡面  
                最後要替代掉的字串_土地管理狀況.Append(tempdata)
            End If
            'If Request("oid") = "60692AAE34137" Then

            'Else
            sFile = sFile.Replace("!重複土地管理狀況", 最後要替代掉的字串_土地管理狀況.ToString())
            'End If

            '[END============================================產權調查-建物目前管理狀況P1-P3=============================================END]

        End If '-------------------------------------------------------------------------判斷成屋還是土地用





        '共用
        '[START============================================產權調查-增建(合法增建)=============================================START]
        '合法增建細項
        Dim myXML_合法增建 As String = ""
        Dim myText_合法增建 As String = ""

        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        Dim srText_合法增建 As New StreamReader(Server.MapPath("..\reports\重複合法增建_V2.txt"))
        myText_合法增建 = srText_合法增建.ReadToEnd()
        srText_合法增建.Close()

        sqlstr = ""
        sqlstr = "Select * From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別='增建' and 項目名稱='未辦理第一次登記'  order by 流水號 "


        Dim table77 As DataTable
        Dim 未登記項次 As Integer
        Dim 未登記座落 As String = ""
        Dim 未登記完成日 As String = ""
        Dim 未登記坪 As String = ""
        Dim 未登記所有權 As String = ""
        Dim 未登記權利範圍 As String = ""
        Dim 未登記坪和 As Decimal = 0

        i = 0
        adpt = New SqlDataAdapter(sqlstr, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table77")
        table77 = ds.Tables("table77")
        Dim 最後要替代掉的字串_合法增建 As New StringBuilder()

        If table77.Rows.Count > 0 Then
            For i = 0 To table77.Rows.Count - 1

                '未登記項次
                未登記項次 = i + 1

                '未登記座落
                If Not IsDBNull(table77.Rows(i)("建號")) Then
                    未登記座落 = table77.Rows(i)("建號").ToString
                Else
                    未登記座落 = ""
                End If

                '未登記完成日
                If Not IsDBNull(table77.Rows(i)("增建完成日期")) Then
                    未登記完成日 = table77.Rows(i)("增建完成日期").ToString
                Else
                    未登記完成日 = ""
                End If

                '未登記坪
                If Not IsDBNull(table77.Rows(i)("實際持有坪")) Then
                    '總和
                    未登記坪和 += table77.Rows(i)("實際持有坪")

                    If Right(table77.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                        未登記坪 = Int(table77.Rows(i)("實際持有坪").ToString)
                    ElseIf table77.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table77.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                        未登記坪 = Left(table77.Rows(i)("實際持有坪").ToString, Len(table77.Rows(i)("實際持有坪").ToString) - 3)
                    ElseIf table77.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table77.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                        未登記坪 = Left(table77.Rows(i)("實際持有坪").ToString, Len(table77.Rows(i)("實際持有坪").ToString) - 2)
                    ElseIf table77.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table77.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                        未登記坪 = Left(table77.Rows(i)("實際持有坪").ToString, Len(table77.Rows(i)("實際持有坪").ToString) - 1)
                    Else
                        未登記坪 = table77.Rows(i)("實際持有坪").ToString
                    End If

                    未登記坪 = "約" & 未登記坪 & "坪"
                Else
                    未登記坪 = ""
                End If

                '未登記所有權
                未登記所有權 = ""

                '未登記權利範圍
                Dim 權利範圍1分子 As String = ""
                Dim 權利範圍1分母 As String = ""
                Dim 權利範圍2分子 As String = ""
                Dim 權利範圍2分母 As String = ""

                '權利範圍1分子
                If Not IsDBNull(table77.Rows(i)("權利範圍1分子")) Then
                    If table77.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                        權利範圍1分子 = "1"
                    Else
                        權利範圍1分子 = table77.Rows(i)("權利範圍1分子").ToString
                    End If
                Else
                    權利範圍1分子 = "1"
                End If

                '權利範圍1分母
                If Not IsDBNull(table77.Rows(i)("權利範圍1分母")) Then
                    If table77.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                        權利範圍1分母 = "1"
                    Else
                        權利範圍1分母 = table77.Rows(i)("權利範圍1分母").ToString
                    End If
                Else
                    權利範圍1分母 = "1"
                End If

                '權利範圍2分子
                If Not IsDBNull(table77.Rows(i)("權利範圍2分子")) Then
                    If table77.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                        權利範圍2分子 = "1"
                    Else
                        權利範圍2分子 = table77.Rows(i)("權利範圍2分子").ToString
                    End If
                Else
                    權利範圍2分子 = "1"
                End If

                '權利範圍2分母
                If Not IsDBNull(table77.Rows(i)("權利範圍2分母")) Then
                    If table77.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                        權利範圍2分母 = "1"
                    Else
                        權利範圍2分母 = table77.Rows(i)("權利範圍2分母").ToString
                    End If
                Else
                    權利範圍2分母 = "1"
                End If

                未登記權利範圍 = CType(CType(權利範圍1分子, Integer) * CType(權利範圍2分子, Integer), String) & "/" & CType(CType(權利範圍1分母, Integer) * CType(權利範圍2分母, Integer), String)

                '先把讀出的xml複製一份，接著開始改
                Dim tempdata As String = myText_合法增建

                tempdata = tempdata.Replace("≠未登記項次", 未登記項次)
                tempdata = tempdata.Replace("≠未登記座落", 未登記座落)
                tempdata = tempdata.Replace("≠未登記完成日", 未登記完成日)
                tempdata = tempdata.Replace("≠未登記坪", 未登記坪)
                tempdata = tempdata.Replace("≠未登記所有權", 未登記所有權)
                tempdata = tempdata.Replace("≠未登記權利範圍", 未登記權利範圍)

                '改完加到最後的字串裡面  
                最後要替代掉的字串_合法增建.Append(tempdata)
            Next
        Else
            Dim tempdata As String = myText_合法增建
            tempdata = tempdata.Replace("≠未登記項次", "")
            tempdata = tempdata.Replace("≠未登記座落", "")
            tempdata = tempdata.Replace("≠未登記完成日", "")
            tempdata = tempdata.Replace("≠未登記坪", "")
            tempdata = tempdata.Replace("≠未登記所有權", "")
            tempdata = tempdata.Replace("≠未登記權利範圍", "")


            '改完加到最後的字串裡面  
            最後要替代掉的字串_合法增建.Append(tempdata)
        End If
        sFile = sFile.Replace("!重複合法增建", 最後要替代掉的字串_合法增建.ToString())

        '未登記坪細項小計
        '未登記坪和
        If Right(未登記坪和.ToString, 5) = ".0000" Then
            未登記坪和 = Int(未登記坪和.ToString)
        ElseIf 未登記坪和.ToString.LastIndexOf(".") > 0 And Right(未登記坪和.ToString, 3) = "000" Then
            未登記坪和 = Left(未登記坪和.ToString, Len(未登記坪和.ToString) - 3)
        ElseIf 未登記坪和.ToString.LastIndexOf(".") > 0 And Right(未登記坪和.ToString, 2) = "00" Then
            未登記坪和 = Left(未登記坪和.ToString, Len(未登記坪和.ToString) - 2)
        ElseIf 未登記坪和.ToString.LastIndexOf(".") > 0 And Right(未登記坪和.ToString, 1) = "0" Then
            未登記坪和 = Left(未登記坪和.ToString, Len(未登記坪和.ToString) - 1)
        Else
            未登記坪和 = 未登記坪和.ToString
        End If

        sFile = sFile.Replace("≠未登記坪和", 未登記坪和)

        '[END============================================產權調查-增建(合法增建)=============================================END]


        '共用
        '[START============================================產權調查-增建(違章建築)=============================================START]
        '違章建築細項
        Dim myXML_違章建築 As String = ""
        Dim myText_違章建築 As String = ""

        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        Dim srText_違章建築 As New StreamReader(Server.MapPath("..\reports\重複違章建築_V2.txt"))
        myText_違章建築 = srText_違章建築.ReadToEnd()
        srText_違章建築.Close()

        sqlstr = ""
        sqlstr = "Select * From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別='增建' and 項目名稱='違章建築'  order by 流水號 "


        Dim table88 As DataTable
        Dim 違建項次 As Integer
        Dim 違建座落 As String = ""
        Dim 違建完成日 As String = ""
        Dim 違建坪 As String = ""
        Dim 違建所有權 As String = ""
        Dim 違建權利範圍 As String = ""
        Dim 違建坪和 As Decimal = 0

        i = 0
        adpt = New SqlDataAdapter(sqlstr, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table88")
        table88 = ds.Tables("table88")
        Dim 最後要替代掉的字串_違章建築 As New StringBuilder()

        If table88.Rows.Count > 0 Then
            For i = 0 To table88.Rows.Count - 1

                '違建項次
                違建項次 = i + 1

                '違建座落
                If Not IsDBNull(table88.Rows(i)("建號")) Then
                    違建座落 = table88.Rows(i)("建號").ToString
                Else
                    違建座落 = ""
                End If

                '違建完成日
                If Not IsDBNull(table88.Rows(i)("增建完成日期")) Then
                    違建完成日 = table88.Rows(i)("增建完成日期").ToString
                Else
                    違建完成日 = ""
                End If

                '違建坪
                If Not IsDBNull(table88.Rows(i)("實際持有坪")) Then
                    '總和
                    違建坪和 += table88.Rows(i)("實際持有坪")

                    If Right(table88.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                        違建坪 = Int(table88.Rows(i)("實際持有坪").ToString)
                    ElseIf table88.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table88.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                        違建坪 = Left(table88.Rows(i)("實際持有坪").ToString, Len(table88.Rows(i)("實際持有坪").ToString) - 3)
                    ElseIf table88.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table88.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                        違建坪 = Left(table88.Rows(i)("實際持有坪").ToString, Len(table88.Rows(i)("實際持有坪").ToString) - 2)
                    ElseIf table88.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table88.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                        違建坪 = Left(table88.Rows(i)("實際持有坪").ToString, Len(table88.Rows(i)("實際持有坪").ToString) - 1)
                    Else
                        違建坪 = table88.Rows(i)("實際持有坪").ToString
                    End If

                    違建坪 = "約" & 違建坪 & "坪"
                Else
                    違建坪 = ""
                End If

                '違建所有權
                違建所有權 = ""

                '違建權利範圍
                Dim 權利範圍1分子 As String = ""
                Dim 權利範圍1分母 As String = ""
                Dim 權利範圍2分子 As String = ""
                Dim 權利範圍2分母 As String = ""

                '權利範圍1分子
                If Not IsDBNull(table88.Rows(i)("權利範圍1分子")) Then
                    If table88.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                        權利範圍1分子 = "1"
                    Else
                        權利範圍1分子 = table88.Rows(i)("權利範圍1分子").ToString
                    End If
                Else
                    權利範圍1分子 = "1"
                End If

                '權利範圍1分母
                If Not IsDBNull(table88.Rows(i)("權利範圍1分母")) Then
                    If table88.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                        權利範圍1分母 = "1"
                    Else
                        權利範圍1分母 = table88.Rows(i)("權利範圍1分母").ToString
                    End If
                Else
                    權利範圍1分母 = "1"
                End If

                '權利範圍2分子
                If Not IsDBNull(table88.Rows(i)("權利範圍2分子")) Then
                    If table88.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                        權利範圍2分子 = "1"
                    Else
                        權利範圍2分子 = table88.Rows(i)("權利範圍2分子").ToString
                    End If
                Else
                    權利範圍2分子 = "1"
                End If

                '權利範圍2分母
                If Not IsDBNull(table88.Rows(i)("權利範圍2分母")) Then
                    If table88.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                        權利範圍2分母 = "1"
                    Else
                        權利範圍2分母 = table88.Rows(i)("權利範圍2分母").ToString
                    End If
                Else
                    權利範圍2分母 = "1"
                End If

                違建權利範圍 = CType(CType(權利範圍1分子, Integer) * CType(權利範圍2分子, Integer), String) & "/" & CType(CType(權利範圍1分母, Integer) * CType(權利範圍2分母, Integer), String)

                '先把讀出的xml複製一份，接著開始改
                Dim tempdata As String = myText_違章建築

                tempdata = tempdata.Replace("≠違建項次", 違建項次)
                tempdata = tempdata.Replace("≠違建座落", 違建座落)
                tempdata = tempdata.Replace("≠違建完成日", 違建完成日)
                tempdata = tempdata.Replace("≠違建坪", 違建坪)
                tempdata = tempdata.Replace("≠違建所有權", 違建所有權)
                tempdata = tempdata.Replace("≠違建權利範圍", 違建權利範圍)

                '改完加到最後的字串裡面  
                最後要替代掉的字串_違章建築.Append(tempdata)
            Next
        Else
            Dim tempdata As String = myText_違章建築
            tempdata = tempdata.Replace("≠違建項次", "")
            tempdata = tempdata.Replace("≠違建座落", "")
            tempdata = tempdata.Replace("≠違建完成日", "")
            tempdata = tempdata.Replace("≠違建坪", "")
            tempdata = tempdata.Replace("≠違建所有權", "")
            tempdata = tempdata.Replace("≠違建權利範圍", "")


            '改完加到最後的字串裡面  
            最後要替代掉的字串_違章建築.Append(tempdata)
        End If
        sFile = sFile.Replace("!重複違章建築", 最後要替代掉的字串_違章建築.ToString())

        '違建坪細項小計
        '違建坪和
        If Right(違建坪和.ToString, 5) = ".0000" Then
            違建坪和 = Int(違建坪和.ToString)
        ElseIf 違建坪和.ToString.LastIndexOf(".") > 0 And Right(違建坪和.ToString, 3) = "000" Then
            違建坪和 = Left(違建坪和.ToString, Len(違建坪和.ToString) - 3)
        ElseIf 未登記坪和.ToString.LastIndexOf(".") > 0 And Right(未登記坪和.ToString, 2) = "00" Then
            違建坪和 = Left(違建坪和.ToString, Len(違建坪和.ToString) - 2)
        ElseIf 違建坪和.ToString.LastIndexOf(".") > 0 And Right(違建坪和.ToString, 1) = "0" Then
            違建坪和 = Left(違建坪和.ToString, Len(違建坪和.ToString) - 1)
        Else
            違建坪和 = 違建坪和.ToString
        End If

        sFile = sFile.Replace("≠違建坪和", 違建坪和)

        '[END============================================產權調查-增建(違章建築)=============================================END]


        '共用
        '[START============================================產權調查-增建(增加建)=============================================START]
        '增加建細項
        Dim myXML_增加建 As String = ""
        Dim myText_增加建 As String = ""

        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        Dim srText_增加建 As New StreamReader(Server.MapPath("..\reports\重複增加建_V2.txt"))
        myText_增加建 = srText_增加建.ReadToEnd()
        srText_增加建.Close()

        sqlstr = ""
        sqlstr = "Select * From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別='增建' and 項目名稱='增建'  order by 流水號 "


        Dim table99 As DataTable
        Dim 增建項次 As Integer
        Dim 增建位置 As String = ""
        Dim 增建用途 As String = ""
        Dim 增建坪 As String = ""
        Dim 增建所有權 As String = ""
        Dim 增建權利範圍 As String = ""
        Dim 增建坪和 As Decimal = 0

        i = 0
        adpt = New SqlDataAdapter(sqlstr, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table99")
        table99 = ds.Tables("table99")
        Dim 最後要替代掉的字串_增加建 As New StringBuilder()

        If table99.Rows.Count > 0 Then
            For i = 0 To table99.Rows.Count - 1

                '增建項次
                增建項次 = i + 1

                '增建位置
                If Not IsDBNull(table99.Rows(i)("建號")) Then
                    增建位置 = table99.Rows(i)("建號").ToString
                Else
                    增建位置 = ""
                End If

                '增建用途
                If Not IsDBNull(table99.Rows(i)("增建用途")) Then
                    增建用途 = table99.Rows(i)("增建用途").ToString
                Else
                    增建用途 = ""
                End If



                '增建坪
                If Not IsDBNull(table99.Rows(i)("實際持有坪")) Then
                    '總和
                    增建坪和 += table99.Rows(i)("實際持有坪")

                    If Right(table99.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                        增建坪 = Int(table99.Rows(i)("實際持有坪").ToString)
                    ElseIf table99.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table99.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                        增建坪 = Left(table99.Rows(i)("實際持有坪").ToString, Len(table99.Rows(i)("實際持有坪").ToString) - 3)
                    ElseIf table99.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table99.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                        增建坪 = Left(table99.Rows(i)("實際持有坪").ToString, Len(table99.Rows(i)("實際持有坪").ToString) - 2)
                    ElseIf table99.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table99.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                        增建坪 = Left(table99.Rows(i)("實際持有坪").ToString, Len(table99.Rows(i)("實際持有坪").ToString) - 1)
                    Else
                        增建坪 = table99.Rows(i)("實際持有坪").ToString
                    End If

                    增建坪 = "約" & 增建坪 & "坪"
                Else
                    增建坪 = ""
                End If

                '增建所有權
                增建所有權 = ""

                '增建權利範圍
                Dim 權利範圍1分子 As String = ""
                Dim 權利範圍1分母 As String = ""
                Dim 權利範圍2分子 As String = ""
                Dim 權利範圍2分母 As String = ""

                '權利範圍1分子
                If Not IsDBNull(table99.Rows(i)("權利範圍1分子")) Then
                    If table99.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                        權利範圍1分子 = "1"
                    Else
                        權利範圍1分子 = table99.Rows(i)("權利範圍1分子").ToString
                    End If
                Else
                    權利範圍1分子 = "1"
                End If

                '權利範圍1分母
                If Not IsDBNull(table99.Rows(i)("權利範圍1分母")) Then
                    If table99.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                        權利範圍1分母 = "1"
                    Else
                        權利範圍1分母 = table99.Rows(i)("權利範圍1分母").ToString
                    End If
                Else
                    權利範圍1分母 = "1"
                End If

                '權利範圍2分子
                If Not IsDBNull(table99.Rows(i)("權利範圍2分子")) Then
                    If table99.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                        權利範圍2分子 = "1"
                    Else
                        權利範圍2分子 = table99.Rows(i)("權利範圍2分子").ToString
                    End If
                Else
                    權利範圍2分子 = "1"
                End If

                '權利範圍2分母
                If Not IsDBNull(table99.Rows(i)("權利範圍2分母")) Then
                    If table99.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                        權利範圍2分母 = "1"
                    Else
                        權利範圍2分母 = table99.Rows(i)("權利範圍2分母").ToString
                    End If
                Else
                    權利範圍2分母 = "1"
                End If

                增建權利範圍 = CType(CType(權利範圍1分子, Integer) * CType(權利範圍2分子, Integer), String) & "/" & CType(CType(權利範圍1分母, Integer) * CType(權利範圍2分母, Integer), String)

                '先把讀出的xml複製一份，接著開始改
                Dim tempdata As String = myText_增加建

                tempdata = tempdata.Replace("≠增建項次", 增建項次)
                tempdata = tempdata.Replace("≠增建位置", 增建位置)
                tempdata = tempdata.Replace("≠增建用途", 增建用途)
                tempdata = tempdata.Replace("≠增建坪", 增建坪)
                tempdata = tempdata.Replace("≠增建所有權", 增建所有權)
                tempdata = tempdata.Replace("≠增建權利範圍", 增建權利範圍)

                '改完加到最後的字串裡面  
                最後要替代掉的字串_增加建.Append(tempdata)
            Next
        Else
            Dim tempdata As String = myText_增加建
            tempdata = tempdata.Replace("≠增建項次", "")
            tempdata = tempdata.Replace("≠增建位置", "")
            tempdata = tempdata.Replace("≠增建用途", "")
            tempdata = tempdata.Replace("≠增建坪", "")
            tempdata = tempdata.Replace("≠增建所有權", "")
            tempdata = tempdata.Replace("≠增建權利範圍", "")


            '改完加到最後的字串裡面  
            最後要替代掉的字串_增加建.Append(tempdata)
        End If
        sFile = sFile.Replace("!重複增加建", 最後要替代掉的字串_增加建.ToString())

        '增建坪細項小計
        '增建坪和
        If Right(增建坪和.ToString, 5) = ".0000" Then
            增建坪和 = Int(增建坪和.ToString)
        ElseIf 增建坪和.ToString.LastIndexOf(".") > 0 And Right(增建坪和.ToString, 3) = "000" Then
            增建坪和 = Left(增建坪和.ToString, Len(增建坪和.ToString) - 3)
        ElseIf 未登記坪和.ToString.LastIndexOf(".") > 0 And Right(未登記坪和.ToString, 2) = "00" Then
            增建坪和 = Left(增建坪和.ToString, Len(增建坪和.ToString) - 2)
        ElseIf 增建坪和.ToString.LastIndexOf(".") > 0 And Right(增建坪和.ToString, 1) = "0" Then
            增建坪和 = Left(增建坪和.ToString, Len(增建坪和.ToString) - 1)
        Else
            增建坪和 = 增建坪和.ToString
        End If

        sFile = sFile.Replace("≠增建坪和", 增建坪和)

        '[END============================================產權調查-增建(增加建)=============================================END]


        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            ''[START============================================產權調查-停車位產權說明=============================================START]
            'Dim 車獨立產權 As String = ""
            'Dim 車權利種類 As String = ""
            'Dim 車建號 As String = ""
            'Dim 車型式 As String = ""
            'Dim 車編號 As String = ""
            'Dim 車長 As String = ""
            'Dim 車寬 As String = ""
            'Dim 車高 As String = ""
            'Dim 車重 As String = ""


            ''建物目前管理狀況
            'Dim myXML_CarDetail As String = ""
            'Dim myText_CarDetail As String = ""

            ''讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            'Dim srText_CarDetail As New StreamReader(Server.MapPath("..\reports\重複車位細項_V2.txt"))
            'myText_CarDetail = srText_CarDetail.ReadToEnd()
            'srText_CarDetail.Close()

            'sqlstr = ""
            'sqlstr = "Select * From 產調_車位 A With(NoLock) Left Join 委賣物件資料表_面積細項 B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號  Where A.物件編號 = '" & Contract_No & "' and A.店代號='" & Request("sid") & "'  order by A.流水號 "


            'Dim table13 As DataTable

            'i = 0
            'adpt = New SqlDataAdapter(sqlstr, conn)
            'ds = New DataSet()
            'adpt.Fill(ds, "table13")
            'table13 = ds.Tables("table13")

            'Dim 最後要替代掉的字串_車位細項 As New StringBuilder()

            'If table13.Rows.Count > 0 Then
            '    For i = 0 To table13.Rows.Count - 1
            '        '車獨立產權
            '        If Not IsDBNull(table13.Rows(i)("獨立產權")) Then
            '            車獨立產權 = table13.Rows(i)("獨立產權").ToString
            '        End If
            '        '車權利種類
            '        If Not IsDBNull(table13.Rows(i)("權利種類")) Then
            '            車權利種類 = table13.Rows(i)("權利種類").ToString
            '        End If
            '        '車建號
            '        If Not IsDBNull(table13.Rows(i)("建號")) Then
            '            車建號 = table13.Rows(i)("建號").ToString
            '        End If
            '        '車型式
            '        If Not IsDBNull(table13.Rows(i)("車位類別")) Then
            '            車型式 = table13.Rows(i)("車位類別").ToString
            '        End If
            '        '車編號
            '        If Not IsDBNull(table13.Rows(i)("車位號碼")) Then
            '            車編號 = table13.Rows(i)("車位號碼").ToString
            '        End If
            '        '車長
            '        If Not IsDBNull(table13.Rows(i)("車位_長")) Then
            '            車長 = table13.Rows(i)("車位_長").ToString
            '        End If
            '        '車寬
            '        If Not IsDBNull(table13.Rows(i)("車位_寬")) Then
            '            車寬 = table13.Rows(i)("車位_寬").ToString
            '        End If
            '        '車高
            '        If Not IsDBNull(table13.Rows(i)("車位_高")) Then
            '            車高 = table13.Rows(i)("車位_高").ToString
            '        End If
            '        '車重
            '        If Not IsDBNull(table13.Rows(i)("車位_承重")) Then
            '            車重 = table13.Rows(i)("車位_承重").ToString
            '        End If




            '        '先把讀出的xml複製一份，接著開始改
            '        Dim tempdata As String = myText_CarDetail
            '        tempdata = tempdata.Replace("≠車獨立產權", 車獨立產權)
            '        tempdata = tempdata.Replace("≠車權利種類", 車權利種類)
            '        tempdata = tempdata.Replace("≠車建號", 車建號)
            '        tempdata = tempdata.Replace("≠車型式", 車型式)
            '        tempdata = tempdata.Replace("≠車編號", 車編號)
            '        tempdata = tempdata.Replace("≠車長", 車長)
            '        tempdata = tempdata.Replace("≠車寬", 車寬)
            '        tempdata = tempdata.Replace("≠車高", 車高)
            '        tempdata = tempdata.Replace("≠車重", 車重)





            '        '改完加到最後的字串裡面  
            '        最後要替代掉的字串_車位細項.Append(tempdata)

            '    Next
            'Else
            '    Dim tempdata As String = myText_CarDetail
            '    tempdata = tempdata.Replace("≠車獨立產權", "")
            '    tempdata = tempdata.Replace("≠車權利種類", "")
            '    tempdata = tempdata.Replace("≠車建號", "")
            '    tempdata = tempdata.Replace("≠車型式", "")
            '    tempdata = tempdata.Replace("≠車編號", "")
            '    tempdata = tempdata.Replace("≠車長", "")
            '    tempdata = tempdata.Replace("≠車寬", "")
            '    tempdata = tempdata.Replace("≠車高", "")
            '    tempdata = tempdata.Replace("≠車重", "")


            '    '改完加到最後的字串裡面  
            '    最後要替代掉的字串_車位細項.Append(tempdata)
            'End If

            'sFile = sFile.Replace("!重複車位細項", 最後要替代掉的字串_車位細項.ToString())

            ''[END============================================產權調查-停車位產權說明=============================================END]


        End If  '-------------------------------------------------------------------------判斷成屋還是土地用






        If Trim(t2.Rows(0)("物件類別").ToString) = "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-停車位產權說明=============================================START]
            Dim 土地地段地號 As String = ""
            Dim 土地管制 As String = ""
            Dim 土地_使用分區 As String = ""
            Dim 土地建蔽率 As String = ""
            Dim 土地容積率 As String = ""

            '土地使用管制內容
            Dim myXML_PlaceData As String = ""
            Dim myText_PlaceData As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_PlaceData As New StreamReader(Server.MapPath("..\reports\重複土地使用管制內容_V2.txt"))
            myText_PlaceData = srText_PlaceData.ReadToEnd()
            srText_PlaceData.Close()

            sqlstr = ""
            sqlstr = "Select * From 委賣物件資料表_面積細項 With(NoLock)  Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "'  order by 流水號 "


            Dim table57 As DataTable

            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table57")
            table57 = ds.Tables("table57")

            Dim 最後要替代掉的字串_土地使用管制內容 As New StringBuilder()

            If table57.Rows.Count > 0 Then
                For i = 0 To table57.Rows.Count - 1
                    '土地地段地號
                    If Not IsDBNull(table57.Rows(i)("建號")) Then
                        土地地段地號 = table57.Rows(i)("建號").ToString
                    End If
                    '土地管制
                    If Not IsDBNull(table57.Rows(i)("管制")) Then
                        土地管制 = table57.Rows(i)("管制").ToString
                    End If
                    '土地使用分區
                    If Not IsDBNull(table57.Rows(i)("使用分區")) Then
                        土地_使用分區 = table57.Rows(i)("使用分區").ToString
                    End If
                    '土地建蔽率
                    If Not IsDBNull(table57.Rows(i)("法定建蔽率")) Then
                        土地建蔽率 = table57.Rows(i)("法定建蔽率").ToString
                    End If
                    '土地容積率
                    If Not IsDBNull(table57.Rows(i)("法定容積率")) Then
                        土地容積率 = table57.Rows(i)("法定容積率").ToString
                    End If


                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_PlaceData
                    tempdata = tempdata.Replace("≠土地地段地號", 土地地段地號)
                    tempdata = tempdata.Replace("≠土地管制", 土地管制)
                    tempdata = tempdata.Replace("≠土地使用分區", 土地_使用分區)
                    tempdata = tempdata.Replace("≠土地建蔽率", 土地建蔽率)
                    tempdata = tempdata.Replace("≠土地容積率", 土地容積率)






                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_土地使用管制內容.Append(tempdata)

                Next
            Else
                Dim tempdata As String = myText_PlaceData
                tempdata = tempdata.Replace("≠土地地段地號", "")
                tempdata = tempdata.Replace("≠土地管制", "")
                tempdata = tempdata.Replace("≠土地使用分區", "")
                tempdata = tempdata.Replace("≠土地建蔽率", "")
                tempdata = tempdata.Replace("≠土地容積率", "")


                '改完加到最後的字串裡面  
                最後要替代掉的字串_土地使用管制內容.Append(tempdata)
            End If

            'If Request("oid") = "60692AAE34137" Then

            'Else
            sFile = sFile.Replace("!重複土地使用管制內容", 最後要替代掉的字串_土地使用管制內容.ToString())
            'End If

            '[END============================================產權調查-停車位產權說明=============================================END]


        End If  '-------------------------------------------------------------------------判斷成屋還是土地用


        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '基地面積細項
            Dim myXML_Place As String = ""
            Dim myText_Place As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_Place As New StreamReader(Server.MapPath("..\reports\重複基地面積細項_V2.txt"))
            myText_Place = srText_Place.ReadToEnd()
            srText_Place.Close()


            'sqlstr = "Select *,(權利範圍1分母+'分之'+權利範圍1分子) as 權利範圍,(權利範圍2分母+'分之'+權利範圍2分子) as 權利範圍2 From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別='土地面積'  order by 流水號 "
            '20160422 使用細項所有權人 by nick
            sqlstr = "Select (  a.權利範圍_分子  +'/'+ a.權利範圍_分母  ) as 權利範圍,( a.出售權利範圍_分子 +'/'+ a.出售權利範圍_分母  ) as 權利範圍2, (權利範圍1分子  +'/'+ 權利範圍1分母  ) as 權利範圍1_old, (權利範圍2分子  +'/'+ 權利範圍2分母  ) as 權利範圍2_old, * "
            sqlstr &= " From 委賣物件資料表_面積細項 b With(NoLock) "
            sqlstr &= " LEFT JOIN 委賣物件資料表_細項所有權人 a ON a.物件編號 = b.物件編號 AND a.店代號 = b.店代號 AND a.細項流水號 = b.流水號  "
            sqlstr &= " Where b.物件編號 = '" & Contract_No & "' and b.店代號='" & Request("sid") & "' and b.類別='土地面積'  order by b.流水號"


            Dim table2 As DataTable
            Dim 基地地號 As String = ""
            Dim 基地分區 As String = ""
            Dim 基地總面積 As String = ""
            Dim 基地權利範圍 As String = ""
            Dim 基地持份面積 As String = ""
            Dim 基地持份坪 As String = ""
            Dim 基地面積平方和 As Decimal = 0
            Dim 基地面積坪和 As Decimal = 0


            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            Dim 最後要替代掉的字串_基地面積細項 As New StringBuilder()

            If table2.Rows.Count > 0 Then
                For i = 0 To table2.Rows.Count - 1


                    '建地號
                    If Not IsDBNull(table2.Rows(i)("建號")) Then
                        基地地號 = table2.Rows(i)("建號").ToString
                    Else
                        基地地號 = ""
                    End If

                    '使用分區
                    If Not IsDBNull(table2.Rows(i)("使用分區")) Then
                        基地分區 = table2.Rows(i)("使用分區").ToString
                    Else
                        基地分區 = ""
                    End If



                    '總面積
                    If Not IsDBNull(table2.Rows(i)("總面積平方公尺")) Then
                        If Right(table2.Rows(i)("總面積平方公尺").ToString, 5) = ".0000" Then
                            基地總面積 = Int(table2.Rows(i)("總面積平方公尺").ToString)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 3) = "000" Then
                            基地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 3)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 2) = "00" Then
                            基地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 2)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 1) = "0" Then
                            基地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 1)
                        Else
                            基地總面積 = table2.Rows(i)("總面積平方公尺").ToString
                        End If
                    Else
                        基地總面積 = "0"
                    End If

                    '建物權利範圍
                    Dim 權利範圍1分子 As String = ""
                    Dim 權利範圍1分母 As String = ""
                    Dim 權利範圍2分子 As String = ""
                    Dim 權利範圍2分母 As String = ""

                    ''權利範圍1分子
                    'If Not IsDBNull(table2.Rows(i)("權利範圍1分子")) Then
                    '    If table2.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                    '        權利範圍1分子 = "1"
                    '    Else
                    '        權利範圍1分子 = table2.Rows(i)("權利範圍1分子").ToString
                    '    End If
                    'Else
                    '    權利範圍1分子 = "1"
                    'End If

                    ''權利範圍1分母
                    'If Not IsDBNull(table2.Rows(i)("權利範圍1分母")) Then
                    '    If table2.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                    '        權利範圍1分母 = "1"
                    '    Else
                    '        權利範圍1分母 = table2.Rows(i)("權利範圍1分母").ToString
                    '    End If
                    'Else
                    '    權利範圍1分母 = "1"
                    'End If

                    ''權利範圍2分子
                    'If Not IsDBNull(table2.Rows(i)("權利範圍2分子")) Then
                    '    If table2.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                    '        權利範圍2分子 = "1"
                    '    Else
                    '        權利範圍2分子 = table2.Rows(i)("權利範圍2分子").ToString
                    '    End If
                    'Else
                    '    權利範圍2分子 = "1"
                    'End If

                    ''權利範圍2分母
                    'If Not IsDBNull(table2.Rows(i)("權利範圍2分母")) Then
                    '    If table2.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                    '        權利範圍2分母 = "1"
                    '    Else
                    '        權利範圍2分母 = table2.Rows(i)("權利範圍2分母").ToString
                    '    End If
                    'Else
                    '    權利範圍2分母 = "1"
                    'End If

                    '基地權利範圍 = CType(CType(權利範圍1分子, Integer) * CType(權利範圍2分子, Integer), String) & "/" & CType(CType(權利範圍1分母, Integer) * CType(權利範圍2分母, Integer), String)

                    ' 2016.06.08 by Finch 如果"委賣物件資料表_細項所有權人"權利範圍沒有值，改從"委賣物件資料表_面積細項"裡抓舊版輸入的值
                    If Not IsDBNull(table2.Rows(i)("權利範圍")) Then
                        基地權利範圍 = table2.Rows(i)("權利範圍")
                    Else
                        If Not IsDBNull(table2.Rows(i)("權利範圍1_old")) Then
                            基地權利範圍 = table2.Rows(i)("權利範圍1_old")
                        Else
                            基地權利範圍 = "1/1"
                        End If
                    End If

                    '基地持份平方 改用 委賣物件資料表_細項所有權人. 持有面積 
                    'If Not IsDBNull(table2.Rows(i)("實際持有平方公尺")) Then

                    '    '總和
                    '    基地面積平方和 += table2.Rows(i)("實際持有平方公尺")

                    '    If Right(table2.Rows(i)("實際持有平方公尺").ToString, 5) = ".0000" Then
                    '        基地持份面積 = Int(table2.Rows(i)("實際持有平方公尺").ToString)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 3) = "000" Then
                    '        基地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 3)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 2) = "00" Then
                    '        基地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 2)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 1) = "0" Then
                    '        基地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 1)
                    '    Else
                    '        基地持份面積 = table2.Rows(i)("實際持有平方公尺").ToString
                    '    End If

                    'Else
                    '    基地持份面積 = ""
                    'End If

                    If Not IsDBNull(table2.Rows(i)("持有面積")) Then
                        '總和
                        基地面積平方和 += table2.Rows(i)("持有面積")
                        基地持份面積 = table2.Rows(i)("持有面積").ToString
                    ElseIf Not IsDBNull(table2.Rows(i)("實際持有平方公尺")) Then
                        基地面積平方和 += CDec(table2.Rows(i)("實際持有平方公尺"))
                        基地持份面積 = table2.Rows(i)("實際持有平方公尺")
                    Else
                        基地持份面積 = ""
                    End If



                    '土地持份坪
                    'If Not IsDBNull(table2.Rows(i)("實際持有坪")) Then
                    '    '總和
                    '    基地面積坪和 += table2.Rows(i)("實際持有坪")

                    '    If Right(table2.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                    '        基地持份坪 = Int(table2.Rows(i)("實際持有坪").ToString)
                    '    ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                    '        基地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 3)
                    '    ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                    '        基地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 2)
                    '    ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                    '        基地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 1)
                    '    Else
                    '        基地持份坪 = table2.Rows(i)("實際持有坪").ToString
                    '    End If
                    'Else
                    '    基地持份坪 = ""
                    'End If
                    基地面積坪和 = Math.Round(CDec(基地面積平方和) * 0.3025, 4)


                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_Place

                    tempdata = tempdata.Replace("≠基地地號", 基地地號)
                    tempdata = tempdata.Replace("≠基地分區", 基地分區)
                    tempdata = tempdata.Replace("≠基地總面積", 基地總面積)
                    tempdata = tempdata.Replace("≠基地權利範圍", 基地權利範圍)
                    tempdata = tempdata.Replace("≠基地持份面積", 基地持份面積)


                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_基地面積細項.Append(tempdata)

                Next
            Else
                Dim tempdata As String = myText_Place
                tempdata = tempdata.Replace("≠基地地號", "")
                tempdata = tempdata.Replace("≠基地分區", 基地分區)
                tempdata = tempdata.Replace("≠基地總面積", "")
                tempdata = tempdata.Replace("≠基地權利範圍", "")
                tempdata = tempdata.Replace("≠基地持份面積", "")


                '改完加到最後的字串裡面  
                最後要替代掉的字串_基地面積細項.Append(tempdata)
            End If
            sFile = sFile.Replace("!重複基地面積細項", 最後要替代掉的字串_基地面積細項.ToString())

            '基地細項小計
            '基地平方公尺
            If Right(基地面積平方和.ToString, 5) = ".0000" Then
                基地面積平方和 = Int(基地面積平方和.ToString)
            ElseIf 基地面積平方和.ToString.LastIndexOf(".") > 0 And Right(基地面積平方和.ToString, 3) = "000" Then
                基地面積平方和 = Left(基地面積平方和.ToString, Len(基地面積平方和.ToString) - 3)
            ElseIf 基地面積平方和.ToString.LastIndexOf(".") > 0 And Right(基地面積平方和.ToString, 2) = "00" Then
                基地面積平方和 = Left(基地面積平方和.ToString, Len(基地面積平方和.ToString) - 2)
            ElseIf 基地面積平方和.ToString.LastIndexOf(".") > 0 And Right(基地面積平方和.ToString, 1) = "0" Then
                基地面積平方和 = Left(基地面積平方和.ToString, Len(基地面積平方和.ToString) - 1)
            Else
                基地面積平方和 = 基地面積平方和.ToString
            End If
            '土地坪
            If Right(基地面積坪和.ToString, 5) = ".0000" Then
                基地面積坪和 = Int(基地面積坪和.ToString)
            ElseIf 基地面積坪和.ToString.LastIndexOf(".") > 0 And Right(基地面積坪和.ToString, 3) = "000" Then
                基地面積坪和 = Left(基地面積坪和.ToString, Len(基地面積坪和.ToString) - 3)
            ElseIf 基地面積坪和.ToString.LastIndexOf(".") > 0 And Right(基地面積坪和.ToString, 2) = "00" Then
                基地面積坪和 = Left(基地面積坪和.ToString, Len(基地面積坪和.ToString) - 2)
            ElseIf 基地面積坪和.ToString.LastIndexOf(".") > 0 And Right(基地面積坪和.ToString, 1) = "0" Then
                基地面積坪和 = Left(基地面積坪和.ToString, Len(基地面積坪和.ToString) - 1)
            Else
                基地面積坪和 = 基地面積坪和.ToString
            End If

            sFile = sFile.Replace("≠基地面積平方和", 基地面積平方和)
            sFile = sFile.Replace("≠基地面積坪和", 基地面積坪和)
        Else '土地
            '基地面積細項
            Dim myXML_Place As String = ""
            Dim myText_Place As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_Place As New StreamReader(Server.MapPath("..\reports\重複土地面積細項_V2.txt"))
            myText_Place = srText_Place.ReadToEnd()
            srText_Place.Close()


            'sqlstr = "Select *,(權利範圍1分母+'分之'+權利範圍1分子) as 權利範圍,(權利範圍2分母+'分之'+權利範圍2分子) as 權利範圍2 From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 類別='土地面積'  order by 流水號 "
            '20160422 使用細項所有權人 by nick
            sqlstr = "Select ( a.權利範圍_分母   +'分之'+ a.權利範圍_分子  ) as 權利範圍,(a.出售權利範圍_分母  +'分之'+ a.出售權利範圍_分子  ) as 權利範圍2, (權利範圍1分子  +'/'+ 權利範圍1分母  ) as 權利範圍1_old, (權利範圍2分子  +'/'+ 權利範圍2分母  ) as 權利範圍2_old, a.所有權人, b.所有權人 AS 所有權人_old, * "
            sqlstr &= " From 委賣物件資料表_面積細項 b With(NoLock) "
            sqlstr &= " LEFT JOIN 委賣物件資料表_細項所有權人 a ON a.物件編號 = b.物件編號 AND a.細項流水號 = b.流水號 and a.店代號=b.店代號 "
            sqlstr &= " Where b.物件編號 = '" & Contract_No & "' and b.店代號='" & Request("sid") & "' and b.類別='土地面積'  order by b.流水號"

            Dim table2 As DataTable
            Dim 土地地號 As String = ""
            Dim 土地總面積 As String = ""
            Dim 土地權利範圍 As String = ""
            Dim 土地持份面積 As String = ""
            Dim 土地持份坪 As String = ""
            Dim 土地權利人 As String = ""
            Dim 土地面積平方和 As Decimal = 0
            Dim 土地面積坪和 As Decimal = 0
            Dim 土地筆數 As Integer = 0

            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            Dim 最後要替代掉的字串_土地面積細項 As New StringBuilder()

            土地筆數 = table2.DefaultView.ToTable(True, "建號").Rows.Count
            If table2.Rows.Count > 0 Then
                For i = 0 To table2.Rows.Count - 1


                    '建地號
                    If Not IsDBNull(table2.Rows(i)("建號")) Then
                        土地地號 = table2.Rows(i)("建號").ToString
                    Else
                        土地地號 = ""
                    End If

                    '總面積
                    If Not IsDBNull(table2.Rows(i)("總面積平方公尺")) Then
                        '統計筆數
                        '土地筆數 += 1

                        If Right(table2.Rows(i)("總面積平方公尺").ToString, 5) = ".0000" Then
                            土地總面積 = Int(table2.Rows(i)("總面積平方公尺").ToString)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 3) = "000" Then
                            土地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 3)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 2) = "00" Then
                            土地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 2)
                        ElseIf table2.Rows(i)("總面積平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("總面積平方公尺").ToString, 1) = "0" Then
                            土地總面積 = Left(table2.Rows(i)("總面積平方公尺").ToString, Len(table2.Rows(i)("總面積平方公尺").ToString) - 1)
                        Else
                            土地總面積 = table2.Rows(i)("總面積平方公尺").ToString
                        End If
                    Else
                        土地總面積 = "0"
                    End If

                    '土地權利範圍1
                    Dim 權利範圍1分子 As String = ""
                    Dim 權利範圍1分母 As String = ""
                    Dim 權利範圍2分子 As String = ""
                    Dim 權利範圍2分母 As String = ""

                    '權利範圍1分子
                    If Not IsDBNull(table2.Rows(i)("權利範圍_分子")) Then
                        If table2.Rows(i)("權利範圍_分子") <> "0" Then
                            If table2.Rows(i)("權利範圍_分子").ToString.Trim = "" Then
                                權利範圍1分子 = "1"
                            Else
                                權利範圍1分子 = table2.Rows(i)("權利範圍_分子").ToString
                            End If
                        End If
                    ElseIf IsDBNull(table2.Rows(i)("權利範圍1分子")) Then
                        If table2.Rows(i)("權利範圍1分子") <> "0" Then
                            If table2.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                                權利範圍1分子 = "1"
                            Else
                                權利範圍1分子 = table2.Rows(i)("權利範圍1分子").ToString
                            End If
                        End If
                    Else
                        權利範圍1分子 = "1"
                    End If
                    'If Not IsDBNull(table2.Rows(i)("權利範圍1分子")) Then
                    '    If table2.Rows(i)("權利範圍1分子").ToString.Trim = "" Then
                    '        權利範圍1分子 = "1"
                    '    Else
                    '        權利範圍1分子 = table2.Rows(i)("權利範圍1分子").ToString
                    '    End If
                    'Else
                    '    權利範圍1分子 = "1"
                    'End If

                    '權利範圍1分母
                    If Not IsDBNull(table2.Rows(i)("權利範圍_分母")) Then
                        If table2.Rows(i)("權利範圍_分母") <> "0" Then
                            If table2.Rows(i)("權利範圍_分母").ToString.Trim = "" Then
                                權利範圍1分母 = "1"
                            Else
                                權利範圍1分母 = table2.Rows(i)("權利範圍_分母").ToString
                            End If
                        End If
                    ElseIf IsDBNull(table2.Rows(i)("權利範圍1分母")) Then
                        If table2.Rows(i)("權利範圍1分母") <> "0" Then
                            If table2.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                                權利範圍1分母 = "1"
                            Else
                                權利範圍1分母 = table2.Rows(i)("權利範圍1分母").ToString
                            End If
                        End If
                    Else
                        權利範圍1分母 = "1"
                    End If
                    'If Not IsDBNull(table2.Rows(i)("權利範圍1分母")) Then
                    '    If table2.Rows(i)("權利範圍1分母").ToString.Trim = "" Then
                    '        權利範圍1分母 = "1"
                    '    Else
                    '        權利範圍1分母 = table2.Rows(i)("權利範圍1分母").ToString
                    '    End If
                    'Else
                    '    權利範圍1分母 = "1"
                    'End If

                    '權利範圍2分子
                    If Not IsDBNull(table2.Rows(i)("出售權利範圍_分子")) Then
                        If table2.Rows(i)("出售權利範圍_分子").ToString.Trim = "" Then
                            權利範圍2分子 = "1"
                        Else
                            權利範圍2分子 = table2.Rows(i)("出售權利範圍_分子").ToString
                        End If
                    Else
                        權利範圍2分子 = "1"
                    End If
                    'If Not IsDBNull(table2.Rows(i)("權利範圍2分子")) Then
                    '    If table2.Rows(i)("權利範圍2分子").ToString.Trim = "" Then
                    '        權利範圍2分子 = "1"
                    '    Else
                    '        權利範圍2分子 = table2.Rows(i)("權利範圍2分子").ToString
                    '    End If
                    'Else
                    '    權利範圍2分子 = "1"
                    'End If

                    '權利範圍2分母
                    If Not IsDBNull(table2.Rows(i)("出售權利範圍_分母")) Then
                        If table2.Rows(i)("出售權利範圍_分母").ToString.Trim = "" Then
                            權利範圍2分母 = "1"
                        Else
                            權利範圍2分母 = table2.Rows(i)("出售權利範圍_分母").ToString
                        End If
                    Else
                        權利範圍2分母 = "1"
                    End If
                    'If Not IsDBNull(table2.Rows(i)("權利範圍2分母")) Then
                    '    If table2.Rows(i)("權利範圍2分母").ToString.Trim = "" Then
                    '        權利範圍2分母 = "1"
                    '    Else
                    '        權利範圍2分母 = table2.Rows(i)("權利範圍2分母").ToString
                    '    End If
                    'Else
                    '    權利範圍2分母 = "1"
                    'End If

                    '土地權利範圍 = CType(CType(權利範圍1分子, Long) * CType(權利範圍2分子, Long), String) & "/" & CType(CType(權利範圍1分母, Long) * CType(權利範圍2分母, Long), String)
                    土地權利範圍 = CType(CType(權利範圍1分子, Long), String) & "/" & CType(CType(權利範圍1分母, Long), String)



                    '基地持份平方
                    If Not IsDBNull(table2.Rows(i)("出售面積")) And table2.Rows(i)("出售面積").ToString <> "0.0000" And table2.Rows(i)("出售面積").ToString <> "" Then
                        '總和
                        土地面積平方和 += table2.Rows(i)("出售面積")

                        If Right(table2.Rows(i)("出售面積").ToString, 5) = ".0000" Then
                            土地持份面積 = Int(table2.Rows(i)("出售面積").ToString)
                        ElseIf table2.Rows(i)("出售面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("出售面積").ToString, 3) = "000" Then
                            土地持份面積 = Left(table2.Rows(i)("出售面積").ToString, Len(table2.Rows(i)("出售面積").ToString) - 3)
                        ElseIf table2.Rows(i)("出售面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("出售面積").ToString, 2) = "00" Then
                            土地持份面積 = Left(table2.Rows(i)("出售面積").ToString, Len(table2.Rows(i)("出售面積").ToString) - 2)
                        ElseIf table2.Rows(i)("出售面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("出售面積").ToString, 1) = "0" Then
                            土地持份面積 = Left(table2.Rows(i)("出售面積").ToString, Len(table2.Rows(i)("出售面積").ToString) - 1)
                        Else
                            土地持份面積 = table2.Rows(i)("出售面積").ToString
                        End If

                    ElseIf Not IsDBNull(table2.Rows(i)("持有面積")) And table2.Rows(i)("持有面積").ToString <> "0.0000" And table2.Rows(i)("持有面積").ToString <> "" Then
                        '總和
                        土地面積平方和 += table2.Rows(i)("持有面積")

                        If Right(table2.Rows(i)("持有面積").ToString, 5) = ".0000" Then
                            土地持份面積 = Int(table2.Rows(i)("持有面積").ToString)
                        ElseIf table2.Rows(i)("持有面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("持有面積").ToString, 3) = "000" Then
                            土地持份面積 = Left(table2.Rows(i)("持有面積").ToString, Len(table2.Rows(i)("持有面積").ToString) - 3)
                        ElseIf table2.Rows(i)("持有面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("持有面積").ToString, 2) = "00" Then
                            土地持份面積 = Left(table2.Rows(i)("持有面積").ToString, Len(table2.Rows(i)("持有面積").ToString) - 2)
                        ElseIf table2.Rows(i)("持有面積").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("持有面積").ToString, 1) = "0" Then
                            土地持份面積 = Left(table2.Rows(i)("持有面積").ToString, Len(table2.Rows(i)("持有面積").ToString) - 1)
                        Else
                            土地持份面積 = table2.Rows(i)("持有面積").ToString
                        End If

                    ElseIf Not IsDBNull(table2.Rows(i)("實際持有平方公尺")) Then
                        '總和
                        土地面積平方和 += table2.Rows(i)("實際持有平方公尺")

                        If Right(table2.Rows(i)("實際持有平方公尺").ToString, 5) = ".0000" Then
                            土地持份面積 = Int(table2.Rows(i)("實際持有平方公尺").ToString)
                        ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 3) = "000" Then
                            土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 3)
                        ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 2) = "00" Then
                            土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 2)
                        ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 1) = "0" Then
                            土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 1)
                        Else
                            土地持份面積 = table2.Rows(i)("實際持有平方公尺").ToString
                        End If
                    Else
                        土地持份面積 = ""
                    End If


                    'If Not IsDBNull(table2.Rows(i)("實際持有平方公尺")) Then

                    '    '總和
                    '    土地面積平方和 += table2.Rows(i)("實際持有平方公尺")

                    '    If Right(table2.Rows(i)("實際持有平方公尺").ToString, 5) = ".0000" Then
                    '        土地持份面積 = Int(table2.Rows(i)("實際持有平方公尺").ToString)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 3) = "000" Then
                    '        土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 3)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 2) = "00" Then
                    '        土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 2)
                    '    ElseIf table2.Rows(i)("實際持有平方公尺").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有平方公尺").ToString, 1) = "0" Then
                    '        土地持份面積 = Left(table2.Rows(i)("實際持有平方公尺").ToString, Len(table2.Rows(i)("實際持有平方公尺").ToString) - 1)
                    '    Else
                    '        土地持份面積 = table2.Rows(i)("實際持有平方公尺").ToString
                    '    End If

                    'Else
                    '    土地持份面積 = ""
                    'End If

                    '土地持份坪

                    If Not IsDBNull(table2.Rows(i)("出售面積")) And table2.Rows(i)("出售面積").ToString <> "0.0000" And table2.Rows(i)("出售面積").ToString <> "" Then
                        Dim 土地面積平方公尺轉坪 As Decimal = 0

                        '總和

                        Try
                            土地面積平方公尺轉坪 = System.Convert.ToDecimal(table2.Rows(i)("出售面積").ToString) * 0.3025
                            土地面積坪和 += 土地面積平方公尺轉坪
                        Catch exception As System.OverflowException
                            System.Console.WriteLine(
                                "Overflow in string-to-decimal conversion.")
                        Catch exception As System.FormatException
                            System.Console.WriteLine(
                                "The string is not formatted as a decimal.")
                        Catch exception As System.ArgumentException
                            System.Console.WriteLine("The string is null.")
                        End Try

                        If Right(土地面積平方公尺轉坪.ToString, 5) = ".0000" Then
                            土地持份坪 = Int(土地面積平方公尺轉坪)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 3) = "000" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 3)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 2) = "00" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 2)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 1) = "0" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 1)
                        Else
                            土地持份坪 = 土地面積平方公尺轉坪
                        End If

                    ElseIf Not IsDBNull(table2.Rows(i)("持有面積")) And table2.Rows(i)("持有面積").ToString <> "0.0000" And table2.Rows(i)("持有面積").ToString <> "" Then
                        Dim 土地面積平方公尺轉坪 As Decimal = 0

                        '總和

                        Try
                            土地面積平方公尺轉坪 = System.Convert.ToDecimal(table2.Rows(i)("持有面積").ToString) * 0.3025
                            土地面積坪和 += 土地面積平方公尺轉坪
                        Catch exception As System.OverflowException
                            System.Console.WriteLine(
                                "Overflow in string-to-decimal conversion.")
                        Catch exception As System.FormatException
                            System.Console.WriteLine(
                                "The string is not formatted as a decimal.")
                        Catch exception As System.ArgumentException
                            System.Console.WriteLine("The string is null.")
                        End Try

                        If Right(土地面積平方公尺轉坪.ToString, 5) = ".0000" Then
                            土地持份坪 = Int(土地面積平方公尺轉坪.ToString)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 3) = "000" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 3)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 2) = "00" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 2)
                        ElseIf 土地面積平方公尺轉坪.ToString.LastIndexOf(".") > 0 And Right(土地面積平方公尺轉坪.ToString, 1) = "0" Then
                            土地持份坪 = Left(土地面積平方公尺轉坪.ToString, Len(土地面積平方公尺轉坪.ToString) - 1)
                        Else
                            土地持份坪 = 土地面積平方公尺轉坪
                        End If

                    ElseIf Not IsDBNull(table2.Rows(i)("實際持有坪")) Then
                        '總和
                        土地面積坪和 += table2.Rows(i)("實際持有坪")

                        If Right(table2.Rows(i)("實際持有坪").ToString, 5) = ".0000" Then
                            土地持份坪 = Int(table2.Rows(i)("實際持有坪").ToString)
                        ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 3) = "000" Then
                            土地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 3)
                        ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 2) = "00" Then
                            土地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 2)
                        ElseIf table2.Rows(i)("實際持有坪").ToString.LastIndexOf(".") > 0 And Right(table2.Rows(i)("實際持有坪").ToString, 1) = "0" Then
                            土地持份坪 = Left(table2.Rows(i)("實際持有坪").ToString, Len(table2.Rows(i)("實際持有坪").ToString) - 1)
                        Else
                            土地持份坪 = table2.Rows(i)("實際持有坪").ToString
                        End If
                    Else
                        土地持份坪 = ""
                    End If

                    '土地權利人
                    If Not IsDBNull(table2.Rows(i)("所有權人")) And table2.Rows(i)("所有權人").ToString.Trim <> "" Then
                        土地權利人 = table2.Rows(i)("所有權人").ToString
                    ElseIf Not IsDBNull(table2.Rows(i)("所有權人_old")) And table2.Rows(i)("所有權人_old").ToString.Trim <> "" Then
                        土地權利人 = table2.Rows(i)("所有權人_old").ToString
                    End If

                    'If Request("oid") = "60692AAE34137" Then

                    'Else
                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_Place

                    tempdata = tempdata.Replace("≠土地地號", 土地地號)
                    tempdata = tempdata.Replace("≠土地總面積", 土地總面積)
                    tempdata = tempdata.Replace("≠土地權利範圍", 土地權利範圍)
                    tempdata = tempdata.Replace("≠土地持份面積", 土地持份面積)
                    tempdata = tempdata.Replace("≠土地持份坪", 土地持份坪)
                    tempdata = tempdata.Replace("≠土地權利人", 土地權利人)

                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_土地面積細項.Append(tempdata)
                    'End If
                Next
            Else
                Dim tempdata As String = myText_Place
                tempdata = tempdata.Replace("≠土地地號", "")
                tempdata = tempdata.Replace("≠土地總面積", "")
                tempdata = tempdata.Replace("≠土地權利範圍", "")
                tempdata = tempdata.Replace("≠土地持份面積", "")
                tempdata = tempdata.Replace("≠土地持份坪", "")
                tempdata = tempdata.Replace("≠土地權利人", "")


                '改完加到最後的字串裡面  
                最後要替代掉的字串_土地面積細項.Append(tempdata)
            End If
            sFile = sFile.Replace("!重複土地面積細項", 最後要替代掉的字串_土地面積細項.ToString())


            sFile = sFile.Replace("≠土地筆數", "共" & 土地筆數 & "筆")

            '基地細項小計
            '基地平方公尺
            If Right(土地面積平方和.ToString, 5) = ".0000" Then
                土地面積平方和 = Int(土地面積平方和.ToString)
            ElseIf 土地面積平方和.ToString.LastIndexOf(".") > 0 And Right(土地面積平方和.ToString, 3) = "000" Then
                土地面積平方和 = Left(土地面積平方和.ToString, Len(土地面積平方和.ToString) - 3)
            ElseIf 土地面積平方和.ToString.LastIndexOf(".") > 0 And Right(土地面積平方和.ToString, 2) = "00" Then
                土地面積平方和 = Left(土地面積平方和.ToString, Len(土地面積平方和.ToString) - 2)
            ElseIf 土地面積平方和.ToString.LastIndexOf(".") > 0 And Right(土地面積平方和.ToString, 1) = "0" Then
                土地面積平方和 = Left(土地面積平方和.ToString, Len(土地面積平方和.ToString) - 1)
            Else
                土地面積平方和 = 土地面積平方和.ToString
            End If
            '土地坪
            If Right(土地面積坪和.ToString, 5) = ".0000" Then
                土地面積坪和 = Int(土地面積坪和.ToString)
            ElseIf 土地面積坪和.ToString.LastIndexOf(".") > 0 And Right(土地面積坪和.ToString, 3) = "000" Then
                土地面積坪和 = Left(土地面積坪和.ToString, Len(土地面積坪和.ToString) - 3)
            ElseIf 土地面積坪和.ToString.LastIndexOf(".") > 0 And Right(土地面積坪和.ToString, 2) = "00" Then
                土地面積坪和 = Left(土地面積坪和.ToString, Len(土地面積坪和.ToString) - 2)
            ElseIf 土地面積坪和.ToString.LastIndexOf(".") > 0 And Right(土地面積坪和.ToString, 1) = "0" Then
                土地面積坪和 = Left(土地面積坪和.ToString, Len(土地面積坪和.ToString) - 1)
            Else
                土地面積坪和 = 土地面積坪和.ToString
            End If

            sFile = sFile.Replace("≠土地面積平方和", 土地面積平方和)
            sFile = sFile.Replace("≠土地面積坪和", 土地面積坪和)
        End If '-------------------------------------------------------------------------判斷成屋還是土地用
        '[START============================================產權調查-基地調查表P1+土地標示說明=============================================START]



        '共用
        '基地權利人
        Dim myXML_Bank As String = ""
        Dim myText_Bank As String = ""

        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        Dim srText_Bank As New StreamReader(Server.MapPath("..\reports\重複基地權利人_V2.txt"))
        myText_Bank = srText_Bank.ReadToEnd()
        srText_Bank.Close()


        sqlstr = "Select * From 物件他項權利細項  With(NoLock) Where 物件編號 = '" & Contract_No & "' and 店代號='" & Request("sid") & "' and 權利類別='土地'  order by 順位 asc,權利類別 desc  "

        Dim table3 As DataTable
        Dim 基地_權利種類 As String = ""
        Dim 基地登記日期 As String = ""
        Dim 基地設定性質設定金額 As String = ""
        Dim 基地權利人 As String = ""
        Dim 基地管理人 As String = ""
        Dim 處理方式 As String = ""
        i = 0
        adpt = New SqlDataAdapter(sqlstr, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table3")
        table3 = ds.Tables("table3")
        Dim 最後要替代掉的字串_基地權利人 As New StringBuilder()



        If t2.Rows(0)("與土地他項權利部相同").ToString = "1" Then
            '判斷基地他項無值時，才執行下段
            If table3.Rows.Count = 0 Then
                '基地權利種類
                基地_權利種類 = "與建物他項權利部同"
                '基地登記日期
                基地登記日期 = "-------"
                '基地設定性質設定金額
                基地設定性質設定金額 = "---------"
                '土地權利人
                基地權利人 = "---------------------"
                '基地管理人
                基地管理人 = "-------"


                '先把讀出的xml複製一份，接著開始改
                Dim tempdata As String = myText_Bank

                tempdata = tempdata.Replace("≠基地權利種類", 基地_權利種類)
                tempdata = tempdata.Replace("≠基地登記日期", 基地登記日期)
                tempdata = tempdata.Replace("≠基地設定性質設定金額", 基地設定性質設定金額)
                tempdata = tempdata.Replace("≠基地權利人", 基地權利人)
                tempdata = tempdata.Replace("≠基地管理人", 基地管理人)
                tempdata = tempdata.Replace("≠這裡放處理方式", "與建物他項權利部同")
                '改完加到最後的字串裡面  
                最後要替代掉的字串_基地權利人.Append(tempdata)
            Else


                For i As Integer = 0 To table3.Rows.Count - 1
                    處理方式 = ""
                    If Not IsDBNull(table3.Rows(i)("處理方式1")) Then
                        If table3.Rows(i)("處理方式1") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式1") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式2")) Then
                        If table3.Rows(i)("處理方式2") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式2") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式3")) Then
                        If table3.Rows(i)("處理方式3") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式3") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式4")) Then
                        If table3.Rows(i)("處理方式4") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式4") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式5")) Then
                        If table3.Rows(i)("處理方式5") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式5") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式6")) Then
                        If table3.Rows(i)("處理方式6") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式6") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("其他說明")) Then
                        If table3.Rows(i)("其他說明") <> "" Then
                            處理方式 &= table3.Rows(i)("其他說明")
                        End If
                    End If
                    If Right(處理方式, 1) = "," Then
                        處理方式 = 處理方式.Substring(0, 處理方式.Length - 1)
                    End If
                    '基地權利種類
                    If Not IsDBNull(table3.Rows(i)("權利種類")) Then
                        基地_權利種類 = table3.Rows(i)("權利種類")
                    Else
                        基地_權利種類 = ""
                    End If


                    '基地登記日期
                    If Not IsDBNull(table3.Rows(i)("登記日期")) Then
                        基地登記日期 = table3.Rows(i)("登記日期").ToString
                    Else
                        基地登記日期 = ""
                    End If

                    '基地設定性質設定金額
                    If Not IsDBNull(table3.Rows(i)("設定")) Then
                        基地設定性質設定金額 = table3.Rows(i)("設定").ToString & "萬"
                    Else
                        基地設定性質設定金額 = ""
                    End If


                    '基地權利人
                    If Not IsDBNull(table3.Rows(i)("設定權利人")) Then
                        基地權利人 = table3.Rows(i)("設定權利人").ToString
                    Else
                        基地權利人 = ""
                    End If

                    '基地管理人
                    If Not IsDBNull(table3.Rows(i)("管理人")) Then
                        基地管理人 = table3.Rows(i)("管理人").ToString
                    Else
                        基地管理人 = ""
                    End If



                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_Bank

                    tempdata = tempdata.Replace("≠基地權利種類", 基地_權利種類)
                    tempdata = tempdata.Replace("≠基地登記日期", 基地登記日期)
                    tempdata = tempdata.Replace("≠基地設定性質設定金額", 基地設定性質設定金額)
                    tempdata = tempdata.Replace("≠基地權利人", 基地權利人)
                    tempdata = tempdata.Replace("≠基地管理人", 基地管理人)
                    tempdata = tempdata.Replace("≠這裡放處理方式", 處理方式)
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_基地權利人.Append(tempdata)

                Next
            End If





            If t2.Rows(0)("其他如下").ToString = "1" Then
                Response.Write(t2.Rows(0)("其他如下").ToString)
                If table3.Rows.Count > 0 Then
                    '判斷是否勾選,建物與土地權利部同
                    i = 0
                    For i = 0 To table3.Rows.Count - 1
                        處理方式 = ""
                        If Not IsDBNull(table3.Rows(i)("處理方式1")) Then
                            If table3.Rows(i)("處理方式1") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式1") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("處理方式2")) Then
                            If table3.Rows(i)("處理方式2") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式2") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("處理方式3")) Then
                            If table3.Rows(i)("處理方式3") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式3") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("處理方式4")) Then
                            If table3.Rows(i)("處理方式4") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式4") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("處理方式5")) Then
                            If table3.Rows(i)("處理方式5") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式5") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("處理方式6")) Then
                            If table3.Rows(i)("處理方式6") <> "" Then
                                處理方式 &= table3.Rows(i)("處理方式6") & ","
                            End If
                        End If
                        If Not IsDBNull(table3.Rows(i)("其他說明")) Then
                            If table3.Rows(i)("其他說明") <> "" Then
                                處理方式 &= table3.Rows(i)("其他說明")
                            End If
                        End If
                        If Right(處理方式, 1) = "," Then
                            處理方式 = 處理方式.Substring(0, 處理方式.Length - 1)
                        End If

                        '基地權利種類
                        If Not IsDBNull(table3.Rows(i)("權利種類")) Then
                            基地_權利種類 = table3.Rows(i)("權利種類")
                        Else
                            基地_權利種類 = ""
                        End If

                        '基地登記日期
                        If Not IsDBNull(table3.Rows(i)("登記日期")) Then
                            基地登記日期 = table3.Rows(i)("登記日期").ToString
                        Else
                            基地登記日期 = ""
                        End If

                        '基地設定性質設定金額
                        If Not IsDBNull(table3.Rows(i)("設定")) Then
                            基地設定性質設定金額 = table3.Rows(i)("設定").ToString & "萬"
                        Else
                            基地設定性質設定金額 = ""
                        End If

                        '土地權利人
                        If Not IsDBNull(table3.Rows(i)("設定權利人")) Then
                            基地權利人 = table3.Rows(i)("設定權利人").ToString
                        Else
                            基地權利人 = ""
                        End If

                        '基地管理人
                        If Not IsDBNull(table3.Rows(i)("管理人")) Then
                            基地管理人 = table3.Rows(i)("管理人").ToString
                        Else
                            基地管理人 = ""
                        End If


                        '先把讀出的xml複製一份，接著開始改
                        Dim tempdata As String = myText_Bank

                        tempdata = tempdata.Replace("≠基地權利種類", 基地_權利種類)
                        tempdata = tempdata.Replace("≠基地登記日期", 基地登記日期)
                        tempdata = tempdata.Replace("≠基地設定性質設定金額", 基地設定性質設定金額)
                        tempdata = tempdata.Replace("≠基地權利人", 基地權利人)
                        tempdata = tempdata.Replace("≠基地管理人", 基地管理人)
                        tempdata = tempdata.Replace("≠這裡放處理方式", 處理方式)
                        '改完加到最後的字串裡面  
                        最後要替代掉的字串_基地權利人.Append(tempdata)

                    Next

                End If


            End If


        Else

            If table3.Rows.Count > 0 Then
                '判斷是否勾選,建物與土地權利部同

                For i = 0 To table3.Rows.Count - 1
                    處理方式 = ""
                    If Not IsDBNull(table3.Rows(i)("處理方式1")) Then
                        If table3.Rows(i)("處理方式1") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式1") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式2")) Then
                        If table3.Rows(i)("處理方式2") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式2") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式3")) Then
                        If table3.Rows(i)("處理方式3") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式3") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式4")) Then
                        If table3.Rows(i)("處理方式4") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式4") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式5")) Then
                        If table3.Rows(i)("處理方式5") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式5") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("處理方式6")) Then
                        If table3.Rows(i)("處理方式6") <> "" Then
                            處理方式 &= table3.Rows(i)("處理方式6") & ","
                        End If
                    End If
                    If Not IsDBNull(table3.Rows(i)("其他說明")) Then
                        If table3.Rows(i)("其他說明") <> "" Then
                            處理方式 &= table3.Rows(i)("其他說明")
                        End If
                    End If
                    If Right(處理方式, 1) = "," Then
                        處理方式 = 處理方式.Substring(0, 處理方式.Length - 1)
                    End If

                    '基地權利種類
                    If Not IsDBNull(table3.Rows(i)("權利種類")) Then
                        基地_權利種類 = table3.Rows(i)("權利種類")
                    Else
                        基地_權利種類 = ""
                    End If

                    '基地登記日期
                    If Not IsDBNull(table3.Rows(i)("登記日期")) Then
                        基地登記日期 = table3.Rows(i)("登記日期").ToString
                    Else
                        基地登記日期 = ""
                    End If

                    '基地設定性質設定金額
                    If Not IsDBNull(table3.Rows(i)("設定")) Then
                        基地設定性質設定金額 = table3.Rows(i)("設定").ToString & "萬"
                    Else
                        基地設定性質設定金額 = ""
                    End If


                    '基地權利人
                    If Not IsDBNull(table3.Rows(i)("設定權利人")) Then
                        基地權利人 = table3.Rows(i)("設定權利人").ToString
                    Else
                        基地權利人 = ""
                    End If

                    '基地管理人
                    If Not IsDBNull(table3.Rows(i)("管理人")) Then
                        基地管理人 = table3.Rows(i)("管理人").ToString
                    Else
                        基地管理人 = ""
                    End If



                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_Bank

                    tempdata = tempdata.Replace("≠基地權利種類", 基地_權利種類)
                    tempdata = tempdata.Replace("≠基地登記日期", 基地登記日期)
                    tempdata = tempdata.Replace("≠基地設定性質設定金額", 基地設定性質設定金額)
                    tempdata = tempdata.Replace("≠基地權利人", 基地權利人)
                    tempdata = tempdata.Replace("≠基地管理人", 基地管理人)
                    tempdata = tempdata.Replace("≠這裡放處理方式", 處理方式)
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_基地權利人.Append(tempdata)

                Next
            Else
                Dim tempdata As String = myText_Bank

                tempdata = tempdata.Replace("≠基地權利種類", "")
                tempdata = tempdata.Replace("≠基地登記日期", "")
                tempdata = tempdata.Replace("≠基地設定性質設定金額", "")
                tempdata = tempdata.Replace("≠基地權利人", "")
                tempdata = tempdata.Replace("≠基地管理人", "")
                tempdata = tempdata.Replace("≠這裡放處理方式", "")
                '改完加到最後的字串裡面  
                最後要替代掉的字串_基地權利人.Append(tempdata)
            End If
        End If
        'If Request("oid") = "60692AAE34137" Then

        'Else
        sFile = sFile.Replace("!重複基地權利人", 最後要替代掉的字串_基地權利人.ToString())
        'End If
        '[END============================================產權調查-基地調查表P1=============================================



        If Trim(t2.Rows(0)("物件類別").ToString) <> "土地" Then  '-------------------------------------------------------------------------判斷成屋還是土地用
            '[START============================================產權調查-基地目前管理狀況=============================================START]
            Dim 分管協議 As String = ""
            Dim 說明分管協議 As String = ""
            Dim 出租出借 As String = ""
            Dim 說明出租出借 As String = ""
            Dim 公眾通行私有道路 As String = ""
            Dim 說明公眾通行私有道路 As String = ""
            Dim 對外道路 As String = ""
            Dim 說明對外道路 As String = ""
            Dim 界址糾紛 As String = ""
            Dim 糾紛對象說明 As String = ""
            Dim 地籍圖重測 As String = ""
            Dim 標題主管機關已公告辦理 As String = ""
            Dim 主管機關已公告辦理 As String = ""
            Dim 公告徵收 As String = ""
            Dim 說明公告徵收範圍 As String = ""
            Dim 列管山坡地 As String = ""
            Dim 山坡地說明 As String = ""
            Dim 捷運系統穿越地 As String = ""
            Dim 捷運穿越地說明 As String = ""
            Dim 基地占用 As String = ""
            Dim 占用基地說明 As String = ""
            Dim 地號說明 As String = ""
            Dim 開發限制 As String = ""
            Dim 說明開發限制 As String = ""
            Dim 其他重要事項 As String = ""
            Dim 說明其他重要事項 As String = ""

            '建物目前管理狀況
            Dim myXML_PlaceManage As String = ""
            Dim myText_PlaceManage As String = ""

            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
            Dim srText_PlaceManage As New StreamReader(Server.MapPath("..\reports\重複基地管理狀況_V2.txt"))
            myText_PlaceManage = srText_PlaceManage.ReadToEnd()
            srText_PlaceManage.Close()

            sqlstr = ""
            sqlstr = "Select * From 產調_基地 A  With(NoLock) Left Join 委賣物件資料表_面積細項 B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號  Where A.物件編號 = '" & Contract_No & "' and A.店代號='" & Request("sid") & "'  order by A.流水號 "


            Dim table8 As DataTable

            i = 0
            adpt = New SqlDataAdapter(sqlstr, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table8")
            table8 = ds.Tables("table8")

            Dim 最後要替代掉的字串_基地管理狀況 As New StringBuilder()

            If table8.Rows.Count > 0 Then
                For i = 0 To table8.Rows.Count - 1
                    '地號層次說明
                    '地號說明 = "<w:br/>‎"
                    If Not IsDBNull(table8.Rows(i)("建號")) Then
                        If table8.Rows(i)("建號").ToString.Trim <> "" Then
                            地號說明 = "(地號：" & table8.Rows(i)("建號").ToString & ")"
                        End If
                    End If

                    '共有人分管協議或依民法第八百二十六條之ㄧ規定為使用管理或分割等約定之登記
                    If Not IsDBNull(table8.Rows(i)("共有土地有無分管協議")) Then
                        分管協議 = table8.Rows(i)("共有土地有無分管協議").ToString

                        If table8.Rows(i)("共有土地有無分管協議").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("共有土地有無分管協議_說明")) Then
                                If Trim(table8.Rows(i)("共有土地有無分管協議_說明").ToString) <> "" Then
                                    說明分管協議 = table8.Rows(i)("共有土地有無分管協議_說明").ToString
                                End If
                            End If
                        End If
                    Else
                        分管協議 = "無"
                    End If

                    '出租或出借
                    If Not IsDBNull(table8.Rows(i)("有無出租或出借")) Then
                        出租出借 = table8.Rows(i)("有無出租或出借").ToString

                        If table8.Rows(i)("有無出租或出借").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("有無出租或出借_說明")) Then
                                If Trim(table8.Rows(i)("有無出租或出借_說明").ToString) <> "" Then

                                    說明出租出借 = table8.Rows(i)("有無出租或出借_選項").ToString & "," & table8.Rows(i)("有無出租或出借_說明").ToString
                                End If
                            End If
                        End If
                    Else
                        出租出借 = "無"
                    End If
                    '開發限制
                    If Not IsDBNull(table8.Rows(i)("開發限制")) Then
                        開發限制 = table8.Rows(i)("開發限制").ToString

                        If table8.Rows(i)("開發限制").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("開發限制說明")) Then
                                If Trim(table8.Rows(i)("開發限制說明").ToString) <> "" Then
                                    說明開發限制 = table8.Rows(i)("開發限制說明").ToString
                                End If
                            End If
                        Else
                            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                            開發限制 = "無"
                            說明開發限制 = "本案已開發建築，若買方欲增建、改建時，仍須依都市計畫法、建築法規等相關規定辦理"
                            'End If
                        End If
                    Else
                        '開發限制 = "無"
                        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                        開發限制 = "無"
                        說明開發限制 = "本案已開發建築，若買方欲增建、改建時，仍須依都市計畫法、建築法規等相關規定辦理"
                        '    Else
                        '    開發限制 = "無"
                        'End If
                    End If
                    '其他重要事項

                    其他重要事項 = ""
                    說明其他重要事項 = ""

                    If Not IsDBNull(table8.Rows(i)("其他重要事項")) Then
                        其他重要事項 = table8.Rows(i)("其他重要事項").ToString

                        If table8.Rows(i)("其他重要事項").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("其他重要事項說明")) Then
                                If Trim(table8.Rows(i)("其他重要事項說明").ToString) <> "" Then
                                    說明其他重要事項 = table8.Rows(i)("其他重要事項說明").ToString
                                End If
                            End If
                        End If
                    Else
                        其他重要事項 = "無"
                        說明其他重要事項 = ""
                    End If


                    '公眾通行私有道路
                    If Not IsDBNull(table8.Rows(i)("供公眾通行之私有道路")) Then
                        公眾通行私有道路 = table8.Rows(i)("供公眾通行之私有道路").ToString

                        If table8.Rows(i)("供公眾通行之私有道路").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("供公眾通行之私有道路_說明")) Then
                                If Trim(table8.Rows(i)("供公眾通行之私有道路_說明").ToString) <> "" Then
                                    說明公眾通行私有道路 = table8.Rows(i)("供公眾通行之私有道路_說明").ToString
                                End If
                            End If

                        End If
                    Else
                        公眾通行私有道路 = "無"
                    End If

                    '對外道路
                    If Not IsDBNull(table8.Rows(i)("可通行對外道路")) Then
                        對外道路 = table8.Rows(i)("可通行對外道路").ToString

                        If table8.Rows(i)("可通行對外道路").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("可通行對外道路_說明")) Then
                                If Trim(table8.Rows(i)("可通行對外道路_說明").ToString) <> "" Then
                                    說明對外道路 = table8.Rows(i)("可通行對外道路_說明").ToString
                                End If
                            End If
                        End If
                    Else
                        對外道路 = "無"
                    End If

                    '界址糾紛
                    If Not IsDBNull(table8.Rows(i)("界址糾紛")) Then
                        界址糾紛 = table8.Rows(i)("界址糾紛").ToString

                        If table8.Rows(i)("界址糾紛").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("糾紛對象說明")) Then
                                If Trim(table8.Rows(i)("糾紛對象說明").ToString) <> "" Then
                                    糾紛對象說明 = table8.Rows(i)("糾紛對象說明").ToString
                                End If
                            End If

                        End If
                    Else
                        界址糾紛 = "無"
                    End If

                    '地籍圖重測
                    If Not IsDBNull(table8.Rows(i)("辦理地籍圖重測")) Then
                        If table8.Rows(i)("辦理地籍圖重測").ToString <> "" Then
                            地籍圖重測 = table8.Rows(i)("辦理地籍圖重測").ToString

                            If table8.Rows(i)("辦理地籍圖重測").ToString = "有" Then
                                標題主管機關已公告辦理 = ""
                                主管機關已公告辦理 = ""

                            Else
                                標題主管機關已公告辦理 = "主管機關已公告辦理"
                                '主管機關已公告辦理
                                If Not IsDBNull(table8.Rows(i)("主管機關已公告辦理")) Then
                                    If table8.Rows(i)("主管機關已公告辦理").ToString <> "" Then
                                        主管機關已公告辦理 = table8.Rows(0)("主管機關已公告辦理").ToString
                                    End If
                                End If
                            End If
                        Else
                            地籍圖重測 = "有"
                        End If
                    Else
                        地籍圖重測 = "有"
                    End If

                    '公告徵收
                    If Not IsDBNull(table8.Rows(i)("公告徵收")) Then
                        公告徵收 = table8.Rows(i)("公告徵收").ToString

                        If table8.Rows(i)("公告徵收").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("公告徵收範圍")) Then
                                If Trim(table8.Rows(i)("公告徵收範圍").ToString) <> "" Then
                                    說明公告徵收範圍 = table8.Rows(i)("公告徵收範圍").ToString
                                End If
                            End If

                        End If
                    Else
                        公告徵收 = "無"
                    End If

                    '山坡地
                    If Not IsDBNull(table8.Rows(i)("列管山坡地")) Then
                        列管山坡地 = table8.Rows(i)("列管山坡地").ToString

                        If table8.Rows(i)("列管山坡地").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("列管山坡地說明")) Then
                                If Trim(table8.Rows(i)("列管山坡地說明").ToString) <> "" Then
                                    山坡地說明 = table8.Rows(i)("列管山坡地說明").ToString
                                End If
                            End If

                        End If
                    Else
                        列管山坡地 = "無"
                        山坡地說明 = ""
                    End If

                    ' 捷運系統穿越地 捷運穿越地說明 
                    If Not IsDBNull(table8.Rows(i)("捷運穿越")) Then
                        捷運系統穿越地 = table8.Rows(i)("捷運穿越").ToString

                        If table8.Rows(i)("捷運穿越").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("捷運穿越說明")) Then
                                If Trim(table8.Rows(i)("捷運穿越說明").ToString) <> "" Then
                                    捷運穿越地說明 = table8.Rows(i)("捷運穿越說明").ToString
                                End If
                            End If

                        End If
                    Else
                        捷運系統穿越地 = "無"
                        捷運穿越地說明 = ""
                    End If
                    '基地占用 基地占用說明
                    If Not IsDBNull(table8.Rows(i)("基地占用")) Then
                        基地占用 = table8.Rows(i)("基地占用").ToString

                        If table8.Rows(i)("基地占用").ToString = "有" Then
                            If Not IsDBNull(table8.Rows(i)("基地占用說明")) Then
                                If Trim(table8.Rows(i)("基地占用說明").ToString) <> "" Then
                                    占用基地說明 = table8.Rows(i)("基地占用說明").ToString
                                End If
                            End If

                        End If
                    Else
                        基地占用 = "無"
                        占用基地說明 = ""
                    End If


                    '先把讀出的xml複製一份，接著開始改
                    Dim tempdata As String = myText_PlaceManage
                    tempdata = tempdata.Replace("≠地號說明", 地號說明)
                    tempdata = tempdata.Replace("≠分管協議", 分管協議)
                    tempdata = tempdata.Replace("≠說明分管協議", 說明分管協議)
                    tempdata = tempdata.Replace("≠出租出借", 出租出借)
                    tempdata = tempdata.Replace("≠說明出租出借", 說明出租出借)
                    tempdata = tempdata.Replace("≠公眾通行私有道路", 公眾通行私有道路)
                    tempdata = tempdata.Replace("≠說明公眾通行私有道路", 說明公眾通行私有道路)
                    tempdata = tempdata.Replace("≠對外道路", 對外道路)
                    tempdata = tempdata.Replace("≠說明對外道路", 說明對外道路)
                    tempdata = tempdata.Replace("≠界址糾紛", 界址糾紛)
                    tempdata = tempdata.Replace("≠糾紛對象說明", 糾紛對象說明)
                    tempdata = tempdata.Replace("≠地籍圖重測", 地籍圖重測)
                    tempdata = tempdata.Replace("≠標題主管機關已公告辦理", 標題主管機關已公告辦理)
                    tempdata = tempdata.Replace("≠主管機關已公告辦理", 主管機關已公告辦理)
                    tempdata = tempdata.Replace("≠公告徵收", 公告徵收)
                    tempdata = tempdata.Replace("≠說明公告徵收範圍", 說明公告徵收範圍)
                    tempdata = tempdata.Replace("≠列管山坡地", 列管山坡地)
                    tempdata = tempdata.Replace("≠山坡地說明", 山坡地說明)
                    tempdata = tempdata.Replace("≠捷運系統穿越地", 捷運系統穿越地)
                    tempdata = tempdata.Replace("≠系統穿越地說明", 捷運穿越地說明)
                    tempdata = tempdata.Replace("≠基地占用", 基地占用)
                    tempdata = tempdata.Replace("≠占用說明", 占用基地說明)
                    tempdata = tempdata.Replace("≠開發限制", 開發限制)
                    tempdata = tempdata.Replace("≠說明開發限制", 說明開發限制)
                    tempdata = tempdata.Replace("≠其他重要事項", 其他重要事項)
                    tempdata = tempdata.Replace("≠說明其他重要事項", 說明其他重要事項)
                    '改完加到最後的字串裡面  
                    最後要替代掉的字串_基地管理狀況.Append(tempdata)
                Next
            Else
                Dim tempdata As String = myText_PlaceManage
                tempdata = tempdata.Replace("≠地號說明", "")
                tempdata = tempdata.Replace("≠分管協議", "")
                tempdata = tempdata.Replace("≠說明分管協議", "")
                tempdata = tempdata.Replace("≠出租出借", "")
                tempdata = tempdata.Replace("≠說明出租出借", "")
                tempdata = tempdata.Replace("≠公眾通行私有道路", "")
                tempdata = tempdata.Replace("≠說明公眾通行私有道路", "")
                tempdata = tempdata.Replace("≠對外道路", "")
                tempdata = tempdata.Replace("≠說明對外道路", "")
                tempdata = tempdata.Replace("≠界址糾紛", "")
                tempdata = tempdata.Replace("≠糾紛對象說明", "")
                tempdata = tempdata.Replace("≠地籍圖重測", "")
                tempdata = tempdata.Replace("≠標題主管機關已公告辦理", "")
                tempdata = tempdata.Replace("≠主管機關已公告辦理", "")
                tempdata = tempdata.Replace("≠公告徵收", "")
                tempdata = tempdata.Replace("≠說明公告徵收範圍", "")
                tempdata = tempdata.Replace("≠列管山坡地", "")
                tempdata = tempdata.Replace("≠山坡地說明", "")
                tempdata = tempdata.Replace("≠捷運系統穿越地", "")
                tempdata = tempdata.Replace("≠系統穿越地說明", "")
                tempdata = tempdata.Replace("≠基地占用", "")
                tempdata = tempdata.Replace("≠占用說明", "")
                tempdata = tempdata.Replace("≠開發限制", "")
                tempdata = tempdata.Replace("≠說明開發限制", "")
                tempdata = tempdata.Replace("≠其他重要事項", "")
                tempdata = tempdata.Replace("≠說明其他重要事項", "")
                '改完加到最後的字串裡面  
                最後要替代掉的字串_基地管理狀況.Append(tempdata)
            End If

            sFile = sFile.Replace("!重複基地管理狀況", 最後要替代掉的字串_基地管理狀況.ToString())

            '[END============================================產權調查-基地目前管理狀況=============================================END]


        End If '-------------------------------------------------------------------------判斷成屋還是土地用





        'If Request.Cookies("webfly_empno").Value.ToUpper <> "H2L" Then
        '    '共用
        '    '[START============================================品牌成交行情=============================================START]
        '    '成交行情
        '    Dim myXML_ET As String = ""
        '    Dim myText_ET As String = ""

        '    '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        '    Dim srText_ET As New StreamReader(Server.MapPath("..\reports\重複品牌成交行情_V2.txt"))
        '    myText_ET = srText_ET.ReadToEnd()
        '    srText_ET.Close()



        '    '以報告日期計算前六個月
        '    Dim flage As String = ""
        '    If IsDBNull(t2.Rows(0)("經度").ToString) Or t2.Rows(0)("經度").ToString.Trim = "" Or IsDBNull(t2.Rows(0)("緯度").ToString) Or t2.Rows(0)("緯度").ToString.Trim = "" Then
        '        flage = "0"
        '    Else
        '        flage = "1"
        '    End If


        '    If flage = "1" Then

        '        Dim tempday = (DateAdd("m", -6, (Left(t2.Rows(0)("報告日期").ToString, 3) + 1911) & "/" & (Mid(t2.Rows(0)("報告日期").ToString, 4, 2)) & "/" + Right(t2.Rows(0)("報告日期").ToString, 2)))
        '        Dim startday = Right("000" & Year(tempday) - 1911, 3) & Right("00" & Month(tempday), 2) & Right("00" & Day(tempday), 2)

        '        Dim tempday2 = (DateAdd("m", -1, (Left(t2.Rows(0)("報告日期").ToString, 3) + 1911) & "/" & (Mid(t2.Rows(0)("報告日期").ToString, 4, 2)) & "/" + Right(t2.Rows(0)("報告日期").ToString, 2)))
        '        Dim endday = Right("000" & Year(tempday2) - 1911, 3) & Right("00" & Month(tempday2), 2) & Right("00" & Day(tempday2), 2)

        '        If DropDownList1.SelectedValue = "依路名" Then

        '            sqlstr = "select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr += " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr += " 賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And 賣.縣市 = '" & stree1.Text & "' "
        '            sqlstr &= " And 賣.鄉鎮市區 = '" & stree2.Text & "' "

        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If

        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)

        '                '近一個月不撈
        '                If endday >= Right("000" & Year(tempday2) - 1911, 3) & Right("00" & Month(tempday2), 2) & Right("00" & Day(tempday2), 2) Then
        '                    endday = Right("000" & Year(tempday2) - 1911, 3) & Right("00" & Month(tempday2), 2) & Right("00" & Day(tempday2), 2)
        '                End If

        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            '路段名
        '            If TextBox26.Text.Trim <> "" Then
        '                sqlstr &= " and (部份地址 like '%" & TextBox26.Text & "%'"
        '            End If

        '            If TextBox27.Text.Trim <> "" Then
        '                sqlstr &= " or 部份地址 like '%" & TextBox27.Text & "%'"
        '            End If

        '            If TextBox28.Text.Trim <> "" Then
        '                sqlstr &= " or 部份地址 like '%" & TextBox28.Text & "%'"
        '            End If

        '            If TextBox26.Text.Trim <> "" Then
        '                sqlstr &= ") "
        '            End If

        '            sqlstr &= " union all "

        '            sqlstr &= " select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr += " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr += "賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件過期資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And 賣.縣市 = '" & stree1.Text & "' "
        '            sqlstr &= " And 賣.鄉鎮市區 = '" & stree2.Text & "' "
        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If


        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = TextBox2.Text 'Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)

        '                '近一個月不撈
        '                If endday >= Right("000" & Year(tempday2) - 1911, 3) & Right("00" & Month(tempday2), 2) & Right("00" & Day(tempday2), 2) Then
        '                    endday = Right("000" & Year(tempday2) - 1911, 3) & Right("00" & Month(tempday2), 2) & Right("00" & Day(tempday2), 2)
        '                End If
        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            '路段名
        '            If TextBox26.Text.Trim <> "" Then
        '                sqlstr &= " and (部份地址 like '%" & TextBox26.Text & "%'"
        '            End If

        '            If TextBox27.Text.Trim <> "" Then
        '                sqlstr &= " or 部份地址 like '%" & TextBox27.Text & "%'"
        '            End If

        '            If TextBox28.Text.Trim <> "" Then
        '                sqlstr &= " or 部份地址 like '%" & TextBox28.Text & "%'"
        '            End If

        '            If TextBox26.Text.Trim <> "" Then
        '                sqlstr &= ") "
        '            End If



        '        Else '依範圍
        '            sqlstr = "select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr += " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr += " 賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And (6371 * acos( cos( radians(" & t2.Rows(0)("經度").ToString & ") ) * cos( radians( 經度 ) ) * cos( radians( 緯度 ) - radians(" & t2.Rows(0)("緯度").ToString & ") ) + sin( radians(" & t2.Rows(0)("經度").ToString & ") ) * sin( radians( 經度 ) ) ) )  < 0.5 "
        '            sqlstr &= " And 經度 <> " & t2.Rows(0)("經度").ToString & " And 緯度 <> " & t2.Rows(0)("緯度").ToString
        '            sqlstr &= " And 成.店代號<>'A0001'"
        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If

        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = TextBox2.Text 'Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)
        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            sqlstr &= " union all "

        '            sqlstr &= " select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr &= " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr &= " 賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件過期資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And (6371 * acos( cos( radians(" & t2.Rows(0)("經度").ToString & ") ) * cos( radians( 經度 ) ) * cos( radians( 緯度 ) - radians(" & t2.Rows(0)("緯度").ToString & ") ) + sin( radians(" & t2.Rows(0)("經度").ToString & ") ) * sin( radians( 經度 ) ) ) )  < 0.5 "
        '            sqlstr &= " And 經度 <> " & t2.Rows(0)("經度").ToString & " And 緯度 <> " & t2.Rows(0)("緯度").ToString

        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If


        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = TextBox2.Text 'Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = TextBox6.Text 'Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)
        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            sqlstr &= " union all "

        '            sqlstr &= " select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr &= " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr &= " 賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And 經度 = " & t2.Rows(0)("經度").ToString & " And 緯度 = " & t2.Rows(0)("緯度").ToString

        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If

        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = TextBox2.Text 'Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)
        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            sqlstr &= " union all "

        '            sqlstr &= " select DISTINCT(成.成交編號),賣.物件編號,賣.縣市 + 賣.鄉鎮市區 + 賣.部份地址 as 物件_部份地址,賣.物件用途, 賣.物件類別, 賣.竣工日期, 賣.所在樓層, "
        '            sqlstr &= " (case when 賣.物件類別 in ('土地','農地','建地') then ROUND(賣.土地平方公尺,2, 1) else ROUND(賣.總平方公尺,2, 1) end) as 總平方公尺, "
        '            sqlstr &= " 賣.車位,成.成交金額, 成.成交日期,賣.經度, 賣.緯度 "
        '            sqlstr &= " FROM 委賣成交資料表 as 成  With(NoLock) "
        '            sqlstr &= " LEFT JOIN 委賣物件過期資料表 as 賣 ON 賣.物件編號 = 成.物件編號 where  成交原因='買賣' "
        '            sqlstr &= " And 經度 = " & t2.Rows(0)("經度").ToString & " And 緯度 = " & t2.Rows(0)("緯度").ToString

        '            If InStr(Request("sid").ToUpper, "S") = 0 Then
        '                sqlstr &= " and 賣.店代號 like 'A%'"
        '            Else
        '                sqlstr &= " and 賣.店代號 like 'S%'"
        '            End If

        '            If DropDownList3.SelectedValue <> "請選擇" Then
        '                sqlstr &= " and 賣.物件類別 like '%" & DropDownList3.SelectedValue & "%' "
        '            End If


        '            '成交日期
        '            If TextBox2.Text <> "" And TextBox6.Text <> "" Then
        '                startday = TextBox2.Text 'Left(Trim(TextBox2.Text).PadLeft(7, "0"), 5)
        '                endday = TextBox6.Text 'Left(Trim(TextBox6.Text).PadLeft(7, "0"), 5)
        '            End If
        '            sqlstr &= " And 成交日期 between '" & startday & "' and '" & endday & "' "

        '            '成交金額
        '            If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
        '                Dim b1 As Double = 0, b2 As Double = 99999999
        '                'If TextBox1.Text <> "" Then b1 = TextBox1.Text
        '                'If TextBox3.Text <> "" Then b2 = TextBox3.Text
        '                b1 = Convert.ToDouble(TextBox1.Text)
        '                b2 = Convert.ToDouble(TextBox3.Text)
        '                sqlstr &= " And 成交金額 between '" & b1 & "' and '" & b2 & "' "
        '            End If

        '            '坪數
        '            If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
        '                Dim b1 As Single = 0, b2 As Single = 99999999
        '                If TextBox4.Text <> "" Then b1 = TextBox4.Text
        '                If TextBox5.Text <> "" Then b2 = TextBox5.Text
        '                sqlstr &= " And 總平方公尺 between '" & b1 & "' and '" & b2 & "' "
        '            End If
        '        End If

        '        sqlstr &= " order by 成.成交日期 desc"

        '        'If Request("oid") = "61273AAC83599" Then
        '        '    Response.Write(sqlstr & "<br>")
        '        '    Response.End()
        '        'End If

        '        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        '        '    Response.Write(sqlstr & "<br>")
        '        '    Response.End()
        '        'End If

        '        Dim table6 As DataTable
        '        Dim 成交座落 As String
        '        Dim 用途 As String
        '        Dim 類別 As String
        '        Dim 竣工日期 As String
        '        Dim 樓層 As String
        '        Dim 面積 As String
        '        Dim 成交車位 As String
        '        Dim 成交 As String
        '        Dim 日期成交 As String
        '        i = 0
        '        adpt = New SqlDataAdapter(sqlstr, conn)
        '        ds = New DataSet()
        '        adpt.Fill(ds, "table6")
        '        table6 = ds.Tables("table6")
        '        Dim 最後要替代掉的字串 As New StringBuilder()

        '        If table6.Rows.Count > 0 Then
        '            i = 0
        '            For i = 0 To table6.Rows.Count - 1
        '                If Not IsDBNull(table6.Rows(i)("物件_部份地址")) Then
        '                    成交座落 = table6.Rows(i)("物件_部份地址").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("物件用途")) Then
        '                    用途 = table6.Rows(i)("物件用途").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("物件類別")) Then
        '                    類別 = table6.Rows(i)("物件類別").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("竣工日期")) Then
        '                    竣工日期 = table6.Rows(i)("竣工日期").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("所在樓層")) Then
        '                    樓層 = table6.Rows(i)("所在樓層").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("總平方公尺")) Then
        '                    面積 = Math.Round(Convert.ToDouble("0" & table6.Rows(i)("總平方公尺").ToString), 2, MidpointRounding.AwayFromZero)
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("車位")) Then
        '                    成交車位 = table6.Rows(i)("車位").ToString
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("成交金額")) Then
        '                    成交 = String.Format("{0:#,###}", table6.Rows(i)("成交金額"))
        '                End If

        '                If Not IsDBNull(table6.Rows(i)("成交日期")) Then
        '                    日期成交 = table6.Rows(i)("成交日期").ToString
        '                End If

        '                'If Request("oid") = "11335LAA41691" Then

        '                'Else
        '                '先把讀出的xml複製一份， 接著開始改
        '                Dim tempdata As String = myText_ET
        '                tempdata = tempdata.Replace("≠成交座落", 成交座落)
        '                tempdata = tempdata.Replace("≠用途", 用途)
        '                tempdata = tempdata.Replace("≠類別", 類別)
        '                tempdata = tempdata.Replace("≠竣工日期", 竣工日期)
        '                tempdata = tempdata.Replace("≠樓層", 樓層)
        '                tempdata = tempdata.Replace("≠面積", 面積)
        '                tempdata = tempdata.Replace("≠成交車位", 成交車位)
        '                tempdata = tempdata.Replace("≠成交", 成交)
        '                tempdata = tempdata.Replace("≠日期成交", 日期成交)
        '                '改完加到最後的字串裡面  
        '                最後要替代掉的字串.Append(tempdata)
        '                'End If

        '                'End If



        '            Next
        '        Else
        '            Dim tempdata As String = myText_ET
        '            tempdata = tempdata.Replace("≠成交座落", "無成交資料")
        '            tempdata = tempdata.Replace("≠用途", "")
        '            tempdata = tempdata.Replace("≠類別", "")
        '            tempdata = tempdata.Replace("≠竣工日期", "")
        '            tempdata = tempdata.Replace("≠樓層", "")
        '            tempdata = tempdata.Replace("≠面積", "")
        '            tempdata = tempdata.Replace("≠成交車位", "")
        '            tempdata = tempdata.Replace("≠成交", "")
        '            tempdata = tempdata.Replace("≠日期成交", "")
        '            '改完加到最後的字串裡面  
        '            最後要替代掉的字串.Append(tempdata)
        '        End If
        '        sFile = sFile.Replace("!重覆品牌成交行情", 最後要替代掉的字串.ToString())
        '    Else
        '        Dim 最後要替代掉的字串 As New StringBuilder()
        '        Dim tempdata As String = myText_ET
        '        tempdata = tempdata.Replace("≠成交座落", "物件資料未定位")
        '        tempdata = tempdata.Replace("≠用途", "")
        '        tempdata = tempdata.Replace("≠類別", "")
        '        tempdata = tempdata.Replace("≠竣工日期", "")
        '        tempdata = tempdata.Replace("≠樓層", "")
        '        tempdata = tempdata.Replace("≠面積", "")
        '        tempdata = tempdata.Replace("≠成交車位", "")
        '        tempdata = tempdata.Replace("≠成交", "")
        '        tempdata = tempdata.Replace("≠日期成交", "")
        '        '改完加到最後的字串裡面  
        '        最後要替代掉的字串.Append(tempdata)
        '        sFile = sFile.Replace("!重覆品牌成交行情", 最後要替代掉的字串.ToString())
        '    End If


        '    '[END============================================品牌成交行情=============================================END]

        'End If





        '共用
        '[START============================================內政部成交行情=============================================START]
        '內政部成交行情
        Dim myXMLO As String = ""
        Dim myTextO As String = ""

        '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        Dim srTextO As New StreamReader(Server.MapPath("\reports\重複實價成交行情_V2.txt"))
        myTextO = srTextO.ReadToEnd()
        srTextO.Close()



        ''以報告日期計算前三個月
        Dim flageO As String = ""
        If IsDBNull(t2.Rows(0)("經度").ToString) Or t2.Rows(0)("經度").ToString.Trim = "" Or IsDBNull(t2.Rows(0)("緯度").ToString) Or t2.Rows(0)("緯度").ToString.Trim = "" Then
            flageO = "0"
        Else
            flageO = "1"
        End If


        Dim sqlstr1 As String = ""
        If flageO = "1" Then


            If DropDownList1.SelectedValue = "依路名" Then
                sqlstr1 = "select top(200) 土地區段位置_建物區段門牌, 交易標的,主要用途, 建物型態,  建築完成年月, 移轉層次, 建物移轉總面積_平方公尺,土地移轉總面積_平方公尺,車位類別, 總價,交易年月, 經度, 緯度 "
                sqlstr1 &= " From 實價登錄DATA  With(NoLock) "
                sqlstr1 &= " Where 類別 = '房地買賣' "
                sqlstr1 &= "And 縣市 = '" & stree1.Text & "' "
                sqlstr1 &= "And 鄉鎮市區 = '" & stree2.Text & "' "
                If DropDownList3.SelectedValue <> "請選擇" Then
                    If Trim(DropDownList3.SelectedItem.Text) = "套房(1房(1廳)1衛)" Then
                        sqlstr1 &= " and 建物型態 like '%套房(1房1廳1衛)%' " '1040430修正...因為與實價登錄DATA內容值不同，故加此判斷
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "土地" Then
                        sqlstr1 &= " and 交易標的 = '土地'"
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "車位" Then
                        sqlstr1 &= " and 交易標的 = '車位'"
                    ElseIf Left(Trim(DropDownList3.SelectedItem.Text), 2) = "公寓" Then
                        sqlstr1 &= " and 建物型態 like '公寓%'"
                    Else
                        sqlstr1 &= " and 建物型態 like '%" & Trim(DropDownList3.SelectedItem.Text) & "%' "
                    End If

                End If


                '交易年月
                If TextBox22.Text <> "" And TextBox25.Text <> "" Then
                    Dim bdate As String = Left(Trim(TextBox22.Text).PadLeft(7, "0"), 5)
                    Dim edate As String = Left(Trim(TextBox25.Text).PadLeft(7, "0"), 5)
                    sqlstr1 &= " And 交易年月 between '" & bdate & "' and '" & edate & "' "
                End If

                '成交金額
                If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
                    Dim b1 As Double = 0, b2 As Double = 99999999
                    'If TextBox1.Text <> "" Then b1 = TextBox1.Text
                    'If TextBox3.Text <> "" Then b2 = TextBox3.Text
                    b1 = Convert.ToDouble(TextBox1.Text)
                    b2 = Convert.ToDouble(TextBox3.Text)
                    sqlstr1 &= " And 總價 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
                End If

                '坪數
                If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
                    Dim b1 As Single = 0, b2 As Single = 99999999
                    If TextBox4.Text <> "" Then b1 = TextBox4.Text
                    If TextBox5.Text <> "" Then b2 = TextBox5.Text
                    sqlstr1 &= " And 建物移轉總面積_平方公尺 between '" & b1 & "' and '" & b2 & "' "
                End If

                '路段名
                If TextBox26.Text.Trim <> "" Then
                    sqlstr1 &= " and (土地區段位置_建物區段門牌 like '%" & TextBox26.Text.Trim & "%'"
                End If

                If TextBox27.Text.Trim <> "" Then
                    sqlstr1 &= " or 土地區段位置_建物區段門牌 like '%" & TextBox27.Text.Trim & "%'"
                End If

                If TextBox28.Text.Trim <> "" Then
                    sqlstr1 &= " or 土地區段位置_建物區段門牌 like '%" & TextBox28.Text.Trim & "%'"
                End If

                If TextBox26.Text.Trim <> "" Then
                    sqlstr1 &= ") "
                End If

                'If ddl路名.Items.Count - 1 > 0 Then
                '    If ddl路名.SelectedValue <> "請選擇" Then
                '        sqlstr1 &= " and (土地區段位置_建物區段門牌 like '%" & Mid(Trim(ddl路名.SelectedValue), 1, Len(Trim(ddl路名.SelectedValue)) - 1) & "%'"
                '        'temp &= Mid(Trim(ddl路名.SelectedValue), 1, Len(Trim(ddl路名.SelectedValue)) - 1)
                '    End If
                '    If ddl路名2.Items.Count - 1 > 0 Then
                '        If ddl路名2.SelectedValue <> "請選擇" Then
                '            sqlstr1 &= " or 土地區段位置_建物區段門牌 like '%" & Mid(Trim(ddl路名2.SelectedValue), 1, Len(Trim(ddl路名2.SelectedValue)) - 1) & "%'"
                '            'temp &= "及" & Mid(Trim(ddl路名2.SelectedValue), 1, Len(Trim(ddl路名2.SelectedValue)) - 1)
                '        End If
                '    End If
                '    If ddl路名3.Items.Count - 1 > 0 Then
                '        If ddl路名3.SelectedValue <> "請選擇" Then
                '            sqlstr1 &= " or 土地區段位置_建物區段門牌 like '%" & Mid(Trim(ddl路名3.SelectedValue), 1, Len(Trim(ddl路名3.SelectedValue)) - 1) & "%'"
                '            'temp &= "及" & Mid(Trim(ddl路名3.SelectedValue), 1, Len(Trim(ddl路名3.SelectedValue)) - 1)
                '        End If
                '    End If
                '    If ddl路名.SelectedIndex - 1 >= 0 Then
                '        sqlstr1 &= ") "
                '    End If
                'End If

                sqlstr1 &= " Order By 建築完成年月 desc ,土地區段位置_建物區段門牌 asc"

                'If stree1.Text = "嘉義市" Then
                'Response.Write(sqlstr1)
                'Response.End()
                'End If
            Else '依範圍
                Dim todayO As String = (Left(t2.Rows(0)("報告日期").ToString, 3) + 1911) & "/" & (Mid(t2.Rows(0)("報告日期").ToString, 4, 2)) & "/" + Right(t2.Rows(0)("報告日期").ToString, 2)
                Dim tempdayO = (DateAdd("m", -6, todayO))
                Dim startdayO = Right("000" & Year(tempdayO) - 1911, 3) & Right("00" & Month(tempdayO), 2)
                Dim enddayO = Left(Trim(t2.Rows(0)("報告日期").ToString).PadLeft(7, "0"), 5)
                sqlstr1 = "DECLARE @lat float = " & t2.Rows(0)("經度").ToString
                sqlstr1 &= " DECLARE @lng float = " & t2.Rows(0)("緯度").ToString
                sqlstr1 &= " DECLARE @radius float = 0.5 "                                  ' 半徑 500m
                sqlstr1 &= " select top(200) 土地區段位置_建物區段門牌, 交易標的,主要用途, 建物型態,  建築完成年月, 移轉層次, 建物移轉總面積_平方公尺,土地移轉總面積_平方公尺,車位類別, 總價,交易年月, 經度, 緯度 "
                sqlstr1 &= " FROM 實價登錄DATA  With(NoLock) " '20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
                sqlstr1 &= " where (類別 <> '房地租賃') " 'and  (交易年月 between '" & startdayO & "' and '" & enddayO & "')"
                If DropDownList3.SelectedValue <> "請選擇" Then
                    If Trim(DropDownList3.SelectedItem.Text) = "套房(1房(1廳)1衛)" Then
                        sqlstr1 &= " and 建物型態 like '%套房(1房1廳1衛)%' " '1040430修正...因為與實價登錄DATA內容值不同，故加此判斷
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "土地" Then
                        sqlstr1 &= " and 交易標的 = '土地'"
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "車位" Then
                        sqlstr1 &= " and 交易標的 = '車位'"
                    ElseIf Left(Trim(DropDownList3.SelectedItem.Text), 2) = "公寓" Then
                        sqlstr1 &= " and 建物型態 like '公寓%'"
                    Else
                        sqlstr1 &= " and 建物型態 like '%" & Trim(DropDownList3.SelectedItem.Text) & "%' "
                    End If
                End If

                '交易年月
                If TextBox22.Text <> "" And TextBox25.Text <> "" Then
                    Dim bdate As String = Left(Trim(TextBox22.Text).PadLeft(7, "0"), 5)
                    Dim edate As String = Left(Trim(TextBox25.Text).PadLeft(7, "0"), 5)
                    sqlstr1 &= " And 交易年月 between '" & bdate & "' and '" & edate & "' "
                End If

                '成交金額
                If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
                    Dim b1 As Double = 0, b2 As Double = 99999999
                    b1 = Convert.ToDouble(TextBox1.Text)
                    b2 = Convert.ToDouble(TextBox3.Text)
                    sqlstr1 &= " And 總價 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
                End If

                '坪數
                If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
                    Dim b1 As Single = 0, b2 As Single = 99999999
                    If TextBox4.Text <> "" Then b1 = TextBox4.Text
                    If TextBox5.Text <> "" Then b2 = TextBox5.Text
                    sqlstr1 &= " And 建物移轉總面積_平方公尺 between '" & b1 * 1 & "' and '" & b2 * 1 & "'"
                End If
                sqlstr1 &= " AND (經度 <> @lng OR 緯度 <> @lat) "
                sqlstr1 &= " AND (6371 * acos(cos(radians(@lat)) * cos(radians(緯度)) * cos(radians(@lng) - radians(經度)) + sin(radians(@lat)) * sin(radians(緯度)))) < (@radius) "

                sqlstr1 &= " UNION ALL "

                sqlstr1 &= " select top(200) 土地區段位置_建物區段門牌, 交易標的,主要用途, 建物型態,  建築完成年月, 移轉層次, 建物移轉總面積_平方公尺,土地移轉總面積_平方公尺,車位類別, 總價,交易年月, 經度, 緯度 "
                sqlstr1 &= " FROM 實價登錄DATA  With(NoLock) " '20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
                sqlstr1 &= " where (類別 <> '房地租賃') " 'and  (交易年月 between '" & startdayO & "' and '" & enddayO & "')"
                If DropDownList3.SelectedValue <> "請選擇" Then
                    If Trim(DropDownList3.SelectedItem.Text) = "套房(1房(1廳)1衛)" Then
                        sqlstr1 &= " and 建物型態 like '%套房(1房1廳1衛)%' " '1040430修正...因為與實價登錄DATA內容值不同，故加此判斷
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "土地" Then
                        sqlstr1 &= " and 交易標的 = '土地'"
                    ElseIf Trim(DropDownList3.SelectedItem.Text) = "車位" Then
                        sqlstr1 &= " and 交易標的 = '車位'"
                    Else
                        sqlstr1 &= " and 建物型態 like '%" & Trim(DropDownList3.SelectedItem.Text) & "%' "
                    End If
                End If

                '交易年月
                If TextBox22.Text <> "" And TextBox25.Text <> "" Then
                    Dim bdate As String = Left(Trim(TextBox22.Text).PadLeft(7, "0"), 5)
                    Dim edate As String = Left(Trim(TextBox25.Text).PadLeft(7, "0"), 5)
                    sqlstr1 &= " And 交易年月 between '" & bdate & "' and '" & edate & "' "
                End If

                '成交金額
                If TextBox1.Text <> "" Or TextBox3.Text <> "" Then
                    Dim b1 As Double = 0, b2 As Double = 99999999
                    b1 = Convert.ToDouble(TextBox1.Text)
                    b2 = Convert.ToDouble(TextBox3.Text)
                    sqlstr1 &= " And 總價 between '" & b1 * 10000 & "' and '" & b2 * 10000 & "' "
                End If

                '坪數
                If TextBox4.Text <> "" Or TextBox5.Text <> "" Then
                    Dim b1 As Single = 0, b2 As Single = 99999999
                    If TextBox4.Text <> "" Then b1 = TextBox4.Text
                    If TextBox5.Text <> "" Then b2 = TextBox5.Text
                    sqlstr1 &= " And 建物移轉總面積_平方公尺 between '" & b1 * 1 & "' and '" & b2 * 1 & "'"
                End If
                sqlstr1 &= " AND 經度 = @lng AND 緯度 = @lat "
                sqlstr1 &= " order by 交易年月 desc"
            End If

            'If Request("oid") = "61132AAE68871" Then
            '    Response.Write(sqlstr1)
            '    Response.End()
            'End If

            Dim table7 As DataTable
            Dim 內政部成交座落 As String
            Dim 內政部用途 As String
            Dim 內政部類別 As String
            Dim 內政部竣工日期 As String
            Dim 內政部樓層 As String
            Dim 內政部面積 As String
            Dim 內政部成交車位 As String
            Dim 內政部成交 As String
            Dim 內政部日期成交 As String
            i = 0
            adpt = New SqlDataAdapter(sqlstr1, conn)

            ds = New DataSet()
            adpt.Fill(ds, "table7")
            table7 = ds.Tables("table7")
            Dim 內政部最後要替代掉的字串 As New StringBuilder()


            'If stree1.Text = "嘉義市" Then
            'Response.Write(table7.Rows.Count)
            'Response.End()
            'End If

            'If Request("oid") = "61311MAA45599" Then
            '    Response.Write(sqlstr1)
            '    Response.End()
            'End If

            If table7.Rows.Count > 0 Then
                For i = 0 To table7.Rows.Count - 1
                    內政部成交座落 = table7.Rows(i)("土地區段位置_建物區段門牌").ToString
                    內政部用途 = table7.Rows(i)("主要用途").ToString

                    Dim 型態 As String = ""
                    Select Case table7.Rows(i)("建物型態").ToString

                        Case "住宅大樓(11層含以上有電梯)"
                            型態 = "大樓"
                        Case "公寓(5樓含以下)"
                            型態 = "公寓"
                        Case "套房(1房1廳1衛)"
                            型態 = "套房"
                        Case "華廈(10層含以下有電梯)"
                            型態 = "華廈"
                        Case "辦公商業大樓"
                            型態 = "商辦"
                        Case "店面(店鋪)"
                            型態 = "店面"
                        Case Else
                            型態 = table7.Rows(i)("建物型態").ToString
                    End Select
                    內政部類別 = 型態
                    內政部竣工日期 = table7.Rows(i)("建築完成年月").ToString
                    內政部樓層 = table7.Rows(i)("移轉層次").ToString

                    If table7.Rows(i)("交易標的").ToString = "土地" Then
                        內政部面積 = Math.Round(Convert.ToDouble("0" & table7.Rows(i)("土地移轉總面積_平方公尺").ToString), 2, MidpointRounding.AwayFromZero)
                    Else
                        內政部面積 = Math.Round(Convert.ToDouble("0" & table7.Rows(i)("建物移轉總面積_平方公尺").ToString), 2, MidpointRounding.AwayFromZero)
                    End If

                    Dim ping As Double = Math.Round(內政部面積 / 3.305785, 2, MidpointRounding.AwayFromZero)
                    Dim strPing As String = String.Format("(約{0}坪)", ping)
                    內政部成交車位 = table7.Rows(i)("車位類別").ToString
                    內政部成交 = String.Format("{0:#,###.00}", Math.Round(Convert.ToDouble("0" & (table7.Rows(i)("總價") / 10000).ToString), 2, MidpointRounding.AwayFromZero)) 'Math.Round(Convert.ToDouble("0" & (table7.Rows(i)("總價") / 10000).ToString), 2, MidpointRounding.AwayFromZero)
                    內政部日期成交 = table7.Rows(i)("交易年月").ToString

                    'If Request("oid") = "61322AAE63660" Then
                    '先把讀出的xml複製一份， 接著開始改
                    Dim tempdataO As String = myTextO
                    tempdataO = tempdataO.Replace("≠成交座落", 內政部成交座落)
                    tempdataO = tempdataO.Replace("≠用途", 內政部用途)
                    tempdataO = tempdataO.Replace("≠類別", 內政部類別)
                    tempdataO = tempdataO.Replace("≠竣工日期", 內政部竣工日期)
                    tempdataO = tempdataO.Replace("≠樓層", 內政部樓層)
                    tempdataO = tempdataO.Replace("≠面積", 內政部面積 + Environment.NewLine + strPing)
                    tempdataO = tempdataO.Replace("≠成交車位", 內政部成交車位)
                    tempdataO = tempdataO.Replace("≠成交", 內政部成交)
                    tempdataO = tempdataO.Replace("≠日期成交", 內政部日期成交)
                    '改完加到最後的字串裡面
                    內政部最後要替代掉的字串.Append(tempdataO)
                    'Else
                    '    '先把讀出的xml複製一份， 接著開始改
                    '    Dim tempdataO As String = myTextO
                    '    tempdataO = tempdataO.Replace("≠成交座落", 內政部成交座落)
                    '    tempdataO = tempdataO.Replace("≠用途", 內政部用途)
                    '    tempdataO = tempdataO.Replace("≠類別", 內政部類別)
                    '    tempdataO = tempdataO.Replace("≠竣工日期", 內政部竣工日期)
                    '    tempdataO = tempdataO.Replace("≠樓層", 內政部樓層)
                    '    tempdataO = tempdataO.Replace("≠面積", 內政部面積)
                    '    tempdataO = tempdataO.Replace("≠成交車位", 內政部成交車位)
                    '    tempdataO = tempdataO.Replace("≠成交", 內政部成交)
                    '    tempdataO = tempdataO.Replace("≠日期成交", 內政部日期成交)
                    '    '改完加到最後的字串裡面
                    '    內政部最後要替代掉的字串.Append(tempdataO)
                    'End If

                    'End If
                Next
            Else
                Dim tempdataO As String = myTextO
                tempdataO = tempdataO.Replace("≠成交座落", "無成交資料")
                tempdataO = tempdataO.Replace("≠用途", "")
                tempdataO = tempdataO.Replace("≠類別", "")
                tempdataO = tempdataO.Replace("≠竣工日期", "")
                tempdataO = tempdataO.Replace("≠樓層", "")
                tempdataO = tempdataO.Replace("≠面積", "")
                tempdataO = tempdataO.Replace("≠成交車位", "")
                tempdataO = tempdataO.Replace("≠成交", "")
                tempdataO = tempdataO.Replace("≠日期成交", "")
                '改完加到最後的字串裡面  
                內政部最後要替代掉的字串.Append(tempdataO)
            End If
            sFile = sFile.Replace("!重複實價成交行情", 內政部最後要替代掉的字串.ToString())
        Else
            Dim 內政部最後要替代掉的字串 As New StringBuilder()
            Dim tempdataO As String = myTextO
            tempdataO = tempdataO.Replace("≠成交座落", "物件資料未定位")
            tempdataO = tempdataO.Replace("≠用途", "")
            tempdataO = tempdataO.Replace("≠類別", "")
            tempdataO = tempdataO.Replace("≠竣工日期", "")
            tempdataO = tempdataO.Replace("≠樓層", "")
            tempdataO = tempdataO.Replace("≠面積", "")
            tempdataO = tempdataO.Replace("≠成交車位", "")
            tempdataO = tempdataO.Replace("≠成交", "")
            tempdataO = tempdataO.Replace("≠日期成交", "")
            '改完加到最後的字串裡面  
            內政部最後要替代掉的字串.Append(tempdataO)
            sFile = sFile.Replace("!重複實價成交行情", 內政部最後要替代掉的字串.ToString())
        End If
        '[END============================================內政部成交行情=============================================END]

        Dim paths As String = ""
        Dim newpaths As String = Server.MapPath("~/images/newobject.jpg")
        Dim Img_byte() As Byte
        Dim ImgObj As Object

        '1000325小豪新增-相片資料夾區分為過期(expired).有效(available)
        'If t2.Rows(0).Item("委託截止日") >= (Format(Year(Today) - 1911, "00#") & Format(Month(Today), "0#") & Format(Day(Today), "0#")) Then
        Me.Label2.Text = "available"
        'Else
        '    Me.Label2.Text = "expired"
        'End If


        Dim newWidth As Double
        Dim newHeight As Double

        ' 地圖部份 ============================================================

        '判別網路上有沒有地圖

        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "m.jpg"
        newWidth = 189
        newHeight = 189
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "m.jpg", Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg")

                    Dim sFileName As String = Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg"
                    Dim sFileStream As FileStream = File.OpenRead(sFileName)
                    Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                    'If image.Width > 1024 Or image.Height > 715 Then '480,360
                    newWidth = 189
                    newHeight = 189
                    'If image.Width > image.Height Then
                    '    newWidth = 189
                    '    newHeight = image.Height / (image.Width / 189)
                    'Else
                    '    newHeight = 253.5
                    '    newWidth = image.Width / (image.Height / 253.5)
                    'End If
                    image.Dispose()
                    sFileStream.Close()
                    sFileStream.Dispose()

                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "m.jpg")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try


        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPicm", Xml2Doc.MyBase64Encode(ImgObj))
        sFile = sFile.Replace("≠地圖寬", newWidth & "pt")
        sFile = sFile.Replace("≠地圖長", newHeight & "pt")



        '物件照片部份 ============================================================

        '物件照片1
        newpaths = Server.MapPath("~/images/newobject.jpg")
        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "a.jpg"
        newHeight = 156
        newWidth = 224.25
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "a.jpg", Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg")
                    Dim sFileName As String = Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg"
                    Dim sFileStream As FileStream = File.OpenRead(sFileName)
                    Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                    'If image.Width > 1024 Or image.Height > 715 Then '480,360
                    If (image.Width / image.Height) <= 1.4375 Then
                        newHeight = 156
                        newWidth = image.Width / (image.Height / 156)
                    Else
                        newWidth = 224.25
                        newHeight = image.Height / (image.Width / 224.25)
                    End If
                    'If image.Width > image.Height Then
                    '    newWidth = 224.25
                    '    newHeight = image.Height / (image.Width / 224.25)
                    'Else
                    '    newHeight = 156
                    '    newWidth = image.Width / (image.Height / 156)
                    'End If
                    image.Dispose()
                    sFileStream.Close()
                    sFileStream.Dispose()
                    'Else
                    '    newWidth = image.Width
                    '    newHeight = image.Height
                    'End If

                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "a.jpg")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try


        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPic1", Xml2Doc.MyBase64Encode(ImgObj))
        'sFile = sFile.Replace("≠objPic1W", newWidth.ToString)
        'sFile = sFile.Replace("≠objPic1H", newHeight.ToString)
        sFile = sFile.Replace("≠寬1", newWidth & "pt")
        sFile = sFile.Replace("≠高1", newHeight & "pt")
        'sFile = sFile.Replace("≠objPic1_1", "width:224.25pt;height:156pt")

        '物件照片2
        newpaths = Server.MapPath("~/images/newobject.jpg")
        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "b.jpg"
        newHeight = 156
        newWidth = 224.25
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "b.jpg", Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg")
                    Dim sFileName As String = Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg"
                    Dim sFileStream As FileStream = File.OpenRead(sFileName)
                    Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                    'If image.Width > 1024 Or image.Height > 715 Then '480,360
                    If (image.Width / image.Height) <= 1.4375 Then
                        newHeight = 156
                        newWidth = image.Width / (image.Height / 156)
                    Else
                        newWidth = 224.25
                        newHeight = image.Height / (image.Width / 224.25)
                    End If
                    'If image.Width > image.Height Then
                    '    newWidth = 224.25
                    '    newHeight = image.Height / (image.Width / 224.25)
                    'Else
                    '    newHeight = 156
                    '    newWidth = image.Width / (image.Height / 156)
                    'End If
                    image.Dispose()
                    sFileStream.Close()
                    sFileStream.Dispose()

                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "b.jpg")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try


        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPic2", Xml2Doc.MyBase64Encode(ImgObj))
        sFile = sFile.Replace("≠寬2", newWidth & "pt")
        sFile = sFile.Replace("≠高2", newHeight & "pt")

        '物件照片3
        newpaths = Server.MapPath("~/images/newobject.jpg")
        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "c.jpg"

        newHeight = 156
        newWidth = 224.25
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "c.jpg", Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg")
                    Dim sFileName As String = Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg"
                    Dim sFileStream As FileStream = File.OpenRead(sFileName)
                    Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                    'If image.Width > 1024 Or image.Height > 715 Then '480,360
                    If (image.Width / image.Height) <= 1.4375 Then
                        newHeight = 156
                        newWidth = image.Width / (image.Height / 156)
                    Else
                        newWidth = 224.25
                        newHeight = image.Height / (image.Width / 224.25)
                    End If
                    'If image.Width > image.Height Then
                    '    newWidth = 224.25
                    '    newHeight = image.Height / (image.Width / 224.25)
                    'Else
                    '    newHeight = 156
                    '    newWidth = image.Width / (image.Height / 156)
                    'End If
                    image.Dispose()
                    sFileStream.Close()
                    sFileStream.Dispose()

                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "c.jpg")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try


        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPic3", Xml2Doc.MyBase64Encode(ImgObj))
        sFile = sFile.Replace("≠寬3", newWidth & "pt")
        sFile = sFile.Replace("≠高3", newHeight & "pt")

        '物件照片4
        newpaths = Server.MapPath("~/images/newobject.jpg")
        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "d.jpg"
        newHeight = 156
        newWidth = 224.25
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "d.jpg", Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg")
                    Dim sFileName As String = Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg"
                    Dim sFileStream As FileStream = File.OpenRead(sFileName)
                    Dim image As System.Drawing.Image = System.Drawing.Image.FromStream(sFileStream)
                    'If image.Width > 1024 Or image.Height > 715 Then '480,360
                    If (image.Width / image.Height) <= 1.4375 Then
                        newHeight = 156
                        newWidth = image.Width / (image.Height / 156)
                    Else
                        newWidth = 224.25
                        newHeight = image.Height / (image.Width / 224.25)
                    End If
                    'If image.Width > image.Height Then
                    '    newWidth = 224.25
                    '    newHeight = image.Height / (image.Width / 224.25)
                    'Else
                    '    newHeight = 156
                    '    newWidth = image.Width / (image.Height / 156)
                    'End If
                    image.Dispose()
                    sFileStream.Close()
                    sFileStream.Dispose()

                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "d.jpg")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try


        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPic4", Xml2Doc.MyBase64Encode(ImgObj))
        sFile = sFile.Replace("≠寬4", newWidth & "pt")
        sFile = sFile.Replace("≠高4", newHeight & "pt")

        '判別網路上有沒有條碼
        newpaths = Server.MapPath("~/images/newobject.jpg")
        paths = "https://img.etwarm.com.tw/" & Request("sid") & "/" & Label2.Text & "/" & Contract_No & "BC.PNG"

        Try
            Dim requests As WebRequest = HttpWebRequest.Create(paths)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "BC.PNG") Then
                        File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "BC.PNG")
                    End If

                    My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/" & Me.Label2.Text & "/" & Contract_No & "BC.PNG", Server.MapPath("~/PICFILE/") & Contract_No & "BC.PNG")

                    Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "BC.PNG")


                    '20110412刪除UPFILE裡的物件
                    File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "BC.PNG")
                End If
            End Using
        Catch ex As WebException
            'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            'If webResponse.StatusCode = HttpStatusCode.NotFound Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
            'End If
        End Try

        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠objPicBG", Xml2Doc.MyBase64Encode(ImgObj))

        '品牌
        If InStr(Request("sid").ToUpper, "S") = 0 Then
            Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/images/description_etwarm.png"))
            sFile = sFile.Replace("≠品牌", "東森房屋")
        Else
            Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/images/description_life.png"))
            sFile = sFile.Replace("≠品牌", "森活不動產")
        End If
        '將圖檔轉成Base64Code後置換
        ImgObj = Img_byte
        sFile = sFile.Replace("≠logo", Xml2Doc.MyBase64Encode(ImgObj))

        'If Not IsDBNull(t2.Rows(0)("經度")) And Not IsDBNull(t2.Rows(0)("緯度")) Then

        '    Dim store As String = Trim(t2.Rows(0)("店代號").ToString) '"A0001"  '加盟店
        '    Dim lat As String = Trim(t2.Rows(0)("經度").ToString) '"25.0448370467" '緯度
        '    Dim lng As String = Trim(t2.Rows(0)("緯度").ToString) '"121.524685661" '經度 
        '    Dim life1 As String = ""
        '    i = 0
        '    For i = 0 To CheckBoxList1.Items.Count - 1
        '        If CheckBoxList1.Items(i).Selected = True Then
        '            life1 &= CheckBoxList1.Items(i).Value & ","
        '        End If
        '    Next

        '    i = 0
        '    If CheckBoxList2.Visible = True Then
        '        For i = 0 To CheckBoxList2.Items.Count - 1
        '            If CheckBoxList2.Items(i).Selected = True Then
        '                life1 &= CheckBoxList2.Items(i).Value & ","
        '            End If
        '        Next
        '    End If

        '    Dim types As String = ""
        '    Dim str2 As String()
        '    str2 = life1.Split(",")
        '    i = 0
        '    For i = 0 To str2.Length - 1
        '        If str2(i) = "1" Then '公有市場
        '            types &= "'H3201000','H3202000','H3203000',"
        '        ElseIf str2(i) = "2" Then '超級市場
        '            types &= "'H0400001','H0400002','H0400003','H0400004','H0400005','H0400006','H0400007','H0400008','H0400009','H0400010','H0400011','H0400012','H0400013',"
        '        ElseIf str2(i) = "3" Then '學校
        '            types &= "'B0200000','B0302000','B0303000','B0301000','B0400000','B0500000','B0800000',"
        '        ElseIf str2(i) = "4" Then '警察局(分駐所、派出所)
        '            types &= "'A0302000',"
        '        ElseIf str2(i) = "5" Then '行政機關
        '            types &= "'A0100000','A0201000','A0202000','A0203000','A0204000','A0205000','A0206000','A0207000','A0208000','A0210000','A0211000','A0212000','A0213000','A0215000',"
        '        ElseIf str2(i) = "19" Then '體育場
        '            types &= "'G0404000',"
        '        ElseIf str2(i) = "6" Then '醫院
        '            types &= "'C0100000','C0600000','C0800000','C1100000','C1200000',"
        '        ElseIf str2(i) = "7" Then '飛機場 
        '            types &= "'K0400000',"
        '        ElseIf str2(i) = "8" Then '台電變電所用地 
        '            types &= "'X0101000','X0102000','X0200000',"
        '        ElseIf str2(i) = "9" Then '地面高壓電塔(線) 
        '            types &= "'X0300000',"
        '        ElseIf str2(i) = "20" Then '寺廟
        '            types &= "'X0800000',"
        '        ElseIf str2(i) = "10" Then '殯儀館 
        '            types &= "'L0504000',"
        '        ElseIf str2(i) = "11" Then '公墓
        '            types &= "'L0501000',"
        '        ElseIf str2(i) = "12" Then '火化場
        '            types &= "'L0502000',"
        '        ElseIf str2(i) = "13" Then '骨灰(骸)存放設施
        '            types &= "'L0503000',"
        '        ElseIf str2(i) = "14" Then '垃圾場(掩埋場、焚化廠)
        '            types &= "'X0500000','X0600000',"
        '        ElseIf str2(i) = "16" Then '加(氣)油站"
        '            types &= "'J0101000','J0102000','J0103000','X0900000',"
        '        ElseIf str2(i) = "17" Then '瓦斯行(場)
        '            types &= "'X1000000','X1100000',"
        '        ElseIf str2(i) = "18" Then '葬儀社
        '            types &= "'L0505000',"
        '        End If
        '    Next



        '    sql = "SELECT   Code_type, name, Mangt_Add, latitude, longitude, Code, ( 6371 * acos( cos( radians(" & lat & ") ) * cos( radians( Latitude ) ) * cos( radians( Longitude ) - radians(" & lng & ") ) + sin( radians(" & lat & ") ) * sin( radians( Latitude ) ) ) ) AS distance  "
        '    sql &= " FROM life_funtion  where  "
        '    sql &= " Code in (" & Mid(types, 1, Len(types) - 1) & ") "
        '    sql &= " GROUP BY Code_type, name, Mangt_Add, latitude, longitude, Code"
        '    sql &= " HAVING ( 6371 * acos( cos( radians(" & lat & ") ) * cos( radians( Latitude ) ) * cos( radians( Longitude ) - radians(" & lng & ") ) + sin( radians(" & lat & ") ) * sin( radians( Latitude ) ) ) ) < (0.3)"
        '    sql &= " ORDER BY distance ASC"



        '    adpt = New SqlDataAdapter(sql, conn)
        '    ds = New DataSet()
        '    adpt.Fill(ds, "鄰近重要設施")
        '    Dim table_重要設施 As DataTable = ds.Tables("鄰近重要設施")
        '    Dim dt As DataTable = table_重要設施

        '    Dim dt1 As DataTable
        '    Dim row As DataRow
        '    Dim ds1 As New DataSet

        '    '虛擬表格_統計表
        '    ds1.Tables.Add("重要設施")
        '    dt1 = ds1.Tables("重要設施")
        '    dt1.Columns.Add("id", GetType(System.String))
        '    dt1.Columns.Add("Code_type", GetType(System.String))
        '    dt1.Columns.Add("iconName", GetType(System.String))
        '    dt1.Columns.Add("Name", GetType(System.String))
        '    dt1.Columns.Add("Mangt_Add", GetType(System.String))
        '    dt1.Columns.Add("Longitude", GetType(System.String))
        '    dt1.Columns.Add("Latitude", GetType(System.String))


        '    If dt.Rows.Count > 0 Then

        '        For i = 0 To dt.Rows.Count - 1 '超過20筆..網址會過長

        '            row = dt1.NewRow

        '            row("id") = i + 1

        '            If dt.Rows(i)("Code") = "H3201000" Or dt.Rows(i)("Code") = "H3202000" Or dt.Rows(i)("Code") = "H3203000" Then '類別
        '                row("Code_type") = "公有市場"
        '                row("iconName") = "imp1.png"
        '            ElseIf dt.Rows(i)("Code") = "H0400001" Or dt.Rows(i)("Code") = "H0400002" Or dt.Rows(i)("Code") = "H0400003" Or dt.Rows(i)("Code") = "H0400004" Or dt.Rows(i)("Code") = "H0400005" Or dt.Rows(i)("Code") = "H0400006" Or dt.Rows(i)("Code") = "H0400007" Or dt.Rows(i)("Code") = "H0400008" Or dt.Rows(i)("Code") = "H0400009" Or dt.Rows(i)("Code") = "H0400010" Or dt.Rows(i)("Code") = "H0400011" Or dt.Rows(i)("Code") = "H0400012" Or dt.Rows(i)("Code") = "H0400013" Then
        '                row("Code_type") = "超級市場"
        '                row("iconName") = "imp2.png"
        '            ElseIf dt.Rows(i)("Code") = "B0200000" Or dt.Rows(i)("Code") = "B0302000" Or dt.Rows(i)("Code") = "B0303000" Or dt.Rows(i)("Code") = "B0301000" Or dt.Rows(i)("Code") = "B0400000" Or dt.Rows(i)("Code") = "B0500000" Or dt.Rows(i)("Code") = "B0800000" Then
        '                row("Code_type") = "學校"
        '                row("iconName") = "imp3.png"
        '            ElseIf dt.Rows(i)("Code") = "A0302000" Then
        '                row("Code_type") = "警察局(分駐所、派出所)"
        '                row("iconName") = "imp4.png"
        '            ElseIf dt.Rows(i)("Code") = "A0100000" Or dt.Rows(i)("Code") = "A0201000" Or dt.Rows(i)("Code") = "A0202000" Or dt.Rows(i)("Code") = "A0203000" Or dt.Rows(i)("Code") = "A0204000" Or dt.Rows(i)("Code") = "A0205000" Or dt.Rows(i)("Code") = "A0206000" Or dt.Rows(i)("Code") = "A0207000" Or dt.Rows(i)("Code") = "A0208000" Or dt.Rows(i)("Code") = "A0210000" Or dt.Rows(i)("Code") = "A0211000" Or dt.Rows(i)("Code") = "A0212000" Or dt.Rows(i)("Code") = "A0213000" Or dt.Rows(i)("Code") = "A0215000" Then
        '                row("Code_type") = "行政機關"
        '                row("iconName") = "imp5.png"
        '            ElseIf dt.Rows(i)("Code") = "G0404000" Then
        '                row("Code_type") = "體育場"
        '                row("iconName") = "imp19.png"
        '            ElseIf dt.Rows(i)("Code") = "C0100000" Or dt.Rows(i)("Code") = "C0600000" Or dt.Rows(i)("Code") = "C0800000" Or dt.Rows(i)("Code") = "C1100000" Or dt.Rows(i)("Code") = "C1200000" Then
        '                row("Code_type") = "醫院"
        '                row("iconName") = "imp6.png"
        '            ElseIf dt.Rows(i)("Code") = "K0400000" Then
        '                row("Code_type") = "飛機場"
        '                row("iconName") = "imp7.png"
        '            ElseIf dt.Rows(i)("Code") = "X0101000" Or dt.Rows(i)("Code") = "X0102000" Or dt.Rows(i)("Code") = "X0200000" Then
        '                row("Code_type") = "台電變電所用地"
        '                row("iconName") = "imp8.png"
        '            ElseIf dt.Rows(i)("Code") = "X0300000" Then
        '                row("Code_type") = "地面高壓電塔(線)"
        '                row("iconName") = "imp9.png"
        '            ElseIf dt.Rows(i)("Code") = "X0800000" Then
        '                row("Code_type") = "寺廟"
        '                row("iconName") = "imp20.png"
        '            ElseIf dt.Rows(i)("Code") = "L0504000" Then
        '                row("Code_type") = "殯儀館"
        '                row("iconName") = "imp10.png"
        '            ElseIf dt.Rows(i)("Code") = "L0501000" Then
        '                row("Code_type") = "公墓"
        '                row("iconName") = "imp11.png"
        '            ElseIf dt.Rows(i)("Code") = "L0502000" Then
        '                row("Code_type") = "火化場"
        '                row("iconName") = "imp12.png"
        '            ElseIf dt.Rows(i)("Code") = "L0503000" Then
        '                row("Code_type") = "骨灰(骸)存放設施"
        '                row("iconName") = "imp13.png"
        '            ElseIf dt.Rows(i)("Code") = "X0500000" Or dt.Rows(i)("Code") = "X0600000" Then
        '                row("Code_type") = "垃圾場(掩埋場、焚化廠)"
        '                row("iconName") = "imp14.png"
        '            ElseIf dt.Rows(i)("Code") = "J0101000" Or dt.Rows(i)("Code") = "J0102000" Or dt.Rows(i)("Code") = "J0103000" Or dt.Rows(i)("Code") = "X0900000" Then
        '                row("Code_type") = "加(氣)油站"
        '                row("iconName") = "imp16.png"
        '            ElseIf dt.Rows(i)("Code") = "X1000000" Or dt.Rows(i)("Code") = "X1100000" Then
        '                row("Code_type") = "瓦斯行(場)"
        '                row("iconName") = "imp17.png"
        '            ElseIf dt.Rows(i)("Code") = "L0505000" Then
        '                row("Code_type") = "葬儀社"
        '                row("iconName") = "imp18.png"
        '            End If


        '            row("Name") = dt.Rows(i)("Name")
        '            row("Mangt_Add") = dt.Rows(i)("Mangt_Add")
        '            row("Longitude") = dt.Rows(i)("Longitude")
        '            row("Latitude") = dt.Rows(i)("Latitude")

        '            dt1.Rows.Add(row)


        '        Next
        '    End If

        '    Dim location As String = ""
        '    Dim modcount As Integer = 0 '循環次數

        '    For i As Integer = 0 To dt1.Rows.Count - 1


        '        'location &= "&markers=icon:https://superwebnew.etwarm.com.tw/images/" & dt1.Rows(i)("iconName") & "%7C" & dt1.Rows(i)("Latitude") & "," & dt1.Rows(i)("Longitude")
        '        Dim imgname As Image
        '        imgname = Image.FromFile(Server.MapPath("../images/" & dt1.Rows(i)("iconName")))

        '        'location &= "&markers=icon:" & CreateImgWithLabel(imgname, i.ToString) & "%7C" & dt1.Rows(i)("Latitude") & "," & dt1.Rows(i)("Longitude")
        '        location &= "&markers=icon:" & shortImgURL(dt1.Rows(i)("iconName")) & "%7C" & dt1.Rows(i)("Latitude") & "," & dt1.Rows(i)("Longitude")

        '        '每頁20筆 所以當 i mod 20 = 0  中斷到下一頁
        '        If i - (20 * (i \ 20)) = 0 Then
        '            '************** 這一段copy 20151210 by nick**********

        '            '判別地址定位是否能訂到
        '            paths = "https://maps.google.com/maps/api/staticmap?language=zh-tw&center=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&markers=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&zoom=17&size=700x300&sensor=false" & location

        '            '下載靜態地圖
        '            'Try
        '            '    Dim requests As WebRequest = HttpWebRequest.Create(paths)
        '            '    requests.Method = "HEAD"
        '            '    requests.Credentials = System.Net.CredentialCache.DefaultCredentials
        '            '    Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
        '            '        If response.StatusCode = HttpStatusCode.OK Then

        '            '            'UPFILE資料夾內檔案若已存在,先刪除該檔案
        '            '            If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG") Then
        '            '                File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")
        '            '            End If

        '            '            My.Computer.Network.DownloadFile(paths, Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")

        '            '            Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")


        '            '            '20110412刪除UPFILE裡的物件
        '            '            File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")
        '            '        End If
        '            '    End Using
        '            'Catch ex As WebException
        '            '    Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
        '            '    If webResponse.StatusCode = HttpStatusCode.NotFound Then
        '            '        Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
        '            '    End If
        '            'End Try

        '            location = ""



        '            '[START============================================鄰近重要設施=============================================START]
        '            Dim 周邊編號 As String = ""
        '            Dim 周邊種類 As String = ""
        '            Dim 周邊名稱 As String = ""
        '            Dim 周邊地址 As String = ""




        '            '鄰近重要設施
        '            Dim myXML_鄰近重要設施 As String = ""
        '            Dim myText_鄰近重要設施 As String = ""

        '            '讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        '            Dim srText_鄰近重要設施 As New StreamReader(Server.MapPath("..\reports\重複重要設施內容_V3.txt"))
        '            myText_鄰近重要設施 = srText_鄰近重要設施.ReadToEnd()
        '            srText_鄰近重要設施.Close()



        '            Dim 最後要替代掉的字串_鄰近重要設施 As New StringBuilder()

        '            If dt1.Rows.Count > 0 Then

        '                For j As Integer = i To (i + 20)
        '                    If j > dt1.Rows.Count - 1 Then
        '                        Exit For
        '                    End If
        '                    '周邊編號
        '                    周邊編號 = j + 1

        '                    '周邊種類
        '                    If Not IsDBNull(dt1.Rows(j)("Code_type")) Then
        '                        周邊種類 = dt1.Rows(j)("Code_type").ToString
        '                    End If
        '                    '周邊名稱
        '                    If Not IsDBNull(dt1.Rows(j)("Name")) Then
        '                        周邊名稱 = dt1.Rows(j)("Name").ToString
        '                    End If
        '                    '周邊地址
        '                    If Not IsDBNull(dt1.Rows(j)("Mangt_Add")) Then
        '                        周邊地址 = dt1.Rows(j)("Mangt_Add").ToString
        '                    End If


        '                    '先把讀出的xml複製一份，接著開始改
        '                    Dim tempdata As String = myText_鄰近重要設施
        '                    tempdata = tempdata.Replace("≠周邊編號", 周邊編號)
        '                    tempdata = tempdata.Replace("≠周邊種類", 周邊種類)
        '                    tempdata = tempdata.Replace("≠周邊名稱", Server.HtmlEncode(周邊名稱))
        '                    tempdata = tempdata.Replace("≠周邊地址", Server.HtmlEncode(周邊地址))

        '                    '將圖檔轉成Base64Code後置換

        '                    Dim imgbyts As Byte() = New System.Net.WebClient().DownloadData(paths)

        '                    'Dim xmlcarparkimg As String = "<w:br/><w:pict><w:binData w:name=""wordml://" & i & """ xml:space=""preserve"">" & Xml2Doc.MyBase64Encode(imgbyts) & "</w:binData><v:shape id=""_x0000_" & i & """ type=""#_x0000_t75"" style=""width:8.8cm;height:6.6cm""><v:imagedata src=""wordml://" & i & """ o:title=""a0001""/></v:shape></w:pict>"
        '                    'tempdata = tempdata.Replace("≠objEraPic", xmlcarparkimg)




        '                    '改完加到最後的字串裡面  
        '                    最後要替代掉的字串_鄰近重要設施.Append(tempdata)

        '                Next
        '            Else
        '                Dim tempdata As String = myText_鄰近重要設施
        '                tempdata = tempdata.Replace("≠周邊編號", "")
        '                tempdata = tempdata.Replace("≠周邊種類", "")
        '                tempdata = tempdata.Replace("≠周邊名稱", "")
        '                tempdata = tempdata.Replace("≠周邊地址", "")


        '                '改完加到最後的字串裡面  
        '                最後要替代掉的字串_鄰近重要設施.Append(tempdata)
        '            End If

        '            sFile = sFile.Replace("!重複重要設施內容", 最後要替代掉的字串_鄰近重要設施.ToString())

        '            '[END============================================鄰近重要設施=============================================END]
        '            '************** 這一段copy 20151210 by nick**********
        '        End If


        '    Next
        '    '規則:&marker:&markers=icon:https://superwebnew.etwarm.com.tw/images/+圖片名+"%7C"+緯度+","+經度

        '    ''************** 這一段copy 20151210 by nick**********

        '    ''判別地址定位是否能訂到
        '    ''paths = "https://maps.google.com/maps/api/staticmap?language=zh-tw&center=台北市中正區館前路42號&zoom=17&size=700x300&sensor=false&markers=icon:https://superwebnew.etwarm.com.tw/images/imp12.png%7C台北市中正區館前路47號"
        '    'paths = "https://maps.google.com/maps/api/staticmap?language=zh-tw&center=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&markers=" & t2.Rows(0)("經度") & "," & t2.Rows(0)("緯度") & "&zoom=17&size=700x300&sensor=false" & location

        '    ''下載靜態地圖
        '    'Try
        '    '    Dim requests As WebRequest = HttpWebRequest.Create(paths)
        '    '    requests.Method = "HEAD"
        '    '    requests.Credentials = System.Net.CredentialCache.DefaultCredentials
        '    '    Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
        '    '        If response.StatusCode = HttpStatusCode.OK Then

        '    '            'UPFILE資料夾內檔案若已存在,先刪除該檔案
        '    '            If File.Exists(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG") Then
        '    '                File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")
        '    '            End If

        '    '            My.Computer.Network.DownloadFile(paths, Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")

        '    '            Img_byte = My.Computer.FileSystem.ReadAllBytes(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")


        '    '            '20110412刪除UPFILE裡的物件
        '    '            File.Delete(Server.MapPath("~/PICFILE/") & Contract_No & "IM.PNG")
        '    '        End If
        '    '    End Using
        '    'Catch ex As WebException
        '    '    Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
        '    '    If webResponse.StatusCode = HttpStatusCode.NotFound Then
        '    '        Img_byte = My.Computer.FileSystem.ReadAllBytes(newpaths)
        '    '    End If
        '    'End Try

        '    ''將圖檔轉成Base64Code後置換
        '    'ImgObj = Img_byte
        '    'sFile = sFile.Replace("≠ObjPicImport", Xml2Doc.MyBase64Encode(ImgObj))


        '    ''[START============================================鄰近重要設施=============================================START]
        '    'Dim 周邊編號 As String = ""
        '    'Dim 周邊種類 As String = ""
        '    'Dim 周邊名稱 As String = ""
        '    'Dim 周邊地址 As String = ""




        '    ''鄰近重要設施
        '    'Dim myXML_鄰近重要設施 As String = ""
        '    'Dim myText_鄰近重要設施 As String = ""

        '    ''讀出剛剛建立的text,裡頭放等等要重複利用的xml  
        '    'Dim srText_鄰近重要設施 As New StreamReader(Server.MapPath("..\reports\重複重要設施內容_V2.txt"))
        '    'myText_鄰近重要設施 = srText_鄰近重要設施.ReadToEnd()
        '    'srText_鄰近重要設施.Close()



        '    'Dim 最後要替代掉的字串_鄰近重要設施 As New StringBuilder()

        '    'If dt1.Rows.Count > 0 Then

        '    '    For i As Integer = 0 To dt1.Rows.Count - 1
        '    '        '周邊編號
        '    '        周邊編號 = i + 1

        '    '        '周邊種類
        '    '        If Not IsDBNull(dt1.Rows(i)("Code_type")) Then
        '    '            周邊種類 = dt1.Rows(i)("Code_type").ToString
        '    '        End If
        '    '        '周邊名稱
        '    '        If Not IsDBNull(dt1.Rows(i)("Name")) Then
        '    '            周邊名稱 = dt1.Rows(i)("Name").ToString
        '    '        End If
        '    '        '周邊地址
        '    '        If Not IsDBNull(dt1.Rows(i)("Mangt_Add")) Then
        '    '            周邊地址 = dt1.Rows(i)("Mangt_Add").ToString
        '    '        End If


        '    '        '先把讀出的xml複製一份，接著開始改
        '    '        Dim tempdata As String = myText_鄰近重要設施
        '    '        tempdata = tempdata.Replace("≠周邊編號", 周邊編號)
        '    '        tempdata = tempdata.Replace("≠周邊種類", 周邊種類)
        '    '        tempdata = tempdata.Replace("≠周邊名稱", Server.HtmlEncode(周邊名稱))
        '    '        tempdata = tempdata.Replace("≠周邊地址", Server.HtmlEncode(周邊地址))

        '    '        '改完加到最後的字串裡面  
        '    '        最後要替代掉的字串_鄰近重要設施.Append(tempdata)

        '    '    Next
        '    'Else
        '    '    Dim tempdata As String = myText_鄰近重要設施
        '    '    tempdata = tempdata.Replace("≠周邊編號", "")
        '    '    tempdata = tempdata.Replace("≠周邊種類", "")
        '    '    tempdata = tempdata.Replace("≠周邊名稱", "")
        '    '    tempdata = tempdata.Replace("≠周邊地址", "")


        '    '    '改完加到最後的字串裡面  
        '    '    最後要替代掉的字串_鄰近重要設施.Append(tempdata)
        '    'End If

        '    'sFile = sFile.Replace("!重複重要設施內容", 最後要替代掉的字串_鄰近重要設施.ToString())

        '    ''[END============================================鄰近重要設施=============================================END]
        '    ''************** 這一段copy 20151210 by nick**********
        'End If

        'If Request.Cookies("webfly_empno").Value.ToLower <> "92H" Then
        '    Response.End()

        '下載檔案
        Dim arrBytes = Xml2Doc.StringToBytes(sFile, enuStandardCodePages.SCP_CP_UTF8)
        If CheckBox5.Checked = True Then
            '20160525 和物調表一起存在暫存資料表 並且壓縮 讓使用者下載
            '存檔
            File.WriteAllBytes(TempDictforZip & RandomDictName & "\" & Contract_No & "_不動產說明書.doc", arrBytes)
            If File.Exists(TempDictforZip & RandomDictName & "\" & Contract_No & "_不動產說明書.doc") Then
                '壓縮
                Dim fz As New Zip.FastZip
                Dim zipfilename As String = CreateRand(3) & ".zip"
                fz.CreateEmptyDirectories = True
                fz.CreateZip(TempDictforZip & zipfilename, TempDictforZip & RandomDictName, False, "")

                '輸出
                HttpContext.Current.Response.Clear()
                HttpContext.Current.Response.ClearHeaders()
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + zipfilename)
                HttpContext.Current.Response.ContentType = "application/octet-stream"
                HttpContext.Current.Response.WriteFile(TempDictforZip & zipfilename)
                HttpContext.Current.Response.Flush()

                '刪除
                For Each file As String In Directory.GetFiles(TempDictforZip & RandomDictName)
                    System.IO.File.Delete(file)
                Next
                System.IO.Directory.Delete(TempDictforZip & RandomDictName, True)
                System.IO.File.Delete(TempDictforZip & zipfilename)

                HttpContext.Current.Response.End()
            End If
        Else
            '沒有選擇和物調表一起
            ResponseFile(arrBytes, "application/msword", Contract_No & "description.doc")
        End If

        'End If
        conn.Close()
        conn.Dispose()
    End Sub

    Const BASE_API_URL As String = "https://www.googleapis.com/urlshortener/v1/url"
    Const SHORTENER_URL_PATTERN As String = BASE_API_URL & "?key={0}"
    Const EXPAND_URL_PATTERN As String = BASE_API_URL & "?shortUrl={0}"
    '短網址
    Public Function Shorten(ByVal url As String) As String
        If String.IsNullOrEmpty(url) Then
            Throw New ArgumentNullException("url")
        End If
        Dim m_APIKey As String = "AIzaSyAH0X6sizoKDHInBqkiux7pOax8P3tucmg"
        If m_APIKey.Length = 0 Then
            Throw New Exception("APIKey not set!")
        End If

        Const POST_PATTERN As String = "{{""longUrl"": ""{0}""}}"
        Const MATCH_PATTERN As String = """id"": ?""(?<id>.+)"""

        Dim post = String.Format(POST_PATTERN, url)
        Dim request = DirectCast(WebRequest.Create(String.Format(SHORTENER_URL_PATTERN, m_APIKey)), HttpWebRequest)

        request.Method = "POST"
        request.ContentLength = post.Length
        request.ContentType = "application/json"
        request.Headers.Add("Cache-Control", "no-cache")

        Using requestStream As Stream = request.GetRequestStream()
            Dim buffer = Encoding.ASCII.GetBytes(post)
            requestStream.Write(buffer, 0, buffer.Length)
        End Using

        Using responseStream As Stream = request.GetResponse.GetResponseStream()

            Dim xreader As New StreamReader(responseStream)
            Return Regex.Match(xreader.ReadToEnd(), MATCH_PATTERN).Groups("id").Value

        End Using
        Return String.Empty
    End Function

    Function CreateImgWithLabel(ByVal iconImg As Image, ByVal label As String) As String
        'Dim img As Image
        'img = New Bitmap(54, 49)
        Dim g As Graphics
        g = Graphics.FromImage(iconImg)

        Dim drawFont As New Font("Arial", 10)
        Dim drawBrush As New SolidBrush(Color.White)
        Dim drawPoint As New PointF(20.0F, 21.0F)
        g.DrawString(label, drawFont, drawBrush, drawPoint)

        'Dim randname As String = CInt(Int((10000 * Rnd()) + 1)) & "_" & label & ".png"
        Dim randname As String = CreateRand(4) & ".png"
        Dim filenam As String = Server.MapPath("~/temp/" & randname)
        Dim urlnam As String = ("https://superwebnew.etwarm.com.tw/temp/" & randname)
        iconImg.Save(filenam, System.Drawing.Imaging.ImageFormat.Png)

        Return urlnam
    End Function
    Function CreateRand(ByVal maxSize As Integer) As String
        Dim chars(62) As Char
        chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".ToCharArray()
        Dim data As Byte() = New Byte(1) {}
        Using crypto As New RNGCryptoServiceProvider()
            crypto.GetNonZeroBytes(data)
            data = New Byte(maxSize) {}
            crypto.GetNonZeroBytes(data)
        End Using

        Dim result As New StringBuilder()
        For Each b As Byte In data
            result.Append(chars(b Mod chars.Length))
        Next

        Return result.ToString()
    End Function
    Function shortImgURL(ByVal iconName As String) As String
        Dim returnSTR As String = ""
        If iconName = "imp1.png" Then
            returnSTR = "https://goo.gl/kWUBqT"
        End If
        If iconName = "imp2.png" Then
            returnSTR = "https://goo.gl/7SY3vw"
        End If
        If iconName = "imp3.png" Then
            returnSTR = "https://goo.gl/WoPWZT"
        End If
        If iconName = "imp4.png" Then
            returnSTR = "https://goo.gl/XZaM2X"
        End If
        If iconName = "imp5.png" Then
            returnSTR = "https://goo.gl/wXw9ol"
        End If
        If iconName = "imp6.png" Then
            returnSTR = "https://goo.gl/NYQqJo"
        End If

        If iconName = "imp7.png" Then
            returnSTR = "https://goo.gl/WLbtSQ"
        End If
        If iconName = "imp8.png" Then
            returnSTR = "https://goo.gl/CYrKzE"
        End If
        If iconName = "imp9.png" Then
            returnSTR = "https://goo.gl/yZOe7t"
        End If
        If iconName = "imp10.png" Then
            returnSTR = "https://goo.gl/Of9Q5w"
        End If
        If iconName = "imp11.png" Then
            returnSTR = "https://goo.gl/DQffUJ"
        End If
        If iconName = "imp12.png" Then
            returnSTR = "https://goo.gl/JMhQ4Z"
        End If
        If iconName = "imp13.png" Then
            returnSTR = "https://goo.gl/xnXsVp"
        End If

        If iconName = "imp14.png" Then
            returnSTR = "https://goo.gl/w0fnRW"
        End If
        If iconName = "imp15.png" Then
            returnSTR = ""
        End If
        If iconName = "imp16.png" Then
            returnSTR = "https://goo.gl/EvPhuz"
        End If
        If iconName = "imp17.png" Then
            returnSTR = "https://goo.gl/zZcWla"
        End If
        If iconName = "imp18.png" Then
            returnSTR = "https://goo.gl/ASIayz"
        End If
        If iconName = "imp19.png" Then
            returnSTR = "https://goo.gl/c5qWLM"
        End If
        If iconName = "imp20.png" Then
            returnSTR = "https://goo.gl/LNvfvc"
        End If
        Return returnSTR
    End Function

    Private Sub ResponseFinal()
        With Response
            .Flush()
            .Clear()
            .End()
        End With
    End Sub

    Public Expires As Object = -1
    Public ResponseBuffer As Boolean = True

    Private Sub ResponseInit()
        With Response
            .Expires = Expires
            .Buffer = ResponseBuffer
            .Clear()
        End With
    End Sub

    Public Sub SetResponseHeader(ByVal sContentType As String, ByVal sFileName As String, ByVal bSaveFile As Boolean)
        Dim sContentDisposition As String = ""

        With Response
            If bSaveFile Then
                sContentDisposition = "attachment; " ' 強制存檔，未設定則依瀏覽器預設開啟或存檔
            End If

            If Len(sFileName) > 0 Then
                sContentDisposition = sContentDisposition & "filename=" & sFileName ' 檔名
            End If

            If Len(sContentDisposition) > 0 Then
                'sContentDisposition = System.Text.Encoding.GetEncoding(.Charset).GetString(System.Text.Encoding.GetEncoding("unicode").GetBytes(sContentDisposition))
                .AddHeader("Content-disposition", sContentDisposition)
            End If

            .ContentType = sContentType
        End With
    End Sub

    Public Sub ResponseFile(ByVal vOutput As Object, ByVal sContentType As String, ByVal sFileName As String, Optional ByVal bSaveFile As Boolean = False)
        ResponseInit()

        SetResponseHeader(sContentType, sFileName, bSaveFile)

        Select Case VarType(vOutput)
            Case vbString
                Response.Write(vOutput)
            Case Else
                Response.BinaryWrite(vOutput)
        End Select

        ResponseFinal()
    End Sub
    '------------------------------------------------------------------------------

    Sub power()
        If myobj.Object_ALL_Address = "1" Then '完整地址-可選

        ElseIf myobj.Object_Apart_Address = "1" Then '部分地址-不可選
            CheckBox8.Enabled = False
            CheckBox8.Checked = True
        ElseIf myobj.Object_Other_Address = "1" Then '自己全部他店部份
            If Request("sid") = Request.Cookies("store_id").Value Then
            Else
                CheckBox8.Enabled = False
                CheckBox8.Checked = True
            End If
        End If
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        If Me.DropDownList1.SelectedValue = "依範圍" Then
            show = 0
        Else
            show = 1
        End If
    End Sub

    Protected Sub CheckBox1_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Visible = False
            CheckBox2.Checked = False
        Else
            CheckBox2.Visible = True
        End If
    End Sub
End Class
