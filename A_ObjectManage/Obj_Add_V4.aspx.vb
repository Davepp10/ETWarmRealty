Imports System.Data
Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Math
Imports System.IO
Imports System.Net
Imports Rebex.Net
Imports Newtonsoft.Json

Partial Class Obj_Add_V4
    Inherits System.Web.UI.Page
    Dim checkleave As New checkleave

    '表單CLASS
    Dim formnocheck As New formnocheck

    '權限CLASS
    Dim myobj As New clspowerset
    Dim myobj_tabs As New Cls_tabs

    '連線字串
    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    Dim land As String = ConfigurationManager.ConnectionStrings("land").ConnectionString '中華電信用
    Dim mysqletwarmstring As String = ConfigurationManager.ConnectionStrings("mysqletwarm").ConnectionString
    Dim mysqlegoupload As String = ConfigurationManager.ConnectionStrings("mysqlEGOUPLOAD").ConnectionString
    '連線CODE用-MSSQL
    Dim i As Integer = 0
    Dim sql, sql2 As String
    Dim table1, table2, table3, table4, table1_1, table1_2, table1_3 As DataTable

    Public conn, conn1, conn_land As SqlConnection
    Public cmd, cmd2 As SqlCommand
    Public ds As DataSet
    Public adpt As SqlDataAdapter

    '連線CODE用-MySQL
    Public MySQL_cmd As OdbcCommand
    Public MySQL_adpt As OdbcDataAdapter

    Public trans As String = "False"

    Public show As Integer = 0
    Public copy As Integer = 0
    Public objcls As Integer = 0

    Dim splitarray As Array
    Dim FTP1 As New Ftp
    Dim FTP2 As New Ftp
    Dim YN As String = ""

    Public show_New As String = "0"
    Public show_New_New As String = "0"

    Dim sysdate As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)

    Dim jointsalerule As String = ""    '聯賣規則
    Dim 聯賣組別 As String = ""         '聯賣組別
    Dim 一般專約 As String = ""         '一般專約
    Dim IS_DIRECTLY_OPERATION As Boolean '直營識別
    Dim IS_MANAGER_ROLE As Boolean '管理者角色 店東、店長、秘書

    Public facebook_share_url As String = "0"
    Public line_share_url As String = "0"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Request.Cookies("webfly_empno") Is Nothing Then
            Response.Redirect("../indexnew/login3.aspx")
        End If
        TextBox267.Attributes.Add("onkeypress", "if( event.keyCode == 13 ) { retu rn false; }")
        '權限判斷
        myobj.power_object(Request.Cookies("webfly_empno").Value)
        myobj.mstores(Request.Cookies("webfly_empno").Value)
        myobj.mgroup(Request.Cookies("webfly_empno").Value)

        Dim role = checkleave.直營leave(Request.Cookies("webfly_empno").Value)

        IS_DIRECTLY_OPERATION = checkleave.直營(Request.Cookies("store_id").Value) = "Y"
        IS_MANAGER_ROLE = {"店東", "店長", "秘書"}.Contains(role)

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            Dim selstr As String = "SELECT *,isnull(聯賣一般專約,'') as 一般專約 FROM 區域聯賣成員名單  where  聯賣店代號 like '%" & Request.Cookies("store_id").Value & "%' and 啟用='Y'"
            Using cmd As New SqlCommand(selstr, conn)
                Dim dt As New DataTable
                dt.Load(cmd.ExecuteReader())
                If dt.Rows.Count > 0 Then
                    jointsalerule = dt.Rows(0).Item("聯賣規則代號")
                    聯賣組別 = dt.Rows(0).Item("組別")
                    一般專約 = dt.Rows(0).Item("一般專約")
                End If
            End Using
        End Using

        '狀態參數
        Dim state As String = Request("state")
        If state = "add" Then
            copy = 0
            objcls = 1
            show_New = "0"
        ElseIf state = "update" Then
            copy = 0
            objcls = 0
            'If Request.Cookies("store_id").Value.ToString = "A0001" Or Request.Cookies("store_id").Value.ToString = "A0002" Or Request.Cookies("store_id").Value.ToString = "A1322" Then
            '    show_New_New = "1"
            'Else
            '    show_New_New = "0"
            'End If
            show_New_New = "1"
            'If Request.Cookies("webfly_empno").Value = "92H" Or Request.Cookies("webfly_empno").Value = "0026" Then
            '    show_New = "1"
            'Else
            '    show_New = "0"
            'End If
            'show_New = "0"
            show_New = "1"
        ElseIf state = "copy" Then
            copy = 1
            objcls = 0
            show_New = "0"
        End If

        '宜蘭聯賣網取消所以註解段
        'If Request.Cookies("store_id").Value = "A1076" Or Request.Cookies("store_id").Value = "A0641" Or Request.Cookies("store_id").Value = "A1157" Or Request.Cookies("store_id").Value = "A0883" Or Request.Cookies("store_id").Value = "A0855" Or Request.Cookies("store_id").Value = "A1183" Then
        '    ImageButton20.Visible = False
        'End If

        'If Request.Cookies("store_id").Value = "A0001" Then
        '    ImageButton6.visible = True
        'End If
        If Not Page.IsPostBack Then
            If jointsalerule = "3" Then
                If 一般專約 = "1" And ddl契約類別.SelectedValue = "一般" Then
                    CheckBox101.Checked = False
                    CheckBox101.Visible = True
                ElseIf 一般專約 = "2" And ddl契約類別.SelectedValue = "專任" Then
                    CheckBox101.Checked = False
                    CheckBox101.Visible = True
                ElseIf 一般專約 = "3" Then
                    CheckBox101.Checked = False
                    CheckBox101.Visible = True
                Else
                    CheckBox101.Checked = False
                    CheckBox101.Visible = False
                End If
                'CheckBox101.checked = False
                'CheckBox101.visible = True
            Else
                CheckBox101.Checked = False
                CheckBox101.Visible = False
            End If

            '讀取參數
            Load_Paramater()

            document()

            '外部來不要有頭()-1040115先註記掉~不然新增頁面的TAB會失去分頁功能
            If Request("source") = "land" Or Request("source") = "sland" Or Request("source") = "eland" Then
                'main_menu1.Visible = False '上選單
                'uniquename99.Style("Display") = "none"
                'reveserd1.Visible = False '版權        

                '謄本新增-不需要謄本新增按鈕
                'ImageButton16.Visible = False

                '謄本編號
                Land_FileNo.Text = Request("TrueFILENO")

                '資料來源-謄本種類
                Select Case Request("source")
                    Case "land"
                        Data_Source.Text = "謄本"
                    Case "sland"
                        Data_Source.Text = "二類"
                    Case "eland"
                        Data_Source.Text = "電傳"
                End Select
            End If

            '接Request("src")參數,判斷為過期還現有物件資料表
            If Trim(Request("src")) = "OLD" Then
                src.Text = "委賣物件過期資料表"
            ElseIf Trim(Request("src")) = "NOW" Then
                src.Text = "委賣物件資料表"
            End If

            '判斷物件維護權限
            myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")

            '依參數觸發那一段程式
            If state = "add" Then '新增--------------
                'checkleave.使用分析(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "A", "物件新增")
                'sql = " insert into 房仲家使用分析 "
                'sql += " (店代號,員工編號,大標,名稱,時間) "
                'sql += " select '" & Request.Cookies("store_id").Value.ToString & "', "
                'sql += " '" & Request.Cookies("webfly_empno").Value.ToString & "', "
                'sql += " 'A','物件新增',GETDATE() "
                'cmd = New SqlCommand(sql, conn)
                'cmd.ExecuteNonQuery()
                'TAB------------------------------------------------------------------
                li2.Text = ""

                li2.Text &= "<li><a href=""Obj_Add_V4.aspx?state=add&src=NOW"">新增</a></li>"

                If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                    li2.Text &= "<li id=""current""><a href=""Object_Search_ds.aspx"">查詢</a></li>"
                Else
                    li2.Text &= "<li id=""current""><a href=""Object_Search.aspx"">查詢</a></li>"
                End If
                '---------------------------------------------------------------------

                '隱藏複製選項
                copy = 0

                '顯示物件類別
                objcls = 1

                載入新增初始頁面()

                '新增按鈕
                If myobj.AC = "1" Then
                    Me.ImageButton1.Visible = True
                    If Request("source") = "land" Or Request("source") = "sland" Or Request("source") = "eland" Then
                        '謄本新增-不需要謄本新增按鈕
                        'ImageButton16.Visible = False
                    Else
                        'Me.ImageButton16.Visible = True '謄本新增按鈕
                    End If
                    Me.ImageButton5.Visible = True '細項面積新增按鈕
                    Me.ImageButton7.Visible = True '公園新增按鈕
                    Me.ImageButton8.Visible = True '商圈新增按鈕
                    Me.ImageButton4.Visible = True '他項細項新增按鈕
                Else
                    Me.ImageButton1.Visible = False
                    'Me.ImageButton16.Visible = False
                    Me.ImageButton5.Visible = False
                    Me.ImageButton7.Visible = False
                    Me.ImageButton8.Visible = False
                    Me.ImageButton4.Visible = False
                End If
                '修改按鈕
                Me.ImageButton12.Visible = False
                '複製按鈕
                Me.ImageButton13.Visible = False
                Me.CheckBox26.Visible = False
                '刪除按鈕
                Me.ImageButton19.Visible = False

                '預設管理費單位為坪
                DropDownList5.SelectedValue = "坪"
                'DropDownList24.SelectedValue = "坪"

                '載入可用表單
                Label28.Visible = True
                DropDownList4.Visible = True

                ddl契約類別_SelectedIndexChanged(Nothing, Nothing)

                '銷售狀態-ADD
                DropDownList21.SelectedValue = "熱賣中"
                DropDownList21.Enabled = False
            ElseIf state = "update" Then '修改----------------------
                'checkleave.使用分析(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "A", "物件修改")
                'sql = " insert into 房仲家使用分析 "
                'sql += " (店代號,員工編號,大標,名稱,時間) "
                'sql += " select '" & Request.Cookies("store_id").Value.ToString & "', "
                'sql += " '" & Request.Cookies("webfly_empno").Value.ToString & "', "
                'sql += " 'A','物件修改',GETDATE() "
                'cmd = New SqlCommand(sql, conn)
                'cmd.ExecuteNonQuery()
                'TAB------------------------------------------------------------------
                myobj_tabs.power_object(Request.Cookies("webfly_empno").Value, Request("state"), Request("sid"), Request("oid"), Request("src"), "Update", "sell")
                li2.Text = myobj_tabs.Li_Str_All
                '---------------------------------------------------------------------

                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                '過期物件 磁扣去做清空動作
                sql = "Select * FROM " & src.Text
                sql &= " With(NoLock) where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table2")
                table2 = ds.Tables("table2")
                If table2.Rows.Count > 0 Then
                    If src.Text = "委賣物件過期資料表" Then
                        sql = "Update " & src.Text & " set 磁扣編號='' where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                        cmd = New SqlCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                    ElseIf src.Text = "委賣物件資料表" Then
                        If Not IsDBNull(table2.Rows(0).Item("委託截止日")) Then
                            If table2.Rows(0).Item("委託截止日") < sysdate Then
                                sql = "Update " & src.Text & " set 磁扣編號='' where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            ElseIf table2.Rows(0).Item("銷售狀態") = "已成交" Then
                                sql = "Update " & src.Text & " set 磁扣編號='' where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            ElseIf table2.Rows(0).Item("銷售狀態") = "已解約" Then
                                sql = "Update " & src.Text & " set 磁扣編號='' where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            End If
                        End If
                    End If
                End If
                conn.Close()
                conn.Dispose()

                'If Request.Cookies("webfly_empno").Value = "92H" Then
                '    Response.Write(Request("sid") & "_" & Request("src") & "_" & Request("oid"))
                'End If


                '隱藏複製選項
                copy = 0

                '隱藏物件類別
                objcls = 0

                '讀入物件資料
                載入更新或複製初始頁面()

                '讀入不動產說明書資料
                載入更新或複製初始頁面_不動產說明書()

                ' ''傳給屋主頁面的參數-1031225新增----因修改承辦人後並不會即時變更，固廢掉此段，改以轉頁中判斷
                ''Dim sale_1 As String = "no"
                ''Dim sale_2 As String = "no"
                ''Dim sale_3 As String = "no"

                ''If sale1.SelectedValue.Trim <> "選擇人員" Then
                ''    sale_1 = sale1.SelectedValue.Trim
                ''End If

                ''If sale2.SelectedValue.Trim <> "選擇人員" Then
                ''    sale_2 = sale2.SelectedValue.Trim
                ''End If

                ''If sale3.SelectedValue.Trim <> "選擇人員" Then
                ''    sale_3 = sale3.SelectedValue.Trim
                ''End If

                ''li2.Text = Replace(li2.Text, "參數", "&sale1=" & sale_1 & "&sale2=" & sale_2 & "&sale3=" & sale_3)

                '學區使用的SESSION值
                Session("County") = DDL_County.SelectedValue
                Session("Town") = DDL_Area.SelectedValue

                '新增按鈕
                Me.ImageButton1.Visible = False
                '修改按鈕
                If myobj.M = "1" Then
                    Me.ImageButton12.Visible = True
                    'Me.ImageButton16.Visible = True '謄本新增按鈕
                    Me.ImageButton5.Visible = True '細項面積新增按鈕
                    Me.ImageButton7.Visible = True '公園新增按鈕
                    Me.ImageButton8.Visible = True '商圈新增按鈕
                    Me.ImageButton4.Visible = True '他項細項新增按鈕
                Else
                    Me.ImageButton12.Visible = False
                    Me.ImageButton1.Visible = False
                    'Me.ImageButton16.Visible = False
                    Me.ImageButton5.Visible = False
                    Me.ImageButton7.Visible = False
                    Me.ImageButton8.Visible = False
                    Me.ImageButton4.Visible = False
                End If
                '複製按鈕
                Me.ImageButton13.Visible = False
                Me.CheckBox26.Visible = False

                '刪除按鈕
                If Left(Request("oid"), 1) = "9" Then
                    If myobj.D = "1" Then
                        Me.ImageButton19.Visible = True
                    Else
                        Me.ImageButton19.Visible = False
                    End If
                Else
                    Me.ImageButton19.Visible = False
                End If

                'UPDATE狀態下,契約類別+物件編號無法修改
                ddl契約類別.Enabled = False
                TextBox2.Enabled = False

                'ddl契約類別_SelectedIndexChanged(Nothing, Nothing)

                '載入可用表單-隱藏
                Label28.Visible = False
                DropDownList4.Visible = False

                'UPDATE狀態下,委託起迄日期日曆按鈕隱藏
                seeday0.Visible = False
                seeday1.Visible = False

                '暫停銷售
                CheckBox2.Visible = True

                '讀入URL傳入參數
                Dim sid As String = Request("sid")
                Dim objectnum As String = Request("oid")
                Dim message As String = ""

                '判斷若為謄本轉入的物件,委託起始日開放修改(20100722小豪新增)-------------------------------------------
                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                sql = "Select * FROM " & src.Text
                sql &= " With(NoLock) where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")
                If table1.Rows.Count > 0 Then
                    Dim i As Integer = 0
                    If Not IsDBNull(table1.Rows(i)("資料來源")) Then
                        If Trim(table1.Rows(i)("資料來源")) = "謄本" Or Trim(table1.Rows(i)("資料來源")) = "二類" Or Trim(table1.Rows(i)("資料來源")) = "電傳" Then
                            If Not IsDBNull(table1.Rows(i)("委託起始日")) Then
                                If Trim(table1.Rows(i)("委託起始日")) = "" Then
                                    Me.Date2.Enabled = True
                                Else
                                    Me.Date2.Enabled = False
                                End If
                            Else
                                Me.Date2.Enabled = True
                            End If
                        Else
                            Me.Date2.Enabled = False
                        End If
                    Else
                        Me.Date2.Enabled = False
                    End If

                    '要更新時判斷此物件是否成交或解約
                    '先行判別是否合約終止
                    If Not IsDBNull(table1.Rows(0).Item("合約終止")) Then
                        If table1.Rows(0).Item("合約終止") = "1" Then
                            message = "此物件合約已終止，故不開放修改"
                        End If
                    End If
                    If message = "" Then
                        If Not IsDBNull(table1.Rows(0).Item("委託截止日")) Then
                            If table1.Rows(0).Item("委託截止日") < sysdate Then
                                message = "此物件已過期，故不開放修改 " & sysdate
                            ElseIf table1.Rows(0).Item("銷售狀態") = "已成交" Then
                                message = "此物件已成交，故不開放修改"
                            ElseIf table1.Rows(0).Item("銷售狀態") = "已解約" Then
                                message = "此物件已解約，故不開放修改"
                            End If
                        End If
                    End If

                    '網頁連結 20121228 by nick
                    If Not IsDBNull(table1.Rows(0).Item("index_num")) Then
                        HyperLink1.NavigateUrl = "https://www.etwarm.com.tw/houses/buy/" & table1.Rows(0).Item("index_num")
                        HyperLink1.Text = "https://www.etwarm.com.tw/houses/buy/" & table1.Rows(0).Item("index_num")

                        facebook_share_url = String.Format("https://www.facebook.com/share.php?u={0}", HyperLink1.NavigateUrl)
                        Dim objName = Server.UrlEncode(table1.Rows(0)("建築名稱").ToString)
                        line_share_url = String.Format("https://social-plugins.line.me/lineit/share?url={0}&text={1}&from=line_scheme", HyperLink1.NavigateUrl, objName)
                    End If

                    If message <> "" Then
                        Dim script As String = ""
                        script += "alert('" & message & "');"
                        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

                        '修改按鈕
                        Me.ImageButton12.Visible = False

                        '刪除按鈕
                        If Left(Request("oid"), 1) = "9" Then
                            If myobj.D = "1" Then
                                Me.ImageButton19.Visible = True
                            Else
                                Me.ImageButton19.Visible = False
                            End If
                        Else
                            Me.ImageButton19.Visible = False
                        End If

                        '不能修改狀態-呼叫鎖定頁面控制項-->enable=false
                        enablefalse("lock")

                        '面積細項按鈕
                        Me.ImageButton2.Visible = False
                        Me.ImageButton3.Visible = False
                        Me.ImageButton5.Visible = False

                        '建議商圈.建議公園按鈕
                        Button1.Visible = False
                        Button2.Visible = False

                        '完工日期.登記日期按鈕
                        ch1.Visible = False
                        Button3.Visible = False

                        '國小.國中.高中.大專院校按鈕
                        Button4.Visible = False
                        Button6.Visible = False
                        Button7.Visible = False
                        Button8.Visible = False

                        '他項權利按鈕
                        Me.ImageButton4.Visible = False
                        Me.ImageButton9.Visible = False
                        Me.ImageButton10.Visible = False
                        seeday2.Visible = False

                        '計算土增稅按鈕
                        'Me.ImageButton11.Visible = False

                    End If

                    conn.Close()
                    conn.Dispose()
                End If

                Me.Date2.Enabled = False
                Me.Date3.Enabled = False

                '判斷新增轉頁過來，給予訊息
                If Request("trans") = "true" Then
                    Dim script As String = ""
                    script += "alert('新增成功!!');"
                    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

                End If

                'If Request("oid") = "61318AAD81152" Or Request("oid") = "61318AAD81152-01" Then
                '    Me.ImageButton19.Visible = True
                'End If



            ElseIf state = "copy" Then '修改----------------------
                'checkleave.使用分析(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "A", "物件複製")
                'sql = " insert into 房仲家使用分析 "
                'sql += " (店代號,員工編號,大標,名稱,時間) "
                'sql += " select '" & Request.Cookies("store_id").Value.ToString & "', "
                'sql += " '" & Request.Cookies("webfly_empno").Value.ToString & "', "
                'sql += " 'A','物件複製',GETDATE() "
                'cmd = New SqlCommand(sql, conn)
                'cmd.ExecuteNonQuery()
                'TAB------------------------------------------------------------------
                myobj_tabs.power_object(Request.Cookies("webfly_empno").Value, Request("state"), Request("sid"), Request("oid"), Request("src"), "Copy", "sell")
                li2.Text = myobj_tabs.Li_Str_All
                '---------------------------------------------------------------------
                '顯示複製選項
                copy = 1

                '以藏物件類別
                objcls = 0

                '讀入物件資料
                載入更新或複製初始頁面()

                '讀入不動產說明書資料
                載入更新或複製初始頁面_不動產說明書()

                ''傳給屋主頁面的參數-1031225新增----因修改承辦人後並不會即時變更，固廢掉此段，改以轉頁中判斷
                'Dim sale_1 As String = "no"
                'Dim sale_2 As String = "no"
                'Dim sale_3 As String = "no"

                'If sale1.SelectedValue.Trim <> "選擇人員" Then
                '    sale_1 = sale1.SelectedValue.Trim
                'End If

                'If sale2.SelectedValue.Trim <> "選擇人員" Then
                '    sale_2 = sale2.SelectedValue.Trim
                'End If

                'If sale3.SelectedValue.Trim <> "選擇人員" Then
                '    sale_3 = sale3.SelectedValue.Trim
                'End If

                'li2.Text = Replace(li2.Text, "參數", "&sale1=" & sale_1 & "&sale2=" & sale_2 & "&sale3=" & sale_3)

                '載入可用表單
                Label28.Visible = True
                DropDownList4.Visible = True

                ddl契約類別_SelectedIndexChanged(Nothing, Nothing)


                '學區使用的SESSION值
                Session("County") = DDL_County.SelectedValue
                Session("Town") = DDL_Area.SelectedValue

                '新增按鈕
                Me.ImageButton1.Visible = False
                '修改按鈕              
                Me.ImageButton12.Visible = False
                '複製按鈕
                If myobj.AC = "1" Then
                    Me.ImageButton13.Visible = True
                    Me.CheckBox26.Visible = True
                    'Me.ImageButton16.Visible = True '謄本新增按鈕
                    Me.ImageButton5.Visible = True '細項面積新增按鈕
                    Me.ImageButton7.Visible = True '公園新增按鈕
                    Me.ImageButton8.Visible = True '商圈新增按鈕
                    Me.ImageButton4.Visible = True '他項細項新增按鈕
                Else
                    Me.ImageButton13.Visible = False
                    Me.CheckBox26.Visible = False
                    Me.ImageButton12.Visible = False
                    Me.ImageButton1.Visible = False
                    'Me.ImageButton16.Visible = False
                    Me.ImageButton5.Visible = False
                    Me.ImageButton7.Visible = False
                    Me.ImageButton8.Visible = False
                    Me.ImageButton4.Visible = False
                End If

                '刪除按鈕
                Me.ImageButton19.Visible = False

                '複製時清空截止日期
                Date3.Text = ""
            End If

            'If Request.Cookies("webfly_empno").Value.ToString.ToUpper = "0AEU" Then
            '    If myobj.M = "1" Then
            '        Button10.visible = True
            '    Else
            '        Button10.visible = False
            '    End If
            'End If

        End If
    End Sub
    'ADD
    Public Sub 載入新增初始頁面()
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        '帶入店代號資料
        store.Items.Clear()
        store = clspowerset.scope(Request.Cookies("webfly_empno").Value, "object", store) '店代號
        store.SelectedValue = Request.Cookies("store_id").Value

        If Request("state") <> "update" Then
            store.Enabled = True
        Else
            store.Enabled = False
        End If

        '不動產說明書報告單位
        Dim 店名 As Array = Split(store.SelectedItem.Text, ",")
        input2.Value = 店名(1)

        '判斷有無多店權限，來顯示all_people
        If myobj.Objectmstore = "1" Then
            Me.all_people.Visible = True
        Else
            Me.all_people.Visible = False
        End If

        '帶入人員資料
        If myobj.Objectmstore = "1" Or myobj.Objectstore = "1" Then '多店本店        
            sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
            sql &= "where man_dept_no  = '" & Request.Cookies("store_id").Value & "' and man_quit_dt = ''"
        ElseIf myobj.Objectgroup = "1" Then '本組            
            sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
            sql &= "where man_dept_no = '" & Request.Cookies("store_id").Value & "' and man_quit_dt = '' AND (man_emp_no IN (" & myobj.mgroup_id & "))"
        Else '本人
            sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
            sql &= "where man_dept_no = '" & Request.Cookies("store_id").Value & "' and man_quit_dt = '' AND (man_emp_no ='" & Request.Cookies("webfly_empno").Value & "')"
        End If

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")
        sale1.Items.Clear()
        sale2.Items.Clear()
        sale3.Items.Clear()
        sale1.Items.Add("選擇人員")
        sale2.Items.Add("選擇人員")
        sale3.Items.Add("選擇人員")

        For i = 0 To table2.Rows.Count - 1
            sale1.Items.Add(table2.Rows(i)("man_emp_no").ToString & "," & table2.Rows(i)("man_name").ToString)
            sale1.Items(i + 1).Value = table2.Rows(i)("man_emp_no").ToString
            sale2.Items.Add(table2.Rows(i)("man_emp_no").ToString & "," & table2.Rows(i)("man_name").ToString)
            sale2.Items(i + 1).Value = table2.Rows(i)("man_emp_no").ToString
            sale3.Items.Add(table2.Rows(i)("man_emp_no").ToString & "," & table2.Rows(i)("man_name").ToString)
            sale3.Items(i + 1).Value = table2.Rows(i)("man_emp_no").ToString
        Next

        '帶入輸入者資料
        sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
        sql &= "where man_emp_no = '" & Request.Cookies("webfly_empno").Value & "' "

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")

        If table2.Rows.Count > 0 Then
            Label34.Text = Trim(table2.Rows(0)("man_name"))
            Label33.Text = Trim(table2.Rows(0)("man_emp_no"))
        End If

        '使用分區 --------------------------------------------------------------------------------------   
        sql = "select 使用分區大項 from 資料_使用分區 With(NoLock) group by 使用分區大項 "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        '主
        'DropDownList16.Items.Clear()
        'DropDownList16.Items.Add("請選擇")

        'For i As Integer = 0 To table1.Rows.Count - 1
        '    DropDownList16.Items.Add(table1.Rows(i)("使用分區大項").ToString.Trim)
        '    DropDownList16.Items(i + 1).Value = table1.Rows(i)("使用分區大項").ToString.Trim
        'Next

        ''副
        'DropDownList11.Items.Clear()
        'DropDownList11.Items.Add("請選擇")

        'For i As Integer = 0 To table1.Rows.Count - 1
        '    DropDownList11.Items.Add(table1.Rows(i)("使用分區大項").ToString.Trim)
        '    DropDownList11.Items(i + 1).Value = table1.Rows(i)("使用分區大項").ToString.Trim
        'Next

        '細項
        DropDownList65.Items.Clear()
        DropDownList65.Items.Add("請選擇")

        For i = 0 To table1.Rows.Count - 1
            DropDownList65.Items.Add(table1.Rows(i)("使用分區大項").ToString.Trim)
            DropDownList65.Items(i + 1).Value = table1.Rows(i)("使用分區大項").ToString.Trim
        Next
        '--------------------------------------------------------------------------------------------   

        '物件型態(類別)
        sql = "select * from 資料_物件類別 With(NoLock) where 店代號 = 'A0001'"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        DropDownList3.Items.Clear()
        DropDownList3.Items.Add("請選擇")

        For i = 0 To table1.Rows.Count - 1
            DropDownList3.Items.Add(table1.Rows(i)("實價全名").ToString.Trim)
            DropDownList3.Items(i + 1).Value = Trim(table1.Rows(i)("名稱").ToString.Trim)
        Next

        '建築主要用途-自訂+公有       
        'sql = "select * from 不動產說明書_物件用途 With(NoLock) where 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
            sql = " select DISTINCT 名稱 from 不動產說明書_物件用途 With(NoLock) "
            sql &= " where 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "') "
            sql &= " or 店代號 in (select bs_dept from hsbsmg where bs_直營識別='Y' and bs_state in ('1','7')) "
        Else
            sql = " select DISTINCT 名稱 from 不動產說明書_物件用途 With(NoLock) "
            sql &= " where 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "') "
            sql &= " or 店代號 in "
            sql &= " (select 店代號 from HSSTRUCTURE "
            sql &= " where 組別 in (select 組別 from HSSTRUCTURE where 店代號='" & Request.Cookies("store_id").Value & "')) "
        End If
        'sql = " select DISTINCT 名稱 from 不動產說明書_物件用途 With(NoLock) "
        'sql &= " where 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "') "
        'sql &= " or 店代號 in "
        'sql &= " (select 店代號 from HSSTRUCTURE "
        'sql &= " where 組別 in (select 組別 from HSSTRUCTURE where 店代號='" & Request.Cookies("store_id").Value & "')) "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList19.Items.Clear()
        DropDownList19.Items.Add("")

        For i = 0 To table1.Rows.Count - 1
            DropDownList19.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        Next
        DropDownList19.Items.Add("其他")

        '住址-路
        address20 = publicfuntion.address2(address20)

        '新增時帶入當天日期預設為委託起始日
        Date2.Text = Year(Today) - 1911 & Format(Month(Today), "0#") & Format(Day(Today), "0#")

        '不動產說明書報告日期
        Date5.Text = Year(Today) - 1911 & Format(Month(Today), "0#") & Format(Day(Today), "0#")

        '車位類別
        'sql = "select * from 資料_車位 With(NoLock) where 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        'adpt = New SqlDataAdapter(sql, conn)
        'ds = New DataSet()
        'adpt.Fill(ds, "table1")
        'table1 = ds.Tables("table1")
        'ddl車位類別.Items.Clear()
        'ddl車位類別.Items.Add("請選擇")
        'For i As Integer = 0 To table1.Rows.Count - 1
        '    ddl車位類別.Items.Add(table1.Rows(i)("車位").ToString.Trim)
        'Next
        'ddl車位類別.Items.Add("其他")

        '捷運路線
        sql = "SELECT DISTINCT 路線 FROM 資料_捷運 With(NoLock) "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")
        DropDownList8.Items.Clear()
        DropDownList8.Items.Add("選擇路線")

        For i = 0 To table2.Rows.Count - 1
            '依照選擇的縣市來決定路線
            If DDL_County.SelectedValue = "新北市" Or DDL_County.SelectedValue = "台北市" Then
                If Not (table2.Rows(i)("路線") = "高雄紅線" Or table2.Rows(i)("路線") = "高雄橘線" Or table2.Rows(i)("路線").ToString = "高雄輕軌" Or table2.Rows(i)("路線") = "台中綠線" Or table2.Rows(i)("路線") = "台中藍線") Then
                    DropDownList8.Items.Add(table2.Rows(i)("路線").ToString)
                End If
            ElseIf DDL_County.SelectedValue = "高雄市" Then
                If table2.Rows(i)("路線").ToString = "高雄紅線" Or table2.Rows(i)("路線").ToString = "高雄橘線" Or table2.Rows(i)("路線").ToString = "高雄輕軌" Then
                    DropDownList8.Items.Add(table2.Rows(i)("路線").ToString)
                End If
            ElseIf DDL_County.SelectedValue = "台中市" Then
                If table2.Rows(i)("路線") = "台中綠線" Or table2.Rows(i)("路線") = "台中藍線" Then
                    DropDownList8.Items.Add(table2.Rows(i)("路線"))
                End If
            ElseIf DDL_County.SelectedValue = "桃園市" Then
                If table2.Rows(i)("路線") = "新莊線" Then
                    DropDownList8.Items.Add(table2.Rows(i)("路線"))
                End If
            End If
        Next

        '隨ddl契約類別的選擇，變換網頁刊登及暫停銷售的顯示
        If ddl契約類別.SelectedValue = "一般" Or ddl契約類別.SelectedValue = "流通" Or ddl契約類別.SelectedValue = "同意書" Then
            If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                CheckBox1.Visible = False
                CheckBox1.Checked = False
                CheckBox1.Enabled = False
            Else
                CheckBox1.Visible = True
                CheckBox1.Checked = True
                CheckBox1.Enabled = True
            End If


        ElseIf ddl契約類別.SelectedValue = "專任" Then

            If IS_DIRECTLY_OPERATION And IS_MANAGER_ROLE Then
                CheckBox1.Visible = True
            Else
                CheckBox1.Visible = False
            End If

            If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                CheckBox1.Checked = False
            End If

        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            CheckBox1.Checked = False
            CheckBox1.Enabled = False
        End If

        ''管理組織及管理方式       
        'sql = "select * from 資料_不動產說明書 With(NoLock) where 類別='管理組織及管理方式 ' and 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        'adpt = New SqlDataAdapter(sql, conn)
        'ds = New DataSet()
        'adpt.Fill(ds, "table1")
        'table1 = ds.Tables("table1")
        'DropDownList14.Items.Clear()
        'DropDownList14.Items.Add("")
        'i = 0
        'For i = 0 To table1.Rows.Count - 1
        '    DropDownList14.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        'Next
        'DropDownList14.Items.Add("其他")

        '外牆外飾        
        sql = "select * from 資料_不動產說明書 With(NoLock) where 類別='外牆外飾' and 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList43.Items.Clear()
        DropDownList43.Items.Add("")

        For i = 0 To table1.Rows.Count - 1
            DropDownList43.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        Next
        DropDownList43.Items.Add("其他")

        '地板        
        sql = "select * from 資料_不動產說明書 With(NoLock) where 類別='地板' and 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList44.Items.Clear()
        DropDownList44.Items.Add("")

        For i = 0 To table1.Rows.Count - 1
            DropDownList44.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        Next
        DropDownList44.Items.Add("其他")

        '室內建材        
        sql = "select * from 資料_不動產說明書 With(NoLock) where 類別='室內建材' and 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList48.Items.Clear()
        DropDownList48.Items.Add("")

        For i = 0 To table1.Rows.Count - 1
            DropDownList48.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        Next
        DropDownList48.Items.Add("其他")

        '隔間材料
        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
            sql = " select * from 資料_不動產說明書 With(NoLock) "
            sql += " where 類別='隔間材料' "
            sql += " and (店代號 in ('A0001') or 店代號 in (select bs_dept from hsbsmg where bs_直營識別='Y' and bs_state in ('1','7'))) "
        Else
            sql = " select * from 資料_不動產說明書 With(NoLock) "
            sql += " where 類別='隔間材料' "
            sql += " And 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "') "
        End If

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList49.Items.Clear()
        DropDownList49.Items.Add("")

        For i = 0 To table1.Rows.Count - 1
            DropDownList49.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        Next
        DropDownList49.Items.Add("其他")

        '不動產說明書-重要交易條件
        讀入前次所填之值(Request.Cookies("store_id").Value)

        conn.Close()
        conn.Dispose()

        '取得縣市
        City_Data()

        '中華電信謄本
        If Request("FILENO") <> "" And Request("sid") <> "" Then
            Dim script As String = ""
            script += "alert('注意!!謄本面積細項部份會在物件儲存後才會轉入!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            landcase()
        End If

        '取得面積細項-類別+項目
        類別名稱()
        項目名稱("主建物")

        '判斷為細項面積新增還是修改ㄉ狀態
        If Me.Label4.Text = "0" Then
            Me.ImageButton5.Visible = True
            Me.ImageButton3.Visible = False
        Else
            Me.ImageButton5.Visible = False
            Me.ImageButton3.Visible = True
        End If

        '判斷為他項權利細項新增還是修改ㄉ狀態
        If Me.Label56.Text = "0" Then
            Me.ImageButton4.Visible = True
            Me.ImageButton9.Visible = False
        Else
            Me.ImageButton4.Visible = False
            Me.ImageButton9.Visible = True
        End If

        '學區用-session
        Session("County") = ""
        Session("Town") = ""

        If (checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任") Or (checkleave.直營leave(Request.Cookies("webfly_empno").Value) = "秘書") Then
            Button14.Visible = True
            Button15.Visible = True
        Else
            Button14.Visible = False
            Button15.Visible = False
        End If
        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Button14.visible = True
        '    Button15.visible = True
        'Else
        '    Button14.visible = False
        '    Button15.visible = False
        'End If

    End Sub

    '中華電信謄本新增
    Sub landcase()
        Dim lno, sion As String
        Using conn As New SqlConnection(land)
            conn.Open()
            sql = "select * from saleobject With(NoLock) WHERE 物件編號='" & Request("FILENO") & "' AND 店代號 = '" & Request("sid") & "'"

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.SelectCommand.CommandTimeout = 200
            adpt.Fill(ds, "table1")
            Dim table1 As DataTable = ds.Tables("table1")
            TextBox5.Text = table1.Rows(0)("主建物平方公尺")
            TextBox6.Text = table1.Rows(0)("主建物")
            TextBox7.Text = table1.Rows(0)("附屬建物平方公尺")
            TextBox8.Text = table1.Rows(0)("附屬建物")
            TextBox21.Text = table1.Rows(0)("公設內車位平方公尺")
            TextBox23.Text = table1.Rows(0)("公設內車位坪數")
            TextBox9.Text = table1.Rows(0)("公共設施平方公尺")
            TextBox10.Text = table1.Rows(0)("公共設施")
            TextBox30.Text = table1.Rows(0)("土地平方公尺")
            TextBox31.Text = table1.Rows(0)("土地坪數")
            TextBox28.Text = table1.Rows(0)("總平方公尺")
            TextBox29.Text = table1.Rows(0)("總坪數")
            TextBox17.Text = IIf(table1.Rows(0)("地號") = "地號", "", Replace(table1.Rows(0)("地號"), ",", ""))
            TextBox18.Text = IIf(table1.Rows(0)("建號") = "建號", "", Replace(table1.Rows(0)("建號"), ",", ""))

            'input50.value = table1.Rows(0)("貸款銀行")
            'If Not IsDBNull(table1.Rows(0)("貸款金額")) Then
            '    input80.Value = CInt(table1.Rows(0)("貸款金額"))
            'End If
            'Text5.value = table1.Rows(0)("貸款銀行2")
            'If Not IsDBNull(table1.Rows(0)("貸款金額2")) Then
            '    Text6.Value = CInt(table1.Rows(0)("貸款金額2"))
            'End If
            'Text7.value = table1.Rows(0)("貸款銀行3")
            'If Not IsDBNull(table1.Rows(0)("貸款金額3")) Then
            '    Text8.Value = CInt(table1.Rows(0)("貸款金額3"))
            'End If
            'Text9.value = table1.Rows(0)("貸款銀行4")
            'If Not IsDBNull(table1.Rows(0)("貸款金額4")) Then
            '    Text10.Value = CInt(table1.Rows(0)("貸款金額4"))
            'End If
            Text2.Value = table1.Rows(0)("竣工日期")
            add1.Text = table1.Rows(0)("村里")
            add2.Text = table1.Rows(0)("鄰")
            add3.Text = table1.Rows(0)("路名")
            add4.Text = table1.Rows(0)("段")
            add5.Text = table1.Rows(0)("巷")
            add6.Text = table1.Rows(0)("弄")
            add7.Text = table1.Rows(0)("號")
            add8.Text = table1.Rows(0)("之")
            add10.Text = table1.Rows(0)("樓之")
            DDL_County.SelectedValue = table1.Rows(0)("縣市")

            DDL_County_SelectedIndexChanged(Nothing, Nothing)

            For i = 0 To DDL_Area.Items.Count - 1
                If DDL_Area.Items(i).Text = Trim(table1.Rows(0)("鄉鎮市區")) Then
                    'DDL_Area.SelectedValue = table1.Rows(0)("鄉鎮市區")
                    DDL_Area.SelectedIndex = i
                    Exit For
                End If
            Next

            TB_AreaCode.Text = table1.Rows(0)("郵遞區號")

            Text11.Value = table1.Rows(0)("register_date")
            conn.Close()
            conn.Dispose()
        End Using
    End Sub
    'UPDATE
    Public Sub 載入更新或複製初始頁面()

        Dim sid As String = Request("sid")
        Label12.Text = Request("sid")

        Dim objectnum As String = Request("oid")
        Label11.Text = Request("oid")
        Label57.Text = Request("oid")

        Dim webfly_empno As String = Request("webfly_empno")

        Dim message As String = ""

        '更新使用分區()

        '先呼叫原始該載入的控制項
        載入新增初始頁面()

        '讀取他項權利
        Load_他項權利Data("OLD")

        '讀取細項面積

        Load_Data("OLD")

        'If Request.Cookies("webfly_empno").Value = "92H" Then
        更新圖片相關資訊()
        'End If

        更新車位()

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        更新委託期間()
        'End If

        sql = "Select "
        sql &= "主建物平方公尺,主建物,"
        sql &= "公共設施平方公尺,公共設施 ,"
        sql &= "地下室平方公尺,地下室,"
        sql &= "土地平方公尺,土地坪數,"
        sql &= "附屬建物平方公尺,附屬建物,"
        sql &= "公設內車位平方公尺,公設內車位坪數,"
        sql &= "車位平方公尺,車位坪數,"
        sql &= "總平方公尺,總坪數,"
        sql &= " * FROM " & src.Text & " With(NoLock) where 物件編號 = '" & objectnum & "' and 店代號 = '" & sid & "' "
        'If Request.Cookies("store_id").Value = "A0001" Then

        '    Response.Write(sql)
        '    'Response.End()
        'End If

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")

        If table2.Rows.Count <> 0 Then
            '帶入輸入者資料--------------------------------------------------------
            sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "

            '複製時抓新增者
            If Request("state") = "copy" Then
                sql &= "where man_emp_no = '" & Request.Cookies("webfly_empno").Value & "' "
            Else
                sql &= "where man_emp_no = '" & table2.Rows(0)("輸入者").ToString & "' "
            End If

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table3")
            table3 = ds.Tables("table3")

            '輸入者員編+姓名
            If table3.Rows.Count > 0 Then
                Label34.Text = Trim(table3.Rows(0)("man_name"))
                Label33.Text = Trim(table3.Rows(0)("man_emp_no"))
            End If
            '----------------------------------------------------------------------

            '物件編號
            If Left(table2.Rows(0)("物件編號").ToString, 1) = 1 Then
                ddl契約類別.SelectedValue = "專任"
            ElseIf Left(table2.Rows(0)("物件編號").ToString, 1) = 6 Then
                ddl契約類別.SelectedValue = "一般"
            ElseIf Left(table2.Rows(0)("物件編號").ToString, 1) = 7 Then
                ddl契約類別.SelectedValue = "同意書"
            ElseIf Left(table2.Rows(0)("物件編號").ToString, 1) = 5 Then
                ddl契約類別.SelectedValue = "流通"
            ElseIf Left(table2.Rows(0)("物件編號").ToString, 1) = 9 Then
                ddl契約類別.SelectedValue = "庫存"
            End If

            If ddl契約類別.SelectedValue = "一般" Or ddl契約類別.SelectedValue = "流通" Or ddl契約類別.SelectedValue = "同意書" Then
                If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                    CheckBox1.Visible = False
                    CheckBox1.Checked = False
                Else
                    CheckBox1.Visible = True
                End If
            ElseIf ddl契約類別.SelectedValue = "專任" Then

                If IS_DIRECTLY_OPERATION And IS_MANAGER_ROLE Then
                    CheckBox1.Visible = True
                Else
                    CheckBox1.Visible = False
                End If

                If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                    CheckBox1.Checked = False
                End If
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                CheckBox1.Checked = False
                CheckBox1.Enabled = False
            End If

            ddl契約類別_SelectedIndexChanged(Nothing, Nothing)

            If IsDBNull(table2.Rows(0)("物件編號")) = False Then
                TextBox2.Text = Mid(table2.Rows(0)("物件編號"), 6)
            End If

            '簡碼
            If IsDBNull(table2.Rows(0)("ezcode")) = False Then
                ez_code.Text = Trim(table2.Rows(0)("ezcode"))
            End If

            '長短案名
            input4.Value = table2.Rows(0)("建築名稱").ToString
            Text15.Value = table2.Rows(0)("張貼卡案名").ToString

            '使用型態-----------------------------------------------------------------
            DropDownList3.SelectedValue = table2.Rows(0)("物件類別").ToString
            '如果使用形態為"預售屋",顯示棟別
            If DropDownList3.SelectedValue = "預售屋" Then
                Label58.Visible = True
                add11.Visible = True
                If Not IsDBNull(table2.Rows(0)("棟別")) Then
                    add11.Text = table2.Rows(0)("棟別").ToString
                End If
            Else
                Label58.Visible = False
                add11.Visible = False
                add11.Text = ""
            End If

            If Not IsDBNull(table2.Rows(0)("合約終止")) Then
                If table2.Rows(0)("合約終止").ToString = "1" Then
                    CheckBox100.Checked = True
                    CheckBox100.Enabled = False
                    '    ImageButton20.Enabled = False
                Else
                    CheckBox100.Checked = False
                    CheckBox100.Enabled = True
                    ImageButton20.Enabled = True
                End If
            Else
                CheckBox100.Checked = False
                CheckBox100.Enabled = True
                ImageButton20.Enabled = True
            End If

            '如果是土地則關閉一些房屋稅的選項
            'If Trim(DropDownList3.SelectedValue) = "土地" Then
            '    input105.Disabled = True
            '    input107.Disabled = True
            '    input113.Disabled = True
            '    input116.Disabled = True
            '    input115.Disabled = True
            '    input114.Disabled = True
            'Else
            '    input105.Disabled = False
            '    input107.Disabled = False
            '    input113.Disabled = False
            '    input116.Disabled = False
            '    input115.Disabled = False
            '    input114.Disabled = False
            'End If

            '1040519新增
            Address_change()
            '-------------------------------------------------------------------------

            '物件地址_區域    
            i = 0
            For i = 0 To DDL_County.Items.Count - 1
                If Trim(DDL_County.Items(i).Value) = Trim(table2.Rows(0)("縣市").ToString) Then
                    DDL_County.SelectedIndex = i
                    Exit For
                End If
            Next
            DDL_County_SelectedIndexChanged(Nothing, Nothing)
            DDL_Area.SelectedValue = table2.Rows(0)("郵遞區號").ToString
            TB_AreaCode.Text = table2.Rows(0)("郵遞區號").ToString

            add1.Text = table2.Rows(0)("村里").ToString

            zone3.SelectedValue = table2.Rows(0)("村里別").ToString

            add2.Text = table2.Rows(0)("鄰").ToString

            add3.Text = table2.Rows(0)("路名").ToString

            address20.SelectedValue = table2.Rows(0)("路別").ToString

            add4.Text = table2.Rows(0)("段").ToString
            add5.Text = table2.Rows(0)("巷").ToString

            '如果為自己開發的物件,可以看到完整地址,其它則依權限
            If webfly_empno.Trim = Trim(table2.Rows(0).Item("經紀人代號")) Or webfly_empno.Trim = Trim(table2.Rows(0).Item("營業員代號1")) Or webfly_empno.Trim = Trim(table2.Rows(0).Item("營業員代號2")) Then '判斷是否為本人物件
                add6.Text = table2.Rows(0)("弄").ToString
                add7.Text = table2.Rows(0)("號").ToString
                add8.Text = table2.Rows(0)("之").ToString
                add9.Text = table2.Rows(0)("所在樓層").ToString
                add10.Text = table2.Rows(0)("樓之").ToString
            Else
                myobj.Detail_Power(webfly_empno, "159", "ALL")
                If myobj.AC = "1" Or myobj.M = "1" Then '判斷是否有新增、複制、修改的權限，這三個權限大於看地址的權限
                    myobj.power_object(webfly_empno)
                    If myobj.Object_ALL_Address = "1" Then '查看完整地址
                        add6.Text = table2.Rows(0)("弄").ToString
                        add7.Text = table2.Rows(0)("號").ToString
                        add8.Text = table2.Rows(0)("之").ToString
                        add9.Text = table2.Rows(0)("所在樓層").ToString
                        add10.Text = table2.Rows(0)("樓之").ToString
                    ElseIf myobj.Object_Other_Address = "1" Then
                        If Trim(table2.Rows(0).Item("店代號")) = sid.Trim Then
                            add6.Text = table2.Rows(0)("弄").ToString
                            add7.Text = table2.Rows(0)("號").ToString
                            add8.Text = table2.Rows(0)("之").ToString
                            add9.Text = table2.Rows(0)("所在樓層").ToString
                            add10.Text = table2.Rows(0)("樓之").ToString
                        End If
                    End If

                End If
            End If

            'If Right(table2.Rows(0)("刊登售價").ToString, 5) = ".0000" Then
            '    TextBox12.Text = Int(table2.Rows(0)("刊登售價").ToString)
            'Else
            '    TextBox12.Text = table2.Rows(0)("刊登售價").ToString
            'End If

            Label462.Text = ""
            If Convert.ToDouble("0" & table2.Rows(0)("刊登售價").ToString) > 0 Then
                TextBox12.Text = Convert.ToDouble(table2.Rows(0)("刊登售價").ToString)
                Label462.Text = Convert.ToDouble(table2.Rows(0)("刊登售價").ToString)
            End If

            '格局-房
            If table2.Rows(0)("房").ToString = "-1" Then
                C1.Checked = True
                C1_CheckedChanged(Nothing, Nothing)
            Else
                C1.Checked = False
                If IsDBNull(table2.Rows(0)("房")) = False Then
                    TextBox13.Text = table2.Rows(0)("房").ToString
                End If
            End If

            '格局-廳
            If IsDBNull(table2.Rows(0)("廳")) = False Then
                TextBox14.Text = table2.Rows(0)("廳").ToString
            End If

            '格局-衛
            If IsDBNull(table2.Rows(0)("衛")) = False Then
                TextBox15.Text = Val(table2.Rows(0)("衛"))
            End If

            '格局-室
            If IsDBNull(table2.Rows(0)("室")) = False Then
                TextBox16.Text = table2.Rows(0)("室").ToString
            End If

            '臨路寬
            If IsDBNull(table2.Rows(0)("臨路寬")) = False Then
                TextBox245.Text = table2.Rows(0)("臨路寬").ToString
            End If

            '面寬
            If IsDBNull(table2.Rows(0)("面寬")) = False Then
                TextBox39.Text = table2.Rows(0)("面寬").ToString
            End If

            '縱深
            If IsDBNull(table2.Rows(0)("縱深")) = False Then
                TextBox40.Text = table2.Rows(0)("縱深").ToString
            End If
            ''磁扣配對--複製時不用複製
            If copy = 1 Then
                TextBox267.Text = ""
            Else
                TextBox267.Text = table2.Rows(0)("磁扣編號").ToString
            End If

            '主建物
            If IsDBNull(table2.Rows(0)("主建物平方公尺")) = False Then
                TextBox5.Text = table2.Rows(0)("主建物平方公尺").ToString
                If Trim(TextBox5.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox5, "F")
                End If
                'If Right(table2.Rows(0)("主建物平方公尺").ToString, 5) = ".0000" Then
                '    TextBox5.Text = Int(table2.Rows(0)("主建物平方公尺").ToString)
                'Else
                '    TextBox5.Text = table2.Rows(0)("主建物平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("主建物")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox6.Text = table2.Rows(0)("主建物").ToString
                If Trim(TextBox6.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox6, "T")
                End If

                'If Right(table2.Rows(0)("主建物").ToString, 5) = ".0000" Then
                '    TextBox6.Text = Int(table2.Rows(0)("主建物").ToString)
                'Else
                '    TextBox6.Text = table2.Rows(0)("主建物").ToString
                'End If
            End If

            '公共設施
            If IsDBNull(table2.Rows(0)("公共設施平方公尺")) = False Then
                TextBox9.Text = table2.Rows(0)("公共設施平方公尺").ToString
                If Trim(TextBox9.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox9, "F")
                End If

                'If Right(table2.Rows(0)("公共設施平方公尺").ToString, 5) = ".0000" Then
                '    TextBox9.Text = Int(table2.Rows(0)("公共設施平方公尺").ToString)
                'Else
                '    TextBox9.Text = table2.Rows(0)("公共設施平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("公共設施")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox10.Text = table2.Rows(0)("公共設施").ToString
                If Trim(TextBox10.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox10, "T")
                End If

                'If Right(table2.Rows(0)("公共設施").ToString, 5) = ".0000" Then
                '    TextBox10.Text = Int(table2.Rows(0)("公共設施").ToString)
                'Else
                '    TextBox10.Text = table2.Rows(0)("公共設施").ToString
                'End If
            End If

            '含公設車位坪數
            If IsDBNull(table2.Rows(0)("含公設車位坪數")) = False Then
                If table2.Rows(0)("含公設車位坪數").ToString = "Y" Then
                    Checkbox27.Checked = True
                End If
            End If

            '地下室
            If IsDBNull(table2.Rows(0)("地下室平方公尺")) = False Then
                TextBox19.Text = table2.Rows(0)("地下室平方公尺").ToString
                If Trim(TextBox19.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox19, "F")
                End If

                'If Right(table2.Rows(0)("地下室平方公尺").ToString, 5) = ".0000" Then
                '    TextBox19.Text = Int(table2.Rows(0)("地下室平方公尺").ToString)
                'Else
                '    TextBox19.Text = table2.Rows(0)("地下室平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("地下室")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox20.Text = table2.Rows(0)("地下室").ToString
                If Trim(TextBox20.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox20, "T")
                End If

                'If Right(table2.Rows(0)("地下室").ToString, 5) = ".0000" Then
                '    TextBox20.Text = Int(table2.Rows(0)("地下室").ToString)
                'Else
                '    TextBox20.Text = table2.Rows(0)("地下室").ToString
                'End If
            End If


            '土地
            If IsDBNull(table2.Rows(0)("土地平方公尺")) = False Then
                TextBox30.Text = table2.Rows(0)("土地平方公尺").ToString
                If Trim(TextBox30.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox30, "F")
                End If

                'If Right(table2.Rows(0)("土地平方公尺").ToString, 5) = ".0000" Then
                '    TextBox30.Text = Int(table2.Rows(0)("土地平方公尺").ToString)
                'Else
                '    TextBox30.Text = table2.Rows(0)("土地平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("土地坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox31.Text = table2.Rows(0)("土地坪數").ToString
                If Trim(TextBox31.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox31, "T")
                End If

                'If Right(table2.Rows(0)("土地坪數").ToString, 5) = ".0000" Then
                '    TextBox31.Text = Int(table2.Rows(0)("土地坪數").ToString)
                'Else
                '    TextBox31.Text = table2.Rows(0)("土地坪數").ToString
                'End If
            End If

            '庭院
            If IsDBNull(table2.Rows(0)("庭院平方公尺")) = False Then
                TextBox32.Text = table2.Rows(0)("庭院平方公尺").ToString
                If Trim(TextBox32.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox32, "F")
                End If

                'If Right(table2.Rows(0)("庭院平方公尺").ToString, 5) = ".0000" Then
                '    TextBox32.Text = Int(table2.Rows(0)("庭院平方公尺").ToString)
                'Else
                '    TextBox32.Text = table2.Rows(0)("庭院平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("庭院坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox33.Text = table2.Rows(0)("庭院坪數").ToString
                If Trim(TextBox33.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox33, "T")
                End If
                'If Right(table2.Rows(0)("庭院坪數").ToString, 5) = ".0000" Then
                '    TextBox33.Text = Int(table2.Rows(0)("庭院坪數").ToString)
                'Else
                '    TextBox33.Text = table2.Rows(0)("庭院坪數").ToString
                'End If
            End If




            '附屬建物
            If IsDBNull(table2.Rows(0)("附屬建物平方公尺")) = False Then
                TextBox7.Text = table2.Rows(0)("附屬建物平方公尺").ToString
                If Trim(TextBox7.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox7, "F")
                End If

                'If Right(table2.Rows(0)("附屬建物平方公尺").ToString, 5) = ".0000" Then
                '    TextBox7.Text = Int(table2.Rows(0)("附屬建物平方公尺").ToString)
                'Else
                '    TextBox7.Text = table2.Rows(0)("附屬建物平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("附屬建物")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox8.Text = table2.Rows(0)("附屬建物").ToString
                If Trim(TextBox8.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox8, "T")
                End If
                'If Right(table2.Rows(0)("附屬建物").ToString, 5) = ".0000" Then
                '    TextBox8.Text = Int(table2.Rows(0)("附屬建物").ToString)
                'Else
                '    TextBox8.Text = table2.Rows(0)("附屬建物").ToString
                'End If
            End If

            '公設內車位
            If IsDBNull(table2.Rows(0)("公設內車位平方公尺")) = False Then
                TextBox21.Text = table2.Rows(0)("公設內車位平方公尺").ToString
                If Trim(TextBox21.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox21, "F")
                End If
                'If Right(table2.Rows(0)("公設內車位平方公尺").ToString, 5) = ".0000" Then
                '    TextBox21.Text = Int(table2.Rows(0)("公設內車位平方公尺").ToString)
                'Else
                '    TextBox21.Text = table2.Rows(0)("公設內車位平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("公設內車位坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox23.Text = table2.Rows(0)("公設內車位坪數").ToString
                If Trim(TextBox23.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox23, "T")
                End If

                'If Right(table2.Rows(0)("公設內車位坪數").ToString, 5) = ".0000" Then
                '    TextBox23.Text = Int(table2.Rows(0)("公設內車位坪數").ToString)
                'Else
                '    TextBox23.Text = table2.Rows(0)("公設內車位坪數").ToString
                'End If
            End If

            '產權獨立車位
            If IsDBNull(table2.Rows(0)("車位平方公尺")) = False Then
                TextBox26.Text = table2.Rows(0)("車位平方公尺").ToString
                If Trim(TextBox26.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox26, "F")
                End If

                'If Right(table2.Rows(0)("車位平方公尺").ToString, 5) = ".0000" Then
                '    TextBox26.Text = Int(table2.Rows(0)("車位平方公尺").ToString)
                'Else
                '    TextBox26.Text = table2.Rows(0)("車位平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("車位坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox27.Text = table2.Rows(0)("車位坪數").ToString
                If Trim(TextBox27.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox27, "T")
                    Label77.Text = "(E)公設車位"
                    Label78.Visible = True
                    UpdatePanel17.Visible = True
                End If

                'If Right(table2.Rows(0)("車位坪數").ToString, 5) = ".0000" Then
                '    TextBox27.Text = Int(table2.Rows(0)("車位坪數").ToString)
                'Else
                '    TextBox27.Text = table2.Rows(0)("車位坪數").ToString
                'End If
            End If

            '增建
            If IsDBNull(table2.Rows(0)("加蓋平方公尺")) = False Then
                TextBox34.Text = table2.Rows(0)("加蓋平方公尺").ToString
                If Trim(TextBox34.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox34, "F")
                End If

                'If Right(table2.Rows(0)("加蓋平方公尺").ToString, 5) = ".0000" Then
                '    TextBox34.Text = Int(table2.Rows(0)("加蓋平方公尺").ToString)
                'Else
                '    TextBox34.Text = table2.Rows(0)("加蓋平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("加蓋坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox35.Text = table2.Rows(0)("加蓋坪數").ToString
                If Trim(TextBox35.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox35, "T")
                End If

                'If Right(table2.Rows(0)("加蓋坪數").ToString, 5) = ".0000" Then
                '    TextBox35.Text = Int(table2.Rows(0)("加蓋坪數").ToString)
                'Else
                '    TextBox35.Text = table2.Rows(0)("加蓋坪數").ToString
                'End If
            End If


            '1040415新增有無增建
            If IsDBNull(table2.Rows(0)("增建")) = False Then
                If table2.Rows(0)("增建") = "Y" Then
                    RadioButtonList3.SelectedIndex = 0
                ElseIf table2.Rows(0)("增建") = "N" Then
                    RadioButtonList3.SelectedIndex = 1
                End If
            Else '若為NULL值，改判段加蓋平方公尺是否有值
                If IsDBNull(table2.Rows(0)("加蓋平方公尺")) = False Then
                    If table2.Rows(0)("加蓋平方公尺") <> 0 Then
                        RadioButtonList3.SelectedIndex = 0
                    Else
                        RadioButtonList3.SelectedIndex = 1
                    End If
                Else
                    RadioButtonList3.SelectedIndex = 1
                End If
            End If

            '總坪數
            If IsDBNull(table2.Rows(0)("總平方公尺")) = False Then
                TextBox28.Text = table2.Rows(0)("總平方公尺").ToString
                If Trim(TextBox28.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox28, "F")
                End If

                'If Right(table2.Rows(0)("總平方公尺").ToString, 5) = ".0000" Then
                '    TextBox28.Text = Int(table2.Rows(0)("總平方公尺").ToString)
                'Else
                '    TextBox28.Text = table2.Rows(0)("總平方公尺").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("總坪數")) = False Then
                '1040522修正-新增自訂小數點位數判斷-坪
                TextBox29.Text = table2.Rows(0)("總坪數").ToString
                If Trim(TextBox29.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox29, "T")
                End If
                'If Right(table2.Rows(0)("總坪數").ToString, 5) = ".0000" Then
                '    TextBox29.Text = Int(table2.Rows(0)("總坪數").ToString)
                'Else
                '    TextBox29.Text = table2.Rows(0)("總坪數").ToString
                'End If
            End If

            If IsDBNull(table2.Rows(0)("地上層數")) = False Then
                TextBox88.Text = table2.Rows(0)("地上層數").ToString
            End If
            If IsDBNull(table2.Rows(0)("地下層數")) = False Then
                TextBox89.Text = table2.Rows(0)("地下層數").ToString
            End If
            If IsDBNull(table2.Rows(0)("銷售樓層")) = False Then
                TextBox90.Text = table2.Rows(0)("銷售樓層").ToString
            End If

            If IsDBNull(table2.Rows(0)("座向")) = False Then
                DropDownList22.SelectedValue = table2.Rows(0)("座向").ToString
            End If

            If IsDBNull(table2.Rows(0)("每層戶數")) = False Then
                TextBox91.Text = table2.Rows(0)("每層戶數").ToString
            End If
            If IsDBNull(table2.Rows(0)("每層電梯數")) = False Then
                TextBox92.Text = table2.Rows(0)("每層電梯數").ToString
            End If

            If IsDBNull(table2.Rows(0)("竣工日期")) = False Then
                Text2.Value = table2.Rows(0)("竣工日期").ToString
            End If

            If IsDBNull(table2.Rows(0)("管理費")) = False Then
                If Right(table2.Rows(0)("管理費").ToString, 5) = ".0000" Then
                    TextBox36.Text = Int(table2.Rows(0)("管理費").ToString)
                Else
                    TextBox36.Text = table2.Rows(0)("管理費").ToString
                End If
            End If

            If IsDBNull(table2.Rows(0)("管理費單位")) = False Then
                DropDownList5.SelectedValue = table2.Rows(0)("管理費單位").ToString
            End If

            If IsDBNull(table2.Rows(0)("管理費繳交方式")) = False Then
                TextBox266.Text = table2.Rows(0)("管理費繳交方式").ToString
            End If

            'If IsDBNull(table2.Rows(0)("車位管理費類別")) = False Then
            '    DropDownList25.SelectedValue = table2.Rows(0)("車位管理費類別").ToString
            'End If

            'If IsDBNull(table2.Rows(0)("車位管理費")) = False Then
            '    TextBox94.Text = table2.Rows(0)("車位管理費").ToString
            'End If

            'If IsDBNull(table2.Rows(0)("車位管理費")) = False Then
            '    DropDownList24.SelectedValue = table2.Rows(0)("車位管理費單位").ToString
            'End If

            'input50.Value = table2.Rows(0)("貸款銀行").ToString
            'If Not IsDBNull(table2.Rows(0)("貸款金額")) Then
            '    input80.Value = CInt(table2.Rows(0)("貸款金額"))
            'End If
            'Text5.Value = table2.Rows(0)("貸款銀行2").ToString
            'If Not IsDBNull(table2.Rows(0)("貸款金額2")) Then
            '    Text6.Value = CInt(table2.Rows(0)("貸款金額2"))
            'End If
            'Text7.Value = table2.Rows(0)("貸款銀行3").ToString
            'If Not IsDBNull(table2.Rows(0)("貸款金額3")) Then
            '    Text8.Value = CInt(table2.Rows(0)("貸款金額3"))
            'End If
            'Text9.Value = table2.Rows(0)("貸款銀行4").ToString
            'If Not IsDBNull(table2.Rows(0)("貸款金額4")) Then
            '    Text10.Value = CInt(table2.Rows(0)("貸款金額4"))
            'End If

            'If IsDBNull(table2.Rows(0)("車位租售")) = False Then
            '    If table2.Rows(0)("車位租售").ToString = "租" Then
            '        DropDownList6.Items(1).Selected = True
            '    ElseIf table2.Rows(0)("車位租售").ToString = "售" Then
            '        DropDownList6.Items(2).Selected = True
            '    End If
            'End If

            'If IsDBNull(table2.Rows(0)("車位數量")) = False Then
            '    input53.Value = table2.Rows(0)("車位數量").ToString
            'End If
            'If IsDBNull(table2.Rows(0)("車位說明")) = False Then
            '    TextBox93.Text = table2.Rows(0)("車位說明").ToString
            'End If
            If IsDBNull(table2.Rows(0)("車位價格")) = False Then
                input55.Value = table2.Rows(0)("車位價格").ToString

                If input55.Value = "含於開價中" Then
                    CheckBox3.Checked = True
                End If
            End If

            If Request("state") = "update" Or Request("state") = "copy" Then
                TextBox17.Text = table2.Rows(0)("土地標示").ToString
                TextBox18.Text = table2.Rows(0)("建號").ToString
            End If

            If IsDBNull(table2.Rows(0)("公車站名")) = False Then
                input60.Value = table2.Rows(0)("公車站名").ToString
            End If

            If IsDBNull(table2.Rows(0)("委託起始日")) = False Then
                If myobj.M = "1" Then
                    Date2.Text = table2.Rows(0)("委託起始日").ToString
                Else
                    If ShowDt_Start.Text = "1" Then
                        Date2.Text = table2.Rows(0)("委託起始日").ToString
                    Else
                        Date2.Text = ""
                    End If
                End If
            End If

            If Request("state") = "update" Or Len(RadioButtonList1.SelectedValue) <> 0 Then '避免複製忘記改截止日變過期物件
                If IsDBNull(table2.Rows(0)("委託截止日")) = False Then
                    If myobj.M = "1" Then
                        Date3.Text = table2.Rows(0)("委託截止日").ToString
                    Else
                        If ShowDt.Text = "1" Then
                            Date3.Text = table2.Rows(0)("委託截止日").ToString
                        Else
                            Date3.Text = ""
                        End If
                    End If

                End If
            End If

            If IsDBNull(table2.Rows(0)("公設比")) = False Then
                TextBox41.Text = table2.Rows(0)("公設比").ToString
            End If

            '阿甘新增委託起始日六個月後不得在網站上曝光(根據20111101店東顧問團決議)--新版取消1030830
            'If Not IsDBNull(table2.Rows(0)("委託起始日")) Then
            '    Dim netdate = DateSerial(Left(table2.Rows(0)("委託起始日").ToString.Trim, 3) + 1911, Mid(table2.Rows(0)("委託起始日").ToString.Trim, 4, 2) + 6, Right(table2.Rows(0)("委託起始日").ToString.Trim, 2))
            '    Label24.Text = Trim(Year(netdate) - 1911).PadLeft(3, "0") & Trim(Month(netdate)).PadLeft(2, "0") & Trim(Day(netdate)).PadLeft(2, "0")
            'End If

            '1000325小豪新增-相片資料夾區分為過期(expired).有效(available)-----------------------------------------------------------------------
            'If Not IsDBNull(table2.Rows(0)("委託截止日")) Then
            '    If table2.Rows(0)("委託截止日") >= (Format(Year(Now.AddDays(-7)) - 1911, "00#") & Format(Month(Now.AddDays(-7)), "0#") & Format(Day(Now.AddDays(-7)), "0#")) Then
            '        Me.Label22.Text = "available"
            '        Me.Label62.Text = "expired"
            '    Else
            '        Me.Label22.Text = "expired"
            '        Me.Label62.Text = "available"
            '    End If
            'Else
            '    Me.Label22.Text = "available"
            '    Me.Label62.Text = "expired"
            'End If
            If Trim(Request.QueryString("src")) = "OLD" Then
                Me.Label22.Text = "expired"
                Me.Label62.Text = "available"
            ElseIf Trim(Request("src")) = "NOW" Then
                Me.Label22.Text = "available"
                Me.Label62.Text = "expired"
            End If

            '------------------------------------------------------------------------------------------------------------------------------------
            If Not IsDBNull(table2.Rows(0)("鑰匙編號")) Then
                input66.Value = table2.Rows(0)("鑰匙編號").ToString
            End If

            CheckBox102.Checked = False
            CheckBox103.Checked = False
            CheckBox102.Visible = False
            CheckBox103.Visible = False
            If IsDBNull(table2.Rows(0)("社區養寵")) = False Then
                If table2.Rows(0)("社區養寵") = "1" Then
                    RadioButton3.Checked = False
                    RadioButton4.Checked = True
                    CheckBox102.Visible = True
                    CheckBox103.Visible = True
                    If IsDBNull(table2.Rows(0)("養貓")) = False Then
                        If table2.Rows(0)("養貓") = "1" Then
                            CheckBox102.Checked = True
                        End If
                    End If
                    If IsDBNull(table2.Rows(0)("養狗")) = False Then
                        If table2.Rows(0)("養狗") = "1" Then
                            CheckBox103.Checked = True
                        End If
                    End If
                Else
                    RadioButton3.Checked = True
                    RadioButton4.Checked = False
                End If
            Else
                RadioButton3.Checked = True
                RadioButton4.Checked = False
            End If

            If table2.Rows(0)("網頁刊登").ToString = "是" Then
                CheckBox1.Checked = True
            Else
                CheckBox1.Checked = False
            End If
            '強銷物件 20151228 by nick
            If Not IsDBNull(table2.Rows(0)("強銷物件")) Then
                If table2.Rows(0)("強銷物件").ToString = "是" Then
                    CheckBox5.Checked = True
                Else
                    CheckBox5.Checked = False
                End If
            End If
            If table2.Rows(0)("上傳註記").ToString = "D" Then
                CheckBox2.Checked = True
                '20120109新增-"暫停銷售"異動判斷用
                Me.Label25.Text = "True"
            Else
                CheckBox2.Checked = False
                '20120109新增-"暫停銷售"異動判斷用
                Me.Label25.Text = "False"
            End If

            If Not IsDBNull(table2.Rows(0)("訴求重點")) Then
                If table2.Rows(0)("訴求重點").ToString <> "" Then
                    '    TextBox102.Text = "建議填寫本區域有特色以及較吸引客戶的文字，能大大提升物件在網路曝光的成效喔！"
                    'Else
                    TextBox102.Text = table2.Rows(0)("訴求重點").ToString
                End If
            End If

            If Not IsDBNull(table2.Rows(0)("備註")) Then
                input79.Value = table2.Rows(0)("備註").ToString
            End If

            ''使用分區-主
            'Select Case Trim(table2.Rows(0)("物件用途").ToString)
            '    Case "住一", "住一之一", "住二", "住二之一", "住二之二", "住三", "住三之一", "住三之二", "住四", "住四之一", "住五", "特定住宅區(二)", "第一種住宅區"
            '        DropDownList16.SelectedValue = "住宅區"
            '        DropDownList16_SelectedIndexChanged(Nothing, Nothing)
            '        DropDownList17.SelectedValue = Trim(table2.Rows(0)("物件用途").ToString)
            '    Case "商一", "商二", "商三", "商四", "商五"
            '        DropDownList16.SelectedValue = "商業區"
            '        DropDownList16_SelectedIndexChanged(Nothing, Nothing)
            '        DropDownList17.SelectedValue = Trim(table2.Rows(0)("物件用途").ToString)
            '    Case "工二", "工三", "特種工業區", "甲種工業區", "乙種工業區", "零星工業區"
            '        DropDownList16.SelectedValue = "工業區"
            '        DropDownList16_SelectedIndexChanged(Nothing, Nothing)
            '        DropDownList17.SelectedValue = Trim(table2.Rows(0)("物件用途").ToString)
            '    Case "甲種建築用地", "乙種建築用地", "丙種建築用地", "丁種建築用地"
            '        DropDownList16.SelectedValue = "非都市計畫區"
            '        DropDownList16_SelectedIndexChanged(Nothing, Nothing)
            '        DropDownList17.SelectedValue = Trim(table2.Rows(0)("物件用途").ToString)
            '    Case Else
            '        i = 0
            '        For i = 0 To DropDownList16.Items.Count - 1
            '            If DropDownList16.Items(i).Value = Trim(table2.Rows(0)("物件用途").ToString) Then
            '                DropDownList16.SelectedIndex = i
            '                Exit For
            '            End If

            '        Next

            'End Select

            ''使用分區 -副
            'Me.TextBox253.Text = Trim(table2.Rows(0)("其他使用分區").ToString)

            ''不動產說明書的使用分區
            'Label47.Text = table2.Rows(0)("物件用途").ToString

            'If Not IsDBNull(table2.Rows(0)("其他使用分區")) Then
            '    If table2.Rows(0)("其他使用分區").ToString.Trim <> "" Then
            '        Label47.Text &= "," & Trim(table2.Rows(0)("其他使用分區").ToString)
            '    End If
            'End If

            'If Not IsDBNull(table2.Rows(0)("車位")) Then
            '    If table2.Rows(0)("車位").ToString.Trim = "有" Then
            '        Checkbox27.Checked = True
            '    End If
            'End If

            '    i = 0
            '    For i = 0 To ddl車位類別.Items.Count - 1
            '        If ddl車位類別.Items(i).Value = Trim(table2.Rows(0)("車位").ToString) Then
            '            ddl車位類別.SelectedIndex = i
            '            Exit For
            '        End If

            '    Next
            'End If

            '先呼叫所屬區域的社區大樓資料
            社區大樓()

            If Not IsDBNull(table2.Rows(0)("大樓代號")) Then
                If ddl社區大樓.Items.Count - 1 > 0 Then
                    For i = 0 To ddl社區大樓.Items.Count - 1
                        If Trim(ddl社區大樓.Items(i).Value) = Trim(table2.Rows(0)("大樓代號")) Then
                            ddl社區大樓.Items(i).Selected = True
                        End If
                    Next
                End If
            End If

            '店代號
            store.SelectedValue = table2.Rows(0)("店代號").ToString
            'store.Enabled = False

            '經紀人代號
            For i = 0 To sale1.Items.Count - 1
                splitarray = Split(sale1.Items(i).Value, ",")
                If splitarray(0) = table2.Rows(0)("經紀人代號").ToString Then
                    sale1.SelectedIndex = i
                End If
            Next

            If table2.Rows(0)("經紀人代號").ToString <> "" And sale1.SelectedIndex = 0 Then
                sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
                If myobj.Objectmstore = "1" Then '掛名的經紀人或營業員不在同一樣店,多店時會發生      
                    sql &= "where man_dept_no IN (" & myobj.mstore_id & ") and man_emp_no = '" & table2.Rows(0)("經紀人代號").ToString & "' "
                Else '有填資料但沒有選擇到表示 => 已離職員工                   
                    sql &= "where man_dept_no  = '" & sid & "' and man_emp_no = '" & table2.Rows(0)("經紀人代號").ToString & "' "
                End If

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table9")
                Dim table9 As DataTable = ds.Tables("table9")

                If table9.Rows.Count Then
                    sale1.Items.Insert(1, table9.Rows(0)("man_emp_no") & "," & table9.Rows(0)("man_name"))
                    sale1.SelectedIndex = 1
                End If
            End If

            '判斷是否異動的原始值-經紀人代號
            Label59.Text = Trim(sale1.SelectedValue)

            '營業員代號1
            For i = 0 To sale2.Items.Count - 1
                splitarray = Split(sale2.Items(i).Value, ",")
                If splitarray(0) = table2.Rows(0)("營業員代號1").ToString Then
                    sale2.SelectedIndex = i
                End If
            Next

            If table2.Rows(0)("營業員代號1").ToString <> "" And sale2.SelectedIndex = 0 Then
                sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
                If myobj.Objectmstore = "1" Then '掛名的經紀人或營業員不在同一樣店,多店時會發生 
                    sql &= "where man_dept_no in (" & myobj.mstore_id & ") and man_emp_no = '" & table2.Rows(0)("營業員代號1").ToString & "' "
                Else '有填資料但沒有選擇到表示 => 已離職員工                
                    sql &= "where man_dept_no = '" & sid & "' and man_emp_no = '" & table2.Rows(0)("營業員代號1").ToString & "' "
                End If

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table9")
                Dim table9 As DataTable = ds.Tables("table9")
                If table9.Rows.Count Then
                    sale2.Items.Insert(1, table9.Rows(0)("man_emp_no") & "," & table9.Rows(0)("man_name"))
                    sale2.SelectedIndex = 1
                End If
            End If

            '判斷是否異動的原始值-營業員代號1
            Label60.Text = Trim(sale2.SelectedValue)

            '營業員代號2
            For i = 0 To sale3.Items.Count - 1
                splitarray = Split(sale3.Items(i).Value, ",")
                If splitarray(0) = table2.Rows(0)("營業員代號2").ToString Then
                    sale3.SelectedIndex = i
                End If
            Next

            If table2.Rows(0)("營業員代號2").ToString <> "" And sale3.SelectedIndex = 0 Then
                sql = "SELECT man_emp_no,man_name FROM psman With(NoLock) "
                If myobj.Objectmstore = "1" Then '掛名的經紀人或營業員不在同一樣店,多店時會發生     
                    sql &= "where man_dept_no in  (" & myobj.mstore_id & ") and man_emp_no = '" & table2.Rows(0)("營業員代號2").ToString & "' "
                Else      '有填資料但沒有選擇到表示 => 已離職員工                  
                    sql &= "where man_dept_no = '" & sid & "' and man_emp_no = '" & table2.Rows(0)("營業員代號2").ToString & "' "
                End If

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table9")
                Dim table9 As DataTable = ds.Tables("table9")
                If table9.Rows.Count Then
                    sale3.Items.Insert(1, table9.Rows(0)("man_emp_no") & "," & table9.Rows(0)("man_name"))
                    sale3.SelectedIndex = 1
                End If
            End If

            '判斷是否異動的原始值-營業員代號2
            Label61.Text = Trim(sale3.SelectedValue)

            If Not IsDBNull(table2.Rows(0)("銷售狀態")) Then
                DropDownList21.SelectedValue = table2.Rows(0)("銷售狀態").ToString
            End If
            DropDownList21.Enabled = False

            If Not IsDBNull(table2.Rows(0)("帶看方式")) Then
                If table2.Rows(0)("帶看方式").ToString = "洽經紀人" Then
                    DropDownList20.SelectedIndex = 0
                ElseIf table2.Rows(0)("帶看方式").ToString = "洽營業員" Then
                    DropDownList20.SelectedIndex = 1
                ElseIf table2.Rows(0)("帶看方式").ToString = "洽管理員" Then
                    DropDownList20.SelectedIndex = 2
                ElseIf table2.Rows(0)("帶看方式").ToString = "鑰匙" Then
                    DropDownList20.SelectedIndex = 3
                ElseIf table2.Rows(0)("帶看方式").ToString = "密碼鎖" Then
                    DropDownList20.SelectedIndex = 4
                ElseIf table2.Rows(0)("帶看方式").ToString = "需帶看" Then
                    DropDownList20.SelectedIndex = 5
                ElseIf table2.Rows(0)("帶看方式").ToString = "約看" Then
                    DropDownList20.SelectedIndex = 6
                ElseIf table2.Rows(0)("帶看方式").ToString = "自由" Then
                    DropDownList20.SelectedIndex = 7
                ElseIf table2.Rows(0)("帶看方式").ToString = "自住" Then
                    DropDownList20.SelectedIndex = 8
                End If
            End If

            '國小
            If IsDBNull(table2.Rows(0)("國小")) = False And IsDBNull(table2.Rows(0)("國小1")) = False Then
                If table2.Rows(0)("國小") <> "" Or table2.Rows(0)("國小1") <> "" Then
                    '名稱
                    TextBox98.Text = table2.Rows(0)("國小1")
                    '代號(隱藏)
                    TextBox246.Text = table2.Rows(0)("國小")
                End If
            End If

            '國中
            If IsDBNull(table2.Rows(0)("國中")) = False And IsDBNull(table2.Rows(0)("國中1")) = False Then
                If table2.Rows(0)("國中") <> "" Or table2.Rows(0)("國中1") <> "" Then
                    '名稱
                    TextBox99.Text = table2.Rows(0)("國中1")
                    '代號(隱藏)
                    TextBox247.Text = table2.Rows(0)("國中")
                End If
            End If

            '高中
            If IsDBNull(table2.Rows(0)("高中")) = False And IsDBNull(table2.Rows(0)("高中1")) = False Then
                If table2.Rows(0)("高中") <> "" Or table2.Rows(0)("高中1") <> "" Then
                    '名稱
                    TextBox100.Text = table2.Rows(0)("高中1")
                    '代號(隱藏)
                    TextBox248.Text = table2.Rows(0)("高中")
                End If
            End If

            '大專院校
            '%%%%%%%%%%%%% 2016.03.21 Modify by Finch %%%%%%%%%%%%%
            '處理有多筆大專院校時資料撈不出來的問題，目前已有大專院校代號，故直接取欄位值。
            If IsDBNull(table2.Rows(0)("大專院校")) = False And IsDBNull(table2.Rows(0)("其他學校")) = False Then
                If table2.Rows(0)("大專院校") <> "" Or table2.Rows(0)("其他學校") <> "" Then
                    '名稱
                    TextBox101.Text = table2.Rows(0)("其他學校")
                    '代號(隱藏)
                    TextBox249.Text = table2.Rows(0)("大專院校")
                End If
            End If
            '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            ''因單機版沒有存大專院校的代號,所以從字串來找出代號
            'If IsDBNull(table2.Rows(0)("其他學校")) = False Then
            '    If table2.Rows(0)("其他學校") <> "" Then
            '        '******* 2010514修改 by 奕凱
            '        Dim schstr As String = table2.Rows(0)("其他學校")
            '        If Right(schstr, 1) = "." Then
            '            schstr = schstr.Remove(schstr.LastIndexOf("."), 1)
            '        End If

            '        sql = "Select * FROM 資料_大專院校 With(NoLock) where 校名 = '" & schstr & "' "
            '        '**************
            '        adpt = New SqlDataAdapter(sql, conn)
            '        ds = New DataSet()
            '        adpt.Fill(ds, "table3")
            '        table3 = ds.Tables("table3")

            '        If table3.Rows.Count <> 0 Then
            '            '名稱
            '            TextBox101.Text = table3.Rows(0)("校名")
            '            '代號(隱藏)
            '            TextBox249.Text = table3.Rows(0)("代號")

            '        End If

            '    End If
            'End If

            Dim s, j, k, m As Integer
            '公園
            If IsDBNull(table2.Rows(0)("公園代號")) = False Then
                If table2.Rows(0)("公園代號") <> "" Then
                    '為了舊資料相容,要再加上.
                    s = InStr(table2.Rows(0)("公園代號"), ".")
                    If s = 0 Then
                        table2.Rows(0)("公園代號") = table2.Rows(0)("公園代號") + "."
                    End If
                    '代號(隱藏)
                    TextBox251.Text = table2.Rows(0)("公園代號")

                    Dim item1 As Array = Split(table2.Rows(0)("公園代號"), ".")
                    Dim temp As String = ""
                    For i = 0 To item1.Length - 1
                        If item1(i) <> "" Then
                            temp &= "'" & item1(i) & "',"
                            k += 1
                        End If
                    Next i
                    Dim cc As String = ""
                    cc = Mid(temp, 1, Len(temp) - 1)

                    sql = "Select 公園名稱 FROM 公園資料表 With(NoLock) where 公園代號 in (" & cc & ") and 核准否='1'"
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table4")
                    table4 = ds.Tables("table4")
                    i = 0

                    If table4.Rows.Count > 0 Then
                        For i = 0 To table4.Rows.Count - 1
                            If table4.Rows(i)("公園名稱") <> "" And IsDBNull(table4.Rows(i)("公園名稱")) = False Then
                                '名稱
                                TextBox97.Text &= table4.Rows(i)("公園名稱") & "."
                            End If
                        Next i
                    End If
                End If
            End If

            '商圈
            If IsDBNull(table2.Rows(0)("商圈代號")) = False Then
                If table2.Rows(0)("商圈代號") <> "" Then
                    '為了舊資料相容,要再加上.
                    j = InStr(table2.Rows(0)("商圈代號"), ".")
                    If j = 0 Then
                        table2.Rows(0)("商圈代號") = table2.Rows(0)("商圈代號") + "."
                    End If

                    '代號(隱藏)
                    TextBox250.Text = table2.Rows(0)("商圈代號")

                    Dim item1 As Array = Split(table2.Rows(0)("商圈代號"), ".")
                    Dim temp As String = ""
                    For i = 0 To item1.Length - 1
                        If item1(i) <> "" Then
                            temp &= "'" & item1(i) & "',"
                            m += 1
                        End If
                    Next i
                    Dim cc As String = ""
                    cc = Mid(temp, 1, Len(temp) - 1)

                    sql = "Select 商圈名稱 FROM 精華生活圈資料表 With(NoLock) where 商圈代號 in (" & cc & ") and 核准否='1'"

                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table4")
                    table4 = ds.Tables("table4")
                    If table4.Rows.Count > 0 Then
                        i = 0
                        For i = 0 To table4.Rows.Count - 1
                            If table4.Rows(i)("商圈名稱") <> "" And IsDBNull(table4.Rows(i)("商圈名稱")) = False Then
                                '名稱
                                TextBox96.Text &= table4.Rows(i)("商圈名稱") & "."
                            End If
                        Next i
                    End If
                End If
            End If

            If Not IsDBNull(table2.Rows(0)("聯賣")) Then
                If table2.Rows(0)("聯賣").ToString = "Y" Then
                    CheckBox101.Checked = True
                    Label465.Text = "Y"
                Else
                    CheckBox101.Checked = False
                    Label465.Text = "N"
                End If
            Else
                CheckBox101.Checked = False
                Label465.Text = "N"
            End If

            '捷運
            If Not IsDBNull(table2.Rows(0)("捷運")) Then
                sql = "SELECT 路線,站名 FROM 資料_捷運 With(NoLock) "
                sql &= "where 代號 = '" & table2.Rows(0)("捷運") & "'"
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                If table1.Rows.Count <> 0 Then
                    DropDownList8.SelectedValue = table1.Rows(0)("路線")
                    DropDownList8_SelectedIndexChanged(Nothing, Nothing)
                    DropDownList9.SelectedValue = table2.Rows(0)("捷運")
                End If
            End If

            '新增登記日期欄位:register_date 20110620
            If Not IsDBNull(table2.Rows(0)("register_date")) Then
                Text11.Value = table2.Rows(0)("register_date")
            End If

            'Land_FileNO
            If Not IsDBNull(table2.Rows(0)("Land_FileNO")) Then
                Land_FileNo.Text = table2.Rows(0)("Land_FileNO")
            End If

            '資料來源
            If Not IsDBNull(table2.Rows(0)("資料來源")) Then
                Data_Source.Text = table2.Rows(0)("資料來源")
            End If

            '判斷Movie_h是否有值20110620
            If Not IsDBNull(table2.Rows(0)("movie_h")) Then
                movie_h.Text = table2.Rows(0)("movie_h")
            End If

            If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
                Me.Label464.Visible = True
                Me.RadioButton1.Visible = True
                Me.RadioButton2.Visible = True
                If Not IsDBNull(table2.Rows(0)("提供個資")) Then
                    If table2.Rows(0)("提供個資") = "Y" Then
                        Me.RadioButton1.Checked = True
                    Else
                        Me.RadioButton2.Checked = True
                    End If
                End If
            Else
                Me.Label464.Visible = False
                Me.RadioButton1.Visible = False
                Me.RadioButton2.Visible = False
                Me.RadioButton1.Checked = False
                Me.RadioButton2.Checked = False
            End If
        End If

        conn.Close()
        conn.Dispose()
    End Sub

    Public Sub 載入更新或複製初始頁面_不動產說明書()
        Dim flag As String = "N"
        Dim sid As String = Request("sid")
        Dim oid As String = Request("oid")
        Dim storeid As String = Request("sid")
        Dim src As String = Request.QueryString("src")

        Dim 資料夾 As String = ""
        If src = "OLD" Then
            資料夾 = "expired"
        Else
            資料夾 = "available"
        End If

        '判別網路上有沒有條碼BC
        Dim picurl As String = "https://img.etwarm.com.tw/" & Request("sid") & "/" & 資料夾 & "/" & oid & "BC.PNG"

        Try
            Dim requests As WebRequest = HttpWebRequest.Create(picurl)

            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials

            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then
                    With Image1
                        .ImageUrl = "https://img.etwarm.com.tw/" & Request("sid") & "/" & 資料夾 & "/" & oid & "BC.PNG"
                        .Visible = True
                    End With

                    'OBJ_BARCODE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists("C:\OBJ_BARCODE\" & oid & "BC.jpg") Then
                        '20110412刪除OBJ_BARCODE裡的物件
                        File.Delete("C:\OBJ_BARCODE\" & oid & "BC.jpg")
                    End If
                End If
            End Using
        Catch ex As WebException
            If src <> "OLD" Then
                'Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                'If webResponse.StatusCode = HttpStatusCode.NotFound Then
                '    '找不到，產生條碼
                '    'With Image1
                '    '    .ImageUrl = "https://el.etwarm.com.tw/new_eip/tool/code41.ashx?id=" & oid
                '    '    .Visible = True
                '    'End With
                '    'Dim href As String = "https://el.etwarm.com.tw/new_eip/tool/t_條碼分類.aspx?oid=" & oid & "&sid=" & sid & "&folder=available&rsid=" & Request("sid") & "&src=NOW&check=description"  'check參數用意在不顯示"下一步按鈕"
                '    'Dim NSCRIPT As String = "GB_showCenter('產生條碼', '" & href & "',320,580);"
                '    'Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", NSCRIPT, True)
                'End If
            End If
        End Try

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        sql = "select * "
        sql &= "from 委賣_房地產說明書 With(NoLock) "

        sql &= "where 物件編號 = '" & oid & "' and 店代號 = '" & sid & "'"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count <> 0 Then
            '載入資料

            '報告單位	varchar	20		
            If Not IsDBNull(table1.Rows(0)("報告單位")) Then
                input2.Value = table1.Rows(0)("報告單位").ToString
            End If

            '報告日期	varchar	7		
            If Not IsDBNull(table1.Rows(0)("報告日期")) Then
                Date5.Text = table1.Rows(0)("報告日期").ToString
            End If

            '地政事務所所在區域	varchar	8	
            If Not IsDBNull(table1.Rows(0)("地政事務所所在區域")) Then
                input102.Value = table1.Rows(0)("地政事務所所在區域").ToString
            End If
            '地政事務所	varchar	8	
            If Not IsDBNull(table1.Rows(0)("地政事務所")) Then
                input5.Value = table1.Rows(0)("地政事務所").ToString
            End If
            '謄本核發日期	varchar	7
            If Not IsDBNull(table1.Rows(0)("謄本核發日期")) Then
                Date7.Text = table1.Rows(0)("謄本核發日期").ToString
            End If

            '產權調查	varchar	1	
            If Not IsDBNull(table1.Rows(0)("產權調查")) Then
                If table1.Rows(0)("產權調查") <> "" Then
                    If table1.Rows(0)("產權調查").ToString = 1 Then CheckBoxList1.Items(0).Selected = True
                End If
            End If
            '物件個案調查	varchar	1	
            If Not IsDBNull(table1.Rows(0)("物件個案調查")) Then
                If table1.Rows(0)("物件個案調查") <> "" Then
                    If table1.Rows(0)("物件個案調查").ToString = 1 Then CheckBoxList1.Items(1).Selected = True
                End If
            End If
            '照片說明	varchar	1	
            If Not IsDBNull(table1.Rows(0)("照片說明")) Then
                If table1.Rows(0)("照片說明") <> "" Then
                    If table1.Rows(0)("照片說明").ToString = 1 Then CheckBoxList1.Items(2).Selected = True
                End If
            End If

            '重要交易條件	varchar	1	-----------NEW
            If Not IsDBNull(table1.Rows(0)("重要交易條件")) Then
                If table1.Rows(0)("重要交易條件") <> "" Then
                    If table1.Rows(0)("重要交易條件").ToString = 1 Then CheckBoxList1.Items(3).Selected = True
                End If
            End If
            '成交行情	varchar	1	    -----------NEW
            If Not IsDBNull(table1.Rows(0)("成交行情")) Then
                If table1.Rows(0)("成交行情") <> "" Then
                    If table1.Rows(0)("成交行情").ToString = 1 Then CheckBoxList1.Items(4).Selected = True
                End If
            End If
            '重要設施	varchar	1	  -----------NEW
            If Not IsDBNull(table1.Rows(0)("重要設施")) Then
                If table1.Rows(0)("重要設施") <> "" Then
                    If table1.Rows(0)("重要設施").ToString = 1 Then CheckBoxList1.Items(5).Selected = True
                End If
            End If
            '其他說明	varchar	1
            If Not IsDBNull(table1.Rows(0)("其他說明")) Then
                If table1.Rows(0)("其他說明") <> "" Then
                    If table1.Rows(0)("其他說明").ToString = 1 Then CheckBoxList1.Items(6).Selected = True
                End If
            End If

            '土地權狀影本	varchar	1	
            If Not IsDBNull(table1.Rows(0)("土地權狀影本")) Then
                If table1.Rows(0)("土地權狀影本") <> "" Then
                    If table1.Rows(0)("土地權狀影本").ToString = 1 Then CheckBoxList2.Items(0).Selected = True
                End If
            End If
            '建物權狀影本	varchar	1	
            If Not IsDBNull(table1.Rows(0)("建物權狀影本")) Then
                If table1.Rows(0)("建物權狀影本") <> "" Then
                    If table1.Rows(0)("建物權狀影本").ToString = 1 Then CheckBoxList2.Items(1).Selected = True
                End If
            End If
            '標的現況說明書	varchar	1	
            If Not IsDBNull(table1.Rows(0)("房地產標的現況說明書")) Then
                If table1.Rows(0)("房地產標的現況說明書") <> "" Then
                    If table1.Rows(0)("房地產標的現況說明書").ToString = 1 Then CheckBoxList2.Items(2).Selected = True
                End If
            End If

            '土地謄本	varchar	1	
            If Not IsDBNull(table1.Rows(0)("土地謄本")) Then
                If table1.Rows(0)("土地謄本") <> "" Then
                    If table1.Rows(0)("土地謄本").ToString = 1 Then CheckBoxList2.Items(3).Selected = True
                End If
            End If
            '建物謄本	varchar	1	
            If Not IsDBNull(table1.Rows(0)("建物謄本")) Then
                If table1.Rows(0)("建物謄本") <> "" Then
                    If table1.Rows(0)("建物謄本").ToString = 1 Then CheckBoxList2.Items(4).Selected = True
                End If
            End If
            '預售買賣契約書	varchar	1	
            If Not IsDBNull(table1.Rows(0)("預售買賣契約書")) Then
                If table1.Rows(0)("預售買賣契約書") <> "" Then
                    If table1.Rows(0)("預售買賣契約書").ToString = 1 Then CheckBoxList2.Items(5).Selected = True
                End If
            End If

            '地籍圖	varchar	1		
            If Not IsDBNull(table1.Rows(0)("地籍圖")) Then
                If table1.Rows(0)("地籍圖") <> "" Then
                    If table1.Rows(0)("地籍圖").ToString = 1 Then CheckBoxList2.Items(6).Selected = True
                End If
            End If
            '建物勘測成果圖	varchar	1		
            If Not IsDBNull(table1.Rows(0)("建物勘測成果圖")) Then
                If table1.Rows(0)("建物勘測成果圖") <> "" Then
                    If table1.Rows(0)("建物勘測成果圖").ToString = 1 Then CheckBoxList2.Items(7).Selected = True
                End If
            End If
            '住戶規約	varchar	1		
            If Not IsDBNull(table1.Rows(0)("住戶規約")) Then
                If table1.Rows(0)("住戶規約") <> "" Then
                    If table1.Rows(0)("住戶規約").ToString = 1 Then CheckBoxList2.Items(8).Selected = True
                End If
            End If

            '土地相關位置略圖	varchar	1		
            If Not IsDBNull(table1.Rows(0)("土地相關位置略圖")) Then
                If table1.Rows(0)("土地相關位置略圖") <> "" Then
                    If table1.Rows(0)("土地相關位置略圖").ToString = 1 Then CheckBoxList2.Items(9).Selected = True
                End If
            End If
            '建物相關位置略圖	varchar	1		
            If Not IsDBNull(table1.Rows(0)("建物相關位置略圖")) Then
                If table1.Rows(0)("建物相關位置略圖") <> "" Then
                    If table1.Rows(0)("建物相關位置略圖").ToString = 1 Then CheckBoxList2.Items(10).Selected = True
                End If
            End If
            '停車位位置圖	varchar	1	
            If Not IsDBNull(table1.Rows(0)("停車位位置圖")) Then
                If table1.Rows(0)("停車位位置圖") <> "" Then
                    If table1.Rows(0)("停車位位置圖").ToString = 1 Then CheckBoxList2.Items(11).Selected = True
                End If
            End If

            '土地分管協議	varchar	1	
            If Not IsDBNull(table1.Rows(0)("土地分管協議")) Then
                If table1.Rows(0)("土地分管協議") <> "" Then
                    If table1.Rows(0)("土地分管協議").ToString = 1 Then CheckBoxList2.Items(12).Selected = True
                End If
            End If
            '建物分管協議	varchar	1	
            If Not IsDBNull(table1.Rows(0)("建物分管協議")) Then
                If table1.Rows(0)("建物分管協議") <> "" Then
                    If table1.Rows(0)("建物分管協議").ToString = 1 Then CheckBoxList2.Items(13).Selected = True
                End If
            End If
            '樑柱顯見裂痕照片	varchar	1	
            If Not IsDBNull(table1.Rows(0)("樑柱顯見裂痕照片")) Then
                If table1.Rows(0)("樑柱顯見裂痕照片") <> "" Then
                    If table1.Rows(0)("樑柱顯見裂痕照片").ToString = 1 Then CheckBoxList2.Items(14).Selected = True
                End If
            End If

            '土地分區種類證明	varchar	1	
            If Not IsDBNull(table1.Rows(0)("分區使用證明")) Then
                If table1.Rows(0)("分區使用證明") <> "" Then
                    If table1.Rows(0)("分區使用證明").ToString = 1 Then CheckBoxList2.Items(15).Selected = True
                End If
            End If
            '房屋稅籍相關證明	varchar	1	
            If Not IsDBNull(table1.Rows(0)("房屋稅單")) Then
                If table1.Rows(0)("房屋稅單") <> "" Then
                    If table1.Rows(0)("房屋稅單").ToString = 1 Then CheckBoxList2.Items(16).Selected = True
                End If
            End If
            '其他	varchar	1	
            If Not IsDBNull(table1.Rows(0)("其他")) Then
                If table1.Rows(0)("其他") <> "" Then
                    If table1.Rows(0)("其他").ToString = 1 Then CheckBoxList2.Items(17).Selected = True
                End If
            End If

            '土增稅增值概算表	varchar	1		
            If Not IsDBNull(table1.Rows(0)("增值稅概算")) Then
                If table1.Rows(0)("增值稅概算") <> "" Then
                    If table1.Rows(0)("增值稅概算").ToString = 1 Then CheckBoxList2.Items(18).Selected = True
                End If
            End If

            '使用執照	varchar	1		
            If Not IsDBNull(table1.Rows(0)("使用執照")) Then
                If table1.Rows(0)("使用執照") <> "" Then
                    If table1.Rows(0)("使用執照").ToString = 1 Then CheckBoxList2.Items(19).Selected = True
                End If
            End If

            ''基地面積1	varchar	100		
            'If Not IsDBNull(table1.Rows(0)("基地面積1")) Then
            '    TextBox231.Text = table1.Rows(0)("基地面積1").ToString
            '    If Trim(TextBox231.Text) <> "" Then
            '        自訂坪數_小數點位數(TextBox231, "F")
            '    End If
            'End If
            'If Not IsDBNull(table1.Rows(0)("基地面積坪1")) Then
            '    TextBox230.Text = table1.Rows(0)("基地面積坪1").ToString
            '    '1040522修正-新增自訂小數點位數判斷-坪    
            '    If Trim(TextBox230.Text) <> "" Then
            '        自訂坪數_小數點位數(TextBox230, "T")
            '    End If

            'End If
            ''土地權利範圍1	varchar	100	
            'If Not IsDBNull(table1.Rows(0)("土地權利範圍1")) Then
            '    input30.Value = table1.Rows(0)("土地權利範圍1").ToString
            'End If
            ''法定建蔽率	varchar	10	
            'If Not IsDBNull(table1.Rows(0)("法定建蔽率")) Then
            '    input31.Value = table1.Rows(0)("法定建蔽率").ToString
            'End If
            ''法定容積率	varchar	10
            'If Not IsDBNull(table1.Rows(0)("法定容積率")) Then
            '    input32.Value = table1.Rows(0)("法定容積率").ToString
            'End If
            ''開發限制方式	varchar	20	
            'If Not IsDBNull(table1.Rows(0)("開發限制方式")) Then
            '    DropDownList26.SelectedValue = table1.Rows(0)("開發限制方式").ToString
            'End If
            ''所有權型態為	varchar	8	
            'If Not IsDBNull(table1.Rows(0)("所有權型態為")) Then
            '    DropDownList27.SelectedValue = table1.Rows(0)("所有權型態為").ToString
            'End If
            ''共有土地有無分管協議	varchar	2	
            'If Not IsDBNull(table1.Rows(0)("共有土地有無分管協議")) Then
            '    DropDownList28.SelectedValue = table1.Rows(0)("共有土地有無分管協議").ToString
            'End If
            ''是否受限制處分	varchar	8
            'If Not IsDBNull(table1.Rows(0)("是否受限制處分")) Then
            '    DropDownList29.SelectedValue = table1.Rows(0)("是否受限制處分").ToString
            'End If
            ''有無出租或占用	varchar	2		
            'If Not IsDBNull(table1.Rows(0)("有無出租或占用")) Then
            '    DropDownList30.SelectedValue = table1.Rows(0)("有無出租或占用").ToString
            'End If
            ''地目	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("地目")) Then
            '    DropDownList31.SelectedValue = table1.Rows(0)("地目").ToString
            'End If

            ''建物權利範圍	varchar	20	
            'If Not IsDBNull(table1.Rows(0)("建物權利範圍")) Then
            '    input33.Value = table1.Rows(0)("建物權利範圍").ToString
            'End If
            ''建物若無使用執照一併說明	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("建物若無使用執照一併說明")) Then
            '    input34.Value = table1.Rows(0)("建物若無使用執照一併說明").ToString
            'End If
            ''建物目前管理及使用情況	varchar	20	
            'If Not IsDBNull(table1.Rows(0)("建物目前管理及使用情況")) Then
            '    DropDownList32.SelectedValue = table1.Rows(0)("建物目前管理及使用情況").ToString
            'End If
            ''專有部分之範圍	varchar	20	
            'If Not IsDBNull(table1.Rows(0)("專有部分之範圍")) Then
            '    input35.Value = table1.Rows(0)("專有部分之範圍").ToString
            'End If
            ''共有部分之範圍	varchar	20	
            'If Not IsDBNull(table1.Rows(0)("共有部分之範圍")) Then
            '    input36.Value = table1.Rows(0)("共有部分之範圍").ToString
            'End If
            ''共有部分之範圍1	varchar	60		
            'If Not IsDBNull(table1.Rows(0)("共有部分之範圍1")) Then
            '    input36.Value = table1.Rows(0)("共有部分之範圍1").ToString
            'End If
            ''建物有無共有約定專用部分	varchar	2
            'If Not IsDBNull(table1.Rows(0)("建物有無共有約定專用部分")) Then
            '    DropDownList33.SelectedValue = table1.Rows(0)("建物有無共有約定專用部分").ToString
            'End If
            ''建物有無專有部分約定共用	varchar	2
            'If Not IsDBNull(table1.Rows(0)("建物有無專有部分約定共用")) Then
            '    DropDownList10.SelectedValue = table1.Rows(0)("建物有無專有部分約定共用").ToString
            'End If
            ''建物範圍為	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("建物範圍為")) Then
            '    input37.Value = table1.Rows(0)("建物範圍為").ToString
            'End If
            ''使用方式	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("使用方式")) Then
            '    input38.Value = table1.Rows(0)("使用方式").ToString
            'End If
            ''建物範圍為_共有	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("建物範圍為_共有")) Then
            '    input37_NEW.Value = table1.Rows(0)("建物範圍為_共有").ToString
            'End If
            ''使用方式_共有	varchar	20		
            'If Not IsDBNull(table1.Rows(0)("使用方式_共有")) Then
            '    input38_NEW.Value = table1.Rows(0)("使用方式_共有").ToString
            'End If
            ''管理組織及其管理方式	varchar	10	
            'If Not IsDBNull(table1.Rows(0)("管理組織及其管理方式")) Then
            '    DropDownList14.SelectedValue = table1.Rows(0)("管理組織及其管理方式").ToString
            'End If
            ''有否使用手冊	varchar	2		
            'If Not IsDBNull(table1.Rows(0)("有否使用手冊")) Then
            '    DropDownList15.SelectedValue = table1.Rows(0)("有否使用手冊").ToString
            'End If
            ''有否受限制處分	varchar	8		
            'If Not IsDBNull(table1.Rows(0)("有否受限制處分")) Then
            '    DropDownList34.SelectedValue = table1.Rows(0)("有否受限制處分").ToString
            'End If
            ''有無檢測海砂	varchar	2		
            'If Not IsDBNull(table1.Rows(0)("有無檢測海砂")) Then
            '    DropDownList35.SelectedValue = table1.Rows(0)("有無檢測海砂").ToString
            'End If
            ''有無檢測輻射含量	varchar	2		
            'If Not IsDBNull(table1.Rows(0)("有無檢測輻射含量")) Then
            '    DropDownList18.SelectedValue = table1.Rows(0)("有無檢測輻射含量").ToString
            'End If

            ''有否辦理單獨區分有建物登記	varchar	2
            'If Not IsDBNull(table1.Rows(0)("有否辦理單獨區分有建物登記")) Then
            '    DropDownList7.SelectedValue = table1.Rows(0)("有否辦理單獨區分有建物登記").ToString
            'End If
            ''使用約定方式	varchar	40		
            'If Not IsDBNull(table1.Rows(0)("使用約定方式")) Then
            '    TextBox95.Text = table1.Rows(0)("使用約定方式").ToString
            'End If
            ''進出口為	varchar	8		
            'If Not IsDBNull(table1.Rows(0)("進出口為")) Then
            '    DropDownList23.SelectedValue = table1.Rows(0)("進出口為").ToString
            'End If
            ''車位號碼	varchar	10		
            'If Not IsDBNull(table1.Rows(0)("車位號碼")) Then
            '    input42.Value = table1.Rows(0)("車位號碼").ToString
            'End If
            ''1040417新增位置地上地下	nvarchar	2
            'If IsDBNull(table1.Rows(0)("位置地上地下")) = False Then
            '    DropDownList64.SelectedValue = table1.Rows(0)("位置地上地下")
            'Else '若為NULL值
            '    DropDownList64.SelectedValue = "地下"
            'End If

            ''位置地下	varchar	3	
            'If Not IsDBNull(table1.Rows(0)("位置地下")) Then
            '    input43.Value = table1.Rows(0)("位置地下").ToString
            'End If

            ''車位性質	nvarchar	20	
            'If Not IsDBNull(table1.Rows(0)("車位性質")) Then
            '    DropDownList67.SelectedValue = table1.Rows(0)("車位性質").ToString
            'End If

            '將舊有他項權利新增至他項權利細項表格-舊有的將不使用---新系統上線一段時間後，此段將可廢除，所有舊資料應已移轉完畢
            'Dim conn_他項權利 = New SqlConnection(EGOUPLOADSqlConnStr)

            'conn_他項權利.Open()
            'i = 1
            'For i = 1 To 8
            '    If Not IsDBNull(table1.Rows(0)("權利種類" & i).ToString) Then
            '        If table1.Rows(0)("權利種類" & i).ToString <> "" And GridView2.Rows.Count = 0 Then


            '            Dim count As Integer = 0
            '            Dim sql2 As String = "select top 1 Num as Num from 物件他項權利細項 With(NoLock) where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "' order by Num desc"

            '            adpt = New SqlDataAdapter(sql2, conn_他項權利)
            '            ds = New DataSet()
            '            adpt.Fill(ds, "table2")
            '            table2 = ds.Tables("table2")
            '            If table2.Rows.Count > 0 Then
            '                count = table2.Rows(0)("Num") + 1
            '            Else
            '                count = 0
            '            End If

            '            Dim 權利類別 As String = ""
            '            If i <= 4 Then
            '                權利類別 = "建物"
            '            Else
            '                權利類別 = "土地"
            '            End If

            '            Dim sql As String = "insert 物件他項權利細項(物件編號,店代號, Num, 權利類別, 權利種類,順位,登記日期,設定,設定權利人) values ('" & Me.Label11.Text & "','" & Me.Label12.Text & "','" & count & "','" & 權利類別 & "','" & table1.Rows(0)("權利種類" & i).ToString & "','" & table1.Rows(0)("順位" & i).ToString & "','" & table1.Rows(0)("登記日期" & i).ToString & "','" & table1.Rows(0)("設定" & i).ToString & "','" & table1.Rows(0)("設定權利人" & i).ToString & "')"
            '            Dim cmd As New SqlCommand(sql, conn_他項權利)
            '            cmd.CommandType = CommandType.Text
            '            cmd.ExecuteNonQuery()

            '            flag = "Y"
            '        End If
            '    End If

            'Next
            'conn_他項權利.Close()
            'conn_他項權利.Dispose()

            ''權利種類1	varchar	10		
            'DropDownList22.SelectedValue = table1.Rows(0)("權利種類1").ToString
            ''順位1	varchar	1		
            'input44.Value = table1.Rows(0)("順位1").ToString
            ''登記日期1	varchar	7		
            'input45.Value = table1.Rows(0)("登記日期1").ToString
            ''設定1	varchar	6		
            'input46.Value = table1.Rows(0)("設定1").ToString
            ''設定權利人1	varchar	20		
            'input47.Value = table1.Rows(0)("設定權利人1").ToString
            ''權利種類2	varchar	10		
            'DropDownList23.SelectedValue = table1.Rows(0)("權利種類2").ToString
            ''順位2	varchar	1		
            'input48.Value = table1.Rows(0)("順位2").ToString
            ''登記日期2	varchar	7		
            'input49.Value = table1.Rows(0)("登記日期2").ToString
            ''設定2	varchar	6		
            'input50.Value = table1.Rows(0)("設定2").ToString
            ''設定權利人2	varchar	20		
            'input51.Value = table1.Rows(0)("設定權利人2").ToString
            ''權利種類3	varchar	10		
            'DropDownList24.SelectedValue = table1.Rows(0)("權利種類3").ToString
            ''順位3	varchar	1		
            'input52.Value = table1.Rows(0)("順位3").ToString
            ''登記日期3	varchar	7		
            'input53.Value = table1.Rows(0)("登記日期3").ToString
            ''設定3	varchar	6		
            'input54.Value = table1.Rows(0)("設定3").ToString
            ''設定權利人3	varchar	20		
            'input55.Value = table1.Rows(0)("設定權利人3").ToString
            '與土地他項權利部相同	varchar	1		
            If table1.Rows(0)("與土地他項權利部相同").ToString = "1" Then CheckBox24.Checked = True
            '其他如下	varchar	1		
            If table1.Rows(0)("其他如下").ToString = "1" Then CheckBox25.Checked = True
            ''權利種類4	varchar	10		
            'DropDownList25.SelectedValue = table1.Rows(0)("權利種類4").ToString
            ''順位4	varchar	1		
            'input58.Value = table1.Rows(0)("順位4").ToString
            ''登記日期4	varchar	7		
            'input59.Value = table1.Rows(0)("登記日期4").ToString
            ''設定4	varchar	6		
            'input60.Value = table1.Rows(0)("設定4").ToString
            ''設定權利人4	varchar	20		
            'input61.Value = table1.Rows(0)("設定權利人4").ToString
            ''權利種類5	varchar	10		
            'DropDownList26.SelectedValue = table1.Rows(0)("權利種類5").ToString
            ''順位5	varchar	1		
            'input62.Value = table1.Rows(0)("順位5").ToString
            ''登記日期5	varchar	7		
            'input63.Value = table1.Rows(0)("登記日期5").ToString
            ''設定5	varchar	6		
            'input64.Value = table1.Rows(0)("設定5").ToString
            ''設定權利人5	varchar	20		
            'input65.Value = table1.Rows(0)("設定權利人5").ToString
            ''權利種類6	varchar	10		
            'DropDownList27.SelectedValue = table1.Rows(0)("權利種類6").ToString
            ''順位6	varchar	1		
            'input66.Value = table1.Rows(0)("順位6").ToString
            ''登記日期6	varchar	7		
            'input67.Value = table1.Rows(0)("登記日期6").ToString
            ''設定6	varchar	6		
            'input68.Value = table1.Rows(0)("設定6").ToString
            ''設定權利人6	varchar	20		
            'input69.Value = table1.Rows(0)("設定權利人6").ToString
            ''權利種類7	varchar	10		
            'DropDownList12.SelectedValue = table1.Rows(0)("權利種類7").ToString
            ''順位7	varchar	1		
            'input111.Value = table1.Rows(0)("順位7").ToString
            ''登記日期7	varchar	7		
            'input112.Value = table1.Rows(0)("登記日期7").ToString
            ''設定7	varchar	6		
            'input113.Value = table1.Rows(0)("設定7").ToString
            ''設定權利人7	varchar	20		
            'input114.Value = table1.Rows(0)("設定權利人7").ToString
            ''權利種類8	varchar	10		
            'DropDownList11.SelectedValue = table1.Rows(0)("權利種類8").ToString
            ''順位8	varchar	1		
            'input115.Value = table1.Rows(0)("順位8").ToString
            ''登記日期8	varchar	7		
            'input116.Value = table1.Rows(0)("登記日期8").ToString
            ''設定8	varchar	6		
            'input117.Value = table1.Rows(0)("設定8").ToString
            ''設定權利人8	varchar	20		
            'input118.Value = table1.Rows(0)("設定權利人8").ToString

            ''固定物	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("固定物")) Then
            '    If table1.Rows(0)("固定物").ToString = 1 Then CheckBox4.Checked = True
            'End If

            ''20151127 remove by nick 電話欄位改為存流理台
            'If Not IsDBNull(table1.Rows(0)("電話")) Then
            '    If table1.Rows(0)("電話").ToString = 1 Then
            '        CheckBox29.Checked = True
            '    End If

            'End If
            '電話	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("電話")) Then
            '    If table1.Rows(0)("電話").ToString = 1 Then CheckBox5.Checked = True
            'End If
            '電話線	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("電話線")) Then
            '    TextBox241.Text = table1.Rows(0)("電話線").ToString
            'End If
            ''梳妝台	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("梳妝台")) Then
            '    If table1.Rows(0)("梳妝台").ToString = 1 Then CheckBox6.Checked = True
            'End If
            ''燈飾	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("燈飾")) Then
            '    If table1.Rows(0)("燈飾").ToString = 1 Then CheckBox7.Checked = True
            'End If
            ''冷氣	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("冷氣")) Then
            '    If table1.Rows(0)("冷氣").ToString = 1 Then CheckBox8.Checked = True
            'End If
            ''冷氣台	varchar	1
            'If Not IsDBNull(table1.Rows(0)("冷氣台")) Then
            '    TextBox242.Text = table1.Rows(0)("冷氣台").ToString
            'End If
            ''窗簾	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("窗簾")) Then
            '    If table1.Rows(0)("窗簾").ToString = 1 Then CheckBox9.Checked = True
            'End If
            ''床組	varchar	1
            'If Not IsDBNull(table1.Rows(0)("床組")) Then
            '    If table1.Rows(0)("床組").ToString = 1 Then CheckBox10.Checked = True
            'End If
            ''冰箱	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("冰箱")) Then
            '    If table1.Rows(0)("冰箱").ToString = 1 Then CheckBox11.Checked = True
            'End If
            ''冰箱台	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("冰箱台")) Then
            '    TextBox38.Text = table1.Rows(0)("冰箱台").ToString
            'End If
            ''熱水器	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("熱水器")) Then
            '    If table1.Rows(0)("熱水器").ToString = 1 Then CheckBox12.Checked = True
            'End If
            ''沙發組	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("沙發組")) Then
            '    If table1.Rows(0)("沙發組").ToString = 1 Then CheckBox13.Checked = True

            '    '將系統櫥櫃組的欄位用來當沙發組數用 20151127 by nick
            '    If Not IsDBNull(table1.Rows(0)("系統櫥櫃組")) Then
            '        If Trim(table1.Rows(0)("系統櫥櫃組")) <> "" Then
            '            Safa_count.Text = Trim(table1.Rows(0)("系統櫥櫃組"))
            '        End If
            '    End If
            'End If

            ''瓦斯廚具	varchar	1
            'If Not IsDBNull(table1.Rows(0)("瓦斯廚具")) Then
            '    If table1.Rows(0)("瓦斯廚具").ToString = 1 Then CheckBox14.Checked = True
            'End If
            ''20151127 remove by nick
            ''瓦斯廚具樣式	varchar	8	
            ''If Not IsDBNull(table1.Rows(0)("瓦斯廚具樣式")) Then
            ''    TextBox39.Text = table1.Rows(0)("瓦斯廚具樣式").ToString
            ''End If
            ''壁櫥	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("壁櫥")) Then
            '    If table1.Rows(0)("壁櫥").ToString = 1 Then CheckBox15.Checked = True
            'End If
            ''酒櫃	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("酒櫃")) Then
            '    If table1.Rows(0)("酒櫃").ToString = 1 Then CheckBox16.Checked = True
            'End If
            ''自來瓦斯	varchar	1		
            'If Not IsDBNull(table1.Rows(0)("自來瓦斯")) Then
            '    If table1.Rows(0)("自來瓦斯").ToString = 1 Then CheckBox17.Checked = True
            'End If
            ''20151127 remove by nick
            ''飲水機
            ''If Not IsDBNull(table1.Rows(0)("飲水機")) Then
            ''    If table1.Rows(0)("飲水機").ToString = "1" Then
            ''        CheckBox18.Checked = True
            ''    End If
            ''End If
            ''洗衣機	varchar	1	
            'If Not IsDBNull(table1.Rows(0)("洗衣機")) Then
            '    If table1.Rows(0)("洗衣機").ToString = "1" Then
            '        CheckBox19.Checked = True

            '        If Not IsDBNull(table1.Rows(0)("洗衣機台")) Then
            '            If Trim(table1.Rows(0)("洗衣機台")) <> "" Then
            '                TextBox40.Text = Trim(table1.Rows(0)("洗衣機台"))
            '            End If
            '        End If
            '    End If
            'End If
            ''乾衣機	varchar	1
            'If Not IsDBNull(table1.Rows(0)("乾衣機")) Then
            '    If table1.Rows(0)("乾衣機").ToString = "1" Then
            '        CheckBox20.Checked = True
            '        If Not IsDBNull(table1.Rows(0)("乾衣機台")) Then
            '            If Trim(table1.Rows(0)("乾衣機台")) <> "" Then
            '                TextBox41.Text = Trim(table1.Rows(0)("乾衣機台"))
            '            End If
            '        End If
            '    End If
            'End If
            ''系統櫥櫃 varchar 1
            'If Not IsDBNull(table1.Rows(0)("系統櫥櫃")) Then
            '    If table1.Rows(0)("系統櫥櫃").ToString = "1" Then
            '        CheckBox21.Checked = True
            '    End If
            'End If
            'If Not IsDBNull(table1.Rows(0)("系統櫥櫃組")) Then
            '    If Trim(table1.Rows(0)("系統櫥櫃組")) <> "" Then
            '        TextBox42.Text = Trim(table1.Rows(0)("系統櫥櫃組"))
            '    End If
            'End If
            ''天然瓦斯度數表	varchar	1
            'If Not IsDBNull(table1.Rows(0)("天然瓦斯度數表")) Then
            '    If table1.Rows(0)("天然瓦斯度數表").ToString = "1" Then
            '        CheckBox22.Checked = True
            '    End If
            'End If
            ''其他項目	varchar	1		
            'If table1.Rows(0)("其他項目").ToString = 1 Then CheckBox23.Checked = True
            ''附贈設備其他                
            'If Not IsDBNull(table1.Rows(0)("其他項目內容")) Then
            '    If Trim(table1.Rows(0)("其他項目內容")) <> "" Then
            '        TextBox43.Text = Trim(table1.Rows(0)("其他項目內容"))
            '    End If
            'End If

            '物件標的	varchar	6	
            If Not IsDBNull(table1.Rows(0)("物件標的")) Then
                DropDownList38.SelectedValue = table1.Rows(0)("物件標的").ToString
            End If
            '現況	varchar	4	
            If Not IsDBNull(table1.Rows(0)("現況")) Then
                DropDownList39.SelectedValue = table1.Rows(0)("現況").ToString
            End If
            '交屋情況	varchar	4	
            If Not IsDBNull(table1.Rows(0)("交屋情況")) Then
                DropDownList40.SelectedValue = table1.Rows(0)("交屋情況").ToString
            End If
            '商談交屋情況	varchar	10		
            If Not IsDBNull(table1.Rows(0)("商談交屋情況")) Then
                input89.Value = table1.Rows(0)("商談交屋情況").ToString
            End If
            '中庭花園	varchar	2	
            If Not IsDBNull(table1.Rows(0)("中庭花園")) Then
                DropDownList41.SelectedValue = table1.Rows(0)("中庭花園").ToString
            End If
            '其他中庭花園	varchar	12		
            If Not IsDBNull(table1.Rows(0)("其他中庭花園")) Then
                input90.Value = table1.Rows(0)("其他中庭花園").ToString
            End If
            '警衛管理	varchar	2		
            If Not IsDBNull(table1.Rows(0)("警衛管理")) Then
                DropDownList42.SelectedValue = table1.Rows(0)("警衛管理").ToString
            End If
            '其他警衛管理	varchar	12	
            If Not IsDBNull(table1.Rows(0)("其他警衛管理")) Then
                input91.Value = table1.Rows(0)("其他警衛管理").ToString
            End If
            '外牆外飾	varchar	8		
            If Not IsDBNull(table1.Rows(0)("外牆外飾")) Then
                DropDownList43.SelectedValue = table1.Rows(0)("外牆外飾").ToString
            End If
            '其他外牆外飾   varchar(10)  
            If Not IsDBNull(table1.Rows(0)("其他外牆外飾")) Then
                input92.Value = table1.Rows(0)("其他外牆外飾").ToString
            End If
            '地板	varchar	6		
            If Not IsDBNull(table1.Rows(0)("地板")) Then
                DropDownList44.SelectedValue = table1.Rows(0)("地板").ToString
            End If
            '其他地板	varchar	10		
            If Not IsDBNull(table1.Rows(0)("其他地板")) Then
                input93.Value = table1.Rows(0)("其他地板").ToString
            End If
            '自來水	varchar	6		
            If Not IsDBNull(table1.Rows(0)("自來水")) Then
                DropDownList45.SelectedValue = table1.Rows(0)("自來水").ToString
            End If
            '未安裝自來水原因	varchar	20
            If Not IsDBNull(table1.Rows(0)("未安裝自來水原因")) Then
                input94.Value = table1.Rows(0)("未安裝自來水原因").ToString
            End If
            '電力系統	varchar	6
            If Not IsDBNull(table1.Rows(0)("電力系統")) Then
                DropDownList46.SelectedValue = table1.Rows(0)("電力系統").ToString
            End If
            '有無獨立電錶	varchar	2	
            If Not IsDBNull(table1.Rows(0)("有無獨立電錶")) Then
                DropDownList47.SelectedValue = table1.Rows(0)("有無獨立電錶").ToString
            End If
            '室內建材	varchar	20		
            If Not IsDBNull(table1.Rows(0)("室內建材")) Then
                DropDownList48.SelectedValue = table1.Rows(0)("室內建材").ToString
            End If
            '隔間材料	varchar	4		
            If Not IsDBNull(table1.Rows(0)("隔間材料")) Then
                '1040505新增判斷，自行新增的選項不在選單裡
                Dim flags As String = "False"
                For i = 0 To DropDownList49.Items.Count - 1
                    If DropDownList49.Items(i).Value = table1.Rows(0)("隔間材料").ToString Then
                        DropDownList49.SelectedIndex = i
                        flags = "True"
                        Exit For
                    End If
                Next

                If flags = "False" Then
                    DropDownList49.Items.Add(New ListItem(table1.Rows(0)("隔間材料").ToString, table1.Rows(0)("隔間材料").ToString))
                    DropDownList49.SelectedValue = table1.Rows(0)("隔間材料").ToString
                End If
                'DropDownList49.SelectedValue = table1.Rows(0)("隔間材料").ToString
            End If
            '電話系統	varchar	6		
            If Not IsDBNull(table1.Rows(0)("電話系統")) Then
                DropDownList50.SelectedValue = table1.Rows(0)("電話系統").ToString
            End If
            '未安裝電話系統原因	varchar	20		
            If Not IsDBNull(table1.Rows(0)("未安裝電話系統原因")) Then
                input95.Value = table1.Rows(0)("未安裝電話系統原因").ToString
            End If
            '瓦斯系統	varchar	6		
            If Not IsDBNull(table1.Rows(0)("瓦斯系統")) Then
                DropDownList51.SelectedValue = table1.Rows(0)("瓦斯系統").ToString
            End If
            '未安裝瓦斯系統	varchar	20		
            If Not IsDBNull(table1.Rows(0)("未安裝瓦斯系統")) Then
                input96.Value = table1.Rows(0)("未安裝瓦斯系統").ToString
            End If
            '主要建材	varchar	20		
            If Not IsDBNull(table1.Rows(0)("建築結構")) Then
                input97.Value = table1.Rows(0)("建築結構").ToString
            End If

            '簽約金	varchar	4	%
            If Not IsDBNull(table1.Rows(0)("簽約金")) Then
                TextBox258.Text = table1.Rows(0)("簽約金").ToString
            End If
            '第一期金額	Decimal	
            If Not IsDBNull(table1.Rows(0)("第一期金額")) Then
                TextBox262.Text = table1.Rows(0)("第一期金額").ToString
            End If
            '備証款	varchar	4	%	
            If Not IsDBNull(table1.Rows(0)("備証款")) Then
                TextBox259.Text = table1.Rows(0)("備証款").ToString
            End If
            '第二期金額	Decimal	
            If Not IsDBNull(table1.Rows(0)("第二期金額")) Then
                TextBox263.Text = table1.Rows(0)("第二期金額").ToString
            End If
            '完稅款	varchar	4	%	
            If Not IsDBNull(table1.Rows(0)("完稅款")) Then
                TextBox260.Text = table1.Rows(0)("完稅款").ToString
            End If
            '第三期金額	Decimal	
            If Not IsDBNull(table1.Rows(0)("第三期金額")) Then
                TextBox264.Text = table1.Rows(0)("第三期金額").ToString
            End If
            '尾款	varchar	4	%
            If Not IsDBNull(table1.Rows(0)("尾款")) Then
                TextBox261.Text = table1.Rows(0)("尾款").ToString
            End If
            '第四期金額	Decimal	
            If Not IsDBNull(table1.Rows(0)("第四期金額")) Then
                TextBox265.Text = table1.Rows(0)("第四期金額").ToString
            End If

            '自用土地增值稅約	nvarchar	20		
            'If Not IsDBNull(table1.Rows(0)("自用土地增值稅約")) Then
            '    input103.Value = table1.Rows(0)("自用土地增值稅約").ToString
            'End If
            '一般增值稅約	nvarchar	20		
            'If Not IsDBNull(table1.Rows(0)("一般增值稅約")) Then
            '    input104.Value = table1.Rows(0)("一般增值稅約").ToString
            'End If
            '契稅約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("契稅約")) Then
            '    input105.Value = table1.Rows(0)("契稅約").ToString
            'End If
            '----------------------------------------新版新增項目------------------------------------------
            '地價稅約	nvarchar	20		
            'If Not IsDBNull(table1.Rows(0)("地價稅約")) Then
            '    input106.Value = table1.Rows(0)("地價稅約").ToString
            'End If
            '房屋稅約	nvarchar	20		
            'If Not IsDBNull(table1.Rows(0)("房屋稅約")) Then
            '    input107.Value = table1.Rows(0)("房屋稅約").ToString
            'End If
            '工程受益費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("工程受益費約")) Then
            '    input108.Value = table1.Rows(0)("工程受益費約").ToString
            'End If
            '代書費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("代書費約")) Then
            '    input109.Value = table1.Rows(0)("代書費約").ToString
            'End If
            '登記規費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("登記規費約")) Then
            '    input110.Value = table1.Rows(0)("登記規費約").ToString
            'End If
            '公證費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("公證費約")) Then
            '    input111.Value = table1.Rows(0)("公證費約").ToString
            'End If
            '印花稅約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("印花稅約")) Then
            '    input112.Value = table1.Rows(0)("印花稅約").ToString
            'End If
            '水電費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("水電費約")) Then
            '    input113.Value = table1.Rows(0)("水電費約").ToString
            'End If
            '管理費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("管理費約")) Then
            '    input114.Value = table1.Rows(0)("管理費約").ToString
            'End If
            '電話費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("電話費約")) Then
            '    input115.Value = table1.Rows(0)("電話費約").ToString
            'End If
            '瓦斯費約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("瓦斯費約")) Then
            '    input116.Value = table1.Rows(0)("瓦斯費約").ToString
            'End If
            '奢侈稅約	nvarchar	20			
            'If Not IsDBNull(table1.Rows(0)("奢侈稅約")) Then
            '    input117.Value = table1.Rows(0)("奢侈稅約").ToString
            'End If
            '----------------------------------------新版新增項目------------------------------------------

            '增值稅	varchar	12		
            If Not IsDBNull(table1.Rows(0)("增值稅")) Then
                DropDownList52.SelectedValue = table1.Rows(0)("增值稅").ToString
            End If
            '契稅	varchar	12	
            If Not IsDBNull(table1.Rows(0)("契稅")) Then
                DropDownList53.SelectedValue = table1.Rows(0)("契稅").ToString
            End If
            '地價稅	varchar	12		
            If Not IsDBNull(table1.Rows(0)("地價稅")) Then
                DropDownList54.SelectedValue = table1.Rows(0)("地價稅").ToString
            End If
            '房屋稅	varchar	12	
            If Not IsDBNull(table1.Rows(0)("房屋稅")) Then
                DropDownList55.SelectedValue = table1.Rows(0)("房屋稅").ToString
            End If
            '工程受益費	varchar	12	
            If Not IsDBNull(table1.Rows(0)("工程受益費")) Then
                DropDownList56.SelectedValue = table1.Rows(0)("工程受益費").ToString
            End If
            '代書費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("代書費")) Then
                DropDownList57.SelectedValue = table1.Rows(0)("代書費").ToString
            End If
            '登記規費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("登記規費")) Then
                DropDownList58.SelectedValue = table1.Rows(0)("登記規費").ToString
            End If
            '公證費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("公證費")) Then
                DropDownList59.SelectedValue = table1.Rows(0)("公證費").ToString
            End If
            '印花稅	varchar	12		
            If Not IsDBNull(table1.Rows(0)("印花稅")) Then
                DropDownList60.SelectedValue = table1.Rows(0)("印花稅").ToString
            End If
            '水電費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("水電費")) Then
                DropDownList61.SelectedValue = table1.Rows(0)("水電費").ToString
            End If
            '管理費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("管理費")) Then
                DropDownList62.SelectedValue = table1.Rows(0)("管理費").ToString
            End If
            '電話費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("電話費")) Then
                DropDownList63.SelectedValue = table1.Rows(0)("電話費").ToString
            End If
            '瓦斯費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("瓦斯費")) Then
                DropDownList13.SelectedValue = table1.Rows(0)("瓦斯費").ToString
            End If
            '奢侈稅	varchar	12		
            If Not IsDBNull(table1.Rows(0)("奢侈稅")) Then
                'DropDownList68.SelectedValue = table1.Rows(0)("奢侈稅").ToString
            End If
            '實價登錄費	varchar	12		
            If Not IsDBNull(table1.Rows(0)("實價登錄費")) Then
                DropDownList70.SelectedValue = table1.Rows(0)("實價登錄費").ToString
            End If
            '代書費New	varchar	12		
            If Not IsDBNull(table1.Rows(0)("代書費New")) Then
                DropDownList71.SelectedValue = table1.Rows(0)("代書費New").ToString
            End If
            '建築主要用途	varchar	20	     
            If Not IsDBNull(table1.Rows(0)("建築用途")) Then
                '1040505新增判斷，自行新增的選項不在選單裡
                Dim flags As String = "False"
                For i = 0 To DropDownList19.Items.Count - 1
                    If DropDownList19.Items(i).Value = table1.Rows(0)("建築用途").ToString Then
                        DropDownList19.SelectedIndex = i
                        flags = "True"
                        Exit For
                    End If
                Next
                If flags = "False" Then
                    DropDownList19.Items.Add(New ListItem(table1.Rows(0)("建築用途").ToString, table1.Rows(0)("建築用途").ToString))
                    DropDownList19.SelectedValue = table1.Rows(0)("建築用途").ToString
                End If
                'DropDownList19.SelectedValue = table1.Rows(0)("建築用途").ToString
            End If

            '20140224新增-Step1.先讀入使用分區
            '讀入物件資料(Request("oid"), "使用分區")
            '20140224新增-Step2.再寫入覆蓋不動產說明書原使用管制內容

            '讀物件物件用途-故駐記
            'sql = "update 委賣_房地產說明書   set 使用管制內容='" & Label47.Text & "' WHERE 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "'  "
            'cmd = New SqlCommand(sql, conn)
            'cmd.ExecuteNonQuery()
        End If

        conn.Close()
        conn = Nothing

        '如果有跑進就他項權利移轉至新資料表時,flag=TRUE,要重新載入他項權利資料表
        If flag = "Y" Then
            Load_他項權利Data("OLD")
        End If
    End Sub

    '面積細項區塊(START)-------------------------------------------------------------------------------------------------------------------
    Sub 類別名稱()

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        sql = "Select DISTINCT(類別) From 資料_委賣物件資料表自訂面積細項 With(NoLock) where (店代號='A0001') and 類別 in ('土地面積','主建物','增建')"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        DropDownList2.Items.Clear()
        DropDownList2.Items.Add("選擇類別名稱")

        For i = 0 To table1.Rows.Count - 1
            DropDownList2.Items.Add(table1.Rows(i)("類別"))
        Next

        conn.Close()
        conn.Dispose()
    End Sub

    Sub 項目名稱(ByVal itemname As String, Optional ByVal defaultselect As String = "")

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        'sql = "Select 項目名稱 From 資料_委賣物件資料表自訂面積細項 With(NoLock) Where 類別 in ('" & itemname & "') and (店代號='" & Request.Cookies("store_id").Value & "' or 店代號='A0001')"
        sql = "Select * FROM (Select DISTINCT(項目名稱),indexSortUse From 資料_委賣物件資料表自訂面積細項 With(NoLock) Where 類別 in ('" & itemname & "') and (店代號 in (" & myobj.mstoresnew(Request.Cookies("webfly_empno").Value) & ",'A0001'))) AS Tab ORDER BY indexSortUse, 項目名稱"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")


        DropDownList1.Items.Clear()
        DropDownList1.Items.Add("選擇項目名稱")

        For i = 0 To table1.Rows.Count - 1
            DropDownList1.Items.Add(table1.Rows(i)("項目名稱"))
        Next

        'DropDownList1.Items.Add("車位")
        'If itemname = "附屬物" Then
        '    DropDownList1.Items.Add("車位(僅使用權)")
        'End If
        'If itemname = "共有部分" Then
        '    DropDownList1.Items.Add("無")
        'End If
        DropDownList1.Items.Add("其他")

        conn.Close()
        conn.Dispose()

        If defaultselect.Length > 0 Then
            For Each li As ListItem In DropDownList1.Items
                If li.Text = defaultselect Then
                    li.Selected = True
                End If
            Next
        End If

    End Sub

    Protected Sub DropDownList2_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged

        'CheckBox28.Checked = False
        CheckBox98.Visible = False
        CheckBox99.Visible = False

        If Me.DropDownList2.SelectedItem.Text = "選擇類別名稱" Then
            clear()
            Label69.Text = "建號"
        Else
            項目名稱(Me.DropDownList2.SelectedItem.Text)

            If Me.DropDownList2.SelectedItem.Text = "土地面積" Then
                Label69.Text = "地號"
                Label70.Text = "土地使用分區"

                Label72.Text = "法定建蔽率"
                Label73.Text = "法定容積率"

                DropDownList69.Visible = True
                DropDownList65.Visible = True

                '法定建蔽率
                TextBox256.Visible = True
                Label74.Visible = True
                '法定容積率
                TextBox257.Visible = True
                Label75.Visible = True

                TextBox254.Text = ""
                TextBox254.Visible = False
                Image15.Visible = False
                Date8.Text = ""
                Date8.Visible = False
                seeday4.Visible = False

                '移除權利人
                'Label71.Text = "權利人"
                'TextBox255.Visible = True

            ElseIf Me.DropDownList2.SelectedItem.Text = "增建" Then
                Label69.Text = "座落門牌或位置"
                Label70.Text = "增建用途"
                Label71.Text = "增建日期"
                Label72.Text = ""
                Label73.Text = ""

                DropDownList69.SelectedValue = "請選擇"
                DropDownList69.Visible = False
                DropDownList65.SelectedValue = "請選擇"
                DropDownList65_SelectedIndexChanged(Nothing, Nothing)
                DropDownList65.Visible = False
                DropDownList66.Visible = False
                TextBox255.Text = ""
                TextBox255.Visible = False
                '法定建蔽率
                TextBox256.Text = ""
                TextBox256.Visible = False
                Label74.Visible = False
                '法定容積率
                TextBox257.Text = ""
                TextBox257.Visible = False
                Label75.Visible = False

                TextBox254.Visible = True
                Image15.Visible = True
                Date8.Visible = True
                seeday4.Visible = True
            Else
                Label69.Text = "建號"
                Label70.Text = ""
                Label71.Text = ""
                Label72.Text = ""
                Label73.Text = ""

                DropDownList69.SelectedValue = "請選擇"
                DropDownList69.Visible = False
                DropDownList65.SelectedValue = "請選擇"
                DropDownList65_SelectedIndexChanged(Nothing, Nothing)
                DropDownList65.Visible = False
                DropDownList66.Visible = False
                TextBox255.Text = ""
                TextBox255.Visible = False
                '法定建蔽率
                TextBox256.Text = ""
                TextBox256.Visible = False
                Label74.Visible = False
                '法定容積率
                TextBox257.Text = ""
                TextBox257.Visible = False
                Label75.Visible = False

                TextBox254.Text = ""
                TextBox254.Visible = False
                Image15.Visible = False
                Date8.Text = ""
                Date8.Visible = False
                seeday4.Visible = False
            End If

            If Me.DropDownList2.SelectedItem.Text = "主建物" Then
                '20151106 顯示第二層
                DDL_level2.Visible = True

                '本建號為公設、車位            
                CheckBox98.Visible = True
                CheckBox99.Visible = True
            Else
                DDL_level2.Visible = False
                DDL_level2.SelectedValue = "主建物','附屬物"
                '本建號為公設、車位            
                CheckBox98.Visible = False
                CheckBox99.Visible = False
                CheckBox98.Checked = False
                CheckBox99.Checked = False

                TextBox1.Visible = False
                TextBox3.Visible = False
                Label26.Visible = False

                TextBox44.Visible = False
                TextBox45.Visible = False
                Label46.Visible = False
            End If
        End If

        Me.TextBox11.Text = ""
        Me.TextBox11.Visible = False
    End Sub

    Protected Sub DDL_level2_SelectedIndexChanged(sender As Object, e As EventArgs)

        '顯示權利範圍
        If DDL_level2.SelectedValue = "共有部分" Then
            TextBox1.Visible = True
            TextBox3.Visible = True
            Label26.Visible = True
            Label70.Text = "標示部權利範圍"
            Label72.Text = "共有部分車位權利範圍"

            TextBox44.Visible = True
            TextBox45.Visible = True
            Label46.Visible = True
        Else
            TextBox1.Visible = False
            TextBox3.Visible = False
            Label26.Visible = False
            Label70.Text = ""
            Label72.Text = ""

            TextBox44.Visible = False
            TextBox45.Visible = False
            Label46.Visible = False
        End If

        項目名稱(Me.DDL_level2.SelectedValue)
    End Sub
    Sub chk_項目()
        Dim 類別代號 As String = ""
        Dim 名稱 As String = ""

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        Dim sql As String = ""

        Select Case Trim(Me.DropDownList2.SelectedItem.Text)
            Case "土地面積"
                類別代號 = "E"
            Case "公共設施"
                類別代號 = "C"
            Case "主建物"
                類別代號 = "A"
            Case "地下室", "地下層"
                類別代號 = "D"
            Case "車位面積(公設內)"
                類別代號 = "H"
            Case "車位面積(產權獨立)"
                類別代號 = "I"
            Case "附屬物"
                類別代號 = "B"
            Case "庭院坪數"
                類別代號 = "F"
            Case "增建"
                類別代號 = "G"
        End Select
        If DDL_level2.Visible = True Then
            名稱 = ""
            類別代號 = ""
            Select Case Trim(Me.DDL_level2.SelectedItem.Text)
                Case "層次"
                    名稱 = "主建物"
                    類別代號 = "A"
                Case "附屬建物"
                    名稱 = "附屬物"
                    類別代號 = "B"
                Case "共有部分"
                    名稱 = "共有部分"
                    類別代號 = "J"
            End Select
            sql = "Select * From 資料_委賣物件資料表自訂面積細項 With(NoLock) where 店代號 in ('" & Request.Cookies("store_id").Value & "','A0001') and 類別='" & 名稱 & "' and 項目名稱='" & Trim(Me.TextBox11.Text) & "'"
        Else
            sql = "Select * From 資料_委賣物件資料表自訂面積細項 With(NoLock) where 店代號 in ('" & Request.Cookies("store_id").Value & "','A0001') and 類別='" & Trim(Me.DropDownList2.SelectedItem.Text) & "' and 項目名稱='" & Trim(Me.TextBox11.Text) & "'"
        End If

        'Dim sql As String = "Select * From 資料_委賣物件資料表自訂面積細項 With(NoLock) where 店代號 in ('" & Request.Cookies("store_id").Value & "','A0001') and 類別='" & Trim(Me.DropDownList2.SelectedItem.Text) & "' and 項目名稱='" & Trim(Me.TextBox11.Text) & "'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count = 0 Then
            Select Case Trim(Me.DropDownList2.SelectedItem.Text)
                Case "土地面積"
                    類別代號 = "E"
                Case "公共設施"
                    類別代號 = "C"
                Case "主建物"
                    類別代號 = "A"
                Case "地下室", "地下層"
                    類別代號 = "D"
                Case "車位面積(公設內)"
                    類別代號 = "H"
                Case "車位面積(產權獨立)"
                    類別代號 = "I"
                Case "附屬物"
                    類別代號 = "B"
                Case "庭院坪數"
                    類別代號 = "F"
                Case "增建"
                    類別代號 = "G"
            End Select

            Dim sql2 As String = ""
            If DDL_level2.Visible = True Then
                名稱 = ""
                類別代號 = ""
                Select Case Trim(Me.DDL_level2.SelectedItem.Text)
                    Case "層次"
                        名稱 = "主建物"
                        類別代號 = "A"
                    Case "附屬建物"
                        名稱 = "附屬物"
                        類別代號 = "B"
                    Case "共有部分"
                        名稱 = "共有部分"
                        類別代號 = "J"
                End Select
                sql2 = "insert 資料_委賣物件資料表自訂面積細項(類別代號,類別,項目名稱,店代號) values ('" & 類別代號 & "','" & 名稱 & "','" & Trim(Me.TextBox11.Text) & "','" & Request.Cookies("store_id").Value & "')"
            Else
                sql2 = "insert 資料_委賣物件資料表自訂面積細項(類別代號,類別,項目名稱,店代號) values ('" & 類別代號 & "','" & Trim(Me.DropDownList2.SelectedItem.Text) & "','" & Trim(Me.TextBox11.Text) & "','" & Request.Cookies("store_id").Value & "')"
            End If
            Dim cmd As New SqlCommand(sql2, conn)
            cmd.CommandType = CommandType.Text

            conn.Open()

            cmd.ExecuteNonQuery()

            conn.Close()
            conn.Dispose()
        End If
    End Sub

    Sub clear()
        類別名稱()
        項目名稱("主建物")
        Me.Label4.Text = "0"
        TextBox25.Text = "" '建號
        TextBox77.Text = "" '總面積(㎡)
        TextBox73.Text = "" '總面積(坪)
        TextBox1.Text = ""  '權利範圍1分母
        TextBox3.Text = ""  '權利範圍1分子
        TextBox44.Text = ""  '權利範圍2分母
        TextBox45.Text = ""  '權利範圍2分子
        TextBox22.Text = "" '持有面積(㎡)
        TextBox24.Text = "" '持有面積(坪) 

        TextBox11.Text = "" '自訂項目名稱
        TextBox11.Visible = False

        '1040616-V2版新增
        '土地使用分區

        DropDownList69.Visible = False
        DropDownList69.SelectedValue = "請選擇"

        DropDownList65.SelectedValue = "請選擇"
        DropDownList65_SelectedIndexChanged(Nothing, Nothing)
        DropDownList65.Visible = False
        DropDownList66.Visible = False

        '增建用途
        Me.TextBox254.Text = ""
        Me.TextBox254.Visible = False
        '增建日期
        Date8.Text = ""
        Date8.Visible = False
        seeday4.Visible = False

        '權利人
        Me.TextBox255.Text = ""
        Me.TextBox255.Visible = False

        '法定建蔽率
        Me.TextBox256.Text = ""
        Me.TextBox256.Visible = False
        Label74.Visible = False

        '法定容積率
        Me.TextBox257.Text = ""
        Me.TextBox257.Visible = False
        Label75.Visible = False

        Label69.Text = "建號"

        Label70.Text = ""
        Label71.Text = ""
        Label72.Text = ""
        Label73.Text = ""

        '是否為公設
        CheckBox98.Checked = False
        '是否為車位
        CheckBox99.Checked = False
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        If Me.DropDownList1.SelectedItem.Text = "其他" Then
            Me.TextBox11.Visible = True

            If DropDownList2.SelectedItem.Text = "增建" Then
                Me.Label69.Text = "座落門牌或位置"
            End If
        Else
            Me.TextBox11.Text = ""
            Me.TextBox11.Visible = False

            If Me.DropDownList2.SelectedItem.Text = "增建" Then
                If Me.DropDownList1.SelectedItem.Text = "增建" Then
                    Me.Label69.Text = "座落位置"
                Else
                    Me.Label69.Text = "座落門牌"
                End If
            End If
        End If
    End Sub

    '權利範圍1(分母)
    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        owner()

        'AJAX-焦點
        'ScriptManager1.SetFocus(TextBox3)
    End Sub

    '權利範圍1(分子)
    Protected Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        owner()

        'AJAX-焦點
        'ScriptManager1.SetFocus(TextBox22)
    End Sub

    '權利範圍2(分母)
    Protected Sub TextBox44_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox44.TextChanged
        owner()

        'AJAX-焦點
        'ScriptManager1.SetFocus(TextBox3)
    End Sub

    '權利範圍2(分子)
    Protected Sub TextBox45_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox45.TextChanged
        owner()

        'AJAX-焦點
        'ScriptManager1.SetFocus(TextBox22)
    End Sub


    '總面積(平方公尺=>坪)
    Protected Sub TextBox77_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox77.TextChanged
        owner()

        'AJAX-焦點
        'ScriptManager1.SetFocus(TextBox1)
    End Sub

    '計算實際持有
    Sub owner()
        If IsNumeric(TextBox77.Text) Then   'TextBox77：總面積(平方公尺)
            TextBox73.Text = square(TextBox77.Text) '換算成坪   TextBox73：總面積(坪)
            '權利範圍1
            If IsNumeric(TextBox1.Text) And IsNumeric(TextBox3.Text) Then '用權利範圍1算持有        標示部權利範圍 TextBox1：分母、TextBox3：分子
                TextBox22.Text = TextBox77.Text / TextBox1.Text * TextBox3.Text
                TextBox24.Text = square(TextBox22.Text)
            Else
                TextBox22.Text = TextBox77.Text
                TextBox24.Text = square(TextBox22.Text)
            End If

            'If (IsNumeric(TextBox1.Text) And IsNumeric(TextBox3.Text)) Then '用權利範圍1算持有
            '    TextBox22.Text = TextBox77.Text * (TextBox3.Text / TextBox1.Text)
            '    TextBox24.Text = square(TextBox22.Text)
            'End If

            '權利範圍2
            If IsNumeric(TextBox44.Text) And IsNumeric(TextBox45.Text) Then '用權利範圍2算持有      共有部分車位權利範圍  TextBox44：分母、TextBox45：分子
                TextBox22.Text = TextBox22.Text / TextBox44.Text * TextBox45.Text
                TextBox24.Text = square(TextBox22.Text)
            End If

            'If (IsNumeric(TextBox44.Text) And IsNumeric(TextBox45.Text)) Then '用權利範圍2算持有
            '    TextBox22.Text = TextBox77.Text * (TextBox45.Text / TextBox44.Text)
            '    TextBox24.Text = square(TextBox22.Text)
            'End If

            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
            If Me.DropDownList2.SelectedItem.Text = "主建物" And DDL_level2.SelectedValue = "共有部分" And (DropDownList1.SelectedItem.Text = "車位" Or DropDownList1.SelectedItem.Text = "停車空間" Or (Me.DropDownList1.SelectedItem.Text = "其他" And TextBox11.Text = "停車空間")) Then
                If IsNumeric(TextBox44.Text) And IsNumeric(TextBox45.Text) Then
                    TextBox22.Text = TextBox77.Text / TextBox44.Text * TextBox45.Text
                    TextBox24.Text = square(TextBox22.Text)
                End If
            End If
            'End If

        End If
    End Sub

    '取到小數點第4位(平方公尺=>坪)
    Function square(ByVal value As String)
        '1040521新增參數-小數點後幾位判斷判斷
        'Return Round(value * 0.3025, 4, MidpointRounding.AwayFromZero) '不加第3個參數.ROUBD的模式是4捨6入5成雙.加上第3個參數值.則會4捨5入
        Return Round(value * 0.3025, CType(Float.Text, Integer))
    End Function

    '新增細項
    Protected Sub ImageButton5_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton5.Click
        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    If ImageButton1.visible = True Then
        '        Dim script As String = ""
        '        script += "alert('新增物件，請先按下確認新增，才可輸入謄本資料');"
        '        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        '        Exit Sub
        '    End If
        'End If

        '判斷有無物件編號,無則跳出
        If Trim(Me.TextBox2.Text) = "" Then
            Dim Script As String = ""
            Script += "alert('請先輸入物件編號!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)
            Exit Sub
        End If

        '組合物件編號
        '物件編號-第1碼
        Dim 物件編號 As String = ""

        If Request("oid") = "" Then '新增時
            If ddl契約類別.SelectedValue = "專任" Then
                物件編號 = "1"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                物件編號 = "6"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                物件編號 = "7"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                物件編號 = "5"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                物件編號 = "9"
            End If

            '物件編號-第2-5碼(店代號)+第6-13碼(表單編號)
            If store.SelectedValue = "請選擇" Then
                物件編號 &= Mid(Request.Cookies("store_id").Value, 2) & TextBox2.Text.Trim
            Else
                物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            End If
        Else '修改複製
            物件編號 = Request("oid")
        End If

        '新增時如是空值，給予當下的資料值為預設值(整筆資料未存檔前)
        If Me.Label11.Text = "" And Me.Label12.Text = "" Then
            '物件編號
            Me.Label11.Text = 物件編號

            '店代號
            Me.Label12.Text = store.SelectedValue
        End If

        If Me.DropDownList2.SelectedItem.Text <> "選擇類別名稱" Then
            '如果編號跟店代號跟預設值不一樣，先下列步驟
            If Me.Label11.Text <> 物件編號 Or Me.Label12.Text <> store.SelectedValue Then
                Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
                Dim sql_upt As String = "update 委賣物件資料表_面積細項 set 物件編號='" & 物件編號 & "',店代號='" & store.SelectedValue & "' where 物件編號='" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"
                Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
                cmd_upt.CommandType = CommandType.Text

                conn_upt.Open()

                cmd_upt.ExecuteNonQuery()

                conn_upt.Close()
                conn_upt.Dispose()

                'UPDATE已存資料後給予新的值
                '物件編號
                Me.Label11.Text = 物件編號

                '店代號
                Me.Label12.Text = store.SelectedValue
            End If

            Dim 項目 As String = ""

            If Me.DropDownList1.SelectedItem.Text = "其他" Then
                If Trim(Me.TextBox11.Text) = "" Then
                    項目 = "其他"
                Else
                    項目 = Trim(Me.TextBox11.Text)
                    chk_項目()
                End If
            Else
                If DropDownList1.SelectedItem.Text = "選擇項目名稱" Then
                    項目 = "其他"
                Else
                    項目 = Trim(Me.DropDownList1.SelectedItem.Text)
                End If
            End If

            conn = New SqlConnection(EGOUPLOADSqlConnStr)
            Dim count As Integer
            Dim sql2 As String = "select MAX(流水號) as 流水號 from 委賣物件資料表_面積細項 With(NoLock) where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"
            Dim cmd2 As New SqlCommand(sql2, conn)
            cmd2.CommandType = CommandType.Text
            adpt = New SqlDataAdapter(sql2, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")

            count = 0
            If table2.Rows.Count > 0 And Not IsDBNull(table2.Rows(0)("流水號")) Then
                count = table2.Rows(0)("流水號") + 1
            End If

            '更新委賣物件資料表_細項所有權人 的流水號
            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
            '    'eip_usual.Show("新增狀態 hidtempcode:" & hidtempcode.Value & ",count:" & count)
            'End If
            If hidtempcode.Value.Length > 5 Then
                Using con_細項所有權人 As New SqlConnection(EGOUPLOADSqlConnStr)
                    con_細項所有權人.Open()

                    hidtempcode.Value = hidtempcode.Value.Replace(",,", ",")
                    If Left(hidtempcode.Value, 1) = "," Then
                        hidtempcode.Value = hidtempcode.Value.Substring(1, hidtempcode.Value.Length - 1)
                    End If
                    If Right(hidtempcode.Value, 1) = "," Then
                        hidtempcode.Value = hidtempcode.Value.Substring(0, hidtempcode.Value.Length - 1)
                    End If

                    Dim updatstr As String = "update 委賣物件資料表_細項所有權人 set 細項流水號 = '" & count & "' where TempCode in ('" & hidtempcode.Value.Replace(",", "','") & "')" 'and 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"

                    'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
                    '    eip_usual.Show("updatstr:" & updatstr)
                    'End If

                    Using cmd_細項所有權人 As New SqlCommand(updatstr, con_細項所有權人)
                        Try
                            cmd_細項所有權人.ExecuteNonQuery()
                            hidtempcode.Value = ""

                            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                            '    'eip_usual.Show("after:" & hidtempcode.Value)
                            'End If
                        Catch ex As Exception
                            Response.Write("錯誤:" & updatstr)
                            Response.End()
                        End Try
                    End Using
                End Using
            End If

            '避免持份未填值.導致出錯-------------
            '總面積平方公尺有值但坪沒有值的情況下 幫填寫
            If TextBox77.Text.Length > 0 And TextBox73.Text.Length = 0 Then
                If IsNumeric(TextBox77.Text) Then
                    TextBox73.Text = Round(CDec(TextBox77.Text) * 0.3025, 4)
                End If
            End If
            '權利範圍1
            If Trim(Me.TextBox1.Text) = "" Then
                Me.TextBox1.Text = "1"
            End If
            If Trim(Me.TextBox3.Text) = "" Then
                Me.TextBox3.Text = "1"
            End If
            '權利範圍2
            If Trim(Me.TextBox44.Text) = "" Then
                Me.TextBox44.Text = "1"
            End If
            If Trim(Me.TextBox45.Text) = "" Then
                Me.TextBox45.Text = "1"
            End If
            '----------------------------------

            '實際持有平方公尺 & 坪
            If Trim(Me.TextBox22.Text) = "" Then
                Me.TextBox22.Text = "0"
            End If
            If Trim(Me.TextBox24.Text) = "" Then
                Me.TextBox24.Text = "0"
            End If
            If DropDownList2.SelectedValue = "車位面積(公設內)" And TextBox73.Text = "" Then
                TextBox73.Text = "0"
            End If

            Dim sql As String = ""
            sql &= "insert 委賣物件資料表_面積細項(物件編號, 流水號, 建號, 類別,項目名稱,總面積平方公尺,總面積坪,權利範圍1分母,權利範圍1分子,權利範圍2分母,權利範圍2分子,實際持有平方公尺,實際持有坪,店代號,使用分區,增建用途,增建完成日期,管制,所有權人,法定建蔽率,法定容積率,DL_level2_selectindex,是否為公設,是否為車位) values ('" & Me.Label11.Text & "','" & count & "','" & TextBox25.Text & "','" & DropDownList2.SelectedValue & "','" & 項目 & "','" & TextBox77.Text & "','" & TextBox73.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "','" & TextBox44.Text & "','" & TextBox45.Text & "','" & TextBox22.Text & "','" & TextBox24.Text & "','" & Me.Label12.Text & "',"

            '1050616-V2版新增
            '土地使用分區-物件用途 
            If DropDownList66.Visible = True Then
                If DropDownList66.SelectedValue = "請選擇" Then
                    sql &= "'" & DropDownList65.SelectedValue & "',"
                Else
                    sql &= "'" & DropDownList66.SelectedValue & "',"
                End If
            Else
                sql &= "'" & DropDownList65.SelectedValue & "',"
            End If

            '增建用途
            sql &= "'" & TextBox254.Text & "',"
            '增建完成日期
            sql &= "'" & Date8.Text & "',"

            '土地使用分區-管制 
            If DropDownList69.Visible = True Then
                If DropDownList69.SelectedValue = "請選擇" Then
                    sql &= "'',"
                Else
                    sql &= "'" & DropDownList69.SelectedValue & "',"
                End If
            Else
                sql &= "'',"
            End If

            '所有權人
            sql &= "'" & TextBox255.Text & "',"

            '法定建蔽率
            sql &= "'" & TextBox256.Text & "',"
            '法定容積率
            sql &= "'" & TextBox257.Text & "',"
            'DDL_level2 
            sql &= DDL_level2.SelectedIndex & ","
            '是否為公設
            sql &= "'" & IIf(CheckBox98.Checked = True, "Y", "N") & "',"
            '是否為車位
            sql &= "'" & IIf(CheckBox99.Checked = True, "Y", "N") & "'"
            sql &= ")"
            Dim cmd As New SqlCommand(sql, conn)
            cmd.CommandType = CommandType.Text

            conn.Open()
            Try
                cmd.ExecuteNonQuery()

                'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then

                'Else
                Dim script As String = ""
                script += "alert('新增成功!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                'End If

                '恢復成原始狀態-----------------------
                clear()
                類別名稱()
                項目名稱("主建物")
                DDL_level2.SelectedIndex = 0
                '-------------------------------------

                '讀取資料
                Load_Data("OLD")
            Catch
                'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then

                'Else
                Response.Write(sql)
                Response.Write("<br>請洽資訊人員")
                Response.End()

                Dim script As String = ""
                script += "alert('新增失敗!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                'End If
            End Try

            'cmd2.ExecuteNonQuery()
            conn.Close()
            conn.Dispose()
        Else
            Dim script As String = ""
            script += "alert('請選擇類別名稱!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        End If

        '計算面積
        Total()

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        寫入坪數()
        'End If

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    If ImageButton1.visible = True Then
        '        ImageButton1_Click(sender, e)
        '    Else
        '        ImageButton12_Click(sender, e)
        '    End If
        '    'Response.Write(TextBox28.TEXT)
        '    'Exit Sub
        '    'ImageButton12_Click(sender, e)
        'Else
        '    'trans = "False"
        '    'If ImageButton1.visible = True Then
        '    '    ImageButton1_Click(sender, e)
        '    'Else
        '    '    ImageButton12_Click(sender, e)
        '    'End If
        '    'If trans = "False" Then
        '    '    Exit Sub
        '    'End If
        'End If
    End Sub

    '計算面積
    Sub Total()
        Dim gr As GridViewRow

        '土地面積細項
        Dim 土地面積_持有_平方公尺 As Double = 0
        Dim 土地面積_持有_坪 As Double = 0
        Dim 土地面積_所有_平方公尺 As Double = 0
        Dim 土地面積_所有_坪 As Double = 0
        Dim 土地面積_出售_平方公尺 As Double = 0
        Dim 土地面積_出售_坪 As Double = 0

        '公共設施細項
        Dim 公共設施_持有_平方公尺 As Double = 0
        Dim 公共設施_持有_坪 As Double = 0
        Dim 公共設施_出售_平方公尺 As Double = 0
        Dim 公共設施_出售_坪 As Double = 0

        '主建物細項
        Dim 主建物_持有_平方公尺 As Double = 0
        Dim 主建物_持有_坪 As Double = 0

        '地下室細項
        Dim 地下室_持有_平方公尺 As Double = 0
        Dim 地下室_持有_坪 As Double = 0

        '車位面積_公設內細項
        Dim 車位面積_公設內_持有_平方公尺 As Double = 0
        Dim 車位面積_公設內_持有_坪 As Double = 0
        Dim 車位面積_公設內_出售_平方公尺 As Double = 0
        Dim 車位面積_公設內_出售_坪 As Double = 0

        '車位面積_公設內細項
        Dim 車位面積_產權獨立_持有_平方公尺 As Double = 0
        Dim 車位面積_產權獨立_持有_坪 As Double = 0

        '附屬物細項
        Dim 附屬物_持有_平方公尺 As Double = 0
        Dim 附屬物_持有_坪 As Double = 0

        '增建細項
        Dim 增建_持有_平方公尺 As Double = 0
        Dim 增建_持有_坪 As Double = 0

        '庭院坪數細項
        Dim 庭院坪數_持有_平方公尺 As Double = 0
        Dim 庭院坪數_持有_坪 As Double = 0

        '總面積
        Dim 總面積_平方公尺 As Double = 0
        Dim 總面積_坪 As Double = 0

        '土地面積-地號
        Dim 土地面積_地號 As String = ""
        '建物面積-建號
        Dim 建物面積_建號 As String = ""
        '權利範圍
        Dim 權利範圍 As String = ""

        RadioButtonList3.SelectedIndex = 1
        'UpdatePanel21.visible = False

        For Each gr In GridView1.Rows
            Dim lbl類別 As Label = CType(gr.FindControl("Label37"), Label)
            Dim lbl子類別 As Label = CType(gr.FindControl("Label24"), Label)
            '項目
            Dim lbl項目 As Label = CType(gr.FindControl("Label38"), Label)

            Dim lbl總面積平方公尺 As Label = CType(gr.FindControl("Label40"), Label)
            Dim lbl總面積坪 As Label = CType(gr.FindControl("Label41"), Label)

            Dim lbl實際持有平方公尺 As Label = CType(gr.FindControl("Label43"), Label)
            Dim lbl實際持有坪 As Label = CType(gr.FindControl("Label44"), Label)

            Dim lbl實際出售平方公尺 As Label = CType(gr.FindControl("Label88"), Label)
            Dim lbl實際出售坪 As Label = CType(gr.FindControl("Label89"), Label)

            Dim lbl土地面積地號 As Label = CType(gr.FindControl("Label39"), Label)

            Dim lbl權利範圍 As Label = CType(gr.FindControl("Label42"), Label)
            Dim lbl權利範圍2 As Label = CType(gr.FindControl("Label20"), Label)

            Dim lbl是否為公設 As Label = CType(gr.FindControl("Label98"), Label)
            Dim lbl是否為車位 As Label = CType(gr.FindControl("Label99"), Label)

            Dim lbl_子類別 As Label = CType(gr.FindControl("Label24"), Label)
            If lbl_子類別.Text = "共有部分" Then
                lbl類別.Text = "公共設施"
            End If
            If lbl_子類別.Text = "附屬建物" Then
                lbl類別.Text = "附屬物"
            End If

            Select Case Trim(lbl類別.Text)
                Case "土地面積"

                    '持有
                    'If Trim(lbl實際持有平方公尺.Text) <> "" Then
                    '    土地面積_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    'End If
                    'If Trim(lbl實際持有坪.Text) <> "" Then
                    '    土地面積_持有_坪 += CType(lbl實際持有坪.Text, Double)
                    'End If

                    '所有(基地面積)1040113新增
                    If Trim(lbl總面積平方公尺.Text) <> "" Then
                        土地面積_所有_平方公尺 += CType(lbl總面積平方公尺.Text, Double)
                    End If
                    If Trim(lbl總面積坪.Text) <> "" Then
                        土地面積_所有_坪 += CType(lbl總面積坪.Text, Double)
                    End If

                    '出售1050614新增
                    If Trim(lbl實際出售平方公尺.Text) <> "" Then
                        土地面積_出售_平方公尺 += CType(lbl實際出售平方公尺.Text, Double)
                    ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                        土地面積_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    End If

                    If Trim(lbl實際出售坪.Text) <> "" Then
                        土地面積_出售_坪 += CType(lbl實際出售坪.Text, Double)
                    ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                        土地面積_持有_坪 += CType(lbl實際持有坪.Text, Double)
                    End If


                    '地號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        土地面積_地號 = Replace(土地面積_地號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If

                    '1040522新增-------------------------------------------------------------------------------
                    'step1-先取得自訂小數點位數後的數字
                    lbl實際持有坪.Text = Round(CType(lbl實際持有坪.Text, Double), CType(Float.Text, Integer), MidpointRounding.AwayFromZero)

                    'step2-去掉尾巴的0
                    If Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 4) = "0000" Then
                        lbl實際持有坪.Text = Int(lbl實際持有坪.Text)
                    ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 3) = "000" Then
                        lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 3)
                    ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 2) = "00" Then
                        lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 2)
                    ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 1) = "0" Then
                        lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 1)
                    Else
                        lbl實際持有坪.Text = lbl實際持有坪.Text
                    End If
                    '--------------------------------------------------------------------------------------------

                    Dim lbl權利範圍_ALL As String = ""

                    If Left(lbl權利範圍.Text, 2) = "A." Then
                        lbl權利範圍_ALL &= Mid(lbl權利範圍.Text, 3)
                    Else
                        lbl權利範圍_ALL &= lbl權利範圍.Text
                    End If

                    If Left(lbl權利範圍2.Text, 2) = "B." Then
                        lbl權利範圍_ALL &= "*" & Mid(lbl權利範圍2.Text, 3)
                    End If


                    '權利範圍
                    If 權利範圍 = "" Then
                        '持份面積 += "地號: " & 地號 & " ,權利範圍: " & 權利範圍分母 & " 分之 " & 權利範圍分子 & " ,約 " & 實際持有坪 & " 坪"
                        權利範圍 += "權利範圍: " & lbl權利範圍_ALL & " ,約 " & lbl實際持有坪.Text & " 坪"
                    Else
                        '持份面積 += ",地號 " & 地號 & " ,權利範圍: " & 權利範圍分母 & " 分之 " & 權利範圍分子 & " ,約 " & 實際持有坪 & " 坪"
                        權利範圍 += ",權利範圍: " & lbl權利範圍_ALL & " ,約 " & lbl實際持有坪.Text & " 坪"
                    End If

                Case "公共設施"
                    '公共設施 的車位要算在 公設車位 這項目內 20160428
                    If lbl項目.Text = "停車空間" Or lbl項目.Text = "車位" Or lbl是否為車位.Text = "Y" Then

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            車位面積_公設內_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            車位面積_公設內_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            車位面積_公設內_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            車位面積_公設內_持有_坪 += tot
                            總面積_坪 += tot
                        End If

                    Else
                        'If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        '    公共設施_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                        '    總面積_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                        'End If
                        'If Trim(lbl實際持有坪.Text) <> "" Then
                        '    公共設施_持有_坪 += CType(lbl實際持有坪.Text, Double)
                        '    總面積_坪 += CType(lbl實際持有坪.Text, Double)
                        'End If

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            公共設施_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)

                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            公共設施_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            公共設施_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            公共設施_持有_坪 += tot
                            總面積_坪 += tot
                        End If

                    End If


                    '建號1040113修正-公設不帶入建號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        '建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If
                Case "主建物"

                    Select Case (lbl_子類別.Text)
                        Case "0"
                            lbl_子類別.Text = "層次"
                        Case "1"
                            lbl_子類別.Text = "附屬建物"
                        Case "2"
                            lbl_子類別.Text = "共有部分"

                    End Select

                    'Response.Write(lbl類別.Text & lbl_子類別.Text)

                    If lbl是否為公設.Text = "Y" Then

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            公共設施_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)

                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            公共設施_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            公共設施_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            公共設施_持有_坪 += tot
                            總面積_坪 += tot
                        End If

                    ElseIf (lbl子類別.Text = "層次" And lbl項目.Text = "車位") Or lbl是否為車位.Text = "Y" Then

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            車位面積_公設內_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            車位面積_公設內_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            車位面積_公設內_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            車位面積_公設內_持有_坪 += tot
                            總面積_坪 += tot
                        End If




                    Else

                        If Trim(lbl實際持有平方公尺.Text) <> "" Then
                            主建物_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                            總面積_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                        End If
                        If Trim(lbl實際持有坪.Text) <> "" Then
                            主建物_持有_坪 += CType(lbl實際持有坪.Text, Double)
                            總面積_坪 += CType(lbl實際持有坪.Text, Double)
                        End If

                        '建號
                        If Trim(lbl土地面積地號.Text) <> "" Then
                            建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                        End If

                    End If





                Case "地下室", "地下層"

                    If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        地下室_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                        總面積_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    End If
                    If Trim(lbl實際持有坪.Text) <> "" Then
                        地下室_持有_坪 += CType(lbl實際持有坪.Text, Double)
                        總面積_坪 += CType(lbl實際持有坪.Text, Double)
                    End If


                    '建號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If
                Case "車位面積(公設內)"

                    If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        Dim tot As Double = 0
                        If lbl實際持有平方公尺.Text = "--" Then
                            tot = 0
                        Else
                            tot = CDec(lbl實際持有平方公尺.Text)
                        End If
                        車位面積_公設內_持有_平方公尺 += tot
                        總面積_平方公尺 += CType(tot, Double)
                    End If
                    If Trim(lbl實際持有坪.Text) <> "" Then
                        Dim tot As Double = 0
                        If lbl實際持有坪.Text = "--" Then
                            tot = 0
                        Else
                            tot = CDec(lbl實際持有坪.Text)
                        End If

                        車位面積_公設內_持有_坪 += tot
                        總面積_坪 += tot
                    End If

                    '建號-1040113修正-公設不帶入建號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        '建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If
                Case "車位面積(產權獨立)"

                    If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        車位面積_產權獨立_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    End If
                    If Trim(lbl實際持有坪.Text) <> "" Then
                        車位面積_產權獨立_持有_坪 += CType(lbl實際持有坪.Text, Double)
                    End If

                    '建號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If
                Case "附屬物"

                    If lbl項目.Text = "車位" Or lbl是否為車位.Text = "Y" Then

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            車位面積_公設內_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            車位面積_公設內_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            車位面積_公設內_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            車位面積_公設內_持有_坪 += tot
                            總面積_坪 += tot
                        End If

                    ElseIf lbl是否為公設.Text = "Y" Then

                        If Trim(lbl實際出售平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售平方公尺.Text)
                            End If
                            公共設施_出售_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        ElseIf Trim(lbl實際持有平方公尺.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有平方公尺.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有平方公尺.Text)
                            End If
                            公共設施_持有_平方公尺 += tot
                            總面積_平方公尺 += CType(tot, Double)
                        End If

                        If Trim(lbl實際出售坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際出售坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際出售坪.Text)
                            End If

                            公共設施_出售_坪 += tot
                            總面積_坪 += tot
                        ElseIf Trim(lbl實際持有坪.Text) <> "" Then
                            Dim tot As Double = 0
                            If lbl實際持有坪.Text = "--" Then
                                tot = 0
                            Else
                                tot = CDec(lbl實際持有坪.Text)
                            End If

                            公共設施_持有_坪 += tot
                            總面積_坪 += tot
                        End If

                    Else
                        If Trim(lbl實際持有平方公尺.Text) <> "" Then
                            附屬物_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                            總面積_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                        End If
                        If Trim(lbl實際持有坪.Text) <> "" Then
                            附屬物_持有_坪 += CType(lbl實際持有坪.Text, Double)
                            總面積_坪 += CType(lbl實際持有坪.Text, Double)
                        End If
                    End If

                    '建號
                    If Trim(lbl土地面積地號.Text) <> "" Then
                        建物面積_建號 = Replace(建物面積_建號, CType(lbl土地面積地號.Text, String) & ",", "") & CType(lbl土地面積地號.Text, String) & ","
                    End If
                Case "增建"

                    If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        增建_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    End If
                    If Trim(lbl實際持有坪.Text) <> "" Then
                        增建_持有_坪 += CType(lbl實際持有坪.Text, Double)
                    End If
                    'UpdatePanel21.visible = True
                    RadioButtonList3.SelectedIndex = 0
                Case "庭院坪數"

                    If Trim(lbl實際持有平方公尺.Text) <> "" Then
                        庭院坪數_持有_平方公尺 += CType(lbl實際持有平方公尺.Text, Double)
                    End If
                    If Trim(lbl實際持有坪.Text) <> "" Then
                        庭院坪數_持有_坪 += CType(lbl實際持有坪.Text, Double)
                    End If

            End Select

        Next

        '20130910小豪新增判斷-不要填回0

        '傳回值
        '主建物
        If 主建物_持有_平方公尺 = 0 Then
            TextBox5.Text = ""
        Else
            'If Right(主建物_持有_平方公尺.ToString, 5) = ".0000" Then
            '    主建物_持有_平方公尺 = Replace(主建物_持有_平方公尺, ".0000", "")
            'End If

            TextBox5.Text = 主建物_持有_平方公尺
            自訂坪數_小數點位數(TextBox5, "F")
        End If

        If 主建物_持有_坪 = 0 Then
            TextBox6.Text = ""
        Else
            'If Right(主建物_持有_坪.ToString, 5) = ".0000" Then
            '    主建物_持有_坪 = Replace(主建物_持有_坪, ".0000", "")
            'End If

            TextBox6.Text = 主建物_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox6, "T")
        End If

        '附屬物
        If 附屬物_持有_平方公尺 = 0 Then
            TextBox7.Text = ""
        Else
            'If Right(附屬物_持有_平方公尺.ToString, 5) = ".0000" Then
            '    附屬物_持有_平方公尺 = Replace(附屬物_持有_平方公尺, ".0000", "")
            'End If

            TextBox7.Text = 附屬物_持有_平方公尺
            自訂坪數_小數點位數(TextBox7, "F")
        End If

        If 附屬物_持有_坪 = 0 Then
            TextBox8.Text = ""
        Else
            'If Right(附屬物_持有_坪.ToString, 5) = ".0000" Then
            '    附屬物_持有_坪 = Replace(附屬物_持有_坪, ".0000", "")
            'End If

            TextBox8.Text = 附屬物_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox8, "T")
        End If


        '公共設施
        If 公共設施_出售_平方公尺 = 0 And 公共設施_持有_平方公尺 = 0 Then

            TextBox9.Text = ""

        ElseIf 公共設施_出售_平方公尺 = 0 Then

            TextBox9.Text = 公共設施_持有_平方公尺
            自訂坪數_小數點位數(TextBox9, "F")

        Else

            TextBox9.Text = 公共設施_出售_平方公尺 + 公共設施_持有_平方公尺
            自訂坪數_小數點位數(TextBox9, "F")

        End If

        If 公共設施_出售_坪 = 0 And 公共設施_持有_坪 = 0 Then

            TextBox10.Text = ""

        ElseIf 公共設施_出售_坪 = 0 Then

            TextBox10.Text = 公共設施_持有_坪
            自訂坪數_小數點位數(TextBox10, "T")

        Else

            TextBox10.Text = 公共設施_出售_坪 + 公共設施_持有_坪
            自訂坪數_小數點位數(TextBox10, "T")

        End If


        '地下室
        If 地下室_持有_平方公尺 = 0 Then
            TextBox19.Text = ""
        Else
            'If Right(地下室_持有_平方公尺.ToString, 5) = ".0000" Then
            '    地下室_持有_平方公尺 = Replace(地下室_持有_平方公尺, ".0000", "")
            'End If

            TextBox19.Text = 地下室_持有_平方公尺
            自訂坪數_小數點位數(TextBox19, "F")
        End If

        If 地下室_持有_坪 = 0 Then
            TextBox20.Text = ""
        Else
            'If Right(地下室_持有_坪.ToString, 5) = ".0000" Then
            '    地下室_持有_坪 = Replace(地下室_持有_坪, ".0000", "")
            'End If

            TextBox20.Text = 地下室_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox20, "T")
        End If


        '車位面積(公設內)
        If 車位面積_公設內_出售_平方公尺 <> 0 And 車位面積_公設內_持有_平方公尺 <> 0 Then
            TextBox21.Text = 車位面積_公設內_出售_平方公尺 + 車位面積_公設內_持有_平方公尺
            自訂坪數_小數點位數(TextBox21, "F")
        ElseIf 車位面積_公設內_出售_平方公尺 <> 0 Then
            TextBox21.Text = 車位面積_公設內_出售_平方公尺
            自訂坪數_小數點位數(TextBox21, "F")
        ElseIf 車位面積_公設內_持有_平方公尺 <> 0 Then
            TextBox21.Text = 車位面積_公設內_持有_平方公尺
            自訂坪數_小數點位數(TextBox21, "F")
        Else
            TextBox21.Text = ""
        End If



        'If 車位面積_公設內_持有_平方公尺 = 0 Then
        '    TextBox21.Text = ""
        'Else
        '    'If Right(車位面積_公設內_持有_平方公尺.ToString, 5) = ".0000" Then
        '    '    車位面積_公設內_持有_平方公尺 = Replace(車位面積_公設內_持有_平方公尺, ".0000", "")
        '    'End If

        '    TextBox21.Text = 車位面積_公設內_持有_平方公尺
        '    自訂坪數_小數點位數(TextBox21, "F")
        'End If


        If 車位面積_公設內_出售_坪 <> 0 And 車位面積_公設內_持有_坪 <> 0 Then
            TextBox23.Text = 車位面積_公設內_出售_坪 + 車位面積_公設內_持有_坪
            自訂坪數_小數點位數(TextBox23, "F")
        ElseIf 車位面積_公設內_出售_坪 <> 0 Then
            TextBox23.Text = 車位面積_公設內_出售_坪
            自訂坪數_小數點位數(TextBox23, "F")
        ElseIf 車位面積_公設內_持有_坪 <> 0 Then
            TextBox23.Text = 車位面積_公設內_持有_坪
            自訂坪數_小數點位數(TextBox23, "F")
        Else
            TextBox23.Text = ""
        End If

        'If 車位面積_公設內_持有_坪 = 0 Then
        '    TextBox23.Text = ""
        'Else
        '    'If Right(車位面積_公設內_持有_坪.ToString, 5) = ".0000" Then
        '    '    車位面積_公設內_持有_坪 = Replace(車位面積_公設內_持有_坪, ".0000", "")
        '    'End If

        '    TextBox23.Text = 車位面積_公設內_持有_坪
        '    '1040522修正-新增自訂小數點位數判斷-坪
        '    自訂坪數_小數點位數(TextBox23, "T")
        'End If


        '車位面積(產權獨立)
        If 車位面積_產權獨立_持有_平方公尺 = 0 Then
            TextBox26.Text = ""
        Else
            'If Right(車位面積_產權獨立_持有_平方公尺.ToString, 5) = ".0000" Then
            '    車位面積_產權獨立_持有_平方公尺 = Replace(車位面積_產權獨立_持有_平方公尺, ".0000", "")
            'End If

            TextBox26.Text = 車位面積_產權獨立_持有_平方公尺
            自訂坪數_小數點位數(TextBox26, "F")
        End If

        If 車位面積_產權獨立_持有_坪 = 0 Then
            TextBox27.Text = ""
        Else
            'If Right(車位面積_產權獨立_持有_坪.ToString, 5) = ".0000" Then
            '    車位面積_產權獨立_持有_坪 = Replace(車位面積_產權獨立_持有_坪, ".0000", "")
            'End If

            TextBox27.Text = 車位面積_產權獨立_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox27, "T")
        End If



        '總面積
        If 總面積_平方公尺 = 0 Then
            TextBox28.Text = "0"
        Else
            'If Right(總面積_平方公尺.ToString, 5) = ".0000" Then
            '    總面積_平方公尺 = Replace(總面積_平方公尺, ".0000", "")
            'End If

            TextBox28.Text = 總面積_平方公尺
            自訂坪數_小數點位數(TextBox28, "F")
        End If

        If 總面積_坪 = 0 Then
            TextBox29.Text = "0"
        Else
            'If Right(總面積_坪.ToString, 5) = ".0000" Then
            '    總面積_坪 = Replace(總面積_坪, ".0000", "")
            'End If

            TextBox29.Text = 總面積_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox29, "T")
        End If


        '土地面積細項

        If 土地面積_出售_平方公尺 <> 0 And 土地面積_持有_平方公尺 <> 0 Then
            TextBox30.Text = 土地面積_出售_平方公尺 + 土地面積_持有_平方公尺
            自訂坪數_小數點位數(TextBox30, "F")
        ElseIf 土地面積_出售_平方公尺 <> 0 Then
            TextBox30.Text = 土地面積_出售_平方公尺
            自訂坪數_小數點位數(TextBox30, "F")
        ElseIf 土地面積_持有_平方公尺 <> 0 Then
            TextBox30.Text = 土地面積_持有_平方公尺
            自訂坪數_小數點位數(TextBox30, "F")
        Else
            TextBox30.Text = ""
        End If

        'If 土地面積_持有_平方公尺 = 0 Then
        '    TextBox30.Text = ""
        'Else
        '    'If Right(土地面積_持有_平方公尺.ToString, 5) = ".0000" Then
        '    '    土地面積_持有_平方公尺 = Replace(土地面積_持有_平方公尺, ".0000", "")
        '    'End If

        '    TextBox30.Text = 土地面積_持有_平方公尺
        '    自訂坪數_小數點位數(TextBox30, "F")
        'End If
        If 土地面積_出售_坪 <> 0 And 土地面積_持有_坪 <> 0 Then
            TextBox31.Text = 土地面積_出售_坪 + 土地面積_持有_坪
            自訂坪數_小數點位數(TextBox31, "F")
        ElseIf 土地面積_出售_坪 <> 0 Then
            TextBox31.Text = 土地面積_出售_坪
            自訂坪數_小數點位數(TextBox31, "F")
        ElseIf 土地面積_持有_坪 <> 0 Then
            TextBox31.Text = 土地面積_持有_坪
            自訂坪數_小數點位數(TextBox31, "F")
        Else
            TextBox31.Text = ""
        End If

        'If 土地面積_持有_坪 = 0 Then
        '    TextBox31.Text = ""
        'Else
        '    'If Right(土地面積_持有_坪.ToString, 5) = ".0000" Then
        '    '    土地面積_持有_坪 = Replace(土地面積_持有_坪, ".0000", "")
        '    'End If

        '    TextBox31.Text = 土地面積_持有_坪
        '    '1040522修正-新增自訂小數點位數判斷-坪
        '    自訂坪數_小數點位數(TextBox31, "T")
        'End If

        ''基地面積-1040113新增
        'If 土地面積_所有_平方公尺 = 0 Then
        '    TextBox231.Text = ""
        'Else
        '    'If Right(土地面積_所有_平方公尺.ToString, 5) = ".0000" Then
        '    '    土地面積_所有_平方公尺 = Replace(土地面積_所有_平方公尺, ".0000", "")
        '    'End If

        '    TextBox231.Text = 土地面積_所有_平方公尺
        '    自訂坪數_小數點位數(TextBox231, "F")
        'End If

        'If 土地面積_所有_坪 = 0 Then
        '    TextBox230.Text = ""
        'Else
        '    'If Right(土地面積_所有_坪.ToString, 5) = ".0000" Then
        '    '    土地面積_所有_坪 = Replace(土地面積_所有_坪, ".0000", "")
        '    'End If

        '    TextBox230.Text = 土地面積_所有_坪
        '    '1040522修正-新增自訂小數點位數判斷-坪
        '    自訂坪數_小數點位數(TextBox230, "T")
        'End If


        '庭院坪數細項
        If 庭院坪數_持有_平方公尺 = 0 Then
            TextBox32.Text = ""
        Else
            'If Right(庭院坪數_持有_平方公尺.ToString, 5) = ".0000" Then
            '    庭院坪數_持有_平方公尺 = Replace(庭院坪數_持有_平方公尺, ".0000", "")
            'End If

            TextBox32.Text = 庭院坪數_持有_平方公尺
            自訂坪數_小數點位數(TextBox32, "F")
        End If

        If 庭院坪數_持有_坪 = 0 Then
            TextBox33.Text = ""
        Else
            'If Right(庭院坪數_持有_坪.ToString, 5) = ".0000" Then
            '    庭院坪數_持有_坪 = Replace(庭院坪數_持有_坪, ".0000", "")
            'End If

            TextBox33.Text = 庭院坪數_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox33, "T")
        End If


        '增建細項
        If 增建_持有_平方公尺 = 0 Then
            TextBox34.Text = ""
        Else
            'If Right(增建_持有_平方公尺.ToString, 5) = ".0000" Then
            '    增建_持有_平方公尺 = Replace(增建_持有_平方公尺, ".0000", "")
            'End If

            TextBox34.Text = 增建_持有_平方公尺
            自訂坪數_小數點位數(TextBox34, "F")
        End If

        If 增建_持有_坪 = 0 Then
            TextBox35.Text = ""
        Else
            'If Right(增建_持有_坪.ToString, 5) = ".0000" Then
            '    增建_持有_坪 = Replace(增建_持有_坪, ".0000", "")
            'End If

            TextBox35.Text = 增建_持有_坪
            '1040522修正-新增自訂小數點位數判斷-坪
            自訂坪數_小數點位數(TextBox35, "T")
        End If


        '土地面積_地號
        TextBox17.Text = 土地面積_地號
        '建物面積_建號
        TextBox18.Text = 建物面積_建號
        ''不動產說明書_基地面積
        'input30.Value = 權利範圍




    End Sub


    '讀取細項面積內容
    Sub Load_Data(ByVal cls As String, Optional ByVal alseCallOLD As Boolean = False)
        'If Request.Cookies("webfly_empno").Value = "00P" Then
        'Response.Write(cls & "_" & alseCallOLD)
        'eip_usual.show(cls)
        'End If

        '清空所有細項
        clear()

        '判斷物件維護權限
        myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")
        'Response.Write("權限" & Request.Cookies("webfly_empno").Value)
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        If cls = "OLD" Then
            '1040616-V2版修正
            sql = "Select *,(權利範圍1分子+'/'+權利範圍1分母) as 權利範圍,(權利範圍2分子+'/'+權利範圍2分母) as 權利範圍2 From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'  order by 類別,項目名稱 "
            '加入空表單
            Dim htmlstr As String = "<table cellspacing=""0"" rules=""all"" border=""1"" id=""Table_right"" style=""width:100%;border-collapse:collapse;"">"
            htmlstr += "<tr><td colspan=""7"" >權利人部分</td></tr>"
            htmlstr += "<tr><th scope=""col"">&nbsp;</th><th scope=""col"">姓名</th><th scope=""col"">權利範圍</th><th scope=""col"">面積</th><th scope=""col"">出售權利種類</th><th scope=""col"">所有權種類</th><th scope=""col""></th></tr>"
            Dim txt_name As String = "<input type='text' id='tb_Textbox50_0' style='width:50px;'  />"
            Dim txtbox_right As String = " 權利範圍.<input type='text' id='tb_Numerator_1_0' title='分母' style='width:50px;'  /> &nbsp;分之 &nbsp;<input type=text' id='tb_Denominator_1_0' title='分子' style='width:50px;' onblur='calculatorarea(this.id);'  /></br>出售權利範圍.<input type='text'  id='tb_Numerator_2_0' title='分母' style='width:50px;'  /> &nbsp;分之 &nbsp;<input type='text' id='tb_Denominator_2_0' title='分子' style='width:50px;' onblur='calculatorarea2(this.id);'  />"
            Dim lb_holder As String = "持有面積:<span id='tb_Label21_1_0'></span><br />出售面積:<span id='tb_Label21_2_0'></span>"
            Dim select1 As String = "<select id='tb_DropDownList70_0'><option  value='所有權'>所有權</option><option value='地上權' >地上權</option><option value='典權' >典權</option><option value='使用權' >使用權</option></select>"
            Dim bons As String = ""
            'If Request.Cookies("webfly_empno").Value.ToString.ToUpper = "0AEU" Then
            If myobj.M = "1" And myobj.D = "1" Then
                bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            ElseIf myobj.M = "1" And myobj.D = "0" Then
                bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input type='hidden' id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            ElseIf myobj.M = "0" And myobj.D = "1" Then
                bons = "<input type='hidden' id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            Else
                bons = "<input type='hidden' id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input type='hidden' id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            End If
            'Else
            '    bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            'End If
            'bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
            Dim select2 As String = " <select  id='tb_DropDownList71_0'><option value='單獨所有' >單獨所有</option><option  value='分別共有' >分別共有</option><option value='公同共有' >公同共有</option></select>"
            htmlstr += "<tr align=""center""><td><span id=""tb_num0"">1</span></td><td>" & txt_name & "</td><td align=left>" & txtbox_right & "</td><td>" & lb_holder & "</td><td>" & select1 & "</td><td>" & select2 + "</td><td>" + bons + "</td></tr>"
            editTable.Text = htmlstr
        Else
            Dim 編號 As Array = Split(cls, ",")
            sql = "Select * From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'  order by 類別,項目名稱 "
            lb_細項流水.Text = 編號(1)
            '讀取權利人資料
            Using con_權利人 As New SqlConnection(EGOUPLOADSqlConnStr)
                con_權利人.Open()
                Dim selstr As String = "select * from 委賣物件資料表_細項所有權人 where 物件編號 = '" & 編號(0) & "' and 細項流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "' order by 權利人流水號"
                Dim tmpcodevalue As String = ""
                Using cmd_權利人 As New SqlCommand(selstr, conn)
                    Dim dt As New DataTable
                    dt.Load(cmd_權利人.ExecuteReader())
                    If dt.Rows.Count > 0 Then
                        Dim htmlstr As String = "<table cellspacing=""0"" rules=""all"" border=""1"" id=""Table_right"" style=""width:100%;border-collapse:collapse;"">"
                        htmlstr += "<tr><td colspan=""7"" >權利人部分</td></tr>"
                        htmlstr += "<tr><th scope=""col"">&nbsp;</th><th scope=""col"">姓名</th><th scope=""col"">所有權部權利範圍</th><th scope=""col"">面積</th><th scope=""col"">出售權利種類</th><th scope=""col"">所有權種類</th><th scope=""col""></th></tr>"
                        For Each dr As DataRow In dt.Rows
                            tmpcodevalue += dr("TempCode").ToString() & ","
                            Dim txt_name As String = "<input type='text' id='tb_Textbox50_" & dr("權利人流水號") & "' style='width:50px;' value='" & dr("所有權人") & "' />"
                            Dim txtbox_right As String = " 權利範圍.<input type='text' id='tb_Numerator_1_" & dr("權利人流水號") & "' title='分母' style='width:50px;' value='" & dr("權利範圍_分母") & "' /> &nbsp;分之 &nbsp;<input type=text' id='tb_Denominator_1_" & dr("權利人流水號") & "' title='分子' style='width:50px;' onblur='calculatorarea(this.id);' value='" & dr("權利範圍_分子") & "' /></br>出售權利範圍.<input type='text'  id='tb_Numerator_2_" & dr("權利人流水號") & "' title='分母' style='width:50px;' value='" & dr("出售權利範圍_分母") & "' /> &nbsp;分之 &nbsp;<input type='text' id='tb_Denominator_2_" & dr("權利人流水號") & "' title='分子' style='width:50px;' onblur='calculatorarea2(this.id);' value='" & dr("出售權利範圍_分子") & "' />"
                            Dim lb_holder As String = "持有面積:<span id='tb_Label21_1_" & dr("權利人流水號") & "'>" & IIf(IsDBNull(dr("持有面積")), "", dr("持有面積") & "㎡") & "</span><br />出售面積:<span id='tb_Label21_2_" & dr("權利人流水號") & "'>" & IIf(IsDBNull(dr("出售面積")), "", dr("出售面積") & "㎡") & "</span>"
                            Dim select1 As String = "<select id='tb_DropDownList70_" & dr("權利人流水號") & "'><option  value='所有權' " & IIf(dr("權利總類") = "所有權", "selected='selected'", "") & ">所有權</option><option value='地上權' " & IIf(dr("權利總類") = "地上權", "selected='selected'", "") & ">地上權</option><option value='典權' " & IIf(dr("權利總類") = "典權", "selected='selected'", "") & ">典權</option><option value='使用權' " & IIf(dr("權利總類") = "使用權", "selected='selected'", "") & ">使用權</option></select>"
                            Dim bons As String = ""
                            'If Request.Cookies("webfly_empno").Value.ToString.ToUpper = "0AEU" Then
                            If myobj.D = "1" Then
                                bons = "<img src='image/tick_64.png' style='width:30px;height:30px' /><input id='Button12" & dr("權利人流水號") & "' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode" & dr("權利人流水號") & "'  style='font-size :3px ;color:Gray;line-height :6px;'  >" & dr("TempCode").ToString() & "</span>"
                            Else
                                bons = "<img src='image/tick_64.png' style='width:30px;height:30px' /><input type='hidden' id='Button12" & dr("權利人流水號") & "' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode" & dr("權利人流水號") & "'  style='font-size :3px ;color:Gray;line-height :6px;'  >" & dr("TempCode").ToString() & "</span>"
                            End If
                            'Else
                            '    bons = "<img src='image/tick_64.png' style='width:30px;height:30px' /><input id='Button12" & dr("權利人流水號") & "' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode" & dr("權利人流水號") & "'  style='font-size :3px ;color:Gray;line-height :6px;'  >" & dr("TempCode").ToString() & "</span>"
                            'End If
                            'bons = "<img src='image/tick_64.png' style='width:30px;height:30px' /><input id='Button12" & dr("權利人流水號") & "' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode" & dr("權利人流水號") & "'  style='font-size :3px ;color:Gray;line-height :6px;'  >" & dr("TempCode").ToString() & "</span>"
                            Dim select2 As String = " <select  id='tb_DropDownList71_" & dr("權利人流水號") & "'><option value='單獨所有' " & IIf(dr("所有權種類") = "單獨所有", "selected='selected'", "") & ">單獨所有</option><option  value='分別共有' " & IIf(dr("所有權種類") = "分別共有", "selected='selected'", "") & ">分別共有</option><option value='公同共有' " & IIf(dr("所有權種類") = "公同共有", "selected='selected'", "") & ">公同共有</option></select>"
                            htmlstr += "<tr align=""center""><td><span id=""tb_num" & dr("權利人流水號") & """>" & dr("權利人流水號") & "</span></td><td>" & txt_name & "</td><td align=left>" & txtbox_right & "</td><td>" & lb_holder & "</td><td>" & select1 & "</td><td>" & select2 + "</td><td>" + bons + "</td></tr>"
                        Next
                        hidtempcode.Value = tmpcodevalue
                        editTable.Text = htmlstr
                    Else
                        '加入空表單
                        Dim htmlstr As String = "<table cellspacing=""0"" rules=""all"" border=""1"" id=""Table_right"" style=""width:100%;border-collapse:collapse;"">"
                        htmlstr += "<tr><td colspan=""7"" >權利人部分</td></tr>"
                        htmlstr += "<tr><th scope=""col"">&nbsp;</th><th scope=""col"">姓名</th><th scope=""col"">權利範圍</th><th scope=""col"">面積</th><th scope=""col"">出售權利種類</th><th scope=""col"">所有權種類</th><th scope=""col""></th></tr>"
                        Dim txt_name As String = "<input type='text' id='tb_Textbox50_0' style='width:50px;'  />"
                        Dim txtbox_right As String = " 權利範圍.<input type='text' id='tb_Numerator_1_0' title='分母' style='width:50px;'  /> &nbsp;分之 &nbsp;<input type=text' id='tb_Denominator_1_0' title='分子' style='width:50px;' onblur='calculatorarea(this.id);'  /></br>出售權利範圍.<input type='text'  id='tb_Numerator_2_0' title='分母' style='width:50px;'  /> &nbsp;分之 &nbsp;<input type='text' id='tb_Denominator_2_0' title='分子' style='width:50px;' onblur='calculatorarea2(this.id);'  />"
                        Dim lb_holder As String = "持有面積:<span id='tb_Label21_1_0'></span><br />出售面積:<span id='tb_Label21_2_0'></span>"
                        Dim select1 As String = "<select id='tb_DropDownList70_0'><option  value='所有權'>所有權</option><option value='地上權' >地上權</option><option value='典權' >典權</option><option value='使用權' >使用權</option></select>"
                        Dim bons As String = ""
                        'If Request.Cookies("webfly_empno").Value.ToString.ToUpper = "0AEU" Then
                        If myobj.M = "1" And myobj.D = "1" Then
                            bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        ElseIf myobj.M = "1" And myobj.D = "0" Then
                            bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input type='hidden' id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        ElseIf myobj.M = "0" And myobj.D = "1" Then
                            bons = "<input type='hidden' id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        Else
                            bons = "<input type='hidden' id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input type='hidden' id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        End If
                        'Else
                        '    bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        'End If
                        'bons = "<input id='Button11_0' type='button' value='儲存' onclick='SaveRow(this)'  /><input id='Button120' type='button' value='刪除' onclick='DelRow(this)' /><br /><span id='tempcode0'  style='font-size :3px ;color:Gray;line-height :6px;'  ></span>"
                        Dim select2 As String = " <select  id='tb_DropDownList71_0'><option value='單獨所有' >單獨所有</option><option  value='分別共有' >分別共有</option><option value='公同共有' >公同共有</option></select>"
                        htmlstr += "<tr align=""center""><td><span id=""tb_num0"">1</span></td><td>" & txt_name & "</td><td align=left>" & txtbox_right & "</td><td>" & lb_holder & "</td><td>" & select1 & "</td><td>" & select2 + "</td><td>" + bons + "</td></tr>"
                        editTable.Text = htmlstr
                    End If
                End Using
            End Using
        End If

        ''20160616 add by nick
        '將OLD的參數還原到這邊 才不會影響到下面的進行
        If alseCallOLD = True Then
            sql = "Select *,(權利範圍1分子+'/'+權利範圍1分母) as 權利範圍,(權利範圍2分子+'/'+權利範圍2分母) as 權利範圍2 From 委賣物件資料表_面積細項 With(NoLock) Where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'  order by 類別,項目名稱 "
            cls = "OLD"
        End If
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "細項內容")
        Dim tb_細項內容 As DataTable = ds.Tables("細項內容")

        If cls = "OLD" Then
            Me.GridView1.DataSource = tb_細項內容
            Me.GridView1.DataBind()

            If Me.GridView1.Rows.Count = 0 Then
                Me.ImageButton2.Visible = False
            Else
                '以此判斷物件是否過期被LOCK住不能修改
                '20151007 Textbox253 移除改用 Textbox4 來判斷
                'If TextBox253.Enabled = True Then
                If TextBox4.Enabled = True Then
                    '面積細項刪除按鈕
                    If myobj.D = "1" Then
                        Me.ImageButton2.Visible = True
                    Else
                        Me.ImageButton2.Visible = False
                    End If
                End If
            End If
        Else
            '類別
            If Not IsDBNull(tb_細項內容.Rows(0)("類別")) Then
                For Each lis As ListItem In DropDownList2.Items
                    If lis.Text = tb_細項內容.Rows(0)("類別") Then
                        lis.Selected = True
                        Exit For
                    End If
                Next

                If Not IsDBNull(tb_細項內容.Rows(0)("DL_level2_selectindex")) Then
                    DDL_level2.SelectedIndex = tb_細項內容.Rows(0)("DL_level2_selectindex")
                Else
                    DDL_level2.SelectedIndex = 0
                End If

                If tb_細項內容.Rows(0)("類別") = "土地面積" Or tb_細項內容.Rows(0)("類別") = "主建物" Or tb_細項內容.Rows(0)("類別").ToString.IndexOf("車位") >= 0 Then
                    Label76.Visible = True '警示文字
                    Me.DropDownList2.Enabled = False
                Else
                    Label76.Visible = False '警示文字
                    Me.DropDownList2.Enabled = True
                End If
                DropDownList2_SelectedIndexChanged1(Nothing, Nothing)

                If tb_細項內容.Rows(0)("類別") = "主建物" And DDL_level2.SelectedIndex = 2 Then      ' DDL_level2：0 (層次)、1 (附屬建物)、2 (共有部分)
                    TextBox1.Visible = True
                    TextBox3.Visible = True
                    Label26.Visible = True
                    Label70.Text = "標示部權利範圍"

                    TextBox44.Visible = True
                    TextBox45.Visible = True
                    Label46.Visible = True
                    Label72.Text = "共有部分車位權利範圍"
                End If

                '項目名稱
                '項目名稱(tb_細項內容.Rows(0)("類別"))----20150828不使用
                If Not IsDBNull(tb_細項內容.Rows(0)("項目名稱")) Then
                    For Each lis As ListItem In DropDownList1.Items
                        If lis.Text = tb_細項內容.Rows(0)("項目名稱") Then
                            lis.Selected = True
                            Exit For
                        End If
                    Next
                End If
            End If
            '是否為公設
            If Not IsDBNull(tb_細項內容.Rows(0)("是否為公設")) Then
                If tb_細項內容.Rows(0)("是否為公設") = "Y" Then
                    CheckBox98.Checked = True
                End If
            End If

            '是否為車位
            If Not IsDBNull(tb_細項內容.Rows(0)("是否為車位")) Then
                If tb_細項內容.Rows(0)("是否為車位") = "Y" Then
                    CheckBox99.Checked = True
                End If
            End If

            '建號
            If Not IsDBNull(tb_細項內容.Rows(0)("建號")) Then
                Me.TextBox25.Text = tb_細項內容.Rows(0)("建號")
            End If

            '總面積平方公尺
            If Not IsDBNull(tb_細項內容.Rows(0)("總面積平方公尺")) Then
                Me.TextBox77.Text = tb_細項內容.Rows(0)("總面積平方公尺")
                If Trim(TextBox77.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox77, "F")
                End If
            End If

            '總面積坪
            If Not IsDBNull(tb_細項內容.Rows(0)("總面積坪")) Then
                Me.TextBox73.Text = tb_細項內容.Rows(0)("總面積坪")
                '1040522修正-新增自訂小數點位數判斷-坪
                If Trim(TextBox73.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox73, "T")
                End If
            End If

            '權利範圍1分母
            If Not IsDBNull(tb_細項內容.Rows(0)("權利範圍1分母")) Then
                Me.TextBox1.Text = tb_細項內容.Rows(0)("權利範圍1分母")
            End If

            '權利範圍1分子
            If Not IsDBNull(tb_細項內容.Rows(0)("權利範圍1分子")) Then
                Me.TextBox3.Text = tb_細項內容.Rows(0)("權利範圍1分子")
            End If

            '權利範圍2分母
            If Not IsDBNull(tb_細項內容.Rows(0)("權利範圍2分母")) Then
                Me.TextBox44.Text = tb_細項內容.Rows(0)("權利範圍2分母")
            End If

            '權利範圍2分子
            If Not IsDBNull(tb_細項內容.Rows(0)("權利範圍2分子")) Then
                Me.TextBox45.Text = tb_細項內容.Rows(0)("權利範圍2分子")
            End If

            '實際持有平方公尺
            If Not IsDBNull(tb_細項內容.Rows(0)("實際持有平方公尺")) Then
                Me.TextBox22.Text = tb_細項內容.Rows(0)("實際持有平方公尺")
                If Trim(TextBox22.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox22, "F")
                End If
            End If

            '實際持有坪
            If Not IsDBNull(tb_細項內容.Rows(0)("實際持有坪")) Then
                Me.TextBox24.Text = tb_細項內容.Rows(0)("實際持有坪")
                '1040522修正-新增自訂小數點位數判斷-坪
                If Trim(TextBox24.Text) <> "" Then
                    自訂坪數_小數點位數(TextBox24, "T")
                End If
            End If

            '增建用途
            If Not IsDBNull(tb_細項內容.Rows(0)("增建用途")) Then
                Me.TextBox254.Text = tb_細項內容.Rows(0)("增建用途")
            End If

            '增建完成日期
            If Not IsDBNull(tb_細項內容.Rows(0)("增建完成日期")) Then
                Date8.Text = tb_細項內容.Rows(0)("增建完成日期")
            End If

            '管制
            If Not IsDBNull(tb_細項內容.Rows(0)("管制")) Then
                If Trim(tb_細項內容.Rows(0)("管制")) = "" Then
                    DropDownList69.SelectedValue = "請選擇"
                Else
                    DropDownList69.SelectedValue = Trim(tb_細項內容.Rows(0)("管制"))
                End If
            Else
                DropDownList69.SelectedValue = "請選擇"
            End If

            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            '因電傳導入，會自動新增至 資料_使用分區 ，故改成抓取資料表資料
            Select_分區細項New(Trim(tb_細項內容.Rows(0)("使用分區").ToString))
            '    Else
            '    '使用分區1040616-V2版新增
            '    Select Case Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "住一", "住一之一", "住二", "住二之一", "住二之二", "住三", "住三之一", "住三之二", "住四", "住四之一", "住五", "特定住宅區(二)", "第一種住宅區", "第一之B種住宅區", "第七種住宅區", "住五之一"
            '            DropDownList65.SelectedValue = "住宅區"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "商一", "商二", "商三", "商四", "商五", "第二之一種商業區", "第三之一種商業區", "第四之一種商業區"
            '            DropDownList65.SelectedValue = "商業區"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "工二", "工三", "特種工業區", "甲種工業區", "乙種工業區", "零星工業區", "策略型工業區", "科技工業區(A區)", "科技工業區(B區)"
            '            DropDownList65.SelectedValue = "工業區"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "甲種建築用地", "乙種建築用地", "丙種建築用地", "丁種建築用地"
            '            DropDownList65.SelectedValue = "非都市計畫區"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "農牧用地"
            '            DropDownList65.SelectedValue = "河川區"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case "上下水道", "公園", "市場", "民用航空站", "兒童遊樂場", "其他", "河道港埠用地", "社教機關", "停車場所", "捷運系統用地", "郵政", "道路", "電信", "綠地", "廣場", "學校", "醫療衛生機構", "變電所", "體育場所"
            '            DropDownList65.SelectedValue = "公共設施用地"
            '            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            '            DropDownList66.SelectedValue = Trim(tb_細項內容.Rows(0)("使用分區").ToString)
            '        Case Else
            '            For i  = 0 To DropDownList65.Items.Count - 1
            '                If DropDownList65.Items(i).Value = Trim(tb_細項內容.Rows(0)("使用分區").ToString) Then
            '                    DropDownList65.SelectedIndex = i
            '                    Exit For
            '                End If
            '            Next
            '    End Select
            'End If

            '權利人
            'If Not IsDBNull(tb_細項內容.Rows(0)("所有權人")) Then
            '    TextBox255.Text = tb_細項內容.Rows(0)("所有權人")
            'End If
            '20160125 權利人改一對多 by nick

            '法定建蔽率
            If Not IsDBNull(tb_細項內容.Rows(0)("法定建蔽率")) Then
                TextBox256.Text = tb_細項內容.Rows(0)("法定建蔽率")
            End If

            '法定容積率
            If Not IsDBNull(tb_細項內容.Rows(0)("法定容積率")) Then
                TextBox257.Text = tb_細項內容.Rows(0)("法定容積率")
            End If

            '序號
            Me.Label4.Text = cls

            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then
                '判斷為新增還是修改ㄉ狀態
                If Me.Label4.Text = "0" Then
                    '面積細項新增按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton5.Visible = True
                    Else
                        Me.ImageButton5.Visible = False
                    End If
                    Me.ImageButton3.Visible = False
                Else
                    Me.ImageButton5.Visible = False
                    '面積細項修改按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton3.Visible = True
                    Else
                        Me.ImageButton3.Visible = False
                    End If
                End If

                If tb_細項內容.Rows(0)("類別") = "土地面積" Then
                    'Me.ImageButton17.Visible = True
                Else
                    'Me.ImageButton17.Visible = False
                End If
            End If
        End If

        conn.Close()
        conn.Dispose()

        If Me.GridView1.Rows.Count = 0 Then
            Me.ImageButton2.Visible = False
        Else
            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then
                '面積細項刪除按鈕
                If myobj.D = "1" Then
                    Me.ImageButton2.Visible = True
                Else
                    Me.ImageButton2.Visible = False
                End If
            End If
        End If

        '項目名稱
        If tb_細項內容.Rows.Count > 0 Then
            If Not IsDBNull(tb_細項內容.Rows(0)("項目名稱")) And tb_細項內容.Rows(0)("類別") = "主建物" Then
                DDL_level2_SelectedIndexChanged(Nothing, Nothing)
                For Each lis As ListItem In DropDownList1.Items
                    If lis.Text = tb_細項內容.Rows(0)("項目名稱") Then
                        lis.Selected = True
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub


    '主建物
    Protected Sub TextBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) Then
            TextBox6.Text = square(TextBox5.Text)
        End If

        ScriptManager1.SetFocus(TextBox7)

        計算總坪數()
    End Sub


    '附屬物
    Protected Sub TextBox7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox7.Text) Then
            TextBox8.Text = square(TextBox7.Text)
        End If

        ScriptManager1.SetFocus(TextBox9)

        Call 計算總坪數()

    End Sub


    '公共設施
    Protected Sub TextBox9_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox9.Text) Then
            TextBox10.Text = square(TextBox9.Text)
        End If

        ScriptManager1.SetFocus(TextBox19)


        Call 計算總坪數()
    End Sub

    '地下室
    Protected Sub TextBox19_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox19.Text) Then
            TextBox20.Text = square(TextBox19.Text)
        End If

        ScriptManager1.SetFocus(TextBox21)

        Call 計算總坪數()
    End Sub

    '公設車位
    Protected Sub TextBox21_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox21.Text) Then
            TextBox23.Text = square(TextBox21.Text)
        End If

        ScriptManager1.SetFocus(TextBox26)

        Call 計算總坪數()


    End Sub

    '產權車位
    Protected Sub TextBox26_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged
        If IsNumeric(TextBox26.Text) Then
            TextBox27.Text = square(TextBox26.Text)
        End If

        ScriptManager1.SetFocus(TextBox30)

        Call 計算總坪數()
    End Sub

    '土地面積
    Protected Sub TextBox30_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox30.Text) Then
            TextBox31.Text = square(TextBox30.Text)
        End If

        ScriptManager1.SetFocus(TextBox28)

        Call 計算總坪數()
    End Sub

    '總坪數
    Protected Sub TextBox28_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox28.Text) Then
            TextBox29.Text = square(TextBox28.Text)
        End If

        ScriptManager1.SetFocus(TextBox26)

        Call 計算總坪數()
    End Sub

    '庭院坪數
    Protected Sub TextBox32_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox32.Text) Then
            TextBox33.Text = square(TextBox32.Text)
        End If

        ScriptManager1.SetFocus(TextBox34)

        Call 計算總坪數()
    End Sub

    '增建坪數
    Protected Sub TextBox34_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox34.TextChanged
        If IsNumeric(TextBox34.Text) Then
            TextBox35.Text = square(TextBox34.Text)
        End If

        Call 計算總坪數()
    End Sub

    '基地面積
    'Protected Sub TextBox231_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox231.TextChanged
    'If IsNumeric(TextBox231.Text) Then
    '    TextBox230.Text = square(TextBox231.Text)
    'End If
    'End Sub

    Sub 計算總坪數()

        Dim sum1 As Double = 0
        Dim sum2 As Double = 0

        If IsNumeric(TextBox5.Text) Then sum1 += TextBox5.Text '主建物
        If IsNumeric(TextBox7.Text) Then sum1 += TextBox7.Text '附屬物
        If IsNumeric(TextBox9.Text) Then sum1 += TextBox9.Text '公共設施
        If IsNumeric(TextBox19.Text) Then sum1 += TextBox19.Text '地下室
        If IsNumeric(TextBox21.Text) Then sum1 += TextBox21.Text '公設車位

        If IsNumeric(TextBox6.Text) Then sum2 += TextBox6.Text
        If IsNumeric(TextBox8.Text) Then sum2 += TextBox8.Text
        If IsNumeric(TextBox10.Text) Then sum2 += TextBox10.Text
        If IsNumeric(TextBox20.Text) Then sum2 += TextBox20.Text
        If IsNumeric(TextBox23.Text) Then sum2 += TextBox23.Text

        TextBox28.Text = sum1
        If Trim(TextBox28.Text) <> "" Then
            自訂坪數_小數點位數(TextBox28, "F")
        End If

        '總坪數部分位數(1040522修正-新增自訂小數點位數判斷-坪)
        TextBox29.Text = sum2 'Round(TextBox28.Text * 0.3025, 4)
        '1040522修正-新增自訂小數點位數判斷-坪
        If Trim(TextBox29.Text) <> "" Then
            自訂坪數_小數點位數(TextBox29, "T")
        End If

        'AJAX-總坪數
        UpdatePanel19.Update()

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            '物件編號LABEL
            Dim lbl物件編號 As Label = e.Row.FindControl("Label35")

            '流水號LABEL
            Dim lbl流水號 As Label = e.Row.FindControl("Label36")

            '類別
            Dim lbl類別 As Label = e.Row.FindControl("Label37")

            '子類別
            Dim lbl子類別 As Label = e.Row.FindControl("Label24")

            '項目
            Dim lbl項目 As Label = e.Row.FindControl("Label38")

            '建號為車位
            Dim lbl建號車位 As Label = e.Row.FindControl("Label99")

            If lbl子類別.Text = "0" Then
                lbl子類別.Text = "層次"
            End If
            If lbl子類別.Text = "1" Then
                lbl子類別.Text = "附屬建物"
            End If
            If lbl子類別.Text = "2" Then
                lbl子類別.Text = "共有部分"
            End If
            '土地面積不要顯示層次
            If lbl類別.Text = "土地面積" And lbl子類別.Text = "層次" Then
                lbl子類別.Text = ""
            End If
            '編輯button
            Dim btn編輯 As Button = e.Row.FindControl("Button5")
            btn編輯.CommandArgument = lbl物件編號.Text & "," & lbl流水號.Text
            'If Request.Cookies("webfly_empno").Value.ToString.ToUpper = "0AEU" Then
            If myobj.M = "1" Then
                btn編輯.Visible = True
            Else
                btn編輯.Visible = False
            End If
            'End If

            '調查button
            Dim btn調查 As Button = e.Row.FindControl("Button6")
            btn調查.CommandArgument = lbl物件編號.Text & "," & lbl流水號.Text & "," & lbl類別.Text & "," & lbl項目.Text & "," & lbl建號車位.Text

            If lbl類別.Text = "土地面積" Or lbl類別.Text = "車位面積(公設內)" Or lbl類別.Text = "車位面積(產權獨立)" Or (lbl類別.Text = "主建物" And (lbl子類別.Text = "層次" Or lbl子類別.Text = "")) Or lbl項目.Text.IndexOf("車位") >= 0 Or lbl項目.Text.IndexOf("停車空間") >= 0 And (lbl子類別.Text <> "附屬建物" Or lbl子類別.Text <> "共有部分") Then
                btn調查.Visible = True
            Else
                btn調查.Visible = False
            End If

            '權利範圍1LABEL
            Dim lbl權利範圍1 As Label = e.Row.FindControl("Label42")
            '權利範圍2LABEL
            Dim lbl權利範圍2 As Label = e.Row.FindControl("Label20")

            ''1040616-V2修正
            If lbl權利範圍2.Text = "1/1" Or lbl權利範圍2.Text = "" Or lbl項目.Text = "車位" Or lbl項目.Text = "停車空間" Then
                lbl權利範圍2.Visible = False
                'Else
                '    lbl權利範圍2.Visible = True

                '    lbl權利範圍1.Text = "A." & lbl權利範圍1.Text
                '    lbl權利範圍2.Text = "B." & lbl權利範圍2.Text
            End If

            '2016.06.26 by Finch 
            If lbl類別.Text = "主建物" And lbl子類別.Text = "共有部分" And (lbl項目.Text = "車位" Or lbl項目.Text = "停車空間") Then
                lbl權利範圍1.Visible = False
                lbl權利範圍2.Visible = True
            Else
                lbl權利範圍1.Visible = True
                lbl權利範圍2.Visible = False
            End If

            '1050317 去細項所有權人撈權利範圍顯示
            Dim totalholderArea As Decimal = 0
            Dim totalSellArea As Decimal = 0
            Using conn_權利範圍 As New SqlConnection(EGOUPLOADSqlConnStr)
                conn_權利範圍.Open()
                'Dim selstr As String = "select 所有權人, 權利範圍_分子, 權利範圍_分母, 出售權利範圍_分子, 出售權利範圍_分母, 持有面積, 出售面積 FROM 委賣物件資料表_細項所有權人 where 物件編號 = '" & lbl物件編號.Text & "' and 細項流水號 = '" & lbl流水號.Text & "' AND 店代號 = '" & store.SelectedValue & "'"
                Dim selstr As String = "select 所有權人, 權利範圍_分子, 權利範圍_分母, 出售權利範圍_分子, 出售權利範圍_分母, isnull(持有面積,0) as 持有面積, isnull(出售面積,0) as 出售面積 FROM 委賣物件資料表_細項所有權人 where 物件編號 = '" & lbl物件編號.Text & "' and 細項流水號 = '" & lbl流水號.Text & "' AND 店代號 = '" & Me.Label12.Text & "' order by 權利人流水號 "
                'Response.Write(store.SelectedValue)
                'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                '    Response.Write(selstr)
                '    'eip_usual.show(selstr)
                'End If

                Using cmd As New SqlCommand(selstr, conn_權利範圍)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader())
                    If dt.Rows.Count > 0 Then
                        lbl權利範圍1.Text = ""
                        For Each dr As DataRow In dt.Rows
                            totalholderArea += dr("持有面積")
                            totalSellArea += dr("出售面積")
                            lbl權利範圍1.Text += dr("所有權人") & ":" & dr("權利範圍_分子") & "/" & dr("權利範圍_分母") & "<br/>"
                        Next
                    End If
                End Using
            End Using

            '總面積平方公尺LABEL
            Dim lbl總面積平方公尺 As Label = e.Row.FindControl("Label40")
            '1040522新增-------------------------------------------------------------------------------
            '去掉尾巴的0
            If Left(Right(lbl總面積平方公尺.Text, 5), 1) = "." And Right(lbl總面積平方公尺.Text, 4) = "0000" Then
                lbl總面積平方公尺.Text = Int(lbl總面積平方公尺.Text)
            ElseIf Left(Right(lbl總面積平方公尺.Text, 5), 1) = "." And Right(lbl總面積平方公尺.Text, 3) = "000" Then
                lbl總面積平方公尺.Text = Left(lbl總面積平方公尺.Text, lbl總面積平方公尺.Text.Length - 3)
            ElseIf Left(Right(lbl總面積平方公尺.Text, 5), 1) = "." And Right(lbl總面積平方公尺.Text, 2) = "00" Then
                lbl總面積平方公尺.Text = Left(lbl總面積平方公尺.Text, lbl總面積平方公尺.Text.Length - 2)
            ElseIf Left(Right(lbl總面積平方公尺.Text, 5), 1) = "." And Right(lbl總面積平方公尺.Text, 1) = "0" Then
                lbl總面積平方公尺.Text = Left(lbl總面積平方公尺.Text, lbl總面積平方公尺.Text.Length - 1)
            Else
                lbl總面積平方公尺.Text = lbl總面積平方公尺.Text
            End If
            '--------------------------------------------------------------------------------------------

            '總面積坪LABEL
            Dim lbl總面積坪 As Label = e.Row.FindControl("Label41")
            '1040522新增-------------------------------------------------------------------------------
            'step1-先取得自訂小數點位數後的數字
            'If lbl總面積坪.Text = "" Then
            '    lbl總面積坪.Text = "0"
            'Else
            lbl總面積坪.Text = Round(CType(lbl總面積坪.Text, Double), CType(Float.Text, Integer), MidpointRounding.AwayFromZero)
            'End If

            'step2-去掉尾巴的0
            If Left(Right(lbl總面積坪.Text, 5), 1) = "." And Right(lbl總面積坪.Text, 4) = "0000" Then
                lbl總面積坪.Text = Int(lbl總面積坪.Text)
            ElseIf Left(Right(lbl總面積坪.Text, 5), 1) = "." And Right(lbl總面積坪.Text, 3) = "000" Then
                lbl總面積坪.Text = Left(lbl總面積坪.Text, lbl總面積坪.Text.Length - 3)
            ElseIf Left(Right(lbl總面積坪.Text, 5), 1) = "." And Right(lbl總面積坪.Text, 2) = "00" Then
                lbl總面積坪.Text = Left(lbl總面積坪.Text, lbl總面積坪.Text.Length - 2)
            ElseIf Left(Right(lbl總面積坪.Text, 5), 1) = "." And Right(lbl總面積坪.Text, 1) = "0" Then
                lbl總面積坪.Text = Left(lbl總面積坪.Text, lbl總面積坪.Text.Length - 1)
            Else
                lbl總面積坪.Text = lbl總面積坪.Text
            End If
            '--------------------------------------------------------------------------------------------


            '實際持有平方公尺LABEL
            Dim lbl實際持有平方公尺 As Label = e.Row.FindControl("Label43")

            '20160322 將所有權利人的相加
            If totalholderArea > 0 Then
                lbl實際持有平方公尺.Text = totalholderArea.ToString
            End If


            '實際出售平方公尺LABEL
            Dim lbl實際出售平方公尺 As Label = e.Row.FindControl("Label88")

            '20160614 將所有權利人的相加
            If totalSellArea > 0 Then
                lbl實際出售平方公尺.Text = totalSellArea.ToString
            End If


            '1040522新增-------------------------------------------------------------------------------
            '去掉尾巴的0
            'If lbl實際持有平方公尺.Text = "" Then
            '    lbl實際持有平方公尺.Text = "0"
            'Else
            If Left(Right(lbl實際持有平方公尺.Text, 5), 1) = "." And Right(lbl實際持有平方公尺.Text, 4) = "0000" Then
                lbl實際持有平方公尺.Text = Int(lbl實際持有平方公尺.Text)
            ElseIf Left(Right(lbl實際持有平方公尺.Text, 5), 1) = "." And Right(lbl實際持有平方公尺.Text, 3) = "000" Then
                lbl實際持有平方公尺.Text = Left(lbl實際持有平方公尺.Text, lbl實際持有平方公尺.Text.Length - 3)
            ElseIf Left(Right(lbl實際持有平方公尺.Text, 5), 1) = "." And Right(lbl實際持有平方公尺.Text, 2) = "00" Then
                lbl實際持有平方公尺.Text = Left(lbl實際持有平方公尺.Text, lbl實際持有平方公尺.Text.Length - 2)
            ElseIf Left(Right(lbl實際持有平方公尺.Text, 5), 1) = "." And Right(lbl實際持有平方公尺.Text, 1) = "0" Then
                lbl實際持有平方公尺.Text = Left(lbl實際持有平方公尺.Text, lbl實際持有平方公尺.Text.Length - 1)
            Else
                lbl實際持有平方公尺.Text = lbl實際持有平方公尺.Text
            End If
            'End If

            If Left(Right(lbl實際出售平方公尺.Text, 5), 1) = "." And Right(lbl實際出售平方公尺.Text, 4) = "0000" Then
                lbl實際出售平方公尺.Text = Int(lbl實際出售平方公尺.Text)
            ElseIf Left(Right(lbl實際出售平方公尺.Text, 5), 1) = "." And Right(lbl實際出售平方公尺.Text, 3) = "000" Then
                lbl實際出售平方公尺.Text = Left(lbl實際出售平方公尺.Text, lbl實際出售平方公尺.Text.Length - 3)
            ElseIf Left(Right(lbl實際出售平方公尺.Text, 5), 1) = "." And Right(lbl實際出售平方公尺.Text, 2) = "00" Then
                lbl實際出售平方公尺.Text = Left(lbl實際出售平方公尺.Text, lbl實際出售平方公尺.Text.Length - 2)
            ElseIf Left(Right(lbl實際出售平方公尺.Text, 5), 1) = "." And Right(lbl實際出售平方公尺.Text, 1) = "0" Then
                lbl實際出售平方公尺.Text = Left(lbl實際出售平方公尺.Text, lbl實際出售平方公尺.Text.Length - 1)
            Else
                lbl實際出售平方公尺.Text = lbl實際出售平方公尺.Text
            End If
            '--------------------------------------------------------------------------------------------


            '實際持有坪LABEL
            Dim lbl實際持有坪 As Label = e.Row.FindControl("Label44")
            If totalholderArea > 0 Then
                lbl實際持有坪.Text = Format(totalholderArea * 0.3025, "0.000000")
            End If

            '實際出售坪LABEL
            Dim lbl實際出售坪 As Label = e.Row.FindControl("Label89")
            If totalSellArea > 0 Then
                'If totalSellArea = 0.000006 Then
                '    Response.Write(format(0.3025 * totalSellArea, "0.000000"))
                '    Exit Sub
                'Else
                lbl實際出售坪.Text = Format(0.3025 * totalSellArea, "0.000000")
                'End If
            End If

            '1040522新增-------------------------------------------------------------------------------
            'step1-先取得自訂小數點位數後的數字
            'If lbl實際持有坪.Text = "" Then
            '    lbl實際持有坪.Text = "0"
            'Else
            'lbl實際持有坪.Text = Round(CType(lbl實際持有坪.Text, Double), CType(Float.Text, Integer), MidpointRounding.AwayFromZero)
            'End If


            'step2-去掉尾巴的0
            'If Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 4) = "0000" Then
            '    lbl實際持有坪.Text = Int(lbl實際持有坪.Text)
            'ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 3) = "000" Then
            '    lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 3)
            'ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 2) = "00" Then
            '    lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 2)
            'ElseIf Left(Right(lbl實際持有坪.Text, 5), 1) = "." And Right(lbl實際持有坪.Text, 1) = "0" Then
            '    lbl實際持有坪.Text = Left(lbl實際持有坪.Text, lbl實際持有坪.Text.Length - 1)
            'Else
            '    lbl實際持有坪.Text = lbl實際持有坪.Text
            'End If

            'If Left(Right(lbl實際出售坪.Text, 5), 1) = "." And Right(lbl實際出售坪.Text, 4) = "0000" Then
            '    lbl實際出售坪.Text = Int(lbl實際出售坪.Text)
            'ElseIf Left(Right(lbl實際出售坪.Text, 5), 1) = "." And Right(lbl實際出售坪.Text, 3) = "000" Then
            '    lbl實際出售坪.Text = Left(lbl實際出售坪.Text, lbl實際出售坪.Text.Length - 3)
            'ElseIf Left(Right(lbl實際出售坪.Text, 5), 1) = "." And Right(lbl實際出售坪.Text, 2) = "00" Then
            '    lbl實際出售坪.Text = Left(lbl實際出售坪.Text, lbl實際出售坪.Text.Length - 2)
            'ElseIf Left(Right(lbl實際出售坪.Text, 5), 1) = "." And Right(lbl實際出售坪.Text, 1) = "0" Then
            '    lbl實際出售坪.Text = Left(lbl實際出售坪.Text, lbl實際出售坪.Text.Length - 1)
            'Else
            '    lbl實際出售坪.Text = lbl實際出售坪.Text
            'End If
            '--------------------------------------------------------------------------------------------

            If (lbl類別.Text = "車位面積(公設內)" Or (lbl類別.Text = "主建物" And lbl子類別.Text = "共有部分" And (lbl項目.Text = "車位" Or lbl項目.Text = "停車空間"))) And lbl總面積平方公尺.Text = "0" Then
                Checkbox27.Checked = True
                lbl總面積坪.Text = "--"
                lbl總面積平方公尺.Text = "--"
                lbl實際持有平方公尺.Text = "--"
                lbl實際持有坪.Text = "--"
            End If

            Dim LandType As Label = CType(e.Row.FindControl("Label37"), Label)
            If LandType IsNot Nothing AndAlso LandType.Text <> "土地面積" Then
                ' 找到按鈕並隱藏
                Dim btnLandTaxCount As Button = CType(e.Row.FindControl("btnlandtaxcount"), Button)
                If btnLandTaxCount IsNot Nothing Then
                    btnLandTaxCount.Visible = False
                End If
            End If

        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "edits" Then

            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
            'eip_usual.Show(e.CommandArgument)
            '因有發現到修改存檔時 hidtempcode.Value原值還在，導致所有權部存檔錯亂，按下編輯時，先行清空
            hidtempcode.Value = ""
            'Exit Sub
            'End If

            Load_Data(e.CommandArgument)

            '判斷物件維護權限
            myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")

            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then


                '判斷為新增還是修改ㄉ狀態
                If Me.Label4.Text = "0" Then
                    '面積細項新增按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton5.Visible = True
                    Else
                        Me.ImageButton5.Visible = False
                    End If
                    Me.ImageButton3.Visible = False
                Else
                    Me.ImageButton5.Visible = False
                    '面積細項修改按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton3.Visible = True
                    Else
                        Me.ImageButton3.Visible = False
                    End If

                End If

            End If

        ElseIf e.CommandName = "check" Then
            check_產調(e.CommandArgument)
        ElseIf e.CommandName = "landtaxcount" Then

            Dim rowIndex = e.CommandArgument
            Dim row As GridViewRow = GridView1.Rows(rowIndex)

            Dim nscript As String
            Dim href As String = ""
            Dim ntitle As String = ""
            Dim 物件編號 As String = Request("oid")
            Dim placeNo As Label = CType(row.FindControl("Label39"), Label)
            Dim landArea As Label = CType(row.FindControl("Label40"), Label)
            Dim setScope As Label = CType(row.FindControl("Label42"), Label)
            Dim input = setScope.Text

            'Dim colonIndex As Integer = input.IndexOf(":"c)
            'Dim frontValue As String
            'Dim backValue As String
            '' 如果找到 ":"，則提取後面的值
            'If colonIndex <> -1 Then
            '    ' 取得 ":" 之後的部分
            '    Dim afterColon As String = input.Substring(colonIndex + 1) ' 取得 "1/1"
            '    afterColon = afterColon.Replace("<br/>", "")
            '    ' 取得 "/" 的索引
            '    Dim slashIndex As Integer = afterColon.IndexOf("/"c)

            '    ' 取得前面的1和後面的1
            '    frontValue = afterColon.Substring(0, slashIndex) ' 取得 "1"
            '    backValue = afterColon.Substring(slashIndex + 1) ' 取得 "1"

            'End If

            Dim slashIndex As Integer = input.IndexOf("/"c)
            Dim frontValue As String = String.Empty
            Dim backValue As String = String.Empty

            If slashIndex <> -1 Then
                ' 取得 "/" 前面的部分，並只保留數字
                frontValue = System.Text.RegularExpressions.Regex.Replace(input.Substring(0, slashIndex), "[^\d.]", "").Trim()

                ' 取得 "/" 後面的部分，並只保留數字
                backValue = System.Text.RegularExpressions.Regex.Replace(input.Substring(slashIndex + 1), "[^\d]", "").Trim()
            End If

            'href = "../TOP_tools/tool_landcount.aspx?state=update&sid=" & Request("sid") & "&oid=" & 物件編號 & "&src=" & Request("src")
            '本機
            'href = "https://localhost:44386/TaxCount/LandTaxCount.aspx?oid=" & 物件編號 & "&sid=" & Request("sid") & "&placeNo=" & placeNo.Text & "&landArea=" & landArea.Text & "&setScope_2=" & frontValue & " &setScope_1=" & backValue & "&src= " & Request("src")
            '測試
            'href = "https://superwebnew6.etwarm.com.tw/LandTaxCount/TaxCount/LandTaxCount.aspx?oid=" & 物件編號 & "&sid=" & Request("sid") & "&placeNo=" & placeNo.Text & "&landArea=" & landArea.Text & "&setScope_2=" & frontValue & " &setScope_1=" & backValue & "&src= " & Request("src")
            '正式
            href = "https://superwebnew.etwarm.com.tw/LandTaxCount/TaxCount/LandTaxCount.aspx?oid=" & 物件編號 & "&sid=" & Request("sid") & "&placeNo=" & placeNo.Text & "&landArea=" & landArea.Text & "&setScope_2=" & frontValue & " &setScope_1=" & backValue & "&src= " & Request("src")

            href = href.Replace(" ", "")
            'ntitle = "土增稅計算"

            nscript = "window.open('"
            nscript += href
            nscript += " '"
            nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

            Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)


        End If


    End Sub

    Sub check_產調(ByVal cls As String)

        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""

        Dim 編號 As Array = Split(cls, ",")

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Response.Write(cls)
        '    'Exit Sub
        'End If

        '組合物件編號
        '物件編號-第1碼
        Dim 物件編號 As String = ""

        If Request("oid") = "" Then '新增時



            If ddl契約類別.SelectedValue = "專任" Then
                物件編號 = "1"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                物件編號 = "6"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                物件編號 = "7"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                物件編號 = "5"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                物件編號 = "9"
            End If

            '物件編號-第2-5碼(店代號)+第6-13碼(表單編號)
            If store.SelectedValue = "請選擇" Then
                物件編號 &= Mid(Request.Cookies("store_id").Value, 2) & TextBox2.Text.Trim
            Else
                物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            End If

        Else '修改複製
            物件編號 = Request("oid")
        End If

        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        '    'For i As Integer = 0 To 編號.Length - 1
        '    '    eip_usual.Show(編號(i))
        '    'Next
        'End If


        If 編號(2) = "土地面積" And 編號(3).ToString.IndexOf("車位") < 0 Then
            If DropDownList3.SelectedValue = "土地" Then
                href = "Place_Data.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue

                ntitle = "產權調查_土地"
                nscript = "window.open('"
                nscript += href
                nscript += " '"
                nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"
            Else

                'href = "Land_Data.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
                'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                href = "Land_DataV2.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
                'End If
                ntitle = "產權調查_基地"
                nscript = "window.open('"
                nscript += href
                nscript += " '"
                nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

            End If

        ElseIf 編號(2) = "車位面積(產權獨立)" Or 編號(2) = "車位面積(公設內)" Or 編號(3).ToString.IndexOf("車位") >= 0 Or 編號(3).ToString = "停車空間" Then
            href = "Car_DataV2.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue

            ntitle = "產權調查_車位"
            nscript = "window.open('"
            nscript += href
            nscript += " '"
            nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"
        ElseIf 編號(2) = "主建物" Then
            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            If 編號(4) = "Y" Then
                href = "Car_DataV2.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue

                ntitle = "產權調查_車位"
                nscript = "window.open('"
                nscript += href
                nscript += " '"
                nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"
            Else
                'href = "Build_Data.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
                'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                href = "Build_DataV2.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
                'End If
                ntitle = "產權調查_建物"
                nscript = "window.open('"
                nscript += href
                nscript += " '"
                nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"
            End If
            'Else
            '    'href = "Build_Data.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
            '    'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
            '    href = "Build_DataV2.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 編號(0) & "&Num=" & 編號(1) & "&uoid=" & 物件編號 & "&usid=" & store.SelectedValue
            '    'End If
            '    ntitle = "產權調查_建物"
            '    nscript = "window.open('"
            '    nscript += href
            '    nscript += " '"
            '    nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"
            'End If
        End If

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub
    '細項更新按鈕 事件
    Protected Sub ImageButton3_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton3.Click
        If Me.DropDownList1.SelectedIndex = 0 Or Me.DropDownList2.SelectedIndex = 0 Then
            Dim script As String = ""
            script += "alert('請選擇類別名稱.項目名稱!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        Else
            updt()
            '讀取資料
            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            Load_Data("OLD")
            'End If

            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            'trans = "False"
            'If ImageButton1.visible = True Then
            '    ImageButton1_Click(sender, e)
            'Else
            '    ImageButton12_Click(sender, e)
            'End If
            'If trans = "False" Then
            '    Exit Sub
            'End If
            'End If
        End If
    End Sub

    '修改
    Sub updt()
        '更新委賣物件資料表_細項所有權人 的流水號
        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        '    'eip_usual.Show("before:" & hidtempcode.Value)
        '    'Exit Sub
        'End If

        If hidtempcode.Value.Length > 5 Then
            If lb_細項流水.Text.Length > 0 Then
                Using con_細項所有權人 As New SqlConnection(EGOUPLOADSqlConnStr)
                    con_細項所有權人.Open()
                    hidtempcode.Value = hidtempcode.Value.Replace(",,", ",")
                    If Left(hidtempcode.Value, 1) = "," Then
                        hidtempcode.Value = hidtempcode.Value.Substring(1, hidtempcode.Value.Length - 1)
                    End If
                    If Right(hidtempcode.Value, 1) = "," Then
                        hidtempcode.Value = hidtempcode.Value.Substring(0, hidtempcode.Value.Length - 1)
                    End If

                    Dim updatstr As String = "update 委賣物件資料表_細項所有權人 set 細項流水號 = '" & lb_細項流水.Text & "' where TempCode in ('" & hidtempcode.Value.Replace(",", "','") & "')" ' and 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"

                    'If Request.Cookies("webfly_empno").Value.ToUpper = "00SF" Then
                    '    'eip_usual.Show("updatstr:" & updatstr)
                    'End If

                    Using cmd_細項所有權人 As New SqlCommand(updatstr, con_細項所有權人)
                        Try


                            cmd_細項所有權人.ExecuteNonQuery()
                            hidtempcode.Value = ""

                            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                            '    'eip_usual.Show("after:" & hidtempcode.Value)
                            'End If
                        Catch ex As Exception
                            Response.Write(updatstr)
                            Response.End()
                        End Try

                    End Using
                    '將權利範圍內容回寫回面積細項

                End Using
            End If
        End If

        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        '    eip_usual.Show("after:" & hidtempcode.Value)
        '    hidtempcode.Value = ""
        'End If

        '修改面積細項
        update_面積細項()


        '重整GRIDVIEW
        'Load_Data("OLD")
        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Load_Data("OLD")
        'Else
        Load_Data(Request.QueryString("oid") & "," & lb_細項流水.Text, True)
        'End If
        '清空所有細項
        clear()

        '計算面積
        Total()

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        寫入坪數()
        'End If

        'If Request.Cookies("webfly_empno").Value = "00P" Then
        '    Response.Write(Request.QueryString("oid") & "," & lb_細項流水.Text)
        '    'eip_usual.Show(Request.QueryString("oid") & "," & lb_細項流水.Text)
        '    Exit Sub
        'Else
        Dim script As String = ""
        script += "alert('修改成功1!!');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        hidtempcode.Value = ""
        'End If

        'End If

        '判斷物件維護權限
        myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")

        '判斷為新增還是修改ㄉ狀態
        If Me.Label4.Text = "0" Then
            '面積細項新增按鈕
            If myobj.AC = "1" Then
                Me.ImageButton5.Visible = True
            Else
                Me.ImageButton5.Visible = False
            End If
            Me.ImageButton3.Visible = False

            '新增的限制解除
            DropDownList2.Enabled = True
            Label76.Visible = False
        Else
            Me.ImageButton5.Visible = False
            '面積細項修改按鈕
            If myobj.AC = "1" Then
                Me.ImageButton3.Visible = True
            Else
                Me.ImageButton3.Visible = False
            End If

        End If
    End Sub

    '修改面積細項
    Sub update_面積細項()


        Dim 編號 As Array = Split(Me.Label4.Text, ",")

        '避免持份未填值.導致出錯-------------
        If Trim(Me.TextBox1.Text) = "" Then
            Me.TextBox1.Text = "1"
        End If
        If Trim(Me.TextBox3.Text) = "" Then
            Me.TextBox3.Text = "1"
        End If
        '權利範圍2
        If Trim(Me.TextBox44.Text) = "" Then
            Me.TextBox44.Text = "1"
        End If
        If Trim(Me.TextBox45.Text) = "" Then
            Me.TextBox45.Text = "1"
        End If
        If Trim(Me.TextBox22.Text) = "" Then
            Me.TextBox22.Text = "0"
        End If
        If Trim(Me.TextBox24.Text) = "" Then
            Me.TextBox24.Text = "0"
        End If
        If Trim(Me.TextBox77.Text) = "" Then
            Me.TextBox77.Text = "0"
        End If
        If Trim(Me.TextBox73.Text) = "" Then
            Me.TextBox73.Text = "0"
        End If
        '----------------------------------

        Dim 項目 As String = ""

        If Me.DropDownList1.SelectedItem.Text = "其他" Then
            If Trim(Me.TextBox11.Text) = "" Then
                項目 = "其他"
            Else
                項目 = Trim(Me.TextBox11.Text)
                chk_項目()
            End If
        Else
            If DropDownList1.SelectedItem.Text = "選擇項目名稱" Then
                項目 = "其他"
            Else
                項目 = Trim(Me.DropDownList1.SelectedItem.Text)
            End If
        End If

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        sql = "Update 委賣物件資料表_面積細項   "
        sql &= " Set "
        sql &= " 類別 = '" & Me.DropDownList2.SelectedItem.Text & "' ,"
        sql &= " 項目名稱 = '" & 項目 & "' ,"
        sql &= " 建號 = '" & TextBox25.Text & "' ,"
        sql &= " 總面積平方公尺 = " & TextBox77.Text & " ,"
        sql &= " 總面積坪 = " & TextBox73.Text & " ,"
        sql &= " 權利範圍1分母 = " & TextBox1.Text & " ,"
        sql &= " 權利範圍1分子 = " & TextBox3.Text & " ,"
        sql &= " 權利範圍2分母 = " & TextBox44.Text & " ,"
        sql &= " 權利範圍2分子 = " & TextBox45.Text & " ,"
        sql &= " 實際持有平方公尺 = " & TextBox22.Text & " ,"
        sql &= " 實際持有坪 = " & TextBox24.Text & ", "

        '1050616-V2版新增
        '土地使用分區-物件用途 
        If DropDownList66.Visible = True Then
            If DropDownList66.SelectedIndex > 0 Then
                sql &= " 使用分區='" & DropDownList66.SelectedValue & "', "
            Else
                If DropDownList65.SelectedValue <> "請選擇" Then
                    sql &= " 使用分區='" & DropDownList65.SelectedValue & "', "
                Else
                    sql &= " 使用分區='請選擇', "
                End If

            End If
        Else
            If DropDownList65.SelectedValue <> "請選擇" Then
                sql &= " 使用分區='" & DropDownList65.SelectedValue & "', "
            Else
                sql &= " 使用分區='請選擇', "
            End If
        End If

        '增建用途
        sql &= "增建用途='" & TextBox254.Text & "',"
        '增建完成日期
        sql &= "增建完成日期='" & Date8.Text & "',"

        '土地使用分區-管制 
        If DropDownList69.Visible = True Then
            If DropDownList69.SelectedValue = "請選擇" Then
                sql &= "管制='',"
            Else
                sql &= "管制='" & DropDownList69.SelectedValue & "',"
            End If
        Else
            sql &= "管制='',"
        End If

        '所有權人https://superwebnew.etwarm.com.tw
        sql &= "所有權人='" & TextBox255.Text & "', "

        '法定建蔽率
        sql &= "法定建蔽率='" & TextBox256.Text & "',"
        '法定容積率
        sql &= "法定容積率='" & TextBox257.Text & "',"
        'DDL_level2
        sql &= "DL_level2_selectindex=" & DDL_level2.SelectedIndex & ","
        '是否為公設
        sql &= "是否為公設='" & IIf(CheckBox98.Checked = True, "Y", "N") & "',"
        '是否為車位
        sql &= "是否為車位='" & IIf(CheckBox99.Checked = True, "Y", "N") & "'"
        sql &= " Where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

        Dim cmd As New SqlCommand(sql, conn)
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Dim 目前行號 As String = New System.Diagnostics.StackTrace(True).GetFrame(0).GetFileLineNumber().ToString()
            myobj.SQL_Error(Request.Url.ToString(), 目前行號, ex.ToString(), sql)
            '跳轉至出錯頁面 
            Response.Redirect("https://superwebnew.etwarm.com.tw/indexnew/ErrorPageNew.aspx")
        End Try


        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        If Me.DropDownList2.SelectedItem.Text = "土地面積" And 項目.ToString.IndexOf("車位") < 0 Then
            If DropDownList3.SelectedValue = "土地" Then
                '只有土地
                '刪除單筆產調_基地----------------------------
                sql = "Delete "
                sql &= " From 產調_基地 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_車位----------------------------
                sql = "Delete "
                sql &= " From 產調_車位 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_建物----------------------------
                sql = "Delete "
                sql &= " From 產調_建物 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            Else
                '只留基地
                '刪除單筆產調_土地----------------------------
                sql = "Delete "
                sql &= " From 產調_土地 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_車位----------------------------
                sql = "Delete "
                sql &= " From 產調_車位 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_建物----------------------------
                sql = "Delete "
                sql &= " From 產調_建物 "
                sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If

        ElseIf Me.DropDownList2.SelectedItem.Text = "車位面積(產權獨立)" Or Me.DropDownList2.SelectedItem.Text = "車位面積(公設內)" Or 項目.ToString.IndexOf("車位") >= 0 Or 項目.ToString = "停車空間" Then
            '只留車位
            '刪除單筆產調_基地----------------------------
            sql = "Delete "
            sql &= " From 產調_基地 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

            '刪除單筆產調_土地----------------------------
            sql = "Delete "
            sql &= " From 產調_土地 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

            '刪除單筆產調_建物----------------------------
            sql = "Delete "
            sql &= " From 產調_建物 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
        ElseIf Me.DropDownList2.SelectedItem.Text = "主建物" Then
            '只留建物

            '刪除單筆產調_基地----------------------------
            sql = "Delete "
            sql &= " From 產調_基地 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

            '刪除單筆產調_土地----------------------------
            sql = "Delete "
            sql &= " From 產調_土地 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

            '刪除單筆產調_車位----------------------------
            sql = "Delete "
            sql &= " From 產調_車位 "
            sql &= " where 物件編號 = '" & 編號(0) & "' and 流水號='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End If
        'End If

        conn.Close()
        conn.Dispose()

    End Sub

    '刪除細項面積內容
    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        For i = 0 To Me.GridView1.Rows.Count - 1

            '判斷刪除CHK有選取時執行刪除動作
            If CType(Me.GridView1.Rows(i).FindControl("chkSelect1"), CheckBox).Checked Then

                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()

                '刪除單筆面積細項----------------------------
                sql = "Delete "
                sql &= " From 委賣物件資料表_面積細項 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_基地----------------------------
                sql = "Delete "
                sql &= " From 產調_基地 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_土地----------------------------
                sql = "Delete "
                sql &= " From 產調_土地 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_車位----------------------------
                sql = "Delete "
                sql &= " From 產調_車位 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除單筆產調_建物----------------------------
                sql = "Delete "
                sql &= " From 產調_建物 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除細項所有權人 20160614 by Finch
                sql = "Delete "
                sql &= " From 委賣物件資料表_細項所有權人 "
                sql &= " where 物件編號 = '" & CType(Me.GridView1.Rows(i).FindControl("Label35"), Label).Text & "' and 細項流水號='" & CType(Me.GridView1.Rows(i).FindControl("Label36"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                conn.Close()
                conn.Dispose()

            End If
        Next

        '重整GRIDVIEW
        Load_Data("OLD")


        更新車位()


        ''重新計算是否還有車位細項 如果都沒有則將此物件的車位欄位更新 2016/04/12 by nick
        'Dim GOTOUPDATECARPARK As Boolean = False
        'For i As Integer = 0 To Me.GridView1.Rows.Count - 1
        '    Dim lb37 As Label = CType(Me.GridView1.Rows(i).FindControl("Label37"), Label) '類別
        '    Dim lb38 As Label = CType(Me.GridView1.Rows(i).FindControl("Label38"), Label) '項目

        '    If lb37.Text.IndexOf("車位") >= 0 Or lb38.Text.IndexOf("車位") >= 0 Then
        '        GOTOUPDATECARPARK = True
        '    End If
        'Next
        'If GOTOUPDATECARPARK = False Then
        '    Dim updstr As String = "Update 委賣物件資料表 set 車位 = '無' where 物件編號 = '" & Request.QueryString("oid") & "' and 店代號 = '" & Request.QueryString("sid") & "'"
        '    Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
        '        conn.Open()
        '        Using cmd As New SqlCommand(updstr, conn)
        '            cmd.ExecuteNonQuery()
        '        End Using
        '    End Using
        'End If

        '重算面積
        Total()

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        寫入坪數()
        'End If

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    If ImageButton1.visible = True Then
        '        ImageButton1_Click(sender, e)
        '    Else
        '        ImageButton12_Click(sender, e)
        '    End If
        'End If

        Dim script As String = ""
        script += "alert('刪除成功!!');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
        hidtempcode.Value = ""

        '恢復成原始狀態-----------------------
        clear()
        類別名稱()
        項目名稱("主建物")
        DDL_level2.SelectedIndex = 0
        '-------------------------------------

        '面積細項新增按鈕
        If myobj.AC = "1" Then
            Me.ImageButton5.Visible = True
        Else
            Me.ImageButton5.Visible = False
        End If
        Me.ImageButton3.Visible = False

        '新增的限制解除
        DropDownList2.Enabled = True
        Label76.Visible = False
        'End If

    End Sub
    '面積細項區塊(END)-----------------------------------------------------------------------------------------------------------------------------------------

















    '他項權利細項區塊(START)-------------------------------------------------------------------------------------------------------------------------------------
    '讀取他項權利內容
    Sub Load_他項權利Data(ByVal cls As String)
        '判斷物件維護權限
        myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        If cls = "OLD" Then
            sql = "Select * From 物件他項權利細項 With(NoLock) Where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'  order by 權利類別 desc ,順位 asc "

        Else
            Dim 編號 As Array = Split(cls, ",")
            sql = "Select * From 物件他項權利細項 With(NoLock) Where 物件編號 = '" & 編號(0) & "' and Num='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'  order by 權利類別 desc,順位 asc "
        End If


        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "細項內容")
        Dim tb_細項內容 As DataTable = ds.Tables("細項內容")

        If cls = "OLD" Then
            Me.GridView2.DataSource = tb_細項內容
            Me.GridView2.DataBind()

            If Me.GridView2.Rows.Count = 0 Then
                Me.ImageButton10.Visible = False
            Else
                '以此判斷物件是否過期被LOCK住不能修改
                '20151007 Textbox253 移除改用 Textbox4 來判斷
                'If TextBox253.Enabled = True Then
                If TextBox4.Enabled = True Then

                    '他項細項刪除按鈕
                    If myobj.D = "1" Then
                        Me.ImageButton10.Visible = True
                    Else
                        Me.ImageButton10.Visible = False
                    End If

                End If


            End If
        Else
            '他項類別
            If Not IsDBNull(tb_細項內容.Rows(0)("權利類別")) Then
                Me.DropDownList37.SelectedValue = tb_細項內容.Rows(0)("權利類別")
            End If

            '他項種類
            If Not IsDBNull(tb_細項內容.Rows(0)("權利種類")) Then
                Me.DropDownList36.SelectedValue = tb_細項內容.Rows(0)("權利種類")
            End If

            '順位
            If Not IsDBNull(tb_細項內容.Rows(0)("順位")) Then
                input44.Value = tb_細項內容.Rows(0)("順位")
            End If

            '登記日期
            If Not IsDBNull(tb_細項內容.Rows(0)("登記日期")) Then
                Date4.Text = tb_細項內容.Rows(0)("登記日期")
            End If

            '設定
            If Not IsDBNull(tb_細項內容.Rows(0)("設定")) Then
                input46.Value = tb_細項內容.Rows(0)("設定")
            End If

            '設定權利人
            If Not IsDBNull(tb_細項內容.Rows(0)("設定權利人")) Then
                input47.Value = tb_細項內容.Rows(0)("設定權利人")
            End If

            ''管理人
            'If Not IsDBNull(tb_細項內容.Rows(0)("管理人")) Then
            '    input122.Value = tb_細項內容.Rows(0)("管理人")
            'End If

            '處理方式

            Dim way1, way2, way3, way4, way5, wayothers, way6 As String
            way1 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式1")), "", tb_細項內容.Rows(0)("處理方式1"))
            way2 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式2")), "", tb_細項內容.Rows(0)("處理方式2"))
            way3 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式3")), "", tb_細項內容.Rows(0)("處理方式3"))
            way4 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式4")), "", tb_細項內容.Rows(0)("處理方式4"))
            way5 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式5")), "", tb_細項內容.Rows(0)("處理方式5"))
            wayothers = IIf(IsDBNull(tb_細項內容.Rows(0)("其他說明")), "", tb_細項內容.Rows(0)("其他說明"))
            way6 = IIf(IsDBNull(tb_細項內容.Rows(0)("處理方式6")), "", tb_細項內容.Rows(0)("處理方式6"))
            For Each li As ListItem In CheckBoxList3.Items

                If (li.Value = "1" And way1.Length > 0) Or (li.Value = "2" And way2.Length > 0) Or (li.Value = "3" And way3.Length > 0) Or (li.Value = "4" And way4.Length > 0) Or (li.Value = "5" And way5.Length > 0) Or (li.Value = "6" And way6.Length > 0) Then
                    li.Selected = True
                Else
                    li.Selected = False
                End If

            Next
            TextBox38.Text = wayothers

            'Num
            Me.Label56.Text = cls


            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then

                '判斷為新增還是修改ㄉ狀態
                If Me.Label56.Text = "0" Then
                    '他項細項新增按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton4.Visible = True
                    Else
                        Me.ImageButton4.Visible = False
                    End If
                    Me.ImageButton9.Visible = False
                Else
                    Me.ImageButton4.Visible = False
                    '他項細項修改按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton9.Visible = True
                    Else
                        Me.ImageButton9.Visible = False
                    End If

                End If

            End If



        End If

        conn.Close()
        conn.Dispose()

        If Me.GridView2.Rows.Count = 0 Then
            Me.ImageButton10.Visible = False
        Else

            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then

                '他項細項刪除按鈕
                If myobj.D = "1" Then
                    Me.ImageButton10.Visible = True
                Else
                    Me.ImageButton10.Visible = False
                End If

            End If

        End If

    End Sub

    Sub clear_他項權利()
        Me.Label56.Text = "0"
        DropDownList37.SelectedIndex = 0
        DropDownList36.SelectedIndex = 0
        input44.Value = ""
        Date4.Text = ""
        input46.Value = ""
        input47.Value = ""
        'input122.Value = ""
    End Sub

    '修改
    Sub updt_他項權利()
        '修改他項權利細項
        update_他項權利細項()


        '重整GRIDVIEW
        Load_他項權利Data("OLD")

        '清空所有細項
        clear_他項權利()

        Dim script As String = ""
        script += "alert('修改成功2!!');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        '判斷物件維護權限
        myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")

        '判斷為新增還是修改ㄉ狀態
        If Me.Label56.Text = "0" Then
            '他項細項新增按鈕
            If myobj.AC = "1" Then
                Me.ImageButton4.Visible = True
            Else
                Me.ImageButton4.Visible = False
            End If
            Me.ImageButton9.Visible = False
        Else
            Me.ImageButton4.Visible = False
            '他項細項修改按鈕
            If myobj.AC = "1" Then
                Me.ImageButton9.Visible = True
            Else
                Me.ImageButton9.Visible = False
            End If

        End If
    End Sub

    '修改他項權利細項
    Sub update_他項權利細項()
        '處理方式
        Dim way1, way2, way3, way4, way5, wayothers, way6 As String
        way1 = ""
        way2 = ""
        way3 = ""
        way4 = ""
        way5 = ""
        way6 = ""
        wayothers = ""
        For Each li As ListItem In CheckBoxList3.Items
            If li.Selected = True Then
                If li.Value = "1" Then
                    way1 = li.Text
                End If
                If li.Value = "2" Then
                    way2 = li.Text
                End If
                If li.Value = "3" Then
                    way3 = li.Text
                End If
                If li.Value = "4" Then
                    way4 = li.Text
                End If
                If li.Value = "5" Then
                    way5 = li.Text
                    wayothers = TextBox38.Text
                End If
                If li.Value = "6" Then
                    way6 = li.Text
                End If
            End If
        Next


        Dim 編號 As Array = Split(Me.Label56.Text, ",")

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        sql = "Update 物件他項權利細項   "
        sql &= " Set "
        sql &= " 權利類別 = '" & Me.DropDownList37.SelectedItem.Text & "' ,"
        sql &= " 權利種類 = '" & Me.DropDownList36.SelectedItem.Text & "' ,"
        sql &= " 順位 = '" & input44.Value & "' ,"
        sql &= " 登記日期 = '" & Left(Trim(Date4.Text), 7) & "' ,"
        sql &= " 設定 = '" & input46.Value & "' ,"
        sql &= " 設定權利人 = '" & input47.Value & "' ,"
        sql &= " 處理方式1 = '" & way1 & "',處理方式2 = '" & way2 & "',處理方式3 = '" & way3 & "',處理方式4 = '" & way4 & "',處理方式5 = '" & way5 & "',其他說明 = '" & wayothers & "'"
        sql &= " ,處理方式6 = '" & way6 & "' "
        sql &= " Where 物件編號 = '" & 編號(0) & "' and Num='" & 編號(1) & "' and 店代號='" & Me.Label12.Text & "'"

        'Response.Write(sql)


        Dim cmd As New SqlCommand(sql, conn)
        cmd.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

    End Sub

    '刪除他項權利細項
    Protected Sub ImageButton10_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton10.Click



        For i = 0 To Me.GridView2.Rows.Count - 1

            '判斷刪除CHK有選取時執行刪除動作
            If CType(Me.GridView2.Rows(i).FindControl("chkSelect2"), CheckBox).Checked Then

                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()

                '刪除他項權利細項----------------------------
                sql = "Delete "
                sql &= " From 物件他項權利細項 "
                sql &= " where 物件編號 = '" & CType(Me.GridView2.Rows(i).FindControl("Label48"), Label).Text & "' and Num='" & CType(Me.GridView2.Rows(i).FindControl("Label49"), Label).Text & "' and 店代號='" & Me.Label12.Text & "'"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                conn.Close()
                conn.Dispose()

            End If
        Next

        '重整GRIDVIEW
        Load_他項權利Data("OLD")

        Dim script As String = ""
        script += "alert('刪除成功!!');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)


    End Sub

    '新增他項權利
    Protected Sub ImageButton4_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton4.Click
        '判斷有無物件編號,無則跳出
        If Trim(Me.TextBox2.Text) = "" Then
            Dim script As String = ""
            script += "alert('請先輸入物件編號!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            Exit Sub
        End If

        '組合物件編號
        '物件編號-第1碼
        Dim 物件編號 As String = ""

        If Request("oid") = "" Then '新增時



            If ddl契約類別.SelectedValue = "專任" Then
                物件編號 = "1"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                物件編號 = "6"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                物件編號 = "7"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                物件編號 = "5"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                物件編號 = "9"
            End If

            '物件編號-第2-5碼(店代號)+第6-13碼(表單編號)
            If store.SelectedValue = "請選擇" Then
                物件編號 &= Mid(Request.Cookies("store_id").Value, 2) & TextBox2.Text.Trim
            Else
                物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            End If

        Else '修改複製
            物件編號 = Request("oid")
        End If

        '新增時如是空值，給予當下的資料值為預設值(整筆資料未存檔前)
        If Me.Label11.Text = "" And Me.Label12.Text = "" Then

            '物件編號
            Me.Label11.Text = 物件編號

            '店代號
            Me.Label12.Text = store.SelectedValue
        End If


        If Me.DropDownList37.SelectedItem.Text <> "" Then

            '如果編號跟店代號跟預設值不一樣，先下列步驟
            If Me.Label11.Text <> 物件編號 Or Me.Label12.Text <> store.SelectedValue Then
                Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
                Dim sql_upt As String = "update 物件他項權利細項 set 物件編號='" & 物件編號 & "',店代號='" & store.SelectedValue & "' where 物件編號='" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"
                Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
                cmd_upt.CommandType = CommandType.Text

                conn_upt.Open()

                cmd_upt.ExecuteNonQuery()

                conn_upt.Close()
                conn_upt.Dispose()


                'UPDATE已存資料後給予新的值
                '物件編號
                Me.Label11.Text = 物件編號

                '店代號
                Me.Label12.Text = store.SelectedValue

            End If

            conn = New SqlConnection(EGOUPLOADSqlConnStr)
            Dim count As Integer
            Dim sql2 As String = "select top 1 Num as Num from 物件他項權利細項 With(NoLock)  where 物件編號 = '" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "' order by Num desc"
            Dim cmd2 As New SqlCommand(sql2, conn)
            cmd2.CommandType = CommandType.Text
            adpt = New SqlDataAdapter(sql2, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count > 0 Then
                count = table2.Rows(0)("Num") + 1
            Else
                count = 0
            End If

            '處理方式
            Dim way1, way2, way3, way4, way5, wayothers, way6 As String
            way1 = ""
            way2 = ""
            way3 = ""
            way4 = ""
            way5 = ""
            wayothers = ""
            way6 = ""
            For Each li As ListItem In CheckBoxList3.Items
                If li.Selected = True Then
                    If li.Value = "1" Then
                        way1 = li.Text
                    End If
                    If li.Value = "2" Then
                        way2 = li.Text
                    End If
                    If li.Value = "3" Then
                        way3 = li.Text
                    End If
                    If li.Value = "4" Then
                        way4 = li.Text
                    End If
                    If li.Value = "5" Then
                        way5 = li.Text
                        wayothers = TextBox38.Text
                    End If
                    If li.Value = "6" Then
                        way6 = li.Text
                    End If
                End If
            Next

            Dim sql As String = "insert 物件他項權利細項(物件編號,店代號, Num, 權利類別, 權利種類,順位,登記日期,設定,設定權利人,處理方式1,處理方式2,處理方式3,處理方式4,處理方式5,處理方式6,其他說明) values ('" & Me.Label11.Text & "','" & Me.Label12.Text & "','" & count & "','" & DropDownList37.SelectedValue & "','" & DropDownList36.SelectedValue & "','" & input44.Value & "','" & Left(Trim(Date4.Text), 7) & "','" & input46.Value & "','" & input47.Value & "','" & way1 & "','" & way2 & "','" & way3 & "','" & way4 & "','" & way5 & "','" & way6 & "','" & wayothers & "')"
            Dim cmd As New SqlCommand(sql, conn)
            cmd.CommandType = CommandType.Text

            conn.Open()
            Try
                cmd.ExecuteNonQuery()

                Dim script As String = ""
                script += "alert('新增成功!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

                '恢復成原始狀態-----------------------
                clear_他項權利()
                '-------------------------------------

                '讀取資料
                Load_他項權利Data("OLD")
            Catch
                Dim script As String = ""
                script += "alert('新增失敗!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            End Try

            cmd2.ExecuteNonQuery()
            conn.Close()
            conn.Dispose()

        Else
            Dim script As String = ""
            script += "alert('請選擇類別名稱!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        End If

    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then

            '物件編號LABEL
            Dim lbl物件編號 As Label = e.Row.FindControl("Label48")

            'NUMLABEL
            Dim lblNUM As Label = e.Row.FindControl("Label49")

            '編輯button
            Dim btn編輯 As Button = e.Row.FindControl("Button9")
            btn編輯.CommandArgument = lbl物件編號.Text & "," & lblNUM.Text

        End If
    End Sub

    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        If e.CommandName = "edits" Then

            Load_他項權利Data(e.CommandArgument)

            '判斷物件維護權限
            myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")


            '以此判斷物件是否過期被LOCK住不能修改
            '20151007 Textbox253 移除改用 Textbox4 來判斷
            'If TextBox253.Enabled = True Then
            If TextBox4.Enabled = True Then
                '判斷為新增還是修改ㄉ狀態
                If Me.Label56.Text = "0" Then
                    '他項細項新增按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton4.Visible = True
                    Else
                        Me.ImageButton4.Visible = False
                    End If
                    Me.ImageButton9.Visible = False
                Else
                    Me.ImageButton4.Visible = False
                    '他項細項修改按鈕
                    If myobj.AC = "1" Then
                        Me.ImageButton9.Visible = True
                    Else
                        Me.ImageButton9.Visible = False
                    End If

                End If

            End If

        End If
    End Sub

    Protected Sub ImageButton9_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton9.Click
        If Me.DropDownList37.SelectedIndex = 0 Or Me.DropDownList36.SelectedIndex = 0 Then
            Dim script As String = ""
            script += "alert('請選擇他項權利類別.他項權利種類!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        Else
            updt_他項權利()
        End If
    End Sub
    '他項權利細項區塊(END)---------------------------------------------------------------------------------------------------------------------------------------
















    '車位區塊(START)---------------------------------------------------------------------------------------------------------------------------------------------------
    '車位種類
    'Protected Sub ddl車位類別_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If ddl車位類別.SelectedValue.Trim = "其他" Then
    '        Me.TextBox37.Visible = True
    '    Else
    '        TextBox37.Text = ""
    '        Me.TextBox37.Visible = False

    '    End If

    '    'ScriptManager1.SetFocus(ddl車位類別)

    'End Sub

    Protected Sub CheckBox3_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged

        If CheckBox3.Checked = True Then
            Me.input55.Value = "含於開價中"
            Me.input55.Attributes("ReadOnly") = "True"
        Else
            If Me.input55.Value = "含於開價中" Then
                Me.input55.Value = ""
                Me.input55.Attributes.Remove("ReadOnly")
            End If
        End If

    End Sub
    ''車位管理費CHANGE事件
    'Protected Sub DropDownList24_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList24.SelectedIndexChanged
    '    If DropDownList24.SelectedValue = "無" Or DropDownList24.SelectedValue = "未知" Then
    '        TextBox94.Text = ""
    '        TextBox94.Visible = False
    '        Label46.Visible = False
    '    Else
    '        TextBox94.Visible = True
    '        Label46.Visible = True
    '    End If
    'End Sub
    '車位區塊(END)-----------------------------------------------------------------------------------------------------------------------------------------------------










    '生活週遭區塊(START)------------------------------------------------------------------------------------------------------------------------------------------------
    '捷運
    Protected Sub DropDownList8_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList8.SelectedIndexChanged
        Dim conn_捷運 As SqlConnection
        conn_捷運 = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_捷運.Open()

        sql = "Select 站名,代號 From 資料_捷運 With(NoLock)  "
        sql &= "Where 路線 = '" & DropDownList8.SelectedValue & "' order by sort,代號 "
        adpt = New SqlDataAdapter(sql, conn_捷運)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        DropDownList9.Items.Clear()
        With DropDownList9
            .DataSource = table1.DefaultView
            .DataTextField = "站名"
            .DataValueField = "代號"
            .DataBind()
        End With
        DropDownList9.Items.Insert(0, "選擇站名")

        conn_捷運.Close()
        conn_捷運 = Nothing

        'ScriptManager1.SetFocus(DropDownList8)
    End Sub

    '生活週遭區塊(END)------------------------------------------------------------------------------------------------------------------------------------------------









    Protected Sub DDL_County_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL_County.SelectedIndexChanged
        city_change("F")
    End Sub

    '縣市資料讀取
    Sub City_Data()
        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()

            sql = " Select 縣市 from 資料_縣市 With(NoLock) "

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            DDL_County.Items.Clear()
            With DDL_County
                .DataSource = table1.DefaultView
                .DataTextField = Trim("縣市")
                .DataValueField = Trim("縣市")
                .DataBind()
                .Items.Insert(0, New ListItem("選擇縣市", "選擇縣市"))
            End With
            conn.Close()
            conn.Dispose()
        End Using

    End Sub

    '縣市Change事件
    Sub city_change(ByVal TorF As String)
        Dim conn_縣市 As New SqlConnection(EGOUPLOADSqlConnStr)
        conn_縣市.Open()
        If TorF = "T" Or TorF = "F" Then
            sql = " Select 鄉鎮市區,郵遞區號 from 資料_鄉鎮市區 With(NoLock) where 縣市名='" & DDL_County.SelectedValue & "'"

        End If


        adpt = New SqlDataAdapter(sql, conn_縣市)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If TorF = "T" Or TorF = "F" Then
            DDL_Area.Items.Clear()
            With DDL_Area
                .DataSource = table1.DefaultView
                .DataTextField = Trim("鄉鎮市區")
                .DataValueField = Trim("郵遞區號")
                .DataBind()
                .Items.Insert(0, New ListItem("選擇鄉鎮市區", "選擇鄉鎮市區"))
            End With
        End If

        '捷運路線
        sql = "SELECT DISTINCT 路線 FROM 資料_捷運 With(NoLock) "
        adpt = New SqlDataAdapter(sql, conn_縣市)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        DropDownList8.Items.Clear()
        DropDownList8.Items.Add("選擇路線")

        For i = 0 To table1.Rows.Count - 1
            '依照選擇的縣市來決定路線
            If DDL_County.SelectedValue = "新北市" Or DDL_County.SelectedValue = "台北市" Then
                If Not (table1.Rows(i)("路線") = "高雄紅線" Or table1.Rows(i)("路線") = "高雄橘線" Or table1.Rows(i)("路線") = "高雄輕軌" Or table1.Rows(i)("路線") = "台中綠線" Or table1.Rows(i)("路線") = "台中藍線") Then
                    DropDownList8.Items.Add(table1.Rows(i)("路線"))
                End If
            ElseIf DDL_County.SelectedValue = "高雄市" Then
                If table1.Rows(i)("路線") = "高雄紅線" Or table1.Rows(i)("路線") = "高雄橘線" Or table1.Rows(i)("路線") = "高雄輕軌" Then
                    DropDownList8.Items.Add(table1.Rows(i)("路線"))
                End If
            ElseIf DDL_County.SelectedValue = "台中市" Then
                If table1.Rows(i)("路線") = "台中綠線" Or table1.Rows(i)("路線") = "台中藍線" Then
                    DropDownList8.Items.Add(table1.Rows(i)("路線"))
                End If
            ElseIf DDL_County.SelectedValue = "桃園市" Then
                If table1.Rows(i)("路線") = "機場捷運線" Then
                    DropDownList8.Items.Add(table1.Rows(i)("路線"))
                End If
            End If
        Next

        conn_縣市.Close()
        conn_縣市.Dispose()

        If TorF = "F" Then
            TB_AreaCode.Text = ""
        End If
    End Sub

    '鄉鎮市區CHANGE事件
    Protected Sub DDL_Area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL_Area.SelectedIndexChanged
        If DDL_Area.SelectedValue = "選擇鄉鎮市區" Then
            TB_AreaCode.Text = ""
        Else
            TB_AreaCode.Text = DDL_Area.SelectedValue

            '學區用-session
            Session("County") = DDL_County.SelectedValue
            Session("Town") = DDL_Area.SelectedValue
        End If

        社區大樓()

    End Sub

    Sub 社區大樓()
        Dim conn_社區大樓 As SqlConnection = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_社區大樓.Open()

        '社區大樓
        sql = "select 大樓編號,社區建案名稱 from 社區大樓資料表 With(NoLock) Where 郵遞區號 = '" & TB_AreaCode.Text & "' "
        If Trim(TextBox252.Text) <> "" Then
            sql &= " and 社區建案名稱 like '%" & TextBox252.Text & "%' "
        End If
        sql &= " and 狀態='A' "
        sql &= " order by 社區建案名稱 "
        adpt = New SqlDataAdapter(sql, conn_社區大樓)
        ds = New DataSet
        adpt.Fill(ds, "mytable")
        Dim mytable As DataTable = ds.Tables("mytable")

        'ddl社區大樓.Items.Clear()
        With ddl社區大樓
            .DataSource = mytable.DefaultView
            .DataTextField = "社區建案名稱"
            .DataValueField = "大樓編號"
            .DataBind()
            .Items.Insert(0, New ListItem("請選擇", "請選擇"))
        End With

        conn_社區大樓.Close()
        conn_社區大樓.Dispose()
    End Sub

    '郵遞區號TEXTBOX的CHANGE事件
    Protected Sub TB_AreaCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_AreaCode.TextChanged
        '取得社區大樓資料
        社區大樓()

        conn = New SqlConnection(EGOUPLOADSqlConnStr)

        sql = " Select * from 資料_鄉鎮市區 With(NoLock) where 郵遞區號='" & Trim(TB_AreaCode.Text) & "'"

        Dim cmd As New SqlCommand(sql, conn)

        conn.Open()

        Dim dr As SqlDataReader = cmd.ExecuteReader

        If dr.Read() Then
            City_Data()
            DDL_County.SelectedValue = Trim(dr("縣市名"))
            city_change("T")
            DDL_Area.SelectedValue = TB_AreaCode.Text
        Else
            City_Data()

            DDL_Area.Items.Clear()
            DDL_Area.Items.Insert(0, New ListItem("選擇鄉鎮市區", "選擇鄉鎮市區"))

        End If

        conn.Close()
        conn.Dispose()
    End Sub

    'Protected Sub ddl契約類別_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl契約類別.SelectedIndexChanged

    'End Sub
    '店代號CHANGE事件
    Protected Sub store_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles store.SelectedIndexChanged

        sale1.Items.Clear()
        sale2.Items.Clear()
        sale3.Items.Clear()

        sale1.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人2
        sale2.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人2
        sale3.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人3

        sale1 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale1) '承辦人1 
        sale2 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale2) '承辦人2 
        sale3 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale3) '承辦人3 

        '取得該店可用表單
        form_own("'" & store.SelectedValue.ToString & "'", ddl契約類別.SelectedValue)

        '不動產說明書報告單位
        Dim 店名 As Array = Split(store.SelectedItem.Text, ",")
        input2.Value = 店名(1)

        '取得該店可用表單
        'form_own("'" & store.SelectedValue.ToString & "'", ddl契約類別.SelectedValue)
    End Sub

    '顯示全部人員(多店)CHANGE事件
    Protected Sub all_people_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles all_people.CheckedChanged
        sale1.Items.Clear()
        sale2.Items.Clear()
        sale3.Items.Clear()

        sale1.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人2
        sale2.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人2
        sale3.Items.Add(New ListItem("選擇人員", "選擇人員")) '承辦人3

        If all_people.Checked Then '顯示所有人員(多店用)
            sale1 = clspowerset.people(Request.Cookies("webfly_empno").Value, "請選擇", sale1) '承辦人1 
            sale2 = clspowerset.people(Request.Cookies("webfly_empno").Value, "請選擇", sale2) '承辦人2 
            sale3 = clspowerset.people(Request.Cookies("webfly_empno").Value, "請選擇", sale3) '承辦人3 
        Else
            sale1 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale1) '承辦人1 
            sale2 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale2) '承辦人2 
            sale3 = clspowerset.people(Request.Cookies("webfly_empno").Value, store.SelectedValue, sale3) '承辦人3 
        End If
    End Sub

    '表單類別CHANGE事件
    Protected Sub ddl契約類別_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl契約類別.SelectedIndexChanged
        If ddl契約類別.SelectedValue = "一般" Or ddl契約類別.SelectedValue = "同意書" Or ddl契約類別.SelectedValue = "流通" Then
            If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                CheckBox1.Visible = False
                CheckBox1.Checked = False 'change事件完設為TRUE
            Else
                CheckBox1.Visible = True
                CheckBox1.Checked = True 'change事件完設為TRUE
            End If
            Label27.Visible = False

            '已購買表單編號-顯示
            Me.Label28.Visible = True
            Me.DropDownList4.Visible = True
        ElseIf ddl契約類別.SelectedValue = "專任" Then
            ' 僅直營且為管理角色時顯示核取方塊
            CheckBox1.Visible = IS_DIRECTLY_OPERATION And IS_MANAGER_ROLE

            ' 直營且非管理角色時不勾選，其餘預設為已勾選
            CheckBox1.Checked = Not (IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE)

            Label27.Visible = False

            '已購買表單編號-顯示
            Me.Label28.Visible = True
            Me.DropDownList4.Visible = True
        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            CheckBox1.Visible = False
            CheckBox1.Checked = False '庫存為False
            Label27.Visible = True

            '已購買表單編號-隱藏
            Me.Label28.Visible = False
            Me.DropDownList4.Visible = False
        End If

        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            Me.Label464.Visible = True
            Me.RadioButton1.Visible = True
            Me.RadioButton2.Visible = True
            Me.RadioButton1.Checked = True
            Me.RadioButton2.Checked = False
        Else
            Me.Label464.Visible = False
            Me.RadioButton1.Visible = False
            Me.RadioButton2.Visible = False
            Me.RadioButton1.Checked = False
            Me.RadioButton2.Checked = False
        End If

        If jointsalerule = "3" Then
            If 一般專約 = "1" And ddl契約類別.SelectedValue = "一般" Then
                CheckBox101.Checked = False
                CheckBox101.Visible = True
            ElseIf 一般專約 = "2" And ddl契約類別.SelectedValue = "專任" Then
                CheckBox101.Checked = False
                CheckBox101.Visible = True
            ElseIf 一般專約 = "3" Then
                CheckBox101.Checked = False
                CheckBox101.Visible = True
            Else
                CheckBox101.Checked = False
                CheckBox101.Visible = False
            End If
            'CheckBox101.checked = False
            'CheckBox101.visible = True
        Else
            CheckBox101.Checked = False
            CheckBox101.Visible = False
        End If

        '取得該店可用表單
        form_own("'" & store.SelectedValue.ToString & "'", ddl契約類別.SelectedValue)


    End Sub

    '讀取該店所購買表單列表
    Sub form_own(ByVal storeid As String, ByVal Cls As String)

        Dim 類別 As String = ""
        Dim conn_own As SqlConnection
        Select Case Cls
            Case "一般"
                類別 = "'M','A','R','1'"  'M-土地一般 A-物件一般 R-預售屋 1-直營物件一般
            Case "專任"
                類別 = "'L','B','S','2'" 'L-土地專任 B-物件專任 S-預售屋
            Case "同意書"
                類別 = "'N','C'" 'N-土地同意書 C-物件同意書
            Case "流通"
                類別 = "'X'"
        End Select

        If Cls <> "庫存" Then
            conn_own = New SqlConnection(EGOUPLOADSqlConnStr)
            conn_own.Open()

            sql = " Select distinct 合約編號 From 秘書_合約管制檔可用 With(NoLock) Where Left(合約編號,1) in (" & 類別 & ")"
            sql &= " and 店代號 IN (" & storeid & ") and 上傳註記<>'D'  and 處理類別 not in ('作廢','遺失')"
            sql &= " order by 合約編號"

            'If Request.Cookies("webfly_empno").Value = "D69" Then
            '    'eip_usual.Show(sql)
            '    Response.Write(sql)
            'End If

            adpt = New SqlDataAdapter(sql, conn_own)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            DropDownList4.Items.Clear()
            DropDownList4.Items.Add(New ListItem("選擇表單", "選擇表單"))

            With DropDownList4
                .DataSource = table1.DefaultView
                .DataTextField = "合約編號"
                .DataValueField = "合約編號"
                .DataBind()
            End With

            conn_own.Close()
            conn_own.Dispose()
        End If



    End Sub

    '表單編號TEXTBOX的CHANGE事件-判斷是否為已購買表單
    Protected Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If Trim(TextBox2.Text) <> "" Then
            '判斷輸入表單編號是否正確
            Dim formtype As String
            Dim objecttype As String = Right(DropDownList3.SelectedValue, 1)

            If objecttype = "地" Then
                If ddl契約類別.SelectedValue = "專任" Then
                    formtype = "1233,1201,2201,1243"
                ElseIf ddl契約類別.SelectedValue = "一般" Then
                    'formtype = "1235" 因目前暫無故會以1229代
                    formtype = "1235,1229,2229,1242"
                ElseIf ddl契約類別.SelectedValue = "同意書" Then
                    formtype = "1240,1231"
                ElseIf ddl契約類別.SelectedValue = "流通" Then
                    formtype = "0000"
                ElseIf ddl契約類別.SelectedValue = "庫存" Then
                    formtype = "9999"
                End If
            Else
                If ddl契約類別.SelectedValue = "專任" Then
                    formtype = "1201,2201,1243"
                ElseIf ddl契約類別.SelectedValue = "一般" Then
                    formtype = "1229,2229,1242"
                ElseIf ddl契約類別.SelectedValue = "同意書" Then
                    formtype = "1231"
                ElseIf ddl契約類別.SelectedValue = "流通" Then
                    formtype = "0000"
                ElseIf ddl契約類別.SelectedValue = "庫存" Then
                    formtype = "9999"
                End If
            End If

            'If objecttype = "地" Then
            '    If ddl契約類別.SelectedValue = "專任" Then
            '        formtype = "1233,1201"
            '    ElseIf ddl契約類別.SelectedValue = "一般" Then
            '        'formtype = "1235" 因目前暫無故會以1229代
            '        formtype = "1235,1229"
            '    ElseIf ddl契約類別.SelectedValue = "同意書" Then
            '        formtype = "1240,1231"
            '    ElseIf ddl契約類別.SelectedValue = "流通" Then
            '        formtype = "0000"
            '    ElseIf ddl契約類別.SelectedValue = "庫存" Then
            '        formtype = "9999"
            '    End If
            'Else
            '    If ddl契約類別.SelectedValue = "專任" Then
            '        formtype = "1201"
            '    ElseIf ddl契約類別.SelectedValue = "一般" Then
            '        formtype = "1229"
            '    ElseIf ddl契約類別.SelectedValue = "同意書" Then
            '        formtype = "1231"
            '    ElseIf ddl契約類別.SelectedValue = "流通" Then
            '        formtype = "0000"
            '    ElseIf ddl契約類別.SelectedValue = "庫存" Then
            '        formtype = "9999"
            '    End If
            'End If

            '檢查物件編號是否重複
            'If TextBox2.Text.Length >= 8 Then
            '    Dim objid As String = ""
            '    If ddl契約類別.SelectedValue = "專任" Then
            '        objid = "1"
            '    ElseIf ddl契約類別.SelectedValue = "一般" Then
            '        objid = "6"
            '    ElseIf ddl契約類別.SelectedValue = "同意書" Then
            '        objid = "7"
            '    ElseIf ddl契約類別.SelectedValue = "流通" Then
            '        objid = "5"
            '    ElseIf ddl契約類別.SelectedValue = "庫存" Then
            '        objid = "9"
            '    End If
            '    objid &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            '    Dim selstr As String = "select 店代號 from 委賣物件資料表 where 物件編號 = '" & objid & "' and 店代號 = '" & store.SelectedValue & "' union select 店代號 from 委賣物件過期資料表 where 物件編號 = '" & objid & "' and 店代號 = '" & store.SelectedValue & "'"
            '    Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString)
            '        conn.Open()
            '        Using cmd As New SqlCommand(selstr, conn)
            '            Try
            '                Dim dt As New DataTable
            '                dt.Load(cmd.ExecuteReader())
            '                If dt.Rows.Count > 0 Then
            '                    Dim script As String = ""
            '                    script += "alert('物件編號重複!');"
            '                    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            '                End If

            '            Catch ex As Exception

            '            End Try
            '        End Using
            '    End Using
            'End If



            If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
                Me.Label464.Visible = True
                Me.RadioButton1.Visible = True
                Me.RadioButton2.Visible = True
                Me.RadioButton1.Checked = False
                Me.RadioButton2.Checked = False
            Else
                Me.Label464.Visible = False
                Me.RadioButton1.Visible = False
                Me.RadioButton2.Visible = False
                Me.RadioButton1.Checked = False
                Me.RadioButton2.Checked = False
            End If

            'Dim message As String = formnocheck.checkform(formtype, TextBox2.Text, store.SelectedValue)
            ''Response.Write(formtype & "," & TextBox2.Text & "," & store.SelectedValue)
            'If Len(message) > 4 Then
            '    Dim script As String = ""
            '    script += "alert('" & message & "');"
            '    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            'End If
        End If

    End Sub

    '可用表單CHANGE事件
    Protected Sub DropDownList4_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList4.SelectedIndexChanged
        Me.TextBox2.Text = Me.DropDownList4.SelectedValue
        Me.DropDownList4.SelectedIndex = 0
        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            Me.Label464.Visible = True
            Me.RadioButton1.Visible = True
            Me.RadioButton2.Visible = True
            Me.RadioButton1.Checked = True
            Me.RadioButton2.Checked = False
        Else
            Me.Label464.Visible = False
            Me.RadioButton1.Visible = False
            Me.RadioButton2.Visible = False
            Me.RadioButton1.Checked = False
            Me.RadioButton2.Checked = False
        End If
    End Sub

    '土地使用分區CHANGE事件(主要)
    'Protected Sub DropDownList16_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList16.SelectedIndexChanged
    '    If DropDownList16.SelectedValue.Trim = "住宅區" Or DropDownList16.SelectedValue.Trim = "商業區" Or DropDownList16.SelectedValue.Trim = "工業區" Or DropDownList16.SelectedValue.Trim = "非都市計畫區" Or DropDownList16.SelectedValue.Trim = "公共設施用地" Or DropDownList16.SelectedValue.Trim = "河川區" Then
    '        DropDownList17.Visible = True
    '        Select_分區細項("主")
    '    Else
    '        DropDownList17.Visible = False
    '    End If

    '    'If DropDownList16.SelectedValue <> "請選擇" Then
    '    '    Label47.Text = DropDownList16.SelectedValue
    '    '    If Trim(TextBox253.text) <> "" Then
    '    '        Label47.Text &= "," & Trim(TextBox253.text)
    '    '    End If
    '    'Else
    '    '    Label47.Text = ""
    '    'End If
    'End Sub

    '土地使用分區CHANGE事件(副)
    'Protected Sub DropDownList11_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList11.SelectedIndexChanged
    '    If DropDownList11.SelectedValue.Trim = "住宅區" Or DropDownList11.SelectedValue.Trim = "商業區" Or DropDownList11.SelectedValue.Trim = "工業區" Or DropDownList11.SelectedValue.Trim = "非都市計畫區" Then
    '        DropDownList12.Visible = True
    '        Select_分區細項("副")
    '    Else
    '        DropDownList12.Visible = False
    '    End If

    'End Sub

    '土地使用分區CHANGE事件(細項)
    Protected Sub DropDownList65_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList65.SelectedIndexChanged
        If DropDownList65.SelectedValue.Trim = "住宅區" Or DropDownList65.SelectedValue.Trim = "風景區" Or DropDownList65.SelectedValue.Trim = "商業區" Or DropDownList65.SelectedValue.Trim = "工業區" Or DropDownList65.SelectedValue.Trim = "非都市計畫區" Or DropDownList65.SelectedValue.Trim = "公共設施用地" Or DropDownList65.SelectedValue.Trim = "鄉村區" Or DropDownList65.SelectedValue.Trim = "一般農業區" Or DropDownList65.SelectedValue.Trim = "山坡地保育區" Or DropDownList65.SelectedValue.Trim = "特定農業區" Or DropDownList65.SelectedValue.Trim = "森林區" Or DropDownList65.SelectedValue.Trim = "部分保護區" Or DropDownList65.SelectedValue.Trim = "國家公園區" Or DropDownList65.SelectedValue.Trim = "部分漁港專用區" Or DropDownList65.SelectedValue.Trim = "市場" Or DropDownList65.SelectedValue.Trim = "特定專用區" Then
            DropDownList66.Visible = True
            Select_分區細項("細項")
        Else
            DropDownList66.Visible = False
        End If
    End Sub

    '土地使用分區小分類CHANGE事件(主要)
    'Protected Sub DropDownList17_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList17.SelectedIndexChanged
    'If DropDownList17.SelectedValue <> "請選擇" Then
    '    Label47.Text = DropDownList17.SelectedValue
    'Else
    '    Label47.Text = DropDownList16.SelectedValue
    'End If
    'End Sub

    Sub Select_分區細項(ByRef cls As String)
        '使用分區    
        Dim conn_使用分區 As New SqlConnection(EGOUPLOADSqlConnStr)
        conn_使用分區.Open()
        sql = "select * from 資料_使用分區 With(NoLock) where 使用分區大項='" & Me.DropDownList65.SelectedValue & "' order by 使用分區小項"

        adpt = New SqlDataAdapter(sql, conn_使用分區)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If cls = "細項" Then
            DropDownList66.Items.Clear()
            DropDownList66.Items.Add("請選擇")
            For i = 0 To table1.Rows.Count - 1
                DropDownList66.Items.Add(table1.Rows(i)("使用分區小項").ToString.Trim)
                DropDownList66.Items(i + 1).Value = table1.Rows(i)("使用分區小項").ToString.Trim
            Next
        End If

        conn_使用分區.Close()
        conn_使用分區.Dispose()
    End Sub

    '建築主要用途CHANGE事件
    Protected Sub DropDownList19_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList19.SelectedIndexChanged
        If DropDownList19.SelectedValue.Trim = "其他" Then
            TextBox4.Visible = True
        Else
            TextBox4.Text = ""
            TextBox4.Visible = False
        End If

    End Sub

    'Protected Sub DropDownList14_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList14.SelectedIndexChanged
    '    If DropDownList14.SelectedValue.Trim = "其他" Then
    '        TextBox232.Visible = True
    '    Else
    '        TextBox232.Text = ""
    '        TextBox232.Visible = False
    '    End If
    '    'ScriptManager1.SetFocus(DropDownList14)
    'End Sub

    '開放(格局)CHANGE事件
    Protected Sub C1_CheckedChanged(sender As Object, e As System.EventArgs) Handles C1.CheckedChanged
        If C1.Checked = True Then
            TextBox13.Text = "-1"
            TextBox13.Visible = False
            Label29.Visible = False
            TextBox14.Text = ""
            TextBox14.Visible = False
            Label30.Visible = False
            TextBox15.Text = ""
            TextBox15.Visible = False
            Label31.Visible = False
            TextBox16.Text = ""
            TextBox16.Visible = False
            Label32.Visible = False

        Else
            TextBox13.Text = ""
            TextBox13.Visible = True
            Label29.Visible = True
            TextBox14.Text = ""
            TextBox14.Visible = True
            Label30.Visible = True
            TextBox15.Text = ""
            TextBox15.Visible = True
            Label31.Visible = True
            TextBox16.Text = ""
            TextBox16.Visible = True
            Label32.Visible = True
        End If

    End Sub

    '管理費單位CHANGE事件
    Protected Sub DropDownList5_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged ', DropDownList24.SelectedIndexChanged
        If DropDownList5.SelectedValue = "無" Or DropDownList5.SelectedValue = "未知" Then
            TextBox36.Text = ""
            TextBox36.Visible = False
            Label45.Visible = False
        Else
            TextBox36.Visible = True
            Label45.Visible = True
        End If

    End Sub
    'CHECK平方公尺不得為空值
    Function chk_平方公尺() As String
        Dim error_msg As String = ""

        If Me.TextBox6.Text <> "" And Trim(Me.TextBox5.Text) = "" Then
            error_msg += "主建物平方公尺不能為空白 \n"
        End If

        If Me.TextBox8.Text <> "" And Trim(Me.TextBox7.Text) = "" Then
            error_msg += "附屬建物平方公尺不能為空白 \n"
        End If

        If Me.TextBox10.Text <> "" And Trim(Me.TextBox9.Text) = "" Then
            error_msg += "公共設施平方公尺不能為空白 \n"
        End If

        If Me.TextBox20.Text <> "" And Trim(Me.TextBox19.Text) = "" Then
            error_msg += "地下室平方公尺不能為空白 \n"
        End If

        If Me.TextBox23.Text <> "" And Trim(Me.TextBox21.Text) = "" Then
            error_msg += "車位面積(公設內)平方公尺不能為空白 \n"
        End If

        If Me.TextBox27.Text <> "" And Trim(Me.TextBox26.Text) = "" Then
            error_msg += "車位面積(產權)平方公尺不能為空白 \n"
        End If

        If Me.TextBox31.Text <> "" And Trim(Me.TextBox30.Text) = "" Then
            error_msg += "土地面積平方公尺不能為空白 \n"
        End If

        If Me.TextBox29.Text <> "" And Trim(Me.TextBox28.Text) = "" Then
            error_msg += "總坪數平方公尺不能為空白 \n"
        End If

        If Me.TextBox33.Text <> "" And Trim(Me.TextBox32.Text) = "" Then
            error_msg += "庭院坪數平方公尺不能為空白 \n"
        End If

        '1040417修正-取消該判斷..
        'If Me.TextBox35.Text <> "" And Trim(Me.TextBox34.Text) = "" Then
        '    error_msg += "增建坪數平方公尺不能為空白 \n"
        'End If


        Return error_msg

    End Function

    Function 驗證數字是否阿拉伯數字() As Integer
        Dim strC As String = "", compare As String = ""
        Dim j As Integer = 0

        Dim 錯誤訊息 As String = ""

        For maxC As Integer = 0 To Me.Controls.Count - 1
            For Each minC As Control In Me.Controls(maxC).Controls
                If minC.GetType.Name = "HtmlInputText" Then
                    compare = CType(minC, HtmlInputText).Value
                    For i = 1 To Len(compare)
                        If IsNumeric(Mid(compare, i, 1)) = True Then
                            If Not (Hex(Asc(Mid(compare, i, 1))).Length = 2) Then
                                錯誤訊息 += compare & ","
                                j += 1
                            End If
                        End If
                    Next
                ElseIf minC.GetType.Name = "TextBox" Then
                    compare = CType(minC, TextBox).Text
                    For i = 1 To Len(compare)
                        If IsNumeric(Mid(compare, i, 1)) = True Then
                            If Not (Hex(Asc(Mid(compare, i, 1))).Length = 2) Then
                                錯誤訊息 += compare & ","
                                j += 1
                            End If
                        End If
                    Next
                End If
            Next
        Next

        If j > 0 Then
            Dim script As String = ""
            script += "alert('請檢查數字是否為半型阿拉伯數字 輸入內容為：" & 錯誤訊息 & " 內容有誤');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        End If
        Return j
    End Function

    '儲存
    Protected Sub ImageButton1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click

        '判斷輸入表單編號是否正確-----------------------------------------
        Dim formtype As String
        Dim objecttype As String
        Dim script As String = ""
        '物件用途是否為土地
        If DropDownList3.SelectedValue = "其他" And TextBox29.Text.Trim <> "" Then
            objecttype = Right(TextBox29.Text.Trim, 1)
        Else
            objecttype = Right(DropDownList3.SelectedValue, 1)
        End If


        If objecttype = "地" Then
            If ddl契約類別.SelectedValue = "專任" Then
                formtype = "1233,1201,2201,1243"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                'formtype = "1235" 因目前暫無故會以1229代
                formtype = "1235,1229,2229,1242"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                formtype = "1240,1231"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                formtype = "0000"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                formtype = "9999"
            End If
        Else
            If ddl契約類別.SelectedValue = "專任" Then  '1246 直營專任約
                formtype = "1201,2201,1243"
            ElseIf ddl契約類別.SelectedValue = "一般" Then '1246 直營一般約
                formtype = "1229,2229,1242"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                formtype = "1231"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                formtype = "0000"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                formtype = "9999"
            End If
        End If


        Dim message As String = formnocheck.checkform(formtype, TextBox2.Text, store.SelectedValue)

        If Len(message) > 4 Then
            script += "alert('" & message & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            Exit Sub
        End If
        Dim amessage As String = ""
        If IsNumeric(TextBox13.Text.ToString) Or TextBox13.Text.Trim = "" Then '房
            If InStr(TextBox13.Text.ToString.ToLower, "+") = 0 And InStr(TextBox13.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox13.Text.ToString.ToLower, "-") = 0 And InStr(TextBox13.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox13.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－房僅可填入數字\n"
                    End If
                Else
                    If C1.Checked = True Then
                    Else
                        amessage &= "建物基本資料－房僅可填入數字\n"
                    End If

                End If
            Else
                amessage &= "建物基本資料－房僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－房僅可填入數字\n"
        End If

        If IsNumeric(TextBox14.Text.ToString) Or TextBox14.Text.Trim = "" Then '廳
            If InStr(TextBox14.Text.ToString.ToLower, "+") = 0 And InStr(TextBox14.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox14.Text.ToString.ToLower, "-") = 0 And InStr(TextBox14.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox14.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－廳僅可填入數字\n"
                    End If
                Else
                    amessage &= "建物基本資料－廳僅可填入數字\n"
                End If
            Else
                amessage &= "建物基本資料－廳僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－廳僅可填入數字\n"
        End If

        If IsNumeric(TextBox15.Text.ToString) Then '衛
            If InStr(TextBox15.Text.ToString.ToLower, "+") = 0 And InStr(TextBox15.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox15.Text.ToString.ToLower, "-") = 0 And InStr(TextBox15.Text.ToString.ToLower, "－") = 0 Then
                Else
                    If TextBox15.Text.Trim = "" Then
                    Else
                        amessage &= "建物基本資料－衛僅可填入數字\n"
                    End If
                End If
            Else
                amessage &= "建物基本資料－衛僅可填入數字\n"
            End If
        Else
            If IsNumeric(TextBox15.Text.Replace(".", "")) Then '允許小數點
            Else
                If TextBox15.Text.Trim = "" Then
                Else
                    amessage &= "建物基本資料－衛僅可填入數字\n"
                End If
            End If
        End If

        If IsNumeric(TextBox16.Text.ToString) Or TextBox16.Text.Trim = "" Then '室
            If InStr(TextBox16.Text.ToString.ToLower, "+") = 0 And InStr(TextBox16.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox16.Text.ToString.ToLower, "-") = 0 And InStr(TextBox16.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox16.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－室僅可填入數字\n"
                    End If
                Else
                    amessage &= "建物基本資料－室僅可填入數字\n"
                End If
            Else
                amessage &= "建物基本資料－室僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－室僅可填入數字\n"
        End If

        If DropDownList3.Text <> "土地" Then
            If TextBox88.Text = "" Or TextBox89.Text = "" Or TextBox90.Text = "" Then
                amessage &= "當不為土地時，地上、地下、所在樓層 不可為空 \n"
            End If
            If Trim(Text2.Value) = "" Then
                amessage &= "當不為土地時，完工年月 不可為空 \n"
            Else
                If Trim(Text2.Value) <> "00000" Then
                    If Not IsDate(Left(Text2.Value, 3) + 1911 & "/" & Mid(Text2.Value, 4, 2) & "/" & "01") Then
                        amessage &= "完工年月輸入錯誤 \n"
                    End If
                End If
            End If
        End If
        If Trim(Text11.Value) <> "" Then
            If Not IsDate(Left(Text11.Value, 3) + 1911 & "/" & Mid(Text11.Value, 4, 2) & "/" & Mid(Text11.Value, 6, 2)) Then
                amessage &= "登記日期輸入錯誤 \n"
            End If
        End If
        If Trim(Date2.Text) <> "" Then
            If Not IsDate(Left(Date2.Text, 3) + 1911 & "/" & Mid(Date2.Text, 4, 2) & "/" & Mid(Date2.Text, 6, 2)) Then
                amessage &= "委託起始日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託起始日期 \n"
        End If
        If Trim(Date3.Text) <> "" Then
            If Not IsDate(Left(Date3.Text, 3) + 1911 & "/" & Mid(Date3.Text, 4, 2) & "/" & Mid(Date3.Text, 6, 2)) Then
                amessage &= "委託截止日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託截止日期 \n"
        End If
        If Trim(Date2.Text) = "" And Trim(Date3.Text) = "" Then

        Else
            If CType(Left(Trim(Date3.Text), 7), Integer) < CType(Left(Trim(Date2.Text), 7), Integer) Then
                amessage &= "委託截止日須在委託起始日之後 \n"
            End If
        End If

        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Or RadioButton2.Checked = True Then

            Else
                amessage &= "消費者是否願意提供個資 要強制選擇 \n"
            End If
        End If

        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "透天" Then

        Else
            If TextBox91.Text = "" Or TextBox92.Text = "" Then
                amessage &= "非土地時，每層戶數及電梯數不可為空 \n"
            End If
        End If

        If amessage = "" Then

        Else
            eip_usual.Show(amessage)
            Exit Sub
        End If
        '使用平方公尺欄位不得為空值---------------------------------------
        Dim Str As String = chk_平方公尺()
        If Str <> "" Then
            script += "alert('" & Str & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            Exit Sub
        End If

        '驗證數字是否阿拉伯數字()------------------------------------------
        Dim j As Integer = 驗證數字是否阿拉伯數字()
        If j > 0 Then
            Exit Sub
        End If

        Dim sid As String = Request("sid")
        Dim 物件編號 As String = ""
        '1010630 by佩嬬
        If ddl契約類別.SelectedValue = "專任" Then
            物件編號 = "1"
        ElseIf ddl契約類別.SelectedValue = "一般" Then
            物件編號 = "6"
        ElseIf ddl契約類別.SelectedValue = "同意書" Then
            物件編號 = "7"
        ElseIf ddl契約類別.SelectedValue = "流通" Then
            物件編號 = "5"
        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            物件編號 = "9"
        End If

        If store.SelectedValue = "請選擇" Then
            物件編號 &= Mid(sid, 2) & TextBox2.Text.Trim
        Else
            物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
        End If
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        Label57.Text = ""
        If Trim(TextBox267.Text) <> "" Then
            Dim sql As String = ""
            sql = "select 物件編號,磁扣編號,委託截止日,銷售狀態 from 委賣物件資料表 where 磁扣編號='" & Trim(TextBox267.Text) & "' and 物件編號 <> '" & 物件編號 & "'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "t1")
            table1 = ds.Tables("t1")
            If table1.Rows.Count > 0 Then
                If Not IsDBNull(table1.Rows(0).Item("委託截止日")) Then
                    If table1.Rows(0).Item("委託截止日") < sysdate Then
                        sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                        cmd = New SqlCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                    ElseIf table1.Rows(0).Item("銷售狀態") = "已成交" Then
                        sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                        cmd = New SqlCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                    ElseIf table1.Rows(0).Item("銷售狀態") = "已解約" Then
                        sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                        cmd = New SqlCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                    Else
                        script += "alert('磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對');"
                        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                        conn.Close()
                        conn = Nothing
                        Exit Sub
                    End If
                Else
                    script += "alert('磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對');"
                    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                    conn.Close()
                    conn = Nothing
                    Exit Sub
                End If
            Else
                sql = "select 物件編號,磁扣編號,委託截止日,租賃狀態 from 委租物件資料表 where 磁扣編號='" & Trim(TextBox267.Text) & "' and 物件編號 <> '" & 物件編號 & "'"
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "t1")
                table1 = ds.Tables("t1")
                If table1.Rows.Count > 0 Then
                    If Not IsDBNull(table1.Rows(0).Item("委託截止日")) Then
                        If table1.Rows(0).Item("委託截止日") < sysdate Then
                            sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        ElseIf table1.Rows(0).Item("租賃狀態") = "已成交" Then
                            sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        ElseIf table1.Rows(0).Item("租賃狀態") = "已解約" Then
                            sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        Else
                            script += "alert('磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對');"
                            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                            conn.Close()
                            conn = Nothing
                            Exit Sub
                        End If
                    Else
                        script += "alert('磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對');"
                        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                        conn.Close()
                        conn = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If

        '新增-物件資料
        新增記錄()
        UpdateAIVoiceOver(物件編號, store.SelectedValue)
        If trans = "True" Then
            '新增-不動產說明書
            '該FUNCTION含新增+修改(會自行判斷)
            新增不動產說明書()

            '新增-面積細項
            面積細項_判斷編號是否相同()

            '新增-他項權利細項
            他項權利_判斷編號是否相同()

            '新增多筆土增稅
            土增稅_判斷編號是否相同()
        Else
            script += "alert('新增失敗');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            Exit Sub
        End If
        conn.Close()
        conn = Nothing
        '確定資料新增成功，長出其他功能列
        If trans = "True" Then
            '-------------------------------------------------------------------------------------------------------------------

            ''Dim exist As String = chk_exist(store.SelectedValue, Label57.Text)
            ''If exist <> "True" Then
            If Left(Label57.Text, 1) <> "9" Then
                '判斷是否有購買廣告版位
                Dim torf As String = web_no(store.SelectedValue)
                Dim num As String = index_num(Label57.Text, store.SelectedValue)

                Dim url As String = ""
                If torf = "True" Then
                    url = "https://home.etwarm.com.tw/sale-" & Trim(num)
                ElseIf torf = "False" Then
                    url = "https://www.etwarm.com.tw/sale-" & Trim(num)
                End If

                '寫入資料表
                ''20190910 10.40.20.66先行拿掉==================================
                ''voice_objects("insert", store.SelectedValue, Label57.Text, url, 1)
                ''20190910 10.40.20.66先行拿掉==================================
            End If
            ''    '---------------------------------------------------------------------------------------------------------------------
            ''End If
            Dim redirectUrl As String = "Obj_Update_V4.aspx?state=update&oid=" & Label57.Text & "&sid=" & store.SelectedValue & "&src=NOW&from=Upd&trans=true"
            Dim currentStoreId As String = ""
            If Not Request.Cookies("store_id") Is Nothing Then
                currentStoreId = Request.Cookies("store_id").Value
            End If

            If Request("state") = "add" AndAlso RadioButtonList2.SelectedValue = "售件" AndAlso IsICBAutoUpdateEnabled(currentStoreId) Then
                Dim redirectScript As String = "updateICBObjectInfoAndRedirect('" & JsEncode(Label57.Text) & "','" & JsEncode(store.SelectedValue) & "','1','" & JsEncode(redirectUrl) & "');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "ICBAutoUpdateAfterAdd", redirectScript, True)
            Else
                Response.Redirect(redirectUrl)
            End If
        End If

    End Sub

    Function chk_exist(ByVal sid As String, ByVal oid As String) As String
        Dim torf As String
        Dim conn_voice As MySqlConnection = New MySqlConnection(mysqlegoupload)
        If conn_voice.State = ConnectionState.Closed Then
        Else
            conn_voice.Open()

            sql = "Select * from voice_objects where sid='" & sid & "' and oid='" & oid & "' "


            Dim mycmd = New MySqlCommand(sql, conn_voice)

            Dim mydr As MySqlDataReader = mycmd.ExecuteReader
            If mydr.Read Then
                torf = "True"
            Else
                torf = "False"
            End If


            conn_voice.Close()
            conn_voice.Dispose()

            Return torf
        End If



    End Function

    '產調區塊(START)-----------------------------------------------------------------------------------------------------------------------------------
    Protected Sub DropDownList48_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList48.SelectedIndexChanged
        If DropDownList48.SelectedValue.Trim = "其他" Then
            TextBox243.Visible = True
        Else
            TextBox243.Text = ""
            TextBox243.Visible = False
        End If
    End Sub

    Protected Sub DropDownList49_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList49.SelectedIndexChanged
        If DropDownList49.SelectedValue.Trim = "其他" Then
            TextBox244.Visible = True
        Else
            TextBox244.Text = ""
            TextBox244.Visible = False
        End If
    End Sub


    '1050506表單控管
    Function check_formno(ByVal formno As String, ByVal sid As String, ByVal oid As String, ByVal oidname As String, ByVal state As String) As String
        Dim fun As String = System.IO.Path.GetFileName(Request.PhysicalPath)
        Dim status As String = ""
        Dim used As String = ""
        Dim sqlm As String = ""
        Dim sqlu As String = ""
        sql = "select * from "
        sqlm = "insert into " '新增至XX管制檔_管理
        sqlu = "update " '更新秘書XX管制檔
        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            'Response.Write("<br>" & fun & "<br>")
            If fun = "Rent_Obj_Add.aspx" Or fun = "Rent_Obj_Copy.aspx" Or fun = "Rent_Obj_Update.aspx" Then '租委託1202
                sql &= "(select 店代號, 合約編號 as 編號,上傳註記 from 秘書_合約管制檔可用 where 店代號='" & sid & "' and 合約編號='" & formno & "'"
                If Request("state") = "update" Or Request("state") = "copy" Then
                    sql &= ")"
                Else
                    sql &= " and 上傳註記<>'D')"
                End If
                If Left(formno, 1) = "F" Then
                    status = "OK"
                    sqlm &= "秘書_合約管制檔_管理 "
                    sqlu &= "秘書_合約管制檔可用 set "
                Else
                    status = "租賃"
                End If
                If status = "OK" Then
                    sqlm &= " (類別, 店代號, 編號, 人員, 日期, 備註, 異動人,物件編號) VALUES "
                    If state = "新增" Or state = "修改" Then
                        sqlm &= "('使用','"
                        sqlu &= " 處理類別='使用',上傳註記='U',"
                    ElseIf state = "刪除" Then
                        sqlm &= "('作廢','"
                        sqlu &= " 處理類別='作廢',上傳註記='A', "
                    End If
                    sqlm &= sid & "','" & formno & "','" & sale1.SelectedValue & "','" & sysdate & "','物件:" & oidname & "," & state & "後直接寫入,使用店代號為:" & sid & "','" & Request.Cookies("webfly_empno").Value & "','" & oid & "')"
                    sqlu &= " 經紀人代號='" & sale1.SelectedValue & "', 備註='物件" & state & "後直接寫入,使用店代號為:" & sid & "', 修改日期='" & sysdate & "' where 店代號='" & sid & "' and "
                    sqlu &= "合約編號='" & formno & "'"
                Else '非出貨合約書編號
                    sqlm = ""
                    sqlu = ""
                End If
            ElseIf fun = "Obj_Add_V4.aspx" Or fun = "Obj_Copy_V3.aspx" Or fun = "Obj_Update_V4.aspx" Then
                If state = "新增" Or state = "刪除" Then
                    sql &= "(select 店代號, 合約編號 as 編號,上傳註記 from 秘書_合約管制檔可用 where 店代號='" & sid & "' and 合約編號='" & formno & "'"
                Else
                    '1050524修
                    If state = "修改" And Left(formno, 1).ToUpper = "X" Then
                        sql &= "(select 店代號, 合約編號 as 編號,上傳註記 from 秘書_合約管制檔可用 where 店代號='" & sid & "' and 合約編號='" & formno & "'"
                    Else
                        sql &= "(select 店代號, 合約編號 as 編號,上傳註記 from 秘書_合約管制檔已用 where 店代號='" & sid & "' and 合約編號='" & formno & "'"
                    End If

                End If

                If Request("state") = "update" Or Request("state") = "copy" Then
                    sql &= ")"
                Else
                    sql &= " and 上傳註記<>'D')"
                End If
                If Left(formno, 1) = "A" Or Left(formno, 1) = "B" Or Left(formno, 1) = "C" Or Left(formno, 1) = "M" Or Left(formno, 1) = "L" Or Left(formno, 1) = "N" Or Left(formno, 1) = "X" Then
                    status = "OK"
                    sqlm &= "秘書_合約管制檔_管理 "
                    'sqlu &= "秘書_合約管制檔可用 set "
                    '1050524修
                    If state = "新增" Or Left(formno, 1) = "X" Then
                        sqlu &= "秘書_合約管制檔可用 set "
                    ElseIf state = "修改" Then
                        sqlu &= "秘書_合約管制檔已用 set "
                    End If
                Else
                    status = "買賣"
                End If
                If status = "OK" Then
                    sqlm &= " (類別, 店代號, 編號, 人員, 日期, 備註, 異動人,物件編號) VALUES "
                    If state = "新增" Or state = "修改" Then
                        sqlm &= "('使用','"
                        sqlu &= " 處理類別='使用',上傳註記='U',"
                    ElseIf state = "刪除" Then
                        sqlm &= "('作廢','"
                        sqlu &= " 處理類別='作廢',上傳註記='D', "
                    End If
                    sqlm &= sid & "','" & formno & "','" & sale1.SelectedValue & "','" & sysdate & "','物件:" & oidname & "," & state & "後直接寫入,使用店代號為:" & sid & "','" & Request.Cookies("webfly_empno").Value & "','" & oid & "')"
                    sqlu &= " 經紀人代號='" & sale1.SelectedValue & "', 備註='物件" & state & "後直接寫入,使用店代號為:" & sid & "', 修改日期='" & sysdate & "' where 店代號='" & sid & "' and "
                    sqlu &= "合約編號='" & formno & "'"
                Else '非出貨合約書編號
                    sqlm = ""
                    sqlu = ""
                End If
            End If
            sql &= " as a "
            'Response.Write("<br>" & sqlm & "<br>")
            'Response.Write("<br>" & sqlu & "<br>")
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            Try
                adpt.Fill(ds, "table1")
                Dim table1 As DataTable = ds.Tables("table1")
                If table1.Rows.Count > 0 Then '沒有被刪除             

                    Dim com As New SqlCommand(sqlm, conn)
                    Try
                        If state = "修改" Then
                            If (Label59.Text.Trim <> sale1.SelectedValue.Trim) Then '舊值 新值
                                com.ExecuteNonQuery()
                            End If
                        Else
                            com.ExecuteNonQuery()
                        End If


                        'Response.Write("<br>sqlm:" & sqlm & "<br>")
                    Catch ex As Exception
                        'Response.Write("<br>sqlm:" & ex.ToString & "<br>")
                        'Response.Write(selstr)
                    End Try

                    Dim com1 As New SqlCommand(sqlu, conn)
                    Try
                        If state = "修改" Then
                            If (Label59.Text.Trim <> sale1.SelectedValue.Trim) Then '舊值 新值
                                com1.ExecuteNonQuery()
                            End If
                        Else
                            com1.ExecuteNonQuery()
                        End If
                        'com1.ExecuteNonQuery()
                        'Response.Write("<br>sqlu:" & sqlu & "<br>")
                    Catch ex As Exception
                        'Response.Write("<br>sqlu:" & ex.ToString & "<br>")
                    End Try
                    used = "use"
                Else '管制檔裡沒有
                    used = "no"

                End If
            Catch ex As Exception
                ' Response.Write(sql)
                used = "no"
                'Response.End()
            End Try


        End Using
        Return used
    End Function




    '產調區塊(END)-----------------------------------------------------------------------------------------------------------------------------------

    Public Sub 新增記錄()

        Dim sid As String = Request("sid")

        Dim 物件編號 As String = ""
        Dim 使用分區 As String = ""
        Dim formtype As String = ""

        '1010630 by佩嬬
        If ddl契約類別.SelectedValue = "專任" Then
            物件編號 = "1"
        ElseIf ddl契約類別.SelectedValue = "一般" Then
            物件編號 = "6"
        ElseIf ddl契約類別.SelectedValue = "同意書" Then
            物件編號 = "7"
        ElseIf ddl契約類別.SelectedValue = "流通" Then
            物件編號 = "5"
        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            物件編號 = "9"
        End If

        If store.SelectedValue = "請選擇" Then
            物件編號 &= Mid(sid, 2) & TextBox2.Text.Trim
        Else
            物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
        End If

        Dim count As Integer = 0
        Dim script, message As String
        script = ""
        message = ""

        If Len(Left(Trim(Date2.Text), 7)) <> 7 Or Len(Left(Trim(Date3.Text), 7)) <> 7 Then
            message &= "委託期間格式輸入有誤，例：1030101\n"
        End If

        '同一家店中是否已有相同的物件編號-- 20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
        sql = " Select 物件編號 From 委賣物件資料表 With(NoLock) "
        sql &= " Where 物件編號 = '" & 物件編號 & "' "
        sql &= " and 店代號 = '" & store.SelectedValue & "' "
        sql &= " union all "
        sql &= " Select 物件編號 From 委賣物件過期資料表 With(NoLock) "
        sql &= " Where 物件編號 = '" & 物件編號 & "' "
        sql &= " and 店代號 = '" & store.SelectedValue & "' "



        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")



        If table1.Rows.Count > 0 Then
            message &= "已有相同的物件編號(表單編號已使用)\n"
        Else
            If Left(物件編號, 1) <> "9" Then '若為庫存則不檢查單號
                '1010630 by佩嬬判斷是否為多店,再判斷是否為同一法人 
                If myobj.Objectmstore = "1" Then

                    If formnocheck.form_own(clspowerset.mstoreid_OLD(myobj.mstore_id, store.SelectedValue), UCase(TextBox2.Text)) = "False" Then
                        message &= "此編號不為所購買表單編號\n"
                    End If
                Else
                    If formnocheck.form_own("'" & store.SelectedValue.ToString & "'", UCase(TextBox2.Text)) = "False" Then
                        message &= "此編號不為所購買表單編號\n"
                    End If
                End If
            End If
        End If


        If Left(物件編號, 1) = "6" Then
            '同一家店中是否有和委租有相同的編號
            sql = "Select * From 委租物件資料表 With(NoLock) "
            sql &= "Where 物件編號 = '" & 物件編號 & "' "
            sql &= "and 店代號 = '" & store.SelectedValue & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            If table1.Rows.Count > 0 Then
                message &= "委租已有相同的物件編號，請更換另一個編號\n"
            End If

        End If

        If TextBox102.Text <> "" And CheckBox2.Checked = False Then
            If Replace(TextBox102.Text.ToString, vbNewLine, "").Length > 1000 Then
                message &= "訴求重點字數上限為1000字\n"
            ElseIf Replace(TextBox102.Text.ToString, vbNewLine, "").Length < 20 Then
                message &= "訴求重點字數至少輸入20字\n"
            End If
        End If

        '鑰匙編號不可有'
        If InStr(input66.Value, "'") > 0 Then
            message &= "鑰匙編號有不合法的字元'\n"
        End If


        '建物主要用途  
        If TextBox4.Text.Trim <> "" Then
            sql = "Select DISTINCT 名稱 From 不動產說明書_物件用途 With(NoLock) "
            sql &= "Where 名稱 = '" & TextBox4.Text.Trim & "' "
            sql &= " and (店代號 in ('A0001','" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "') "
            sql &= " or 店代號 in "
            sql &= " (select 店代號 from HSSTRUCTURE "
            sql &= " where 組別 in (select 組別 from HSSTRUCTURE where 店代號='" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "'))) "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            'If Request.Cookies("webfly_empno").Value = "92H" Then
            '    Response.Write(sql)
            '    Exit Sub
            'End If
            If table1.Rows.Count > 0 Then
                message &= "此建物主要用途已存在,請下拉選擇\n"
                TextBox4.Text = ""
            Else
                If Len(TextBox4.Text.Trim) < 50 Then
                    Dim sql = "insert into 不動產說明書_物件用途 (名稱,店代號) values ("
                    sql &= "'" & TextBox4.Text.Trim & "' , '" & IIf(sid = "", store.SelectedValue, sid) & "')"
                    cmd = New SqlCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                Else
                    message &= "此建物主要用途超過長度50個字,請修正\n"
                End If
            End If
        End If

        '完整地址, 

        'address1完整地址,address2部份地址-到"弄"
        Dim address1 As String = "", address2 As String = ""

        '部分地址
        If add1.Text <> "" Then address2 &= add1.Text & zone3.SelectedValue
        If add2.Text <> "" Then address2 &= add2.Text & "鄰"
        If add3.Text <> "" Then address2 &= add3.Text & address20.SelectedValue
        '1040519修正
        'If add4.Text <> "" Then address2 &= add4.Text & "段"
        If add4.Text <> "" Then address2 &= add4.Text & Label64.Text

        If add5.Text <> "" Then address2 &= add5.Text & "巷"
        If add6.Text <> "" Then address2 &= add6.Text & "弄"

        address1 &= address2
        '1040519修正
        'If add7.Text <> "" Then address1 &= add7.Text & "號"
        If Label64.Text = "小段" Then
            If add7.Text <> "" Then address1 &= add7.Text & "地號"
        Else
            If add7.Text <> "" Then address1 &= add7.Text & "號"
        End If
        If add8.Text <> "" Then address1 &= "之" & add8.Text

        '20100607小豪修正("之"後方若有值,則"樓"之前加入空白,避免距離過近造成誤會,ex:101號之1   3樓)----
        If add9.Text <> "" Then
            If add8.Text <> "" Then
                address1 &= "   " & add9.Text & "樓"
            Else
                address1 &= add9.Text & "樓"
            End If
        End If
        '--------------------------------------------------------------------------------------------------
        If add10.Text <> "" Then address1 &= "之" & add10.Text

        If 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")

                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")

                    If table1.Rows.Count > 0 Then
                        For j = 0 To table1.Rows.Count - 1
                            sql = " update 委賣物件資料表 "
                            sql += " set 聯賣='N' "
                            sql += " where 物件編號 = '" & table1.Rows(j)("物件編號").ToString.Trim & "' and 店代號='" & table1.Rows(j)("店代號").ToString.Trim & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        Next
                        'message &= "此區域已存在相同物件\n"
                    End If
                End If
            End If
        End If

        '2017.01.05 by Finch 聯賣規則
        If Not 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")

                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")

                    If table1.Rows.Count > 0 Then
                        message &= "此區域已存在相同物件\n"
                    End If
                Else
                    If table1.Rows(0)("區域名稱").ToString.Trim = "宜蘭區聯賣" Then
                        Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")

                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")

                        If table1.Rows.Count > 0 Then
                            message &= "此區域已存在相同物件\n"
                        End If
                    ElseIf table1.Rows(0)("區域名稱").ToString.Trim = "高屏區聯賣" Then
                        Dim kanPingSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")

                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & kanPingSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")

                        'Response.Write(sql)

                        If table1.Rows.Count > 0 Then
                            Dim queryOid As String = ""
                            For i = 0 To table1.Rows.Count - 1
                                queryOid &= table1.Rows(i)("物件編號") & "', '"
                            Next

                            sql2 = " Select TOP 1 *,convert(char, [借用時間], 111) AS 借用時間西元 From 物件鑰匙管理 With(NoLock) "
                            sql2 &= " Where 物件編號 in ('" & queryOid & "') AND 借用店 = '" & store.SelectedValue & "' "
                            sql2 &= " ORDER BY num DESC "
                            adpt = New SqlDataAdapter(sql2, conn)
                            ds = New DataSet()
                            adpt.Fill(ds, "table2")
                            table2 = ds.Tables("table2")

                            'Response.Write(sql2)

                            If table2.Rows.Count > 0 Then
                                message &= "此區域已存在相同物件，物件編號：" & table2.Rows(0)("物件編號") & "\n"
                                If table2.Rows(0)("借用項目") = "O" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾點閱此物件物調表。\n"
                                ElseIf table2.Rows(0)("借用項目") = "M" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件不動產說明書。\n"
                                ElseIf table2.Rows(0)("借用項目") = "K" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件鑰匙。\n"
                                ElseIf table2.Rows(0)("借用項目") = "C" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾電話詢問此物件。\n"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        '2017.01.19 by Finch 同群組物件不重複        
        sql = "Select * From HSSTRUCTURE With(NoLock) "
        sql &= "Where 店代號 = '" & store.SelectedValue & "' AND 同群物件不重覆 = 'Y'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        Dim groupSid As String = ""

        If table1.Rows.Count > 0 Then
            Dim groupId As String = table1.Rows(0)("組別").ToString.Trim
            sql = "Select * From HSSTRUCTURE With(NoLock) "
            sql &= "Where 組別 = '" & groupId & "'"

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                For i = 0 To table1.Rows.Count - 1
                    groupSid &= table1.Rows(i)("店代號").ToString.Trim & "','"
                Next

                sql = "Select * From 委賣物件資料表 With(NoLock) "
                sql &= "Where 店代號 in ('" & groupSid & "') AND 店代號 <> '" & store.SelectedValue & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' AND 註記 = 'Y' "
                sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                If table1.Rows.Count > 0 Then
                    message &= "同群組已存在相同物件，物件編號：" & table1.Rows(0)("物件編號") & "\n"
                End If

            End If
        End If

        '車位類別
        'If TextBox37.Text.Trim <> "" Then
        '    sql = "Select * From 資料_車位 With(NoLock) "
        '    sql &= "Where 車位 = '" & TextBox37.Text.Trim & "' and 店代號 in ('A0001','" & store.SelectedValue & "') "
        '    adpt = New SqlDataAdapter(sql, conn)
        '    ds = New DataSet()
        '    adpt.Fill(ds, "table1")
        '    table1 = ds.Tables("table1")
        '    Dim flag1 As String = ""
        '    If table1.Rows.Count > 0 Then
        '        message &= "此車位類別已存在,請下拉選擇\n"
        '        TextBox37.Text = ""
        '    Else
        '        If Len(TextBox37.Text.Trim) < 5 Then
        '            Dim sql2 = "insert into 資料_車位 (車位,店代號) values ("
        '            sql2 &= "'" & TextBox37.Text.Trim & "' , '" & Request.Cookies("store_id").Value & "')"
        '            cmd = New SqlCommand(sql2, conn)
        '            cmd.ExecuteNonQuery()
        '            cmd.Dispose()
        '        Else
        '            message &= "此新增車位類別超過長度5個字,請修正\n"
        '        End If

        '    End If
        'End If

        If message <> "" Then
            script += "alert('" & message & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            '判斷失敗OR成功的參數
            trans = "False"
            Exit Sub
        End If

        '2016.06.21 by Finch 從"委賣物件資料表_面積細項"抓第一筆土地面積的使用分區，
        '寫回"委賣物件資料表"的"物件用途"欄位(前台抓使用分區的欄位)
        'If Request.Cookies("webfly_empno").Value = "92H" Then
        sql = "Select TOP 1 * From 委賣物件資料表_面積細項 With(NoLock) "
        sql &= "Where 物件編號 = '" & 物件編號 & "' and 店代號 = '" & store.SelectedValue & "' and 項目名稱 = '土地面積' and 使用分區 IS NOT NULL and 使用分區 <> '' ORDER BY 流水號 "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count > 0 Then
            使用分區 = table1.Rows(0)("使用分區")
        End If
        'End If



        Dim splitarray As Array
        Dim message1 As String = ""

        '左邊單位為平方公尺

        sql = "Insert into 委賣物件資料表 ( "
        sql &= "物件編號,建物型態,物件用途,物件類別,郵遞區號,商圈代號,完整地址,縣市,鄉鎮市區,村里,村里別, "
        sql &= "鄰,路名,路別,段,巷,弄,號,之,所在樓層,樓之,刊登售價,主建物,主建物平方公尺,附屬建物,附屬建物平方公尺, "
        sql &= "公共設施,公共設施平方公尺,公設內車位坪數,公設內車位平方公尺,地下室,地下室平方公尺, "
        sql &= "總坪數,總平方公尺,土地坪數,土地平方公尺,庭院坪數,庭院平方公尺,加蓋坪數,加蓋平方公尺,房,廳,衛,建築名稱,張貼卡案名,地上層數,地下層數, "
        sql &= "每層戶數,每層電梯數,竣工日期,座向,管理費,含公設車位坪數,車位說明,車位坪數,車位平方公尺, "
        sql &= "訴求重點,店代號,經紀人代號,營業員代號1,營業員代號2,委託起始日,委託截止日,公設比,銷售狀態, "
        sql &= "輸入者,新增日期,修改日期,修改時間,網頁刊登,強銷物件,鑰匙編號,帶看方式,土地標示,建號,上傳註記,委託售價,車位價格, "
        sql &= "車位租售,國小,國中,車位數量,國小1,國中1,部份地址,捷運,大專院校,其他學校,公車站名,銷售樓層, "
        sql &= "室,備註,公園代號,高中,高中1, "
        sql &= "管理費單位,管理費繳交方式,車位管理費類別,車位管理費,車位管理費單位,大樓代號,register_date,網站總坪數,棟別,臨路寬,面寬,縱深,磁扣編號,其他使用分區,ezcode,增建,Land_FileNo "
        '資料來源
        If Request("source") = "land" Or Request("source") = "sland" Or Request("source") = "eland" Then
            sql &= ",資料來源"
        End If
        sql &= ",合約終止,提供個資,聯賣,社區養寵,養貓,養狗,聯賣日期 "
        sql &= ") Values ( "



        '物件編號,
        sql &= "'" & UCase(Microsoft.VisualBasic.Strings.StrConv(物件編號, Microsoft.VisualBasic.VbStrConv.Narrow)) & "' , "

        '物件主要用途
        If TextBox4.Text <> "" Then
            sql &= "'" & TextBox4.Text & "' , "
        ElseIf DropDownList19.SelectedValue <> "" Then
            sql &= "'" & DropDownList19.SelectedValue & "' , "
        Else
            sql &= "'' , "
        End If

        '2016.06.21 by Finch 使用分區寫入"物件用途"欄位
        '使用分區
        If 使用分區 <> "" And 使用分區 <> "請選擇" Then
            sql &= "'" & 使用分區 & "' , "
        Else
            sql &= "'' , "
        End If


        '物件型態-物件類別,        
        sql &= "'" & DropDownList3.SelectedValue & "' , "


        '郵遞區號, 
        sql &= "'" & TB_AreaCode.Text & "' , "

        '商圈代號, 
        If Trim(TextBox96.Text) <> "" Then
            'Dim m As Integer
            'Dim 商圈代號 As String = ""
            'Dim sql1 As String = ""

            'Dim item1 As Array = Split(Page.Request.Form(TextBox96.UniqueID), ".")
            'Dim temp As String = ""
            'i = 0
            'For i = 0 To item1.Length - 1
            '    If item1(i) <> "" Then
            '        temp &= "'" & item1(i) & "',"
            '        m += 1
            '    End If
            'Next i
            'Dim cc As String = ""
            'cc = Mid(temp, 1, Len(temp) - 1)
            'i = 0
            'For i = 0 To m - 1
            '    sql1 = "Select 商圈代號 FROM 精華生活圈資料表 where 商圈名稱 in (" & cc & ") and 核准否='1'"
            '    adpt = New SqlDataAdapter(sql1, conn)
            '    ds = New DataSet()
            '    adpt.Fill(ds, "table4")
            '    table4 = ds.Tables("table4")
            '    If table4.Rows.Count > 0 Then
            '        If IsDBNull(table4.Rows(i)("商圈代號")) = False Or table4.Rows(i)("商圈代號") <> "" Then
            '            商圈代號 &= table4.Rows(i)("商圈代號") & "."
            '        End If
            '    End If
            'Next i

            '修正代號改以隱藏的TEXTBOX250.TEXT
            sql &= "'" & TextBox250.Text & "' , "
        Else
            sql &= "'',"
        End If




        sql &= "N'" & address1 & "' , "
        '縣市, 
        sql &= "'" & DDL_County.SelectedValue & "' , "
        '鄉鎮市區, 
        sql &= "'" & DDL_Area.SelectedItem.Text & "' , "
        '村里, 
        sql &= "'" & add1.Text & "' , "
        '村里別,
        sql &= "'" & zone3.SelectedValue & "' , "
        '鄰,
        sql &= "'" & add2.Text & "' , "
        '路名,
        sql &= "N'" & add3.Text & "' , "
        '路別, 
        sql &= "N'" & address20.SelectedValue & "' , "
        '段, 
        sql &= "'" & add4.Text & "' , "
        '巷, 
        sql &= "'" & add5.Text & "' , "
        '弄, 
        sql &= "'" & add6.Text & "' , "
        '號, 
        sql &= "'" & add7.Text & "' , "
        '之
        sql &= "'" & add8.Text & "' , "
        '所在樓層,-住址的"樓"存在"所在樓層"欄位
        sql &= "'" & add9.Text & "' , "
        '樓之, 
        sql &= "'" & add10.Text & "' , "


        '刊登售價num, 
        If TextBox12.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox12.Text & " , "
        End If



        '主建物num, 
        If TextBox6.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox6.Text & " , "
        End If
        '主建物平方公尺num, 
        If TextBox5.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox5.Text & " , "
        End If

        '附屬建物num, 
        If TextBox8.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox8.Text & " , "
        End If
        '附屬建物平方公尺num,
        If TextBox7.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox7.Text & " , "
        End If

        '公共設施num,
        If TextBox10.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox10.Text & " , "
        End If
        '公共設施平方公尺num, 
        If TextBox9.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox9.Text & " , "
        End If

        '公設內車位坪數num, 
        If TextBox23.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox23.Text & " , "
        End If
        '公設內車位平方公尺num, 
        If TextBox21.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox21.Text & " , "
        End If

        '地下室num, 
        If TextBox20.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox20.Text & " , "
        End If
        '地下室平方公尺,num
        If TextBox19.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox19.Text & " , "
        End If

        '總坪數,num
        If TextBox29.Text = "" Then
            sql &= "0 ,"
        Else
            sql &= TextBox29.Text & " , "
        End If
        '總平方公尺,num 
        If TextBox28.Text = "" Then
            sql &= "0 ,"
        Else
            sql &= TextBox28.Text & " , "
        End If

        '土地坪數, num
        If TextBox31.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox31.Text & " , "
        End If
        '土地平方公尺, num
        If TextBox30.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox30.Text & " , "
        End If

        '庭院坪數, num
        If TextBox33.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox33.Text & " , "
        End If
        '庭院平方公尺, num
        If TextBox32.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox32.Text & " , "
        End If

        '加蓋坪數, num
        If TextBox35.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox35.Text & " , "
        End If
        '加蓋平方公尺, num
        If TextBox34.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox34.Text & " , "
        End If

        '房, 
        If C1.Checked = True Then
            sql &= "'-1' , "
        Else
            sql &= "'" & TextBox13.Text & "' , "
        End If
        '廳, 
        sql &= "'" & TextBox14.Text & "' , "
        '衛, num
        If TextBox15.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox15.Text & " , "
        End If

        '建築名稱, 
        sql &= "N'" & input4.Value & "' , "
        '張貼卡案名, 
        If Trim(Text15.Value) = "一般案名最多為15個字" Then
            sql &= "'' , "
        Else
            sql &= "N'" & Text15.Value & "' , "
        End If
        '地上層數, 
        sql &= "'" & TextBox88.Text & "' , "
        '地下層數,
        sql &= "'" & TextBox89.Text & "' , "
        '每層戶數, 
        sql &= "'" & TextBox91.Text & "' , "
        '每層電梯數, 
        sql &= "'" & TextBox92.Text & "' , "
        '竣工日期,
        sql &= "'" & Text2.Value & "' , "
        '座向, 
        sql &= "'" & DropDownList22.SelectedValue & "' , "

        '管理費, num 
        If DropDownList5.SelectedValue <> "無" And DropDownList5.SelectedValue <> "未知" And DropDownList5.SelectedValue <> "" Then
            If TextBox36.Text = "" Then
                sql &= "null ,"
            Else
                sql &= TextBox36.Text & " , "
            End If
        Else
            sql &= "null ,"
        End If

        '含公設車位坪數
        If Checkbox27.Checked = True Then '含公設車位坪數有勾
            sql &= "'Y' , "
        Else
            sql &= "'N' , "
        End If

        '車位類別, 
        'Dim carstr As String = ""
        'If input55.Value.Length > 0 And input55.Value <> "0" Then '車位售價有填
        '    carstr = "有"

        'End If
        'If Checkbox27.Checked = True Then '含公設車位坪數有勾
        '    carstr = "有"

        'End If
        'If TextBox21.Text.Length > 0 And TextBox21.Text <> "0.000000" And TextBox21.Text <> "0" Then '公設車位平方公尺有填
        '    carstr = "有"

        'End If
        'sql &= "'" & carstr & "' , "

        '車位說明, 
        'sql &= "'" & TextBox93.Text & "' , "
        sql &= "'' , "

        '車位坪數,num
        If TextBox27.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox27.Text & " , "
        End If

        '車位平方公尺,num
        If TextBox26.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox26.Text & " , "
        End If

        '訴求重點,
        If TextBox102.Text = "特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。" Then
            sql &= "'' , "
        Else
            sql &= "N'" & TextBox102.Text & "' , "
        End If

        '店代號, 
        If store.SelectedValue <> "請選擇" Then
            sql &= "'" & store.SelectedValue & "' , "
        Else
            message1 = "請選擇店代號"
        End If

        '經紀人代號, 
        If sale1.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale1.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '營業員代號1, 
        If sale2.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale2.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '營業員代號2, 
        If sale3.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale3.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '委託起始日, 
        sql &= "'" & Left(Trim(Date2.Text), 7) & "' , "
        '委託截止日, 
        sql &= "'" & Left(Trim(Date3.Text), 7) & "' , "

        '公設比
        If TextBox41.Text <> "" Then
            sql &= TextBox41.Text & ", "
        Else
            sql &= "null , "
        End If

        '銷售狀態,
        sql &= "'" & DropDownList21.SelectedValue & "' , "

        '輸入者,
        sql &= "'" & Label33.Text & "' , "
        '新增日期, 
        sql &= "'" & sysdate & "' , "
        '修改日期, 
        sql &= "'" & sysdate & "' , "
        '修改時間,
        sql &= "GETDATE(), "


        If CheckBox1.Checked And CheckBox1.Visible Then
            sql &= "'是' , "
        Else
            Select Case ddl契約類別.SelectedValue
                Case "專任"
                    ' 專任約：除了直營+非管理，其他情況都為 "是"
                    If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                        sql &= If(CheckBox1.Checked, "'是' , ", "'否' , ")
                    Else
                        sql &= "'是' , "
                    End If
                Case "庫存"
                    sql &= "'否' , "
                Case Else
                    sql &= "'否' , "
            End Select
        End If


        '強銷物件 20151228 by nick
        If CheckBox5.Checked = True Then
            sql &= "'是' , "
        Else
            sql &= "'否' , "
        End If

        '鑰匙編號, 
        sql &= "'" & input66.Value & "' , "

        '帶看方式, 
        sql &= "'" & DropDownList20.SelectedValue & "' , "

        '土地標示, 
        If TextBox17.Text = "請填寫地號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox17.Text & "' , "
        End If
        '建號, 
        If TextBox18.Text = "請填寫建號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox18.Text & "' , "
        End If



        '上傳註記, 
        If CheckBox2.Checked = True Then
            sql &= "'D' , "
        Else
            sql &= "'A' , "
        End If

        '委託售價num, 
        If TextBox12.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox12.Text & " , "
        End If

        '車位價格,
        sql &= "'" & input55.Value & "' , "

        '車位租售,
        'sql &= "'" & DropDownList6.SelectedValue & "' , "
        sql &= "'' , "

        '國小, 
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox98.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX246.TEXT
            sql &= "'" & TextBox246.Text & "' , "
        Else
            sql &= "'',"
        End If

        '國中, 
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox99.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If


        '車位數量, 
        'sql &= "'" & input53.Value & "' , "
        sql &= "'' , "

        '國小1, 
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox98.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX98.TEXT
            sql &= "'" & TextBox98.Text & "' , "
        Else
            sql &= "'',"
        End If

        '國中1, 
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox99.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX99.TEXT
            sql &= "'" & TextBox99.Text & "' , "
        Else
            sql &= "'',"
        End If

        '部份地址, 
        sql &= "N'" & address2 & "' , "

        '捷運, 
        If DropDownList9.SelectedIndex > 0 Then
            sql &= "'" & DropDownList9.SelectedValue & "' , "
        Else
            sql &= "'',"
        End If

        '大專院校, 
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox101.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX249.TEXT
            sql &= "'" & TextBox249.Text & "' , "
        Else
            sql &= "'',"
        End If

        '其他學校-大專院校,
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox101.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX101.TEXT
            sql &= "'" & TextBox101.Text & "' , "
        Else
            sql &= "'',"
        End If

        '公車站名, 
        sql &= "'" & input60.Value & "' , "

        '銷售樓層-所在樓層,
        sql &= "'" & TextBox90.Text & "' , "

        '室,
        sql &= "'" & TextBox16.Text & "' , "

        '備註, 
        sql &= "'" & input79.Value & "' , "

        '公園代號,        
        If Trim(TextBox97.Text) <> "" Then
            'Dim 公園代號 As String = ""
            'Dim sql1 As String = ""
            'Dim n As Integer
            'Dim item1 As Array = Split(Page.Request.Form(TextBox97.UniqueID), ".")
            'Dim temp As String = ""
            'For i = 0 To item1.Length - 1
            '    If item1(i) <> "" Then
            '        temp &= "'" & item1(i) & "',"
            '        n += 1
            '    End If
            'Next i
            'Dim cc As String = ""
            'cc = Mid(temp, 1, Len(temp) - 1)

            'For i = 0 To n - 1
            '    sql1 = "Select 公園代號,公園名稱 FROM 公園資料表 where 公園名稱 in (" & cc & ") and 核准否='1'"
            '    adpt = New SqlDataAdapter(sql1, conn)
            '    ds = New DataSet()
            '    adpt.Fill(ds, "table4")
            '    table4 = ds.Tables("table4")
            '    If table4.Rows.Count > 0 Then
            '        If table4.Rows(i)("公園代號") <> "" Or IsDBNull(table4.Rows(i)("公園代號")) = False Then
            '            公園代號 &= table4.Rows(i)("公園代號") & "."
            '        End If
            '    End If

            'Next i

            '修正代號改以隱藏的TEXTBOX251.TEXT
            Dim parkCode As String = String.Join(".", TextBox251.Text.Split(".").Distinct().ToArray())      '
            sql &= "'" & parkCode & "' , "
        Else
            sql &= "'',"
        End If



        '高中,
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox100.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If

        '高中1, 
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox100.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取的TEXTBOX100.TEXT
            sql &= "'" & TextBox100.Text & "' , "
        Else
            sql &= "'' , "
        End If

        '管理費單位,
        sql &= "'" & DropDownList5.SelectedValue & "' , "

        '管理費繳交方式, 
        sql &= "'" & TextBox266.Text & "' , "

        '車位管理費類別,
        'sql &= "'" & DropDownList25.SelectedValue & "' , "
        sql &= "'' , "

        '車位管理費, num
        'If DropDownList24.SelectedValue <> "無" And DropDownList24.SelectedValue <> "未知" And DropDownList24.SelectedValue <> "" Then
        '    If TextBox94.Text = "" Then
        '        sql &= "null ,"
        '    Else
        '        sql &= TextBox94.Text & " , "
        '    End If
        'Else
        '    sql &= "null ,"
        'End If
        sql &= "null ,"

        '車位管理費單位
        'sql &= "'" & DropDownList24.SelectedValue & "', "
        sql &= "'', "

        '大樓代號
        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "其他" Then
            sql &= "'',"
        Else
            If ddl社區大樓.SelectedIndex > 0 Then
                sql &= "'" & ddl社區大樓.SelectedValue & "',"
            Else
                sql &= "'',"
            End If
        End If


        '登記日期20110620
        If Trim(Text11.Value) <> "" Then
            sql &= "'" & Left(Trim(Text11.Value), 7) & "', "
        Else
            sql &= "'', "
        End If


        '網站總坪數20110803
        '網站總坪數 20151231 修正不管是否含於開價中 總坪數皆需加入
        If IsNumeric(TextBox26.Text) Then
            Dim sum1 As Double = 0

            If IsNumeric(TextBox28.Text) Then sum1 += TextBox28.Text '總坪數
            If IsNumeric(TextBox26.Text) Then sum1 += TextBox26.Text '產權獨立車位面積

            sql &= Round(sum1 * 0.3025, 4) & ","
        Else
            '總坪數,num
            If TextBox29.Text = "" Then
                sql &= "0,"
            Else
                sql &= TextBox29.Text & ","
            End If

        End If

        '棟別
        If Trim(add11.Text) <> "" Then
            sql &= "'" & Trim(add11.Text) & "', "
        Else
            sql &= "'', "
        End If

        '臨路寬
        If Trim(TextBox245.Text) <> "" Then
            sql &= "'" & Trim(TextBox245.Text) & "', "
        Else
            sql &= "'', "
        End If

        '面寬
        If Trim(TextBox39.Text) <> "" Then
            sql &= "'" & Trim(TextBox39.Text) & "', "
        Else
            sql &= "'', "
        End If

        '縱深
        If Trim(TextBox40.Text) <> "" Then
            sql &= "'" & Trim(TextBox40.Text) & "', "
        Else
            sql &= "'', "
        End If
        '磁扣配對
        If Trim(TextBox267.Text) <> "" Then
            sql &= "'" & Trim(TextBox267.Text) & "', "
        Else
            sql &= "'', "
        End If
        '其他使用分區
        'If Trim(TextBox253.Text) <> "" Then
        '    sql &= "'" & Trim(TextBox253.Text) & "', "
        'Else
        '    sql &= "'', "
        'End If
        sql &= "'', "
        '簡碼
        If Left(物件編號, 1) <> "9" Then
            Dim ezcode As String = Trim(get_num(store.SelectedValue))
            ez_code.Text = Trim(ezcode)
            sql &= "'" & ezcode & "', "
        Else
            sql &= "'', "
        End If

        '增建
        sql &= "'" & RadioButtonList3.SelectedValue & "', "

        '謄本編號
        sql &= "'" & Land_FileNo.Text & "' "

        '資料來源
        If Request("source") = "land" Or Request("source") = "sland" Or Request("source") = "eland" Then
            sql &= ",'" & Data_Source.Text & "' "
        End If
        '新增時合約終止強制為空
        sql &= ",'' "
        '提供個資
        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Then
                sql &= ",'Y' "
            Else
                sql &= ",'N' "
            End If
        Else
            '當不屬於一般約及委託編號<AAD80001時，直接帶入空值
            sql &= ",'' "
        End If

        '聯賣
        If CheckBox101.Visible = False Then
            sql &= " ,'N' "
        Else
            If CheckBox101.Checked = True Then
                sql &= " ,'Y' "
            Else
                sql &= " ,'N' "
            End If
        End If

        '社區養寵
        If Trim(DropDownList3.SelectedValue) = "預售屋" Or Trim(DropDownList3.SelectedValue) = "土地" Then
            sql &= " ,'0','0','0' "
        Else
            If RadioButton4.Checked = True Then
                sql &= " ,'1' "
                If CheckBox102.Checked = True Then
                    sql &= " ,'1' "
                Else
                    sql &= " ,'0' "
                End If
                If CheckBox103.Checked = True Then
                    sql &= " ,'1' "
                Else
                    sql &= " ,'0' "
                End If
            Else
                sql &= " ,'0','0','0' "
            End If
        End If

        '聯賣日期
        If CheckBox101.Visible = False Then
            sql &= " ,NULL "
            Label465.Text = "N"
        Else
            If CheckBox101.Checked = True Then
                sql &= " ,GETDATE() "
                Label465.Text = "Y"
            Else
                sql &= " ,NULL "
                Label465.Text = "N"
            End If
        End If

        sql &= ") "



        If message1 <> "" Then
            script += "alert('" & message1 & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            Exit Sub
        End If



        cmd = New SqlCommand(sql, conn)

        Try
            'If Request.Cookies("store_id").Value = "A0279" Then
            '    checkleave.錯誤訊息(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "物件新增", "", sql, "")
            '    trans = "False"
            '    'Response.Write(sql)
            '    Exit Sub
            'End If

            Dim message2 As String
            cmd.ExecuteNonQuery()

            'If Request.Cookies("store_id").Value = "A0001" Then
            If CheckBox101.Checked = True Then
                sql = " insert into 聯賣資料 "
                sql += " (聯賣區域,聯賣編號,店代號,物件編號,異動日期) "
                sql += " select '" & 聯賣組別 & "' "
                sql += " ,isnull((select max(聯賣編號) from 聯賣資料 where 聯賣區域='" & 聯賣組別 & "'),0)+1 "
                sql += " ,'" & store.SelectedValue & "','" & UCase(Microsoft.VisualBasic.Strings.StrConv(物件編號, Microsoft.VisualBasic.VbStrConv.Narrow)) & "',GETDATE() "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If
            'End If


            '1050506表單控管
            If ddl契約類別.SelectedValue <> "庫存" Then
                'If Request.Cookies("webfly_empno").Value = "92H" Then
                If check_formno(UCase(Mid(物件編號, 6)), store.SelectedValue, 物件編號, input4.Value, "新增") = "use" Then
                    eip_usual.Show("一併寫入表單管理")
                Else
                    eip_usual.Show("輸入編號不為所購買表單，無法寫入表單管理")
                End If
                'End If
            End If

            If ddl契約類別.SelectedValue = "一般" Or ddl契約類別.SelectedValue = "專任" Or ddl契約類別.SelectedValue = "同意書" Then
                message2 = formnocheck.form_usestate(UCase(物件編號), store.SelectedValue)

                '1040114修正謄本帶入新增細項面積
                If message2 = "新增成功" And Request("FILENO") <> "" Then
                    Call landcase_detail(UCase(物件編號))
                    Call landcase_他項權利(UCase(物件編號))
                End If
            Else
                Call landcase_detail(UCase(物件編號))
                Call landcase_他項權利(UCase(物件編號))
                message2 = "新增成功"
            End If



            ''重整可使用表單內容
            'ddl契約類別_SelectedIndexChanged(Nothing, Nothing)

            'With Image1
            '    .ImageUrl = "https://el.etwarm.com.tw/new_eip/tool/code41.ashx?id=" & 物件編號
            '    .Visible = True
            'End With

            Label57.Text = 物件編號
            Page.SetFocus(ddl契約類別)

            If message2 = "新增成功" Then
                '判斷是否轉頁的參數
                trans = "True"
            Else
                '判斷是否轉頁的參數
                trans = "False"
            End If

            script += "alert('" & message2 & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            'eip_usual.Show(message2)

            'Dim href As String = "https://el.etwarm.com.tw/new_eip/tool/t_條碼分類.aspx?state=update&oid=" & 物件編號 & "&sid=" & store.SelectedValue & "&folder=available&rsid=" & Request("sid") & "&src=" & Request("src") & "&check=description" 'check參數用意在不顯示"下一步按鈕"
            'Dim NSCRIPT As String = "GB_showCenter('產生條碼', '" & href & "',320,580);"
            'Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", NSCRIPT, True)


        Catch ex As Exception
            '判斷是否轉頁的參數
            trans = "False"
            checkleave.錯誤訊息(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "物件新增", Request.Url.ToString, sql, "")
            script += "alert('新增失敗-新增" & Replace(ex.ToString, "'", "") & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        End Try
    End Sub

    Public Sub 更新記錄()
        Dim j = 驗證數字是否阿拉伯數字()
        If j > 0 Then
            Exit Sub
        End If

        Dim 物件編號 As String = ""
        Dim 使用分區 As String = ""
        Dim sid As String = Request("sid")
        Dim oid As String = Request("oid")

        物件編號 = oid.Trim
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        Dim count As Integer = 0
        Dim script, message As String
        script = ""
        message = ""

        '物件編號限制11碼到13碼
        If 物件編號 = "" Then
            message &= "物件編號為空值，請查明之\n"
        End If

        'If Len(Date2.Text) <> 7 Or Len(Date3.Text) <> 7 Then
        '    message &= "委託期間格式為0960108\n"
        'End If

        If TextBox102.Text <> "" And CheckBox2.Checked = False Then
            If Replace(TextBox102.Text.ToString, vbNewLine, "").Length > 1000 Then
                message &= "訴求重點字數上限為1000字\n"
            ElseIf Replace(TextBox102.Text.ToString, vbNewLine, "").Length < 20 Then
                message &= "訴求重點字數至少輸入20字\n"
            End If
        End If

        'address1完整地址,address2部份地址-到"弄"
        Dim address1 As String = "", address2 As String = ""

        '部分地址
        If add1.Text <> "" Then address2 &= add1.Text & zone3.SelectedValue
        If add2.Text <> "" Then address2 &= add2.Text & "鄰"
        If add3.Text <> "" Then address2 &= add3.Text & address20.SelectedValue

        '1040519修正
        'If add4.Text <> "" Then address2 &= add4.Text & "段"
        If add4.Text <> "" Then address2 &= add4.Text & Label64.Text

        If add5.Text <> "" Then address2 &= add5.Text & "巷"
        If add6.Text <> "" Then address2 &= add6.Text & "弄"
        address1 &= address2

        '1040519修正
        'If add7.Text <> "" Then address1 &= add7.Text & "號"
        If Label64.Text = "小段" Then
            If add7.Text <> "" Then address1 &= add7.Text & "地號"
        Else
            If add7.Text <> "" Then address1 &= add7.Text & "號"
        End If

        If add8.Text <> "" Then address1 &= "之" & add8.Text

        '20100607小豪修正("之"後方若有值,則"樓"之前加入空白,避免距離過近造成誤會,ex:101號之1   3樓)----
        If add9.Text <> "" Then
            If add8.Text <> "" Then
                address1 &= "   " & add9.Text & "樓"
            Else
                address1 &= add9.Text & "樓"
            End If
        End If
        '--------------------------------------------------------------------------------------------------
        If add10.Text <> "" Then address1 &= "之" & add10.Text


        '2017.01.05 by Finch 聯賣規則
        If 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    'If Request.Cookies("webfly_empno").Value = "05W" Then
                    '    Response.Write(sql)
                    '    Exit Sub
                    'End If
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")
                    If table1.Rows.Count > 0 Then
                        'If Request.Cookies("webfly_empno").Value = "05W" Then
                        For j = 0 To table1.Rows.Count - 1
                            sql = " update 委賣物件資料表 "
                            sql += " set 聯賣='N' "
                            sql += " where 物件編號 = '" & table1.Rows(j)("物件編號").ToString.Trim & "' and 店代號='" & table1.Rows(j)("店代號").ToString.Trim & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        Next
                        'End If
                        'message &= "此區域已存在相同物件\n"
                    End If
                End If
            End If
        End If

        If Not 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    'If Request.Cookies("webfly_empno").Value = "08AZ" Then
                    '    Response.Write(sql)
                    '    'Exit Sub
                    'End If
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")
                    If table1.Rows.Count > 0 Then
                        message &= "此區域已存在相同物件\n"
                    End If
                Else
                    If table1.Rows(0)("區域名稱").ToString.Trim = "宜蘭區聯賣" And CType(Trim(Me.Date2.Text), Integer) >= "1051101" Then
                        Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")
                        If table1.Rows.Count > 0 Then
                            message &= "此區域已存在相同物件\n"
                        End If
                    ElseIf table1.Rows(0)("區域名稱").ToString.Trim = "高屏區聯賣" Then
                        Dim kanPingSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & kanPingSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")

                        'Response.Write(sql)

                        If table1.Rows.Count > 0 Then
                            Dim queryOid As String = ""
                            For i = 0 To table1.Rows.Count - 1
                                queryOid &= table1.Rows(i)("物件編號") & "', '"
                            Next

                            sql2 = " Select TOP 1 *,convert(char, [借用時間], 111) AS 借用時間西元 From 物件鑰匙管理 With(NoLock) "
                            sql2 &= " Where 物件編號 in ('" & queryOid & "') AND 借用店 = '" & store.SelectedValue & "' "
                            sql2 &= " ORDER BY num DESC "
                            adpt = New SqlDataAdapter(sql2, conn)
                            ds = New DataSet()
                            adpt.Fill(ds, "table2")
                            table2 = ds.Tables("table2")

                            'Response.Write(sql2)

                            If table2.Rows.Count > 0 Then
                                message &= "此區域已存在相同物件，物件編號：" & table2.Rows(0)("物件編號") & "\n"
                                If table2.Rows(0)("借用項目") = "O" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾點閱此物件物調表。"
                                ElseIf table2.Rows(0)("借用項目") = "M" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件不動產說明書。"
                                ElseIf table2.Rows(0)("借用項目") = "K" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件鑰匙。"
                                ElseIf table2.Rows(0)("借用項目") = "C" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾電話詢問此物件。"
                                End If
                            End If
                        End If


                        For i = 0 To table1.Rows.Count - 1

                        Next

                    End If
                End If
            End If

        End If

        '2017.01.19 by Finch 同群組物件不重複        
        sql = "Select * From HSSTRUCTURE With(NoLock) "
        sql &= "Where 店代號 = '" & store.SelectedValue & "' AND 同群物件不重覆 = 'Y'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        Dim groupSid As String = ""

        If table1.Rows.Count > 0 Then
            Dim groupId As String = table1.Rows(0)("組別").ToString.Trim
            sql = "Select * From HSSTRUCTURE With(NoLock) "
            sql &= "Where 組別 = '" & groupId & "'"

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                For i = 0 To table1.Rows.Count - 1
                    groupSid &= table1.Rows(i)("店代號").ToString.Trim & "','"
                Next

                sql = "Select * From 委賣物件資料表 With(NoLock) "
                sql &= "Where 店代號 in ('" & groupSid & "') AND 店代號 <> '" & store.SelectedValue & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' AND 註記 = 'Y' "
                sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                If table1.Rows.Count > 0 Then
                    message &= "同群組已存在相同物件，物件編號：" & table1.Rows(0)("物件編號") & "\n"
                End If

            End If
        End If

        'If TextBox37.Text.Trim <> "" Then
        '    sql = "Select * From 資料_車位 With(NoLock) "
        '    sql &= "Where 車位 = '" & TextBox37.Text.Trim & "' and 店代號 in ('A0001','" & store.SelectedValue & "') "
        '    adpt = New SqlDataAdapter(sql, conn)
        '    ds = New DataSet()
        '    adpt.Fill(ds, "table1")
        '    table1 = ds.Tables("table1")
        '    Dim flag1 As String = ""
        '    If table1.Rows.Count > 0 Then
        '        message &= "此車位類別已存在,請下拉選擇\n"
        '        TextBox37.Text = ""
        '    Else
        '        If Len(TextBox37.Text.Trim) < 5 Then
        '            Dim sql2 = "insert into 資料_車位 (車位,店代號) values ("
        '            sql2 &= "'" & TextBox37.Text.Trim & "' , '" & Request.Cookies("store_id").Value & "')"
        '            cmd = New SqlCommand(sql2, conn)
        '            cmd.ExecuteNonQuery()
        '            cmd.Dispose()
        '        Else
        '            message &= "此新增車位類別超過長度,請修正\n"
        '        End If

        '    End If
        'End If

        If DDL_County.SelectedValue = "選擇縣市" Then
            message &= "尚未選擇縣市,請選擇\n"
        End If

        If DDL_Area.SelectedValue = "選擇鄉鎮市區" Then
            message &= "尚未選擇鄉鎮市區,請選擇\n"
        End If


        If message <> "" Then
            script += "alert('" & message & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            Exit Sub
        End If

        '2016.06.21 by Finch 從"委賣物件資料表_面積細項"抓第一筆土地面積的使用分區，
        '寫回"委賣物件資料表"的"物件用途"欄位(前台抓使用分區的欄位)
        'If Request.Cookies("webfly_empno").Value = "92H" Then
        sql = "Select TOP 1 * From 委賣物件資料表_面積細項 With(NoLock) "
        sql &= "Where 物件編號 = '" & 物件編號 & "' and 店代號 = '" & store.SelectedValue & "' and 項目名稱 = '土地面積' and 使用分區 IS NOT NULL and 使用分區 <> '' ORDER BY 流水號 "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count > 0 Then
            使用分區 = table1.Rows(0)("使用分區")
        Else
            使用分區 = ""
        End If
        'End If        
        '============================================================================
        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Response.Write("0001" & TextBox28.TEXT & src.Text)
        '    Exit Sub
        'End If

        sql = "Update " & src.Text '--20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
        sql &= " Set 物件編號 = "
        sql &= "'" & 物件編號 & "' , "

        '物件主要用途
        sql &= "建物型態 ="
        If TextBox4.Text <> "" Then
            sql &= "'" & TextBox4.Text & "' , "
        ElseIf DropDownList19.SelectedValue <> "" Then
            sql &= "'" & DropDownList19.SelectedValue & "' , "
        Else
            sql &= "'' , "
        End If

        '使用分區-主
        '2016.06.21 by Finch
        If 使用分區 <> "請選擇" Then
            sql &= "物件用途 ="
            sql &= "'" & 使用分區 & "' , "
        Else
            sql &= "物件用途 ="
            sql &= "'' , "
        End If
        'sql &= "物件用途 ="
        'If DropDownList17.Visible = True Then
        '    If DropDownList17.SelectedValue = "請選擇" Then
        '        sql &= "'" & DropDownList16.SelectedValue & "' , "
        '    Else
        '        sql &= "'" & DropDownList17.SelectedValue & "' , "
        '    End If
        'Else
        '    sql &= "'" & DropDownList16.SelectedValue & "' , "
        'End If

        ''其他使用分區
        'sql &= "其他使用分區 ="
        'If Trim(TextBox253.Text) <> "" Then
        '    sql &= "'" & Trim(TextBox253.Text) & "', "
        'Else
        '    sql &= "'', "
        'End If


        '物件型態
        sql &= "物件類別 ="
        sql &= "'" & DropDownList3.SelectedValue & "' , "


        sql &= "郵遞區號 ="
        sql &= "'" & TB_AreaCode.Text & "' , "
        sql &= "商圈代號 ="

        '商圈
        If Trim(TextBox96.Text) <> "" Then
            'Dim j, m As Integer
            'Dim store As String = ""
            'Dim sql1 As String = ""


            ''為了舊資料相容,要再加上.
            'j = InStr(Page.Request.Form(TextBox96.UniqueID), ".")
            'If j = 0 Then
            '    Page.Request.Form(TextBox96.UniqueID) = Page.Request.Form(TextBox96.UniqueID) + "."
            'End If
            'Dim item1 As Array = Split(Page.Request.Form(TextBox96.UniqueID), ".")
            'Dim temp As String = ""
            'For i = 0 To item1.Length - 1
            '    If item1(i) <> "" Then
            '        temp &= "'" & item1(i) & "',"
            '        m += 1
            '    End If
            'Next i
            'Dim cc As String = ""
            'cc = Mid(temp, 1, Len(temp) - 1)

            'For i = 0 To m - 1
            '    sql1 = "Select 商圈代號,商圈名稱 FROM 精華生活圈資料表 where 商圈名稱 in (" & cc & ") and 核准否='1'"
            '    adpt = New SqlDataAdapter(sql1, conn)
            '    ds = New DataSet()
            '    adpt.Fill(ds, "table4")
            '    table4 = ds.Tables("table4")
            '    If IsDBNull(table4.Rows(i)("商圈代號")) = False Or table4.Rows(i)("商圈代號") <> "" Then
            '        store &= table4.Rows(i)("商圈代號") & "."
            '    End If
            'Next i

            'sql &= "'" & store & "' , "

            '修正代號改以隱藏的TEXTBOX250.TEXT
            sql &= "'" & TextBox250.Text & "' , "
        Else
            sql &= "'',"
        End If



        sql &= "完整地址 ="
        sql &= "N'" & address1 & "' , "
        sql &= "縣市 ="
        sql &= "'" & DDL_County.SelectedValue & "' , "
        sql &= "鄉鎮市區 ="
        sql &= "'" & DDL_Area.SelectedItem.Text & "' , "
        sql &= "村里 ="
        sql &= "'" & add1.Text & "' , "
        sql &= "村里別 ="
        sql &= "'" & zone3.SelectedValue & "' , "
        sql &= "鄰 ="
        sql &= "'" & add2.Text & "' , "
        sql &= "路名 ="
        sql &= "N'" & add3.Text & "' , "
        sql &= "路別 ="
        sql &= "N'" & address20.SelectedValue & "' , "
        sql &= "段 ="
        sql &= "'" & add4.Text & "' , "
        sql &= "巷 ="
        sql &= "'" & add5.Text & "' , "
        sql &= "弄 ="
        sql &= "'" & add6.Text & "' , "
        sql &= "號 ="
        sql &= "'" & add7.Text & "' , "
        sql &= "之 ="
        sql &= "'" & add8.Text & "' , "
        sql &= "所在樓層 ="
        sql &= "'" & add9.Text & "' , "
        sql &= "樓之 ="
        sql &= "'" & add10.Text & "' , "

        sql &= "刊登售價 ="
        If TextBox12.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox12.Text & " , "
        End If
        If TextBox12.Text <> Label462.Text Then
            sql &= "原始售價 =null,"
            sql &= "降價幅度 =null,"
        End If

        sql &= "主建物 ="
        If TextBox6.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox6.Text & " , "
        End If
        sql &= "主建物平方公尺 ="
        If TextBox5.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox5.Text & " , "
        End If

        sql &= "附屬建物 ="
        If TextBox8.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox8.Text & " , "
        End If
        sql &= "附屬建物平方公尺 ="
        If TextBox7.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox7.Text & " , "
        End If

        sql &= "公共設施 ="
        If TextBox10.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox10.Text & " , "
        End If
        sql &= "公共設施平方公尺 ="
        If TextBox9.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox9.Text & " , "
        End If

        sql &= "公設內車位坪數 ="
        If TextBox23.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox23.Text & " , "
        End If
        sql &= "公設內車位平方公尺 ="
        If TextBox21.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox21.Text & " , "
        End If

        sql &= "地下室 ="
        If TextBox20.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox20.Text & " , "
        End If
        sql &= "地下室平方公尺 ="
        If TextBox19.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox19.Text & " , "
        End If

        sql &= "總坪數 ="
        If TextBox29.Text = "" Then
            sql &= "0,"
        Else
            sql &= TextBox29.Text & " , "
        End If
        sql &= "總平方公尺 ="
        If TextBox28.Text = "" Then
            sql &= "0,"
        Else
            sql &= TextBox28.Text & " , "
        End If

        sql &= "土地坪數 ="
        If TextBox31.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox31.Text & " , "
        End If
        sql &= "土地平方公尺 ="
        If TextBox30.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox30.Text & " , "
        End If

        sql &= "車位坪數 ="
        If TextBox27.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox27.Text & " , "
        End If
        sql &= "車位平方公尺 ="
        If TextBox26.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox26.Text & " , "
        End If

        sql &= "庭院坪數 ="
        If TextBox33.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox33.Text & " , "
        End If
        sql &= "庭院平方公尺 ="
        If TextBox32.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox32.Text & " , "
        End If
        sql &= "加蓋坪數 ="
        If TextBox35.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox35.Text & " , "
        End If
        sql &= "加蓋平方公尺 ="
        If TextBox34.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox34.Text & " , "
        End If

        sql &= "房 ="
        If C1.Checked = True Then
            sql &= "'-1' , "
        Else
            sql &= "'" & TextBox13.Text & "' , "
        End If
        sql &= "廳 ="
        sql &= "'" & TextBox14.Text & "' , "
        sql &= "衛 ="
        If TextBox15.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox15.Text & " , "
        End If
        sql &= "室 ="
        sql &= "'" & TextBox16.Text & "' , "

        sql &= "建築名稱 ="
        sql &= "N'" & input4.Value & "' , "
        sql &= "張貼卡案名 ="
        sql &= "N'" & Text15.Value & "' , "
        sql &= "地上層數 ="
        sql &= "'" & TextBox88.Text & "' , "
        sql &= "地下層數 ="
        sql &= "'" & TextBox89.Text & "' , "
        sql &= "銷售樓層 ="
        sql &= "'" & TextBox90.Text & "' , " '所在樓層

        sql &= "每層戶數 ="
        sql &= "'" & TextBox91.Text & "' , "
        sql &= "每層電梯數 ="
        sql &= "'" & TextBox92.Text & "' , "

        sql &= "竣工日期 ="
        sql &= "'" & Text2.Value & "' , "
        sql &= "座向 ="
        sql &= "'" & DropDownList22.SelectedValue & "' , "
        sql &= "管理費 ="
        If TextBox36.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox36.Text & " , "
        End If

        '含公設車位坪數
        sql &= "含公設車位坪數 ="
        If Checkbox27.Checked = True Then '含公設車位坪數有勾
            sql &= "'Y' , "
        Else
            sql &= "'N' , "
        End If

        '車位類別
        sql &= "車位 ="
        Dim carstr As String = ""
        If input55.Value.Length > 0 And input55.Value <> "0" Then '車位售價有填
            carstr = "有"
        End If
        If Checkbox27.Checked = True Then '含公設車位坪數有勾
            carstr = "有"
        End If
        If TextBox21.Text.Length > 0 And TextBox21.Text <> "0.000000" And TextBox21.Text <> "0" Then '公設車位平方公尺有填
            carstr = "有"
        End If
        sql &= "'" & carstr & "' , "

        sql &= "車位說明 ="
        'sql &= "'" & TextBox93.Text & "' , "
        sql &= "'' , "
        sql &= "訴求重點 ="
        If TextBox102.Text = "特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。" Then
            sql &= "'' , "
        Else
            sql &= "N'" & TextBox102.Text & "' , "
        End If

        sql &= "店代號 ="
        sql &= "'" & store.SelectedValue & "' , "
        sql &= "經紀人代號 ="
        If sale1.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale1.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If
        sql &= "營業員代號1 ="
        If sale2.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale2.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If
        sql &= "營業員代號2 ="
        If sale3.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale3.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If
        sql &= "委託起始日 ="
        sql &= "'" & Left(Trim(Date2.Text), 7) & "' , "
        sql &= "委託截止日 ="
        sql &= "'" & Left(Trim(Date3.Text), 7) & "' , "
        sql &= "公設比 ="
        If TextBox41.Text <> "" Then
            sql &= TextBox41.Text & ", "
        Else
            sql &= "null , "
        End If

        If CheckBox100.Checked = True Then
            sql &= "銷售狀態 = '已終止' , "
        Else
            sql &= "銷售狀態 = '" & DropDownList21.SelectedValue & "' , "
        End If

        sql &= "輸入者 ="
        sql &= "'" & Label33.Text & "' , "
        sql &= "修改日期 ="
        sql &= "'" & sysdate & "' , "
        sql &= "修改時間 ="
        sql &= "GETDATE(), "
        sql &= "網頁刊登 ="
        If CheckBox1.Checked = True And ddl契約類別.SelectedValue <> "庫存" Then
            sql &= "'是' , "
        Else
            sql &= "'否' , "
        End If
        '強銷物件 20151228 BY NICK
        sql &= "強銷物件 ="
        If CheckBox5.Checked = True Then
            sql &= "'是' , "
        Else
            sql &= "'否' , "
        End If
        sql &= "鑰匙編號 ="
        sql &= "'" & input66.Value & "' , "
        sql &= "帶看方式 ="
        sql &= "'" & DropDownList20.SelectedValue & "' , "
        sql &= "土地標示 ="
        If TextBox17.Text = "請填寫地號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox17.Text & "' , "
        End If
        sql &= "建號 ="
        If TextBox18.Text = "請填寫建號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox18.Text & "' , "
        End If
        sql &= "上傳註記 ="
        If CheckBox2.Checked = True Then
            sql &= "'D' , "
        Else
            sql &= "'U' , "
        End If
        sql &= "委託售價 ="
        sql &= TextBox12.Text & " , "
        sql &= "車位價格 ="
        sql &= "'" & input55.Value & "' , "
        sql &= "車位租售 ="
        sql &= "'',"
        'If DropDownList6.Items(1).Selected = True Then
        '    sql &= "'租' , "
        'ElseIf DropDownList6.Items(2).Selected = True Then
        '    sql &= "'售' , "
        'Else
        '    sql &= "'',"
        'End If

        sql &= "國小 ="
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox26.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX246.TEXT
            sql &= "'" & TextBox246.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "國中 ="
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox24.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "國小1 ="
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox26.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX98.TEXT
            sql &= "'" & TextBox98.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "國中1 ="
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox24.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "
            '修正校名直接讀取TEXTBOX99.TEXT
            sql &= "'" & TextBox99.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "高中 ="
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox27.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If
        sql &= "高中1 ="
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox27.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "'"

            '修正校名直接讀取的TEXTBOX100.TEXT
            sql &= "'" & TextBox100.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "大專院校 ="
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox28.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX249.TEXT
            sql &= "'" & TextBox249.Text & "' , "
        Else
            sql &= "'',"
        End If

        sql &= "其他學校 ="
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox28.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX101.TEXT
            sql &= "'" & TextBox101.Text & "' , "
        Else
            sql &= "'',"
        End If


        ''貸款銀行
        'sql &= "貸款銀行 ="
        'sql &= "'" & input50.Value & "' , "
        'sql &= "貸款金額 ="
        'If Trim(input80.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= input80.Value & " , "
        'End If
        ''貸款銀行2
        'sql &= "貸款銀行2 ="
        'sql &= "'" & Text5.Value & "' , "
        'sql &= "貸款金額2 ="
        'If Trim(Text6.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text6.Value & " , "
        'End If
        ''貸款銀行3
        'sql &= "貸款銀行3 ="
        'sql &= "'" & Text7.Value & "' , "
        'sql &= "貸款金額3 ="
        'If Trim(Text8.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text8.Value & " , "
        'End If
        ''貸款銀行4
        'sql &= "貸款銀行4 ="
        'sql &= "'" & Text9.Value & "' , "
        'sql &= "貸款金額4 ="
        'If Trim(Text10.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text10.Value & " , "
        'End If

        sql &= "車位數量 ="
        'sql &= "'" & input53.Value & "' , "
        sql &= "'' , "

        sql &= "部份地址 ="
        sql &= "N'" & address2 & "' , "
        sql &= "捷運 ="
        If DropDownList9.SelectedIndex > 0 Then
            sql &= "'" & DropDownList9.SelectedValue & "' , "
        Else
            sql &= "'',"
        End If


        sql &= "公車站名 ="
        sql &= "'" & input60.Value & "' , "


        sql &= "備註 ="
        sql &= "'" & input79.Value & "' , "

        '公園代號,  
        sql &= "公園代號 ="
        If Trim(TextBox97.Text) <> "" Then
            '    Dim s, n As Integer
            '    Dim park As String = ""
            '    Dim sql1 As String = ""
            '    '為了舊資料相容,要再加上.
            '    s = InStr(Page.Request.Form(TextBox34.UniqueID), ".")
            '    If s = 0 Then
            '        Page.Request.Form(TextBox34.UniqueID) = Page.Request.Form(TextBox34.UniqueID) + "."
            '    End If
            '    Dim item1 As Array = Split(Page.Request.Form(TextBox34.UniqueID), ".")
            '    Dim temp As String = ""
            '    For i = 0 To item1.Length - 1
            '        If item1(i) <> "" Then
            '            temp &= "'" & item1(i) & "',"
            '            n += 1
            '        End If
            '    Next i
            '    Dim cc As String = ""
            '    cc = Mid(temp, 1, Len(temp) - 1)

            '    For i = 0 To n - 1
            '        sql1 = "Select 公園代號,公園名稱 FROM 公園資料表 where 公園名稱 in (" & cc & ") and 核准否='1'"
            '        adpt = New SqlDataAdapter(sql1, conn)
            '        ds = New DataSet()
            '        adpt.Fill(ds, "table4")
            '        table4 = ds.Tables("table4")
            '        If table4.Rows(i)("公園代號") <> "" Or IsDBNull(table4.Rows(i)("公園代號")) = False Then
            '            park &= table4.Rows(i)("公園代號") & "."
            '        End If

            '    Next i

            '    sql &= "'" & park & "' , "

            '修正代號改以隱藏的TEXTBOX251.TEXT
            Dim parkCode As String = String.Join(".", TextBox251.Text.Split(".").Distinct().ToArray())      '
            sql &= "'" & parkCode & "' , "
        Else
            sql &= "'',"
        End If


        sql &= "管理費單位 ="
        sql &= "'" & DropDownList5.SelectedValue & "' , "
        sql &= "管理費繳交方式 ="
        sql &= "'" & TextBox266.Text & "' , "
        sql &= "車位管理費類別 ="
        'sql &= "'" & DropDownList25.SelectedValue & "' , "
        sql &= "'' , "
        sql &= "車位管理費 ="
        sql &= "null,"
        'If TextBox94.Text = "" Then
        '    sql &= "null,"
        'Else
        '    sql &= TextBox94.Text & " , "
        'End If

        sql &= "車位管理費單位 ="
        'sql &= "'" & DropDownList24.SelectedValue & "',  "
        sql &= "'',  "

        '大樓代號
        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "其他" Then
            sql &= "大樓代號='', "
        Else
            If ddl社區大樓.SelectedIndex > 0 Then
                sql &= "大樓代號='" & ddl社區大樓.SelectedValue & "', "
            Else
                sql &= "大樓代號='', "
            End If
        End If


        '登記日期20110620
        If Trim(Text11.Value) <> "" Then
            sql &= "register_date='" & Left(Trim(Text11.Value), 7) & "', "
        Else
            sql &= "register_date='', "
        End If

        '網站總坪數20110803
        '網站總坪數 20151231 修正不管是否含於開價中 總坪數皆需加入
        If IsNumeric(TextBox26.Text) Then
            Dim sum1 As Double = 0
            If IsNumeric(TextBox28.Text) Then sum1 += TextBox28.Text
            If IsNumeric(TextBox26.Text) Then sum1 += TextBox26.Text


            sql &= " 網站總坪數=" & Round(sum1 * 0.3025, 4) & ","

            'Response.Write(sum1 & "=" & TextBox28.Text & "+" & TextBox26.Text & "----" & Round(sum1 * 0.3025, 4) & "----" & sum1 * 0.3025)

        Else
            '總坪數,num
            If TextBox29.Text = "" Then
                sql &= " 網站總坪數=0, "
            Else
                sql &= " 網站總坪數=" & TextBox29.Text & " , "
            End If

        End If


        '登記日期20110620
        If Trim(add11.Text) <> "" Then
            sql &= "棟別='" & add11.Text & "', "
        Else
            sql &= "棟別='', "
        End If

        '臨路寬
        If Trim(TextBox245.Text) <> "" Then
            sql &= "臨路寬='" & TextBox245.Text & "', "
        Else
            sql &= "臨路寬='', "
        End If

        '面寬
        If Trim(TextBox39.Text) <> "" Then
            sql &= "面寬='" & TextBox39.Text & "', "
        Else
            sql &= "面寬='', "
        End If

        '縱深
        If Trim(TextBox40.Text) <> "" Then
            sql &= "縱深='" & TextBox40.Text & "', "
        Else
            sql &= "縱深='', "
        End If
        '磁扣配對
        If Trim(TextBox267.Text) <> "" Then
            sql &= "磁扣編號='" & TextBox267.Text & "', "
        Else
            sql &= "磁扣編號='', "
        End If
        '增建
        sql &= "增建='" & RadioButtonList3.SelectedValue & "' "

        'ezcode-補值
        If Trim(Request("src")) = "NOW" And ez_code.Text = "" And Left(物件編號, 1) <> "9" Then
            sql &= ",ezcode='" & get_num(Request("sid")) & "' "
        End If

        If CheckBox100.Checked = True Then
            sql &= ",合約終止='1' "
        Else
            sql &= ",合約終止='' "
        End If
        '提供個資
        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Then
                sql &= ",提供個資='Y' "
            Else
                sql &= ",提供個資='N' "
            End If
        Else
            '當不屬於一般約及委託編號<AAD80001時，直接帶入空值
            sql &= ",提供個資='' "
        End If

        ''網站下架日20111229--取消規定
        'If Trim(Me.Label24.Text) = "" Then
        '    If Trim(Me.TextBox22.Text) <> "" Then
        '        Dim netdate = DateSerial(Left(Trim(Me.TextBox22.Text), 3) + 1911, Mid(Trim(Me.TextBox22.Text), 4, 2) + 6, Right(Trim(Me.TextBox22.Text), 2))
        '        Label24.Text = Trim(Year(netdate) - 1911).PadLeft(3, "0") & Trim(Month(netdate)).PadLeft(2, "0") & Trim(Day(netdate)).PadLeft(2, "0")

        '        sql &= " 網站下架日='" & Label24.Text & "' "

        '    End If
        'Else
        '    sql &= " 網站下架日='" & Label24.Text & "' "
        'End If

        '聯賣
        If CheckBox101.Visible = False Then
            sql &= " ,聯賣='N' "
        Else
            If CheckBox101.Checked = True Then
                sql &= " ,聯賣='Y' "
            Else
                sql &= " ,聯賣='N' "
            End If
        End If

        '社區養寵
        If Trim(DropDownList3.SelectedValue) = "預售屋" Or Trim(DropDownList3.SelectedValue) = "土地" Then
            sql &= " ,社區養寵='0',養貓='0',養狗='0' "
        Else
            If RadioButton4.Checked = True Then
                sql &= " ,社區養寵='1' "
                If CheckBox102.Checked = True Then
                    sql &= " ,養貓='1' "
                Else
                    sql &= " ,養貓='0' "
                End If
                If CheckBox103.Checked = True Then
                    sql &= " ,養狗='1' "
                Else
                    sql &= " ,養狗='0' "
                End If
            Else
                sql &= " ,社區養寵='0',養貓='0',養狗='0' "
            End If
        End If

        '聯賣日期
        If CheckBox101.Visible = False Then
            sql &= " ,聯賣日期=NULL "
            Label465.Text = "N"
        Else
            If CheckBox101.Checked = True Then
                If Label465.Text <> "Y" Then
                    sql &= " ,聯賣日期=(case when 聯賣日期 is null then GETDATE() else 聯賣日期 end) "
                    Label465.Text = "Y"
                End If
            Else
                sql &= " ,聯賣日期=NULL "
                Label465.Text = "N"
            End If
        End If

        sql &= " Where 物件編號 = '" & 物件編號 & "' and 店代號 = '" & Request("sid") & "' "

        'If Request.Cookies("webfly_empno").Value = "1EI" Then
        '    Response.Write(sql)
        '    Response.End()
        'End If

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Response.Write(sql)
        '    Exit Sub
        'End If

        cmd = New SqlCommand(sql, conn)

        Try
            cmd.ExecuteNonQuery()
            'log 修改記錄
            LogTrack(sql)
            'If Request.Cookies("store_id").Value = "A0001" Then
            If CheckBox101.Checked = True Then
                sql = " select * "
                sql += " from 聯賣資料 "
                sql += " where 聯賣區域='" & 聯賣組別 & "' and 物件編號 = '" & 物件編號 & "' and 店代號='" & Request("sid") & "' "
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1_3")
                table1_3 = ds.Tables("table1_3")
                If table1_3.Rows.Count > 0 Then

                Else
                    sql = " insert into 聯賣資料 "
                    sql += " (聯賣區域,聯賣編號,店代號,物件編號,異動日期) "
                    sql += " select '" & 聯賣組別 & "' "
                    sql += " ,isnull((select max(聯賣編號) from 聯賣資料 where 聯賣區域='" & 聯賣組別 & "'),0)+1 "
                    sql += " ,'" & store.SelectedValue & "','" & UCase(Microsoft.VisualBasic.Strings.StrConv(物件編號, Microsoft.VisualBasic.VbStrConv.Narrow)) & "',GETDATE() "
                    cmd = New SqlCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                End If
            End If
            sql = " delete from 社區大樓歸戶 "
            sql += " where 店代號='" & store.SelectedValue & "' and 物件編號='" & UCase(Microsoft.VisualBasic.Strings.StrConv(物件編號, Microsoft.VisualBasic.VbStrConv.Narrow)) & "' "
            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

            If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "其他" Then

            Else
                If ddl社區大樓.SelectedIndex > 0 Then
                    sql = " insert into 社區大樓歸戶 "
                    sql += " (店代號, 物件編號,社區代號, "
                    sql += " 郵遞區號, 縣市, 鄉鎮市區, "
                    sql += " 村里, 村里別, 鄰, 路名, 路別, 段, 巷, 弄, 號, 之,所在樓層, 樓之, 完整地址, 更新日期) "
                    sql &= " select '" & store.SelectedValue & "','" & UCase(Microsoft.VisualBasic.Strings.StrConv(物件編號, Microsoft.VisualBasic.VbStrConv.Narrow)) & "', "
                    sql &= " '" & ddl社區大樓.SelectedValue & "', "
                    sql &= " '" & TB_AreaCode.Text & "','" & DDL_County.SelectedValue & "','" & DDL_Area.SelectedItem.Text & "', "
                    sql &= " '" & add1.Text & "' , '" & zone3.SelectedValue & "' , '" & add2.Text & "' , "
                    sql &= " N'" & add3.Text & "' ,N'" & address20.SelectedValue & "' ,'" & add4.Text & "' , "
                    sql &= " '" & add5.Text & "' ,'" & add6.Text & "' ,'" & add7.Text & "' ,'" & add8.Text & "' , "
                    sql &= " '" & add9.Text & "' , '" & add10.Text & "',N'" & address1 & "', "
                    sql &= " GETDATE() "
                    cmd = New SqlCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                End If
            End If

            'End If


            '1050506表單控管
            If ddl契約類別.SelectedValue <> "庫存" Then
                'If Request.Cookies("webfly_empno").Value = "92H" Then
                If check_formno(UCase(Mid(物件編號, 6)), store.SelectedValue, 物件編號, input4.Value, "修改") = "use" Then
                    'eip_usual.Show("一併寫入表單管理")
                Else
                    'eip_usual.Show("輸入編號不為所購買表單，無法寫入表單管理")
                End If
                'End If

            End If

            TextBox24.Text = Page.Request.Form(TextBox24.UniqueID)
            TextBox26.Text = Page.Request.Form(TextBox26.UniqueID)
            TextBox27.Text = Page.Request.Form(TextBox27.UniqueID)
            TextBox28.Text = Page.Request.Form(TextBox28.UniqueID)
            TextBox33.Text = Page.Request.Form(TextBox33.UniqueID)
            TextBox34.Text = Page.Request.Form(TextBox34.UniqueID)

        Catch ex As SqlException

            'Response.Write(sql)
            '店代號,員工編號,功能名稱,網址,語法,訊息
            checkleave.錯誤訊息(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "物件修改", Request.Url.ToString, sql, "")
            eip_usual.Show("修改物件資料失敗")
            'script += "alert('修改物件資料失敗');"
            'ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        End Try
    End Sub

    '謄本新增細項面積-在儲存時觸發
    Sub landcase_detail(ByVal objid As String)
        Dim sql2 As String
        conn1 = New SqlConnection(land)
        conn1.Open()
        sql = "select * from saleobject_areadetail With(NoLock) WHERE 物件編號='" & Request("FILENO") & "' AND 店代號 = '" & Request("sid") & "'"

        adpt = New SqlDataAdapter(sql, conn1)
        ds = New DataSet()
        adpt.SelectCommand.CommandTimeout = 200
        adpt.Fill(ds, "table1")
        Dim table1 As DataTable = ds.Tables("table1")

        conn_land = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_land.Open()

        For i = 0 To table1.Rows.Count - 1
            sql2 = "insert 委賣物件資料表_面積細項(物件編號, 流水號, 建號, 類別,項目名稱,總面積平方公尺,總面積坪,權利範圍1分母,權利範圍1分子,實際持有平方公尺,實際持有坪,店代號) values ('" & objid & "','" & table1.Rows(i)("流水號") & "','" & table1.Rows(i)("建號") & "','" & table1.Rows(i)("類別") & "','" & table1.Rows(i)("項目名稱") & "','" & table1.Rows(i)("總面積平方公尺") & "','" & table1.Rows(i)("總面積坪") & "','" & table1.Rows(i)("權利範圍1分母") & "','" & table1.Rows(i)("權利範圍1分子") & "','" & table1.Rows(i)("實際持有平方公尺") & "','" & table1.Rows(i)("實際持有坪") & "','" & Trim(store.SelectedValue) & "')"

            Dim cmd2 As New SqlCommand(sql2, conn_land)
            cmd2.CommandType = CommandType.Text
            Try
                cmd2.ExecuteNonQuery()

            Catch

            End Try
        Next
        conn1.Close()
        conn1.Dispose()
        conn_land.Close()
        conn_land.Dispose()
    End Sub

    '謄本新增他項權利
    Sub landcase_他項權利(ByVal objid As String)
        Dim 權利類別 As String = ""
        Dim 權利種類 As String = ""
        'Dim 順位 As String = ""
        Dim 登記日期 As String = ""
        Dim 設定 As String = ""
        Dim 設定權利人 As String = ""

        Dim sql2 As String
        conn1 = New SqlConnection(land)
        conn1.Open()
        sql = "select * from TG1 With(NoLock) WHERE T1FileNo='" & Request("TrueFILENO") & "' order by T1Seq"

        adpt = New SqlDataAdapter(sql, conn1)
        ds = New DataSet()
        adpt.SelectCommand.CommandTimeout = 200
        adpt.Fill(ds, "table1")
        Dim table1 As DataTable = ds.Tables("table1")

        conn_land = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_land.Open()

        For i = 0 To table1.Rows.Count - 1
            If table1.Rows(i)("T1Type").ToString.Trim = "C" Then
                權利類別 = "土地"
            Else
                權利類別 = "建物"
            End If

            權利種類 = table1.Rows(i)("T1RIGHTKIND").ToString.Trim
            '順位 = table1.Rows(i)("T1REGODR").ToString.Trim

            If Not IsDBNull(table1.Rows(i)("T1REGDT")) Then
                登記日期 = table1.Rows(i)("T1REGDT").ToString.Trim
            End If

            If Not IsDBNull(table1.Rows(i)("T1GUARAMT")) Then
                設定 = (table1.Rows(i)("T1GUARAMT") / 10000)
            End If

            If Not IsDBNull(table1.Rows(i)("T1OWNERNAME")) Then
                設定權利人 = table1.Rows(i)("T1OWNERNAME")
            End If


            sql2 = "insert 物件他項權利細項(物件編號,店代號, Num, 權利類別, 權利種類,登記日期,設定,設定權利人) values ('" & objid & "','" & Trim(store.SelectedValue) & "','" & i & "','" & 權利類別 & "','" & 權利種類 & "','" & 登記日期 & "','" & 設定 & "','" & 設定權利人 & "')"


            Dim cmd2 As New SqlCommand(sql2, conn_land)
            cmd2.CommandType = CommandType.Text
            Try
                cmd2.ExecuteNonQuery()

            Catch

            End Try
        Next
        conn1.Close()
        conn1.Dispose()
        conn_land.Close()
        conn_land.Dispose()


    End Sub




    '新增不動產說明書
    Sub 新增不動產說明書()
        Dim message As String = ""
        Dim sid As String = store.SelectedValue '選擇的店代號
        Dim oid As String = Label57.Text '物件編號



        '先查有無對應資料
        sql = "select * "
        sql &= "from 委賣_房地產說明書 With(NoLock) "
        sql &= "where 物件編號 = '" & oid & "' and 店代號 = '" & sid & "'"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        '20160225 註記
        '大部分的選項都移至 產調_建物 or 產調_土地 or 產調_車位 內 

        If table1.Rows.Count = 0 Then
            '新增
            sql = "Insert into 委賣_房地產說明書 ( "
            sql &= "物件編號, 報告單位,報告編號_年,報告編號_部門, "
            sql &= "報告編號_序號,報告日期,地政事務所所在區域,地政事務所,謄本核發日期,產權調查, "
            sql &= "物件個案調查,照片說明,其他說明,土地權狀影本,建物權狀影本,土地謄本, "
            sql &= "建物謄本,地籍圖,建物平面圖,房屋稅單,房地產標的現況說明書,都市土地使用分區域非都市土地使用種類證明, "
            sql &= "使用執照,住戶規約,預售買賣契約書,其他, "
            sql &= "與土地他項權利部相同,其他如下, " '---
            sql &= "個案特色, 物件標的, 現況, 交屋情況, 商談交屋情況, "
            sql &= "中庭花園,其他中庭花園,警衛管理,其他警衛管理,外牆外飾, "
            sql &= "其他外牆外飾,地板,其他地板,自來水,未安裝自來水原因, "
            sql &= "電力系統,有無獨立電錶,電話系統,未安裝電話系統原因,瓦斯系統, "
            sql &= "未安裝瓦斯系統,委託價格新台幣,簽約金,備証款,完稅款,"
            sql &= "尾款, 增值稅,"
            sql &= "契稅,地價稅,房屋稅,工程受益費,代書費,"
            sql &= "登記規費,公證費,印花稅,水電費,管理費, "
            sql &= "電話費,店代號,經紀人代號,營業員代號1,營業員代號2,"
            sql &= "輸入者, 上傳註記,上傳日期,分區使用證明,委託銷售契約書,"
            sql &= "建物物件與環境相關費用, 增值稅概算, " '---
            sql &= "室內建材, 建築結構, 隔間材料,建築用途,新增日期, "
            sql &= "修改日期, "
            sql &= "重新產生xml, "
            sql &= "重要交易條件,成交行情,重要設施,建物勘測成果圖,土地相關位置略圖,建物相關位置略圖,停車位位置圖,土地分管協議,建物分管協議,樑柱顯見裂痕照片,"

            sql &= "瓦斯費,奢侈稅,奢侈稅約,第一期金額,第二期金額,第三期金額,第四期金額,  "
            sql &= "實價登錄費,代書費New"
            sql &= " ) Values ("
            sql &= "@pa1,@pa2,@pa3,@pa4"
            sql &= ",@pa5,@pa6,@pa7,@pa8,@pa9,@pa10,@pa11,@pa12,@pa13,@pa14,@pa15,@pa16"
            sql &= ",@pa17,@pa18,@pa19,@pa20,@pa21,@pa22,@pa23,@pa24,@pa25,@pa26"
            sql &= ",@pa74,@pa75" '與土地他項權利部相同,其他如下, 
            sql &= ",@pa107,@pa108,@pa109,@pa110,@pa111"
            sql &= ",@pa112,@pa113,@pa114,@pa115,@pa116"
            sql &= ",@pa117,@pa118,@pa119,@pa120,@pa121"
            sql &= ",@pa122,@pa123,@pa124,@pa125,@pa126"
            sql &= ",@pa127,@pa128,@pa129,@pa130,@pa131"
            sql &= ",@pa132,@pa136"
            sql &= ",@pa137,@pa138,@pa139,@pa140,@pa141"
            sql &= ",@pa142,@pa143,@pa144,@pa145,@pa146"
            sql &= ",@pa147,@pa148,@pa149,@pa150,@pa151"
            sql &= ",@pa152,@pa153,@pa154,@pa155,@pa156"
            sql &= ",@pa157,@pa158" '建物物件與環境相關費用,增值稅概算
            sql &= ",@pa163,@pa164,@pa165,@pa169,@pa170"
            sql &= ",@pa171"
            sql &= ",@pa192"
            sql &= ",@pa196,@pa197,@pa198,@pa199,@pa200,@pa201,@pa202,@pa203,@pa204,@pa205"
            sql &= ""
            sql &= ",@pa216,@pa217,@pa219,@pa220,@pa221,@pa222,@pa223"
            sql &= ",@pa224,@pa225"
            sql &= ")"
        Else
            '修改
            sql = "update 委賣_房地產說明書 "
            sql &= "set 物件編號=@pa1 , 報告單位=@pa2, 報告編號_年=@pa3, 報告編號_部門=@pa4,   "
            sql &= "報告編號_序號=@pa5 , 報告日期=@pa6 , 地政事務所所在區域=@pa7 , 地政事務所=@pa8 , 謄本核發日期=@pa9 , 產權調查=@pa10 , "
            sql &= "物件個案調查=@pa11 , 照片說明=@pa12 , 其他說明=@pa13 , 土地權狀影本=@pa14 , 建物權狀影本=@pa15 , 土地謄本=@pa16 ,  "
            sql &= "建物謄本=@pa17 , 地籍圖=@pa18 , 建物平面圖=@pa19 , 房屋稅單=@pa20 , 房地產標的現況說明書=@pa21 , 都市土地使用分區域非都市土地使用種類證明=@pa22 ,  "
            sql &= "使用執照=@pa23 , 住戶規約=@pa24 , 預售買賣契約書=@pa25 , 其他=@pa26 , "
            sql &= "與土地他項權利部相同=@pa74 , 其他如下=@pa75 ,"
            sql &= "個案特色=@pa107 , 物件標的=@pa108 , 現況=@pa109 , 交屋情況=@pa110 , 商談交屋情況=@pa111 ,  "
            sql &= "中庭花園=@pa112 , 其他中庭花園=@pa113 , 警衛管理=@pa114 , 其他警衛管理=@pa115 , 外牆外飾=@pa116 , 其他外牆外飾=@pa117 ,  "
            sql &= "地板=@pa118 , 其他地板=@pa119 , 自來水=@pa120 , 未安裝自來水原因=@pa121 , 電力系統=@pa122 , 有無獨立電錶=@pa123 ,  "
            sql &= "電話系統=@pa124 , 未安裝電話系統原因=@pa125 , 瓦斯系統=@pa126 , 未安裝瓦斯系統=@pa127 , 委託價格新台幣=@pa128 ,  "
            sql &= "簽約金=@pa129 , 備証款=@pa130 , 完稅款=@pa131 , 尾款=@pa132 , "
            sql &= "增值稅=@pa136 , 契稅=@pa137 , 地價稅=@pa138 , 房屋稅=@pa139 , 工程受益費=@pa140 , 代書費=@pa141 ,  "
            sql &= "登記規費=@pa142 , 公證費=@pa143 , 印花稅=@pa144 , 水電費=@pa145 , 管理費=@pa146 , 電話費=@pa147 ,  "
            sql &= "店代號=@pa148 , 經紀人代號=@pa149 , 營業員代號1=@pa150 , 營業員代號2=@pa151 , 輸入者=@pa152 , 上傳註記=@pa153 ,  "
            sql &= "上傳日期=@pa154 , 分區使用證明=@pa155 , 委託銷售契約書=@pa156 , 建物物件與環境相關費用=@pa157 , 增值稅概算=@pa158 ,  "
            sql &= "室內建材=@pa163 , 建築結構=@pa164 ,  "
            sql &= "隔間材料=@pa165 , 建築用途=@pa169 , 新增日期=@pa170 ,  "

            sql &= "修改日期=@pa171 , "
            sql &= "重新產生xml=@pa192, "
            sql &= "重要交易條件=@pa196,成交行情=@pa197,重要設施=@pa198,建物勘測成果圖=@pa199,土地相關位置略圖=@pa200,建物相關位置略圖=@pa201,停車位位置圖=@pa202,土地分管協議=@pa203,建物分管協議=@pa204,樑柱顯見裂痕照片=@pa205,  "
            sql &= ""
            sql &= "瓦斯費=@pa216,奢侈稅=@pa217,奢侈稅約=@pa219,第一期金額=@pa220,第二期金額=@pa221,第三期金額=@pa222,第四期金額=@pa223, "
            sql &= "實價登錄費=@pa224,代書費New=@pa225 "
            sql &= "where 物件編號 = '" & oid & "' and 店代號 = '" & sid & "'"
        End If


        cmd = New SqlCommand(sql, conn)

        '物件編號	varchar	16	
        cmd.Parameters.Add(New SqlParameter("@pa1", SqlDbType.VarChar, 16))
        cmd.Parameters("@pa1").Value = oid
        '報告單位	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa2", SqlDbType.VarChar, 20))
        cmd.Parameters("@pa2").Value = input2.Value
        '報告編號_年	varchar	2		-----已無使用,故塞''值
        cmd.Parameters.Add(New SqlParameter("@pa3", SqlDbType.VarChar, 2))
        cmd.Parameters("@pa3").Value = ""
        '報告編號_部門	varchar	8		-----已無使用,故塞''值
        cmd.Parameters.Add(New SqlParameter("@pa4", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa4").Value = ""
        '報告編號_序號	varchar	10		-----已無使用,故塞''值
        cmd.Parameters.Add(New SqlParameter("@pa5", SqlDbType.VarChar, 10))
        cmd.Parameters("@pa5").Value = ""
        '報告日期	varchar	7		
        cmd.Parameters.Add(New SqlParameter("@pa6", SqlDbType.VarChar, 7))
        If Date5.Text = "" Then
            Date5.Text = Year(Today) - 1911 & Format(Month(Today), "0#") & Format(Day(Today), "0#")
        End If
        cmd.Parameters("@pa6").Value = Left(Date5.Text, 7)
        '地政事務所所在區域	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa7", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa7").Value = input102.Value
        '地政事務所	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa8", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa8").Value = input5.Value
        '謄本核發日期	varchar	7		
        cmd.Parameters.Add(New SqlParameter("@pa9", SqlDbType.VarChar, 7))
        cmd.Parameters("@pa9").Value = Left(Date7.Text, 7)

        '產權調查	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa10", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa10").Value = IIf(CheckBoxList1.Items(0).Selected = True, 1, 0)
        '物件個案調查	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa11", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa11").Value = IIf(CheckBoxList1.Items(1).Selected = True, 1, 0)
        '照片說明	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa12", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa12").Value = IIf(CheckBoxList1.Items(2).Selected = True, 1, 0)
        '重要交易條件	varchar	1	--------NEW	
        cmd.Parameters.Add(New SqlParameter("@pa196", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa196").Value = IIf(CheckBoxList1.Items(3).Selected = True, 1, 0)
        '成交行情	varchar	1	--------NEW			
        cmd.Parameters.Add(New SqlParameter("@pa197", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa197").Value = IIf(CheckBoxList1.Items(4).Selected = True, 1, 0)
        '重要設施	varchar	1	--------NEW			
        cmd.Parameters.Add(New SqlParameter("@pa198", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa198").Value = IIf(CheckBoxList1.Items(5).Selected = True, 1, 0)
        '其他說明	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa13", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa13").Value = IIf(CheckBoxList1.Items(6).Selected = True, 1, 0)

        '封面-17項-------------------------------------------------------------------1040701修正為20項
        '土地權狀影本	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa14", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa14").Value = IIf(CheckBoxList2.Items(0).Selected = True, 1, 0)
        '建物權狀影本	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa15", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa15").Value = IIf(CheckBoxList2.Items(1).Selected = True, 1, 0)
        '標的現況說明書	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa21", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa21").Value = IIf(CheckBoxList2.Items(2).Selected = True, 1, 0)

        '土地謄本	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa16", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa16").Value = IIf(CheckBoxList2.Items(3).Selected = True, 1, 0)
        '建物謄本	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa17", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa17").Value = IIf(CheckBoxList2.Items(4).Selected = True, 1, 0)
        '預售買賣契約書	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa25", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa25").Value = IIf(CheckBoxList2.Items(5).Selected = True, 1, 0)

        '地籍圖	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa18", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa18").Value = IIf(CheckBoxList2.Items(6).Selected = True, 1, 0)
        '建物勘測成果圖	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa199", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa199").Value = IIf(CheckBoxList2.Items(7).Selected = True, 1, 0)
        '住戶規約	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa24", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa24").Value = IIf(CheckBoxList2.Items(8).Selected = True, 1, 0)

        '土地相關位置略圖	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa200", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa200").Value = IIf(CheckBoxList2.Items(9).Selected = True, 1, 0)
        '建物相關位置略圖	varchar	1	
        cmd.Parameters.Add(New SqlParameter("@pa201", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa201").Value = IIf(CheckBoxList2.Items(10).Selected = True, 1, 0)
        '停車位位置圖	varchar	1	
        cmd.Parameters.Add(New SqlParameter("@pa202", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa202").Value = IIf(CheckBoxList2.Items(11).Selected = True, 1, 0)

        '土地分管協議	varchar	1	
        cmd.Parameters.Add(New SqlParameter("@pa203", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa203").Value = IIf(CheckBoxList2.Items(12).Selected = True, 1, 0)
        '建物分管協議	varchar	1	
        cmd.Parameters.Add(New SqlParameter("@pa204", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa204").Value = IIf(CheckBoxList2.Items(13).Selected = True, 1, 0)
        '樑柱顯見裂痕照片	varchar	1	
        cmd.Parameters.Add(New SqlParameter("@pa205", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa205").Value = IIf(CheckBoxList2.Items(14).Selected = True, 1, 0)

        '分區使用證明	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa155", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa155").Value = IIf(CheckBoxList2.Items(15).Selected = True, 1, 0)
        '房屋稅籍相關證明	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa20", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa20").Value = IIf(CheckBoxList2.Items(16).Selected = True, 1, 0)
        '其他	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa26", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa26").Value = IIf(CheckBoxList2.Items(17).Selected = True, 1, 0)

        '土地增值稅概算表	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa158", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa158").Value = IIf(CheckBoxList2.Items(18).Selected = True, 1, 0)
        '使用執照	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa23", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa23").Value = IIf(CheckBoxList2.Items(19).Selected = True, 1, 0)


        '建物平面圖	varchar	1		------新版不用
        cmd.Parameters.Add(New SqlParameter("@pa19", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa19").Value = IIf(CheckBoxList2.Items(4).Selected = True, 1, 0)   
        cmd.Parameters("@pa19").Value = 0
        '都市土地使用分區域非都市土地使用種類證明	varchar	1		------新版不用
        cmd.Parameters.Add(New SqlParameter("@pa22", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa22").Value = IIf(CheckBoxList2.Items(15).Selected = True, 1, 0)     
        cmd.Parameters("@pa22").Value = 0
        '委託銷售契約書	varchar	1		------新版不用
        cmd.Parameters.Add(New SqlParameter("@pa156", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa156").Value = IIf(CheckBoxList2.Items(12).Selected = True, 1, 0)
        cmd.Parameters("@pa156").Value = 0
        '建物物件與環境相關費用	varchar	1		------新版不用
        cmd.Parameters.Add(New SqlParameter("@pa157", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa157").Value = IIf(CheckBoxList2.Items(8).Selected = True, 1, 0)
        cmd.Parameters("@pa157").Value = 0




        '基地面積	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa27", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa27").Value = input29.Value
        '權利範圍	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa28", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa28").Value = input30.Value


        '產調-土地-------------------------------------------------------------------
        ''使用管制內容	varchar	20	'20140214修正-使用分區
        'cmd.Parameters.Add(New SqlParameter("@pa29", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa29").Value = Me.Label47.Text
        ''法定建蔽率	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa30", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa30").Value = input31.Value
        ''法定容積率	varchar	10
        'cmd.Parameters.Add(New SqlParameter("@pa31", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa31").Value = input32.Value
        ''開發限制方式	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa32", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa32").Value = DropDownList26.SelectedValue
        ''所有權型態為	varchar	8		
        'cmd.Parameters.Add(New SqlParameter("@pa33", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa33").Value = DropDownList27.SelectedValue
        ''共有土地有無分管協議	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa34", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa34").Value = DropDownList28.SelectedValue
        ''是否受限制處分	varchar	8		
        'cmd.Parameters.Add(New SqlParameter("@pa35", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa35").Value = DropDownList29.SelectedValue
        ''有無出租或占用	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa36", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa36").Value = DropDownList30.SelectedValue
        ''地目	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa168", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa168").Value = DropDownList31.SelectedValue


        '產調-建物-------------------------------------------------------------------
        ''建物權利範圍	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa37", SqlDbType.VarChar, 50))
        'cmd.Parameters("@pa37").Value = input33.Value
        ''建物若無使用執照一併說明	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa38", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa38").Value = input34.Value
        ''建物目前管理及使用情況	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa39", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa39").Value = DropDownList32.SelectedValue

        ''住戶規約內容	varchar	10	---------------------已無此選項	
        'cmd.Parameters.Add(New SqlParameter("@pa40", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa40").Value = ""

        ''專有部分之範圍	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa41", SqlDbType.VarChar, 50))
        'cmd.Parameters("@pa41").Value = input35.Value
        ''共有部分之範圍	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa42", SqlDbType.VarChar, 50))
        'cmd.Parameters("@pa42").Value = input36.Value
        ''建物有無共有約定專用部分	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa43", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa43").Value = DropDownList33.SelectedValue
        ''新版新增"有無共有約定專用部分-範圍為"
        ''新版新增"有無共有約定專用部分-使用方式"

        ''建物有無專有部分約定共用	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa44", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa44").Value = DropDownList10.SelectedValue
        ''建物範圍為	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa45", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa45").Value = input37.Value
        ''使用方式	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa46", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa46").Value = input38.Value

        ''公共基金之數額為新台幣	varchar	8 ---------------------已無此選項			
        'cmd.Parameters.Add(New SqlParameter("@pa47", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa47").Value = "" 'input39.Value
        ''其運用方式為	varchar	20	 ---------------------已無此選項		
        'cmd.Parameters.Add(New SqlParameter("@pa48", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa48").Value = "" 'input40.Value

        ''管理組織及其管理方式	varchar	10		
        ''cmd.Parameters.Add(New SqlParameter("@pa49", SqlDbType.VarChar, 10))
        ''If TextBox232.Text <> "" Then
        ''    cmd.Parameters("@pa49").Value = TextBox232.Text
        ''Else
        ''    cmd.Parameters("@pa49").Value = DropDownList14.SelectedValue
        ''End If

        ''有否使用手冊	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa50", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa50").Value = DropDownList15.SelectedValue
        ''有否受限制處分	varchar	8		
        'cmd.Parameters.Add(New SqlParameter("@pa51", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa51").Value = DropDownList34.SelectedValue
        ''有無檢測海砂	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa52", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa52").Value = DropDownList35.SelectedValue
        ''有無檢測輻射含量	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa53", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa53").Value = DropDownList18.SelectedValue


        '產調-車位-------------------------------------------------------------------
        '有否辦理單獨區分有建物登記	varchar	2		
        'cmd.Parameters.Add(New SqlParameter("@pa54", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa54").Value = DropDownList7.SelectedValue
        ''使用約定方式	varchar	40		
        'cmd.Parameters.Add(New SqlParameter("@pa55", SqlDbType.VarChar, 40))
        'cmd.Parameters("@pa55").Value = TextBox95.Text
        ''進出口為	varchar	8		
        'cmd.Parameters.Add(New SqlParameter("@pa56", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa56").Value = DropDownList23.SelectedValue
        ''車位號碼	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa57", SqlDbType.NVarChar, 100))
        'cmd.Parameters("@pa57").Value = input42.Value
        ''位置地下	varchar	3		
        'cmd.Parameters.Add(New SqlParameter("@pa58", SqlDbType.NVarChar, 100))
        'cmd.Parameters("@pa58").Value = input43.Value
        ''車位性質	nvarchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa224", SqlDbType.NVarChar, 100))
        'cmd.Parameters("@pa224").Value = DropDownList67.SelectedValue


        '他項權利---------------------------------------------------------------------
        ''權利種類1	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa59", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa59").Value = DropDownList22.SelectedValue
        ''順位1	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa60", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa60").Value = input44.Value
        ''登記日期1	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa61", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa61").Value = input45.Value
        ''設定1	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa62", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa62").Value = input46.Value
        ''設定權利人1	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa63", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa63").Value = input47.Value
        ''權利種類2	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa64", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa64").Value = DropDownList23.SelectedValue
        ''順位2	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa65", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa65").Value = input48.Value
        ''登記日期2	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa66", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa66").Value = input49.Value
        ''設定2	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa67", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa67").Value = input50.Value
        ''設定權利人2	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa68", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa68").Value = input51.Value
        ''權利種類3	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa69", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa69").Value = DropDownList24.SelectedValue
        ''順位3	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa70", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa70").Value = input52.Value
        ''登記日期3	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa71", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa71").Value = input53.Value
        ''設定3	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa72", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa72").Value = input54.Value
        ''設定權利人3	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa73", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa73").Value = input55.Value
        '與土地他項權利部相同	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa74", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa74").Value = IIf(CheckBox24.Checked = True, 1, 0)
        '其他如下	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa75", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa75").Value = IIf(CheckBox25.Checked = True, 1, 0)
        ''權利種類4	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa76", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa76").Value = DropDownList25.SelectedValue
        ''順位4	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa77", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa77").Value = input58.Value
        ''登記日期4	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa78", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa78").Value = input59.Value
        ''設定4	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa79", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa79").Value = input60.Value
        ''設定權利人4	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa80", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa80").Value = input61.Value
        ''權利種類5	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa81", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa81").Value = DropDownList26.SelectedValue
        ''順位5	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa82", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa82").Value = input62.Value
        ''登記日期5	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa83", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa83").Value = input63.Value
        ''設定5	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa84", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa84").Value = input64.Value
        ''設定權利人5	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa85", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa85").Value = input65.Value
        ''權利種類6	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa86", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa86").Value = DropDownList27.SelectedValue
        ''順位6	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa87", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa87").Value = input66.Value
        ''登記日期6	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa88", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa88").Value = input67.Value
        ''設定6	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa89", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa89").Value = input68.Value
        ''設定權利人6	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa90", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa90").Value = input69.Value
        '權利種類7	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa172", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa172").Value = DropDownList12.SelectedValue
        ''順位7	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa173", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa173").Value = input111.Value
        ''登記日期7	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa174", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa174").Value = input112.Value
        ''設定7	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa175", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa175").Value = input113.Value
        ''設定權利人7	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa176", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa176").Value = input114.Value
        ''權利種類8	varchar	10		
        'cmd.Parameters.Add(New SqlParameter("@pa177", SqlDbType.VarChar, 10))
        'cmd.Parameters("@pa177").Value = DropDownList11.SelectedValue
        ''順位8	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa178", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa178").Value = input115.Value
        ''登記日期8	varchar	7		
        'cmd.Parameters.Add(New SqlParameter("@pa179", SqlDbType.VarChar, 7))
        'cmd.Parameters("@pa179").Value = input116.Value
        ''設定8	varchar	6		
        'cmd.Parameters.Add(New SqlParameter("@pa180", SqlDbType.VarChar, 6))
        'cmd.Parameters("@pa180").Value = input117.Value
        ''設定權利人8	varchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa181", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa181").Value = input118.Value


        ''附贈設備---------------------------------------------------------------
        ''固定物	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa91", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa91").Value = IIf(CheckBox4.Checked = True, 1, 0)
        ''20151127 marked by nick
        ''電話	varchar	1	電話欄位改存流理台	
        'cmd.Parameters.Add(New SqlParameter("@pa92", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa92").Value = IIf(CheckBox29.Checked = True, 1, 0) 'IIf(CheckBox5.Checked = True, 1, 0)
        ''電話線	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa93", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa93").Value = "" 'TextBox241.Text
        ''梳妝台	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa94", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa94").Value = IIf(CheckBox6.Checked = True, 1, 0)
        ''燈飾	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa95", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa95").Value = IIf(CheckBox7.Checked = True, 1, 0)
        ''冷氣	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa96", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa96").Value = IIf(CheckBox8.Checked = True, 1, 0)
        ''冷氣台	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa97", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa97").Value = TextBox242.Text
        ''窗簾	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa98", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa98").Value = IIf(CheckBox9.Checked = True, 1, 0)
        ''床組	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa99", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa99").Value = IIf(CheckBox10.Checked = True, 1, 0)
        ''冰箱	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa100", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa100").Value = IIf(CheckBox11.Checked = True, 1, 0)
        ''冰箱台	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa101", SqlDbType.VarChar, 2))
        'cmd.Parameters("@pa101").Value = TextBox38.Text
        ''熱水器	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa102", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa102").Value = IIf(CheckBox12.Checked = True, 1, 0)
        ''沙發組	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa103", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa103").Value = IIf(CheckBox13.Checked = True, 1, 0)
        ''沙發組數 由系統櫃組數改 by nick 20151127
        'cmd.Parameters.Add(New SqlParameter("@pa185", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa185").Value = Safa_count.Text
        ''瓦斯廚具	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa104", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa104").Value = IIf(CheckBox14.Checked = True, 1, 0)
        ''瓦斯廚具樣式	varchar	8		
        'cmd.Parameters.Add(New SqlParameter("@pa105", SqlDbType.VarChar, 8))
        'cmd.Parameters("@pa105").Value = "" 'TextBox39.Text
        ''壁櫥	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa160", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa160").Value = IIf(CheckBox15.Checked = True, 1, 0)
        ''酒櫃	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa161", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa161").Value = IIf(CheckBox16.Checked = True, 1, 0)
        ''自來瓦斯	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa162", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa162").Value = IIf(CheckBox17.Checked = True, 1, 0)
        ''飲水機	nvarchar	1	
        'cmd.Parameters.Add(New SqlParameter("@pa190", SqlDbType.NVarChar, 1))
        'cmd.Parameters("@pa190").Value = "" 'IIf(CheckBox18.Checked = True, 1, 0)
        ''洗衣機	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa182", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa182").Value = IIf(CheckBox19.Checked = True, 1, 0)
        ''洗衣機台	nvarchar	2	
        'cmd.Parameters.Add(New SqlParameter("@pa188", SqlDbType.NVarChar, 2))
        'If CheckBox19.Checked = True Then
        '    If Trim(TextBox40.Text & "") <> "" Then
        '        cmd.Parameters("@pa188").Value = TextBox40.Text
        '    Else
        '        cmd.Parameters("@pa188").Value = "1"
        '    End If
        'Else
        '    cmd.Parameters("@pa188").Value = "0"
        'End If
        ''乾衣機	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa183", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa183").Value = IIf(CheckBox20.Checked = True, 1, 0)
        ''乾衣機台	nvarchar	2	
        'cmd.Parameters.Add(New SqlParameter("@pa189", SqlDbType.NVarChar, 2))
        'If CheckBox20.Checked = True Then
        '    If Trim(TextBox41.Text & "") <> "" Then
        '        cmd.Parameters("@pa189").Value = TextBox41.Text
        '    Else
        '        cmd.Parameters("@pa189").Value = "1"
        '    End If
        'Else
        '    cmd.Parameters("@pa189").Value = "0"
        'End If
        ''系統櫥櫃	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa184", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa184").Value = "" 'IIf(CheckBox21.Checked = True, 1, 0)
        ''系統櫥櫃組	varchar	1		
        ''cmd.Parameters.Add(New SqlParameter("@pa185", SqlDbType.VarChar, 1))
        ''cmd.Parameters("@pa185").Value = "" 'TextBox42.Text
        ''天然瓦斯度數表	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa186", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa186").Value = IIf(CheckBox22.Checked = True, 1, 0)
        ''其他項目	varchar	1		
        'cmd.Parameters.Add(New SqlParameter("@pa106", SqlDbType.VarChar, 1))
        'cmd.Parameters("@pa106").Value = IIf(CheckBox23.Checked = True, 1, 0)
        ''其他項目內容
        'cmd.Parameters.Add(New SqlParameter("@pa191", SqlDbType.NVarChar, 12))
        'cmd.Parameters("@pa191").Value = TextBox43.Text



        '現況說明
        '個案特色	text	16	 ---------------------已無此選項			
        cmd.Parameters.Add(New SqlParameter("@pa107", SqlDbType.Text, 16))
        cmd.Parameters("@pa107").Value = ""

        '物件標的	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa108", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa108").Value = DropDownList38.SelectedValue
        '現況	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa109", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa109").Value = DropDownList39.SelectedValue
        '交屋情況	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa110", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa110").Value = DropDownList40.SelectedValue
        '商談交屋情況	varchar	10		
        cmd.Parameters.Add(New SqlParameter("@pa111", SqlDbType.VarChar, 10))
        cmd.Parameters("@pa111").Value = input89.Value
        '中庭花園	varchar	2		
        cmd.Parameters.Add(New SqlParameter("@pa112", SqlDbType.VarChar, 2))
        cmd.Parameters("@pa112").Value = DropDownList41.SelectedValue
        '其他中庭花園	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa113", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa113").Value = input90.Value
        '警衛管理	varchar	2		
        cmd.Parameters.Add(New SqlParameter("@pa114", SqlDbType.VarChar, 2))
        cmd.Parameters("@pa114").Value = DropDownList42.SelectedValue
        '其他警衛管理	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa115", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa115").Value = input91.Value
        '外牆外飾	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa116", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa116").Value = DropDownList43.SelectedValue
        '其他外牆外飾	varchar	10		
        cmd.Parameters.Add(New SqlParameter("@pa117", SqlDbType.VarChar, 10))
        cmd.Parameters("@pa117").Value = input92.Value
        '地板	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa118", SqlDbType.VarChar, 6))
        cmd.Parameters("@pa118").Value = DropDownList44.SelectedValue
        '其他地板	varchar	10		
        cmd.Parameters.Add(New SqlParameter("@pa119", SqlDbType.VarChar, 10))
        cmd.Parameters("@pa119").Value = input93.Value
        '自來水	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa120", SqlDbType.VarChar, 6))
        cmd.Parameters("@pa120").Value = DropDownList45.SelectedValue
        '未安裝自來水原因	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa121", SqlDbType.VarChar, 20))
        cmd.Parameters("@pa121").Value = input94.Value
        '電力系統	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa122", SqlDbType.VarChar, 6))
        cmd.Parameters("@pa122").Value = DropDownList46.SelectedValue
        '有無獨立電錶	varchar	2		
        cmd.Parameters.Add(New SqlParameter("@pa123", SqlDbType.VarChar, 2))
        cmd.Parameters("@pa123").Value = DropDownList47.SelectedValue
        '室內建材	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa163", SqlDbType.VarChar, 20))
        If TextBox243.Text <> "" Then
            cmd.Parameters("@pa163").Value = TextBox243.Text
        Else
            cmd.Parameters("@pa163").Value = DropDownList48.SelectedValue
        End If
        '隔間材料	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa165", SqlDbType.VarChar, 4))
        If TextBox244.Text <> "" Then
            cmd.Parameters("@pa165").Value = TextBox244.Text
        Else
            cmd.Parameters("@pa165").Value = DropDownList49.SelectedValue
        End If
        '電話系統	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa124", SqlDbType.VarChar, 6))
        cmd.Parameters("@pa124").Value = DropDownList50.SelectedValue
        '未安裝電話系統原因	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa125", SqlDbType.VarChar, 20))
        cmd.Parameters("@pa125").Value = input95.Value
        '瓦斯系統	varchar	6		
        cmd.Parameters.Add(New SqlParameter("@pa126", SqlDbType.VarChar, 6))
        cmd.Parameters("@pa126").Value = DropDownList51.SelectedValue
        '未安裝瓦斯系統	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa127", SqlDbType.VarChar, 20))
        cmd.Parameters("@pa127").Value = input96.Value
        '建築結構	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa164", SqlDbType.VarChar, 20))
        cmd.Parameters("@pa164").Value = input97.Value



        '委託價格新台幣	varchar	10		 ---------------------已無此選項		
        cmd.Parameters.Add(New SqlParameter("@pa128", SqlDbType.VarChar, 10))
        cmd.Parameters("@pa128").Value = ""
        '付款方式------------------------------------------------------------
        '簽約金	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa129", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa129").Value = TextBox258.Text
        '第一期金額	Decimal		
        cmd.Parameters.Add(New SqlParameter("@pa220", SqlDbType.Decimal))
        cmd.Parameters("@pa220").Value = IIf(TextBox262.Text = "", 0, TextBox262.Text)
        '備証款	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa130", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa130").Value = TextBox259.Text
        '第二期金額	Decimal		
        cmd.Parameters.Add(New SqlParameter("@pa221", SqlDbType.Decimal))
        cmd.Parameters("@pa221").Value = IIf(TextBox263.Text = "", 0, TextBox263.Text)
        '完稅款	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa131", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa131").Value = TextBox260.Text
        '第三期金額	Decimal		
        cmd.Parameters.Add(New SqlParameter("@pa222", SqlDbType.Decimal))
        cmd.Parameters("@pa222").Value = IIf(TextBox264.Text = "", 0, TextBox264.Text)
        '尾款	varchar	4		
        cmd.Parameters.Add(New SqlParameter("@pa132", SqlDbType.VarChar, 4))
        cmd.Parameters("@pa132").Value = TextBox261.Text
        '第四期金額	Decimal		
        cmd.Parameters.Add(New SqlParameter("@pa223", SqlDbType.Decimal))
        cmd.Parameters("@pa223").Value = IIf(TextBox265.Text = "", 0, TextBox265.Text)

        '土地增值稅
        '自用土地增值稅約	nvarchar	20		
        'cmd.Parameters.Add(New SqlParameter("@pa133", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa133").Value = input103.Value
        '一般增值稅約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa134", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa134").Value = input104.Value
        '契稅約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa135", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa135").Value = input105.Value
        '------------------------------------------新版新增------------------------------------
        '地價稅約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa206", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa206").Value = input106.Value
        '房屋稅約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa207", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa207").Value = input107.Value
        '工程受益費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa208", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa208").Value = input108.Value
        '代書費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa209", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa209").Value = input109.Value
        '登記規費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa210", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa210").Value = input110.Value
        '公證費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa211", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa211").Value = input111.Value
        '印花稅約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa212", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa212").Value = input112.Value
        '水電費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa213", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa213").Value = input113.Value
        '管理費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa214", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa214").Value = input114.Value
        '電話費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa215", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa215").Value = input115.Value
        '瓦斯費約	nvarchar	20				
        'cmd.Parameters.Add(New SqlParameter("@pa218", SqlDbType.NVarChar, 20))
        'cmd.Parameters("@pa218").Value = input116.Value
        '奢侈稅約	nvarchar	20				
        cmd.Parameters.Add(New SqlParameter("@pa219", SqlDbType.NVarChar, 20))
        cmd.Parameters("@pa219").Value = ""
        '------------------------------------------新版新增------------------------------------




        '付費方式---------------------------------------------
        '增值稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa136", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa136").Value = DropDownList52.SelectedValue
        '契稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa137", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa137").Value = DropDownList53.SelectedValue
        '地價稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa138", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa138").Value = DropDownList54.SelectedValue
        '房屋稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa139", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa139").Value = DropDownList55.SelectedValue
        '工程受益費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa140", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa140").Value = DropDownList56.SelectedValue
        '代書費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa141", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa141").Value = DropDownList57.SelectedValue
        '登記規費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa142", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa142").Value = DropDownList58.SelectedValue
        '公證費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa143", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa143").Value = DropDownList59.SelectedValue
        '印花稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa144", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa144").Value = DropDownList60.SelectedValue
        '水電費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa145", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa145").Value = DropDownList61.SelectedValue
        '管理費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa146", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa146").Value = DropDownList62.SelectedValue
        '電話費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa147", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa147").Value = DropDownList63.SelectedValue
        '瓦斯費	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa216", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa216").Value = DropDownList13.SelectedValue
        '奢侈稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa217", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa217").Value = ""
        '增值稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa224", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa224").Value = DropDownList70.SelectedValue
        '增值稅	varchar	12		
        cmd.Parameters.Add(New SqlParameter("@pa225", SqlDbType.VarChar, 12))
        cmd.Parameters("@pa225").Value = DropDownList71.SelectedValue

        '店代號	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa148", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa148").Value = sid
        '經紀人代號	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa149", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa149").Value = ""
        '營業員代號1	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa150", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa150").Value = ""
        '營業員代號2	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa151", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa151").Value = ""
        '輸入者	varchar	8		
        cmd.Parameters.Add(New SqlParameter("@pa152", SqlDbType.VarChar, 8))
        cmd.Parameters("@pa152").Value = Request("webfly_empno")
        '上傳註記	varchar	1		
        cmd.Parameters.Add(New SqlParameter("@pa153", SqlDbType.VarChar, 1))
        cmd.Parameters("@pa153").Value = "A"
        '上傳日期	varchar	7		
        cmd.Parameters.Add(New SqlParameter("@pa154", SqlDbType.VarChar, 7))
        cmd.Parameters("@pa154").Value = sysdate

        '共有部分之範圍1	varchar	60		
        'cmd.Parameters.Add(New SqlParameter("@pa159", SqlDbType.VarChar, 60))
        'cmd.Parameters("@pa159").Value = input36.Value


        ''基地面積1	varchar	100		
        'cmd.Parameters.Add(New SqlParameter("@pa166", SqlDbType.VarChar, 100))
        'cmd.Parameters("@pa166").Value = TextBox231.Text
        ''基地面積坪1	varchar	100	
        'cmd.Parameters.Add(New SqlParameter("@pa187", SqlDbType.VarChar, 100))
        'cmd.Parameters("@pa187").Value = TextBox230.Text
        '土地權利範圍1	varchar	100		
        'cmd.Parameters.Add(New SqlParameter("@pa167", SqlDbType.Text))
        'cmd.Parameters("@pa167").Value = input30.Value


        '物件主要用途	varchar	20		
        cmd.Parameters.Add(New SqlParameter("@pa169", SqlDbType.VarChar, 20))
        If TextBox4.Text <> "" Then
            cmd.Parameters("@pa169").Value = TextBox4.Text
        Else
            cmd.Parameters("@pa169").Value = DropDownList19.SelectedValue
        End If

        '新增日期	varchar	7		
        cmd.Parameters.Add(New SqlParameter("@pa170", SqlDbType.VarChar, 7))
        If table1.Rows.Count = 0 Then
            cmd.Parameters("@pa170").Value = sysdate
        Else
            cmd.Parameters("@pa170").Value = table1.Rows(0)("新增日期").ToString
        End If

        '修改日期	varchar	7		
        cmd.Parameters.Add(New SqlParameter("@pa171", SqlDbType.VarChar, 7))
        cmd.Parameters("@pa171").Value = sysdate


        '重新產生xml
        cmd.Parameters.Add(New SqlParameter("@pa192", SqlDbType.NVarChar, 1))
        cmd.Parameters("@pa192").Value = "Y"



        ''建物範圍為_共有
        'cmd.Parameters.Add(New SqlParameter("@pa193", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa193").Value = input37_NEW.Value

        ''使用方式_共有
        'cmd.Parameters.Add(New SqlParameter("@pa194", SqlDbType.VarChar, 20))
        'cmd.Parameters("@pa194").Value = input38_NEW.Value

        '1040417新增-位置地上地下
        'cmd.Parameters.Add(New SqlParameter("@pa195", SqlDbType.NVarChar, 2))
        'cmd.Parameters("@pa195").Value = DropDownList64.SelectedValue


        '物件主要用途  
        If TextBox4.Text.Trim <> "" Then
            sql2 = "Select DISTINCT 名稱 From 不動產說明書_物件用途 With(NoLock) "
            sql2 &= "Where 名稱 = '" & TextBox4.Text.Trim & "' "
            sql2 &= " and (店代號 in ('A0001','" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "') "
            sql2 &= " or 店代號 in "
            sql2 &= " (select 店代號 from HSSTRUCTURE "
            sql2 &= " where 組別 in (select 組別 from HSSTRUCTURE where 店代號='" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "'))) "
            adpt = New SqlDataAdapter(sql2, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count > 0 Then
                'If Request.Cookies("webfly_empno").Value.ToUpper = "00P" Then
                '    Response.Write(sql2)
                '    Exit Sub
                'End If
                message &= "此建築用途已存在,請下拉選擇\n"
                TextBox4.Text = ""
            Else
                If Len(TextBox4.Text.Trim) < 20 Then
                    Dim sql2 = "insert into 不動產說明書_物件用途 (名稱,店代號) values ("
                    sql2 &= "'" & TextBox4.Text.Trim & "' , '" & sid & "')"
                    cmd2 = New SqlCommand(sql2, conn)
                    cmd2.ExecuteNonQuery()
                    cmd2.Dispose()
                Else
                    message &= "此建築用途超過長度,請修正\n"
                End If
            End If
        End If

        ''管理組織及管理方式 
        'If TextBox232.Text.Trim <> "" Then
        '    sql2 = "Select * From 資料_不動產說明書 With(NoLock) "
        '    sql2 &= "Where 類別='管理組織及管理方式' and 名稱 = '" & TextBox232.Text.Trim & "' and 店代號 in ('A0001','" & sid & "') "
        '    adpt = New SqlDataAdapter(sql2, conn)
        '    ds = New DataSet()
        '    adpt.Fill(ds, "table2")
        '    table2 = ds.Tables("table2")
        '    If table2.Rows.Count > 0 Then
        '        message &= "此管理組織及管理方式已存在,請下拉選擇\n"
        '        TextBox232.Text = ""
        '    Else
        '        If Len(TextBox232.Text.Trim) < 7 Then
        '            Dim sql2 = "insert into 資料_不動產說明書 (類別,名稱,店代號) values ("
        '            sql2 &= "'管理組織及管理方式' , '" & TextBox232.Text.Trim & "' , '" & sid & "')"
        '            cmd2 = New SqlCommand(sql2, conn)
        '            cmd2.ExecuteNonQuery()
        '            cmd2.Dispose()
        '        Else
        '            message &= "此新增管理組織及管理方式超過長度,請修正\n"
        '        End If
        '    End If
        'End If


        '室內建材 
        If TextBox243.Text.Trim <> "" Then
            If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                sql2 = " Select * From 資料_不動產說明書 With(NoLock) "
                sql2 &= " Where 類別='室內建材' and 名稱 = '" & TextBox243.Text.Trim & "' and "
                sql2 &= " (店代號 in ('A0001') or 店代號 in (select bs_dept from hsbsmg where bs_直營識別='Y' and bs_state in ('1','7'))) "
            Else
                sql2 = " Select * From 資料_不動產說明書 With(NoLock) "
                sql2 &= " Where 類別='室內建材' and 名稱 = '" & TextBox243.Text.Trim & "' and 店代號 in ('A0001','" & sid & "') "
            End If
            'sql2 = "Select * From 資料_不動產說明書 With(NoLock) "
            'sql2 &= "Where 類別='室內建材' and 名稱 = '" & TextBox243.Text.Trim & "' and 店代號 in ('A0001','" & sid & "') "
            adpt = New SqlDataAdapter(sql2, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count > 0 Then
                message &= "此室內建材已存在,請下拉選擇\n"
                TextBox243.Text = ""
            Else
                If Len(TextBox243.Text.Trim) < 7 Then
                    Dim sql2 = "insert into 資料_不動產說明書 (類別,名稱,店代號) values ("
                    sql2 &= "'室內建材' , '" & TextBox243.Text.Trim & "' , '" & sid & "')"
                    cmd2 = New SqlCommand(sql2, conn)
                    cmd2.ExecuteNonQuery()
                    cmd2.Dispose()
                Else
                    message &= "此新增室內建材超過長度,請修正\n"
                End If
            End If
        End If


        '隔間材料 
        If TextBox244.Text.Trim <> "" Then
            sql2 = "Select * From 資料_不動產說明書 With(NoLock) "
            sql2 &= "Where 類別='隔間材料' and 名稱 = '" & TextBox244.Text.Trim & "' and 店代號 in ('A0001','" & sid & "') "
            adpt = New SqlDataAdapter(sql2, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count > 0 Then
                message &= "此隔間材料已存在,請下拉選擇\n"
                TextBox244.Text = ""
            Else
                If Len(TextBox244.Text.Trim) < 7 Then
                    Dim sql2 = "insert into 資料_不動產說明書 (類別,名稱,店代號) values ("
                    sql2 &= "'隔間材料' , '" & TextBox244.Text.Trim & "' , '" & sid & "')"
                    cmd2 = New SqlCommand(sql2, conn)
                    cmd2.ExecuteNonQuery()
                    cmd2.Dispose()
                Else
                    message &= "此新增隔間材料超過長度,請修正\n"
                End If
            End If
        End If


        If message <> "" Then
            Dim Script As String = ""
            Script += "alert('" & message & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Exit Sub
        End If

        If table1.Rows.Count = 0 Then
            Try
                cmd.ExecuteNonQuery()

                '判斷是否轉頁的參數
                trans = "True"
                Dim Script As String = ""
                Script += "alert('新增成功');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Catch sqlex As Exception
                Response.Write(sqlex.Message)
                '判斷是否轉頁的參數
                trans = "False"
                Dim Script As String = ""
                Script += "alert('新增失敗-不動產說明書');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            End Try
        Else
            Try
                cmd.ExecuteNonQuery()

                '判斷是否轉頁的參數
                trans = "True"
                Dim Script As String = ""
                Script += "alert('修改成功3');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Catch sqlex As Exception
                Response.Write(sqlex.Message)
                '判斷是否轉頁的參數
                trans = "False"
                Dim Script As String = ""
                Script += "alert('修改失敗-不動產說明書." & sqlex.Message & "');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)


            End Try
        End If


        'Call 是否存檔過()
    End Sub

    Sub 面積細項_判斷編號是否相同()


        '如果編號跟店代號跟預設值不一樣，先下列步驟
        If Me.Label11.Text <> Label57.Text Or Me.Label12.Text <> store.SelectedValue Then
            Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
            Dim sql_upt As String = "update 委賣物件資料表_面積細項 set 物件編號='" & Label57.Text & "',店代號='" & store.SelectedValue & "' where 物件編號='" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"
            Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
            cmd_upt.CommandType = CommandType.Text

            conn_upt.Open()
            Try
                cmd_upt.ExecuteNonQuery()

                '判斷是否轉頁的參數
                trans = "True"

                conn_upt.Close()
                conn_upt.Dispose()

                'UPDATE已存資料後給予新的值
                '物件編號
                Me.Label11.Text = Label57.Text

                '店代號
                Me.Label12.Text = store.SelectedValue
                Dim Script As String = ""
                Script += "alert('新增成功');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Catch ex As Exception

                '判斷是否轉頁的參數
                trans = "False"

                conn_upt.Close()
                conn_upt.Dispose()
                Dim Script As String = ""
                Script += "alert('新增失敗-細項面積');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)
            End Try





        End If
    End Sub

    Sub 他項權利_判斷編號是否相同()


        '新增時如是空值，給予當下的資料值為預設值(整筆資料未存檔前)
        If Me.Label11.Text = "" And Me.Label12.Text = "" Then

            '物件編號
            Me.Label11.Text = Label57.Text

            '店代號
            Me.Label12.Text = store.SelectedValue
        End If


        '如果編號跟店代號跟預設值不一樣，先下列步驟
        If Me.Label11.Text <> Label57.Text Or Me.Label12.Text <> store.SelectedValue Then
            Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
            Dim sql_upt As String = "update 物件他項權利細項 set 物件編號='" & Label57.Text & "',店代號='" & store.SelectedValue & "' where 物件編號='" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"
            Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
            cmd_upt.CommandType = CommandType.Text


            conn_upt.Open()

            Try
                cmd_upt.ExecuteNonQuery()

                conn_upt.Close()
                conn_upt.Dispose()

                '判斷是否轉頁的參數
                trans = "True"


                'UPDATE已存資料後給予新的值
                '物件編號
                Me.Label11.Text = Label57.Text

                '店代號
                Me.Label12.Text = store.SelectedValue
                Dim Script As String = ""
                Script += "alert('新增成功');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)



            Catch ex As Exception
                '判斷是否轉頁的參數
                trans = "False"

                conn_upt.Close()
                conn_upt.Dispose()
                Dim Script As String = ""
                Script += "alert('新增失敗-他項權利');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            End Try




        End If
    End Sub

    Sub 土增稅_判斷編號是否相同()


        '新增時如是空值，給予當下的資料值為預設值(整筆資料未存檔前)
        If Me.Label11.Text = "" And Me.Label12.Text = "" Then

            '物件編號
            Me.Label11.Text = Label57.Text

            '店代號
            Me.Label12.Text = store.SelectedValue
        End If


        '如果編號跟店代號跟預設值不一樣，先下列步驟
        If Me.Label11.Text <> Label57.Text Or Me.Label12.Text <> store.SelectedValue Then
            Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)

            Dim sql_upt As String = "Update 工具_計算土地增值稅 set 物件編號='" & Label57.Text & "',店代號='" & store.SelectedValue & "' where 物件編號='" & Me.Label11.Text & "' and 店代號='" & Me.Label12.Text & "'"

            Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
            cmd_upt.CommandType = CommandType.Text


            conn_upt.Open()

            Try
                cmd_upt.ExecuteNonQuery()

                conn_upt.Close()
                conn_upt.Dispose()

                '判斷是否轉頁的參數
                trans = "True"


                'UPDATE已存資料後給予新的值
                '物件編號
                Me.Label11.Text = Label57.Text

                '店代號
                Me.Label12.Text = store.SelectedValue
                Dim Script As String = ""
                Script += "alert('新增成功');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)



            Catch ex As Exception
                '判斷是否轉頁的參數
                trans = "False"

                conn_upt.Close()
                conn_upt.Dispose()
                Dim Script As String = ""
                Script += "alert('新增失敗-土增稅');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            End Try




        End If
    End Sub


    '不動產說明書-重要交易條件
    Sub 讀入前次所填之值(ByVal sid As String)
        sql = "select TOP 1 * "
        sql &= "from 委賣_房地產說明書 With(NoLock) "
        sql &= "where 店代號 = '" & sid & "' and 增值稅 <> '' ORDER BY  新增日期 DESC,Num Desc"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count <> 0 Then
            '增值稅	varchar	12		
            DropDownList52.SelectedValue = table1.Rows(0)("增值稅").ToString
            '契稅	varchar	12		
            DropDownList53.SelectedValue = table1.Rows(0)("契稅").ToString
            '地價稅	varchar	12		
            DropDownList54.SelectedValue = table1.Rows(0)("地價稅").ToString
            '房屋稅	varchar	12		
            DropDownList55.SelectedValue = table1.Rows(0)("房屋稅").ToString
            '工程受益費	varchar	12		
            DropDownList56.SelectedValue = table1.Rows(0)("工程受益費").ToString
            '代書費	varchar	12		
            DropDownList57.SelectedValue = table1.Rows(0)("代書費").ToString
            '登記規費	varchar	12		
            DropDownList58.SelectedValue = table1.Rows(0)("登記規費").ToString
            '公證費	varchar	12		
            DropDownList59.SelectedValue = table1.Rows(0)("公證費").ToString
            '印花稅	varchar	12		
            DropDownList60.SelectedValue = table1.Rows(0)("印花稅").ToString
            '水電費	varchar	12		
            DropDownList61.SelectedValue = table1.Rows(0)("水電費").ToString
            '管理費	varchar	12		
            DropDownList62.SelectedValue = table1.Rows(0)("管理費").ToString
            '電話費	varchar	12		
            DropDownList63.SelectedValue = table1.Rows(0)("電話費").ToString
            '電話費	varchar	12		
            DropDownList13.SelectedValue = table1.Rows(0)("瓦斯費").ToString
            '增值稅	varchar	12		
            DropDownList70.SelectedValue = table1.Rows(0)("實價登錄費").ToString
            '契稅	varchar	12		
            DropDownList71.SelectedValue = table1.Rows(0)("代書費New").ToString
        End If
    End Sub


    '鎖定頁面控制項的FUNCTION
    Public Sub enablefalse(cls As String)
        Dim web_c As Boolean
        Dim html_c As Boolean
        If cls = "lock" Then
            web_c = False
            html_c = True

        ElseIf cls = "unlock" Then
            web_c = True
            html_c = False
        End If

        'VIEW1---------------------------
        '銷售狀態
        CheckBox1.Enabled = web_c
        CheckBox2.Enabled = web_c
        '物件編號
        ddl契約類別.Enabled = web_c
        TextBox2.Enabled = web_c
        Label28.Visible = web_c
        DropDownList4.Visible = web_c
        '長短案名
        input4.Disabled = html_c
        Text15.Disabled = html_c
        '刊登售價
        TextBox12.Enabled = web_c
        '物件型態
        DropDownList3.Enabled = web_c
        '建物主要用途
        DropDownList19.Enabled = web_c
        If TextBox4.Visible = html_c Then
            TextBox4.Enabled = web_c
        End If
        ''土地使用分區-主
        'DropDownList16.Enabled = web_c
        'If DropDownList17.Visible = html_c Then
        '    DropDownList17.Enabled = web_c
        'End If
        ''土地使用分區-副
        'DropDownList11.Enabled = web_c
        'If DropDownList12.Visible = html_c Then
        '    DropDownList12.Enabled = web_c
        'End If
        'ImageButton18.Visible = web_c
        'TextBox253.Enabled = web_c
        '地址
        TB_AreaCode.Enabled = web_c
        DDL_County.Enabled = web_c
        DDL_Area.Enabled = web_c
        add1.Enabled = web_c
        zone3.Enabled = web_c
        add2.Enabled = web_c
        add3.Enabled = web_c
        address20.Enabled = web_c
        add4.Enabled = web_c
        add5.Enabled = web_c
        add6.Enabled = web_c
        add7.Enabled = web_c
        add8.Enabled = web_c
        add9.Enabled = web_c
        add10.Enabled = web_c
        '土地標示
        TextBox17.Enabled = web_c
        '建物標示
        TextBox18.Enabled = web_c
        '委託起迄日期
        Date2.Enabled = web_c
        Date3.Enabled = web_c
        '銷售狀態
        DropDownList21.Enabled = web_c
        '帶看方式
        DropDownList20.Enabled = web_c

        '店代號----------------------
        store.Enabled = web_c
        all_people.Enabled = web_c
        '同群組複製店代號可以選擇
        If RadioButtonList1.SelectedValue = "flow" Then
            store.Enabled = True
            all_people.Enabled = True
        Else '1040409新增判斷
            store.Enabled = False
            all_people.Enabled = False
        End If
        '-------------------------------------

        '承辦人*3
        sale1.Enabled = web_c
        sale2.Enabled = web_c
        sale3.Enabled = web_c

        '謄本電傳
        ImageButton16.Visible = web_c

        'VIEW2-------------------------------------
        '面積細項
        ImageButton3.Visible = web_c
        ImageButton5.Visible = web_c
        '主建物
        TextBox5.Enabled = web_c
        TextBox6.Enabled = web_c
        '附屬建物
        TextBox7.Enabled = web_c
        TextBox8.Enabled = web_c
        '公共設施
        TextBox9.Enabled = web_c
        TextBox10.Enabled = web_c
        '地下室
        TextBox19.Enabled = web_c
        TextBox20.Enabled = web_c
        '公設車位
        TextBox21.Enabled = web_c
        TextBox23.Enabled = web_c
        '產權車位
        TextBox26.Enabled = web_c
        TextBox27.Enabled = web_c
        '土地坪數
        TextBox30.Enabled = web_c
        TextBox31.Enabled = web_c
        '總坪數
        TextBox28.Enabled = web_c
        TextBox29.Enabled = web_c
        '庭院坪數
        TextBox32.Enabled = web_c
        TextBox33.Enabled = web_c
        '增建坪數
        TextBox34.Enabled = web_c
        TextBox35.Enabled = web_c

        '刪除
        ImageButton2.Visible = web_c


        'VIEW3-------------------------------------
        '開放格局
        C1.Enabled = web_c
        '房
        TextBox13.Enabled = web_c
        '廳
        TextBox14.Enabled = web_c
        '衛
        TextBox15.Enabled = web_c
        '室
        TextBox16.Enabled = web_c
        '座向  
        DropDownList22.Enabled = web_c
        '地上
        TextBox88.Enabled = web_c
        '地下
        TextBox89.Enabled = web_c
        '所在
        TextBox90.Enabled = web_c
        '電梯戶數
        TextBox91.Enabled = web_c
        '幾部
        TextBox92.Enabled = web_c
        '完工年月
        Text2.Disabled = html_c
        ch1.Disabled = html_c
        '登記日期
        Text11.Disabled = html_c
        Button3.Disabled = html_c
        '管理費
        DropDownList5.Enabled = web_c
        TextBox36.Enabled = web_c
        '臨路寬
        TextBox245.Enabled = web_c

        '面寬
        TextBox39.Enabled = web_c

        '縱深
        TextBox40.Enabled = web_c
        '磁扣配對
        TextBox267.Enabled = web_c
        'VIEW4-------------------------------------
        ''車位類別
        'ddl車位類別.Enabled = web_c
        ''進出口為
        'DropDownList23.Enabled = web_c
        ''車位租售
        'DropDownList6.Enabled = web_c
        '車位價格
        input55.Disabled = html_c
        CheckBox3.Enabled = web_c
        ''車位數量
        'input53.Disabled = html_c
        ''車位號碼
        'input42.Disabled = html_c
        ''車位位置
        'input43.Disabled = html_c
        ''車位說明
        'TextBox93.Enabled = web_c
        ''車位管理費
        'DropDownList25.Enabled = web_c
        'DropDownList24.Enabled = web_c
        'TextBox94.Enabled = web_c
        ''使用約定方式
        'TextBox95.Enabled = web_c
        ''有無辦單獨區分所有建物登記
        'DropDownList7.Enabled = web_c


        'VIEW5-------------------------------------
        '商圈資訊
        TextBox96.Enabled = web_c
        Button1.Disabled = html_c
        ImageButton7.Visible = web_c
        '公園綠地
        TextBox97.Enabled = web_c
        Button2.Disabled = html_c
        ImageButton8.Visible = web_c
        '捷運
        DropDownList8.Enabled = web_c
        DropDownList9.Enabled = web_c
        '公車站牌
        input60.Disabled = html_c
        '國小
        TextBox98.Enabled = web_c
        Button4.Disabled = html_c
        '國中
        TextBox99.Enabled = web_c
        Button6.Disabled = html_c
        '高中
        TextBox100.Enabled = web_c
        Button7.Disabled = html_c
        '大專院校
        TextBox101.Enabled = web_c
        Button8.Disabled = html_c
        '社區大樓
        ddl社區大樓.Enabled = web_c
        '鑰匙編號
        input66.Disabled = html_c
        '訴求重點
        TextBox102.Enabled = web_c
        '備註
        input79.Disabled = html_c
    End Sub

    '修改
    Protected Sub ImageButton12_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton12.Click
        '20111111新增判斷,平方公尺欄位不得為空值
        Dim Str As String = chk_平方公尺()
        If Str <> "" Then
            Dim Script As String = ""
            Script += "alert('" & Str & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Exit Sub
        End If

        Dim j = 驗證數字是否阿拉伯數字()
        If j > 0 Then
            Exit Sub
        End If

        'If CType(Trim(Me.Date3.Text), Integer) < CType(Trim(Me.Date2.Text), Integer) Then
        '    Dim Script As String = ""
        '    Script += "alert('委託截止日須在委託起始日之後');"
        '    ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

        '    Exit Sub
        'End If

        'If Trim(Text2.Value) <> "" Then
        '    If Not IsDate(Left(Text2.Value, 3) + 1911 & "/" & Mid(Text2.Value, 4, 2) & "/" & "01") Then
        '        Dim Script As String = ""
        '        Script += "alert('完工年月輸入錯誤');"
        '        ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

        '        Exit Sub
        '    End If
        'End If
        'If Trim(Text11.Value) <> "" Then
        '    If Not IsDate(Left(Text11.Value, 3) + 1911 & "/" & Mid(Text11.Value, 4, 2) & "/" & mid(Text11.Value, 6, 2)) Then
        '        Dim Script As String = ""
        '        Script += "alert('登記日期輸入錯誤');"
        '        ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

        '        Exit Sub
        '    End If
        'End If

        Dim sid As String = Request("sid")
        Dim state As String = Request("state")



        '20120109新增"暫停銷售"記錄
        Insert_物件異動記錄()
        Dim amessage As String = ""
        If IsNumeric(TextBox13.Text.ToString) Or TextBox13.Text.Trim = "" Then '房
            If InStr(TextBox13.Text.ToString.ToLower, "+") = 0 And InStr(TextBox13.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox13.Text.ToString.ToLower, "-") = 0 And InStr(TextBox13.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox13.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－房僅可填入數字\n"
                    End If
                Else
                    If C1.Checked = True Then
                    Else
                        amessage &= "建物基本資料－房僅可填入數字\n"
                    End If
                End If
            Else
                amessage &= "建物基本資料－房僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－房僅可填入數字\n"
        End If

        If IsNumeric(TextBox14.Text.ToString) Or TextBox14.Text.Trim = "" Then '廳
            If InStr(TextBox14.Text.ToString.ToLower, "+") = 0 And InStr(TextBox14.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox14.Text.ToString.ToLower, "-") = 0 And InStr(TextBox14.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox14.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－廳僅可填入數字\n"
                    End If
                Else
                    amessage &= "建物基本資料－廳僅可填入數字\n"
                End If
            Else
                amessage &= "建物基本資料－廳僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－廳僅可填入數字\n"
        End If

        If IsNumeric(TextBox15.Text.ToString) Then '衛
            If InStr(TextBox15.Text.ToString.ToLower, "+") = 0 And InStr(TextBox15.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox15.Text.ToString.ToLower, "-") = 0 And InStr(TextBox15.Text.ToString.ToLower, "－") = 0 Then
                Else
                    If TextBox15.Text.Trim = "" Then
                    Else
                        amessage &= "建物基本資料－衛僅可填入數字\n"
                    End If
                End If
            Else
                amessage &= "建物基本資料－衛僅可填入數字\n"
            End If
        Else
            If IsNumeric(TextBox15.Text.Replace(".", "")) Then '允許小數點
            Else
                If TextBox15.Text.Trim = "" Then
                Else
                    amessage &= "建物基本資料－衛僅可填入數字\n"
                End If
            End If
        End If

        If IsNumeric(TextBox16.Text.ToString) Or TextBox16.Text.Trim = "" Then '室
            If InStr(TextBox16.Text.ToString.ToLower, "+") = 0 And InStr(TextBox16.Text.ToString.ToLower, "＋") = 0 Then
                If InStr(TextBox16.Text.ToString.ToLower, "-") = 0 And InStr(TextBox16.Text.ToString.ToLower, "－") = 0 Then
                    If InStr(TextBox16.Text.ToString.ToLower, ".") = 0 Then
                    Else
                        amessage &= "建物基本資料－室僅可填入數字\n"
                    End If
                Else
                    amessage &= "建物基本資料－室僅可填入數字\n"
                End If
            Else
                amessage &= "建物基本資料－室僅可填入數字\n"
            End If
        Else
            amessage &= "建物基本資料－室僅可填入數字\n"
        End If

        Dim sql As String = ""
        If Trim(TextBox267.Text) <> "" Then
            Dim 物件編號 As String = ""
            '1010630 by佩嬬
            If ddl契約類別.SelectedValue = "專任" Then
                物件編號 = "1"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                物件編號 = "6"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                物件編號 = "7"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                物件編號 = "5"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                物件編號 = "9"
            End If

            If store.SelectedValue = "請選擇" Then
                物件編號 &= Mid(sid, 2) & TextBox2.Text.Trim
            Else
                物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            End If

            If Trim(TextBox267.Text) <> "" Then
                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                sql = "select 物件編號,磁扣編號,委託截止日,銷售狀態 from 委賣物件資料表 where 磁扣編號='" & Trim(TextBox267.Text) & "' and 物件編號 <> '" & 物件編號 & "'"
                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "t1")
                table1 = ds.Tables("t1")
                If table1.Rows.Count > 0 Then
                    If Not IsDBNull(table1.Rows(0).Item("委託截止日")) Then
                        If table1.Rows(0).Item("委託截止日") < sysdate Then
                            sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        ElseIf table1.Rows(0).Item("銷售狀態") = "已成交" Then
                            sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        ElseIf table1.Rows(0).Item("銷售狀態") = "已解約" Then
                            sql = "Update 委賣物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        Else
                            amessage = "磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對"
                            eip_usual.Show(amessage)
                            conn.Close()
                            conn = Nothing
                            Exit Sub
                        End If
                    Else
                        amessage = "磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對"
                        eip_usual.Show(amessage)
                        conn.Close()
                        conn = Nothing
                        Exit Sub
                    End If
                Else
                    sql = "select 物件編號,磁扣編號,委託截止日,租賃狀態 from 委租物件資料表 where 磁扣編號='" & Trim(TextBox267.Text) & "' and 物件編號 <> '" & 物件編號 & "'"
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "t1")
                    table1 = ds.Tables("t1")
                    If table1.Rows.Count > 0 Then
                        If Not IsDBNull(table1.Rows(0).Item("委託截止日")) Then
                            If table1.Rows(0).Item("委託截止日") < sysdate Then
                                sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            ElseIf table1.Rows(0).Item("租賃狀態") = "已成交" Then
                                sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            ElseIf table1.Rows(0).Item("租賃狀態") = "已解約" Then
                                sql = "Update 委租物件資料表 set 磁扣編號='' where 磁扣編號 = '" & Trim(TextBox267.Text) & "' "
                                cmd = New SqlCommand(sql, conn)
                                cmd.ExecuteNonQuery()
                            Else
                                amessage = "磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對"
                                eip_usual.Show(amessage)
                                conn.Close()
                                conn = Nothing
                                Exit Sub
                            End If
                        Else
                            amessage = "磁扣配對失敗，無法存檔!!磁扣編號已和" & table1.Rows(0).Item("物件編號") & "配對過，請查明後再行配對"
                            eip_usual.Show(amessage)
                            conn.Close()
                            conn = Nothing
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        If DropDownList3.Text <> "土地" Then
            If TextBox88.Text = "" Or TextBox89.Text = "" Or TextBox90.Text = "" Then
                amessage &= "當不為土地時，地上、地下、所在樓層 不可為空 \n"
            End If
            If Trim(Text2.Value) = "" Then
                amessage &= "當不為土地時，完工年月 不可為空 \n"
            Else
                If Trim(Text2.Value) <> "00000" Then
                    If Not IsDate(Left(Text2.Value, 3) + 1911 & "/" & Mid(Text2.Value, 4, 2) & "/" & "01") Then
                        amessage &= "完工年月輸入錯誤 \n"
                    End If
                End If
            End If
        End If
        If Trim(Text11.Value) <> "" Then
            If Not IsDate(Left(Text11.Value, 3) + 1911 & "/" & Mid(Text11.Value, 4, 2) & "/" & Mid(Text11.Value, 6, 2)) Then
                amessage &= "登記日期輸入錯誤 \n"
            End If
        End If
        If Trim(Date2.Text) <> "" Then
            If Not IsDate(Left(Date2.Text, 3) + 1911 & "/" & Mid(Date2.Text, 4, 2) & "/" & Mid(Date2.Text, 6, 2)) Then
                amessage &= "委託起始日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託起始日期 \n"
        End If
        If Trim(Date3.Text) <> "" Then
            If Not IsDate(Left(Date3.Text, 3) + 1911 & "/" & Mid(Date3.Text, 4, 2) & "/" & Mid(Date3.Text, 6, 2)) Then
                amessage &= "委託截止日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託截止日期 \n"
        End If
        If Trim(Date2.Text) = "" And Trim(Date3.Text) = "" Then

        Else
            If CType(Left(Trim(Date3.Text), 7), Integer) < CType(Left(Trim(Date2.Text), 7), Integer) Then
                amessage &= "委託截止日須在委託起始日之後 \n"
            End If
        End If

        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Or RadioButton2.Checked = True Then

            Else
                amessage &= "當為一般約，委託編號大於AAD80001時，消費者是否願意提供個資 要強制選擇 \n"
            End If
        End If

        If Trim(DropDownList3.SelectedValue) = "預售屋" Or Trim(DropDownList3.SelectedValue) = "土地" Then

        Else
            If RadioButton4.Checked = True Then
                If CheckBox102.Checked = True Or CheckBox103.Checked = True Then

                Else
                    amessage &= "當選擇社區可養寵物，請在選擇貓或狗 \n"
                End If
            End If
        End If

        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "透天" Then

        Else
            If TextBox91.Text = "" Or TextBox92.Text = "" Then
                amessage &= "非土地時，每層戶數及電梯數不可為空 \n"
            End If
        End If


        If amessage = "" Then
            更新記錄()
            'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
            更新委託期間()
            '    'Exit Sub
            'End If
            更新車位()
        Else
            eip_usual.Show(amessage)
            Exit Sub
        End If

        '該FUNCTION含新增+修改(會自行判斷)
        新增不動產說明書()

        '判斷承辦人是否有異動，有異動則同步更新屋主的承辦人     
        If Trim(sale1.SelectedValue) <> Trim(Label59.Text) Or Trim(sale2.SelectedValue) <> Trim(Label60.Text) Or Trim(sale3.SelectedValue) <> Trim(Label61.Text) Then
            更新屋主承辦人()
            新增屋主異動紀錄()
        End If

        '-------------------------------------------------------------------------------------------------------------------
        '判斷是否有購買廣告版位
        Dim torf As String = web_no(store.SelectedValue)
        Dim num As String = index_num(Label57.Text, store.SelectedValue)

        Dim url As String = ""
        If torf = "True" Then
            url = "https://home.etwarm.com.tw/sale-" & Trim(num)
        ElseIf torf = "False" Then
            url = "https://www.etwarm.com.tw/sale-" & Trim(num)
        End If

        '寫入資料表   
        ''20190910 10.40.20.66先行拿掉==================================
        ''voice_objects("update", store.SelectedValue, Label57.Text, url, 1)
        ''20190910 10.40.20.66先行拿掉==================================
        '---------------------------------------------------------------------------------------------------------------------

        If Trim(movie_h.Text) = "Y" Then
            chk_HD(Request("sid"), Request("oid"), "u")

        End If

        UpdateAIVoiceOver(Request("oid"), store.SelectedValue)

        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        '    Response.Redirect("Obj_Update_V4.aspx?state=update&oid=" & Label57.Text & "&sid=" & store.SelectedValue & "&src=NOW&from=Upd&trans=true")
        'End If



        '檢查是否屬於icb_object_infos的物件



        If Request.Cookies("webfly_empno") Is Nothing Then
            Response.Redirect("../indexnew/login3.aspx")
        Else

            If CheckIsCoSell() Then
                '呼叫更新物件api(webhook)
                UpdateObject("1")
            End If

            'If Request.Cookies("webfly_empno").Value = "92H" Then
            '    If CheckIsCoSell() Then
            '        '呼叫更新物件api(webhook)
            '        UpdateObject("1")
            '    End If
            'End If

        End If

    End Sub

    '檢查是否屬於icb_object_infos的物件(檢查物件是否屬於太平洋的聯賣物件)
    Function CheckIsCoSell() As Boolean

        '90001AAD96740
        Using cn As New SqlConnection(EGOUPLOADSqlConnStr)
            Using cmd As New SqlCommand("select count(*) from icb_object_infos where object_no = @object_no and  agree_cb=@agree_cb", cn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@object_no", SqlDbType.NVarChar, 50)).Value = Request("oid")
                cmd.Parameters.Add(New SqlParameter("@agree_cb", SqlDbType.Bit, 1)).Value = True
                cn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                If count > 0 Then
                    Return True
                Else
                    Return False
                End If

            End Using
        End Using

    End Function

    '呼叫更新物件api(webhook)
    Sub UpdateObject(status As String)

        Dim dt As New DataTable()

        Using cn As New SqlConnection(EGOUPLOADSqlConnStr)
            Using cmd As New SqlCommand("xsp_icb_get_object_info", cn)
                cmd.CommandType = CommandType.StoredProcedure

                ' 加入參數
                cmd.Parameters.Add(New SqlParameter("@obj_no", SqlDbType.NVarChar, 50)).Value = Request("oid")
                ' 開啟連線
                cn.Open()

                ' 用 SqlDataAdapter 將資料填充到 DataTable
                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using

                If dt.Rows.Count > 0 Then
                    Dim firstRow As DataRow = dt.Rows(0)
                    Dim obj = CoObj(firstRow)

                    ' 組成事件
                    Dim eventItem As New EventItem With {
            .action = Convert.ToInt32(status),
            .user_id = firstRow("user_id"),
            .obj_id = firstRow("obj_id"),
            .payload = obj
        }

                    ' 包成 events JSON
                    Dim wrapper As New EventsWrapper With {
            .events = New List(Of EventItem) From {eventItem}
        }

                    ' 輸出 JSON
                    Dim json As String = JsonConvert.SerializeObject(wrapper, Formatting.Indented)

                    '呼叫webhook
                    CallCoObjWebhook(json)

                End If

            End Using
        End Using


    End Sub

    Function CoObj(firstRow As DataRow) As object_info

        Dim obj As New object_info()
        ' 將資料轉入 object_info
        obj.obj_id = If(IsDBNull(firstRow("obj_id")), "", firstRow("obj_id").ToString())
        obj.store_id = If(IsDBNull(firstRow("store_id")), "", firstRow("store_id").ToString())
        obj.user_id = If(IsDBNull(firstRow("user_id")), "", firstRow("user_id").ToString())
        obj.obj_no = If(IsDBNull(firstRow("obj_no")), "", firstRow("obj_no").ToString())
        obj.obj_name = If(IsDBNull(firstRow("obj_name")), "", firstRow("obj_name").ToString())
        obj.case_type = If(IsDBNull(firstRow("case_type")), 0, Convert.ToInt32(firstRow("case_type")))
        obj.sell_rent = If(IsDBNull(firstRow("sell_rent")), 0, Convert.ToInt32(firstRow("sell_rent")))
        obj.case_status = If(IsDBNull(firstRow("case_status")), 0, Convert.ToInt32(firstRow("case_status")))
        obj.ping = If(IsDBNull(firstRow("ping")), 0D, Convert.ToDecimal(firstRow("ping")))
        obj.build_ping = If(IsDBNull(firstRow("build_ping")), 0D, Convert.ToDecimal(firstRow("build_ping")))
        obj.land_ping = If(IsDBNull(firstRow("land_ping")), 0D, Convert.ToDecimal(firstRow("land_ping")))
        obj.price = If(IsDBNull(firstRow("price")), 0D, Convert.ToDecimal(firstRow("price")))
        obj.base_price = If(IsDBNull(firstRow("base_price")), 0D, Convert.ToDecimal(firstRow("base_price")))
        obj.unit_price = If(IsDBNull(firstRow("unit_price")), 0D, Convert.ToDecimal(firstRow("unit_price")))
        obj.city = If(IsDBNull(firstRow("city")), "", firstRow("city").ToString())
        obj.area = If(IsDBNull(firstRow("area")), "", firstRow("area").ToString())
        obj.area_code = If(IsDBNull(firstRow("area_code")), "", firstRow("area_code").ToString())
        obj.address = If(IsDBNull(firstRow("address")), "", firstRow("address").ToString())
        obj.road_name = If(IsDBNull(firstRow("road_name")), "", firstRow("road_name").ToString())
        obj.road_width = If(IsDBNull(firstRow("road_width")), 0D, Convert.ToDecimal(firstRow("road_width")))
        obj.gps_lng = If(IsDBNull(firstRow("gps_lng")), 0D, Convert.ToDecimal(firstRow("gps_lng")))
        obj.gps_lat = If(IsDBNull(firstRow("gps_lat")), 0D, Convert.ToDecimal(firstRow("gps_lat")))
        obj.obj_class = If(IsDBNull(firstRow("obj_class")), "", firstRow("obj_class").ToString())
        obj.legal_use = If(IsDBNull(firstRow("legal_use")), 0, Convert.ToInt32(firstRow("legal_use")))
        obj.case_use = If(IsDBNull(firstRow("case_use")), 0, Convert.ToInt32(firstRow("case_use")))
        obj.current_usage = If(IsDBNull(firstRow("current_usage")), 0, Convert.ToInt32(firstRow("current_usage")))
        obj.contract_type = If(IsDBNull(firstRow("contract_type")), 0, Convert.ToInt32(firstRow("contract_type")))
        obj.contract_date_e = If(IsDBNull(firstRow("contract_date_e")), "", firstRow("contract_date_e").ToString())
        obj.off_date = If(IsDBNull(firstRow("off_date")), "", firstRow("off_date").ToString())
        obj.look_note = If(IsDBNull(firstRow("look_note")), "", firstRow("look_note").ToString())
        obj.key_get = If(IsDBNull(firstRow("key_get")), 0, Convert.ToInt32(firstRow("key_get")))
        obj.room = If(IsDBNull(firstRow("room")), 0, Convert.ToInt32(firstRow("room")))
        obj.living = If(IsDBNull(firstRow("living")), 0, Convert.ToInt32(firstRow("living")))
        obj.bathroom = If(IsDBNull(firstRow("bathroom")), 0, Convert.ToInt32(firstRow("bathroom")))
        obj.balcony = If(IsDBNull(firstRow("balcony")), 0, Convert.ToInt32(firstRow("balcony")))
        obj.floor = If(IsDBNull(firstRow("floor")), 0, Convert.ToInt32(firstRow("floor")))
        obj.floor_to = If(IsDBNull(firstRow("floor_to")), 0, Convert.ToInt32(firstRow("floor_to")))
        obj.up_floors = If(IsDBNull(firstRow("up_floors")), 0, Convert.ToInt32(firstRow("up_floors")))
        obj.down_floors = If(IsDBNull(firstRow("down_floors")), 0, Convert.ToInt32(firstRow("down_floors")))
        obj.toward = If(IsDBNull(firstRow("toward")), "", firstRow("toward").ToString())
        obj.complete_date = If(IsDBNull(firstRow("complete_date")) OrElse String.IsNullOrWhiteSpace(firstRow("complete_date")), Nothing, firstRow("complete_date").ToString())
        obj.age = If(IsDBNull(firstRow("age")), 0, Convert.ToInt32(firstRow("age")))
        obj.build_struct = If(IsDBNull(firstRow("build_struct")), 0, Convert.ToInt32(firstRow("build_struct")))
        obj.wall_material = If(IsDBNull(firstRow("wall_material")), 0, Convert.ToInt32(firstRow("wall_material")))
        obj.land_type = Convert.ToBoolean(firstRow("land_type"))
        obj.land_urban_zone = If(IsDBNull(firstRow("land_urban_zone")), 0, Convert.ToInt32(firstRow("land_urban_zone")))
        obj.land_depth = If(IsDBNull(firstRow("land_depth")), 0D, Convert.ToDecimal(firstRow("land_depth")))
        obj.total_house = If(IsDBNull(firstRow("total_house")), 0, Convert.ToInt32(firstRow("total_house")))
        obj.house_per_floor = If(IsDBNull(firstRow("house_per_floor")), 0, Convert.ToInt32(firstRow("house_per_floor")))
        obj.cover_photo = If(IsDBNull(firstRow("cover_photo")), "", firstRow("cover_photo").ToString())
        obj.update_timestamp = CLng(firstRow("update_timestamp"))
        obj.major_space = If(IsDBNull(firstRow("major_space")), 0D, Convert.ToDecimal(firstRow("major_space")))
        obj.sub_space = If(IsDBNull(firstRow("sub_space")), 0D, Convert.ToDecimal(firstRow("sub_space")))
        obj.balcony_space = If(IsDBNull(firstRow("balcony_space")), 0D, Convert.ToDecimal(firstRow("balcony_space")))
        obj.other_space = If(IsDBNull(firstRow("other_space")), 0D, Convert.ToDecimal(firstRow("other_space")))
        obj.public_space = If(IsDBNull(firstRow("public_space")), 0D, Convert.ToDecimal(firstRow("public_space")))
        obj.pub_parking_space = If(IsDBNull(firstRow("pub_parking_space")), 0D, Convert.ToDecimal(firstRow("pub_parking_space")))
        obj.prv_parking_space = If(IsDBNull(firstRow("prv_parking_space")), 0D, Convert.ToDecimal(firstRow("prv_parking_space")))
        obj.light_side = If(IsDBNull(firstRow("light_side")), 0, Convert.ToInt32(firstRow("light_side")))
        obj.elevators = If(IsDBNull(firstRow("elevators")), 0, Convert.ToInt32(firstRow("elevators")))
        obj.manage_type = If(IsDBNull(firstRow("manage_type")), 0, Convert.ToInt32(firstRow("manage_type")))
        obj.manage_fee = If(IsDBNull(firstRow("manage_fee")), 0, Convert.ToInt32(firstRow("manage_fee")))
        obj.manage_fee_way = If(IsDBNull(firstRow("manage_fee_way")), 0, Convert.ToInt32(firstRow("manage_fee_way")))
        obj.parking_type = If(IsDBNull(firstRow("parking_type")), 0, Convert.ToInt32(firstRow("parking_type")))
        obj.parking_no = If(IsDBNull(firstRow("parking_no")), "", firstRow("parking_no").ToString())
        obj.parking_rent = If(IsDBNull(firstRow("parking_rent")), 0, Convert.ToInt32(firstRow("parking_rent")))
        obj.parking_price = If(IsDBNull(firstRow("parking_price")), 0, Convert.ToInt32(firstRow("parking_price")))
        obj.community = If(IsDBNull(firstRow("community")), "", firstRow("community").ToString())
        obj.land_section = If(IsDBNull(firstRow("land_section")), "", firstRow("land_section").ToString())
        obj.land_no = If(IsDBNull(firstRow("land_no")), "", firstRow("land_no").ToString())
        obj.land_non_urban_zone = If(IsDBNull(firstRow("land_non_urban_zone")), 0, Convert.ToInt32(firstRow("land_non_urban_zone")))
        obj.land_non_urban_zone_type = If(IsDBNull(firstRow("land_non_urban_zone_type")), 0, Convert.ToInt32(firstRow("land_non_urban_zone_type")))
        obj.build_rate = If(IsDBNull(firstRow("build_rate")), 0, Convert.ToInt32(firstRow("build_rate")))
        obj.floor_rate = If(IsDBNull(firstRow("floor_rate")), 0, Convert.ToInt32(firstRow("floor_rate")))
        obj.land_width = If(IsDBNull(firstRow("land_width")), 0, Convert.ToInt32(firstRow("land_width")))
        obj.total_build_ping = If(IsDBNull(firstRow("total_build_ping")), 0, Convert.ToInt32(firstRow("total_build_ping")))
        obj.bus_station = If(IsDBNull(firstRow("bus_station")), "", firstRow("bus_station").ToString())
        obj.mrt_station = If(IsDBNull(firstRow("mrt_station")), "", firstRow("mrt_station").ToString())
        obj.park = If(IsDBNull(firstRow("park")), "", firstRow("park").ToString())
        obj.market = If(IsDBNull(firstRow("market")), "", firstRow("market").ToString())
        obj.hospital = If(IsDBNull(firstRow("hospital")), "", firstRow("hospital").ToString())
        obj.primary_school = If(IsDBNull(firstRow("primary_school")), "", firstRow("primary_school").ToString())
        obj.junior_school = If(IsDBNull(firstRow("junior_school")), "", firstRow("junior_school").ToString())
        obj.feature = If(IsDBNull(firstRow("feature")), "", firstRow("feature").ToString())
        obj.owner_name = If(IsDBNull(firstRow("owner_name")), "", firstRow("owner_name").ToString())
        obj.owner_id = If(IsDBNull(firstRow("owner_id")), "", firstRow("owner_id").ToString())
        obj.owner_birthday = If(IsDBNull(firstRow("owner_birthday")) OrElse String.IsNullOrWhiteSpace(firstRow("owner_birthday").ToString()), Nothing, firstRow("owner_birthday").ToString())
        obj.owner_title = If(IsDBNull(firstRow("owner_title")), 0, Convert.ToInt32(firstRow("owner_title")))
        obj.owner_address = If(IsDBNull(firstRow("owner_address")), "", firstRow("owner_address").ToString())
        obj.owner_home_tel = If(IsDBNull(firstRow("owner_home_tel")), "", firstRow("owner_home_tel").ToString())
        obj.owner_office_tel = If(IsDBNull(firstRow("owner_office_tel")), "", firstRow("owner_office_tel").ToString())
        obj.owner_mobile = If(IsDBNull(firstRow("owner_mobile")), "", firstRow("owner_mobile").ToString())
        obj.owner_email = If(IsDBNull(firstRow("owner_email")), "", firstRow("owner_email").ToString())
        obj.owner_remark = If(IsDBNull(firstRow("owner_remark")), "", firstRow("owner_remark").ToString())
        obj.agent_name = If(IsDBNull(firstRow("agent_name")), "", firstRow("agent_name").ToString())
        obj.agent_id = If(IsDBNull(firstRow("agent_id")), "", firstRow("agent_id").ToString())
        obj.agent_birthday = If(IsDBNull(firstRow("agent_birthday")) OrElse String.IsNullOrWhiteSpace(firstRow("agent_birthday").ToString()), Nothing, firstRow("agent_birthday").ToString())
        obj.agent_title = If(IsDBNull(firstRow("agent_title")), 0, Convert.ToInt32(firstRow("agent_title")))
        obj.agent_address = If(IsDBNull(firstRow("agent_address")), "", firstRow("agent_address").ToString())
        obj.agent_home_tel = If(IsDBNull(firstRow("agent_home_tel")), "", firstRow("agent_home_tel").ToString())
        obj.agent_office_tel = If(IsDBNull(firstRow("agent_office_tel")), "", firstRow("agent_office_tel").ToString())
        obj.agent_mobile = If(IsDBNull(firstRow("agent_mobile")), "", firstRow("agent_mobile").ToString())
        obj.agent_email = If(IsDBNull(firstRow("agent_email")), "", firstRow("agent_email").ToString())
        obj.agent_remark = If(IsDBNull(firstRow("agent_remark")), "", firstRow("agent_remark").ToString())

        Return obj
    End Function

    Sub CallCoObjWebhook(json As String)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        'ServicePointManager.ServerCertificateValidationCallback = Function(sender, certificate, chain, sslPolicyErrors) True

        'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls

        '本機
        'Dim apiUrl As String = "https://localhost:32773/icb/object"
        '測試
        'Dim apiUrl As String = "https://superwebnew6.etwarm.com.tw/WebSite/Site2/Etwarm/icb/Ashx/Handler.ashx?action=update_object"
        '正式
        Dim apiUrl As String = "https://icb.etwarm.com.tw/icb/object"

        Dim requestapi As HttpWebRequest = CType(WebRequest.Create(apiUrl), HttpWebRequest)
        requestapi.Method = "POST"
        requestapi.ContentType = "application/json"
        requestapi.Accept = "application/json"


        Dim empnoCookie = Request.Cookies("webfly_empno")
        If empnoCookie IsNot Nothing Then
            ' 確保 CookieContainer 已初始化
            If requestapi.CookieContainer Is Nothing Then
                requestapi.CookieContainer = New CookieContainer()
            End If

            ' 建立 cookie 並加上 domain（從 URL 擷取 host）
            Dim uri As New Uri(apiUrl)
            Dim cookie As New Cookie("webfly_empno", empnoCookie.Value)
            cookie.Domain = uri.Host  ' 自動設定為 "localhost" 或實際 domain
            requestapi.CookieContainer.Add(uri, cookie)
        End If

        ' 寫入 JSON 到 request body
        Using streamWriter As New StreamWriter(requestapi.GetRequestStream())
            streamWriter.Write(json)
            streamWriter.Flush()
            streamWriter.Close()
        End Using

        ' 取得 response
        Try
            Dim response As HttpWebResponse = CType(requestapi.GetResponse(), HttpWebResponse)

            Using streamReader As New StreamReader(response.GetResponseStream())
                Dim result As String = streamReader.ReadToEnd()
                'Console.WriteLine(result)
            End Using
        Catch ex As Exception
            'Dim script As String = ""
            'script += "alert('發生錯誤：" & ex.Message.Replace("'", "\'") & "');"
            'ScriptManager.RegisterStartupScript(Me, Me.GetType(), "errorAlert", script, True)
        End Try

    End Sub

    '20120109新增"暫停銷售"記錄
    Sub Insert_物件異動記錄()
        Dim check As String = ""
        Dim 異動記錄 As String = ""
        Dim 異動記錄A As String = ""
        Dim 異動記錄B As String = ""

        '判斷是否有選取
        If CheckBox2.Checked = True Then
            check = "True"
            異動記錄B = "暫停銷售"
        Else
            check = "False"
            異動記錄B = "繼續銷售"
        End If

        If Me.Label25.Text = "True" Then
            異動記錄A = "暫停銷售"
        Else
            異動記錄A = "繼續銷售"
        End If

        異動記錄 = "異動內容: [" & 異動記錄A & "] 更改為 [" & 異動記錄B & "] "

        '判斷"暫停銷售"是否有異動
        If check <> Me.Label25.Text Then
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)

                conn.Open()

                sql = "Insert into 物件異動記錄表(物件編號,店代號,異動日期,異動人員,異動記錄,異動項目) values('" & Request("oid") & "','" & Request("sid") & "','" & Now & "','" & Request("webfly_empno") & "','" & 異動記錄 & "','暫停銷售')"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                conn.Close()
                conn.Dispose()
            End Using

        End If
    End Sub

    Sub chk_HD(ByVal sid As String, ByVal oid As String, ByVal state As String)
        ''20190910 10.40.20.66先行拿掉==================================
        ''Dim conn_MYS As New MySqlConnection(mysqletwarmstring)
        ''If conn_MYS.State = ConnectionState.Closed Then
        ''Else
        ''    sql = "select * from youtube_edit where sid ='" & sid & "' and oid ='" & oid & "' and status='n'"

        ''    Dim cmd_MYS As New MySqlCommand(sql, conn_MYS)

        ''    conn_MYS.Open()
        ''    Dim dr_MYS As MySqlDataReader = cmd_MYS.ExecuteReader

        ''    If Not dr_MYS.Read Then
        ''        add_hd(sid, oid, state)
        ''    End If

        ''    conn_MYS.Close()
        ''    conn_MYS.Dispose()

        ''End If
        ''20190910 10.40.20.66先行拿掉==================================
    End Sub

    Sub add_hd(ByVal sid As String, ByVal oid As String, ByVal state As String)
        ''20190910 10.40.20.66先行拿掉==================================
        ''Dim conn_MYS As New MySqlConnection(mysqletwarmstring)
        ''If conn_MYS.State = ConnectionState.Closed Then
        ''Else
        ''    sql = "Insert into youtube_edit (sid,oid,function) values ('" & sid & "','" & oid & "','" & state & "')"

        ''    Dim cmd_MYS As New MySqlCommand(sql, conn_MYS)

        ''    conn_MYS.Open()

        ''    cmd_MYS.ExecuteNonQuery()

        ''    conn_MYS.Close()
        ''    conn_MYS.Dispose()
        ''End If
        ''20190910 10.40.20.66先行拿掉==================================
    End Sub

    Protected Sub ImageButton16_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton16.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""


        'href = "https://superwebnew.etwarm.com.tw/oldwebfly/tool_land_search.aspx"
        href = "https://superwebnew.etwarm.com.tw/TOP_TOOLS/tool_land_search.aspx"
        ntitle = "謄本電傳"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=795, width=1060, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    'Protected Sub ImageButton11_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton11.Click
    '    Dim nscript As String
    '    Dim href As String = ""
    '    Dim ntitle As String = ""


    '    '判斷有無物件編號,無則跳出
    '    If Trim(Me.TextBox2.Text) = "" Then
    '        Dim script As String = ""
    '        script += "alert('請先輸入物件編號!!');"
    '        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
    '        Exit Sub
    '    End If


    '    '組合物件編號
    '    '物件編號-第1碼
    '    Dim 物件編號 As String = ""

    '    If Request("oid") = "" Then '新增時



    '        If ddl契約類別.SelectedValue = "專任" Then
    '            物件編號 = "1"
    '        ElseIf ddl契約類別.SelectedValue = "一般" Then
    '            物件編號 = "6"
    '        ElseIf ddl契約類別.SelectedValue = "同意書" Then
    '            物件編號 = "7"
    '        ElseIf ddl契約類別.SelectedValue = "流通" Then
    '            物件編號 = "5"
    '        ElseIf ddl契約類別.SelectedValue = "庫存" Then
    '            物件編號 = "9"
    '        End If

    '        '物件編號-第2-5碼(店代號)+第6-13碼(表單編號)
    '        If store.SelectedValue = "請選擇" Then
    '            物件編號 &= Mid(Request.Cookies("store_id").Value, 2) & UCase(Microsoft.VisualBasic.Strings.StrConv(TextBox2.Text.Trim, Microsoft.VisualBasic.VbStrConv.Narrow))
    '        Else
    '            物件編號 &= Mid(store.SelectedValue, 2) & UCase(Microsoft.VisualBasic.Strings.StrConv(TextBox2.Text.Trim, Microsoft.VisualBasic.VbStrConv.Narrow))
    '        End If

    '    Else '修改複製
    '        物件編號 = Request("oid")
    '    End If



    '    If Request("state") = "add" Then
    '        href = "../TOP_tools/tool_landcount.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 物件編號 & "&src=NOW"
    '    Else
    '        href = "../TOP_tools/tool_landcount.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 物件編號 & "&src=" & Request("src")
    '    End If

    '    ntitle = "土增稅計算"

    '    nscript = "window.open('"
    '    nscript += href
    '    nscript += " '"
    '    nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

    '    Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    'End Sub

    Protected Sub ImageButton17_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton17.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""


        '判斷有無物件編號,無則跳出
        If Trim(Me.TextBox2.Text) = "" Then
            Dim script As String = ""
            script += "alert('請先輸入物件編號!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            Exit Sub
        End If

        '組合物件編號
        '物件編號-第1碼
        Dim 物件編號 As String = ""

        If Request("oid") = "" Then '新增時



            If ddl契約類別.SelectedValue = "專任" Then
                物件編號 = "1"
            ElseIf ddl契約類別.SelectedValue = "一般" Then
                物件編號 = "6"
            ElseIf ddl契約類別.SelectedValue = "同意書" Then
                物件編號 = "7"
            ElseIf ddl契約類別.SelectedValue = "流通" Then
                物件編號 = "5"
            ElseIf ddl契約類別.SelectedValue = "庫存" Then
                物件編號 = "9"
            End If

            '物件編號-第2-5碼(店代號)+第6-13碼(表單編號)
            If store.SelectedValue = "請選擇" Then
                物件編號 &= Mid(Request.Cookies("store_id").Value, 2) & TextBox2.Text.Trim
            Else
                物件編號 &= Mid(store.SelectedValue, 2) & TextBox2.Text.Trim
            End If

        Else '修改複製
            物件編號 = Request("oid")
        End If



        If Request("state") = "add" Then
            href = "../TOP_tools/tool_landcount.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 物件編號 & "&src=NOW&landNO=" & TextBox25.Text
        Else
            href = "../TOP_tools/tool_landcount.aspx?state=" & Request("state") & "&sid=" & store.SelectedValue & "&oid=" & 物件編號 & "&src=" & Request("src") & "&landNO=" & TextBox25.Text
        End If

        ntitle = "土增稅計算"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    Sub 複製記錄()
        Dim count As Integer = 0
        Dim script, message As String
        script = ""
        message = ""

        Dim sid As String = Request("sid")
        Dim oid As String = Request("oid")

        Dim 使用分區 As String = ""
        Dim 物件編號 As String = ""
        If ddl契約類別.SelectedValue = "專任" Then
            物件編號 = "1"
        ElseIf ddl契約類別.SelectedValue = "一般" Then
            物件編號 = "6"
        ElseIf ddl契約類別.SelectedValue = "同意書" Then
            物件編號 = "7"
        ElseIf ddl契約類別.SelectedValue = "流通" Then
            物件編號 = "5"
        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            物件編號 = "9"
        End If

        物件編號 &= Mid(oid, 2, 4) & TextBox2.Text
        '物件編號 -> 新的物件編號 , oid -> 舊的物件編號
        'Response.Write(物件編號)
        'Response.Write("," & oid)
        'Response.End()

        '1010630物件編號 by佩嬬
        If RadioButtonList1.SelectedValue = "many" Or RadioButtonList1.SelectedValue = "flow" Then '一約多屋或流通件
            ddl契約類別.Enabled = False
            TextBox2.Enabled = False
        Else '不是一約多屋或流通件,即為一般複製,用新增方式檢查                 
            '判斷輸入表單編號是否正確
            Dim formtype As String
            Dim objecttype As String
            '物件用途是否為土地

            objecttype = Right(DropDownList3.SelectedValue, 1)

            If objecttype = "地" Then
                If ddl契約類別.SelectedValue = "專任" Then
                    formtype = "1233,1201,2201,1243"
                ElseIf ddl契約類別.SelectedValue = "一般" Then
                    'formtype = "1235" 因目前暫無故會以1229代
                    formtype = "1235,1229,2229,1242"
                ElseIf ddl契約類別.SelectedValue = "同意書" Then
                    formtype = "1240,1231"
                ElseIf ddl契約類別.SelectedValue = "流通" Then
                    formtype = "0000"
                ElseIf ddl契約類別.SelectedValue = "庫存" Then
                    formtype = "9999"
                End If
            Else
                If ddl契約類別.SelectedValue = "專任" Then
                    formtype = "1201,2201,1243"
                ElseIf ddl契約類別.SelectedValue = "一般" Then
                    formtype = "1229,2229,1242"
                ElseIf ddl契約類別.SelectedValue = "同意書" Then
                    formtype = "1231"
                ElseIf ddl契約類別.SelectedValue = "流通" Then
                    formtype = "0000"
                ElseIf ddl契約類別.SelectedValue = "庫存" Then
                    formtype = "9999"
                End If
            End If


            Dim message1 As String = formnocheck.checkform(formtype, TextBox2.Text, store.SelectedValue)
            If Len(message1) > 4 Then

                script = "alert('" & message1 & "');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

                Exit Sub
            End If
        End If

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        'If Len(left(trim(Date2.Text),7)) <> 7 Or Len(left(trim(Date3.Text),7)) <> 7 Then
        '    message &= "委託期間格式輸入有誤，例：1030101\n"
        'End If

        '同一家店中是否已有相同的物件編號-- 20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
        sql = " Select 物件編號 From 委賣物件資料表 With(NoLock) "
        sql &= " Where 物件編號 = '" & 物件編號 & "' "
        sql &= " and 店代號 = '" & store.SelectedValue & "' "
        sql &= " union all "
        sql &= " Select 物件編號 From 委賣物件過期資料表 With(NoLock) "
        sql &= " Where 物件編號 = '" & 物件編號 & "' "
        sql &= " and 店代號 = '" & store.SelectedValue & "' "

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")



        If table1.Rows.Count > 0 Then '需先有來源物件
            If RadioButtonList1.SelectedValue = "many" Then '一約多屋-編號不變加-流水號

                sql = "Select * From 委賣物件資料表 With(NoLock) "
                sql &= " Where 物件編號 = '" & 物件編號 & "' "
                'sql &= " and 店代號 = 'A" & Mid(物件編號.ToString.Trim, 2, 4) & "' "
                sql &= " and 店代號 = '" & store.SelectedValue & "' "
                sql &= " and 建築名稱 = '" & input4.Value & "' "
                If Right(DropDownList3.SelectedValue, 1) = "地" Then
                    sql &= " and 土地標示 = '" & TextBox17.Text & "' " '土地地號不可重覆
                Else
                    sql &= " and 建號 = '" & TextBox18.Text & "' " '其餘建號不可重覆
                End If
                sql &= " union all "
                sql &= " Select * From 委賣物件過期資料表 With(NoLock) "
                sql &= " Where 物件編號 = '" & 物件編號 & "' "
                'sql &= " and 店代號 = 'A" & Mid(物件編號.ToString.Trim, 2, 4) & "' "
                sql &= " and 店代號 = '" & store.SelectedValue & "' "
                sql &= " and 建築名稱 = '" & input4.Value & "'"
                If Right(DropDownList3.SelectedValue, 1) = "地" Then
                    sql &= " and 土地標示 = '" & TextBox17.Text & "' " '土地地號不可重覆
                ElseIf DropDownList3.SelectedValue = "預售屋" Then
                    If TextBox18.Text = "" Then
                        'message &= "契約編號為必填欄位\n"
                        'message &= "建照字號(建號)為必填欄位\n" 
                    End If
                Else
                    sql &= " and 建號 = '" & TextBox18.Text & "' " '其餘建號不可重覆
                End If

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")
                If table1.Rows.Count > 0 Then
                    message &= "已有相同的網路案名或物件資料\n"
                Else
                    sql = "Select * from (Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " union all "
                    sql &= " Select * From 委賣物件過期資料表 With(NoLock) "
                    sql &= ") as 委賣物件資料表all where 物件編號 like '" & Mid(物件編號.ToString.Trim, 1, 13) & "%'  "
                    'sql &= " and 店代號 = 'A" & Mid(物件編號.ToString.Trim, 2, 4) & "' "
                    sql &= " and 店代號 = '" & store.SelectedValue & "' "
                    sql &= "order by 物件編號 desc"
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")
                    'If "A" & Mid(物件編號.ToString.Trim, 2, 4).Trim.ToString = store.SelectedValue.Trim.ToString Then
                    If table1.Rows.Count < 10 Then
                        物件編號 = Mid(UCase(物件編號), 1, 13) & "-0" & table1.Rows.Count
                    Else
                        '1080923，不限制總筆數
                        '避免走後門，所以故意走錯誤路線，讓一約多屋最多10筆-1040121
                        '物件編號 = Mid(UCase(物件編號), 1, 13) & "-" & CType(table1.Rows.Count, String) 
                        物件編號 = Mid(UCase(物件編號), 1, 13) & "-" & Mid(table1.Rows.Count, 1, 2) '錯的
                    End If

                    '    Else
                    '    message &= "此被複製物件非選擇店之物件，請再次確認!!\n"
                    'End If
                End If
            Else
                message &= "已有相同的物件編號\n"
                conn.Close()
                conn = Nothing
            End If
        Else '沒有物件資料
            If Left(物件編號, 1) <> "9" Then '庫存件可以自行編號
                If RadioButtonList1.SelectedValue = "flow" Then '流通件-編號不變
                    'If Request.Cookies("store_id").Value = "A1051" Then                        
                    '    Response.Write(InStr(myobj.mstoreid_new("'A0095','A0913','A1051','A1150','A1170','A1188'", "A" & "1188"), "A1051"))
                    '    Response.Write(formnocheck.form_own(myobj.mstoreid_new("'A0095','A0913','A1051','A1150','A1170','A1188'", "A1051"), UCase(物件編號)))
                    '    Response.Write(InStr(myobj.mstoreidf("'A0095','A0913','A1051','A1150','A1170','A1188'", "A" & "1188"), "A1051"))
                    '    Response.Write(InStr(myobj.mstoreid_new("'A0095','A0913','A1051','A1150','A1170','A1188'", "S" & "1188"), "A1051"))
                    '    Response.Write(formnocheck.form_own(myobj.mstoreid_new("'A0095','A0913','A1051','A1150','A1170','A1188'", "A1051"), UCase(物件編號)))
                    '    Response.Write(InStr(myobj.mstoreidf("'A0095','A0913','A1051','A1150','A1170','A1188'", "S" & "1188"), "A1051"))
                    '    Response.End()
                    'End If
                    'If ((InStr(myobj.mstoreid_OLD(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0 And formnocheck.form_own(myobj.mstoreid_OLD(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) = "False" And InStr(myobj.mstoreidf(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0) And (InStr(myobj.mstoreid_OLD(myobj.mstore_id, "S" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0 And formnocheck.form_own(myobj.mstoreid_OLD(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) = "False" And InStr(myobj.mstoreidf(myobj.mstore_id, "S" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0)) Then '檢查是否為同法人或簽訂流通單
                    'If ((InStr(myobj.mstoreid_new(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0 And formnocheck.form_own(myobj.mstoreid_new(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) = "False" And InStr(myobj.mstoreidf(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0) And (InStr(myobj.mstoreid_new(myobj.mstore_id, "S" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0 And formnocheck.form_own(myobj.mstoreid_new(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) = "False" And InStr(myobj.mstoreidf(myobj.mstore_id, "S" & Mid(物件編號.ToString.Trim, 2, 4)), store.SelectedValue) = 0)) Then '檢查是否為同法人或簽訂流通單
                    '原先是依照物件編號第二碼取4位當店代號，現改成用物件編號所存的店代號去判別
                    If ((InStr(myobj.mstoreid_new(myobj.mstore_id, Request("sid")), store.SelectedValue) = 0 And formnocheck.form_own(myobj.mstoreid_new(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) = "False" And InStr(myobj.mstoreidf(myobj.mstore_id, Request("sid")), store.SelectedValue) = 0)) Then '檢查是否為同法人或簽訂流通單
                        'Response.Write(myobj.mstore_id & "_" & Request("sid"))
                        'If Request.Cookies("store_id").Value = "A1170" Then
                        '    Response.Write("A" & Mid(物件編號.ToString.Trim, 2, 4) & "法人群組" & myobj.mstoreid_new(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)) & "_秘書_合約管制檔可用" & formnocheck.form_own(myobj.mstoreid_new(myobj.mstore_id, store.SelectedValue), UCase(物件編號)) & "_A" & Mid(物件編號.ToString.Trim, 2, 4) & "流通店" & myobj.mstoreidf(myobj.mstore_id, "A" & Mid(物件編號.ToString.Trim, 2, 4)))
                        '    Response.End()
                        'End If
                        message &= "此物件非同法人之物件請先簽定流通單\n"

                    Else '同體系之間案名不可重覆

                        sql = "Select * from (Select * From 委賣物件資料表 With(NoLock) "
                        sql &= " union all "
                        sql &= " Select * From 委賣物件過期資料表 With(NoLock) "
                        sql &= ") as 委賣物件資料表all"
                        sql &= " Where 物件編號 = '" & 物件編號 & "' "
                        sql &= " and 店代號 in (" & myobj.mstore_id & ") "
                        sql &= " and 建築名稱 = '" & input4.Value & "' "

                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")
                        If table1.Rows.Count > 0 Then
                            message &= "已有相同的網路案名\n"
                        Else
                            sql = " Select * From 委賣物件資料表 With(NoLock) "
                            sql &= " Where 物件編號 = '" & 物件編號 & "' "
                            sql &= " and 店代號 = '" & store.SelectedValue & "' "
                            sql &= " union all "
                            sql &= " Select * From 委賣物件過期資料表 With(NoLock) "
                            sql &= " Where 物件編號 = '" & 物件編號 & "' "
                            sql &= " and 店代號 = '" & store.SelectedValue & "' "

                            adpt = New SqlDataAdapter(sql, conn)
                            ds = New DataSet()
                            adpt.Fill(ds, "table1")
                            table1 = ds.Tables("table1")

                            If table1.Rows.Count > 0 Then '需先有來源物件
                                message &= "已有相同的物件編號\n"
                            End If
                        End If
                    End If
                ElseIf RadioButtonList1.SelectedValue = "many" Then
                    message &= "此被複製物件非選擇店之物件，請再次確認(一約多屋)!!\n"
                Else '一般物件複製
                    '1010630 by佩嬬判斷是否為多店,再判斷是否為同一法人       
                    If myobj.Objectmstore = "1" Then
                        If formnocheck.form_own(myobj.mstoreid_OLD(myobj.mstore_id, store.SelectedValue), UCase(Mid(物件編號.Trim, 6))) = "False" Then
                            message &= "此編號不為所購買表單編號"
                        End If
                    Else
                        If formnocheck.form_own("'" & store.SelectedValue.ToString & "'", UCase(Mid(物件編號.Trim, 6))) = "False" Then
                            message &= "此編號不為所購買表單編號"
                        End If
                    End If

                End If
            End If
            conn.Close()
            conn = Nothing
        End If

        If TextBox102.Text <> "" And CheckBox2.Checked = False Then
            If Replace(TextBox102.Text.ToString, vbNewLine, "").Length > 1000 Then
                message &= "訴求重點字數上限為1000字\n"
                'ElseIf Replace(TextBox102.Text.ToString, vbNewLine, "").Length < 50 Then
                '    message &= "訴求重點字數至少輸入50字\n"
            End If
        End If

        'Response.Write("確定")
        'Exit Sub

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        '建物主要用途  
        If TextBox4.Text.Trim <> "" Then
            sql = "Select DISTINCT 名稱 From 不動產說明書_物件用途 With(NoLock) "
            sql &= "Where 名稱 = '" & TextBox4.Text.Trim & "' "
            sql &= " and (店代號 in ('A0001','" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "') "
            sql &= " or 店代號 in "
            sql &= " (select 店代號 from HSSTRUCTURE "
            sql &= " where 組別 in (select 組別 from HSSTRUCTURE where 店代號='" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "'))) "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            If table1.Rows.Count > 0 Then
                message &= "此建物主要用途已存在,請下拉選擇\n"
                TextBox4.Text = ""
            Else
                If Len(TextBox4.Text.Trim) < 50 Then
                    Dim sql = "insert into 不動產說明書_物件用途 (名稱,店代號) values ("
                    sql &= "'" & TextBox4.Text.Trim & "' , '" & IIf(sid = "", Request.Cookies("store_id").Value, sid) & "')"
                    cmd = New SqlCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                Else
                    message &= "此建物主要用途超過長度50個字,請修正\n"
                End If
            End If
        End If
        '面積細項

        '完整地址, 

        'address1完整地址,address2部份地址-到"弄"
        Dim address1 As String = "", address2 As String = ""

        '部分地址
        If add1.Text <> "" Then address2 &= add1.Text & zone3.SelectedValue
        If add2.Text <> "" Then address2 &= add2.Text & "鄰"
        If add3.Text <> "" Then address2 &= add3.Text & address20.SelectedValue

        '1040519修正
        'If add4.Text <> "" Then address2 &= add4.Text & "段"
        If add4.Text <> "" Then address2 &= add4.Text & Label64.Text

        If add5.Text <> "" Then address2 &= add5.Text & "巷"
        If add6.Text <> "" Then address2 &= add6.Text & "弄"

        address1 &= address2

        '1040519修正
        'If add7.Text <> "" Then address1 &= add7.Text & "號"
        If Label64.Text = "小段" Then
            If add7.Text <> "" Then address1 &= add7.Text & "地號"
        Else
            If add7.Text <> "" Then address1 &= add7.Text & "號"
        End If

        If add8.Text <> "" Then address1 &= "之" & add8.Text

        '20100607小豪修正("之"後方若有值,則"樓"之前加入空白,避免距離過近造成誤會,ex:101號之1   3樓)----
        If add9.Text <> "" Then
            If add8.Text <> "" Then
                address1 &= "   " & add9.Text & "樓"
            Else
                address1 &= add9.Text & "樓"
            End If
        End If
        '--------------------------------------------------------------------------------------------------
        If add10.Text <> "" Then address1 &= "之" & add10.Text

        If 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")
                    If table1.Rows.Count > 0 Then
                        For j = 0 To table1.Rows.Count - 1
                            sql = " update 委賣物件資料表 "
                            sql += " set 聯賣='N' "
                            sql += " where 物件編號 = '" & table1.Rows(j)("物件編號").ToString.Trim & "' and 店代號='" & table1.Rows(j)("店代號").ToString.Trim & "' "
                            cmd = New SqlCommand(sql, conn)
                            cmd.ExecuteNonQuery()
                        Next
                        'message &= "此區域已存在相同物件\n"
                    End If
                End If
            End If
        End If

        '2017.01.05 by Finch 聯賣規則
        If Not 物件編號.Trim.StartsWith("1") Then
            sql = "Select * From 區域聯賣成員名單 With(NoLock) "
            sql &= "Where 聯賣店代號 like '%" & store.SelectedValue & "%' And 啟用 = 'Y'"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")

            If table1.Rows.Count > 0 Then
                If table1.Rows(0)("聯賣規則代號").ToString.Trim = "3" And CheckBox101.Checked = True Then
                    Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                    sql = " Select * From 委賣物件資料表 With(NoLock) "
                    sql &= " Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                    sql &= " and isnull(聯賣,'')='Y' "
                    sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                    adpt = New SqlDataAdapter(sql, conn)
                    ds = New DataSet()
                    adpt.Fill(ds, "table1")
                    table1 = ds.Tables("table1")
                    If table1.Rows.Count > 0 Then
                        message &= "此區域已存在相同物件\n"
                    End If
                Else
                    If table1.Rows(0)("區域名稱").ToString.Trim = "宜蘭區聯賣" Then
                        Dim yilanSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & yilanSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")
                        If table1.Rows.Count > 0 Then
                            message &= "此區域已存在相同物件\n"
                        End If
                    ElseIf table1.Rows(0)("區域名稱").ToString.Trim = "高屏區聯賣" Then
                        Dim kanPingSid As String = table1.Rows(0)("聯賣店代號").ToString.Trim.Replace(",", "','")
                        sql = "Select * From 委賣物件資料表 With(NoLock) "
                        sql &= "Where 店代號 in ('" & kanPingSid & "') AND 店代號 <> '" & sid & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' "
                        sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "
                        adpt = New SqlDataAdapter(sql, conn)
                        ds = New DataSet()
                        adpt.Fill(ds, "table1")
                        table1 = ds.Tables("table1")

                        'Response.Write(sql)

                        If table1.Rows.Count > 0 Then
                            Dim queryOid As String = ""
                            For i = 0 To table1.Rows.Count - 1
                                queryOid &= table1.Rows(i)("物件編號") & "', '"
                            Next

                            sql2 = " Select TOP 1 *,convert(char, [借用時間], 111) AS 借用時間西元 From 物件鑰匙管理 With(NoLock) "
                            sql2 &= " Where 物件編號 in ('" & queryOid & "') AND 借用店 = '" & store.SelectedValue & "' "
                            sql2 &= " ORDER BY num DESC "
                            adpt = New SqlDataAdapter(sql2, conn)
                            ds = New DataSet()
                            adpt.Fill(ds, "table2")
                            table2 = ds.Tables("table2")

                            'Response.Write(sql2)

                            If table2.Rows.Count > 0 Then
                                message &= "此區域已存在相同物件，物件編號：" & table2.Rows(0)("物件編號") & "\n"
                                If table2.Rows(0)("借用項目") = "O" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾點閱此物件物調表。"
                                ElseIf table2.Rows(0)("借用項目") = "M" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件不動產說明書。"
                                ElseIf table2.Rows(0)("借用項目") = "K" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾借用此物件鑰匙。"
                                ElseIf table2.Rows(0)("借用項目") = "C" Then
                                    message &= "員工編號：" & table2.Rows(0)("借用人") & " 於" & table2.Rows(0)("借用時間西元") & "\n曾電話詢問此物件。"
                                End If
                            End If
                        End If

                        For i = 0 To table1.Rows.Count - 1

                        Next
                    End If
                End If
            End If
        End If

        '2017.01.19 by Finch 同群組物件不重複        
        sql = "Select * From HSSTRUCTURE With(NoLock) "
        sql &= "Where 店代號 = '" & store.SelectedValue & "' AND 同群物件不重覆 = 'Y'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        Dim groupSid As String = ""

        If table1.Rows.Count > 0 Then
            Dim groupId As String = table1.Rows(0)("組別").ToString.Trim
            sql = "Select * From HSSTRUCTURE With(NoLock) "
            sql &= "Where 組別 = '" & groupId & "'"

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count > 0 Then
                For i = 0 To table1.Rows.Count - 1
                    groupSid &= table1.Rows(i)("店代號").ToString.Trim & "','"
                Next

                sql = "Select * From 委賣物件資料表 With(NoLock) "
                sql &= "Where 店代號 in ('" & groupSid & "') AND 店代號 <> '" & store.SelectedValue & "' AND 縣市 = '" & DDL_County.SelectedValue & "' AND 鄉鎮市區 = '" & DDL_Area.SelectedItem.Text & "' AND 完整地址 = '" & address1 & "' AND 註記 = 'Y' "
                sql &= " AND 委託截止日 >= CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),dateadd(day, -7, getdate()),20) - 1911) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),6,2) + SUBSTRING(CONVERT(VARCHAR(10),dateadd(day, -7, getdate()),20),9,2) "

                adpt = New SqlDataAdapter(sql, conn)
                ds = New DataSet()
                adpt.Fill(ds, "table1")
                table1 = ds.Tables("table1")

                If table1.Rows.Count > 0 Then
                    message &= "同群組已存在相同物件，物件編號：" & table1.Rows(0)("物件編號") & "\n"
                End If

            End If
        End If

        '車位類別
        'If TextBox37.Text.Trim <> "" Then
        '    sql = "Select * From 資料_車位 With(NoLock) "
        '    sql &= "Where 車位 = '" & TextBox37.Text.Trim & "' and 店代號 in ('A0001','" & store.SelectedValue & "') "
        '    adpt = New SqlDataAdapter(sql, conn)
        '    ds = New DataSet()
        '    adpt.Fill(ds, "table1")
        '    table1 = ds.Tables("table1")
        '    Dim flag1 As String = ""
        '    If table1.Rows.Count > 0 Then
        '        message &= "此車位類別已存在,請下拉選擇\n"
        '        TextBox37.Text = ""
        '    Else
        '        If Len(TextBox37.Text.Trim) < 5 Then
        '            Dim sql2 = "insert into 資料_車位 (車位,店代號) values ("
        '            sql2 &= "'" & TextBox37.Text.Trim & "' , '" & Request.Cookies("store_id").Value & "')"
        '            cmd = New SqlCommand(sql2, conn)
        '            cmd.ExecuteNonQuery()
        '            cmd.Dispose()
        '        Else
        '            message &= "此新增車位類別超過長度5個字,請修正\n"
        '        End If

        '    End If
        'End If

        If DropDownList3.Text <> "土地" Then
            If TextBox88.Text = "" Or TextBox89.Text = "" Or TextBox90.Text = "" Then
                message &= "當不為土地時，地上、地下、所在樓層 不可為空 \n"
            End If
            If Trim(Text2.Value) = "" Then
                message &= "當不為土地時，完工年月 不可為空 \n"
            Else
                If Trim(Text2.Value) <> "00000" Then
                    If Not IsDate(Left(Text2.Value, 3) + 1911 & "/" & Mid(Text2.Value, 4, 2) & "/" & "01") Then
                        message &= "完工年月輸入錯誤 \n"
                    End If
                End If
            End If
        End If
        If Trim(Text11.Value) <> "" Then
            If Not IsDate(Left(Text11.Value, 3) + 1911 & "/" & Mid(Text11.Value, 4, 2) & "/" & Mid(Text11.Value, 6, 2)) Then
                message &= "登記日期輸入錯誤 \n"
            End If
        End If
        If Trim(Date2.Text) <> "" Then
            If Not IsDate(Left(Date2.Text, 3) + 1911 & "/" & Mid(Date2.Text, 4, 2) & "/" & Mid(Date2.Text, 6, 2)) Then
                message &= "委託起始日期輸入錯誤 \n"
            End If
        Else
            message &= "請輸入委託起始日期 \n"
        End If
        If Trim(Date3.Text) <> "" Then
            If Not IsDate(Left(Date3.Text, 3) + 1911 & "/" & Mid(Date3.Text, 4, 2) & "/" & Mid(Date3.Text, 6, 2)) Then
                message &= "委託截止日期輸入錯誤 \n"
            End If
        Else
            message &= "請輸入委託截止日期 \n"
        End If
        If Trim(Date2.Text) = "" And Trim(Date3.Text) = "" Then

        Else
            If CType(Left(Trim(Date3.Text), 7), Integer) < CType(Left(Trim(Date2.Text), 7), Integer) Then
                message &= "委託截止日須在委託起始日之後 \n"
            End If
        End If

        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Or RadioButton2.Checked = True Then

            Else
                message &= "當為一般約，委託編號大於AAD80001時，消費者是否願意提供個資 要強制選擇 \n"
            End If
        End If

        If message <> "" Then
            script = "alert('" & message & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            Exit Sub
        End If

        '2016.06.21 by Finch 從"委賣物件資料表_面積細項"抓第一筆土地面積的使用分區，
        '寫回"委賣物件資料表"的"物件用途"欄位(前台抓使用分區的欄位)
        'If Request.Cookies("webfly_empno").Value = "92H" Then
        'sql = "Select TOP 1 * From 委賣物件資料表_面積細項 With(NoLock) "
        'sql &= "Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' and 項目名稱 = '土地面積' and 使用分區 IS NOT NULL and 使用分區 <> '' ORDER BY 流水號 "

        'sql = "Select 使用分區 "
        'sql &= " From 委賣物件資料表_面積細項 With(NoLock) "
        'sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' and 項目名稱 = '土地面積' and 使用分區 IS NOT NULL and 使用分區 <> '' "
        'sql &= " group by 流水號 "

        'If Request.Cookies("webfly_empno").Value = "92H" Then
        '    Response.Write(sql)
        'End If

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        'If table1.Rows.Count > 0 Then
        '    使用分區 = table1.Rows(0)("使用分區")
        'Else
        '    使用分區 = ""
        'End If

        sql = "Insert into 委賣物件資料表 ( "
        sql &= "物件編號,建物型態,物件用途,物件類別,郵遞區號,商圈代號,完整地址,縣市,鄉鎮市區,村里,村里別, "
        sql &= "鄰,路名,路別,段,巷,弄,號,之,所在樓層,樓之,刊登售價,主建物,主建物平方公尺,附屬建物,附屬建物平方公尺, "
        sql &= "公共設施,公共設施平方公尺,公設內車位坪數,公設內車位平方公尺,地下室,地下室平方公尺, "
        sql &= "總坪數,總平方公尺,土地坪數,土地平方公尺,庭院坪數,庭院平方公尺,加蓋坪數,加蓋平方公尺,房,廳,衛,建築名稱,張貼卡案名,地上層數,地下層數, "
        sql &= "每層戶數,每層電梯數,竣工日期,座向,管理費,車位,車位說明,車位坪數,車位平方公尺, "
        sql &= "訴求重點,店代號,經紀人代號,營業員代號1,營業員代號2,委託起始日,委託截止日,公設比,銷售狀態, "
        sql &= "輸入者,新增日期,修改日期,修改時間,網頁刊登,強銷物件,鑰匙編號,帶看方式,土地標示,建號,上傳註記,委託售價,車位價格, "
        sql &= "車位租售,國小,國中,車位數量,國小1,國中1,部份地址,捷運,大專院校,其他學校,公車站名,銷售樓層, "
        sql &= "室,備註,公園代號,高中,高中1 "
        sql &= ",管理費單位,管理費繳交方式,車位管理費類別,車位管理費,車位管理費單位,大樓代號,複製,register_date,網站總坪數,棟別,臨路寬,面寬,縱深,磁扣編號,其他使用分區,ezcode,增建,Land_FileNo "
        'If RadioButtonList1.SelectedValue = "many" Or RadioButtonList1.SelectedValue = "flow" Then '1010502新增一約多屋及流通件註記
        sql &= ",資料來源,提供個資,聯賣,社區養寵,養貓,養狗,聯賣日期 "
        'End If
        sql &= ") Values ( "



        sql &= "'" & UCase(物件編號) & "' , "

        Label57.Text = UCase(物件編號)

        '物件主要用途
        If TextBox4.Text <> "" Then
            sql &= "'" & TextBox4.Text & "' , "
        ElseIf DropDownList19.SelectedValue <> "" Then
            sql &= "'" & DropDownList19.SelectedValue & "' , "
        Else
            sql &= "'' , "
        End If

        '2016.06.21 by Finch 使用分區寫入"物件用途"欄位
        '使用分區
        If 使用分區 <> "" And 使用分區 <> "請選擇" Then
            sql &= "'" & 使用分區 & "' , "
        Else
            sql &= "'' , "
        End If


        '物件型態-物件類別,        
        sql &= "'" & DropDownList3.SelectedValue & "' , "


        '郵遞區號, 
        sql &= "'" & TB_AreaCode.Text & "' , "

        '商圈代號, 
        If Trim(TextBox96.Text) <> "" Then
            '修正代號改以隱藏的TEXTBOX250.TEXT
            sql &= "'" & TextBox250.Text & "' , "
        Else
            sql &= "'',"
        End If



        sql &= "N'" & address1 & "' , "
        '縣市, 
        sql &= "'" & DDL_County.SelectedValue & "' , "
        '鄉鎮市區, 
        sql &= "'" & DDL_Area.SelectedItem.Text & "' , "
        '村里, 
        sql &= "'" & add1.Text & "' , "
        '村里別,
        sql &= "'" & zone3.SelectedValue & "' , "
        '鄰,
        sql &= "'" & add2.Text & "' , "
        '路名,
        sql &= "N'" & add3.Text & "' , "
        '路別, 
        sql &= "'" & address20.SelectedValue & "' , "
        '段, 
        sql &= "'" & add4.Text & "' , "
        '巷, 
        sql &= "'" & add5.Text & "' , "
        '弄, 
        sql &= "'" & add6.Text & "' , "
        '號, 
        sql &= "'" & add7.Text & "' , "
        '之
        sql &= "'" & add8.Text & "' , "
        '所在樓層,-住址的"樓"存在"所在樓層"欄位
        sql &= "'" & add9.Text & "' , "
        '樓之, 
        sql &= "'" & add10.Text & "' , "


        '刊登售價num, 
        If TextBox12.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox12.Text & " , "
        End If



        '主建物num, 
        If TextBox6.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox6.Text & " , "
        End If
        '主建物平方公尺num, 
        If TextBox5.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox5.Text & " , "
        End If

        '附屬建物num, 
        If TextBox8.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox8.Text & " , "
        End If
        '附屬建物平方公尺num,
        If TextBox7.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox7.Text & " , "
        End If

        '公共設施num,
        If TextBox10.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox10.Text & " , "
        End If
        '公共設施平方公尺num, 
        If TextBox9.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox9.Text & " , "
        End If

        '公設內車位坪數num, 
        If TextBox23.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox23.Text & " , "
        End If
        '公設內車位平方公尺num, 
        If TextBox21.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox21.Text & " , "
        End If

        '地下室num, 
        If TextBox20.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox20.Text & " , "
        End If
        '地下室平方公尺,num
        If TextBox19.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox19.Text & " , "
        End If

        '總坪數,num
        If TextBox29.Text = "" Then
            sql &= "0 ,"
        Else
            sql &= TextBox29.Text & " , "
        End If
        '總平方公尺,num 
        If TextBox28.Text = "" Then
            sql &= "0 ,"
        Else
            sql &= TextBox28.Text & " , "
        End If

        '土地坪數, num
        If TextBox31.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox31.Text & " , "
        End If
        '土地平方公尺, num
        If TextBox30.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox30.Text & " , "
        End If

        '庭院坪數, num
        If TextBox33.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox33.Text & " , "
        End If
        '庭院平方公尺, num
        If TextBox32.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox32.Text & " , "
        End If

        '加蓋坪數, num
        If TextBox35.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox35.Text & " , "
        End If
        '加蓋平方公尺, num
        If TextBox34.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox34.Text & " , "
        End If


        '房, 
        If C1.Checked = True Then
            sql &= "'-1' , "
        Else
            sql &= "'" & TextBox13.Text & "' , "
        End If
        '廳, 
        sql &= "'" & TextBox14.Text & "' , "
        '衛, num
        If TextBox15.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox15.Text & " , "
        End If


        '建築名稱, 
        sql &= "N'" & input4.Value & "' , "
        '張貼卡案名, 
        sql &= "N'" & Text15.Value & "' , "
        '地上層數, 
        sql &= "'" & TextBox88.Text & "' , "
        '地下層數,
        sql &= "'" & TextBox89.Text & "' , "
        '每層戶數, 
        sql &= "'" & TextBox91.Text & "' , "
        '每層電梯數, 
        sql &= "'" & TextBox92.Text & "' , "
        '竣工日期,
        sql &= "'" & Text2.Value & "' , "
        '座向, 
        sql &= "'" & DropDownList22.SelectedValue & "' , "

        '管理費, num 
        If DropDownList5.SelectedValue <> "無" And DropDownList5.SelectedValue <> "未知" And DropDownList5.SelectedValue <> "" Then
            If TextBox36.Text = "" Then
                sql &= "null ,"
            Else
                sql &= TextBox36.Text & " , "
            End If
        Else
            sql &= "null ,"
        End If

        '20151127 複製- 將車位選項移除
        sql &= "'' , "
        ''車位類別, 
        'If TextBox37.Text = "" Then
        '    If Trim(ddl車位類別.SelectedValue) = "請選擇" Then
        '        sql &= "'' , "
        '    Else
        '        sql &= "'" & ddl車位類別.SelectedValue & "' , "
        '    End If

        'Else
        '    sql &= "'" & TextBox37.Text & "' , "
        'End If

        '車位說明, 
        'sql &= "'" & TextBox93.Text & "' , "
        sql &= "'' , "
        '車位坪數,num
        If TextBox27.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox27.Text & " , "
        End If

        '車位平方公尺,num
        If TextBox26.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox26.Text & " , "
        End If

        '訴求重點,
        If TextBox102.Text = "特色推薦可依照最基本的「建物特色」、「附近重大交通建設」、「公園綠地」、「學區介紹」、「生活機能」等五大要點進行填寫。" Then
            sql &= "'' , "
        Else
            sql &= "N'" & TextBox102.Text & "' , "
        End If


        '店代號, 
        If store.SelectedValue <> "請選擇" Then
            sql &= "'" & store.SelectedValue & "' , "
        Else
            sql &= "'' , "
        End If

        '經紀人代號, 
        If sale1.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale1.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '營業員代號1, 
        If sale2.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale2.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '營業員代號2, 
        If sale3.SelectedValue <> "選擇人員" Then
            splitarray = Split(sale3.SelectedValue, ",")
            sql &= "'" & splitarray(0) & "' , "
        Else
            sql &= "'',"
        End If

        '委託起始日, 
        sql &= "'" & Left(Trim(Date2.Text), 7) & "' , "
        '委託截止日, 
        sql &= "'" & Left(Trim(Date3.Text), 7) & "' , "

        '公設比
        If TextBox41.Text <> "" Then
            sql &= TextBox41.Text & ", "
        Else
            sql &= "null , "
        End If

        ''1000325小豪新增-相片資料夾區分為過期(expired).有效(available)-----------------------------------------------------------------------
        'If left(trim(Date3.Text), 7) >= (Format(Year(Now.AddDays(-7)) - 1911, "00#") & Format(Month(Now.AddDays(-7)), "0#") & Format(Day(Now.AddDays(-7)), "0#")) Then
        '    Me.Label23.Text = "available"
        'Else
        '    Me.Label23.Text = "expired"
        'End If
        If Trim(Request.QueryString("src")) = "OLD" Then
            Me.Label23.Text = "expired"
        ElseIf Trim(Request("src")) = "NOW" Then
            Me.Label23.Text = "available"
        End If
        '------------------------------------------------------------------------------------------------------------------------------------


        '銷售狀態,1031211修正-複製一律為"熱賣中"
        'sql &= "'" & DropDownList21.SelectedValue & "' , "
        sql &= "'熱賣中' , "


        '輸入者,
        sql &= "'" & Label33.Text & "' , "
        '新增日期, 
        sql &= "'" & sysdate & "' , "
        '修改日期, 
        sql &= "'" & sysdate & "' , "
        '修改時間, 
        sql &= "GETDATE(), "

        If CheckBox1.Checked And CheckBox1.Visible Then
            sql &= "'是' , "
        Else
            Select Case ddl契約類別.SelectedValue
                Case "專任"
                    ' 專任約：除了直營+非管理，其他情況都為 "是"
                    If IS_DIRECTLY_OPERATION And Not IS_MANAGER_ROLE Then
                        sql &= If(CheckBox1.Checked, "'是' , ", "'否' , ")
                    Else
                        sql &= "'是' , "
                    End If
                Case "庫存"
                    sql &= "'否' , "
                Case Else
                    sql &= "'否' , "
            End Select
        End If

        '強銷物件 20151228 BY NICK
        If CheckBox5.Checked = True Then
            sql &= "'是' , "
        Else
            sql &= "'否' , "
        End If

        '鑰匙編號, 
        sql &= "'" & input66.Value & "' , "

        '帶看方式, 
        sql &= "'" & DropDownList20.SelectedValue & "' , "

        '土地標示, 
        If TextBox17.Text = "請填寫地號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox17.Text & "' , "
        End If
        '建號, 
        If TextBox18.Text = "請填寫建號" Then
            sql &= "'' , "
        Else
            sql &= "'" & TextBox18.Text & "' , "
        End If



        '上傳註記,2011123修正,複製-上傳註記一律為"A"
        sql &= "'A' , "


        '委託售價num, 
        If TextBox12.Text = "" Then
            sql &= "null ,"
        Else
            sql &= TextBox12.Text & " , "
        End If

        '車位價格,
        sql &= "'" & input55.Value & "' , "

        '車位租售,
        'sql &= "'" & DropDownList6.SelectedValue & "' , "
        sql &= "'' , "

        '國小, 
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox98.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX246.TEXT
            sql &= "'" & TextBox246.Text & "' , "
        Else
            sql &= "'',"
        End If

        '國中, 
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox99.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If

        '車位數量, 
        'sql &= "'" & input53.Value & "' , "
        sql &= "'' , "

        '國小1, 
        If Trim(TextBox98.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox98.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX98.TEXT
            sql &= "'" & TextBox98.Text & "' , "
        Else
            sql &= "'',"
        End If

        '國中1, 
        If Trim(TextBox99.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox99.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX99.TEXT
            sql &= "'" & TextBox99.Text & "' , "
        Else
            sql &= "'',"
        End If


        '部份地址, 
        sql &= "N'" & address2 & "' , "

        '捷運, 
        If DropDownList9.SelectedIndex > 0 Then
            sql &= "'" & DropDownList9.SelectedValue & "' , "
        Else
            sql &= "'',"
        End If

        '大專院校, 
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox101.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX249.TEXT
            sql &= "'" & TextBox249.Text & "' , "
        Else
            sql &= "'',"
        End If

        '其他學校-大專院校,
        If Trim(TextBox101.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox101.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取TEXTBOX101.TEXT
            sql &= "'" & TextBox101.Text & "' , "
        Else
            sql &= "'',"
        End If

        '公車站名, 
        sql &= "'" & input60.Value & "' , "

        '銷售樓層-所在樓層,
        sql &= "'" & TextBox90.Text & "' , "

        '室,
        sql &= "'" & TextBox16.Text & "' , "

        '備註, 
        sql &= "'" & input79.Value & "' , "

        '公園代號,        
        If Trim(TextBox97.Text) <> "" Then

            '修正代號改以隱藏的TEXTBOX251.TEXT
            Dim parkCode As String = String.Join(".", TextBox251.Text.Split(".").Distinct().ToArray())      '
            sql &= "'" & parkCode & "' , "
        Else
            sql &= "'',"
        End If



        '高中,
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox100.UniqueID), ",")
            'sql &= "'" & splitarray(1) & "' , "

            '修正代號直接讀取隱藏的TEXTBOX247.TEXT
            sql &= "'" & TextBox247.Text & "' , "
        Else
            sql &= "'',"
        End If

        '高中1, 
        If Trim(TextBox100.Text) <> "" Then
            'splitarray = Split(Page.Request.Form(TextBox100.UniqueID), ",")
            'sql &= "'" & splitarray(0) & "' , "

            '修正校名直接讀取的TEXTBOX100.TEXT
            sql &= "'" & TextBox100.Text & "' , "
        Else
            sql &= "'' , "
        End If

        '管理費單位,
        sql &= "'" & DropDownList5.SelectedValue & "' , "

        '管理費繳交方式,
        sql &= "'" & TextBox266.Text & "' , "

        '車位管理費類別,
        'sql &= "'" & DropDownList25.SelectedValue & "' , "
        sql &= "'' , "
        '車位管理費, num
        sql &= "null ,"
        'If DropDownList24.SelectedValue <> "無" And DropDownList24.SelectedValue <> "未知" And DropDownList24.SelectedValue <> "" Then
        '    If TextBox94.Text = "" Then
        '        sql &= "null ,"
        '    Else
        '        sql &= TextBox94.Text & " , "
        '    End If
        'Else
        '    sql &= "null ,"
        'End If


        '車位管理費單位
        'sql &= "'" & DropDownList24.SelectedValue & "', "
        sql &= "'', "
        '大樓代號
        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "其他" Then
            sql &= "'',"
        Else
            If ddl社區大樓.SelectedIndex > 0 Then
                sql &= "'" & ddl社區大樓.SelectedValue & "',"
            Else
                sql &= "'',"
            End If
        End If

        '複製註記 
        sql &= "'Y', "

        '登記日期20110620
        If Trim(Text11.Value) <> "" Then
            sql &= "'" & Left(Trim(Text11.Value), 7) & "', "
        Else
            sql &= "'', "
        End If


        '網站總坪數20110803
        '網站總坪數 20151231 修正不管是否含於開價中 總坪數皆需加入
        If IsNumeric(TextBox26.Text) Then
            Dim sum1 As Double = 0

            If IsNumeric(TextBox28.Text) Then sum1 += TextBox28.Text '總坪數
            If IsNumeric(TextBox26.Text) Then sum1 += TextBox26.Text '產權獨立車位面積

            sql &= Round(sum1 * 0.3025, 4) & ","
        Else
            '總坪數,num
            If TextBox29.Text = "" Then
                sql &= "0,"
            Else
                sql &= TextBox29.Text & ","
            End If

        End If

        '棟別
        If Trim(add11.Text) <> "" Then
            sql &= "'" & Trim(add11.Text) & "', "
        Else
            sql &= "'', "
        End If

        '臨路寬
        If Trim(TextBox245.Text) <> "" Then
            sql &= "'" & Trim(TextBox245.Text) & "', "
        Else
            sql &= "'', "
        End If

        '面寬
        If Trim(TextBox39.Text) <> "" Then
            sql &= "'" & Trim(TextBox39.Text) & "', "
        Else
            sql &= "'', "
        End If

        '縱深
        If Trim(TextBox40.Text) <> "" Then
            sql &= "'" & Trim(TextBox40.Text) & "', "
        Else
            sql &= "'', "
        End If
        '磁扣編號
        If Trim(TextBox267.Text) <> "" Then
            sql &= "'" & Trim(TextBox267.Text) & "', "
        Else
            sql &= "'', "
        End If
        '其他使用分區
        'If Trim(TextBox253.Text) <> "" Then
        '    sql &= "'" & Trim(TextBox253.Text) & "', "
        'Else
        '    sql &= "'', "
        'End If
        sql &= "'', "

        '簡碼
        If Left(物件編號, 1) <> "9" Then
            Dim ezcode As String = Trim(get_num(store.SelectedValue))
            sql &= "'" & ezcode & "', "
        Else
            sql &= "'', "
        End If

        '增建
        sql &= "'" & RadioButtonList3.SelectedValue & "', "

        'Land_FileNo
        sql &= "'" & Land_FileNo.Text & "' "

        Dim 資料來源 As String = ""

        Select Case Trim(Data_Source.Text)
            Case "謄本", "二類", "電傳"
                資料來源 = Trim(Data_Source.Text) & ","
            Case Else
                資料來源 = ""
        End Select

        '1010502新增一約多屋及流通件註記
        If RadioButtonList1.SelectedValue = "many" Then
            sql &= ",'" & 資料來源 & "一約多屋'"
        ElseIf RadioButtonList1.SelectedValue = "flow" Then
            sql &= ",'" & 資料來源 & "同法人物件'"
        Else
            sql &= ",'" & 資料來源 & "'"
        End If

        '提供個資
        If (Mid(TextBox2.Text.ToUpper, 1, 1) = "A" And TextBox2.Text.ToUpper >= "AAD80001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "B" And TextBox2.Text.ToUpper >= "BAB60001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "F" And TextBox2.Text.ToUpper >= "FAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "L" And TextBox2.Text.ToUpper >= "LAA36001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "M" And TextBox2.Text.ToUpper >= "MAA38001") Or (Mid(TextBox2.Text.ToUpper, 1, 1) = "N" And TextBox2.Text.ToUpper >= "NAA12001") Then
            If RadioButton1.Checked = True Then
                sql &= ",'Y' "
            Else
                sql &= ",'N' "
            End If
        Else
            '當不屬於一般約及委託編號<AAD80001時，直接帶入空值
            sql &= ",'' "
        End If

        '聯賣(複製預設不聯賣)
        sql &= " ,'N' "
        Label465.Text = "N"
        'If CheckBox101.visible = False Then
        '    sql &= " ,'N' "
        '    Label465.text = "N"
        'Else
        '    If CheckBox101.Checked = True Then
        '        sql &= " ,'Y' "
        '        Label465.text = "Y"
        '    Else
        '        sql &= " ,'N' "
        '        Label465.text = "N"
        '    End If
        'End If

        '社區養寵
        If Trim(DropDownList3.SelectedValue) = "預售屋" Or Trim(DropDownList3.SelectedValue) = "土地" Then
            sql &= " ,'0','0','0' "
        Else
            If RadioButton4.Checked = True Then
                sql &= " ,'1' "
                If CheckBox102.Checked = True Then
                    sql &= " ,'1' "
                Else
                    sql &= " ,'0' "
                End If
                If CheckBox103.Checked = True Then
                    sql &= " ,'1' "
                Else
                    sql &= " ,'0' "
                End If
            Else
                sql &= " ,'0','0','0' "
            End If
        End If

        '聯賣日期(複製預設不聯賣)
        sql &= " ,NULL "
        'If CheckBox101.visible = False Then
        '    sql &= " ,NULL "
        'Else
        '    If CheckBox101.Checked = True Then
        '        sql &= " ,GETDATE() "
        '    Else
        '        sql &= " ,NULL "
        '    End If
        'End If

        sql &= ") "


        ''貸款銀行, 
        'sql &= "'" & input50.Value & "' , "
        ''貸款金額
        'If Trim(input80.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= input80.Value & " , "
        'End If
        ''貸款銀行2, 
        'sql &= "'" & Text5.Value & "' , "
        ''貸款金額2
        'If Trim(Text6.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text6.Value & " , "
        'End If
        ''貸款銀行3, 
        'sql &= "'" & Text7.Value & "' , "
        ''貸款金額3
        'If Trim(Text8.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text8.Value & " , "
        'End If
        ''貸款銀行4, 
        'sql &= "'" & Text9.Value & "' , "
        ''貸款金額4
        'If Trim(Text10.Value & "") = "" Then
        '    sql &= "0 , "
        'Else
        '    sql &= Text10.Value & " , "
        'End If



        cmd = New SqlCommand(sql, conn)
        Dim message2 As String
        'If Request.Cookies("webfly_empno").Value = "09W6" Then
        '    Response.Write("A_" & sql & "_A")
        '    Exit Sub
        'End If


        'Try
        If RadioButtonList1.SelectedValue <> "again" Then
            cmd.ExecuteNonQuery()
        End If

        If RadioButtonList1.SelectedValue = "many" Or RadioButtonList1.SelectedValue = "flow" Or ddl契約類別.SelectedValue = "流通" Or ddl契約類別.SelectedValue = "庫存" Then '一般複製要有表單領用紀錄                
            message2 = "新增成功"
        Else
            message2 = formnocheck.form_usestate(UCase(物件編號), store.SelectedValue)
        End If

        '1050506表單控管
        If ddl契約類別.SelectedValue <> "庫存" Then
            'If Request.Cookies("webfly_empno").Value = "92H" Then
            If check_formno(UCase(Mid(物件編號, 6)), store.SelectedValue, 物件編號, input4.Value, "新增") = "use" Then
                eip_usual.Show("一併寫入表單管理")
            Else
                'eip_usual.Show("輸入編號不為所購買表單，無法寫入表單管理")
            End If
            'End If
        End If

        'With Image1
        '    .ImageUrl = "https://el.etwarm.com.tw/new_eip/tool/code41.ashx?id=" & 物件編號
        '    .Visible = True
        'End With


        script += "alert('" & message2 & "');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

        'Dim href As String = "https://el.etwarm.com.tw/new_eip/tool/t_條碼分類.aspx?state=update&oid=" & 物件編號 & "&sid=" & store.SelectedValue & "&folder=available&rsid=" & Request("sid") & "&src=" & Request("src") & "&check=description" 'check參數用意在不顯示"下一步按鈕"
        'Dim NSCRIPT As String = "GB_showCenter('產生條碼', '" & href & "',320,580);"
        'Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", NSCRIPT, True)


        If CheckBox26.Checked = True Then
            複製其他所有資料()
        End If


        conn.Close()
        conn = Nothing

        'Catch
        '    Response.Write(sql)
        'End Try

    End Sub











    Sub 複製其他所有資料()

        Dim 承辦人1 As String, 承辦人2 As String, 承辦人3 As String

        承辦人1 = sale1.SelectedValue
        承辦人2 = sale2.SelectedValue
        承辦人3 = sale3.SelectedValue

        Dim script As String
        Dim 物件編號 As String = ""
        Dim 被複製編號 As String = ""
        Dim 店代號 As String = store.SelectedValue


        If Trim(Request("src")) = "OLD" Then
            src.Text = "委賣物件過期資料表"
        ElseIf Trim(Request("src")) = "NOW" Then
            src.Text = "委賣物件資料表"
        End If

        'old
        被複製編號 = Request("oid")

        'new-註計為BUG
        'If ddl契約類別.SelectedValue = "專任" Then
        '    物件編號 = "1"
        'ElseIf ddl契約類別.SelectedValue = "一般" Then
        '    物件編號 = "6"
        'ElseIf ddl契約類別.SelectedValue = "同意書" Then
        '    物件編號 = "7"
        'ElseIf ddl契約類別.SelectedValue = "流通" Then
        '    物件編號 = "5"
        'ElseIf ddl契約類別.SelectedValue = "庫存" Then
        '    物件編號 = "9"
        'End If
        '物件編號 &= Mid(Request("oid"), 2, 4) & TextBox2.Text

        物件編號 = Label57.Text




        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        If Request.Cookies("webfly_empno").Value.ToUpper <> "92H" Then
            Dim insertStr As String = "insert into 物件照片複製暫存表 (來源物件編號,目標物件編號,來源店代號,目標店代號,來源資料表) values ('" & 被複製編號 & "','" & 物件編號 & "','" & Request("sid") & "','" & 店代號 & "','委賣物件資料表')"
            Using connTmp As New SqlConnection(EGOUPLOADSqlConnStr)
                connTmp.Open()
                Using cmd As New SqlCommand(insertStr, connTmp)
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        Response.Write(ex.Message)
                        Response.End()
                    End Try
                End Using
            End Using
        End If


        '參數順序︰原物件編號,複製到的物件編號,店代號,承辦人,承辦人2,承辦人3,承辦秘書
        Mod_複製委賣相關資料("委賣屋主資料表", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)
        Mod_複製委賣相關資料("委賣_房地產說明書", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)

        Mod_複製委賣物件相片資料表("委賣物件相片", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)
        Mod_複製委賣物件地圖資料表("委賣地圖", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)
        Mod_複製委賣物件格局圖資料表("委賣隔局圖", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)

        Mod_複製checkpic資料表("checkpic", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)

        Mod_複製委賣物件其他資料(src.Text, 被複製編號, 物件編號, 店代號)

        '20150520新增-複製高畫質影音
        Mod_複製Youtube資料表("Youtube", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)

        '20110621複製細項面積--------------------------------------------------------------------------------------------------------
        Mod_複製委賣物件細項面積("委賣物件資料表_面積細項", 被複製編號, 物件編號, 店代號)
        Mod_複製委賣物件細項面積("物件他項權利細項", 被複製編號, 物件編號, 店代號)
        Mod_複製委賣物件細項面積("工具_計算土地增值稅", 被複製編號, 物件編號, 店代號)
        '----------------------------------------------------------------------------------------------------------------------------


        Mod_複製相片影音地圖格局圖檔("", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)
        '20080624新增----------------------------------------------------------------------------------------------------------------------------
        'MySQL複製黃金隔局圖資料表(210.200.219.131 etwarm paint_gold)
        ''20190910 10.40.20.66先行拿掉==================================
        ''Mod_複製paint_gold資料表("paint_gold", 被複製編號, 物件編號, 店代號, 承辦人1, 承辦人2, 承辦人3, Label33.Text)
        ''20190910 10.40.20.66先行拿掉==================================
        '----------------------------------------------------------------------------------------------------------------------------------------


        'Call Mod_複製委賣相關資料("委賣追蹤資料表", Request("oid"), 物件編號, ddl店代號.SelectedValue, 承辦人1(0), 承辦人2(0), 承辦人3(0), Text1.Value)
        'Call Mod_複製委賣相關資料("買方帶看資料表", Request("oid"), 物件編號, DropDownList16.SelectedValue, 承辦人1(0), 承辦人2(0), 承辦人3(0), Text1.Value)
        '影音,格局,



        'Dim href2 As String = "https://superweb2.etwarm.com.tw/720p/copy720.php?fromsid=" & Request("sid") & "&tosid=" & 店代號 & "&fromoid=" & 被複製編號 & "&tooid=" & 物件編號 & "&src=" & Request("src")

        'script = "<script language='javascript'>"

        'script += "window.open('"
        'script += href2
        'script += " '"
        'script += ",'newwindows', 'height=320, width=580, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no,status=no');"
        'script += "<"
        'script += "/script>"
        'Response.Write(script)
        'Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", script, True)
        'If Request("type") = "again" Then
        '    script = "<script language='javascript'>"
        '    script += "window.opener = null;"
        '    script += "window.close();"
        '    script += "<"
        '    script += "/script>"
        '    Response.Write(script)
        'End If

    End Sub

    Sub Mod_複製委賣物件其他資料(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "' "

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)

            conn.Open()

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            table1 = ds.Tables("table1")
            If table1.Rows.Count Then


                sql = "Update 委賣物件資料表" '& TableName
                sql &= " Set "
                sql &= " num = " & IIf(table1.Rows(0)("num").ToString = "", "0", table1.Rows(0)("num").ToString) & ","
                sql &= " movie = " & IIf(table1.Rows(0)("movie").ToString = "", "null", table1.Rows(0)("movie").ToString) & ","
                sql &= " movie_b = '" & table1.Rows(0)("movie_b").ToString & "',"
                sql &= " movie_s = '" & table1.Rows(0)("movie_s").ToString & "',"
                sql &= " movie_h = '" & table1.Rows(0)("movie_h").ToString & "',"
                sql &= " movie_720 = '" & table1.Rows(0)("movie_720").ToString & "',"
                sql &= " map = '" & table1.Rows(0)("map").ToString & "',"
                sql &= " plan1 = '" & table1.Rows(0)("plan1").ToString & "',"
                sql &= " 經度 = " & IIf(table1.Rows(0)("經度").ToString = "", "null", table1.Rows(0)("經度").ToString) & ","
                sql &= " 緯度 = " & IIf(table1.Rows(0)("緯度").ToString = "", "null", table1.Rows(0)("緯度").ToString) & ","
                sql &= " 車位 = '" & table1.Rows(0)("車位").ToString & "' "
                sql &= " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號='" & 店代號1 & "' "

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

            End If

        End Using

    End Sub

    '20090624新增-------------------------------------------------------------------------------------------------------------
    Sub Mod_複製paint_gold資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)
        ''20190910 10.40.20.66先行拿掉==================================
        ''Dim MYSQL_conn As New OdbcConnection(mysqletwarmstring)

        ' ''被複制的物件編號

        ''sql = "Select * From " & TableName & " With(NoLock) Where object_id = '" & 原物件編號 & "' and storeid = '" & Request("sid") & "'"


        ''MySQL_adpt = New OdbcDataAdapter(sql, MYSQL_conn)
        ''MYSQL_conn.Open()
        ''ds = New DataSet()
        ''MySQL_adpt.Fill(ds, "table1")
        ''table1 = ds.Tables("table1")
        ''If table1.Rows.Count Then

        ''    sql = "Select * From " & TableName & " With(NoLock) Where object_id = '" & 複製的物件編號 & "' and storeid = '" & 店代號1 & "' "
        ''    MySQL_adpt = New OdbcDataAdapter(sql, MYSQL_conn)
        ''    ds = New DataSet()
        ''    MySQL_adpt.Fill(ds, "table2")
        ''    table2 = ds.Tables("table2")
        ''    If table2.Rows.Count Then
        ''        '若已存在就先刪除這一筆
        ''        sql = "Delete From " & TableName & "  Where object_id = '" & 複製的物件編號 & "' and storeid = '" & 店代號1 & "' "
        ''        MySQL_cmd = New OdbcCommand(sql, MYSQL_conn)
        ''        MySQL_cmd.ExecuteNonQuery()
        ''    End If

        ''    '新增
        ''    sql = "insert into " & TableName
        ''    sql &= " (object_id,storeid,c0,c1,c2,c3,c4,c5)"
        ''    sql &= " Values ('" & 複製的物件編號 & "','" & store.SelectedValue & "','" & table1.Rows(0)("c0") & "','" & table1.Rows(0)("c1") & "','" & table1.Rows(0)("c2") & "','" & table1.Rows(0)("c3") & "','" & table1.Rows(0)("c4") & "','" & table1.Rows(0)("c5") & "') "
        ''    'sql &= " Values (@pa1,@pa2,@pa3,@pa4,@pa5,@pa6,@pa7,@pa8) "
        ''    Dim Command As OdbcCommand
        ''    Command = New OdbcCommand(sql, MYSQL_conn)

        ''    Command.ExecuteNonQuery()

        ''End If
        ''MYSQL_conn.Close()
        ''MYSQL_conn.Dispose()
        ''20190910 10.40.20.66先行拿掉==================================
    End Sub
    '-------------------------------------------------------------------------------------------------------------------------

    Sub Mod_複製委賣物件相片資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        If table1.Rows.Count Then

            sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count Then
                '若已存在就先刪除這一筆
                sql = "Delete From " & TableName & "  Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If

            '新增
            sql = "insert into " & TableName
            sql &= " (物件編號,店代號,經紀人代號,營業員代號1,營業員代號2,輸入者,"
            sql &= " 新增日期,修改日期,照片1,照片2,照片3,照片4,照片說明1,照片說明2,照片說明3,照片說明4,上傳註記,上傳日期,照片5,照片6,照片7,照片8,照片說明5,照片說明6,照片說明7,照片說明8, "
            sql &= " 照片9,照片10,照片11,照片12,照片13,照片14,照片15, "
            sql &= " 照片說明9,照片說明10,照片說明11,照片說明12,照片說明13,照片說明14,照片說明15) "
            sql &= " Values (@pa1,@pa2,@pa3,@pa4,@pa5,@pa6,"
            sql &= " @pa7,@pa8,@pa9,@pa10,@pa11,@pa12,@pa13,@pa14,@pa15,@pa16,@pa17,@pa18,@pa19,@pa20,@pa21,@pa22,@pa23,@pa24,@pa25,@pa26, "
            sql &= " @pa27,@pa28,@pa29,@pa30,@pa31,@pa32,@pa33, "
            sql &= " @pa34,@pa35,@pa36,@pa37,@pa38,@pa39,@pa40) "


            Dim Command As SqlCommand
            Command = New SqlCommand(sql, conn)

            Dim splitarray As Array, str As String = ""

            '物件編號
            Command.Parameters.Add(New SqlParameter("@pa1", SqlDbType.NVarChar))
            Command.Parameters("@pa1").Value = 複製的物件編號
            '店代號
            Command.Parameters.Add(New SqlParameter("@pa2", SqlDbType.NVarChar))
            Command.Parameters("@pa2").Value = store.SelectedValue

            '經紀人代號, 
            Command.Parameters.Add(New SqlParameter("@pa3", SqlDbType.NVarChar))

            If sale1.SelectedValue <> "選擇人員" Then
                str = sale1.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa3").Value = str

            '營業員代號1
            Command.Parameters.Add(New SqlParameter("@pa4", SqlDbType.NVarChar))

            If sale2.SelectedValue <> "選擇人員" Then
                str = sale2.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa4").Value = str
            '營業員代號2
            Command.Parameters.Add(New SqlParameter("@pa5", SqlDbType.NVarChar))

            If sale3.SelectedValue <> "選擇人員" Then
                str = sale3.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa5").Value = str

            '輸入者
            Command.Parameters.Add(New SqlParameter("@pa6", SqlDbType.NVarChar))
            Command.Parameters("@pa6").Value = Label33.Text
            '新增日期
            Command.Parameters.Add(New SqlParameter("@pa7", SqlDbType.NVarChar))
            Command.Parameters("@pa7").Value = sysdate
            '修改日期
            Command.Parameters.Add(New SqlParameter("@pa8", SqlDbType.NVarChar))
            Command.Parameters("@pa8").Value = sysdate
            '照片1
            Command.Parameters.Add(New SqlParameter("@pa9", SqlDbType.Char))
            Command.Parameters("@pa9").Value = table1.Rows(0)("照片1")
            '照片2
            Command.Parameters.Add(New SqlParameter("@pa10", SqlDbType.Char))
            Command.Parameters("@pa10").Value = table1.Rows(0)("照片2")
            '照片3
            Command.Parameters.Add(New SqlParameter("@pa11", SqlDbType.Char))
            Command.Parameters("@pa11").Value = table1.Rows(0)("照片3")
            '照片4
            Command.Parameters.Add(New SqlParameter("@pa12", SqlDbType.Char))
            Command.Parameters("@pa12").Value = table1.Rows(0)("照片4")
            '照片說明1
            Command.Parameters.Add(New SqlParameter("@pa13", SqlDbType.NVarChar))
            Command.Parameters("@pa13").Value = table1.Rows(0)("照片說明1")
            '照片說明2
            Command.Parameters.Add(New SqlParameter("@pa14", SqlDbType.NVarChar))
            Command.Parameters("@pa14").Value = table1.Rows(0)("照片說明2")
            '照片說明3
            Command.Parameters.Add(New SqlParameter("@pa15", SqlDbType.NVarChar))
            Command.Parameters("@pa15").Value = table1.Rows(0)("照片說明3")
            '照片說明4
            Command.Parameters.Add(New SqlParameter("@pa16", SqlDbType.NVarChar))
            Command.Parameters("@pa16").Value = table1.Rows(0)("照片說明4")
            '上傳註記
            Command.Parameters.Add(New SqlParameter("@pa17", SqlDbType.NVarChar))
            Command.Parameters("@pa17").Value = table1.Rows(0)("上傳註記")
            '上傳日期
            Command.Parameters.Add(New SqlParameter("@pa18", SqlDbType.NVarChar))
            Command.Parameters("@pa18").Value = sysdate
            '20101028小豪新增
            '照片5
            Command.Parameters.Add(New SqlParameter("@pa19", SqlDbType.Char))
            Command.Parameters("@pa19").Value = table1.Rows(0)("照片5")
            '照片6
            Command.Parameters.Add(New SqlParameter("@pa20", SqlDbType.Char))
            Command.Parameters("@pa20").Value = table1.Rows(0)("照片6")
            '照片7
            Command.Parameters.Add(New SqlParameter("@pa21", SqlDbType.Char))
            Command.Parameters("@pa21").Value = table1.Rows(0)("照片7")
            '照片8
            Command.Parameters.Add(New SqlParameter("@pa22", SqlDbType.Char))
            Command.Parameters("@pa22").Value = table1.Rows(0)("照片8")
            '照片說明5
            Command.Parameters.Add(New SqlParameter("@pa23", SqlDbType.NVarChar))
            Command.Parameters("@pa23").Value = table1.Rows(0)("照片說明5")
            '照片說明6
            Command.Parameters.Add(New SqlParameter("@pa24", SqlDbType.NVarChar))
            Command.Parameters("@pa24").Value = table1.Rows(0)("照片說明6")
            '照片說明7
            Command.Parameters.Add(New SqlParameter("@pa25", SqlDbType.NVarChar))
            Command.Parameters("@pa25").Value = table1.Rows(0)("照片說明7")
            '照片說明8
            Command.Parameters.Add(New SqlParameter("@pa26", SqlDbType.NVarChar))
            Command.Parameters("@pa26").Value = table1.Rows(0)("照片說明8")

            '照片9
            Command.Parameters.Add(New SqlParameter("@pa27", SqlDbType.Char))
            Command.Parameters("@pa27").Value = table1.Rows(0)("照片9")
            '照片10
            Command.Parameters.Add(New SqlParameter("@pa28", SqlDbType.Char))
            Command.Parameters("@pa28").Value = table1.Rows(0)("照片10")
            '照片11
            Command.Parameters.Add(New SqlParameter("@pa29", SqlDbType.Char))
            Command.Parameters("@pa29").Value = table1.Rows(0)("照片11")
            '照片12
            Command.Parameters.Add(New SqlParameter("@pa30", SqlDbType.Char))
            Command.Parameters("@pa30").Value = table1.Rows(0)("照片12")
            '照片13
            Command.Parameters.Add(New SqlParameter("@pa31", SqlDbType.Char))
            Command.Parameters("@pa31").Value = table1.Rows(0)("照片13")
            '照片14
            Command.Parameters.Add(New SqlParameter("@pa32", SqlDbType.Char))
            Command.Parameters("@pa32").Value = table1.Rows(0)("照片14")
            '照片15
            Command.Parameters.Add(New SqlParameter("@pa33", SqlDbType.Char))
            Command.Parameters("@pa33").Value = table1.Rows(0)("照片15")

            '照片說明9
            Command.Parameters.Add(New SqlParameter("@pa34", SqlDbType.Char))
            Command.Parameters("@pa34").Value = table1.Rows(0)("照片說明9")
            '照片說明10
            Command.Parameters.Add(New SqlParameter("@pa35", SqlDbType.Char))
            Command.Parameters("@pa35").Value = table1.Rows(0)("照片說明10")
            '照片說明11
            Command.Parameters.Add(New SqlParameter("@pa36", SqlDbType.Char))
            Command.Parameters("@pa36").Value = table1.Rows(0)("照片說明11")
            '照片說明12
            Command.Parameters.Add(New SqlParameter("@pa37", SqlDbType.Char))
            Command.Parameters("@pa37").Value = table1.Rows(0)("照片說明12")
            '照片說明13
            Command.Parameters.Add(New SqlParameter("@pa38", SqlDbType.Char))
            Command.Parameters("@pa38").Value = table1.Rows(0)("照片說明13")
            '照片說明14
            Command.Parameters.Add(New SqlParameter("@pa39", SqlDbType.Char))
            Command.Parameters("@pa39").Value = table1.Rows(0)("照片說明14")
            '照片說明15
            Command.Parameters.Add(New SqlParameter("@pa40", SqlDbType.Char))
            Command.Parameters("@pa40").Value = table1.Rows(0)("照片說明15")

            Command.ExecuteNonQuery()
        End If

    End Sub


    Sub Mod_複製委賣物件格局圖資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        '被複制的物件編號        

        sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        If table1.Rows.Count Then

            sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count Then
                '若已存在就先刪除這一筆
                sql = "Delete From " & TableName & " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If


            '新增
            sql = "Insert Into 委賣隔局圖 (物件編號,店代號,上傳日期,上傳註記,輸入者)"
            sql += " Values('" & 複製的物件編號 & "','" & 店代號1 & "','" & sysdate & "','A','" & Label33.Text & "')"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
            cmd.Dispose()

        End If

    End Sub

    Sub Mod_複製委賣物件地圖資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"



        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        If table1.Rows.Count Then

            sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count Then
                '若已存在就先刪除這一筆
                sql = "Delete From " & TableName & " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If

            '新增
            sql = "insert into " & TableName
            sql &= " (物件編號,環境介紹,上傳註記,上傳日期,店代號,經紀人代號,營業員代號1,"
            sql &= " 營業員代號2,輸入者,新增日期,修改日期) "
            sql &= " Values (@pa1,@pa2,@pa3,@pa4,@pa5,@pa6,@pa7"
            sql &= " ,@pa8,@pa9,@pa10,@pa11) "
            Dim Command As SqlCommand
            Command = New SqlCommand(sql, conn)

            Dim splitarray As Array, str As String = ""

            '物件編號
            Command.Parameters.Add(New SqlParameter("@pa1", SqlDbType.NVarChar))
            Command.Parameters("@pa1").Value = 複製的物件編號
            '環境介紹
            Command.Parameters.Add(New SqlParameter("@pa2", SqlDbType.NVarChar))
            Command.Parameters("@pa2").Value = table1.Rows(0)("環境介紹")
            '上傳註記
            Command.Parameters.Add(New SqlParameter("@pa3", SqlDbType.NVarChar))
            Command.Parameters("@pa3").Value = table1.Rows(0)("上傳註記")
            '上傳日期
            Command.Parameters.Add(New SqlParameter("@pa4", SqlDbType.NVarChar))
            Command.Parameters("@pa4").Value = table1.Rows(0)("上傳日期")
            '店代號
            Command.Parameters.Add(New SqlParameter("@pa5", SqlDbType.NVarChar))
            Command.Parameters("@pa5").Value = store.SelectedValue

            '經紀人代號, 
            Command.Parameters.Add(New SqlParameter("@pa6", SqlDbType.NVarChar))

            If sale1.SelectedValue <> "選擇人員" Then

                str = sale1.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa6").Value = str

            '營業員代號1
            Command.Parameters.Add(New SqlParameter("@pa7", SqlDbType.NVarChar))

            If sale2.SelectedValue <> "選擇人員" Then

                str = sale2.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa7").Value = str
            '營業員代號2
            Command.Parameters.Add(New SqlParameter("@pa8", SqlDbType.NVarChar))

            If sale3.SelectedValue <> "選擇人員" Then
                str = sale3.SelectedValue
            Else
                str = ""
            End If

            Command.Parameters("@pa8").Value = str

            '輸入者
            Command.Parameters.Add(New SqlParameter("@pa9", SqlDbType.NVarChar))
            Command.Parameters("@pa9").Value = Label33.Text
            '新增日期
            Command.Parameters.Add(New SqlParameter("@pa10", SqlDbType.NVarChar))
            Command.Parameters("@pa10").Value = sysdate
            '修改日期
            Command.Parameters.Add(New SqlParameter("@pa11", SqlDbType.NVarChar))
            Command.Parameters("@pa11").Value = sysdate
            Command.ExecuteNonQuery()

        End If

    End Sub

    Sub Mod_複製Youtube資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where y_oid = '" & 原物件編號 & "' and y_sid = '" & Request("sid") & "' "

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count Then

            sql = "Select * From " & TableName & " With(NoLock) Where y_oid = '" & 複製的物件編號 & "' and y_sid = '" & 店代號1 & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count Then
                '若已存在就先刪除這一筆
                sql = "Delete From " & TableName & " Where y_oid = '" & 複製的物件編號 & "' and y_sid = '" & 店代號1 & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If

            sql = "Insert Into " & TableName
            sql &= "(y_oid,y_id,y_sid,y_del,y_datetime)" &
                   " Values( " &
                   "'" & 複製的物件編號 & "','" & table1.Rows(0)("y_id") & "','" & 店代號1 & "'," &
                   "'" & table1.Rows(0)("y_del") & "','" & table1.Rows(0)("y_datetime") & "')"


            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

        End If

    End Sub

    Sub Mod_複製checkpic資料表(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where pic_contract_no = '" & 原物件編號 & "' and pic_dept_no = '" & Request("sid") & "' "

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count Then

            sql = "Select * From " & TableName & " With(NoLock) Where pic_contract_no = '" & 複製的物件編號 & "' and pic_dept_no = '" & 店代號1 & "' "
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table2")
            table2 = ds.Tables("table2")
            If table2.Rows.Count Then
                '若已存在就先刪除這一筆
                sql = "Delete From " & TableName & " Where pic_contract_no = '" & 複製的物件編號 & "' and pic_dept_no = '" & 店代號1 & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If

            sql = "Insert Into checkpic "
            sql &= "(pic_contract_no,pic_1,pic_2,pic_3,pic_4,pic_5,pic_6,pic_7,pic_8,pic_count" &
                   ",pic_map,pic_map_count,pic_360_1,pic_360_2,pic_360_3,pic_360_4,pic_360_count" &
                   ",pic_vcr,pic_vcr_b,pic_plan,pic_dept_no,pic_upd_dt, " &
                   " pic_9,pic_10,pic_11,pic_12,pic_13,pic_14,pic_15, " &
                   " pic_vr,pic_youtube " &
                   " )" &
                   " Values( " &
                   "'" & 複製的物件編號 & "','" & table1.Rows(0)("pic_1") & "','" & table1.Rows(0)("pic_2") & "'," &
                   "'" & table1.Rows(0)("pic_3") & "','" & table1.Rows(0)("pic_4") & "','" & table1.Rows(0)("pic_5") & "','" & table1.Rows(0)("pic_6") & "','" & table1.Rows(0)("pic_7") & "','" & table1.Rows(0)("pic_8") & "','" & table1.Rows(0)("pic_count") & "'," &
                   "'" & table1.Rows(0)("pic_map") & "','" & table1.Rows(0)("pic_map_count") & "','" & table1.Rows(0)("pic_360_1") & "'," &
                   "'" & table1.Rows(0)("pic_360_2") & "','" & table1.Rows(0)("pic_360_3") & "','" & table1.Rows(0)("pic_360_4") & "'," &
                   "'" & table1.Rows(0)("pic_360_count") & "','" & table1.Rows(0)("pic_vcr") & "','" & table1.Rows(0)("pic_vcr_b") & "'," &
                   "'" & table1.Rows(0)("pic_plan") & "','" & 店代號1 & "','" & sysdate & "'," &
                   "'" & table1.Rows(0)("pic_9") & "','" & table1.Rows(0)("pic_10") & "','" & table1.Rows(0)("pic_11") & "','" & table1.Rows(0)("pic_12") & "','" & table1.Rows(0)("pic_13") & "','" & table1.Rows(0)("pic_14") & "','" & table1.Rows(0)("pic_15") & "'," &
                   "'" & table1.Rows(0)("pic_vr") & "','" & table1.Rows(0)("pic_youtube") & "'" &
                   ")"

            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()

        End If

    End Sub

    Sub Mod_複製相片影音地圖格局圖檔(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)
        'Dim insertStr As String = "insert into 物件照片複製暫存表 (來源物件編號,目標物件編號,來源店代號,目標店代號) values ('" & 原物件編號 & "','" & 複製的物件編號 & "','" & Request("sid") & "','" & 店代號1 & "')"
        'Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
        '    conn.Open()
        '    Using cmd As New SqlCommand(insertStr, conn)
        '        cmd.ExecuteNonQuery()
        '    End Using
        'End Using

        Dim picurl As String
        '影音------------------------------------------------------------------------------------------------------------
        '判別網路上有沒有影音-WMV       
        If Request("type") = "again" Then
            picurl = "https://media.etwarm.com.tw/obj_movie/" & Request("csid") & "/" & 原物件編號 & "-b.wmv"
        Else
            picurl = "https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 原物件編號 & "-b.wmv"
        End If
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(picurl)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 複製的物件編號 & "-b.wmv") Then
                        File.Delete("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 複製的物件編號 & "-b.wmv")
                    End If


                    If Request("type") = "again" Then
                        My.Computer.Network.DownloadFile("https://media.etwarm.com.tw/obj_movie/" & Request("csid") & "/" & 原物件編號 & "-b.wmv", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.wmv")
                    Else
                        My.Computer.Network.DownloadFile("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 原物件編號 & "-b.wmv", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.wmv")
                    End If
                    FTP2.Passive = True
                    FTP2.Connect("media.etwarm.com.tw", 21)
                    FTP2.Login("movie", "movie")
                    '20110707
                    FTP2.KeepAliveDuringTransferInterval = 3000

                    If FTP2.DirectoryExists(store.SelectedValue) = False Then
                        FTP2.CreateDirectory(store.SelectedValue)
                    Else
                        If FTP2.FileExists(store.SelectedValue & "\" & 複製的物件編號 & "-b.wmv") Then FTP2.DeleteFile(store.SelectedValue & "\" & 複製的物件編號 & "-b.wmv")
                    End If

                    System.Threading.Thread.Sleep(1000)

                    Dim fa2 As System.IO.Stream

                    '上傳影片
                    If File.Exists("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.wmv") Then

                        fa2 = File.Open("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.wmv", FileMode.Open)
                        FTP2.PutFile(fa2, store.SelectedValue & "\" & 複製的物件編號 & "-b.wmv")
                        '清除檔案佔用
                        fa2.Close()
                        fa2.Dispose()

                        '20110412刪除UPFILE裡的物件
                        File.Delete("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.wmv")
                    End If

                End If
            End Using
        Catch ex As WebException
            Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            If webResponse.StatusCode = HttpStatusCode.NotFound Then


            End If
        End Try



        '判別網路上有沒有影音-mp4       
        If Request("type") = "again" Then
            picurl = "https://media.etwarm.com.tw/obj_movie/" & Request("csid") & "/" & 原物件編號 & "-b.mp4"
        Else
            picurl = "https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 原物件編號 & "-b.mp4"
        End If
        Try
            Dim requests As WebRequest = HttpWebRequest.Create(picurl)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 複製的物件編號 & "-b.mp4") Then
                        File.Delete("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 複製的物件編號 & "-b.mp4")
                    End If


                    If Request("type") = "again" Then
                        My.Computer.Network.DownloadFile("https://media.etwarm.com.tw/obj_movie/" & Request("csid") & "/" & 原物件編號 & "-b.mp4", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.mp4")
                    Else
                        My.Computer.Network.DownloadFile("https://media.etwarm.com.tw/obj_movie/" & Request("sid") & "/" & 原物件編號 & "-b.mp4", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.mp4")
                    End If
                    FTP2.Passive = True
                    FTP2.Connect("media.etwarm.com.tw", 21)
                    FTP2.Login("movie", "movie")
                    '20110707
                    FTP2.KeepAliveDuringTransferInterval = 3000

                    If FTP2.DirectoryExists(store.SelectedValue) = False Then
                        FTP2.CreateDirectory(store.SelectedValue)
                    Else
                        If FTP2.FileExists(store.SelectedValue & "\" & 複製的物件編號 & "-b.mp4") Then FTP2.DeleteFile(store.SelectedValue & "\" & 複製的物件編號 & "-b.mp4")
                    End If

                    System.Threading.Thread.Sleep(1000)

                    Dim fa2 As System.IO.Stream

                    '上傳影片
                    If File.Exists("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.mp4") Then

                        fa2 = File.Open("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.mp4", FileMode.Open)
                        FTP2.PutFile(fa2, store.SelectedValue & "\" & 複製的物件編號 & "-b.mp4")
                        '清除檔案佔用
                        fa2.Close()
                        fa2.Dispose()

                        '20110412刪除UPFILE裡的物件
                        File.Delete("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & "-b.mp4")
                    End If

                End If
            End Using
        Catch ex As WebException
            Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            If webResponse.StatusCode = HttpStatusCode.NotFound Then


            End If
        End Try
        '影音------------------------------------------------------------------------------------------------------------

        '照片、格局圖、地圖=================================================================
        'If Request.Cookies("webfly_empno").Value.ToUpper = "92H" Then
        FTP2.Passive = True
        FTP2.Connect("10.40.20.88", 21)
        FTP2.Login("etwarm_media", "Etwa!r@m")
        '20110707
        FTP2.KeepAliveDuringTransferInterval = 3000

        Dim 圖片 As String() = Split("a,b,c,d,w,x,y,z,e,f,h,s,t,u,v,g,m", ",")
        For i = 0 To 16
            '判別網路上有沒有影音-mp4       
            If Request("type") = "again" Then
                picurl = "https://img.etwarm.com.tw/" & Request("csid") & "/available/" & 原物件編號 & 圖片(i) & ".jpg"
            Else
                picurl = "https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 原物件編號 & 圖片(i) & ".jpg"
            End If
            Try
                Dim requests As WebRequest = HttpWebRequest.Create(picurl)
                requests.Method = "HEAD"
                requests.Credentials = System.Net.CredentialCache.DefaultCredentials
                Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                    If response.StatusCode = HttpStatusCode.OK Then
                        'UPFILE資料夾內檔案若已存在,先刪除該檔案
                        If File.Exists("https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 複製的物件編號 & 圖片(i) & ".jpg") Then
                            File.Delete("https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 複製的物件編號 & 圖片(i) & ".jpg")
                        End If

                        If Request("type") = "again" Then
                            My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("csid") & "/available/" & 原物件編號 & 圖片(i) & ".jpg", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & 圖片(i) & ".jpg")
                        Else
                            My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 原物件編號 & 圖片(i) & ".jpg", "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & 圖片(i) & ".jpg")
                        End If


                        If FTP2.DirectoryExists(store.SelectedValue) = False Then
                            FTP2.CreateDirectory(store.SelectedValue)
                        Else
                            If FTP2.FileExists(store.SelectedValue & "\available\" & 複製的物件編號 & 圖片(i) & ".jpg") Then FTP2.DeleteFile(store.SelectedValue & "\" & 複製的物件編號 & 圖片(i) & ".jpg")
                        End If

                        System.Threading.Thread.Sleep(1000)

                        Dim fa2 As System.IO.Stream

                        '上傳影片
                        If File.Exists("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & 圖片(i) & ".jpg") Then

                            fa2 = File.Open("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & 圖片(i) & ".jpg", FileMode.Open)
                            FTP2.PutFile(fa2, store.SelectedValue & "\available\" & 複製的物件編號 & 圖片(i) & ".jpg")
                            '清除檔案佔用
                            fa2.Close()
                            fa2.Dispose()

                            '20110412刪除UPFILE裡的物件
                            File.Delete("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號 & 圖片(i) & ".jpg")
                        End If

                    End If
                End Using
            Catch ex As WebException
                Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
                If webResponse.StatusCode = HttpStatusCode.NotFound Then

                End If
            End Try
        Next

        FTP2.Disconnect()
        'End If
        '照片、格局圖、地圖=================================================================

    End Sub

    '1040413照片未移動到過期資料夾,額外判斷使用
    Sub chk_again(href As String, 複製的物件編號 As String, 原物件編號 As String)
        Dim FTP目錄 As String = store.SelectedValue & "\" & Me.Label23.Text

        Try
            Dim requests As WebRequest = HttpWebRequest.Create(href)
            requests.Method = "HEAD"
            requests.Credentials = System.Net.CredentialCache.DefaultCredentials
            Using response As HttpWebResponse = DirectCast(requests.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.OK Then

                    'UPFILE資料夾內檔案若已存在,先刪除該檔案
                    If File.Exists("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號) Then
                        File.Delete("C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號)
                    End If
                    '照片路徑必在available，故寫死available
                    If Request("type") = "again" Then
                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("csid") & "/available/" & 原物件編號, "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號)
                    Else
                        My.Computer.Network.DownloadFile("https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 原物件編號, "C:\Inetpub\wwwroot\Filesystem\2014\" & 複製的物件編號)
                    End If

                End If
            End Using

            ' Response.Write(href & "-" & "https://img.etwarm.com.tw/" & Request("sid") & "/available/" & 原物件編號)
        Catch ex As WebException
            Dim webResponse As HttpWebResponse = DirectCast(ex.Response, HttpWebResponse)
            If webResponse.StatusCode = HttpStatusCode.NotFound Then


            End If
        End Try
    End Sub
    Sub Mod_複製委賣物件細項面積(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String)

        '被複制的物件編號
        Dim sql_outside As String = "Select * From " & TableName & " With(NoLock) Where 店代號 = '" & Request("sid") & "' and 物件編號 = '" & 原物件編號 & "' "


        adpt = New SqlDataAdapter(sql_outside, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count > 0 Then
            '刪除已存在的
            Using conn_delexist As New SqlConnection(EGOUPLOADSqlConnStr)
                conn_delexist.Open()
                Dim sql_delexist As String = "Delete From " & TableName & " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
                Using cmd_deleExist As New SqlCommand(sql_delexist, conn_delexist)
                    cmd_deleExist.ExecuteNonQuery()
                End Using
            End Using
            If TableName = "工具_計算土地增值稅" Then
                Dim insertStr As String
                insertStr = "INSERT INTO 工具_計算土地增值稅 (物件編號, 地址, 土地標示, 一般增值稅, 自用增值稅, 本次公告現值, 前次公告現值, 前次取得年月, 物價指數, 土地面積, 持分1, 持分2, 備註1, 備註2, 備註3, 新增日期, "
                insertStr += " 修改日期, 前次取得日期, 稅率說明, 店代號, 所有人姓名, 備註, 土地使用類別, 年限,Num )"
                insertStr += " Select  '" & 複製的物件編號 & "' as 物件編號, 地址, 土地標示, 一般增值稅, 自用增值稅, 本次公告現值, 前次公告現值, 前次取得年月, 物價指數, 土地面積, 持分1, 持分2, 備註1, 備註2, 備註3, 新增日期, "
                insertStr += " 修改日期, 前次取得日期, 稅率說明, '" & 店代號1 & "' as 店代號, 所有人姓名, 備註, 土地使用類別, 年限,Num "
                insertStr += " FROM 工具_計算土地增值稅"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using
            End If
            If TableName = "物件他項權利細項" Then
                Dim insertStr As String
                insertStr = "INSERT INTO 物件他項權利細項 (物件編號, 店代號, 權利類別, 權利種類, 順位, 登記日期, 設定, 設定權利人, 管理人, Num )"
                insertStr += " Select  '" & 複製的物件編號 & "' as 物件編號, '" & 店代號1 & "' as 店代號, 權利類別, 權利種類, 順位, 登記日期, 設定, 設定權利人, 管理人, Num "
                insertStr += " FROM 物件他項權利細項"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using
            End If
            If TableName = "委賣物件資料表_面積細項" Then
                'select into 面積細項
                Dim insertStr As String
                insertStr = "INSERT INTO 委賣物件資料表_面積細項 (物件編號, 流水號, 建號, 類別, 項目名稱, 總面積平方公尺, 總面積坪, 權利範圍1分母, 權利範圍1分子, 權利範圍2分母, 權利範圍2分子, 實際持有平方公尺, 實際持有坪, "
                insertStr += "店代號, 管制, 使用分區, 所有權人, 增建用途, 增建完成日期, 法定建蔽率, 法定容積率, DL_level2_selectindex, 是否為公設)"
                insertStr += " Select  '" & 複製的物件編號 & "' as 物件編號, 流水號, 建號, 類別, 項目名稱, 總面積平方公尺, 總面積坪, 權利範圍1分母, 權利範圍1分子, 權利範圍2分母, 權利範圍2分子, 實際持有平方公尺, 實際持有坪, "
                insertStr += " '" & 店代號1 & "' as 店代號, 管制, 使用分區, 所有權人, 增建用途, 增建完成日期, 法定建蔽率, 法定容積率, DL_level2_selectindex, 是否為公設"
                insertStr += " FROM 委賣物件資料表_面積細項"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using

                'select into 細項所有權人

                insertStr = "INSERT INTO 委賣物件資料表_細項所有權人 (物件編號, 店代號, 細項流水號, 權利人流水號, 所有權人, 權利範圍_分子, 權利範圍_分母, 出售權利範圍_分子, 出售權利範圍_分母, 持有面積, 出售面積, 權利總類, 所有權種類, TempCode, 新增日期) "
                insertStr += "Select '" & 複製的物件編號 & "' as 物件編號, '" & 店代號1 & "' as 店代號, 細項流水號, 權利人流水號, 所有權人, 權利範圍_分子, 權利範圍_分母, 出售權利範圍_分子, 出售權利範圍_分母, 持有面積, 出售面積, 權利總類, 所有權種類, TempCode, 新增日期"
                insertStr += " FROM 委賣物件資料表_細項所有權人"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using

                'select into 調查部分
                '產調_土地
                insertStr = "INSERT INTO 產調_土地 (物件編號, 流水號, 店代號, 限制登記, 限制登記_說明, 信託登記, 信託登記_說明, 其他權利, 其他權利_說明, 是否依慣例使用, 是否依慣例使用_說明, "
                insertStr += "共有土地有無分管協議, 共有土地有無分管協議_說明, 有無出租或出借, 有無出租或出借_說明, 有無被他人無權占用, 有無被他人無權占用_說明, 供公眾通行之私有道路, "
                insertStr += "供公眾通行之私有道路_說明, 辦理地籍圖重測, 主管機關已公告辦理, 是否越界建築, 是否越界建築_說明, 公告徵收, 公告徵收範圍, 有無公共基礎設施, "
                insertStr += "有無公共基礎設施_說明, 禁限建地區, 禁限建地區_說明, 不得興建農舍, 不得興建農舍_說明, 山坡地範圍, 山坡地範圍_說明, 水土保持區, 水土保持區_說明, 河川區域, "
                insertStr += "河川區域_說明, 排水設施, 排水設施_說明, 國家公園, 國家公園_說明, 水質保護區, 水質保護區_說明, 水量保護區, 水量保護區_說明, 地下水汙染場址, 地下水汙染場址_說明, "
                insertStr += "新增日期, 新增者, 修改日期, 修改者, 備註)"
                insertStr += " Select '" & 複製的物件編號 & "' as 物件編號, 流水號, '" & 店代號1 & "' as 店代號, 限制登記, 限制登記_說明, 信託登記, 信託登記_說明, 其他權利, 其他權利_說明, 是否依慣例使用, 是否依慣例使用_說明, "
                insertStr += "共有土地有無分管協議, 共有土地有無分管協議_說明, 有無出租或出借, 有無出租或出借_說明, 有無被他人無權占用, 有無被他人無權占用_說明, 供公眾通行之私有道路,"
                insertStr += "供公眾通行之私有道路_說明, 辦理地籍圖重測, 主管機關已公告辦理, 是否越界建築, 是否越界建築_說明, 公告徵收, 公告徵收範圍, 有無公共基礎設施,"
                insertStr += "有無公共基礎設施_說明, 禁限建地區, 禁限建地區_說明, 不得興建農舍, 不得興建農舍_說明, 山坡地範圍, 山坡地範圍_說明, 水土保持區, 水土保持區_說明, 河川區域,"
                insertStr += "河川區域_說明, 排水設施, 排水設施_說明, 國家公園, 國家公園_說明, 水質保護區, 水質保護區_說明, 水量保護區, 水量保護區_說明, 地下水汙染場址, 地下水汙染場址_說明,"
                insertStr += "新增日期, 新增者, '" & sysdate & "' as 修改日期, '" & Request.Cookies("webfly_empno").Value & "' as  修改者, 備註"
                insertStr += " FROM 產調_土地"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using
                '產調_車位
                insertStr = "INSERT INTO 產調_車位 (物件編號, 流水號, 店代號, 細項編號, 獨立產權, 權利種類, 其他權利種類, 建號, 車位類別, 車位號碼, 車位_長, 車位_寬, 車位_高, 車位_承重, 新增日期, 新增者, "
                insertStr += "修改日期, 修改者, 車位說明, 進出口, 車位性質, 出租或占用情形, 車位管理費_類別, 車位管理費_期間, 車位管理費_金額, 車位照片, 車位位置_上下, 車位位置_樓層, 備註, "
                insertStr += "車位獨立售價) "
                insertStr += "Select '" & 複製的物件編號 & "' as 物件編號, 流水號, '" & 店代號1 & "' as 店代號, 細項編號, 獨立產權, 權利種類, 其他權利種類, 建號, 車位類別, 車位號碼, 車位_長, 車位_寬, 車位_高, 車位_承重, 新增日期, 新增者, "
                insertStr += "修改日期, 修改者, 車位說明, 進出口, 車位性質, 出租或占用情形, 車位管理費_類別, 車位管理費_期間, 車位管理費_金額, 車位照片, 車位位置_上下, 車位位置_樓層, 備註,"
                insertStr += "車位獨立售價"
                insertStr += " FROM 產調_車位"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using

                '產調_建物
                insertStr = "INSERT INTO 產調_建物 (物件編號, 流水號, 店代號, 限制登記, 限制登記_說明, 信託登記, 信託登記_說明, 其他權利, 其他權利_說明, 是否共有, 有無分管協議登記, 有無分管協議登記_說明, "
                insertStr += "專有部分範圍_有無, 專有部分範圍, 共有部分範圍_有無, 共有部分範圍, 專有約定共用, 專有約定共用之範圍, 專有約定共用之使用方式, 共有約定專用, 共有約定專用之範圍,"
                insertStr += "共有約定專用之使用方式, 公共基金有無, 公共基金數額, 公共基金提撥方式, 公共基金運用方式, 管理費使用, 管理費或使用費, 管理費, 管理費單位, 管理費繳交方式,"
                insertStr += "管理組織_有無, 管理組織及方式, 管理組織及方式_其他, 管理公司_有無, 管理公司, 管理手冊, 管理手冊_說明, 獎勵容積開放空間提供公共使用,"
                insertStr += "獎勵容積開放空間提供公共使用_說明, 電梯設備, 張貼合格標章, 張貼合格標章_說明, 出租狀況, 出租範圍, 出租範圍備註, 出租情況類型, 租金, 租期起, 租期迄, 租約是否公證,"
                insertStr += "出租狀況_說明, 出借狀況, 出借狀況_說明, 出借範圍, 出借範圍備註, 出借書面約定與否, 借用人姓名, 出借起日, 出借迄日, 押租保證金, 出借返還條件, 佔用情形,"
                insertStr += "佔用他人建物土地, 佔用他人建物土地_說明, 被他人佔用建物土地, 被他人佔用建物土地_說明, 佔用情形其他, 佔用情形其他_說明, 消防設備, 消防設備_說明, 無障礙設施,"
                insertStr += "無障礙設施_說明, 夾層, 夾層面積, 是否合法登記之夾層, 夾層面積1, 是否合法登記之夾層1, 夾層面積2, 是否合法登記之夾層2, 夾層其他, 獨立供水, 獨立供水_說明, 供水類型,"
                insertStr += "供水是否正常, 獨立電表, 獨立電表_說明, 天然瓦斯, 天然瓦斯_說明, 天然瓦斯_說明2, 水電管線是否更新, 水電管線是否更新_說明, 水管更新日期, 電線更新日期, 積欠應繳費用,"
                insertStr += "積欠應繳費用_說明, 屬工業區或其他分區, 屬工業區或其他分區_說明, 持有期間有無居住, 使用執照有無備註之注意事項, 使用執照有無備註之注意事項_說明,"
                insertStr += "有無公共設施重大修繕, 有無公共設施重大修繕_說明, 有無公共設施重大修繕_金額, 混凝土中氯離子含量, 混凝土中氯離子含量_說明, 輻射檢測, 輻射檢測_說明,"
                insertStr += "曾發生火災或其他災害, 曾發生火災或其他災害_說明, 因地震被公告為危險建築, 因地震被公告為危險建築_說明, 樑柱部分是否有顯見裂痕, 樑柱部分是否有顯見裂痕_說明,"
                insertStr += "裂痕長度, 間隙寬度, 建物鋼筋裸露, 建物鋼筋裸露_說明, 是否為兇宅, 兇宅發生期間, 是否為兇宅_說明, 非持有是否為兇宅, 非持有是否為兇宅_說明, 滲漏水狀態,"
                insertStr += "滲漏水狀態_處理, 滲漏水狀態_說明, 有無禁建情事, 有無禁建情事_說明, 違增建使用權, 違增建使用權_說明, 違增建_面積, 違增建列管情形, 違增建列管情形_說明, 排水系統,"
                insertStr += "排水系統_說明, 其他重要事項, 其他重要事項_說明, 新增日期, 新增者, 修改日期, 修改者, 備註, 頂樓基地台, 頂樓基地台_說明, 衛生下水道工程, 衛生下水道工程_選項,"
                insertStr += "衛生下水道工程_說明, 規約外特殊使用, 規約外特殊使用_共用說明, 規約外特殊使用_專用說明, 住戶規約使用手冊, 隨附設備有無, 隨附設備, 沙發數, 電視數, 冰箱數, 冷氣數,"
                insertStr += "洗衣機數, 乾衣機數,全棟,中繼幫浦,中繼幫浦_說明) "
                insertStr += "Select '" & 複製的物件編號 & "' as 物件編號, 流水號, '" & 店代號1 & "' as 店代號, 限制登記, 限制登記_說明, 信託登記, 信託登記_說明, 其他權利, 其他權利_說明, 是否共有, 有無分管協議登記, 有無分管協議登記_說明, "
                insertStr += "專有部分範圍_有無, 專有部分範圍, 共有部分範圍_有無, 共有部分範圍, 專有約定共用, 專有約定共用之範圍, 專有約定共用之使用方式, 共有約定專用, 共有約定專用之範圍,"
                insertStr += "共有約定專用之使用方式, 公共基金有無, 公共基金數額, 公共基金提撥方式, 公共基金運用方式, 管理費使用, 管理費或使用費, 管理費, 管理費單位, 管理費繳交方式,"
                insertStr += "管理組織_有無, 管理組織及方式, 管理組織及方式_其他, 管理公司_有無, 管理公司, 管理手冊, 管理手冊_說明, 獎勵容積開放空間提供公共使用,"
                insertStr += "獎勵容積開放空間提供公共使用_說明, 電梯設備, 張貼合格標章, 張貼合格標章_說明, 出租狀況, 出租範圍, 出租範圍備註, 出租情況類型, 租金, 租期起, 租期迄, 租約是否公證,"
                insertStr += "出租狀況_說明, 出借狀況, 出借狀況_說明, 出借範圍, 出借範圍備註, 出借書面約定與否, 借用人姓名, 出借起日, 出借迄日, 押租保證金, 出借返還條件, 佔用情形,"
                insertStr += "佔用他人建物土地, 佔用他人建物土地_說明, 被他人佔用建物土地, 被他人佔用建物土地_說明, 佔用情形其他, 佔用情形其他_說明, 消防設備, 消防設備_說明, 無障礙設施,"
                insertStr += "無障礙設施_說明, 夾層, 夾層面積, 是否合法登記之夾層, 夾層面積1, 是否合法登記之夾層1, 夾層面積2, 是否合法登記之夾層2, 夾層其他, 獨立供水, 獨立供水_說明, 供水類型,"
                insertStr += "供水是否正常, 獨立電表, 獨立電表_說明, 天然瓦斯, 天然瓦斯_說明, 天然瓦斯_說明2, 水電管線是否更新, 水電管線是否更新_說明, 水管更新日期, 電線更新日期, 積欠應繳費用,"
                insertStr += "積欠應繳費用_說明, 屬工業區或其他分區, 屬工業區或其他分區_說明, 持有期間有無居住, 使用執照有無備註之注意事項, 使用執照有無備註之注意事項_說明,"
                insertStr += "有無公共設施重大修繕, 有無公共設施重大修繕_說明, 有無公共設施重大修繕_金額, 混凝土中氯離子含量, 混凝土中氯離子含量_說明, 輻射檢測, 輻射檢測_說明,"
                insertStr += "曾發生火災或其他災害, 曾發生火災或其他災害_說明, 因地震被公告為危險建築, 因地震被公告為危險建築_說明, 樑柱部分是否有顯見裂痕, 樑柱部分是否有顯見裂痕_說明,"
                insertStr += "裂痕長度, 間隙寬度, 建物鋼筋裸露, 建物鋼筋裸露_說明, 是否為兇宅, 兇宅發生期間, 是否為兇宅_說明, 非持有是否為兇宅, 非持有是否為兇宅_說明, 滲漏水狀態,"
                insertStr += "滲漏水狀態_處理, 滲漏水狀態_說明, 有無禁建情事, 有無禁建情事_說明, 違增建使用權, 違增建使用權_說明, 違增建_面積, 違增建列管情形, 違增建列管情形_說明, 排水系統,"
                insertStr += "排水系統_說明, 其他重要事項, 其他重要事項_說明, 新增日期, 新增者, 修改日期, 修改者, 備註, 頂樓基地台, 頂樓基地台_說明, 衛生下水道工程, 衛生下水道工程_選項,"
                insertStr += "衛生下水道工程_說明, 規約外特殊使用, 規約外特殊使用_共用說明, 規約外特殊使用_專用說明, 住戶規約使用手冊, 隨附設備有無, 隨附設備, 沙發數, 電視數, 冰箱數, 冷氣數,"
                insertStr += "洗衣機數, 乾衣機數,全棟,中繼幫浦,中繼幫浦_說明 "
                insertStr += " FROM 產調_建物 "
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using

                '產調_基地
                insertStr = "INSERT INTO 產調_基地 ( 物件編號, 流水號, 店代號, 共有土地有無分管協議, 共有土地有無分管協議_說明, 有無出租或出借, 有無出租或出借_選項, 有無出租或出借_說明, "
                insertStr += "      供公眾通行之私有道路, 供公眾通行之私有道路_說明, 可通行對外道路, 可通行對外道路_說明, 界址糾紛, 糾紛對象說明, 辦理地籍圖重測, 主管機關已公告辦理, 公告徵收, "
                insertStr += "      公告徵收範圍, 新增日期, 新增者, 修改日期, 修改者, 列管山坡地, 列管山坡地說明, 備註, 捷運穿越, 捷運穿越說明, 基地占用, 基地占用說明, 開發限制, 開發限制說明, "
                insertStr += "      其他重要事項, 其他重要事項說明) "
                insertStr += "Select '" & 複製的物件編號 & "' as 物件編號, 流水號, '" & 店代號1 & "' as 店代號, 共有土地有無分管協議, 共有土地有無分管協議_說明, 有無出租或出借, 有無出租或出借_選項, 有無出租或出借_說明, "
                insertStr += "      供公眾通行之私有道路, 供公眾通行之私有道路_說明, 可通行對外道路, 可通行對外道路_說明, 界址糾紛, 糾紛對象說明, 辦理地籍圖重測, 主管機關已公告辦理, 公告徵收, "
                insertStr += "      公告徵收範圍, 新增日期, 新增者, 修改日期, 修改者, 列管山坡地, 列管山坡地說明, 備註, 捷運穿越, 捷運穿越說明, 基地占用, 基地占用說明, 開發限制, 開發限制說明, "
                insertStr += "      其他重要事項, 其他重要事項說明 "
                insertStr += " FROM 產調_基地"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using


            End If


        End If

    End Sub
    'Sub Mod_複製委賣物件細項面積(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String)

    '    Dim dr1 As SqlDataReader

    '    '被複制的物件編號

    '    sql = "Select * From " & TableName & " With(NoLock) Where 店代號 = '" & Request("sid") & "' and 物件編號 = '" & 原物件編號 & "' "


    '    adpt = New SqlDataAdapter(sql, conn)
    '    ds = New DataSet()
    '    adpt.Fill(ds, "table1")
    '    table1 = ds.Tables("table1")

    '    If table1.Rows.Count Then

    '        sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
    '        adpt = New SqlDataAdapter(sql, conn)
    '        ds = New DataSet()
    '        adpt.Fill(ds, "table2")
    '        table2 = ds.Tables("table2")
    '        If table2.Rows.Count Then
    '            '若已存在就先刪除這一筆
    '            sql = "Delete From " & TableName & " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
    '            cmd = New SqlCommand(sql, conn)
    '            cmd.ExecuteNonQuery()
    '        End If

    '        Dim z As Integer = 0

    '        For z = 0 To table1.Rows.Count - 1
    '            Dim cmd3 As String = "", cmd4 As String = "", sql2 As String = ""

    '            sql = "Insert Into " & TableName
    '            cmd3 = " (" : cmd4 = " VALUES ("

    '            '被複制的物件編號

    '            sql = "Select * From " & TableName & " With(NoLock) Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "' "


    '            Dim conn2 As New SqlConnection(EGOUPLOADSqlConnStr)
    '            conn2.Open()
    '            cmd = New SqlCommand(sql, conn2)
    '            dr1 = cmd.ExecuteReader

    '            sql = "Insert Into " & TableName

    '            For i  = 0 To dr1.FieldCount - 1
    '                sql2 = dr1.GetName(i)
    '                Select Case dr1.GetFieldType(i).FullName
    '                    Case "System.String", "System.DateTime"
    '                        '判斷是否要對照店代號
    '                        Select Case dr1.GetName(i)
    '                            Case "物件編號"
    '                                cmd3 = cmd3 & sql2 & ","
    '                                cmd4 = cmd4 & "'" & 複製的物件編號 & "',"
    '                            Case "店代號"
    '                                cmd3 = cmd3 & sql2 & ","
    '                                cmd4 = cmd4 & "'" & 店代號1 & "',"
    '                            Case Else
    '                                cmd3 = cmd3 & sql2 & ","
    '                                cmd4 = cmd4 & "'" & table1.Rows(z)(dr1.GetName(i)) & "',"
    '                        End Select
    '                    Case "System.Byte", "System.Decimal", "System.Double", "System.SByte", "System.Single", "System.Int16", "System.Int32", "System.Int64", "System.UInt16", "System.UInt32", "System.UInt64"
    '                        Select Case dr1.GetName(i)
    '                            Case Else
    '                                cmd3 = cmd3 & sql2 & ","
    '                                cmd4 = cmd4 & Val(table1.Rows(z)(sql2)) & ","
    '                        End Select
    '                    Case "System.Boolean"
    '                        cmd3 = cmd3 & sql2 & ","
    '                        cmd4 = cmd4 & IIf(table1.Rows(z)(dr1.GetName(i)) <> "", 1, 0) & ","
    '                    Case Else
    '                End Select
    '            Next

    '            cmd3 = Mid(cmd3, 1, Len(cmd3) - 1) & ")"
    '            cmd4 = Mid(cmd4, 1, Len(cmd4) - 1) & ")"
    '            sql += cmd3 & cmd4
    '            cmd = New SqlCommand(sql, conn)
    '            cmd.ExecuteNonQuery()
    '            conn2.Close()
    '        Next



    '    End If

    'End Sub
    Sub Mod_複製委賣相關資料(ByVal TableName As String, ByVal 原物件編號 As String, ByVal 複製的物件編號 As String, ByVal 店代號1 As String, ByVal 業務1 As String, ByVal 業務2 As String, ByVal 業務3 As String, ByVal 秘書1 As String)

        sql = "Select * From " & TableName & " With(NoLock) Where 店代號 = '" & Request("sid") & "' and 物件編號 = '" & 原物件編號 & "' "
        Dim table1 As New DataTable
        Using Conn_委賣相關資料 As New SqlConnection(EGOUPLOADSqlConnStr)
            Conn_委賣相關資料.Open()
            Using com As New SqlCommand(sql, Conn_委賣相關資料)
                table1.Load(com.ExecuteReader)
            End Using
        End Using
        'Response.Write(sql)

        If table1.Rows.Count > 0 Then

            '刪除已存在的
            Using conn_delexist As New SqlConnection(EGOUPLOADSqlConnStr)
                conn_delexist.Open()
                Dim sql_delexist As String = "Delete From " & TableName & " Where 物件編號 = '" & 複製的物件編號 & "' and 店代號 = '" & 店代號1 & "' "
                'Response.Write(sql_delexist)
                Using cmd_deleExist As New SqlCommand(sql_delexist, conn_delexist)
                    cmd_deleExist.ExecuteNonQuery()
                End Using
            End Using
            If TableName = "委賣屋主資料表" Then
                Dim insertStr As String
                insertStr = "INSERT INTO 委賣屋主資料表 (物件編號, 身分證字號, 客戶姓名, 性別, 婚姻, 結婚紀念日, 人口數, 生日, 郵遞區號, 商圈代號, 大樓代號, 完整地址, 縣市, 鄉鎮市區, 村里, 村里別, 鄰, 路名, 路別, 段, 巷, "
                insertStr += "弄, 號, 之, 所在樓層, 樓之, 電話, 電話1, 電話2, 電子郵件, 行業別, 職稱, 公司名稱, 客戶來源, 銷售狀態, 買方轉屋主, 店代號, 經紀人代號, 營業員代號1, 營業員代號2, 輸入者,"
                insertStr += "新增日期, 修改日期, 上傳註記, 上傳日期, 永久_郵遞區號, 永久_完整地址, 永久_縣市, 永久_鄉鎮市區, 永久_村里, 永久_村里別, 永久_鄰, 永久_路名, 永久_路別, 永久_段,"
                insertStr += "永久_巷, 永久_弄, 永久_號, 永久_之, 永久_所在樓層, 永久_樓之, 同永久地址, 備註, 開放, 簡訊種類)"
                insertStr += " Select  '" & 複製的物件編號 & "' as 物件編號, 身分證字號, 客戶姓名, 性別, 婚姻, 結婚紀念日, 人口數, 生日, 郵遞區號, 商圈代號, 大樓代號, 完整地址, 縣市, 鄉鎮市區, 村里, 村里別, 鄰, 路名, 路別, 段, 巷, "
                insertStr += "弄, 號, 之, 所在樓層, 樓之, 電話, 電話1, 電話2, 電子郵件, 行業別, 職稱, 公司名稱, 客戶來源, 銷售狀態, 買方轉屋主, '" & 店代號1 & "' as 店代號, 經紀人代號, 營業員代號1, 營業員代號2, 輸入者,"
                insertStr += "新增日期, 修改日期, 上傳註記, 上傳日期, 永久_郵遞區號, 永久_完整地址, 永久_縣市, 永久_鄉鎮市區, 永久_村里, 永久_村里別, 永久_鄰, 永久_路名, 永久_路別, 永久_段,"
                insertStr += "永久_巷, 永久_弄, 永久_號, 永久_之, 永久_所在樓層, 永久_樓之, 同永久地址, 備註, 開放, 簡訊種類 "
                insertStr += " FROM 委賣屋主資料表"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                'Response.Write(insertStr)
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using
            End If

            If TableName = "委賣_房地產說明書" Then
                Dim insertStr As String
                insertStr = "INSERT INTO 委賣_房地產說明書 (物件編號, 報告單位, 報告編號_年, 報告編號_部門, 報告編號_序號, 報告日期, 地政事務所所在區域, 地政事務所, 謄本核發日期, 產權調查, 物件個案調查, 照片說明, "
                insertStr += "      重要交易條件, 成交行情, 重要設施, 其他說明, 土地權狀影本, 建物權狀影本, 土地謄本, 建物謄本, 地籍圖, 建物平面圖, 房屋稅單, 房地產標的現況說明書, "
                insertStr += "      都市土地使用分區域非都市土地使用種類證明, 使用執照, 住戶規約, 預售買賣契約書, 建物勘測成果圖, 土地相關位置略圖, 建物相關位置略圖, 停車位位置圖, 土地分管協議, "
                insertStr += "      建物分管協議, 樑柱顯見裂痕照片, 其他, 基地面積, 權利範圍, 使用管制內容, 法定建蔽率, 法定容積率, 開發限制方式, 所有權型態為, 共有土地有無分管協議, 是否受限制處分, "
                insertStr += "      有無出租或占用, 建物權利範圍, 建物若無使用執照一併說明, 建物目前管理及使用情況, 住戶規約內容, 專有部分之範圍, 共有部分之範圍, 建物有無共有約定專用部分, "
                insertStr += "      建物範圍為_共有, 使用方式_共有, 建物有無專有部分約定共用, 建物範圍為, 使用方式, 公共基金之數額為新台幣, 其運用方式為, 管理組織及其管理方式, 有否使用手冊, "
                insertStr += "      有否受限制處分, 有無檢測海砂, 有無檢測輻射含量, 有否辦理單獨區分有建物登記, 使用約定方式, 進出口為, 車位號碼, 位置地上地下, 位置地下, 權利種類1, 順位1, 登記日期1, "
                insertStr += "      設定1, 設定權利人1, 權利種類2, 順位2, 登記日期2, 設定2, 設定權利人2, 權利種類3, 順位3, 登記日期3, 設定3, 設定權利人3, 與土地他項權利部相同, 其他如下, 權利種類4, 順位4, "
                insertStr += "      登記日期4, 設定4, 設定權利人4, 權利種類5, 順位5, 登記日期5, 設定5, 設定權利人5, 權利種類6, 順位6, 登記日期6, 設定6, 設定權利人6, 固定物, 電話, 電話線, 梳妝台, 燈飾, 冷氣, "
                insertStr += "      冷氣台, 窗簾, 床組, 冰箱, 冰箱台, 熱水器, 沙發組, 瓦斯廚具, 瓦斯廚具樣式, 其他項目, 個案特色, 物件標的, 現況, 交屋情況, 商談交屋情況, 中庭花園, 其他中庭花園, 警衛管理, "
                insertStr += "      其他警衛管理, 外牆外飾, 其他外牆外飾, 地板, 其他地板, 自來水, 未安裝自來水原因, 電力系統, 有無獨立電錶, 電話系統, 未安裝電話系統原因, 瓦斯系統, 未安裝瓦斯系統, "
                insertStr += "      委託價格新台幣, 簽約金, 第一期金額, 備証款, 第二期金額, 完稅款, 第三期金額, 尾款, 第四期金額, 自用土地增值稅約, 一般增值稅約, 契稅約, 地價稅約, 房屋稅約, "
                insertStr += "      工程受益費約, 代書費約, 登記規費約, 公證費約, 印花稅約, 水電費約, 管理費約, 電話費約, 瓦斯費約, 奢侈稅約, 增值稅, 契稅, 地價稅, 房屋稅, 工程受益費, 代書費, 登記規費, "
                insertStr += "      公證費, 印花稅, 水電費, 管理費, 電話費, 瓦斯費, 奢侈稅, 店代號, 經紀人代號, 營業員代號1, 營業員代號2, 輸入者, 上傳註記, 上傳日期, 分區使用證明, 委託銷售契約書, "
                insertStr += "      建物物件與環境相關費用, 增值稅概算, 共有部分之範圍1, 壁櫥, 酒櫃, 自來瓦斯, 室內建材, 建築結構, 隔間材料, 基地面積1, 土地權利範圍1, 地目, 建築用途, 新增日期, 修改日期, "
                insertStr += "      權利種類7, 順位7, 登記日期7, 設定7, 設定權利人7, 權利種類8, 順位8, 登記日期8, 設定8, 設定權利人8, 洗衣機, 乾衣機, 系統櫥櫃, 系統櫥櫃組, 天然瓦斯度數表, 基地面積坪1, "
                insertStr += "      洗衣機台, 乾衣機台, 飲水機, 其他項目內容, 重新產生XML, 車位性質, 實價登錄費, 代書費New)"
                insertStr += " Select  '" & 複製的物件編號 & "' as 物件編號, 報告單位, 報告編號_年, 報告編號_部門, 報告編號_序號, 報告日期, 地政事務所所在區域, 地政事務所, 謄本核發日期, 產權調查, 物件個案調查, 照片說明, "
                insertStr += "      重要交易條件, 成交行情, 重要設施, 其他說明, 土地權狀影本, 建物權狀影本, 土地謄本, 建物謄本, 地籍圖, 建物平面圖, 房屋稅單, 房地產標的現況說明書, "
                insertStr += "      都市土地使用分區域非都市土地使用種類證明, 使用執照, 住戶規約, 預售買賣契約書, 建物勘測成果圖, 土地相關位置略圖, 建物相關位置略圖, 停車位位置圖, 土地分管協議, "
                insertStr += "      建物分管協議, 樑柱顯見裂痕照片, 其他, 基地面積, 權利範圍, 使用管制內容, 法定建蔽率, 法定容積率, 開發限制方式, 所有權型態為, 共有土地有無分管協議, 是否受限制處分, "
                insertStr += "      有無出租或占用, 建物權利範圍, 建物若無使用執照一併說明, 建物目前管理及使用情況, 住戶規約內容, 專有部分之範圍, 共有部分之範圍, 建物有無共有約定專用部分, "
                insertStr += "      建物範圍為_共有, 使用方式_共有, 建物有無專有部分約定共用, 建物範圍為, 使用方式, 公共基金之數額為新台幣, 其運用方式為, 管理組織及其管理方式, 有否使用手冊, "
                insertStr += "      有否受限制處分, 有無檢測海砂, 有無檢測輻射含量, 有否辦理單獨區分有建物登記, 使用約定方式, 進出口為, 車位號碼, 位置地上地下, 位置地下, 權利種類1, 順位1, 登記日期1, "
                insertStr += "      設定1, 設定權利人1, 權利種類2, 順位2, 登記日期2, 設定2, 設定權利人2, 權利種類3, 順位3, 登記日期3, 設定3, 設定權利人3, 與土地他項權利部相同, 其他如下, 權利種類4, 順位4, "
                insertStr += "      登記日期4, 設定4, 設定權利人4, 權利種類5, 順位5, 登記日期5, 設定5, 設定權利人5, 權利種類6, 順位6, 登記日期6, 設定6, 設定權利人6, 固定物, 電話, 電話線, 梳妝台, 燈飾, 冷氣, "
                insertStr += "      冷氣台, 窗簾, 床組, 冰箱, 冰箱台, 熱水器, 沙發組, 瓦斯廚具, 瓦斯廚具樣式, 其他項目, 個案特色, 物件標的, 現況, 交屋情況, 商談交屋情況, 中庭花園, 其他中庭花園, 警衛管理, "
                insertStr += "      其他警衛管理, 外牆外飾, 其他外牆外飾, 地板, 其他地板, 自來水, 未安裝自來水原因, 電力系統, 有無獨立電錶, 電話系統, 未安裝電話系統原因, 瓦斯系統, 未安裝瓦斯系統, "
                insertStr += "      委託價格新台幣, 簽約金, 第一期金額, 備証款, 第二期金額, 完稅款, 第三期金額, 尾款, 第四期金額, 自用土地增值稅約, 一般增值稅約, 契稅約, 地價稅約, 房屋稅約, "
                insertStr += "      工程受益費約, 代書費約, 登記規費約, 公證費約, 印花稅約, 水電費約, 管理費約, 電話費約, 瓦斯費約, 奢侈稅約, 增值稅, 契稅, 地價稅, 房屋稅, 工程受益費, 代書費, 登記規費, "
                insertStr += "      公證費, 印花稅, 水電費, 管理費, 電話費, 瓦斯費, 奢侈稅,'" & 店代號1 & "' as 店代號, 經紀人代號, 營業員代號1, 營業員代號2, 輸入者, 上傳註記, 上傳日期, 分區使用證明, 委託銷售契約書, "
                insertStr += "      建物物件與環境相關費用, 增值稅概算, 共有部分之範圍1, 壁櫥, 酒櫃, 自來瓦斯, 室內建材, 建築結構, 隔間材料, 基地面積1, 土地權利範圍1, 地目, 建築用途, 新增日期, 修改日期, "
                insertStr += "      權利種類7, 順位7, 登記日期7, 設定7, 設定權利人7, 權利種類8, 順位8, 登記日期8, 設定8, 設定權利人8, 洗衣機, 乾衣機, 系統櫥櫃, 系統櫥櫃組, 天然瓦斯度數表, 基地面積坪1, "
                insertStr += "      洗衣機台, 乾衣機台, 飲水機, 其他項目內容, 重新產生XML, 車位性質, 實價登錄費, 代書費New"
                insertStr += " FROM 委賣_房地產說明書"
                insertStr += " Where 物件編號 = '" & 原物件編號 & "' and 店代號 = '" & Request("sid") & "'"
                'Response.Write(insertStr)
                Using conn_addDetail As New SqlConnection(EGOUPLOADSqlConnStr)
                    conn_addDetail.Open()
                    Using cmd_addDetail As New SqlCommand(insertStr, conn_addDetail)
                        cmd_addDetail.ExecuteNonQuery()
                    End Using
                End Using
            End If


        End If

    End Sub


    'Function vFieldVal(ByVal fval As Object) As Object
    '    If IsDBNull(fval) Then
    '        vFieldVal = ""
    '    Else
    '        On Error Resume Next
    '        vFieldVal = Trim$(CStr(fval))
    '        If Err.Number = 94 Then vFieldVal = ""
    '        On Error GoTo 0
    '    End If
    'End Function

    'Function vVal(ByVal Rval As Object) As Single
    '    vVal = Val(vFieldVal(Rval))
    'End Function

    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList1.SelectedIndexChanged
        Dim message As String = ""


        '讀入物件資料
        載入更新或複製初始頁面()

        '讀入不動產說明書資料
        載入更新或複製初始頁面_不動產說明書()


        '學區使用的SESSION值
        Session("County") = DDL_County.SelectedValue
        Session("Town") = DDL_Area.SelectedValue

        If Trim(TextBox2.Text) <> "" Then
            '一開始的狀態---------------------------
            '解除頁面控制項
            enablefalse("unlock")
            '銷售狀態
            DropDownList21.Enabled = False

            '複製的按鈕
            Me.ImageButton13.Visible = True
            Me.CheckBox26.Visible = True
            '-----------------------------------------

            If RadioButtonList1.SelectedValue = "many" Then '若為一約多屋,物件編號不得修改
                TextBox2.Enabled = False
                ddl契約類別.Enabled = False

                Label28.Visible = False
                DropDownList4.Visible = False

                If TextBox17.Text = "" Or TextBox17.Text = "請填寫地號" Then
                    '1010501若為一約多屋,需填入建號或地號
                    If Right(Trim(DropDownList3.SelectedValue), 1) = "地" Then
                        message = "原物件之地號為必填欄位，請先修改再複製!!"
                        '複製的按鈕
                        Me.ImageButton13.Visible = False
                        Me.CheckBox26.Visible = False
                    End If
                End If

                If TextBox18.Text.Trim = "" Or TextBox18.Text = "請填寫建號" Then
                    '1010501若為一約多屋,需填入建號或地號
                    If Right(Trim(DropDownList3.SelectedValue), 1) <> "地" Then
                        If DropDownList3.SelectedValue <> "預售屋" Then
                            message = "原物件之建號為必填欄位，請先修改再複製!!"
                            '複製的按鈕
                            Me.ImageButton13.Visible = False
                            Me.CheckBox26.Visible = False
                        End If
                    End If
                End If

                If message <> "" Then
                    Dim script As String = ""
                    script += "alert('" & message & "');"
                    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                    Exit Sub
                End If

            ElseIf RadioButtonList1.SelectedValue = "flow" Then '若為流通件,物件編號不得修改
                '鎖定頁面控制項
                enablefalse("lock")
                '案名
                input4.Visible = True
                input4.Disabled = False
                '經紀人
                sale1.Enabled = True
                '營業員1
                sale2.Enabled = True
                '營業員2
                sale3.Enabled = True
                '訴求重點  
                TextBox102.Enabled = True
                '複製的按鈕
                Me.ImageButton13.Visible = True
                Me.CheckBox26.Visible = True
            End If
        End If
    End Sub
    '複製
    Protected Sub ImageButton13_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton13.Click

        '20111111新增判斷,平方公尺欄位不得為空值
        Dim Str As String = chk_平方公尺()
        If Str <> "" Then
            Dim Script As String = ""
            Script += "alert('" & Str & "');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

            Exit Sub
        End If

        Dim j = 驗證數字是否阿拉伯數字()
        If j > 0 Then
            Exit Sub
        End If

        Dim amessage As String = ""
        If DropDownList3.Text <> "土地" Then
            If TextBox88.Text = "" Or TextBox89.Text = "" Or TextBox90.Text = "" Then
                amessage &= "當不為土地時，地上、地下、所在樓層 不可為空 \n"
            End If
            If Trim(Text2.Value) = "" Then
                amessage &= "當不為土地時，完工年月 不可為空 \n"
            Else
                If Trim(Text2.Value) <> "00000" Then
                    If Not IsDate(Left(Text2.Value, 3) + 1911 & "/" & Mid(Text2.Value, 4, 2) & "/" & "01") Then
                        amessage &= "完工年月輸入錯誤 \n"
                    End If
                End If
            End If
        End If
        If Trim(Text11.Value) <> "" Then
            If Not IsDate(Left(Text11.Value, 3) + 1911 & "/" & Mid(Text11.Value, 4, 2) & "/" & Mid(Text11.Value, 6, 2)) Then
                amessage &= "登記日期輸入錯誤 \n"
            End If
        End If
        If Trim(Date2.Text) <> "" Then
            If Not IsDate(Left(Date2.Text, 3) + 1911 & "/" & Mid(Date2.Text, 4, 2) & "/" & Mid(Date2.Text, 6, 2)) Then
                amessage &= "委託起始日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託起始日期 \n"
        End If
        If Trim(Date3.Text) <> "" Then
            If Not IsDate(Left(Date3.Text, 3) + 1911 & "/" & Mid(Date3.Text, 4, 2) & "/" & Mid(Date3.Text, 6, 2)) Then
                amessage &= "委託截止日期輸入錯誤 \n"
            End If
        Else
            amessage &= "請輸入委託截止日期 \n"
        End If
        If Trim(Date2.Text) = "" And Trim(Date3.Text) = "" Then

        Else
            If CType(Left(Trim(Date3.Text), 7), Integer) < CType(Left(Trim(Date2.Text), 7), Integer) Then
                amessage &= "委託截止日須在委託起始日之後 \n"
            End If
        End If

        If DropDownList3.SelectedValue = "土地" Or DropDownList3.SelectedValue = "透天" Then

        Else
            If TextBox91.Text = "" Or TextBox92.Text = "" Then
                amessage &= "非土地時，每層戶數及電梯數不可為空 \n"
            End If
        End If

        If amessage = "" Then

        Else
            eip_usual.Show(amessage)
            Exit Sub
        End If

        'If CType(Trim(Me.Date3.Text), Integer) < CType(Trim(Me.Date2.Text), Integer) Then
        '    Dim Script As String = ""
        '    Script += "alert('委託截止日須在委託起始日之後');"
        '    ScriptManager.RegisterStartupScript(Me, GetType(String), "", Script, True)

        '    Exit Sub
        'End If



        Dim sid As String = Request("sid")
        Dim state As String = Request("state")

        Dim 物件編號 As String = ""
        If ddl契約類別.SelectedValue = "專任" Then
            物件編號 = "1"
        ElseIf ddl契約類別.SelectedValue = "一般" Then
            物件編號 = "6"
        ElseIf ddl契約類別.SelectedValue = "同意書" Then
            物件編號 = "7"
        ElseIf ddl契約類別.SelectedValue = "流通" Then
            物件編號 = "5"
        ElseIf ddl契約類別.SelectedValue = "庫存" Then
            物件編號 = "9"
        End If
        物件編號 &= Mid(Request("oid"), 2, 4) & TextBox2.Text
        '20160408 如果物件編號相同則用UPDATE 方式更新原物件資料 by nick
        'If RadioButtonList1.SelectedValue = "normal" And 物件編號 = Request("oid") Then
        '    更新記錄()
        '    Dim script As String = ""
        '    script = "alert('更新完成，您複製的物件編號與原來的相同');"
        '    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        '    Exit Sub
        'Else
        複製記錄()
        'End If
        'If Request.Cookies("webfly_empno").Value = "09W6" Then
        '    Exit Sub
        'End If

        '--------------------------------------------------------------------------------
        ''Dim exist As String = chk_exist(store.SelectedValue, Label57.Text)

        ''If exist <> "True" Then
        If Left(Label57.Text, 1) <> "9" Then
            '判斷是否有購買廣告版位
            Dim torf As String = web_no(store.SelectedValue)
            Dim num As String = index_num(Label57.Text, store.SelectedValue)

            Dim url As String = ""
            If torf = "True" Then
                url = "https://home.etwarm.com.tw/sale-" & Trim(num)
            ElseIf torf = "False" Then
                url = "https://www.etwarm.com.tw/sale-" & Trim(num)
            End If

            '寫入資料表
            ''20190910 10.40.20.66先行拿掉==================================
            ''voice_objects("insert", store.SelectedValue, Label57.Text, url, 1)
            ''20190910 10.40.20.66先行拿掉==================================
        End If
        ''    '---------------------------------------------------------------------------------------------------------------------
        ''End If
        '---------------------------------------------------------------------------------


        If Trim(movie_h.Text) = "Y" Then
            chk_HD(Request("sid"), Request("oid"), "c")
        End If

        If YN = "True" Then
            Dim script As String = ""

            script = "alert('該物件的複製來源並無照片檔案,煩請確認並重新上傳!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        End If

    End Sub

    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged

        If Trim(DropDownList3.SelectedValue) = "預售屋" Then
            Me.Label58.Visible = True
            Me.add11.Visible = True
            'RadioButton3.visible = False
        Else
            Me.Label58.Visible = False
            Me.add11.Visible = False
            Me.add11.Text = ""
        End If

        If Trim(DropDownList3.SelectedValue) = "預售屋" Or Trim(DropDownList3.SelectedValue) = "土地" Then
            'e.Row.Attributes.Add("OnMouseover", "this.style.background='#33CCFF'")
            'UpdatePanel10.Attributes.Add("display", "none")
            RadioButton3.Visible = False
            RadioButton4.Visible = False
            CheckBox102.Visible = False
            CheckBox103.Visible = False
            Label466.Visible = False
            Image20.Visible = False
            'UpdatePanel10.visible = False
            'UpdatePanel45.visible = False
        Else
            'UpdatePanel10.Style.Add("display", "block")
            'UpdatePanel10.visible = True
            'UpdatePanel45.visible = True
            RadioButton3.Visible = True
            RadioButton4.Visible = True
            CheckBox102.Visible = True
            CheckBox103.Visible = True
            Label466.Visible = True
            Image20.Visible = True
        End If

        '如果是土地則關閉一些房屋稅的選項 20160407 by nick
        'If Trim(DropDownList3.SelectedValue) = "土地" Then
        '    input105.Disabled = True
        '    input107.Disabled = True
        '    input113.Disabled = True
        '    input116.Disabled = True
        '    input115.Disabled = True
        '    input114.Disabled = True            
        'Else
        '    input105.Disabled = False
        '    input107.Disabled = False
        '    input113.Disabled = False
        '    input116.Disabled = False
        '    input115.Disabled = False
        '    input114.Disabled = False
        'End If

        Address_change()
    End Sub

    Sub Address_change()
        If Trim(DropDownList3.SelectedValue) = "土地" Then
            '村里
            add1.Text = ""
            add1.Visible = False
            zone3.SelectedIndex = 0
            zone3.Visible = False
            '鄰
            add2.Text = ""
            add2.Visible = False
            Label63.Visible = False
            '路名(段)--------------------------------
            add3.Text = ""
            address20.SelectedValue = "段"
            '路別(小段)------------------------------
            add4.Text = ""
            add4.Width = 75
            add4.MaxLength = 5
            Label64.Text = "小段"
            '巷
            add5.Text = ""
            add5.Visible = False
            Label65.Visible = False
            '弄
            add6.Text = ""
            add6.Visible = False
            Label66.Visible = False
            '號-------------------------------------
            add7.Text = ""
            add7.Width = New Unit("60%")
            add7.MaxLength = 100
            Label67.Text = "地號"
            '樓之
            add8.Text = ""
            add8.Visible = False
            '樓
            add9.Text = ""
            add9.Visible = False
            Label68.Visible = False

            add10.Text = ""
            add10.Visible = False
        Else
            '村里
            add1.Text = ""
            add1.Visible = True
            zone3.SelectedIndex = 0
            zone3.Visible = True
            '鄰
            add2.Text = ""
            add2.Visible = True
            Label63.Visible = True
            '路名(段)--------------------------------
            add3.Text = ""
            address20.SelectedIndex = 0
            '路別(小段)------------------------------
            add4.Text = ""
            add4.Width = 31
            add4.MaxLength = 2
            Label64.Text = "段"
            '巷
            add5.Text = ""
            add5.Visible = True
            Label65.Visible = True
            '弄
            add6.Text = ""
            add6.Visible = True
            Label66.Visible = True
            '號-------------------------------------
            add7.Text = ""
            add7.Width = 31
            add7.MaxLength = 10
            Label67.Text = "號之"
            '樓之
            add8.Text = ""
            add8.Visible = True
            '樓
            add9.Text = ""
            add9.Visible = True
            Label68.Visible = True

            add10.Text = ""
            add10.Visible = True

        End If

    End Sub

    Protected Sub RadioButtonList2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList2.SelectedIndexChanged
        If RadioButtonList2.SelectedIndex = 0 Then
            Response.Redirect("Obj_Add_V4.aspx?state=add&src=NOW")
        ElseIf RadioButtonList2.SelectedIndex = 1 Then
            Response.Redirect("Rent_Obj_Add.aspx?state=add&src=NOW")
        ElseIf RadioButtonList2.SelectedIndex = 2 Then

        End If

    End Sub

    Protected Sub TextBox252_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox252.TextChanged
        社區大樓()
    End Sub

    Protected Sub ImageButton7_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton7.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""


        href = "../D_Community/Community_Area.aspx?source=notop"


        ntitle = "建議商圈"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    Protected Sub ImageButton8_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton8.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""


        href = "../D_Community/Community_Park.aspx?source=notop"


        ntitle = "建議公園"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    Protected Sub ImageButton22_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton22.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""


        href = "../D_Community/Community_社區.aspx?source=notop"


        ntitle = "建議社區"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    'Protected Sub ImageButton18_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton18.Click
    '    '土地使用分區-物件用途, 
    '    If DropDownList12.Visible = True Then
    '        If DropDownList12.SelectedValue = "請選擇" Then

    '            If Trim(TextBox253.Text) <> "" Then

    '                TextBox253.Text = Replace(TextBox253.Text, DropDownList11.SelectedValue & ",", "") & DropDownList11.SelectedValue & ","
    '            Else
    '                TextBox253.Text &= DropDownList11.SelectedValue & ","
    '            End If
    '        Else
    '            If Trim(TextBox253.Text) <> "" Then
    '                TextBox253.Text = Replace(TextBox253.Text, DropDownList12.SelectedValue & ",", "") & DropDownList12.SelectedValue & ","
    '            Else
    '                TextBox253.Text &= DropDownList12.SelectedValue & ","
    '            End If
    '        End If

    '    Else
    '        If Trim(TextBox253.Text) <> "" Then
    '            TextBox253.Text = Replace(TextBox253.Text, DropDownList11.SelectedValue & ",", "") & DropDownList11.SelectedValue & ","
    '        Else
    '            TextBox253.Text &= DropDownList11.SelectedValue & ","
    '        End If
    '    End If



    '    'If DropDownList16.SelectedValue <> "請選擇" Then
    '    '    Label47.Text = DropDownList16.SelectedValue
    '    '    If Trim(TextBox253.Text) <> "" Then
    '    '        Label47.Text &= "," & Trim(TextBox253.Text) & ","
    '    '    End If
    '    'Else
    '    '    Label47.Text = ""
    '    'End If

    'End Sub

    Function get_num(ByVal sid As String) As String
        Dim i As Integer = 0
        Dim ezcode As String = ""

        Dim conn_num As SqlConnection = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_num.Open()


        Dim sql_ezcode As String = "select ezcode from 委賣物件資料表 With(NoLock) where 店代號='" & sid & "' and left(物件編號,1)<>'9' order by ezcode asc"
        'Response.Write(sql_ezcode)
        adpt = New SqlDataAdapter(sql_ezcode, conn_num)
        ds = New DataSet()
        adpt.Fill(ds, "簡碼")
        Dim t1 As DataTable = ds.Tables("簡碼")

        'If t1.Rows.Count > 0 Then
        '    If Not IsDBNull(t1.Rows(t1.Rows.Count - 1)("ezcode")) Then
        '        If Trim(t1.Rows(t1.Rows.Count - 1)("ezcode")) <> "" Then
        '            '委賣物件RANGE:0001-7500
        '            If t1.Rows(t1.Rows.Count - 1)("ezcode") < "7500" Then
        '                ezcode = String.Format("{0:0000}", CType(t1.Rows(t1.Rows.Count - 1)("ezcode"), Integer) + 1)
        '            ElseIf t1.Rows(t1.Rows.Count - 1)("ezcode") = "7500" Then
        '                For i = 0 To 7499
        '                    If String.Format("{0:0000}", i + 1) <> String.Format("{0:0000}", CType(t1.Rows(i)("ezcode"), Integer)) Then
        '                        ezcode = String.Format("{0:0000}", i + 1)
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '        Else
        '            ezcode = "0001"
        '        End If

        '    Else
        '        ezcode = "0001"
        '    End If

        'Else
        ezcode = "0001"
        'End If

        conn_num.Close()
        conn_num.Dispose()

        Return ezcode

    End Function

    Sub voice_objects(ByVal state As String, ByVal sid As String, ByVal oid As String, ByVal url As String, ByVal status As Integer)


        Dim conn_voice As MySqlConnection = New MySqlConnection(mysqlegoupload)
        If conn_voice.State = ConnectionState.Closed Then
        Else
            conn_voice.Open()

            If state = "insert" Then
                'sql = "Insert into voice_objects (sid,oid,object_id,title,url,status,update_flag,update_time,create_time,updates) values (@sid,@oid,@object_id,@title,@url,@status,@update_flag,'" & Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second & "','" & Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second & "',@updates)"
                sql = "Insert into voice_objects (sid,oid,object_id,title,url,status,update_flag,update_time,create_time,updates) values (@sid,@oid,@object_id,@title,@url,@status,@update_flag,NOW(),NOW(),@updates)"

            ElseIf state = "update" Then
                'sql = "update voice_objects set object_id=@object_id,title=@title,url=@url,status=@status,update_flag=@update_flag,update_time='" & Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second & "',updates=@updates where sid=@sid and oid=@oid"
                sql = "update voice_objects set object_id=@object_id,title=@title,url=@url,status=@status,update_flag=@update_flag,update_time=NOW(),updates=@updates where sid=@sid and oid=@oid"
            End If

            Dim mycmd = New MySqlCommand(sql, conn_voice)
            'sid	
            mycmd.Parameters.Add(New MySqlParameter("@sid", MySqlDbType.String, 5))
            mycmd.Parameters("@sid").Value = sid 'store.SelectedValue


            'oid		
            mycmd.Parameters.Add(New MySqlParameter("@oid", MySqlDbType.String, 16))
            mycmd.Parameters("@oid").Value = oid 'Label57.Text


            'object_id		
            mycmd.Parameters.Add(New MySqlParameter("@object_id", MySqlDbType.String, 5))
            mycmd.Parameters("@object_id").Value = ez_code.Text

            'title		
            mycmd.Parameters.Add(New MySqlParameter("@title", MySqlDbType.String, 40))
            mycmd.Parameters("@title").Value = input4.Value

            'url		
            mycmd.Parameters.Add(New MySqlParameter("@url", MySqlDbType.String, 255))
            mycmd.Parameters("@url").Value = url

            'status		
            mycmd.Parameters.Add(New MySqlParameter("@status", MySqlDbType.Int16))
            mycmd.Parameters("@status").Value = status

            'update_flag		
            mycmd.Parameters.Add(New MySqlParameter("@update_flag", MySqlDbType.Int16))
            mycmd.Parameters("@update_flag").Value = 0


            'updates		
            mycmd.Parameters.Add(New MySqlParameter("@updates", MySqlDbType.String, 7))
            mycmd.Parameters("@updates").Value = sysdate


            mycmd.ExecuteNonQuery()

            conn_voice.Close()
            conn_voice.Dispose()
        End If






    End Sub

    Function web_no(ByVal 店代號 As String) As String

        Dim TorF As String
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        sql = "Select * from hsbsmg "
        sql &= " Where bs_dept = '" & 店代號 & "' and bs_contract_b='Y' "

        cmd = New SqlCommand(sql, conn)

        Dim dr As SqlDataReader = cmd.ExecuteReader
        If dr.Read Then
            TorF = "True"
        Else
            TorF = "False"
        End If

        conn.Close()
        conn.Dispose()

        Return TorF

    End Function

    Function index_num(ByVal oid As String, ByVal sid As String) As String

        Dim num As String
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        sql = "Select * from 委賣物件資料表 "
        sql &= " Where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "

        cmd = New SqlCommand(sql, conn)

        Dim dr As SqlDataReader = cmd.ExecuteReader
        If dr.Read Then
            num = dr("index_num")
            If Not IsDBNull(dr("ezcode")) Then
                ez_code.Text = dr("ezcode")
            End If

        End If

        conn.Close()
        conn.Dispose()

        Return num

    End Function


    Sub 更新屋主承辦人()

        Dim conn_owner As SqlConnection = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_owner.Open()

        sql = "update 委賣屋主資料表 set 經紀人代號='" & Trim(sale1.SelectedValue) & "',營業員代號1='" & Trim(sale2.SelectedValue) & "',營業員代號2='" & Trim(sale3.SelectedValue) & "' where 店代號='" & Request("sid") & "' and 物件編號='" & Request("oid") & "' "

        cmd = New SqlCommand(sql, conn_owner)
        cmd.ExecuteNonQuery()

        cmd.Dispose()
        conn_owner.Close()
        conn_owner.Dispose()

    End Sub

    Sub 新增屋主異動紀錄()
        Dim 欄位 As String = ""
        Dim 舊值 As String = ""
        Dim 新值 As String = ""

        Dim conn_dif As SqlConnection = New SqlConnection(EGOUPLOADSqlConnStr)
        conn_dif.Open()

        i = 1
        For i = 1 To 3
            Select Case i
                Case 1
                    欄位 = "承辦人1"
                    舊值 = Label59.Text.Trim
                    新值 = sale1.SelectedValue.Trim
                Case 2
                    欄位 = "承辦人2"
                    舊值 = Label60.Text.Trim
                    新值 = sale2.SelectedValue.Trim
                Case 3
                    欄位 = "承辦人3"
                    舊值 = Label61.Text.Trim
                    新值 = sale3.SelectedValue.Trim
            End Select

            If 舊值 <> 新值 Then
                sql = "insert into 委賣屋主資料表_異動記錄 (編號,欄位,舊值,新值,店代號,異動人) values ('" & Request("oid") & "','" & 欄位 & "','" & 舊值 & "','" & 新值 & "','" & Request("sid") & "','" & Request.Cookies("webfly_empno").Value & "')"

                cmd = New SqlCommand(sql, conn_dif)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            End If

        Next

        conn_dif.Close()
        conn_dif.Dispose()


        '異動紀錄完，將原本的判斷值更改成新值
        Label59.Text = sale1.SelectedValue.Trim
        Label60.Text = sale2.SelectedValue.Trim
        Label61.Text = sale3.SelectedValue.Trim
    End Sub



    '刪除
    Protected Sub ImageButton19_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton19.Click

        If CheckIsCoSell() Then
            '呼叫更新物件api(webhook)
            UpdateObject("2")
        End If

        Dim TorF As String = chk_deal(Request("oid"), Request("sid"))

        Dim many As String = "True"
        Dim isnewform As String = ""
        If Left(Request("oid"), 1) <> "9" Then
            isnewform = isValidate(Request("oid"))
        End If

        If Mid(UCase(Request("oid").ToString), 6, 1) = "A" Or Mid(UCase(Request("oid").ToString), 6, 1) = "B" Or Mid(UCase(Request("oid").ToString), 6, 1) = "C" Or Mid(UCase(Request("oid").ToString), 6, 1) = "L" Or Mid(UCase(Request("oid").ToString), 6, 1) = "M" Then
            many = chk_many(Mid(UCase(Request("oid").ToString), 1, 13), Request("sid")) '僅為物件13碼
        End If

        Dim script As String

        'If Request("oid") = "60664AAD63489" Then
        '    isnewform = ""
        'End If


        If TorF = "True" Or many = "False" Or isnewform = "New" Then

            '有成交紀錄,無法刪除,須先刪除成交紀錄
            If TorF = "True" Then
                script = "alert('該物件已有成交紀錄,請先刪除成交紀錄才能刪除此物件!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            End If

            '不為一約多屋之最後一筆
            If many = "False" Then
                script = "alert('該物件為一約多屋物件,僅可刪除最後一筆物件，其餘請用修改方式!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            End If

            '新單號不可刪除
            If isnewform = "New" Then
                script = "alert('該物件編號為新的，不可刪除，請改以修改方式!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
            End If



        Else
            '無成交紀錄,可以刪除    
            Try
                conn = New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()

                '刪除委賣物件資料表--20110715修改(接Request("src")參數,判斷為過期還現有物件資料表)
                sql = "DELETE From " & src.Text
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣屋主資料
                sql = "DELETE From 委賣屋主資料表 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣代理人資料
                sql = "DELETE From 委賣代理人資料表 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣_房地產說明書         
                sql = "DELETE From 委賣_房地產說明書 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣價變資料表         
                sql = "DELETE From 委賣價變資料表 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣續約資料表         
                sql = "DELETE From 委賣續約資料表 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除CheckPic         
                sql = "DELETE From CheckPic "
                sql &= " Where pic_contract_no = '" & Request("oid") & "' and pic_dept_no = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣物件相片        
                sql = "DELETE From 委賣物件相片 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除委賣物件細項面積       
                sql = "DELETE From 委賣物件資料表_面積細項 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除他項權利細項       
                sql = "DELETE From 物件他項權利細項 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除土增稅細項       
                sql = "DELETE From 工具_計算土地增值稅 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除物件異動記錄表20120110       
                sql = "DELETE From 物件異動記錄表 "
                sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                '刪除照片.格局圖.地圖.黃金格局圖照片
                Dim FTP目錄 As String = "media\" & Request("sid") & "\" & Me.Label23.Text

                FTP1.Passive = True
                FTP1.Connect("10.40.20.88", 21)
                FTP1.Login("etwarm_media", "Etwa!r@m")
                '照片*8
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "a.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "a.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "b.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "b.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "c.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "c.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "d.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "d.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "w.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "w.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "x.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "x.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "y.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "y.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "z.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "z.jpg")
                '格局圖*1
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "g.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "g.jpg")
                '地圖*1
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "m.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "m.jpg")
                '黃金格局圖*6
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_0.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_0.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_1.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_1.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_2.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_2.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_3.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_3.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_4.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_4.jpg")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "_5.jpg") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "_5.jpg")
                FTP1.Disconnect()

                '刪除影音

                FTP1.Passive = True
                FTP1.Connect("media.etwarm.com.tw", 21)
                FTP1.Login("movie", "movie")

                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "-s.wmv") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "-s.wmv")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "-b.wmv") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "-b.wmv")
                If FTP1.FileExists(FTP目錄 & "\" & Request("oid") & "-h.wmv") Then FTP1.DeleteFile(FTP目錄 & "\" & Request("oid") & "-h.wmv")

                FTP1.Disconnect()

                '新增刪除記錄到委賣物件刪除資料表
                sql = "Insert into 委賣物件刪除資料表(物件編號,刪除日期,刪除人員,店代號) values('" & Request("oid") & "','" & Today.Date & "','" & Request.Cookies("webfly_empno").Value & "','" & Request("sid") & "')"

                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()

                conn.Close()
                conn = Nothing

                '刪除MySQL黃金格局圖紀錄-------------------------------------------------------------------------------
                ''20190910 10.40.20.66先行拿掉==================================
                ''Using MYSQL_conn As New MySqlConnection(mysqletwarmstring)
                ''    MYSQL_conn.Open()
                ''    sql = "Delete From paint_gold Where object_id = '" & Request("oid") & "' and storeid = '" & Request("sid") & "'"
                ''    Using MYSQL_command As New MySqlCommand(sql, MYSQL_conn)
                ''        MYSQL_command.ExecuteNonQuery()
                ''    End Using
                ''End Using
                ''20190910 10.40.20.66先行拿掉==================================
                ''------------------------------------------------------------------------------------------------------

                '新增記錄到高畫質影音異動記錄表
                '--------------------------------------------------------------------------------------------------------
                If Trim(movie_h.Text) = "Y" Then
                    add_hd(Request("sid"), Request("oid"), "d")
                End If
                '--------------------------------------------------------------------------------------------------------


                script = "alert('刪除物件相關資料成功!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            Catch
                conn.Close()
                conn = Nothing

                script = "alert('刪除物件相關資料失敗!!');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            End Try


        End If


    End Sub

    Function chk_deal(ByVal 物件編號 As String, ByVal 店代號 As String) As String
        Dim TorF As String
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        sql = "Select * from 委賣成交資料表 "
        sql &= " Where 物件編號 = '" & 物件編號 & "' and 店代號 = '" & 店代號 & "' "

        cmd = New SqlCommand(sql, conn)
        Dim dr As SqlDataReader = cmd.ExecuteReader
        If dr.Read Then
            TorF = "True"
        Else
            TorF = "False"
        End If

        conn.Close()
        conn.Dispose()

        Return TorF

    End Function

    Function chk_many(ByVal 物件編號 As String, ByVal 店代號 As String) As String '一約多屋除最後一筆外,其餘不可刪除
        Dim many As String
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()
        sql = "Select * from (Select * From 委賣物件資料表 "
        sql &= " union all "
        sql &= " Select * From 委賣物件過期資料表 "
        sql &= " ) as 委賣物件資料表all where 物件編號 like '" & 物件編號 & "%' and 店代號 = '" & 店代號 & "' order by 物件編號 desc"
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        Dim id As String = table1.Rows(0).Item("物件編號")
        If table1.Rows.Count > 1 Then
            If id = Request("oid").ToString Then
                many = "True" '為最後一筆
            Else
                many = "False"
            End If
        Else
            many = "True" '不為一約多屋之物件
        End If

        conn.Close()
        conn.Dispose()

        Return many

    End Function

    Function isValidate(ByVal form_no As String) As String '檢查是否為新表單        
        Dim type As String = ""
        Dim yes As String = "Old"
        If Len(form_no) = 13 Then
            For i = 0 To 2
                type = Asc(Mid(form_no, 6 + i, 1))
                If (type >= 65 And type <= 90) Or (type >= 98 And type <= 122) Then
                    yes = "New"
                End If
            Next
        End If
        Return yes
    End Function

    Sub Load_Paramater()
        Dim store_id As String = Request.Cookies("store_id").Value
        Dim TorF As String = chk_Set(store_id)

        conn = New SqlConnection(EGOUPLOADSqlConnStr)

        If TorF = "True" Then
            sql = "select * from ParamaterSetting With(NoLock)  where StoreID='" & store_id & "' and paramaterID='A' and UseState='Y' order by Num asc"
        Else
            sql = "select * from ParamaterSetting With(NoLock)  where StoreID='None' and paramaterID='A'  and UseState='Y' order by Num asc" '若判斷資料表無任何值，則抓取預設值'None'
        End If

        conn.Open()

        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        Dim t1 As DataTable = ds.Tables("table1")

        i = 0
        For i = 0 To t1.Rows.Count - 1
            If t1.Rows(i)("Num") = "01" Then '過期N天數提醒
                OverDay.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "02" Then '物件何種狀態不秀
                NotShow.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "06" Then '長案名OR短案名
                LongOrShort.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "07" Then '小數點位數
                Float.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "10" Then '件數視窗
                ShowMsg.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "11" Then '截止日
                ShowDt.Text = t1.Rows(i)("value")
            ElseIf t1.Rows(i)("Num") = "12" Then '起始日
                ShowDt_Start.Text = t1.Rows(i)("value")
            End If
        Next

        conn.Close()
        conn = Nothing

        If Trim(Float.Text) = "" Then
            '無值給予預設值
            Float.Text = "4"
        End If
    End Sub

    Function chk_Set(ByVal store_id As String) As String
        Dim TorF As String = ""

        conn = New SqlConnection(EGOUPLOADSqlConnStr)

        sql = "select * from  ParamaterSetting With(NoLock)  where StoreID='" & store_id & "'"
        Dim cmd As SqlCommand = New SqlCommand(sql, conn)
        conn.Open()

        Dim dr As SqlDataReader = cmd.ExecuteReader
        If dr.Read Then
            TorF = "True"
        Else
            TorF = "False"
        End If

        conn.Close()
        conn.Dispose()

        Return TorF
    End Function

    Private Function IsICBAutoUpdateEnabled(ByVal currentStoreId As String) As Boolean
        If String.IsNullOrWhiteSpace(currentStoreId) Then
            Return False
        End If

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            Using cmd As New SqlCommand("select count(1) from ParamaterSetting With(NoLock) where StoreID=@StoreID and Sort='21' and Num='21' and ParamaterID='A' and UseState='Y' and Value='1'", conn)
                cmd.Parameters.AddWithValue("@StoreID", currentStoreId)
                conn.Open()

                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    Private Function JsEncode(ByVal value As String) As String
        If value Is Nothing Then
            Return ""
        End If

        Return value.Replace("\", "\\").Replace("'", "\'")
    End Function

    Sub 自訂坪數_小數點位數(ByVal TxtBox As TextBox, ByVal TorF As String)
        If TorF = "T" Then
            'step1-先取得自訂小數點位數後的數字
            TxtBox.Text = Round(CType(TxtBox.Text, Double), CType(Float.Text, Integer), MidpointRounding.AwayFromZero)
        End If
        'step2-去掉尾巴的0
        If Left(Right(TxtBox.Text, 5), 1) = "." And Right(TxtBox.Text, 4) = "0000" Then
            TxtBox.Text = Int(TxtBox.Text)
        ElseIf Left(Right(TxtBox.Text, 5), 1) = "." And Right(TxtBox.Text, 3) = "000" Then
            TxtBox.Text = Left(TxtBox.Text, TxtBox.Text.Length - 3)
        ElseIf Left(Right(TxtBox.Text, 5), 1) = "." And Right(TxtBox.Text, 2) = "00" Then
            TxtBox.Text = Left(TxtBox.Text, TxtBox.Text.Length - 2)
        ElseIf Left(Right(TxtBox.Text, 5), 1) = "." And Right(TxtBox.Text, 1) = "0" Then
            TxtBox.Text = Left(TxtBox.Text, TxtBox.Text.Length - 1)
        Else
            TxtBox.Text = TxtBox.Text
        End If
    End Sub

    '重要交易條件
    Protected Sub TextBox258_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox258.TextChanged
        計算金額(TextBox258, TextBox262)
    End Sub

    Protected Sub TextBox259_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox259.TextChanged
        計算金額(TextBox259, TextBox263)
    End Sub

    Protected Sub TextBox260_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox260.TextChanged
        計算金額(TextBox260, TextBox264)
    End Sub

    Protected Sub TextBox261_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox261.TextChanged
        計算金額(TextBox261, TextBox265)
    End Sub

    Sub 計算金額(百分比 As TextBox, 金額 As TextBox)
        If Trim(百分比.Text) <> "" Then
            If Trim(TextBox12.Text) <> "" Then
                金額.Text = CType(Trim(TextBox12.Text), Integer) * CType(Trim(百分比.Text), Integer) / 100
            End If
        End If
    End Sub

    '空白文件及範本
    Sub document()
        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            sql = " SELECT Class_b_ID, Class_s_ID, Class_area_ID, Title_name, Title_ID, Num, FileName, FileIntr, FileLocation, Del, UpdateTime FROM Superweb_List_Detail WHERE (FileName LIKE '%不動產說明書%') or (FileName LIKE '%遠距%')  order by UpdateTime desc"
            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table3")
            table3 = ds.Tables("table3")
            content.Text &= "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" bgcolor=""#e7e7e7"" class=""search"">"
            content.Text &= "<tr style=""font-weight:bold"">"
            content.Text &= "<td width=""370"" height=""30"" align=""center"" bgcolor=""#f7f7f7"">檔案名稱</td>"
            content.Text &= "<td align=""center"" bgcolor=""#f7f7f7"">內容說明</td>"
            content.Text &= "<td width=""150"" align=""center"" bgcolor=""#f7f7f7"">更新日期</td>"
            content.Text &= "</tr>"
            For j = 0 To table3.Rows.Count - 1
                content.Text &= "<tr>"
                content.Text &= "<td height=""35"" align=""left"" bgcolor=""#FFFFFF"" id=""td1""><a href=""https://neweip.etwarm.com.tw/Eip/Fly_Head/Upfile/" & table3.Rows(j)("FileLocation").ToString.Trim & """ target=""_blank"">&nbsp;&nbsp;" & table3.Rows(j)("FileName").ToString.Trim & "</a></td>"
                content.Text &= "<td height=""35"" align=""left"" bgcolor=""#FFFFFF"" id=""td1"">" & table3.Rows(j)("FileIntr").ToString.Trim & "</td>"
                content.Text &= "<td align=""center"" bgcolor=""#FFFFFF"" id=""td1"">" & table3.Rows(j)("Updatetime").ToString.Trim & "</td>"
                content.Text &= "</tr>"
            Next
            content.Text &= "</table>"
        End Using
    End Sub

    Protected Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
        '清空hidtempcode
        'If Request.Cookies("webfly_empno").Value = "92h" Then
        '    'If Request.Params.Get("__EVENTTARGET") = "ImageButton5" And Request.Cookies("webfly_empno").Value = "92h" Then
        hidtempcode.Value = ""
        '    'eip_usual.Show("清空value:" & hidtempcode.Value)
        '    'End If
        'End If

    End Sub

    Protected Sub ImageButton20_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton20.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""

        href = "../A_ObjectManage/object_audit_add.aspx?state=add&sid=" & store.SelectedValue & "&oid=" & Request("oid") & "&src=" & Request("src") & "&source=notop"

        ntitle = "變更委託期間"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1200, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    'Protected Sub ImageButton11_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton11.Click
    '    If Request("src") <> "NOW" Then
    '        eip_usual.Show("不為委賣物件，故無法查詢")
    '        Exit Sub
    '    End If

    '    Dim 縣市 As String = ""
    '    Dim 行政區 As String = ""
    '    Dim 單價 As String = ""
    '    Dim 屋齡 As String = ""
    '    Dim 面積 As String = ""
    '    Dim 物件類型 As String = ""
    '    Dim 總價 As String = ""
    '    Dim 經度 As String = ""
    '    Dim 緯度 As String = ""
    '    Dim 型態 As String = "B"
    '    'If Request.Cookies("webfly_empno").Value = "92H" Then
    '    conn = New SqlConnection(EGOUPLOADSqlConnStr)
    '    conn.Open()

    '    sql = " select 店代號,物件編號,縣市,鄉鎮市區 as 行政區,isnull(經度,'') as 經度,isnull(緯度,'') as 緯度, "
    '    sql += " 總坪數 as 面積,刊登售價 as 總價,(CASE WHEN 總坪數=0 THEN 0 ELSE 刊登售價/總坪數 END) AS 單價, "
    '    sql += " (CASE when 物件類別='大樓' or 物件類別='華廈' then '2' "
    '    sql += "       when 物件類別='公寓' then '3' "
    '    sql += "       when 物件類別='透天' then '1' "
    '    sql += "       when 物件類別='土地' then '' "
    '    sql += "       when 物件類別='車位' then 'N' "
    '    sql += "         Else '' end) as 物件類型, "
    '    sql += " isnull(竣工日期,'') as 屋齡, "
    '    sql += " isnull(地上層數,'') as 總樓層數 "
    '    sql += " from 委賣物件資料表 "
    '    sql += " where 店代號='" & Request("sid") & "' and 物件編號='" & Request("oid") & "' "
    '    'Response.Write(Request("sid") & "_" & Request("src") & "_" & Request("oid"))
    '    adpt = New SqlDataAdapter(sql, conn)
    '    ds = New DataSet()
    '    adpt.Fill(ds, "table1")
    '    Dim table1 As DataTable = ds.Tables("table1")

    '    If table1.Rows.Count > 0 Then
    '        縣市 = Trim(table1.Rows(0)("縣市"))
    '        行政區 = Trim(table1.Rows(0)("行政區"))
    '        單價 = CInt(Trim(table1.Rows(0)("單價")))
    '        屋齡 = Trim(table1.Rows(0)("屋齡"))
    '        面積 = Trim(table1.Rows(0)("面積"))
    '        物件類型 = Trim(table1.Rows(0)("物件類型"))
    '        總價 = Trim(table1.Rows(0)("總價")) * 10000
    '        經度 = Trim(table1.Rows(0)("經度"))
    '        緯度 = Trim(table1.Rows(0)("緯度"))
    '    End If
    '    Dim 建築完成日 As String = 屋齡
    '    If Request("oid") = "60063AAD21188" Then
    '        建築完成日 = "2016-03-30"
    '        屋齡 = "14"
    '    Else
    '        If 建築完成日 = "" Then
    '            建築完成日 = ""
    '            屋齡 = ""
    '        Else
    '            建築完成日 = CInt(Mid(建築完成日, 1, 3)) + 1911
    '            建築完成日 &= "-" & Mid(屋齡, 4, 2) & "-" & "00"
    '            屋齡 = CInt(Mid(屋齡, 1, 3)) + 1911
    '            屋齡 = CInt(Mid(sysdate, 1, 3)) + 1911 - CInt(屋齡)
    '        End If
    '    End If

    '    If 物件類型 = "1" Or 物件類型 = "2" Or 物件類型 = "3" Then
    '        If 經度 = "" Or 緯度 = "" Or 經度 = 0 Or 緯度 = 0 Then
    '            eip_usual.Show("地圖未定位，故無法查詢")
    '            Exit Sub
    '        End If
    '        Dim str_url As String
    '        str_url = "https://map.ctop.tw/ceoCompetitive/index.aspx?account_id=89446259&compay_id=89446259&key=az04OTQ0NjI1OU9E&areatype=" & 縣市 & "&Town=" & 行政區 & "&entrustAPrice=" & 單價 & "&houseAge=" & 屋齡 & "&areaSize=" & 面積 & "&type=" & 物件類型 & "&entrustTPrice=" & 總價 & "&lat=" & 經度 & "&lng=" & 緯度 & "&objectType=" & 型態 & "&BuileYM=" & 建築完成日 & "&TFloor=" & 屋齡 & ""
    '        Response.Write("<script>window.open('" + str_url + "','');</script>")
    '    Else
    '        eip_usual.Show("物件類別不為 大樓 華廈 公寓 透天，故無法查詢")
    '        Exit Sub
    '    End If
    '    'End If
    'End Sub

    Protected Sub RadioButtonList3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList3.SelectedIndexChanged
        'If RadioButtonList3.SelectedIndex = 0 Then
        '    UpdatePanel21.visible = True
        'Else
        '    UpdatePanel21.visible = False
        'End If
    End Sub

    Sub Select_分區細項New(ByRef cls As String)
        Dim 大項名稱 As String

        If cls = "" Then
            DropDownList65.SelectedValue = "請選擇"
            Exit Sub
        End If

        Dim conn_使用分區 As New SqlConnection(EGOUPLOADSqlConnStr)
        conn_使用分區.Open()

        sql = "select * from 資料_使用分區 With(NoLock) where 使用分區小項='" & cls & "' "

        adpt = New SqlDataAdapter(sql, conn_使用分區)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count > 0 Then
            DropDownList66.Items.Clear()
            DropDownList66.Items.Add("請選擇")
            For i = 0 To table1.Rows.Count - 1
                大項名稱 = table1.Rows(i)("使用分區大項").ToString.Trim
                DropDownList66.Items.Add(table1.Rows(i)("使用分區小項").ToString.Trim)
                DropDownList66.Items(i + 1).Value = table1.Rows(i)("使用分區小項").ToString.Trim
            Next
            DropDownList65.SelectedValue = 大項名稱
            DropDownList65_SelectedIndexChanged(Nothing, Nothing)
            DropDownList66.SelectedValue = cls
        Else
            For i = 0 To DropDownList65.Items.Count - 1
                If DropDownList65.Items(i).Value = cls Then
                    DropDownList65.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        conn_使用分區.Close()
        conn_使用分區.Dispose()
    End Sub

    Public Sub 更新委託期間()
        Dim 物件編號 As String = ""
        Dim sid As String = Request("sid")      '店代號
        Dim oid As String = Request("oid")      '物件編號
        Dim 起始申請日 As String = ""
        Dim 起始日期 As String = ""
        Dim 終止申請日 As String = ""
        Dim 終止日期 As String = ""
        Dim 原始起始日期 As String = ""
        Dim 原始終止日期 As String = ""

        物件編號 = oid.Trim
        'conn = New SqlConnection(EGOUPLOADSqlConnStr)
        'conn.Open()
        '撈取原始資料
        sql = " select TOP 1 續約前日期起,續約前日期訖 from 委賣續約資料表 "
        sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
        sql += " order by 新增日期,Num "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1_3")
        table1_3 = ds.Tables("table1_3")

        '撈取異動申請檔
        sql = " select item as 識別,id,add_date as 申請日,oid as 物件編號,sid as 店代號,contents as 日期 "
        sql += " from object_audit "
        sql += " where audit='Y' and oid like '" & Left(oid, 13) & "%' "
        sql += " order by add_date desc,id desc "
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")
        '撈取異動申請明細檔
        sql = " select * from object_audit_detail  "
        sql += " where oid='" & oid & "' and sid='" & sid & "' "
        sql += " order by num desc "
        'eip_usual.Show(sql)
        'Exit Sub
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1_1")
        table1_1 = ds.Tables("table1_1")
        '撈取續約檔
        sql = " select TOP 1 新增日期,續約日期訖 from 委賣續約資料表 "
        sql += " where 刪除註記='N' and 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
        sql += " order by 新增日期 desc,Num desc "
        'eip_usual.Show(sql)
        'Exit Sub
        adpt = New SqlDataAdapter(sql, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table1_2")
        table1_2 = ds.Tables("table1_2")
        'conn.Close()
        'conn.Dispose()
        '先判別委託起始日
        If table1.Rows.Count > 0 Then
            For i = 0 To table1.Rows.Count - 1
                If table1.Rows(i)("識別").ToString.Trim = "委託起始日" Then
                    '再去判別物件編號及店代號是否一樣
                    If table1.Rows(i)("物件編號").ToString.Trim = oid And table1.Rows(i)("店代號").ToString.Trim = sid Then
                        '一樣直接紀錄，跳離迴圈
                        起始申請日 = table1.Rows(i)("申請日").ToString.Trim
                        起始日期 = table1.Rows(i)("日期").ToString.Trim
                        Exit For
                    Else
                        '不一樣，判別明細資料
                        If table1_1.Rows.Count > 0 Then
                            For j As Integer = 0 To table1_1.Rows.Count - 1
                                If table1_1.Rows(j)("num").ToString.Trim = table1.Rows(i)("id").ToString.Trim And table1_1.Rows(j)("oid").ToString.Trim = oid And table1_1.Rows(j)("sid").ToString.Trim = sid Then
                                    起始申請日 = table1.Rows(i)("申請日").ToString.Trim
                                    起始日期 = table1.Rows(i)("日期").ToString.Trim
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If 起始申請日 <> "" Then
                    Exit For
                End If
            Next
        End If

        If table1.Rows.Count > 0 Then
            For i = 0 To table1.Rows.Count - 1
                If table1.Rows(i)("識別").ToString.Trim = "委託截止日" Then
                    '再去判別物件編號及店代號是否一樣
                    If table1.Rows(i)("物件編號").ToString.Trim = oid And table1.Rows(i)("店代號").ToString.Trim = sid Then
                        '一樣直接紀錄，跳離迴圈
                        終止申請日 = table1.Rows(i)("申請日").ToString.Trim
                        終止日期 = table1.Rows(i)("日期").ToString.Trim
                        Exit For
                    Else
                        '不一樣，判別明細資料
                        If table1_1.Rows.Count > 0 Then
                            For j As Integer = 0 To table1_1.Rows.Count - 1
                                If table1_1.Rows(j)("num").ToString.Trim = table1.Rows(i)("id").ToString.Trim And table1_1.Rows(j)("oid").ToString.Trim = oid And table1_1.Rows(j)("sid").ToString.Trim = sid Then
                                    終止申請日 = table1.Rows(i)("申請日").ToString.Trim
                                    終止日期 = table1.Rows(i)("日期").ToString.Trim
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If 終止申請日 <> "" Then
                    Exit For
                End If
            Next
        End If
        'eip_usual.Show("委託起始日：申請日 " & 起始申請日 & " 更改日期 " & 起始日期 & " 委託截止日：申請日 " & 終止申請日 & " 更改日期 " & 終止日期)
        'Exit Sub

        If table1_2.Rows.Count > 0 Then
            For i = 0 To table1_2.Rows.Count - 1
                If table1_2.Rows(i)("新增日期").ToString.Trim >= 終止申請日 Then
                    終止申請日 = table1_2.Rows(i)("新增日期").ToString.Trim
                    終止日期 = table1_2.Rows(i)("續約日期訖").ToString.Trim
                    Exit For
                End If
            Next
        End If
        If table1_3.Rows.Count > 0 Then
            For i = 0 To table1_3.Rows.Count - 1
                原始起始日期 = table1_3.Rows(i)("續約前日期起").ToString.Trim
                原始終止日期 = table1_3.Rows(i)("續約前日期訖").ToString.Trim
            Next
        End If

        'If oid = "60863AAD60460" Then
        '    Response.Write("終止申請日" & 終止申請日)
        'End If
        'eip_usual.Show(" 委託起始日：申請日 " & 起始申請日 & " 更改日期 " & 起始日期 & " 委託截止日：申請日 " & 終止申請日 & " 更改日期 " & 終止日期)
        'conn = New SqlConnection(EGOUPLOADSqlConnStr)
        'conn.Open()
        If 起始申請日 <> "" Then
            sql = " update 委賣物件資料表 "
            sql += " set 委託起始日='" & 起始日期 & "' "
            sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
        Else
            If 原始起始日期 <> "" Then
                sql = " update 委賣物件資料表 "
                sql += " set 委託起始日='" & 原始起始日期 & "' "
                sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If
        End If
        If 終止申請日 <> "" Then
            sql = " update 委賣物件資料表 "
            sql += " set 委託截止日='" & 終止日期 & "' "
            sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
            cmd = New SqlCommand(sql, conn)
            cmd.ExecuteNonQuery()
        Else
            If 原始起始日期 <> "" Then
                sql = " update 委賣物件資料表 "
                sql += " set 委託截止日='" & 原始終止日期 & "' "
                sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
                cmd = New SqlCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End If
        End If
        'conn.Close()
        'conn.Dispose()
    End Sub

    Sub 寫入坪數()
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        sql = " update " & src.Text & " "
        sql += " set  "
        sql &= "主建物 ="
        If TextBox6.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox6.Text & " , "
        End If
        sql &= "主建物平方公尺 ="
        If TextBox5.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox5.Text & " , "
        End If

        sql &= "附屬建物 ="
        If TextBox8.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox8.Text & " , "
        End If
        sql &= "附屬建物平方公尺 ="
        If TextBox7.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox7.Text & " , "
        End If

        sql &= "公共設施 ="
        If TextBox10.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox10.Text & " , "
        End If
        sql &= "公共設施平方公尺 ="
        If TextBox9.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox9.Text & " , "
        End If

        sql &= "公設內車位坪數 ="
        If TextBox23.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox23.Text & " , "
        End If
        sql &= "公設內車位平方公尺 ="
        If TextBox21.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox21.Text & " , "
        End If

        sql &= "地下室 ="
        If TextBox20.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox20.Text & " , "
        End If
        sql &= "地下室平方公尺 ="
        If TextBox19.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox19.Text & " , "
        End If

        sql &= "總坪數 ="
        If TextBox29.Text = "" Then
            sql &= "0,"
        Else
            sql &= TextBox29.Text & " , "
        End If
        sql &= "總平方公尺 ="
        If TextBox28.Text = "" Then
            sql &= "0,"
        Else
            sql &= TextBox28.Text & " , "
        End If

        sql &= "土地坪數 ="
        If TextBox31.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox31.Text & " , "
        End If
        sql &= "土地平方公尺 ="
        If TextBox30.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox30.Text & " , "
        End If

        sql &= " 車位坪數 = "
        If TextBox27.Text = "" Then
            sql &= "null,"
        Else
            sql &= TextBox27.Text & " , "
        End If
        sql &= " 車位平方公尺 = "
        If TextBox26.Text = "" Then
            sql &= " null, "
        Else
            sql &= TextBox26.Text & " , "
        End If

        sql &= " 庭院坪數 = "
        If TextBox33.Text = "" Then
            sql &= " null, "
        Else
            sql &= TextBox33.Text & " , "
        End If
        sql &= " 庭院平方公尺 = "
        If TextBox32.Text = "" Then
            sql &= " null, "
        Else
            sql &= TextBox32.Text & " , "
        End If
        sql &= " 加蓋坪數 = "
        If TextBox35.Text = "" Then
            sql &= " null, "
        Else
            sql &= TextBox35.Text & " , "
        End If
        sql &= "加蓋平方公尺 ="
        If TextBox34.Text = "" Then
            sql &= " null "
        Else
            sql &= TextBox34.Text & " "
        End If
        sql &= " Where 物件編號 = '" & Request("oid") & "' and 店代號 = '" & Request("sid") & "' "
        cmd = New SqlCommand(sql, conn)
        'Response.Write(sql)
        Try
            cmd.ExecuteNonQuery()
        Catch ex As SqlException
            'Response.Write(sql)
        End Try
        conn.Close()
        conn.Dispose()
    End Sub

    'Sub 更新使用分區()
    '    conn = New SqlConnection(EGOUPLOADSqlConnStr)
    '    conn.Open()

    '    sql = " select TOP 1 續約前日期起,續約前日期訖 from 委賣續約資料表 "
    '    sql += " where 物件編號 = '" & oid & "' and 店代號='" & sid & "' "
    '    sql += " order by 新增日期,Num "
    '    adpt = New SqlDataAdapter(sql, conn)
    '    ds = New DataSet()
    '    adpt.Fill(ds, "table1_3")
    '    table1_3 = ds.Tables("table1_3")


    '    Try
    '        cmd.ExecuteNonQuery()
    '    Catch ex As SqlException
    '        'Response.Write(sql)
    '    End Try
    '    conn.Close()
    '    conn.Dispose()

    'End Sub

    Protected Sub RadioButton4_CheckedChanged(sender As Object, e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton3.Checked = True Then
            CheckBox102.Checked = False
            CheckBox103.Checked = False
            CheckBox102.Visible = False
            CheckBox103.Visible = False
        Else
            CheckBox102.Checked = False
            CheckBox103.Checked = False
            CheckBox102.Visible = True
            CheckBox103.Visible = True
        End If
    End Sub

    Protected Sub RadioButton3_CheckedChanged(sender As Object, e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            CheckBox102.Checked = False
            CheckBox103.Checked = False
            CheckBox102.Visible = False
            CheckBox103.Visible = False
        Else
            CheckBox102.Checked = False
            CheckBox103.Checked = False
            CheckBox102.Visible = True
            CheckBox103.Visible = True
        End If
    End Sub


    Protected Sub ImageButton11_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton11.Click
        Dim nscript As String
        Dim href As String = ""
        Dim ntitle As String = ""
        Dim 物件編號 As String = Request("oid")

        href = "../TOP_tools/tool_landcount.aspx?state=update&sid=" & Request("sid") & "&oid=" & 物件編號 & "&src=" & Request("src")

        ntitle = "土增稅計算"

        nscript = "window.open('"
        nscript += href
        nscript += " '"
        nscript += ",'newwindow2', 'height=825, width=1100, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');"

        Page.ClientScript.RegisterStartupScript(Me.GetType(), "MyScript", nscript, True)
    End Sub

    Public Sub 更新車位()
        Dim 車位 As String = ""
        Dim 車位說明 As String = ""
        Dim 車位價格 As String = ""
        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            sql = " select b.車位類別 "
            sql += " From 委賣物件資料表_面積細項 a With(NoLock) "
            sql += " left join 產調_車位 b With(NoLock) on a.店代號=b.店代號 and a.物件編號=b.物件編號 and A.流水號=B.流水號 "
            sql += " where a.店代號='" & Request.QueryString("sid") & "' and a.物件編號='" & Request.QueryString("oid") & "' "
            sql += " and b.物件編號 is not null and isnull(b.車位類別,'')<>'' "
            sql += " group by b.車位類別 "
            Using cmd_cartype As New SqlCommand(sql, conn)
                Dim dt1 As New DataTable
                dt1.Load(cmd_cartype.ExecuteReader)
                For k As Integer = 0 To dt1.Rows.Count - 1
                    車位 &= dt1.Rows(k)("車位類別").ToString.Trim & ","
                Next
            End Using

            sql = " select b.車位說明 "
            sql += " From 委賣物件資料表_面積細項 a With(NoLock) "
            sql += " left join 產調_車位 b With(NoLock) on a.店代號=b.店代號 and a.物件編號=b.物件編號 and A.流水號=B.流水號 "
            sql += " where a.店代號='" & Request.QueryString("sid") & "' and a.物件編號='" & Request.QueryString("oid") & "' "
            sql += " and b.物件編號 is not null and isnull(b.車位說明,'')<>'' "
            sql += " group by b.車位說明 "
            Using cmd_cartype As New SqlCommand(sql, conn)
                Dim dt1 As New DataTable
                dt1.Load(cmd_cartype.ExecuteReader)
                For k As Integer = 0 To dt1.Rows.Count - 1
                    車位說明 &= dt1.Rows(k)("車位說明").ToString.Trim & "、"
                Next
            End Using

            sql = " select b.車位獨立售價 "
            sql += " From 委賣物件資料表_面積細項 a With(NoLock) "
            sql += " left join 產調_車位 b With(NoLock) on a.店代號=b.店代號 and a.物件編號=b.物件編號 and A.流水號=B.流水號 "
            sql += " where a.店代號='" & Request.QueryString("sid") & "' and a.物件編號='" & Request.QueryString("oid") & "' "
            sql += " and b.物件編號 is not null and isnull(b.車位獨立售價,'')<>'' "
            sql += " group by b.車位獨立售價 "
            Using cmd_cartype As New SqlCommand(sql, conn)
                Dim dt1 As New DataTable
                dt1.Load(cmd_cartype.ExecuteReader)
                For k As Integer = 0 To dt1.Rows.Count - 1
                    車位價格 &= dt1.Rows(k)("車位獨立售價").ToString.Trim & "、"
                Next
            End Using
        End Using
        If 車位價格 <> "" Then 車位價格 = Mid(車位價格, 1, Len(車位價格) - 1)

        If 車位 = "" Then
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                sql = " Select * From 委賣物件資料表_面積細項 "
                sql += " Where 店代號='" & Request.QueryString("sid") & "' and 物件編號='" & Request.QueryString("oid") & "' "
                sql += " And (類別 Like '%車位%' "
                sql += " or 項目名稱 Like '%車位%' "
                sql += " or isnull(是否為車位,'') ='Y') "
                Using cmd_cartype As New SqlCommand(sql, conn)
                    Dim dt1 As New DataTable
                    dt1.Load(cmd_cartype.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        車位 &= "有"
                    End If
                End Using
            End Using
        End If

        If 車位 = "" Then
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                sql = " Select * From 委賣物件資料表 "
                sql += " Where 店代號='" & Request.QueryString("sid") & "' and 物件編號='" & Request.QueryString("oid") & "' "
                sql += " And 含公設車位坪數='Y' "
                Using cmd_cartype As New SqlCommand(sql, conn)
                    Dim dt1 As New DataTable
                    dt1.Load(cmd_cartype.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        車位 &= "有"
                    End If
                End Using
            End Using
        End If

        Dim updatestr As String = ""
        updatestr = "update " & src.Text & " set "
        updatestr += " 車位 = '" & 車位 & "' "
        updatestr += " ,車位說明 = '" & 車位說明 & "' "
        updatestr += " ,車位價格報表 = '" & 車位價格 & "' "
        updatestr += " where 物件編號 = '" & Request.QueryString("oid") & "' and 店代號 = '" & Request.QueryString("sid") & "' "
        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            conn.Open()
            Using cmd As New SqlCommand(updatestr, conn)
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    'Response.Write(updatestr)
                    'Response.End()
                End Try
            End Using
        End Using
    End Sub

    Sub 更新圖片相關資訊()
        Dim filenames As String
        Dim sid As String = Request("sid")
        Dim objectnum As String = Request("oid")

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        Dim 日期 As String = Format(Year(Today) - 1911, "0##") & Format(Today, "MMdd")
        '產生 checkpic、委租物件相片
        sql = "select count(*) from checkpic where pic_contract_no='" & objectnum & "' and pic_dept_no='" & sid & "'"
        Using cmd3 As New SqlCommand(sql, conn)
            If cmd3.ExecuteScalar = 0 Then
                sql = "Insert Into checkpic (pic_contract_no,pic_dept_no,pic_upd_dt,pic_1,pic_2,pic_3,pic_4,pic_5,pic_6,pic_7,pic_8,pic_9,pic_10,pic_11,pic_12,pic_13,pic_14,pic_15,pic_plan,pic_count)"
                sql += " Values('" & objectnum & "','" & sid & "','" & 日期 & "','N','N','N','N','N','N','N','N','N','N','N','N','N','N','N','N',0)"
            End If
            cmd2 = New SqlCommand(sql, conn)
            cmd2.ExecuteNonQuery()
        End Using
        sql = "select count(*) from 委賣物件相片 where 物件編號='" & objectnum & "' and 店代號='" & sid & "'"
        Using cmd3 As New SqlCommand(sql, conn)
            If cmd3.ExecuteScalar = 0 Then
                sql = " Insert Into 委賣物件相片 (物件編號,店代號,新增日期,修改日期 "
                sql += " ,照片1,照片2,照片3,照片4,照片5,照片6,照片7,照片8,照片9,照片10,照片11,照片12,照片13,照片14,照片15 "
                sql += " ,照片說明1,照片說明2,照片說明3,照片說明4,照片說明5,照片說明6,照片說明7,照片說明8,照片說明9,照片說明10,照片說明11,照片說明12,照片說明13,照片說明14,照片說明15 "
                sql += " ,輸入者,上傳註記) "
                sql += " Values('" & objectnum & "','" & sid & "','" & 日期 & "','" & 日期 & "' "
                sql += " ,'','','','','','','','','','','','','','','', "
                sql += " '','','','','','','','','','','','','','','', '" & Request.Cookies("webfly_empno").Value & "','v')"
            End If
            cmd2 = New SqlCommand(sql, conn)
            cmd2.ExecuteNonQuery()
        End Using

        '更新
        Dim wc As New System.Net.WebClient
        Dim 照片 As String = ""
        Dim pic(17) As String
        Dim 數量 As Integer = 0
        Dim 圖片 As String() = Split("a,b,c,d,w,x,y,z,e,f,h,s,t,u,v,g,m", ",")
        For i = 0 To 16
            filenames = "https://img.etwarm.com.tw/" & sid & "/available/" & objectnum & 圖片(i) & ".jpg"
            Try
                Dim data As Byte() = wc.DownloadData(filenames)
                pic(i + 1) = "'Y'"
                If i >= 0 And i <= 14 Then
                    數量 += 1
                End If
            Catch ex As System.Net.WebException
                pic(i + 1) = "'N'"
            End Try
        Next
        sql = "Update checkpic set pic_1=" & pic(1) & ",pic_2=" & pic(2) & ",pic_3=" & pic(3) & ",pic_4=" & pic(4) & ",pic_5=" & pic(5) & ",pic_6=" & pic(6) & ",pic_7=" & pic(7) & ",pic_8=" & pic(8) & ",pic_9=" & pic(9) & ",pic_10=" & pic(10) & ",pic_11=" & pic(11) & ",pic_12=" & pic(12) & ",pic_13=" & pic(13) & ",pic_14=" & pic(14) & ",pic_15=" & pic(15) & ",pic_plan=" & pic(16) & ",pic_Count=" & 數量
        sql += " where pic_contract_no='" & objectnum & "' and pic_dept_no='" & sid & "'"
        'Response.Write(sql)
        'Exit Sub
        cmd2 = New SqlCommand(sql, conn)
        cmd2.ExecuteNonQuery()
        sql = "Update a "
        sql += " set 照片1=" & pic(1) & ",照片2=" & pic(2) & ",照片3=" & pic(3) & ",照片4=" & pic(4) & ",照片5=" & pic(5) & ",照片6=" & pic(6) & ",照片7=" & pic(7) & ",照片8=" & pic(8)
        sql += " ,照片說明1=(case when " & pic(1) & "='' then '' else a.照片說明1 end) "
        sql += " ,照片說明2=(case when " & pic(2) & "='' then '' else a.照片說明2 end) "
        sql += " ,照片說明3=(case when " & pic(3) & "='' then '' else a.照片說明3 end) "
        sql += " ,照片說明4=(case when " & pic(4) & "='' then '' else a.照片說明4 end) "
        sql += " ,照片說明5=(case when " & pic(5) & "='' then '' else a.照片說明5 end) "
        sql += " ,照片說明6=(case when " & pic(6) & "='' then '' else a.照片說明6 end) "
        sql += " ,照片說明7=(case when " & pic(7) & "='' then '' else a.照片說明7 end) "
        sql += " ,照片說明8=(case when " & pic(8) & "='' then '' else a.照片說明8 end) "
        sql += " ,照片說明9=(case when " & pic(9) & "='' then '' else a.照片說明9 end) "
        sql += " ,照片說明10=(case when " & pic(10) & "='' then '' else a.照片說明10 end) "
        sql += " ,照片說明11=(case when " & pic(11) & "='' then '' else a.照片說明11 end) "
        sql += " ,照片說明12=(case when " & pic(12) & "='' then '' else a.照片說明12 end) "
        sql += " ,照片說明13=(case when " & pic(13) & "='' then '' else a.照片說明13 end) "
        sql += " ,照片說明14=(case when " & pic(14) & "='' then '' else a.照片說明14 end) "
        sql += " ,照片說明15=(case when " & pic(15) & "='' then '' else a.照片說明15 end) "
        sql += " ,輸入者='" & Request.Cookies("webfly_empno").Value & "'"
        sql += " from 委賣物件相片 a "
        sql += " where 物件編號='" & objectnum & "' and 店代號='" & sid & "' "
        cmd2 = New SqlCommand(sql, conn)
        cmd2.ExecuteNonQuery()

        sql = " update 委賣物件資料表 set num=" & 數量 & ",plan1=" & pic(16) & ",map=" & pic(17) & " "
        sql += " where 物件編號='" & objectnum & "' and 店代號='" & sid & "' "
        cmd2 = New SqlCommand(sql, conn)
        cmd2.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()
    End Sub

    Protected Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        checkleave.使用分析(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "Z", "AI生成")

        Dim message As String
        Dim script As String = ""
        Dim 經度 As String = ""
        Dim 緯度 As String = ""

        Dim conn_定位 As New SqlConnection(EGOUPLOADSqlConnStr)
        conn_定位.Open()

        sql = " select isnull(經度,0) as 經度,isnull(緯度,0) as 緯度 "
        If Trim(Request("src")) = "OLD" Then
            sql += " from 委賣物件過期資料表 With(NoLock) "
        ElseIf Trim(Request("src")) = "NOW" Then
            sql += " from 委賣物件資料表 With(NoLock)"
        End If
        sql += " where 物件編號='" & Request("oid") & "' "
        sql += " and 店代號='" & Request("sid") & "' "
        sql += " and 經度 is not null and isnull(經度,0)<>0 "

        adpt = New SqlDataAdapter(sql, conn_定位)
        ds = New DataSet()
        adpt.Fill(ds, "table1")
        table1 = ds.Tables("table1")

        If table1.Rows.Count > 0 Then
            If table1.Rows(0)("經度").ToString.Trim = "0" Then
                conn_定位.Close()
                conn_定位.Dispose()
                message = "請先定位此物件"
                If message <> "" Then
                    script += "alert('" & message & "');"
                    ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                    Exit Sub
                End If
            Else
                經度 = table1.Rows(0)("經度").ToString.Trim
                緯度 = table1.Rows(0)("緯度").ToString.Trim
            End If
        Else
            conn_定位.Close()
            conn_定位.Dispose()
            message = "請先定位此物件"
            If message <> "" Then
                script += "alert('" & message & "');"
                ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
                Exit Sub
            End If
        End If

        conn_定位.Close()
        conn_定位.Dispose()


        Dim address1 As String = "", address2 As String = ""

        '部分地址
        If add1.Text <> "" Then address2 &= add1.Text & zone3.SelectedValue
        If add2.Text <> "" Then address2 &= add2.Text & "鄰"
        If add3.Text <> "" Then address2 &= add3.Text & address20.SelectedValue
        '1040519修正
        'If add4.Text <> "" Then address2 &= add4.Text & "段"
        If add4.Text <> "" Then address2 &= add4.Text & Label64.Text

        'If add5.Text <> "" Then address2 &= add5.Text & "巷"
        'If add6.Text <> "" Then address2 &= add6.Text & "弄"

        'address1 &= address2
        ''1040519修正
        ''If add7.Text <> "" Then address1 &= add7.Text & "號"
        'If Label64.Text = "小段" Then
        '    If add7.Text <> "" Then address1 &= add7.Text & "地號"
        'Else
        '    If add7.Text <> "" Then address1 &= add7.Text & "號"
        'End If
        'If add8.Text <> "" Then address1 &= "之" & add8.Text

        ''20100607小豪修正("之"後方若有值,則"樓"之前加入空白,避免距離過近造成誤會,ex:101號之1   3樓)----
        'If add9.Text <> "" Then
        '    If add8.Text <> "" Then
        '        address1 &= "   " & add9.Text & "樓"
        '    Else
        '        address1 &= add9.Text & "樓"
        '    End If
        'End If
        ''--------------------------------------------------------------------------------------------------
        'If add10.Text <> "" Then address1 &= "之" & add10.Text


        Dim webClient = New WebClient()
        ' 指定 WebClient 編碼
        webClient.Encoding = Encoding.UTF8
        ' 指定 WebClient 的 authorization header
        webClient.Headers.Add("authorization", "token {apitoken}")
        '要傳送的資料內容
        Dim nameValues As NameValueCollection = New NameValueCollection()
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

        Dim 網址 As String = ""
        '網址 = "https://bot.etwarm.net.tw/webhook_797dhxrq/AI/objTitle.php?old=" & TextBox2.Text & "&address=" & DDL_County.SelectedValue & DDL_Area.SelectedItem.Text & address2 & "&kind=" & DropDownList3.SelectedValue & "&lat=" & 經度 & "&lng=" & 緯度

        網址 = "https://bot.etwarm.net.tw/webhook_797dhxrq/AI/objTitle.php?oid=" & TextBox2.Text & "&address=" & DDL_County.SelectedValue & DDL_Area.SelectedItem.Text & address2 & "&kind=" & DropDownList3.SelectedValue & "&lat=" & 經度 & "&lng=" & 緯度


        Try
            Dim result = webClient.UploadValues(網址, nameValues)
            '將 post 結果轉為 string
            Dim resultstr As String = Encoding.UTF8.GetString(result)
            'Label10.Text = resultstr

            Dim Obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(resultstr)
            '取得物件明細的方法 依據階層 Obj.item("階層1")("階層2") 以此類推
            Dim 物件特色 As String = Obj.Item("msg").ToString
            TextBox102.Text = 物件特色
        Catch ex As Exception
            TextBox102.Text = ""
        End Try

        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "unblockUI", "$.unblockUI();", True)
    End Sub

    Protected Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        checkleave.使用分析(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "Z", "AI文案載入")

        Dim webClient = New WebClient()
        ' 指定 WebClient 編碼
        webClient.Encoding = Encoding.UTF8
        ' 指定 WebClient 的 authorization header
        webClient.Headers.Add("authorization", "token {apitoken}")
        '要傳送的資料內容
        Dim nameValues As NameValueCollection = New NameValueCollection()
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

        Dim 網址 As String = ""
        網址 = "https://bot.etwarm.net.tw/webhook_797dhxrq/AI/loadAI.php?sid=" & store.SelectedValue & "&no=" & TextBox2.Text

        Try
            Dim result = webClient.UploadValues(網址, nameValues)
            '將 post 結果轉為 string
            Dim resultstr As String = Encoding.UTF8.GetString(result)
            'Label10.Text = resultstr

            Dim Obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(resultstr)
            '取得物件明細的方法 依據階層 Obj.item("階層1")("階層2") 以此類推
            Dim 物件特色 As String = Obj.Item("msg").ToString
            TextBox102.Text = 物件特色
        Catch ex As Exception
            TextBox102.Text = ""
        End Try

        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "unblockUI", "$.unblockUI();", True)
    End Sub

    Protected Sub UpdateAIVoiceOver(objectNo As String, storeNo As String)

        Dim indexNum = GetIndexNumByObjectNo(objectNo, storeNo)
        If String.IsNullOrEmpty(indexNum) Then
            Exit Sub
        End If
        Dim url = "https://queue.etwarm.com.tw/webhook/d1715378-639f-4cef-b039-961f19bea551" ' 替換為你的 URL

        Dim request = CType(WebRequest.Create(url), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"

        Dim postData As String = String.Format("objectNo={0}&objectType={1}", indexNum, DropDownList3.SelectedValue)
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)

        Try
            request.ContentLength = byteArray.Length
            Using dataStream As Stream = request.GetRequestStream()
                dataStream.Write(byteArray, 0, byteArray.Length)
            End Using

            Dim response = CType(request.GetResponse(), HttpWebResponse)
            Using responseStream As Stream = response.GetResponseStream()
                Dim reader As New StreamReader(responseStream, Encoding.UTF8)
                Dim responseFromServer As String = reader.ReadToEnd()

            End Using
            response.Close()
        Catch ex As WebException
            Console.WriteLine("Error: " & ex.Message)
            If ex.Response IsNot Nothing Then
                Using responseStream As Stream = ex.Response.GetResponseStream()
                    Dim reader As New StreamReader(responseStream, Encoding.UTF8)
                    Console.WriteLine(reader.ReadToEnd())
                End Using
            End If
        End Try
    End Sub

    Public Function GetIndexNumByObjectNo(objectNo As String, storeNo As String) As String

        If (String.IsNullOrEmpty(objectNo)) Then
            Return String.Empty
        End If
        Dim cnString As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString

        Const query = "sp_get_object_index_num_by_object_no"
        Dim result = String.Empty

        Using cn As New SqlConnection(cnString)
            Using cm As New SqlCommand(query, cn)
                cm.CommandType = CommandType.StoredProcedure
                cm.Parameters.AddWithValue("@object_no", objectNo)
                cm.Parameters.AddWithValue("@store_no", storeNo)
                Dim returnParameter = cm.Parameters.Add("@result", SqlDbType.BigInt)
                returnParameter.Direction = ParameterDirection.Output
                Try
                    cn.Open()
                    cm.ExecuteNonQuery()
                    result = Convert.ToString(returnParameter.Value)
                Catch ex As Exception
                    Dim message = ex.Message
                Finally
                    cn.Close()
                End Try
            End Using
        End Using

        Return result
    End Function


    Protected Sub LogTrack(ByVal message As String)
        Dim store_no As String = Request.Cookies("store_id").Value
        Dim emp_no As String = Request("webfly_empno")
        Dim log_title As String = "物件修改"
        Dim url As String = Request.Url.ToString
        Dim ip As String = Request.ServerVariables("REMOTE_ADDR")
        Dim query As String = ""
        錯誤訊息(store_no, emp_no, log_title, url, query, message)

    End Sub

    Protected Sub 錯誤訊息(ByVal 店代號 As String,
                  ByVal 員工編號 As String,
                  ByVal 功能名稱 As String,
                  ByVal 網址 As String,
                  ByVal 語法 As String,
                  ByVal 訊息 As String)

        Const sql As String =
            "INSERT INTO 房仲家錯誤訊息紀錄 " &
            "(店代號, 員工編號, 功能名稱, 網址, 語法, 訊息) " &
            "VALUES (@店代號, @員工編號, @功能名稱, @網址, @語法, @訊息);"

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr),
              cmd As New SqlCommand(sql, conn)

            '建議：明確指定型別 + 長度（長度請依你 DB 欄位實際長度調整）
            cmd.Parameters.Add("@店代號", SqlDbType.VarChar, 10).Value = If(店代號, "")
            cmd.Parameters.Add("@員工編號", SqlDbType.VarChar, 20).Value = If(員工編號, "")
            cmd.Parameters.Add("@功能名稱", SqlDbType.NVarChar, 100).Value = If(功能名稱, "")
            cmd.Parameters.Add("@網址", SqlDbType.NVarChar, 500).Value = If(網址, "")
            cmd.Parameters.Add("@語法", SqlDbType.NVarChar, -1).Value = If(語法, "")    '-1 = NVARCHAR(MAX)
            cmd.Parameters.Add("@訊息", SqlDbType.NVarChar, -1).Value = If(訊息, "")

            Try
                conn.Open()
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                '你原本是吃掉例外，這邊保留同樣行為（要記錄就自行加 log）
                '例如：Debug.WriteLine(ex.ToString())
            End Try
        End Using

    End Sub
End Class


Public Class object_info

    'obj_id
    Public Property obj_id As String


    '店家代碼（必要欄位）
    Public Property store_id As String


    '使用者代號（必要欄位）
    Public Property user_id As String

    '物件編號
    Public Property obj_no As String


    '物件名稱（必要欄位）
    Public Property obj_name As String

    '物件類型（必要欄位，參考 CaseType 代碼表）
    Public Property case_type As Integer

    '租售型態（0: 出租, 1: 出售，必要欄位）
    Public Property sell_rent As Integer


    '物件狀態（必要欄位，參考 CaseStatus 代碼表）
    Public Property case_status As Integer

    '坪數（查詢與配對）
    Public Property ping As Decimal

    '建物坪數（必要欄位）
    Public Property build_ping As Decimal

    '土地坪數（必要欄位）
    Public Property land_ping As Decimal

    '售價(元) / 月租金（必要欄位）
    Public Property price As Decimal

    '底價(元)
    Public Property base_price As Decimal

    '每坪單價(元)
    Public Property unit_price As Decimal

    '縣市名稱（必要欄位）
    Public Property city As String

    '區域名稱（必要欄位）
    Public Property area As String

    '郵遞區號（必要欄位）
    Public Property area_code As String

    '完整地址
    Public Property address As String

    '路段名稱（必要欄位）
    Public Property road_name As String

    '路寬（米）
    Public Property road_width As Decimal

    '座標經度
    Public Property gps_lng As Decimal

    '座標緯度
    Public Property gps_lat As Decimal

    '物件等級（A B C D E）
    Public Property obj_class As String

    '法定用途（參考 CaseLegalUse 代碼表）
    Public Property legal_use As Integer

    '現況用途（參考 CaseUse 代碼表）
    Public Property case_use As Integer

    '使用現況（參考 CaseCurrentUsage 代碼表）
    Public Property current_usage As Integer

    '委託型態（參考 CaseContractType 代碼表）
    Public Property contract_type As Integer

    '合約到期日
    Public Property contract_date_e As String

    '下架日期
    Public Property off_date As String

    '帶看注意事項
    Public Property look_note As String

    '鑰匙取得方式（參考 CaseKeyGet 代碼表）KEY
    Public Property key_get As Integer

    '格局-房（必要欄位）
    Public Property room As Integer

    '格局-廳（必要欄位）
    Public Property living As Integer

    '格局-衛（必要欄位）
    Public Property bathroom As Integer

    '格局-陽台
    Public Property balcony As Integer

    '所在樓層(起)（0 表示整棟，負值表示地下室，必要欄位）
    Public Property floor As Integer

    '所在樓層(迄)（負值表示地下室，必要欄位）
    Public Property floor_to As Integer

    '地上樓層（必要欄位）
    Public Property up_floors As Integer

    '地下樓層
    Public Property down_floors As Integer

    '朝向（參考 CaseToward 代碼表）
    Public Property toward As String

    '完工日期
    Public Property complete_date As String

    '屋齡
    Public Property age As Integer

    '建物結構（參考 CaseBuildStruct 代碼表）
    Public Property build_struct As Integer

    '外牆材質（參考 CaseWallMaterial 代碼表）
    Public Property wall_material As Integer

    '土地類型（0: 非都市土地, 1: 都市土地）
    Public Property land_type As Boolean

    '土地分區（參考 CaseUrbanZone 代碼表）
    Public Property land_urban_zone As Integer

    '土地深度（米）
    Public Property land_depth As Decimal

    '總戶數
    Public Property total_house As Integer

    '每層戶數
    Public Property house_per_floor As Integer

    '封面照片
    Public Property cover_photo As String

    '照片檔案清單
    Public Property photo_list As List(Of String)

    '附件檔案清單
    Public Property attachments As List(Of String)

    '時間戳記
    Public Property update_timestamp As Long

    Public Property major_space As Decimal
    Public Property sub_space As Decimal
    Public Property balcony_space As Decimal
    Public Property other_space As Decimal
    Public Property public_space As Decimal
    Public Property pub_parking_space As Decimal
    Public Property prv_parking_space As Decimal
    Public Property light_side As Integer
    Public Property elevators As Integer
    Public Property manage_type As Integer
    Public Property manage_fee As Integer
    Public Property manage_fee_way As Integer
    Public Property parking_type As Integer
    Public Property parking_no As String
    Public Property parking_rent As Integer
    Public Property parking_price As Integer
    Public Property community As String
    Public Property land_section As String
    Public Property land_no As String
    Public Property land_non_urban_zone As Integer
    Public Property land_non_urban_zone_type As Integer
    Public Property build_rate As Integer
    Public Property floor_rate As Integer
    Public Property land_width As Integer
    Public Property total_build_ping As Integer
    Public Property bus_station As String
    Public Property mrt_station As String
    Public Property park As String
    Public Property market As String
    Public Property hospital As String
    Public Property primary_school As String
    Public Property junior_school As String
    Public Property feature As String
    Public Property owner_name As String
    Public Property owner_id As String
    Public Property owner_birthday As String
    Public Property owner_title As Integer
    Public Property owner_address As String
    Public Property owner_home_tel As String
    Public Property owner_office_tel As String
    Public Property owner_mobile As String
    Public Property owner_email As String
    Public Property owner_remark As String

    Public Property agent_name As String
    Public Property agent_id As String
    Public Property agent_birthday As String
    Public Property agent_title As Integer
    Public Property agent_address As String
    Public Property agent_home_tel As String
    Public Property agent_office_tel As String
    Public Property agent_mobile As String
    Public Property agent_email As String
    Public Property agent_remark As String
End Class


'Public Class Payload
'    Public Property case_status As Integer
'    Public Property off_date As String
'End Class

Public Class EventItem
    Public Property action As Integer
    Public Property user_id As String
    Public Property obj_id As String
    Public Property payload As object_info
End Class

Public Class EventsWrapper
    Public Property events As List(Of EventItem)
End Class


