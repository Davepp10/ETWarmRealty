Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient

Partial Class usercontrol_main_menu
    Inherits UserControl
    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    Dim sql, sql1 As String
    Dim ds1, ds2 As DataSet
    Dim adpt1, adpt2 As SqlDataAdapter
    Dim checkvpn As New checkvpn
    Dim checkleave As New checkleave
    Dim myobj As New clspowerset
    Dim 法拍密碼 As String = ""
    Dim mysql As String = ConfigurationManager.ConnectionStrings("mysqletwarm").ConnectionString
    Public 直營 As String = "3"
    Public HouseHunterUrl As String
    'Public Show As String = "0"
    Dim ENCRYPT_IV As String = "etwarm88etwarm88"
    Dim ENCRYPT_KEY As String = "qwsdg4hh234dfhrw"
    Dim COOKIE_DOMAIN As String = ConfigurationManager.AppSettings("cookie_domain")
    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        'If (Request.Cookies("webfly_empno").Value = "CUD") Then
        '    Response.Write("aaa")
        '    Response.End()
        'End If

        If Request.Cookies("webfly_empno") Is Nothing And (Request.Cookies("store_id") Is Nothing) Then
            eip_usual.Show("登入錯誤，請重新登入!!")
            'Response.Write(" <script> setTimeout('location.href=""https://superwebnew.etwarm.com.tw/""',3000); </script> ")
            Response.Redirect("https://superwebnew.etwarm.com.tw/")
            'Else
            '    Response.Write(Request.Cookies("webfly_empno").Value)
        ElseIf (Request.Cookies("webfly_empno").Value = "") Then
            Response.Redirect("https://superwebnew.etwarm.com.tw/")
        End If

        Dim empNo = Request.Cookies("webfly_empno").Value
        Dim storeNo = Request.Cookies("store_id").Value
        Dim empValid = String.Format("{0}-{1}", empNo.Trim.ToUpper, storeNo)

        '如果還未加密過，則進行加密
        Dim encryptedCode As String
        If Request.Cookies("valid_code") Is Nothing Then
            encryptedCode = GetEncryptResultWithTimestamp(empValid, ENCRYPT_KEY, ENCRYPT_IV)
        Else
            encryptedCode = Request.Cookies("valid_code").Value
        End If

        Dim decryptedCode = GetDecryptResultWithTimestamp(encryptedCode, ENCRYPT_KEY, ENCRYPT_IV)
        If empValid <> decryptedCode Then
            eip_usual.Show("登入錯誤，請重新登入!!")
            'Response.Write(" <script> setTimeout('location.href=""https://superwebnew.etwarm.com.tw/""',3000); </script> ")
            Response.Redirect("/indexnew/login3.aspx")
        End If

        '更新 valid_code----------------
        empValid = GetEncryptResultWithTimestamp(empValid, ENCRYPT_KEY, ENCRYPT_IV)
        Dim cookieValidCode As HttpCookie
        If Request.Cookies("valid_code") Is Nothing Then
            cookieValidCode = New HttpCookie("valid_code")
        Else
            cookieValidCode = Request.Cookies("valid_code")
        End If

        Dim expires = Now.AddDays(2)
        cookieValidCode.Value = empValid
        cookieValidCode.Expires = expires
        cookieValidCode.Domain = COOKIE_DOMAIN

        Response.AppendCookie(cookieValidCode)
        Response.Cookies.Add(cookieValidCode)
        '-------------------------------

        'If Request.Cookies("store_id").Value.ToString = "A0001" Or Request.Cookies("store_id").Value.ToString = "A0002" Or Request.Cookies("store_id").Value.ToString = "A1322" Or Request.Cookies("store_id").Value.ToString = "A0279" Or Request.Cookies("store_id").Value.ToString = "A1294" Or Request.Cookies("store_id").Value.ToString = "A0718" Then
        '    'If (checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任") Then
        '    Show = "1"
        '    'End If
        'End If
        'Show = "1"

        myobj.power_object(Request.Cookies("webfly_empno").Value)
        myobj.mstores(Request.Cookies("webfly_empno").Value)
        myobj.mgroup(Request.Cookies("webfly_empno").Value)

        Using conn As New MySqlConnection(mysql)
            conn.Open()
            Dim linksel As String = "select * from foreclose where f_sid='" & Request.Cookies("store_id").Value & "' and del='N' "
            Using cmd As New MySqlCommand(linksel, conn)
                Dim dtmy As New DataTable
                dtmy.Load(cmd.ExecuteReader)
                If dtmy.Rows.Count > 0 Then
                    法拍密碼 = dtmy.Rows(0)("f_pwd")
                Else
                    法拍密碼 = ""
                End If
            End Using
        End Using

        Dim 法拍 As String = ""

        If 法拍密碼 <> "" Then
            法拍 = "http://www.54168.com.tw/AuthorizeCenter/Auction2010/AutoPass.asp?id=" & Request.Cookies("store_id").Value & "&amp;pd=" & 法拍密碼 & ""
        Else
            法拍 = ""
        End If

        Dim a As String
        Dim ans As String
        Dim str() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}

        Dim dt As New DataTable()
        Using conn1 As New SqlConnection(EGOUPLOADSqlConnStr)
            conn1.Open()

            sql1 = "select man_tree,man_tree_new from PSMAN With(NoLock) where man_emp_no='" & Request.Cookies("webfly_empno").Value & "' and  man_dept_no='" & Request.Cookies("store_id").Value & "' and  (man_quit_dt='' or man_quit_dt<='" & eip_usual.toROCyear(Now) & "')"
            'If Request.Cookies("webfly_empno").Value = "TWW" Then
            '    Response.Write(sql1)
            'End If

            adpt2 = New SqlDataAdapter(sql1, conn1)
            ds2 = New DataSet()
            adpt2.Fill(ds2, "table2")
            Dim table2 As DataTable = ds2.Tables("table2")
            If table2.Rows.Count > 0 Then
                Dim str1 As String = ""
                Dim str2 As String = ""
                If Not IsDBNull(table2.Rows(0)("man_tree_new")) Or table2.Rows(0)("man_tree_new").ToString.Trim <> "" Then
                    If checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東" Or Request.Cookies("webfly_empno").Value = "92H" Then
                        str1 = table2.Rows(0)("man_tree_new").ToString & ",319,349,344,224,266,510,214,213,88,217,238,227,327,343,330,310,336,341,212,500,502,514,239,253,274,237,504,246,335,107,103,73,78,70,77,71,93,262,354,337,230,323,322,503,360,207,101,100,130,279,297,283,354,360,361,362,500,501,502,503,504,505,315,237,239,247,275,226,321,515,516,517,519,521,522,523,524,525,"
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "業助" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務助理" Or checkleave.直營leave(Request.Cookies("webfly_empno").Value) = "秘書" Then
                        str1 = table2.Rows(0)("man_tree_new").ToString & ",283,354,360,361,362,500,501,502,503,504,505,315,237,239,247,275,226,321,515,516,517,519,522,523,524,"
                        'If Request.Cookies("webfly_empno").Value = "5WH" Then
                        '    Response.Write(str1)
                        'End If
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "店長" Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(一)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(二)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(三)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(四)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(五)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務副理(六)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(一)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(二)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(三)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(四)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(五)") Or (checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務經理(六)") Then
                        str1 = table2.Rows(0)("man_tree_new").ToString & ",283,361,362,501,315,237,247,275,226,321,516,517,519,521,523,524,"
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "經紀人" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "營業員" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "組長" Then
                        str1 = table2.Rows(0)("man_tree_new").ToString & ",283,361,362,501,315,237,247,275,226,321,516,517,519,523,524,"
                    Else
                        str1 = table2.Rows(0)("man_tree_new").ToString & ",283,361,362,501,315,237,247,275,226,321,516,517,519,523,524,"
                        'str1 = ",283,361,362,501,315,516,519,"
                    End If
                    '客戶權限
                    myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "158", "ALL")
                    If myobj.Brw = "1" Then
                        str1 &= "124,"
                    End If
                    If myobj.AC = "1" Then
                        str1 &= "158,1581,"
                    End If

                    '物件權限
                    myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "159", "ALL")
                    If myobj.Brw = "1" Then
                        str1 &= "2,"
                    End If
                    If myobj.AC = "1" Then
                        str1 &= "159,"
                    End If

                    '有成交就有同行合作成交
                    myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "348", "ALL")
                    If myobj.AC = "1" Then
                        str1 &= "175,"
                    End If
                    If checkvpn.checkvpn(Request.Cookies("store_id").Value.ToString) = "Y" And myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "357", "NO") = "1" Then
                        str1 &= "357,"
                    End If

                    If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                        str1 += "1597,"
                    End If

                    'If Request.Cookies("webfly_empno").Value = "AB3" Then
                    '    Response.Write(str1)
                    'End If

                Else
                    '= Replace(table2.Rows(0)("man_tree").ToString, ",", "','")
                    'response.write(Request.Cookies("webfly_empno").Value)
                    If checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東" Or Request.Cookies("webfly_empno").Value = "92H" Then
                        str1 = table2.Rows(0)("man_tree").ToString & ",319,349,344,224,266,510,214,213,88,217,238,227,327,343,330,310,336,341,212,500,502,514,239,253,274,237,504,246,335,107,103,73,78,70,77,71,93,262,354,337,230,323,322,503,360,207,101,100,130,279,297,283,354,360,361,362,500,501,502,503,504,505,315,237,239,247,275,226,321,515,516,517,519,521,522,523,524,525,"
                        str2 = "TOP1,TOP2,:A17,A18,A4,A5,A6,A7,A9,A12,A20,A21,A22,"
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "業助" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "業務助理" Or checkleave.直營leave(Request.Cookies("webfly_empno").Value) = "秘書" Then
                        str1 = table2.Rows(0)("man_tree").ToString & ",283,354,360,361,362,500,501,502,503,504,505,315,237,239,247,275,226,321,515,516,517,519,"
                        str2 = ":A17,A18,A4,A5,A6,A7,A9,A12,A20,A21,A22,"
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "店長" Then
                        str1 = table2.Rows(0)("man_tree").ToString & ",283,354,361,362,501,315,247,275,226,321,516,517,519,521,"
                        str2 = "TOP1,TOP2,:A17,A9,A20,A21,A22,"
                    ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "經紀人" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "營業員" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "組長" Then
                        str1 = table2.Rows(0)("man_tree").ToString & ",283,354,361,362,501,315,247,275,226,321,516,517,519,"
                        str2 = "TOP1,TOP2,:A17,A9,A20,A21,A22,"
                    Else
                        str1 = ",283,354,361,362,501,315,516,519,"
                        str2 = ":A18,A20,A21,A22,"
                    End If

                    sql = "update PSMAN set man_tree_new='" & str1 & "',man_power_new='" & str2 & "' where man_emp_no='" & Request.Cookies("webfly_empno").Value & "'"
                    'If Request.Cookies("webfly_empno").Value = "02K3" Or Request.Cookies("webfly_empno").Value = "02k3" Then
                    '    eip_usual.Show(sql)
                    Dim com As New SqlCommand(sql, conn1)
                    Try
                        com.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                    'End If
                End If
                str1 += "1587,1589,1590,1591,1593,"
                str1 = Replace(str1, ",", "','")

                For i = 0 To str.Length - 1
                    'If checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任" Then
                    '    sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and menu_type in ('A','B','C','D','E','F','G','H','I')"
                    '    sql &= " order by sort"
                    'Else
                    '    sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and levels ='0'"
                    '    sql &= " order by sort"
                    'End If
                    'If Request.Cookies("webfly_empno").Value = "92h" Or Request.Cookies("webfly_empno").Value = "92H" Then
                    If str(i) = "F" Then
                        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' "
                        Else
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id not in ('1595') "
                        End If
                    ElseIf str(i) = "J" Then
                        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' "
                        Else
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id not in ('1592') "
                        End If
                    Else
                        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id in ('" & str1 & "') "
                        Else
                            sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id in ('" & str1 & "') and id not in ('1591') "
                        End If
                        'If 法拍密碼 <> "" Then
                        '    sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id in ('" & str1 & "') "
                        'Else
                        '    sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id in ('" & str1 & "') and id not in ('266') "
                        'End If
                    End If

                    'sql = "SELECT id,menu_type,sort,menu_name as name,url,target,levels FROM web_powerset_menu_test With(NoLock) where affect='1' and menu_name<>'' and url<>'' and menu_type='" & str(i) & "' and id in ('" & str1 & "') "
                    If Left(Request.Cookies("store_id").Value.ToUpper, 1) = "K" Then
                        sql &= " and brand ='K'"
                    End If
                    'If Request.Cookies("webfly_empno").Value.ToUpper <> "92H" Then
                    '    sql &= " and menu_type in ('A','B','C','D','E','F','G','H')"
                    'End If
                    sql &= " order by sort"
                    'If Request.Cookies("webfly_empno").Value = "MRZ" Then
                    '    Response.Write(sql)
                    '    Response.End()
                    'End If
                    'Response.Write(sql)
                    'Response.End()
                    'End If

                    adpt1 = New SqlDataAdapter(sql, conn1)
                    ds1 = New DataSet()
                    adpt1.Fill(ds1, "table1")
                    Dim table1 As DataTable = ds1.Tables("table1")
                    If table1.Rows.Count > 0 Then
                        a += """" & str(i) & """:["
                        For j = 0 To table1.Rows.Count - 1
                            If table1.Rows(j)("id").ToString.Trim = "266" Then
                                '法拍暫時拿掉
                                'a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                'a += """,""url"":""" & 法拍 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                'a += """},"
                            ElseIf table1.Rows(j)("id").ToString.Trim = "2" And checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                                Dim 網址 As String = "../A_ObjectManage/Object_Search_ds.aspx"
                                a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                a += """,""url"":""" & 網址 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                a += """},"
                            ElseIf table1.Rows(j)("id").ToString.Trim = "124" And checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                                Dim 工作日誌 As String = "../B_CustomerManage/customer_search_v2.aspx"
                                a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                a += """,""url"":""" & 工作日誌 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                a += """},"
                            ElseIf table1.Rows(j)("id").ToString.Trim = "230" And checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                                'If Request.Cookies("webfly_empno").Value.ToUpper = "H2L" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CT8" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CPM" Or Request.Cookies("webfly_empno").Value.ToUpper = "05YZ" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CQ0" Or Request.Cookies("webfly_empno").Value.ToUpper = "0026" Or Request.Cookies("webfly_empno").Value.ToUpper = "0AG2" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CFY" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CVZ" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CQ1" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CUN" Or Request.Cookies("webfly_empno").Value.ToUpper = "M3C" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CVA" Then
                                Dim 工作日誌 As String = "../WORK_NOTE/DS_WorkNote_Calendar_New.aspx"
                                a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                a += """,""url"":""" & 工作日誌 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                a += """},"
                                'Else
                                '    Dim 工作日誌 As String = "../WORK_NOTE/DS_WorkNote_Calendar.aspx"
                                '    a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                '    a += """,""url"":""" & 工作日誌 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                '    a += """},"
                                'End If
                            ElseIf table1.Rows(j)("id").ToString.Trim = "212" And checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                                Dim 工作日誌 As String = "../WORK_NOTE/DS_WorkNote_assign_calendar_New.aspx"
                                a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                a += """,""url"":""" & 工作日誌 & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                a += """},"
                            Else
                                a += "{""name"":""" & table1.Rows(j)("menu_type").ToString.Trim + "-" + table1.Rows(j)("sort").ToString.Trim + "." + table1.Rows(j)("name").ToString.Trim
                                a += """,""url"":""" & table1.Rows(j)("url").ToString.Trim & """,""target"":""" & table1.Rows(j)("target").ToString.Trim
                                a += """},"
                            End If

                            '測試新版不動產說明書()
                            If table1.Rows(j)("url").ToString.Trim = "../A_ObjectManage/Obj_Add_V3.aspx?state=add&src=NOW" Then
                                'Select Case Trim(Request.Cookies("webfly_empno").Value)
                                '    Case "06TU", "P7B", "KJD", "03T0", "05KV", "06V3", "05QW", "NGR", "EUL", "01MM", "LC0", "NMG", "MPZ", "038Q" '土城體系.東門.諾北爾體系.嘉義仁愛.景美
                                a = Replace(a, "../A_ObjectManage/Obj_Add_V4.aspx?state=add&src=NOW", "../A_ObjectManage/Obj_Add_V4.aspx?state=add&src=NOW")
                                '    Case Else
                                'a = a
                                'End Select
                            End If
                        Next
                        'If Request.Cookies("webfly_empno").Value.Trim = "92H" And Left(Request.Cookies("store_id").Value.ToUpper, 1) = "A" And str(i) = "I" Then
                        '    a += "{""name"":""I-10.語音系統設定"",""url"":""../I_SystemSetting/voice_set.aspx"",""target"":""_self""},"
                        'End If
                        a += "],"
                    End If
                Next
            End If

            'Dim str1 As String = ""
            'If checkleave.leave(Request.Cookies("webfly_empno").Value) = "店東" Or Request.Cookies("webfly_empno").Value = "92H" Then
            '    str1 = table2.Rows(0)("man_tree").ToString & ",283,354,360,361,362,500,501,502,503,504,505,315,"
            'ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書主任" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "秘書" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "業助" Then
            '    str1 = table2.Rows(0)("man_tree").ToString & ",283,354,360,361,362,500,501,502,503,504,505,315,"
            'ElseIf checkleave.leave(Request.Cookies("webfly_empno").Value) = "經紀人" Or checkleave.leave(Request.Cookies("webfly_empno").Value) = "營業員" Then
            '    str1 = table2.Rows(0)("man_tree").ToString & ",283,354,361,362,501,504,315,"
            'Else

            ' End If
            Try
                ans = "{" & Mid(Replace(a, "},]", "}]"), 1, Len(Replace(a, "},]", "}]")) - 1) & "}"
            Catch ex As Exception

            End Try


            mainmenutext.Text = ans


            conn1.Close()

            '點人頭出現
            Dim info As String = ""
            Dim days As String = Trim(Year(Today) - 1911).PadLeft(3, "0") & Month(Today).ToString.PadLeft(2, "0")
            If myobj.Objectmstore = "1" Then '多店
                info = myobj.mstore_id
            Else
                info = "'" & Request.Cookies("store_id").Value & "'"
            End If
            work_person1.open_database(Request.Cookies("webfly_empno").Value, Request.Cookies("store_id").Value, days, info)

            '系統更新則數
            sql = "select * from system_for_update_read With(NoLock) "
            sql += " where emp_no='" & Request.Cookies("webfly_empno").Value & "'"
            Dim table3 As New DataTable
            Dim adp1 As SqlDataAdapter = New SqlDataAdapter(sql, conn1)
            adp1.Fill(table3)
            If table3.Rows.Count > 0 Then
                Dim cmessage As Integer = getSyscountRow(table3.Rows(0)("readdate"))
                'Response.Write(cmessage)
                'If cmessage = 0 Then
                If cmessage <> 0 Then
                    Label7.Text = cmessage
                Else
                    ring_warn.Style("Display") = "none"
                End If
                'End If
                'Dim dr As Data.DataRow = getSyscountRow(table3.Rows(0)("readdate"))
                'If Not dr Is Nothing Then
                '    If dr("sum") <> 0 Then
                '        Label7.Text = dr("sum")
                '    Else
                '        ring_warn.Style("Display") = "none"
                '    End If
                'End If
            Else
                Dim cmessage As Integer = getSyscountRow("2015/6/3 上午 10:42:01")
                'If cmessage = 0 Then
                If cmessage <> 0 Then
                    Label7.Text = cmessage
                Else
                    ring_warn.Style("Display") = "none"
                End If
                'End If
                'Dim dr As Data.DataRow = getSyscountRow("2015/6/3 上午 10:42:01")
                'If Not dr Is Nothing Then
                '    If dr("sum") <> 0 Then
                '        Label7.Text = dr("sum")
                '    Else
                '        ring_warn.Style("Display") = "none"
                '    End If
                'End If
            End If

            ''系統更新則數
            'Dim dr As Data.DataRow = getSyscountRow()
            'If Not dr Is Nothing Then
            '    If dr("sum") <> 0 Then
            '        Label7.Text = dr("sum")
            '    Else
            '        ring_warn.Style("Display") = "none"
            '    End If
            'End If
        End Using

        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
            'If Request.Cookies("webfly_empno").Value.ToUpper = "H2L" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CT8" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CPM" Or Request.Cookies("webfly_empno").Value.ToUpper = "05YZ" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CQ0" Or Request.Cookies("webfly_empno").Value.ToUpper = "0026" Or Request.Cookies("webfly_empno").Value.ToUpper = "0AG2" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CFY" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CVZ" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CQ1" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CUN" Or Request.Cookies("webfly_empno").Value.ToUpper = "M3C" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CVA" Or Request.Cookies("webfly_empno").Value.ToUpper = "0CWV" Then
            直營 = "1"
            'Else
            '    直營 = "2"
            'End If
        Else
            直營 = "3"
        End If

    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Request.Cookies("webfly_empno").Value = "PL7" Or Request.Cookies("webfly_empno").Value = "pl7" Then
        '    Response.Write(Session.Item("Safeloginkey"))
        'End If

        Dim ip As String = Request.ServerVariables("REMOTE_ADDR") 'Left(ip, 10) <>ip 10.40.201.
        If Left(ip, 3) <> "10." Then
            Label2.Text = ""
        Else
            Label2.Text = "<img id=""vpn"" src=""../images/lock_icon.png"" width=""14"" height=""16"" align=""absmiddle"" />"
        End If
        'Response.Write("aa")
        If Request.Cookies("webfly_level") Is Nothing Or Request.Cookies("webfly_empno") Is Nothing Or Request.Cookies("store_id") Is Nothing Then  '判斷 cookie aa 是否存在            
            'Response.Redirect("../indexnew/terminated.aspx")
            Response.End()
        End If

        Dim store_no = Request.Cookies("store_id").Value.Trim()
        Dim emp_no = Request("webfly_empno").Trim()
        Dim man_level = Request("webfly_level").Trim()

        If Not IsPostBack Then
            'store_no = Request.Cookies("store_id").Value.Trim()
            'emp_no = Request("webfly_empno").Trim()
            Dim log_title As String = "頁面使用記錄"
            Dim url As String = Request.Url.ToString
            Dim query As String = ""
            Dim message As String = "user ip:" & ip
            checkleave.錯誤訊息(store_no, emp_no, log_title, url, query, message)

            'Label1.Text = System.Web.HttpUtility.UrlDecode(Request.Cookies("webfly_level").Value, System.Text.Encoding.Default) & " " & System.Web.HttpUtility.UrlDecode(Request.Cookies("webfly_name").Value, System.Text.Encoding.Default)
            Label1.Text = checkleave.leave(Request.Cookies("webfly_empno").Value) & " " & checkleave.name(Request.Cookies("webfly_empno").Value)
            Label5.Text = checkleave.store_name(Request.Cookies("store_id").Value)

            '1050217超過60分登出(無formsumbit)
            'Page.ClientScript.RegisterStartupScript(Me.GetType(), "onLoad", "DisplaySessionTimeout();", True)
        End If

        Dim data = String.Format("uid={0}&sid={1}", emp_no, store_no)
        Dim dataEncrypt = GetEncryptResult(data, "etwarmetwarmetwarmetwarmetwarmet", "0123456789ABCDEF")
        HouseHunterUrl = String.Format("https://bot.etwarm.net.tw/rakuya/superweb_redirect.php?data={0}", dataEncrypt)
        'spy需有申請vpn,357(作癈，20170710，阿甘)
        'spy需有申請vpn和有繳費
        If checkvpn.checkvpn(Request.Cookies("store_id").Value.ToString) = "Y" Then '有申請vpn   myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "357", "NO") = "1" Then
            'Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            '    conn.Open()
            'If checkvpn.spytimes(Request.Cookies("webfly_empno").Value.ToString) = "Y" Or Request.Cookies("webfly_empno").Value.ToString = "92H" Then
            'Label3.Text = "<a href='#' onClick=""window.open('../A_ObjectManage/spy_Redirect.aspx', 'new', 'status=no,toolbar=no,scrollbars=yes,resizable=yes,width=1100,Height=600')"";><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'sql = "select man_emp_no, man_password FROM PSMAN WHERE  man_emp_no= '" & Request.Cookies("webfly_empno").Value.ToString & "' and man_dept_no  = '" & Request.Cookies("store_id").Values.ToString & "'"
            'adpt1 = New SqlDataAdapter(sql, conn)
            'ds1 = New DataSet()
            'adpt1.SelectCommand.CommandTimeout = 200
            'adpt1.Fill(ds1, "table1")
            'Dim table1 As DataTable = ds1.Tables("table1")
            'If table1.Rows.Count > 0 Then

            'sql = "update web_powerset set log_times=log_times+1 where man_emp_no='" & Request.Cookies("webfly_empno").Value.ToString & "'"
            'Response.Write(sql)
            'Dim cmd As New SqlCommand(sql, conn)
            'cmd.CommandType = CommandType.Text
            'Try
            '    cmd.ExecuteNonQuery()
            'Catch ex As Exception
            '    Response.Write(ex.ToString)
            'End Try
            '    Label3.Text = "<a href='#' onClick=""window.open('http://flydatas.etwarm.com.tw/ETSpy/ws.asmx/CheckUser?account=" & Request.Cookies("webfly_empno").Value & "&password=" & table1.Rows(0)("man_password").ToString & "', 'new', 'status=no,toolbar=no,scrollbars=yes,resizable=yes,width=1100,Height=600')"";><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'Else
            '    Label3.Text = "<a href=""#"" onClick=""alert('尚未有同業物件使用權限,請先申請VPN!')""><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'End If
            'Else
            'Label3.Text = "<a href=""#"" onClick=""alert('已達本日使用上限,請明日再做使用!')""><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'End If
            ' End Using
            'If checkvpn.checkspyuse(Request.Cookies("store_id").Value.ToString) = "Y" Then '有訂購同業比對
            Label3.Text = String.Format("<a href='{0}' onClick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}});""><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>", HouseHunterUrl, store_no, emp_no, "Z", "房產獵人")

            'Label3.Text = String.Format("<a href='#' onClick=""window.open('{0}', 'new', 'status=no,toolbar=no,scrollbars=yes,resizable=yes,width=1100,Height=600')"";><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>", HouseHunterUrl)
            'Label3.Text = "<a href='#' onClick=""window.open('../A_ObjectManage/spy_Redirect.aspx', 'new', 'status=no,toolbar=no,scrollbars=yes,resizable=yes,width=1100,Height=600')"";><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'Else
            '    Label3.Text = "<a href=""#"" onClick=""alert('尚未訂購同業物件比對系統，故無使用權限，請洽詢各專戶!')""><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
            'End If
        Else '有vpn並有spy權限

            Label3.Text = "<a href=""#"" onClick=""alert('尚未申請VPN，故無使用權限,請先申請VPN或洽詢各專戶!')""><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"
        End If
        '商品訂購217
        If myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "217", "NO") = "1" Then
            Label4.Text = "<a href=""../C_ShopKeeping/goods_order.aspx"" ><img src=""../images/quick_menu_02.png"" width=""80"" height=""34"" /></a>"
        End If

        ' Label3.Text = "<a href='#' onClick=""window.open('http://10.40.200.3/ETSpy/ws.asmx/CheckUser?account=92h&password=89446259', 'new', 'status=no,toolbar=no,scrollbars=yes,resizable=yes,width=1100,Height=600')"";><img src=""../images/quick_menu_07.png"" width=""102"" height=""34"" /></a>"

        'Label8.text = "<ul>"
        Label6.Text = ""
        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
            Dim 網址 As String
            網址 = "https://superwebnew.etwarm.com.tw/dm.aspx"
            Label6.Text = "<li><a href=""" & 網址 & """ target=""_blank"" >個人DM</a></li>"
            Label6.Text += "<li><a href = ""../TOP_TOOLS/tool_letter_DS.aspx"" > 開發信</a></li>"
            'Label6.Text += "<li><a href = ""https://realprice.etwarm.net.tw/"" target=""_blank"" style=""color: red;"">新版成交行情</a></li>"
            Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">新版成交行情</a></li>", "https://realprice.etwarm.net.tw", store_no, emp_no, "Z", "新版成交行情")
        Else
            Dim 網址 As String
            網址 = "https://superwebnew.etwarm.com.tw/dm.aspx"
            Label6.Text = "<li><a href=""" & 網址 & """ target=""_blank"" >個人DM</a></li>"
            Label6.Text += "<li><a href = ""../A_ObjectManage/tool_dm_print.aspx"">型錄列印</a></li>"
            Label6.Text += "<li><a href = ""../TOP_TOOLS/tool_letter.aspx"" > 開發信</a></li>"
        End If
        'Label6.Text += "<li><a href = ""../D_community/deal_price_s.aspx"" >成交行情</a></li>"
        Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">成交行情</a></li>", "https://realprice.etwarm.net.tw", store_no, emp_no, "Z", "成交行情")
        Label6.Text += "<li><a href = ""../G_WebService/SMS_main.aspx"" >簡訊發送及購買</a></li>"
        Label6.Text += "<li><a href = ""../TOP_TOOLS/各項稅務試算.aspx"" >各項稅務試算</a></li>"
        Label6.Text += "<li><a href = ""http://www.etwarm.com.tw/efinger/index.html"" target=""_blank"" >e指通</a></li>"
        Label6.Text += "<li><a href = ""https://www.etwarm.com.tw/linebot"" target=""_blank"" >東屋小秘書</a></li>"
        'Label6.Text += "<li><a href = ""https://www.etwarm.com.tw/actions/linephoto"" target=""_blank"" >Line貼圖製作</a></li>"

        'If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
        Dim urlAIO = "https://www.etwarm.com.tw/AIO/index.php"
        Label6.Text += "<li><a href=""" & urlAIO & """ target=""_blank"" >產調懶人包</a></li>"
        '2024/05/27 改為直營即可使用
        'If HavePermissionForHouseHounter(emp_no) Then 
        urlAIO = String.Format("https://bot.etwarm.net.tw/rakuya/superweb_redirect.php?data={0}", dataEncrypt)
        Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">房產獵人</a></li>", urlAIO, store_no, emp_no, "Z", "房產獵人")

        'If store_no = "A0001" OrElse checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
        urlAIO = String.Format("https://house-report.etwarm.net.tw/index.php?agentNo={0}&token=", emp_no)
        Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">銷售企劃書</a></li>", urlAIO, store_no, emp_no, "Z", "銷售報告書")
        'End If 

        If store_no = "A0001" Or store_no = "A0002" Then
            Dim urlStore = "/encrypt_redirect.aspx?type_id=1"
            Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">網路商城部落格</a></li>", urlStore, store_no, emp_no, "Z", "網路商城部落格")
            Dim urlSecretaryInventory = "https://liff.line.me/1654077574-QXMWr3GO"
            Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">小秘書庫存案件</a></li>", urlSecretaryInventory, store_no, emp_no, "Z", "小秘書庫存案件")
        End If

        '直營店
        If checkleave.直營(Request.Cookies("store_id").Value) = "Y" AndAlso InStr(checkleave.直營leave(Request.Cookies("webfly_empno").Value), "秘書") > 0 Then
            'urlAIO = "https://superwebnew6.etwarm.com.tw/ExtendWebsite/PropertyHandover/index.aspx"
            urlAIO = "https://superwebnew.etwarm.com.tw/ExtendWebsite/PropertyHandover/index.aspx"
            Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">點交系統</a></li>", urlAIO, store_no, emp_no, "Z", "點交系統")
        End If

        '所有店
        'If InStr(checkleave.直營leave(Request.Cookies("webfly_empno").Value), "秘書") > 0 Then
        '    urlAIO = "https://superwebnew6.etwarm.com.tw/ExtendWebsite/PropertyHandover/index.aspx"
        '    Label6.Text += String.Format("<li><a href='{0}' target='_blank' onclick=""addUsageRecord(event,this,{{store_no:'{1}', emp_no:'{2}', category_no:'{3}', usage_name:'{4}'}})"">點交系統</a></li>", urlAIO, store_no, emp_no, "Z", "點交系統")
        'End If
        'End If
        'End If
        'Label8.text += "<div style = ""height: 14px;"" <> img src=""../images/sub_list_bg_bottom.png"" alt=""""></div>"
        'Label8.text += "</ul>"

        Session("key") = "123"
        '每10分檢查一次
        'Response.Write("aa" & DateTime.Now.AddSeconds(600).ToString)
        'loglist()
    End Sub
    Private Sub InitializeComponent()
        AddHandler Me.Load, AddressOf Me.Page_Load
    End Sub

    '系統更新則數
    Function getSyscountRow(ByVal readdate As String) As Integer
        Dim sqlsor As New SqlDataSource
        sqlsor.ConnectionString = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
        Dim cmessage As Integer = 0
        '業務員代號
        Try
            Dim 條件 As String
            If checkleave.直營(Request.Cookies("store_id").Value) = "Y" Then
                條件 = " and mode in ('2','3') "
            Else
                條件 = " and mode in ('1','3') "
            End If

            sqlsor.SelectCommand = "select updatedate from system_for_update With(NoLock) where used='Y' " & 條件
            'If Request.Cookies("webfly_empno").Value = "02K3" Then
            'Response.Write(sqlsor.SelectCommand & "<br>")
            'Response.Write(DateDiff("n", updatedate, readdate))
            'Response.End()
            'End If
            sqlsor.DataBind()
            Dim dv As Data.DataView
            dv = CType(sqlsor.Select(New System.Web.UI.DataSourceSelectArguments), Data.DataView)
            'Response.Write("<br>" & readdate & "count:" & dv.Count & "<br>")
            ' Response.End()
            For i = 0 To dv.Count - 1
                'Response.Write("<br>update:" & dv.Table.Rows(i)(0) & "read:" & readdate & "count:" & dv.Count & "<br>")
                'Response.Write(DateDiff("m", dv.Table.Rows(i)(0), readdate))
                If DateDiff("s", dv.Table.Rows(i)(0), readdate) <= 0 Then
                    cmessage += 1
                End If
            Next
            'If Not IsDBNull(dv.Table.Rows(0)(0)) Then
            '    Return dv.Table.Rows(0)
            'Else
            '    Return Nothing
            'End If
            'Response.Write("<br>count:" & cmessage & "<br>")
            Return cmessage

        Catch ex As Exception
            'Response.Write(ex.ToString)
            'Throw ex
        Finally
            sqlsor.Dispose()
        End Try
    End Function


    '系統更新則數
    'Function getSyscountRow(ByVal readdate As String) As Data.DataRow
    '    Dim sqlsor As New SqlDataSource
    '    sqlsor.ConnectionString = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    '    '業務員代號
    '    Try
    '        sqlsor.SelectCommand = "select count(*) as sum from system_for_update where used='Y' and (updatedate > '" & Now() & "' and updatedate > '" & readdate & "')"
    '        'If Request.Cookies("webfly_empno").Value = "02K3" Then
    '        '    Response.Write(sqlsor.SelectCommand)
    '        'End If
    '        sqlsor.DataBind()
    '        Dim dv As Data.DataView
    '        dv = CType(sqlsor.Select(New System.Web.UI.DataSourceSelectArguments), Data.DataView)
    '        If Not IsDBNull(dv.Table.Rows(0)(0)) Then
    '            Return dv.Table.Rows(0)
    '        Else
    '            Return Nothing
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        sqlsor.Dispose()
    '    End Try
    'End Function

    'Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
    '    '登出系統
    '    eip_usual.Show("log out")
    'End Sub

    ''檢查是否重覆登入
    'Sub loglist()
    '    Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
    '        conn.Open()
    '        sql1 = "select ip from Superweb_Online where empno='" & Request.Cookies("webfly_empno").Value & "'"
    '        adpt1 = New SqlDataAdapter(sql1, conn)
    '        ds1 = New DataSet()
    '        adpt1.Fill(ds1, "table1")
    '        table1 = ds1.Tables("table1")
    '        If table1.Rows.Count > 1 Then '表有1人以上登入
    '            sql = "select top 1 ip,passkey from Superweb_LoginRecord where empno='" & Request.Cookies("webfly_empno").Value & "' and login_check='Y' order by login_time desc"
    '            Dim cmd As New SqlCommand(sql, conn)
    '            cmd.CommandType = CommandType.Text
    '            Dim dr As SqlDataReader = cmd.ExecuteReader
    '            If dr.Read() Then
    '                If dr("ip") = Request.ServerVariables("REMOTE_ADDR") Then 'ip 
    '                    log_delete()
    '                    log_add()
    '                Else
    '                    If Session("key").ToString <> dr("passkey").ToString Then
    '                        eip_usual.Show("同一帳號不可同時登入")
    '                    Else
    '                        log_delete()
    '                        log_add()
    '                    End If
    '                End If
    '            Else
    '                log_update()
    '            End If
    '        End If
    '    End Using

    'End Sub

    'Sub log_add()
    '    Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
    '        conn.Open()
    '        sql1 = "insert Superweb_Online (empno,ip,login_time) values ('" & Request.Cookies("webfly_empno").Value & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Now() & "')"
    '        Dim cmd1 As New SqlCommand(sql1, conn)
    '        cmd1.CommandType = CommandType.Text
    '    End Using
    'End Sub

    'Sub log_update()
    '    Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
    '        conn.Open()
    '        sql1 = "update Superweb_Online set login_time = '" & Now & "' where empno='" & Request.Cookies("webfly_empno").Value & "'"
    '        Dim cmd1 As New SqlCommand(sql1, conn)
    '        cmd1.CommandType = CommandType.Text
    '    End Using
    'End Sub

    'Sub log_delete()
    '    Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
    '        conn.Open()
    '        sql1 = "delete Superweb_Online where empno='" & Request.Cookies("webfly_empno").Value & "'"
    '        Dim cmd1 As New SqlCommand(sql1, conn)
    '        cmd1.CommandType = CommandType.Text
    '    End Using
    'End Sub

    Function HavePermissionForHouseHounter(ByVal empNo As String) As Boolean
        Dim cnString As String = EGOUPLOADSqlConnStr

        Dim query As String = "" &
"Select COUNT(1) " &
"From PSMAN emp WITH (NOLOCK) " &
"Join HSBSMG dep WITH (NOLOCK) " &
"On emp.man_dept_no = dep.bs_dept " &
"WHERE dep.bs_直營識別 = 'Y' " &
"And ISNULL(man_quit_dt, '') = '' " &
"And ( " &
"( " &
"emp.man_level Like '%經理%' " &
"Or emp.man_level Like '%副理%' " &
") " &
"Or emp.man_emp_no IN ('H2I','0026','H2L') " &
") " &
"And TRIM(emp.man_emp_no) = @emp_no"

        Dim count As Integer = 0

        Using cn As New SqlConnection(cnString)
            Using cm As New SqlCommand(query, cn)
                cm.Parameters.AddWithValue("@emp_no", empNo)

                Try
                    cn.Open()
                    count = Convert.ToInt32(cm.ExecuteScalar())
                Catch ex As Exception
                Finally
                    cn.Close()
                End Try
            End Using
        End Using
        Dim result = If(count > 0, True, False)
        Return result
    End Function

    Public Shared Function GetEncryptResult(ByVal content As String, ByVal key As String, ByVal iv As String) As String
        Dim keyBytes = Encoding.ASCII.GetBytes(key)
        Dim ivBytes = Encoding.UTF8.GetBytes(iv)
        'var ivBytes = GenerateRandomBytes(16);

        Using aes = New AesCryptoServiceProvider()
            aes.KeySize = 256
            aes.BlockSize = 128
            aes.Key = keyBytes
            aes.IV = ivBytes
            aes.Mode = CipherMode.CBC
            aes.Padding = PaddingMode.PKCS7

            Dim encryptor As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)

            Dim encryptedBytes As Byte() = Nothing
            Dim result = String.Empty
            Using ms = New MemoryStream()
                Using cs = New CryptoStream(ms, encryptor, CryptoStreamMode.Write)
                    Dim plainBytes = Encoding.UTF8.GetBytes(content)
                    cs.Write(plainBytes, 0, plainBytes.Length)
                End Using

                encryptedBytes = ms.ToArray()
                result = Convert.ToBase64String(encryptedBytes)
            End Using
            Return result
        End Using
    End Function


    Public Shared Function GetEncryptResultWithTimestamp(ByVal content As String, ByVal key As String, ByVal iv As String) As String
        Dim keyBytes = Encoding.ASCII.GetBytes(key)
        Dim ivBytes = Encoding.UTF8.GetBytes(iv)
        'var ivBytes = GenerateRandomBytes(16);
        Dim timestamp As String = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString()
        Dim dataWithTimestamp As String = String.Format("{0}:{1}", timestamp, content)

        Using aes = New AesCryptoServiceProvider()
            aes.KeySize = 256
            aes.BlockSize = 128
            aes.Key = keyBytes
            aes.IV = ivBytes
            aes.Mode = CipherMode.CBC
            aes.Padding = PaddingMode.PKCS7

            Dim encryptor As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)

            Dim encryptedBytes As Byte() = Nothing
            Dim result = String.Empty
            Using ms = New MemoryStream()
                Using cs = New CryptoStream(ms, encryptor, CryptoStreamMode.Write)
                    Dim plainBytes = Encoding.UTF8.GetBytes(dataWithTimestamp)
                    cs.Write(plainBytes, 0, plainBytes.Length)
                End Using

                encryptedBytes = ms.ToArray()
                result = Convert.ToBase64String(encryptedBytes)
            End Using
            Return result
        End Using
    End Function

    Public Shared Function GetDecryptResultWithTimestamp(ByVal encryptedContent As String, ByVal key As String, ByVal iv As String) As String
        Dim keyBytes = Encoding.ASCII.GetBytes(key)
        Dim ivBytes = Encoding.UTF8.GetBytes(iv)

        Using aes = New AesCryptoServiceProvider()
            aes.KeySize = 256
            aes.BlockSize = 128
            aes.Key = keyBytes
            aes.IV = ivBytes
            aes.Mode = CipherMode.CBC
            aes.Padding = PaddingMode.PKCS7

            Dim decryptor = aes.CreateDecryptor(aes.Key, aes.IV)

            Dim decryptedBytes As Byte() = Nothing
            Dim plainText = String.Empty
            Using ms = New MemoryStream()
                Using cs = New CryptoStream(ms, decryptor, CryptoStreamMode.Write)
                    Dim encryptedBytes = Convert.FromBase64String(encryptedContent)
                    cs.Write(encryptedBytes, 0, encryptedBytes.Length)
                End Using

                decryptedBytes = ms.ToArray()
                plainText = Encoding.UTF8.GetString(decryptedBytes)
            End Using

            Dim parts As String() = plainText.Split(":"c)
            If parts.Length <> 2 Then
                Return String.Empty
            End If

            Dim timestamp As Long = Long.Parse(parts(0))
            Dim originalData As String = parts(1)

            ' 驗證時間戳是否過期
            Dim currentTimestamp As Long = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
            If currentTimestamp - timestamp > 8640 Then
                Return String.Empty
            End If

            Return originalData
        End Using
    End Function
End Class
