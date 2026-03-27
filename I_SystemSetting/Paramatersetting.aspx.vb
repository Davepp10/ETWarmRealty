Imports System.Data
Imports System.Data.SqlClient
Partial Class System_Paramatersetting
    Inherits System.Web.UI.Page

    Private Const LFTAutoSyncParamName As String = "新增物件自動同步至聯房通"
    Private Const ICBBatchSyncSort As String = "21"
    Private Const ICBBatchSyncParamaterID As String = "A"
    Private Const ICBBatchSyncNum As String = "21"

    Public conn As SqlConnection
    Public cmd As SqlCommand
    Public ds As DataSet
    Public adpt As SqlDataAdapter
    Public sysdate As String
    Dim i As Integer
    Dim sql As String
    Dim table1, table2 As DataTable


    Public mstore_id1 As Object = ""
    Dim myobj As New clspowerset


    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    Dim same_EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("same_EGOUPLOADConnectionString").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.Cookies("webfly_empno") Is Nothing Then
            Response.Redirect("../indexnew/login3.aspx")
        End If

        If Not IsPostBack Then
            '帶入店代號資料
            store.Items.Clear()
            If Left(Request.Cookies("store_id").Value, 1) <> "K" Then
                store = clspowerset.scope(Request.Cookies("webfly_empno").Value, "store", store) '店代號
            Else
                store = clspowerset.scope_same(Request.Cookies("webfly_empno").Value, "store", store) '店代號
            End If

            store.SelectedValue = Request.Cookies("store_id").Value

            Load_Paramater()
        End If

    End Sub

    Sub Load_Paramater()
        Dim dbcon As String = ""
        dbcon = GetDbConnectionString(store.SelectedValue)

        Using conn As New SqlConnection(dbcon)
            'If TorF = "True" Then
            'sql = "select a.sort,b.UseState,a.ParamaterID,a.Num,a.ParamaterName,b.Value,a.ParamaterContent from ParamaterContent a left join ParamaterSetting b on a.ParamaterID=b.ParamaterID and a.Num=b.Num where StoreID='" & store_id & "' and Show='Y' order by sort asc"
            'Else
            '一開始都先讀預設，避免新增的參數，已存的資料會讀不出來
            sql = "select a.sort,b.UseState,a.ParamaterID,a.Num,a.ParamaterName,b.Value,a.ParamaterContent from ParamaterContent a left join ParamaterSetting b on a.ParamaterID=b.ParamaterID and a.Num=b.Num where StoreID='None' and Show='Y'  order by sort " '若判斷資料表無任何值，則抓取預設值'None'
            'End If

            'Response.Write(sql)
            conn.Open()

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            Dim t1 As DataTable = ds.Tables("table1")
            'Response.Write("aa" & t1.Rows.Count & "aa" & "<br>")

            With GridView1
                .DataSource = t1.DefaultView
                .DataBind()
            End With
        End Using


    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Dim store_id As String = Request.Cookies("store_id").Value
        Dim dbcon As String = ""
        dbcon = GetDbConnectionString(store.SelectedValue)

        Using conn As New SqlConnection(dbcon)
            sql = "select a.sort,b.UseState,a.ParamaterID,a.Num,a.ParamaterName,b.Value,a.ParamaterContent from ParamaterContent a left join ParamaterSetting b on a.ParamaterID=b.ParamaterID and a.Num=b.Num where StoreID='" & store.SelectedValue & "' and Show='Y' order by sort asc"
            'Response.Write(sql & "<br>")
            conn.Open()

            adpt = New SqlDataAdapter(sql, conn)
            ds = New DataSet()
            adpt.Fill(ds, "table1")
            Dim t1 As DataTable = ds.Tables("table1")
            'Response.Write(t1.Rows.Count & "<br>")

            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim i As Integer = 0
                Dim ParamaterID As Label = e.Row.FindControl("ParamaterID")
                Dim Num As Label = e.Row.FindControl("Num")
                Dim chk As CheckBox = e.Row.FindControl("UseState")
                Dim paramaterName As Label = e.Row.FindControl("Label3")
                Dim value As TextBox = e.Row.FindControl("Value")
                Dim valueRadio As RadioButtonList = e.Row.FindControl("ValueRadio")
                Dim btnSyncICBStoreObjects As Button = e.Row.FindControl("btnSyncICBStoreObjects")
                If t1.Rows.Count > 0 Then
                    For i = 0 To t1.Rows.Count - 1
                        If ParamaterID.Text = t1.Rows(i)("ParamaterID") And Num.Text = t1.Rows(i)("Num") Then

                            If t1.Rows(i)("UseState") = "Y" Then
                                chk.Checked = True
                            Else
                                chk.Checked = False
                            End If

                            value.Text = t1.Rows(i)("Value")
                            Exit For
                        End If
                    Next
                End If
                If paramaterName IsNot Nothing AndAlso paramaterName.Text = LFTAutoSyncParamName Then
                    value.Visible = False
                    valueRadio.Visible = True
                    If value.Text = "1" Or value.Text = "Y" Or value.Text = "是" Then
                        valueRadio.SelectedValue = "1"
                    Else
                        valueRadio.SelectedValue = "0"
                    End If
                End If
                If btnSyncICBStoreObjects IsNot Nothing Then
                    btnSyncICBStoreObjects.Visible = (Trim(ParamaterID.Text) = ICBBatchSyncParamaterID AndAlso Trim(Num.Text) = ICBBatchSyncNum AndAlso Trim(CType(e.Row.FindControl("sort"), Label).Text) = ICBBatchSyncSort)
                End If
                If Num.Text = "19" Then '登出時間
                    chk.Checked = True
                    chk.Enabled = False
                End If
            End If
        End Using

    End Sub

    Function chk_Set(ByVal store_id As String, ByVal ParamaterID As String, ByVal Num As String) As String
        Dim TorF As String = ""

        Dim dbcon As String = ""
        dbcon = GetDbConnectionString(store.SelectedValue)

        Using conn As New SqlConnection(dbcon)
            sql = "select * from  ParamaterSetting  where StoreID='" & store.SelectedValue & "' and ParamaterID='" & ParamaterID & "' and Num='" & Num & "'"
            'Response.Write(sql)
            cmd = New SqlCommand(sql, conn)
            conn.Open()

            Dim dr As SqlDataReader = cmd.ExecuteReader
            If dr.Read Then
                TorF = "True"
            Else
                TorF = "False"
            End If

        End Using

        Return TorF
    End Function

  
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        'Dim store_id As String = Request.Cookies("store_id").Value
        Dim messages As String = "成功"
        Dim counts As Integer = 0
        Dim gr As GridViewRow

        Dim dbcon As String = ""
        dbcon = GetDbConnectionString(store.SelectedValue)

        Using conn As New SqlConnection(dbcon)
            conn.Open()

            For Each gr In GridView1.Rows
                Dim Sort As Label = CType(gr.FindControl("Sort"), Label)
                Dim ParamaterID As Label = CType(gr.FindControl("ParamaterID"), Label)
                Dim Num As Label = CType(gr.FindControl("Num"), Label)
                Dim ParamaterName As Label = CType(gr.FindControl("Label3"), Label)
                Dim UseState As CheckBox = CType(gr.FindControl("UseState"), CheckBox)
                Dim Value As TextBox = CType(gr.FindControl("Value"), TextBox)
                Dim ValueRadio As RadioButtonList = CType(gr.FindControl("ValueRadio"), RadioButtonList)
                Dim TorF As String = chk_Set(store.SelectedValue, ParamaterID.Text, Num.Text)
                If TorF = "True" Then
                    sql = "Update  ParamaterSetting set Value=@Value,UseState=@UseState where StoreID=@StoreID and Sort=@Sort"
                    'Response.Write(sql & "   " & Num.Text & "<br>")
                    cmd = New SqlCommand(sql, conn)

                    '參數值
                    cmd.Parameters.Add(New SqlParameter("@Value", SqlDbType.NVarChar))
                    If ParamaterName.Text = LFTAutoSyncParamName Then
                        cmd.Parameters("@Value").Value = ValueRadio.SelectedValue
                    ElseIf Num.Text <> "19" Then
                        cmd.Parameters("@Value").Value = Value.Text
                    Else
                        If Value.Text = "" Or Val(Value.Text) < 20 Or Val(Value.Text) > 120 Then
                            '登出時間長短,預設20Min
                            messages = "時間需介於20-120分"
                            counts += 1
                        Else
                            cmd.Parameters("@Value").Value = Value.Text
                            messages = "修改成功"
                        End If
                    End If


                    '是否使用
                    cmd.Parameters.Add(New SqlParameter("@UseState", SqlDbType.NVarChar))
                    If Num.Text <> "19" Then '登出時間一定有
                        If UseState.Checked = True Then
                            cmd.Parameters("@UseState").Value = "Y"
                        Else
                            cmd.Parameters("@UseState").Value = "N"
                        End If
                    Else
                        cmd.Parameters("@UseState").Value = "Y"
                    End If


                    'Cookie店代號StoreID
                    cmd.Parameters.Add(New SqlParameter("@StoreID", SqlDbType.NVarChar))
                    cmd.Parameters("@StoreID").Value = store.SelectedValue

                    '排序
                    cmd.Parameters.Add(New SqlParameter("@Sort", SqlDbType.NVarChar))
                    cmd.Parameters("@Sort").Value = Sort.Text

                ElseIf TorF = "False" Then
                    sql = "Insert into  ParamaterSetting (StoreID,Sort,ParamaterID,NUM,UseState,Value) Values(@StoreID,@Sort,@ParamaterID,@NUM,@UseState,@Value)"
                    cmd = New SqlCommand(sql, conn)

                    'Cookie店代號StoreID
                    cmd.Parameters.Add(New SqlParameter("@StoreID", SqlDbType.NVarChar))
                    cmd.Parameters("@StoreID").Value = store.SelectedValue

                    '排序
                    cmd.Parameters.Add(New SqlParameter("@Sort", SqlDbType.NVarChar))
                    cmd.Parameters("@Sort").Value = Sort.Text

                    '參數ID
                    cmd.Parameters.Add(New SqlParameter("@ParamaterID", SqlDbType.NVarChar))
                    cmd.Parameters("@ParamaterID").Value = ParamaterID.Text

                    '參數序號
                    cmd.Parameters.Add(New SqlParameter("@Num", SqlDbType.NVarChar))
                    cmd.Parameters("@Num").Value = Num.Text

                    '是否使用
                    cmd.Parameters.Add(New SqlParameter("@UseState", SqlDbType.NVarChar))
                    If Num.Text <> "19" Then '登出時間
                        If UseState.Checked = True Then
                            cmd.Parameters("@UseState").Value = "Y"
                        Else
                            cmd.Parameters("@UseState").Value = "N"
                        End If
                    Else '登出時間
                        cmd.Parameters("@UseState").Value = "Y"
                    End If

                    '參數值
                    cmd.Parameters.Add(New SqlParameter("@Value", SqlDbType.NVarChar))
                    If ParamaterName.Text = LFTAutoSyncParamName Then
                        cmd.Parameters("@Value").Value = ValueRadio.SelectedValue
                        messages = "新增成功"
                    ElseIf Num.Text <> "19" Then '登出時間長短
                        cmd.Parameters("@Value").Value = Value.Text
                        messages = "新增成功"
                    Else
                        'Response.Write(Value.Text)
                        If Value.Text = "" Or Val(Value.Text) < 20 Or Val(Value.Text) > 120 Then
                            messages = "時間需介於20-120分"
                            counts += 1
                        Else
                            cmd.Parameters("@Value").Value = Value.Text
                            messages = "新增成功"
                        End If
                    End If
                End If
                If InStr(1, messages, "成功") <> 0 Then
                    Try
                        cmd.ExecuteNonQuery()
                        'eip_usual.Show(messages)
                    Catch ex As Exception
                        eip_usual.Show("失敗")
                        Response.Write(ex.ToString)
                    End Try

                Else
                    Try
                        eip_usual.Show(messages)
                        Exit Sub
                    Catch ex As Exception
                        Response.Write(ex.ToString)
                    End Try

                End If
            Next
            If messages = "時間需介於20-120分" And counts > 0 Then
                eip_usual.Show("逾時登出設定" & messages)
            Else
                eip_usual.Show(messages)
            End If
            'Response.Write(messages)
            'eip_usual.Show(messages)
        End Using
    End Sub

    Protected Sub store_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles store.SelectedIndexChanged
        Load_Paramater()
    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName <> "SyncICBStoreObjects" Then
            Return
        End If

        Try
            SyncICBStoreObjects(store.SelectedValue)
            eip_usual.Show("所有物件已同步至聯房通")
        Catch ex As Exception
            eip_usual.Show("同步失敗")
            Response.Write(ex.ToString)
        End Try
    End Sub

    Private Sub SyncICBStoreObjects(ByVal storeNo As String)
        Using conn As New SqlConnection(GetDbConnectionString(storeNo))
            Using cmd As New SqlCommand("xsp_icb_update_object_info_by_store_no", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@store_no", SqlDbType.NVarChar, 10)).Value = storeNo
                cmd.Parameters.Add(New SqlParameter("@agree_cb", SqlDbType.Bit)).Value = True
                cmd.Parameters.Add(New SqlParameter("@info_type", SqlDbType.Int)).Value = 1
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Function GetDbConnectionString(ByVal storeNo As String) As String
        If Not String.IsNullOrWhiteSpace(storeNo) AndAlso Left(storeNo, 1) = "K" Then
            Return same_EGOUPLOADSqlConnStr
        End If

        Return EGOUPLOADSqlConnStr
    End Function
End Class
