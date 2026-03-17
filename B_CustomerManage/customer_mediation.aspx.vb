Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Partial Class B_CustomerManage_customer_mediation
    Inherits System.Web.UI.Page
    Dim checkleave As New checkleave
    Dim object_fun As String = ""
    Dim oid As String = ""
    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString
    Dim conn1, conn2 As SqlConnection
    Dim ds1, ds2 As DataSet
    Dim adpt1, adpt2 As SqlDataAdapter
    Dim table1 As DataTable
    Dim sql As String = ""
    Public myobj As New clspowerset
    Dim sysdate As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)


    Structure customertype
        Dim customertype As String  '那一類客戶
        Dim tablename As String  '那一類客戶存入那個table
        Dim cumname1 As String  '那一類客戶編號
        Dim cumname2 As String  '那一類客戶編號
        Dim cumname3 As String  '那一類客戶編號
        Dim tablename1 As String  '那一類客戶存入那個table
        Dim tablename2 As String  '那一類客戶存入那個table
    End Structure

    Public Function 客戶類別() As customertype()
        Dim b(0) As customertype

        If object_fun = "sellcustomer" Then
            '買方
            b(0).tablename = "web_買方要約資料表"
            b(0).cumname1 = "買方編號"
            b(0).cumname2 = "買方編號"
            b(0).cumname3 = "銷售狀態"
            b(0).tablename1 = "(select 店代號,建築名稱,物件編號 from 委賣物件資料表 With(NoLock) union all select 店代號,建築名稱,物件編號 from 委賣物件過期資料表 With(NoLock))"
            b(0).tablename2 = "web_買方基本資料表"
        ElseIf object_fun = "rentcustomer" Then
            '租方
            b(0).tablename = "web_租方要約資料表"
            b(0).cumname1 = "租方編號"
            b(0).cumname2 = "租方編號"
            b(0).cumname3 = "租賃狀態"
            b(0).tablename1 = "委租物件資料表"
            b(0).tablename2 = "web_租方基本資料表"
        ElseIf object_fun = "othercustomer" Then
            '潛在
            b(0).tablename = "web_潛在要約資料表"
            b(0).cumname1 = "潛在編號"
            b(0).cumname2 = "編號"
            b(0).cumname3 = "銷售狀態"
            b(0).tablename1 = "委賣物件資料表"
            b(0).tablename2 = "web_潛在客戶基本資料表"
        End If
        Return b
    End Function

    '===========================
    ' 斡旋表單號碼（offerid）下拉資料
    '===========================
    Private Sub BindOfferIdDropDown(ByVal storeNo As String)
        If String.IsNullOrWhiteSpace(storeNo) Then
            offerid.Items.Clear()
            offerid.Items.Add(New ListItem("--請先選擇店代號--", ""))
            Return
        End If

        Dim dt As New DataTable()

        Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
            Using cmd As New SqlCommand("dbo.xsp_get_table_nos_by_store_no", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("@store_no", SqlDbType.NVarChar, 5).Value = storeNo.Trim()

                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using

        offerid.Items.Clear()
        offerid.Items.Add(New ListItem("--請選擇--", ""))

        For Each r As DataRow In dt.Rows
            Dim v As String = r("table_no").ToString().Trim()
            If v <> "" Then
                offerid.Items.Add(New ListItem(v, v))
            End If
        Next
    End Sub

    ' 取得目前要用的 offerid（編輯模式 DropDownList 會 Disabled，需用 HiddenField 保留值）
    Private Function GetOfferIdSuffix() As String
        Dim v As String = ""
        If offerid.Enabled Then
            v = If(offerid.SelectedValue, "")
        End If

        If String.IsNullOrWhiteSpace(v) Then
            v = If(hf_offerid.Value, "")
        End If

        Return v.Trim()
    End Function

    ' 設定 dropdown 選擇值，並依模式鎖定/解鎖
    Private Sub SetOfferIdSelection(ByVal offerSuffix As String, ByVal lockControl As Boolean)
        If offerid.Items.Count = 0 Then
            BindOfferIdDropDown(sid.Text)
        End If

        Dim v As String = If(offerSuffix, "").Trim()

        If v <> "" AndAlso offerid.Items.FindByValue(v) Is Nothing Then
            ' 若 SP 回傳的清單沒有舊值，仍要讓編輯畫面看得到
            offerid.Items.Insert(1, New ListItem(v, v))
        End If

        If offerid.Items.FindByValue(v) IsNot Nothing Then
            offerid.SelectedValue = v
        ElseIf offerid.Items.FindByValue("") IsNot Nothing Then
            offerid.SelectedValue = ""
        ElseIf offerid.Items.Count > 0 Then
            offerid.SelectedIndex = 0
        End If

        hf_offerid.Value = v
        offerid.Enabled = Not lockControl
    End Sub



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'PlaceHolder1.Controls.Add(DynamicUserControl)
        object_fun = Request("object_fun") '售或租屋主
        oid = Request("oid") '物件編號
        If Label0.Text = "no" Then
            sid.Text = Request("sid")
        End If


        myobj.power_object(Request.Cookies("webfly_empno").Value)
        myobj.mstores(Request.Cookies("webfly_empno").Value)
        myobj.mgroup(Request.Cookies("webfly_empno").Value)
        '權限
        myobj.Detail_Power(Request.Cookies("webfly_empno").Value, "313", "ALL")

        If Not IsPostBack Then
            offerdateb.Value = eip_usual.toROCyear(Now()) '起始日期
            ' 斡旋編號下拉：依店代號載入可用表單號碼
            BindOfferIdDropDown(sid.Text)
            hf_offerid.Value = ""
            '外部來不要有頭
            If Request("source") = "top" Or Request("source") = "worktop" Or Request("source") = "objecttop" Or Not IsNothing(Request("key")) Then
                main_menu1.Visible = False '上選單
                uniquename99.Style("Display") = "none"
                reveserd1.Visible = False '版權             
            End If

            If Request("state") = "update" Then
                Load_Data斡旋("OLD")
            End If
        End If
    End Sub

    '儲存要約
    Protected Sub save_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles save.Click
        If myobj.AC = "1" Or myobj.M = "1" Then
            Dim selstr As String = ""
            Dim different_state As String = ""
            Dim sp() As String = Split(Request.Form("objectname"), ",")
            Dim b As customertype
            If Request("oid") <> "" Then
                Dim offerSuffix As String = GetOfferIdSuffix()
                If String.IsNullOrWhiteSpace(offerSuffix) Then
                    eip_usual.Show("請先選擇編號!!")
                    Exit Sub
                End If

                If offerSuffix.ToUpper >= "JAB40001" And offerSuffix.ToUpper <= "JZZ99999" Then
                    If RadioButton1.Checked = False And RadioButton2.Checked = False Then
                        eip_usual.Show("請先選擇 是否願意提供個資!!")
                        Exit Sub
                    End If
                End If

                For Each b In 客戶類別()
                    If Label0.Text = "no" Then '新增
                        SqlDataSource1.InsertParameters.Clear()
                        SqlDataSource1.InsertParameters.Add("編號", Request("oid"))
                        SqlDataSource1.InsertParameters.Add("要約編號", sid.Text & offerSuffix)
                        SqlDataSource1.InsertParameters.Add("物件編號", sp(0))
                        SqlDataSource1.InsertParameters.Add("要約金", price.Text)
                        SqlDataSource1.InsertParameters.Add("要約起", offerdateb.Value)
                        SqlDataSource1.InsertParameters.Add("要約訖", offerdatee.Value)
                        SqlDataSource1.InsertParameters.Add("輸入者", Request.Cookies("webfly_empno").Value)
                        SqlDataSource1.InsertParameters.Add("店代號", Request("sid"))
                        SqlDataSource1.InsertParameters.Add("新增日期", eip_usual.toROCyear(Now()))
                        SqlDataSource1.InsertParameters.Add("修改日期", eip_usual.toROCyear(Now()))
                        SqlDataSource1.InsertParameters.Add("斡旋金", offerprice.Text)
                        SqlDataSource1.InsertParameters.Add("票據面額", offertick.Text)
                        If RadioButton1.Checked = True Then
                            SqlDataSource1.InsertParameters.Add("提供個資", "Y")
                        Else
                            SqlDataSource1.InsertParameters.Add("提供個資", "N")
                        End If
                        selstr = "insert into " & b.tablename & " (" & b.cumname1 & ",要約編號,物件編號,要約金,要約起,要約訖,輸入者,店代號,新增日期,修改日期,斡旋金,票據面額,提供個資) values (@編號,@要約編號,@物件編號,@要約金,@要約起,@要約訖,@輸入者,@店代號,@新增日期,@修改日期,@斡旋金,@票據面額,@提供個資)"

                        SqlDataSource1.InsertCommand = selstr
                        Try
                            SqlDataSource1.Insert()
                            eip_usual.Show("新增斡旋成功!")
                            SqlDataSource1.UpdateParameters.Add("編號", Request("oid"))
                            SqlDataSource1.UpdateParameters.Add("店代號", Request("sid"))
                            'If Request.Cookies("webfly_empno").Value = "92H" Then
                            selstr = "update " & b.tablename2 & " set " & b.cumname3 & "='斡旋' where " & b.cumname2 & "=@編號" & " and 店代號=@店代號"
                            'Response.Write(selstr)
                            'Response.End()
                            SqlDataSource1.UpdateCommand = selstr
                            Try
                                SqlDataSource1.Update()
                            Catch ex As Exception
                                Response.Write(ex.ToString)
                            End Try

                            'End If

                            '1050506表單控管
                            'If Request.Cookies("webfly_empno").Value = "92H" Then
                            'Response.Write(check_formno(offerSuffix, sid.Text, sp(0), sp(1), "新增"))
                            If check_formno(offerSuffix, sid.Text, sp(0), sp(1), "新增") = "use" Then
                                eip_usual.Show("一併寫入表單管理新")
                            Else
                                'eip_usual.Show("輸入編號為1050515前出貨之表單或不為所購買表單，無法寫入表單管理")
                            End If
                            'End If

                            clear斡旋()
                        Catch ex As Exception
                            'Response.Write(selstr & "<br>")
                            'Response.Write(ex.ToString)
                            eip_usual.Show("新增斡旋失敗!斡旋重覆使用，請先至物件查詢確認!!")
                        End Try
                    Else
                        SqlDataSource1.UpdateParameters.Add("編號", Request("oid"))
                        SqlDataSource1.UpdateParameters.Add("要約編號", Label0.Text)
                        SqlDataSource1.UpdateParameters.Add("物件編號", sp(0))
                        SqlDataSource1.UpdateParameters.Add("要約金", price.Text)
                        SqlDataSource1.UpdateParameters.Add("要約起", offerdateb.Value)
                        SqlDataSource1.UpdateParameters.Add("要約訖", offerdatee.Value)
                        SqlDataSource1.UpdateParameters.Add("輸入者", Request.Cookies("webfly_empno").Value)
                        SqlDataSource1.UpdateParameters.Add("店代號", Request("sid"))
                        SqlDataSource1.UpdateParameters.Add("修改日期", eip_usual.toROCyear(Now()))
                        SqlDataSource1.UpdateParameters.Add("斡旋金", offerprice.Text)
                        SqlDataSource1.UpdateParameters.Add("票據面額", offertick.Text)
                        If RadioButton1.Checked = True Then
                            SqlDataSource1.UpdateParameters.Add("提供個資", "Y")
                        Else
                            SqlDataSource1.UpdateParameters.Add("提供個資", "N")
                        End If
                        different_state = ""
                        selstr = "update " & b.tablename & " set 物件編號=@物件編號,要約金=@要約金,要約起=@要約起,要約訖=@要約訖,輸入者=@輸入者,店代號=@店代號,修改日期=@修改日期,斡旋金=@斡旋金,票據面額=@票據面額,提供個資=@提供個資 where " & b.cumname1 & "=@編號 and 要約編號= @要約編號"
                        SqlDataSource1.UpdateCommand = selstr
                        Try
                            SqlDataSource1.Update()

                            '1050506表單控管
                            'If Request.Cookies("webfly_empno").Value = "92H" Then
                            If check_formno(offerSuffix, sid.Text, sp(0), sp(1), "修改") = "use" Then
                                eip_usual.Show("一併寫入表單管理修")
                            Else
                                'eip_usual.Show("輸入編號為1050515前出貨之表單或不為所購買表單，無法寫入表單管理")
                            End If
                            'End If

                            Label0.Text = ""
                            clear斡旋()
                            eip_usual.Show("修改斡旋成功!")
                        Catch ex As Exception
                            'Response.Write(ex.ToString)
                            eip_usual.Show("修改斡旋失敗!")
                        End Try
                    End If
                Next
            Else
                eip_usual.Show("請先選擇客戶資料!!")
                Exit Sub
            End If
            Load_Data斡旋("OLD")
            Me.Label464.Visible = False
            Me.RadioButton1.Visible = False
            Me.RadioButton2.Visible = False
            Me.RadioButton1.Checked = False
            Me.RadioButton2.Checked = False
        Else
            eip_usual.Show("目前無斡旋新增修改權限，無法維護任何資料。")
        End If

    End Sub

    Sub Load_Data斡旋(ByVal cls As String)
        Dim b As customertype
        For Each b In 客戶類別()
            Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                conn.Open()
                If cls = "OLD" Then
                    sql = "Select 要約.*,物.建築名稱,isnull(要約.提供個資,'N') as 提供個資New From " & b.tablename & " AS 要約 INNER JOIN " & b.tablename1 & " AS 物 ON 要約.物件編號 = 物.物件編號 Where " & b.cumname1 & " = '" & Request("oid") & "' AND 要約.店代號='" & Request("sid") & "'"
                    'sql = "Select 要約.*,物.建築名稱 From " & b.tablename & " AS 要約 INNER JOIN " & b.tablename1 & " AS 物 ON 要約.物件編號 = 物.物件編號 AND 物.店代號='" & Request.Cookies("store_id").Value & "' Where " & b.cumname1 & " = '" & Request("oid") & "'"
                Else '修改
                    Dim 編號1 As Array = Split(cls, ",")
                    sql = "Select 要約.*,物.建築名稱,isnull(要約.提供個資,'N') as 提供個資New From " & b.tablename & " AS 要約 INNER JOIN " & b.tablename1 & " AS 物 ON 要約.物件編號 = 物.物件編號 Where 要約.要約編號 = '" & 編號1(0) & "' and " & b.cumname1 & "='" & Request("oid") & "'  AND 要約.店代號='" & Request("sid") & "'"
                End If
                '//== 最後，合併成完整的SQL指令（搜尋引擎～專用） ==
                '**********************************************************************
                Dim currentSortColumn, currentSortDirection As String
                currentSortColumn = Me.SortExpression.Split(" "c)(0)
                currentSortDirection = Me.SortExpression.Split(" "c)(1)
                'sql &= " order by " & currentSortColumn & " " & currentSortDirection
                sql &= " order by " & "convert(varchar(255)," & currentSortColumn & ")" & " " & currentSortDirection
                '**********************************************************************
                'If Request.Cookies("webfly_empno").Value = "0265" Then
                'Response.Write(sql)
                '    'Response.End()
                'End If

                adpt1 = New SqlDataAdapter(sql, conn)
                ds1 = New DataSet()
                adpt1.Fill(ds1, "細項內容")
                Dim tb_細項內容 As DataTable = ds1.Tables("細項內容")

                If cls = "OLD" Then
                    If tb_細項內容.Rows.Count > 0 And myobj.Brw = "1" Then
                        del.Visible = True
                        Clear.Visible = True
                        Me.GridView1.DataSource = tb_細項內容
                        Me.GridView1.DataBind()
                    Else
                        del.Visible = False
                        Clear.Visible = False
                        Me.GridView1.DataBind()
                        If myobj.Brw = "0" Then
                            eip_usual.Show("目前無斡旋瀏覽權限，無法查看任何資料。")
                        Else
                            eip_usual.Show("目前的查詢條件，搜尋不到任何資料。")
                        End If

                    End If
                Else
                    Dim offerSuffix As String = Right(tb_細項內容.Rows(0)("要約編號").ToString, Len(tb_細項內容.Rows(0)("要約編號").ToString) - 5) 'tb_細項內容.Rows(0)("要約編號").ToString
                    SetOfferIdSelection(offerSuffix, True)
                    objectname.Text = tb_細項內容.Rows(0)("物件編號").ToString & "," & tb_細項內容.Rows(0)("建築名稱").ToString
                    offerdateb.Value = tb_細項內容.Rows(0)("要約起").ToString
                    offerdatee.Value = tb_細項內容.Rows(0)("要約訖").ToString
                    price.Text = tb_細項內容.Rows(0)("要約金").ToString
                    offerprice.Text = tb_細項內容.Rows(0)("斡旋金").ToString
                    offertick.Text = tb_細項內容.Rows(0)("票據面額").ToString
                    Label0.Text = cls

                    If GetOfferIdSuffix().ToUpper >= "JAB40001" And GetOfferIdSuffix().ToUpper <= "JZZ99999" Then
                        Me.Label464.Visible = True
                        Me.RadioButton1.Visible = True
                        Me.RadioButton2.Visible = True
                        Me.RadioButton1.Checked = False
                        Me.RadioButton2.Checked = False
                        If tb_細項內容.Rows(0)("提供個資").ToString = "Y" Then
                            Me.RadioButton1.Checked = True
                        Else
                            Me.RadioButton2.Checked = True
                        End If
                    Else
                        Me.Label464.Visible = False
                        Me.RadioButton1.Visible = False
                        Me.RadioButton2.Visible = False
                        Me.RadioButton1.Checked = False
                        Me.RadioButton2.Checked = False
                    End If
                End If
            End Using
        Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim gridView As GridView = DirectCast(sender, GridView)
        Dim currentSortColumn, currentSortDirection As String
        currentSortColumn = Me.SortExpression.Split(" "c)(0)
        currentSortDirection = Me.SortExpression.Split(" "c)(1)

        If e.Row.RowType = DataControlRowType.Header Then
            Dim cellIndex As Integer = -1
            For Each field As DataControlField In gridView.Columns
                If field.SortExpression = currentSortColumn Then
                    cellIndex = gridView.Columns.IndexOf(field)
                End If
            Next
            If cellIndex > -1 Then
                '  this is a header row, set the sort style
                e.Row.Cells(cellIndex).CssClass = If(currentSortDirection = "ASC", "sortasc", "sortdesc")
            End If
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then

            '要約LABEL
            Dim lbl流水號 As Label = e.Row.FindControl("Label10")

            '編輯button
            Dim btn編輯 As ImageButton = e.Row.FindControl("Button5")
            btn編輯.CommandArgument = lbl流水號.Text

        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "edits" Then
            Load_Data斡旋(e.CommandArgument)
        End If
    End Sub

    '清除
    Protected Sub clear_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Clear.Click
        Dim isEditMode As Boolean = (Not String.IsNullOrWhiteSpace(Label0.Text)) AndAlso (Label0.Text <> "no") AndAlso (Label0.Text <> "OLD")
        clear斡旋(isEditMode)
    End Sub
    Sub clear斡旋(Optional ByVal keepOfferId As Boolean = False)
        If Not keepOfferId Then
            '新增/非編輯：清空並可選擇
            If offerid.Items.Count = 0 Then BindOfferIdDropDown(sid.Text)
            offerid.ClearSelection()
            If offerid.Items.FindByValue("") IsNot Nothing Then
                offerid.SelectedValue = ""
            ElseIf offerid.Items.Count > 0 Then
                offerid.SelectedIndex = 0
            End If
            offerid.Enabled = True
            hf_offerid.Value = ""
        Else
            '編輯模式：保留既有選擇且不可更動（用 HiddenField 保留值）
            hf_offerid.Value = GetOfferIdSuffix()
            offerid.Enabled = False
        End If

        objectname.Text = "" '物件編號(案名)
        offerdateb.Value = eip_usual.toROCyear(Now())
        offerdatee.Value = ""
        price.Text = ""
        offerprice.Text = ""
        offertick.Text = ""
    End Sub

    '刪除帶看
    Protected Sub del_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles del.Click
        If myobj.D = "1" Then
            Dim script As String = ""
            Dim selstr As String = ""
            Dim i As Integer = 0
            Dim dels As Integer = 0
            Dim b As customertype
            For Each b In 客戶類別()
                Using conn As New SqlConnection(EGOUPLOADSqlConnStr)
                    For i = 0 To Me.GridView1.Rows.Count - 1

                        '判斷刪除CHK有選取時執行刪除動作
                        If CType(Me.GridView1.Rows(i).FindControl("chkSelect1"), CheckBox).Checked Then

                            '刪除單筆
                            Dim num As String = CType(Me.GridView1.Rows(i).FindControl("Label10"), Label).Text
                            Dim oidname As String = CType(Me.GridView1.Rows(i).FindControl("Label11"), Label).Text '物件案名
                            SqlDataSource1.DeleteParameters.Add("序號" & i, num)
                            selstr = "Delete from " & b.tablename & " where 要約編號=@序號" & i & " and 店代號='" & Request("sid") & "' and " & b.cumname1 & "='" & Request("oid") & "'"
                            SqlDataSource1.DeleteCommand = selstr
                            SqlDataSource1.Delete()
                            dels += 1

                            '1050506表單控管
                            'If Request.Cookies("webfly_empno").Value = "92H" Then
                            If check_formno(Mid(num, 6, Len(num) - 1), Request("sid"), oid, oidname, "刪除") = "use" Then
                                eip_usual.Show("一併寫入表單管理刪")
                            Else
                                'eip_usual.Show("輸入編號為1050515前出貨之表單或不為所購買表單，無法寫入表單管理")
                            End If
                            'End If
                        End If
                    Next
                End Using
            Next
            eip_usual.Show("刪除斡旋成功!")
            '重整GRIDVIEW
            Load_Data斡旋("OLD")
        Else
            eip_usual.Show("目前無斡旋刪除權限，無法刪除任何資料。")
        End If

    End Sub
    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        'Response.Write("<span class=""inpage_page_font_02"">" & (CInt(e.NewPageIndex) + 1) & "</span>")
        'Response.Write("目前位於第" & (CInt(e.NewPageIndex) + 1) & "頁<br>")
        '//== 把 GridView1 的 [EnableSortingAndPagingCallBack]屬性關閉(=False)，才會執行到這一行！ ==

        '重整GRIDVIEW
        Load_Data斡旋("OLD")
    End Sub


    Protected Sub GridView1_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles GridView1.Sorting
        Dim currentSortColumn, currentSortDirection As String
        currentSortColumn = Me.SortExpression.Split(" "c)(0)
        currentSortDirection = Me.SortExpression.Split(" "c)(1)
        If e.SortExpression.Equals(currentSortColumn) Then
            ' switch sort direction
            Select Case currentSortDirection.ToUpper
                Case "ASC"
                    Me.SortExpression = currentSortColumn & " DESC"
                Case "DESC"
                    Me.SortExpression = currentSortColumn & " ASC"
            End Select
        Else
            Me.SortExpression = e.SortExpression & " ASC"
        End If
        '重整GRIDVIEW
        Load_Data斡旋("OLD")
    End Sub


    Public Property SortExpression As String
        Get
            If ViewState("SortExpression") Is Nothing Then
                ViewState("SortExpression") = "要約.新增日期 ASC"
            End If
            Return ViewState("SortExpression").ToString
        End Get
        Set(ByVal value As String)
            ViewState("SortExpression") = value
        End Set
    End Property

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Tagstr As String
        If Request("object_fun") = "sellcustomer" Or Request("object_fun") = "othercustomer" Then
            Tagstr = "<script>var objwin1=window.open('choose.aspx?object_fun=sellcustomer','Pn_Part_Query','resizable=yes,scrollbars=yes,top=200,left=240,height=600,width=720,status=no,toolbar=no,menubar=no,location=no','');</script>"
        Else
            Tagstr = "<script>var objwin1=window.open('choose.aspx?object_fun=rentcustomer','Pn_Part_Query','resizable=yes,scrollbars=yes,top=200,left=240,height=600,width=720,status=no,toolbar=no,menubar=no,location=no','');</script>"
        End If

        ScriptManager.RegisterClientScriptBlock(Me, GetType(String), "", Tagstr, False)
    End Sub

    '1090407修正
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
            If fun = "customer_offer.aspx" Then '要約書
                sql &= "(select 店代號, 要約編號 as 編號,上傳註記 from 秘書_要約管制檔 where 店代號='" & sid & "' and 要約編號='" & formno & "' and 上傳註記<>'D' union all select 店代號,斡旋要約編號 as 編號,上傳註記 from 秘書_斡旋要約管制檔 where 店代號='" & sid & "' and 斡旋要約編號='" & formno & "' and 上傳註記<>'D')"
                If Request("object_fun") = "sellcustomer" Then '1232.1204 售
                    If Left(formno, 1) = "I" Then
                        status = "OK"
                        sqlm &= "秘書_斡旋要約管制檔_管理 "
                        sqlu &= "秘書_斡旋要約管制檔 set "
                    ElseIf Left(formno, 1) = "K" Then
                        status = "OK"
                        sqlm &= "秘書_要約管制檔_管理 "
                        sqlu &= "秘書_要約管制檔 set "
                    Else
                        status = "要約"
                    End If
                Else '1206 租
                    If Left(formno, 1) = "H" Then
                        status = "OK"
                        sqlm &= "秘書_要約管制檔_管理 "
                        sqlu &= "秘書_要約管制檔 set "
                    Else
                        status = "要約"
                    End If
                End If
                If status = "OK" Then
                    sqlm &= " (類別, 店代號, 編號, 人員, 日期, 備註, 異動人,物件編號) VALUES "
                    If state = "新增" Or state = "修改" Then
                        sqlm &= "('使用','"
                        sqlu &= " 處理類別='使用', "
                    ElseIf state = "刪除" Then
                        sqlm &= "('繳回','"
                        sqlu &= " 處理類別='作廢', "
                    End If
                    sqlm &= sid & "','" & formno & "','" & CType(Me.Customer_detail1.FindControl("names0"), Label).Text & "','" & sysdate & "','客戶:" & CType(Me.Customer_detail1.FindControl("names"), Label).Text & ";物件:" & oidname & ";要約" & state & "後直接寫入,使用店代號為:" & sid & "','" & Request.Cookies("webfly_empno").Value & "','" & oid & "')"
                    sqlu &= " 經紀人代號='" & CType(Me.Customer_detail1.FindControl("names0"), Label).Text & "', 上傳註記='U', 備註='要約" & state & "後直接寫入,使用店代號為:" & sid & "', 修改日期='" & sysdate & "' where 店代號='" & sid & "' and "
                    If Left(formno, 1) = "K" Or Left(formno, 1) = "H" Then
                        sqlu &= "要約編號='" & formno & "'"
                    Else 'I
                        sqlu &= "斡旋要約編號='" & formno & "'"
                    End If
                Else '非出貨要約書編號
                    sqlm = ""
                    sqlu = ""
                End If
            ElseIf fun = "customer_mediation.aspx" Then '斡旋1232.1225
                If Request("object_fun") = "sellcustomer" Then '1232.1204 售
                    sql &= "(select 店代號, 斡旋編號 as 編號 from 秘書_斡旋管制檔 where 店代號='" & sid & "' and 斡旋編號='" & formno & "' and 上傳註記<>'D' union all select 店代號,斡旋要約編號 as 編號 from 秘書_斡旋要約管制檔 where 店代號='" & sid & "' and 斡旋要約編號='" & formno & "' and 上傳註記<>'D')"
                    If Left(formno, 1) = "I" Then
                        status = "OK"
                        sqlm &= "秘書_斡旋要約管制檔_管理 "
                        sqlu &= "秘書_斡旋要約管制檔 set "
                    ElseIf Left(formno, 1) = "J" Then
                        status = "OK"
                        sqlm &= "秘書_斡旋管制檔_管理 "
                        sqlu &= "秘書_斡旋管制檔 set "
                    Else
                        status = "斡旋"
                    End If
                Else '1206 租
                    sql &= "(select 店代號, 要約編號 as 編號 from 秘書_要約管制檔 where 店代號='" & sid & "' and 要約編號='" & formno & "' and 上傳註記<>'D')"
                    If Left(formno, 1) = "H" Then
                        status = "OK"
                        sqlm &= "秘書_要約管制檔_管理 "
                        sqlu &= "秘書_要約管制檔 set "
                    Else
                        status = "斡旋"
                    End If
                End If
                If status = "OK" Then
                    sqlm &= " (類別, 店代號, 編號, 人員, 日期, 備註, 異動人,物件編號) VALUES "
                    If state = "新增" Or state = "修改" Then
                        sqlm &= "('使用','"
                        sqlu &= " 處理類別='使用', "
                    ElseIf state = "刪除" Then
                        sqlm &= "('繳回','"
                        sqlu &= " 處理類別='作廢', "
                    End If
                    sqlm &= sid & "','" & formno & "','" & CType(Me.Customer_detail1.FindControl("names0"), Label).Text & "','" & sysdate & "','客戶:" & CType(Me.Customer_detail1.FindControl("names"), Label).Text & ";物件:" & oidname & ";斡旋" & state & "後直接寫入,使用店代號為:" & sid & "','" & Request.Cookies("webfly_empno").Value & "','" & oid & "')"
                    sqlu &= " 經紀人代號='" & CType(Me.Customer_detail1.FindControl("names0"), Label).Text & "', 上傳註記='U', 備註='斡旋" & state & "後直接寫入,使用店代號為:" & sid & "', 修改日期='" & sysdate & "' where 店代號='" & sid & "' "
                    If Left(formno, 1) = "H" Then
                        sqlu &= " and 要約編號='" & formno & "'"
                    ElseIf Left(formno, 1) = "J" Then
                        sqlu &= " and 斡旋編號='" & formno & "'"
                    End If
                Else '非出貨斡旋書編號
                    sqlm = ""
                    sqlu = ""
                End If

            End If
            sql &= " as a "
            'Response.Write("<br>" & sql & "<br>")
            'Response.Write("<br>" & sqlu & "<br>")
            adpt1 = New SqlDataAdapter(sql, conn)
            ds1 = New DataSet()
            adpt1.Fill(ds1, "table1")
            Dim table1 As DataTable = ds1.Tables("table1")
            'Response.Write(sql & "<br>sql筆數:" & table1.Rows.Count & "<br>")
            If table1.Rows.Count > 0 Then '沒有被刪除
                If Request("object_fun") = "sellcustomer" Then '1232.1204 售
                    If Left(table1.Rows(0)("編號").ToString, 1) = "I" Or Left(table1.Rows(0)("編號").ToString, 1) = "K" Or Left(table1.Rows(0)("編號").ToString, 1) = "J" Then
                        Dim com As New SqlCommand(sqlm, conn)
                        Try
                            com.ExecuteNonQuery()

                        Catch ex As Exception
                            'Response.Write("<br>sqlm:" & ex.ToString & "<br>")
                            'Response.Write(selstr)
                        End Try

                        Dim com1 As New SqlCommand(sqlu, conn)
                        Try
                            com1.ExecuteNonQuery()

                        Catch ex As Exception
                            'Response.Write("<br>sqlu:" & ex.ToString & "<br>")
                        End Try
                        used = "use"
                    Else
                        used = "no"
                    End If
                Else
                    If Left(table1.Rows(0)("編號").ToString, 1) = "H" Then
                        Dim com As New SqlCommand(sqlm, conn)
                        Try
                            com.ExecuteNonQuery()

                        Catch ex As Exception
                            'Response.Write("<br>sqlm:" & ex.ToString & "<br>")
                            'Response.Write(selstr)
                        End Try

                        Dim com1 As New SqlCommand(sqlu, conn)
                        Try
                            com1.ExecuteNonQuery()

                        Catch ex As Exception
                            'Response.Write("<br>sqlu:" & ex.ToString & "<br>")
                        End Try
                        used = "use"
                    Else
                        used = "no"
                    End If
                End If

                'used = "<br>管理:" & sqlm & "<br>管制:" & sqlu & "<br>"
                'Return status & "<br>管理:" & sqlm & "<br>管制:" & sqlu & "<br>"
            Else '管制檔裡沒有
                used = "no"
                'Return status & "<br>管理:" & sqlm & "<br>管制:" & sqlu & "<br>"
            End If
        End Using
        Return used
    End Function

    Protected Sub offerid_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles offerid.SelectedIndexChanged
        Dim offerSuffix As String = GetOfferIdSuffix()
        hf_offerid.Value = offerSuffix

        If offerSuffix.ToUpper >= "JAB40001" And offerSuffix.ToUpper <= "JZZ99999" Then
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
    End Sub
End Class