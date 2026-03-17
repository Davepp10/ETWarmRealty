Imports System.Data
Imports clspowerset
Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Imports System.Net
Imports System.IO

Partial Class G_WebService_vacancy_search
    Inherits System.Web.UI.Page
    Dim mysql As String = ConfigurationManager.ConnectionStrings("mysqletwarm").ConnectionString

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            If Request.Cookies("webfly_empno").Value Is Nothing Then
                Response.Write("Cookie expired")
                Response.End()
            End If
            'If Request.Cookies("webfly_empno").Value = "92H" Then
            Label1.Text = "<li id=""current""><a href=""/C_ShopKeeping/1111.aspx"" >相關文件</a></li>"
            'End If
            DropDownList1 = clspowerset.scope(Request.Cookies("webfly_empno").Value, "store", DropDownList1) '店代號
            DropDownList1.SelectedValue = Request.Cookies("store_id").Value

            bindDt()

        End If
    End Sub
    Sub bindDt()
        GridView2.Visible = False
        '" & Request.Cookies("store_id").Value & "
        Dim selsql As String = "select * from vacancy where vacancy_storeid = '" & DropDownList1.SelectedValue & "'"
        If TextBox1.Text <> "" Then
            selsql += " and vacancy_name like '%" & eip_usual.replaceUnsafeStr(TextBox1.Text) & "%'"
        End If
        If TextBox2.Text <> "" And TextBox3.Text <> "" Then
            selsql += " and (vacancy_deadline between '" & eip_usual.replaceUnsafeStr(TextBox2.Text) & "' and '" & eip_usual.replaceUnsafeStr(TextBox3.Text) & "')"
        End If
        'If Request.Cookies("store_id").Value <> "A0001" Then
        selsql += " and vacancy_del != 'y'"
        'End If
        selsql += "  order by vacancy_maketime desc "
        SqlDataSource1.SelectCommand = selsql
        GridView1.DataSourceID = "SqlDataSource1"
        GridView1.DataBind()
        Label3.Text = GridView1.Rows.Count

        Dim selstr As String = " SELECT * FROM recruitment "
        selstr += " where referral_to = '" & Request.Cookies("store_id").Value & "' and del = 'n' order by create_at desc"

        SqlDataSource3.SelectCommand = selstr
        SqlDataSource3.DataBind()
        GridView3.DataSourceID = "SqlDataSource3"
        GridView3.DataBind()
        GridView3.Visible = True
    End Sub
    Protected Sub GridView1_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Dele" Then

            Dim updatestr As String = "update vacancy set vacancy_del = 'y',vacancy_deltime = NOW(),vacancy_delmember = '" & Request.Cookies("webfly_empno").Value & "' where vacancy_oid = '" & e.CommandArgument & "'"
            Response.Write(updatestr)
            Using conn As New MySqlConnection(mysql)
                conn.Open()
                Using cmd As New MySqlCommand(updatestr, conn)
                    cmd.ExecuteNonQuery()
                    eip_usual.Show("刪除完成")
                    bindDt()
                End Using
            End Using
        End If

        If e.CommandName = "seedetail" Then
            Dim selstr As String = "SELECT * FROM vacancy_bio_history a "
            selstr += " left join vacancy_bio b on a.bioid = b.vacancy_bio_oid "
            selstr += " where a.jid = '" & e.CommandArgument & "' and a.del = 'n' and IFNULL(b.vacancy_bio_del,'')<>'y' order by bio_date desc"
            'eip_usual.Show(e.CommandArgument)
            'Exit Sub
            SqlDataSource2.SelectCommand = selstr
            SqlDataSource2.DataBind()
            GridView2.DataSourceID = "SqlDataSource2"
            GridView2.DataBind()
            GridView2.Visible = True
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow And (e.Row.RowState = DataControlRowState.Normal Or e.Row.RowState = DataControlRowState.Alternate) Then
            Dim lb1 As Label = e.Row.Cells(2).FindControl("Label1")
            lb1.Text = e.Row.DataItem("vacancy_maketime").ToString.Substring(0, 10)
            Dim lb2 As Label = e.Row.Cells(3).FindControl("Label2")
            lb2.Text = e.Row.DataItem("vacancy_deadline").ToString


            '應徵筆數
            Dim lb3 As Label = e.Row.Cells(4).FindControl("Label3")
            lb3.Text = "0"
            Using conn As New MySqlConnection(mysql)
                conn.Open()
                Using cmd As New MySqlCommand("SELECT count(*) FROM etwarm.vacancy_bio_history a left join etwarm.vacancy_bio b on a.bioid = b.vacancy_bio_oid where a.jid = '" & e.Row.DataItem("vacancy_oid") & "' and a.del = 'n' and IFNULL(b.vacancy_bio_del,'')<>'y' ", conn)
                    Dim dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    lb3.Text = dt.Rows(0)(0)

                    Dim btdetail As Button = e.Row.Cells(5).FindControl("Button2")
                    If lb3.Text <> "0" Then
                        btdetail.Visible = True
                    Else
                        btdetail.Visible = False
                    End If
                End Using
            End Using
        End If
    End Sub

    Protected Sub GridView2_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView2.RowCommand
        If e.CommandName = "Dele" Then

            'eip_usual.Show("刪除完成" & e.CommandArgument)

            Dim updatestr As String = "update vacancy_bio set vacancy_bio_del='y' where vacancy_bio_oid = '" & e.CommandArgument & "'"
            Response.Write(updatestr)
            Using conn As New MySqlConnection(mysql)
                conn.Open()
                Using cmd As New MySqlCommand(updatestr, conn)
                    cmd.ExecuteNonQuery()
                    eip_usual.Show("刪除完成" & e.CommandArgument)
                    bindDt()
                End Using
            End Using
        End If

        'If e.CommandName = "seedetail" Then
        '    Dim selstr As String = "SELECT * FROM vacancy_bio_history a "
        '    selstr += " left join vacancy_bio b on a.bioid = b.vacancy_bio_oid "
        '    selstr += " where a.jid = '" & e.CommandArgument & "' and a.del = 'n' order by bio_date desc"
        '    eip_usual.Show(e.CommandArgument)
        '    Exit Sub
        '    SqlDataSource2.SelectCommand = selstr
        '    SqlDataSource2.DataBind()
        '    GridView2.DataSourceID = "SqlDataSource2"
        '    GridView2.DataBind()
        '    GridView2.Visible = True
        'End If
    End Sub

    Protected Sub Button1_Click(sender As Object, e As ImageClickEventArgs) Handles Button1.Click
        bindDt()
    End Sub

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then '排除標頭的判斷
            Dim HyLnk1 As HyperLink = CType(e.Row.Cells(0).FindControl("HyperLink1"), HyperLink)
            Dim ID As Label = CType(e.Row.Cells(0).FindControl("Label5"), Label)
            Dim store_report As Label = CType(e.Row.Cells(0).FindControl("Label6"), Label)
            'HyLnk1.Text = "<a href='#' onClick=""GB_showCenter('商品訂購', '../vacancy/vacancy_biodetailNew.aspx?id=" & CType(e.Row.Cells(0).FindControl("id"), Label).text & "&dlr=" & CType(e.Row.Cells(0).FindControl("store_report"), Label).text & "',560,1100)"";><img src=""../images/comstore_bt_01.gif""/></a>"
            HyLnk1.Text = "<a href='#' onClick=""GB_showCenter('查看', 'https://superwebnew.etwarm.com.tw/C_ShopKeeping/vacancy/vacancy_biodetailNew.aspx?id=" & ID.text & "&dlr=" & store_report.text & "',560,1100)"";><img src=""../images/land_bt_01.gif""/></a>"
            'Response.Write(ID.text & "_" & store_report.text)
        End If
    End Sub

    Protected Sub GridView3_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView3.RowCommand
        If e.CommandName = "Del" Then
            Dim updatestr As String = "update recruitment set del = 'y' where id = '" & e.CommandArgument & "'"
            Using conn As New MySqlConnection(mysql)
                conn.Open()
                Using cmd As New MySqlCommand(updatestr, conn)
                    cmd.ExecuteNonQuery()
                    eip_usual.Show("刪除完成")
                    bindDt()
                End Using
            End Using
        End If
    End Sub
End Class
