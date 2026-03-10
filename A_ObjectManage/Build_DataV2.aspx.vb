Imports System.Data
Imports System.Data.SqlClient
Imports clspowerset
Partial Class A_ObjectManage_Build_Data
    Inherits System.Web.UI.Page

    '權限CLASS
    Dim myobj As New clspowerset


    '連線字串
    Dim EGOUPLOADSqlConnStr As String = ConfigurationManager.ConnectionStrings("EGOUPLOADConnectionString").ConnectionString

    '連線CODE用-MSSQL
    Dim i As Integer = 0
    Dim sql, sql2 As String
    Dim table1, table2, table3, table4 As DataTable

    Public conn, conn1, conn_land As SqlConnection
    Public cmd, cmd2 As SqlCommand
    Public ds As DataSet
    Public adpt As SqlDataAdapter

    Dim sysdate As String = Right("000" & Year(Now) - 1911, 3) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)
    Dim checkleave As New checkleave

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        TextBox253.Attributes("readonly") = "readonly"

        If Not Page.IsPostBack Then
            'URL傳過來的參數
            'Gridview的值
            sid.Text = Request("sid")
            oid.Text = Request("oid")
            NUM.Text = Request("NUM")

            '物件編號
            Me.Label11.Text = Request("oid")
            '店代號
            Me.Label12.Text = Request("sid")


            '控制項的值
            usid.Text = Request("usid")
            uoid.Text = Request("uoid")


            get_Data()
        End If


    End Sub

    Sub get_Data()

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        '管理組織及管理方式       
        'sql = "select * from 資料_不動產說明書 With(NoLock) where 類別='管理組織及管理方式 ' and 店代號 in ('A0001','" & Request.Cookies("store_id").Value & "')"
        'adpt = New SqlDataAdapter(sql, conn)
        'ds = New DataSet()
        'adpt.Fill(ds, "table1")
        'table1 = ds.Tables("table1")
        'DropDownList34.Items.Clear()
        'i = 0
        'For i = 0 To table1.Rows.Count - 1
        '    DropDownList34.Items.Add(table1.Rows(i)("名稱").ToString.Trim)
        'Next
        'DropDownList34.Items.Add("其他")

        Dim sql2 As String = "select * from 委賣物件資料表_面積細項 A With(NoLock) left join 產調_建物  B on A.物件編號=B.物件編號 and A.店代號=B.店代號 and A.流水號=B.流水號 where A.物件編號 = '" & oid.Text & "' and A.店代號='" & sid.Text & "' and A.流水號='" & NUM.Text & "'  "
        Dim cmd2 As New SqlCommand(sql2, conn)
        cmd2.CommandType = CommandType.Text
        adpt = New SqlDataAdapter(sql2, conn)
        ds = New DataSet()
        adpt.Fill(ds, "table2")
        table2 = ds.Tables("table2")

        '地建號
        If Not IsDBNull(table2.Rows(0)("建號")) Then
            Label5.Text = table2.Rows(0)("建號")
        End If


        If IsDBNull(table2.Rows(0)("新增者")) Then
            'If myobj.AC = "1" Then
            ImageButton1.Visible = True
            'Else
            '    ImageButton1.Visible = False
            'End If
            ImageButton12.Visible = False
            ImageButton19.Visible = False
        Else
            ImageButton1.Visible = False
            'If myobj.M = "1" Then
            ImageButton12.Visible = True
            'Else
            '    ImageButton12.Visible = False
            'End If
            'If myobj.D = "1" Then
            ImageButton19.Visible = True
            'Else
            '    ImageButton19.Visible = False
            'End If

            '-------------------------------------------------------------個案權利說明-------------------------------------------------------------
            '建物是否共有
            If Not IsDBNull(table2.Rows(0)("是否共有")) Then
                DropDownList58.SelectedValue = table2.Rows(0)("是否共有")

                If DropDownList58.SelectedValue = "無" Then
                    PlaceHolder19.Visible = False

                Else
                    PlaceHolder19.Visible = True

                    If Not IsDBNull(table2.Rows(0)("有無分管協議登記")) Then
                        RadioButtonList4.SelectedValue = table2.Rows(0)("有無分管協議登記")
                    End If
                    If Not IsDBNull(table2.Rows(0)("有無分管協議登記_說明")) Then
                        TextBox68.Text = table2.Rows(0)("有無分管協議登記_說明")
                    End If

                End If
            Else
                DropDownList58.SelectedValue = "無"
            End If

            '限制登記
            If Not IsDBNull(table2.Rows(0)("限制登記")) Then
                DropDownList41.SelectedValue = table2.Rows(0)("限制登記")

                If DropDownList41.SelectedValue = "無" Then
                    Label38.Visible = False
                    TextBox246.Text = ""
                    TextBox246.Visible = False
                Else
                    Label38.Visible = True
                    TextBox246.Visible = True
                    If Not IsDBNull(table2.Rows(0)("限制登記_說明")) Then
                        TextBox246.Text = table2.Rows(0)("限制登記_說明")
                    Else
                        TextBox246.Text = ""
                    End If
                End If
            Else
                DropDownList41.SelectedValue = "無"
            End If

            '信託登記
            If Not IsDBNull(table2.Rows(0)("信託登記")) Then
                DropDownList29.SelectedValue = table2.Rows(0)("信託登記")

                If DropDownList29.SelectedValue = "無" Then
                    Label39.Visible = False
                    TextBox247.Text = ""
                    TextBox247.Visible = False
                Else
                    Label39.Visible = True
                    TextBox247.Visible = True
                    If Not IsDBNull(table2.Rows(0)("信託登記_說明")) Then
                        TextBox247.Text = table2.Rows(0)("信託登記_說明")
                    Else
                        TextBox247.Text = ""
                    End If
                End If
            Else
                DropDownList29.SelectedValue = "無"
            End If

            '其他事項(其他權利)
            If Not IsDBNull(table2.Rows(0)("其他權利")) Then
                DropDownList1.SelectedValue = table2.Rows(0)("其他權利")

                If DropDownList1.SelectedValue = "無" Then
                    Label40.Visible = False
                    TextBox248.Text = ""
                    TextBox248.Visible = False
                Else
                    Label40.Visible = True
                    TextBox248.Visible = True
                    If Not IsDBNull(table2.Rows(0)("其他權利_說明")) Then
                        TextBox248.Text = table2.Rows(0)("其他權利_說明")
                    Else
                        TextBox248.Text = ""
                    End If
                End If
            Else
                DropDownList1.SelectedValue = "無"
            End If



            '獎勵容積開放空間提供公共使用
            If Not IsDBNull(table2.Rows(0)("獎勵容積開放空間提供公共使用")) Then
                DropDownList36.SelectedValue = table2.Rows(0)("獎勵容積開放空間提供公共使用")

                If DropDownList36.SelectedValue = "無" Then
                    Label17.Visible = False
                    TextBox236.Text = ""
                    TextBox236.Visible = False
                Else
                    Label17.Visible = True
                    TextBox236.Visible = True
                    If Not IsDBNull(table2.Rows(0)("獎勵容積開放空間提供公共使用_說明")) Then
                        TextBox236.Text = table2.Rows(0)("獎勵容積開放空間提供公共使用_說明")
                    Else
                        TextBox236.Text = ""
                    End If
                End If
            Else
                DropDownList36.SelectedValue = "無"
            End If



            '屬工業區或其他分區
            If Not IsDBNull(table2.Rows(0)("屬工業區或其他分區")) Then
                DropDownList10.SelectedValue = table2.Rows(0)("屬工業區或其他分區")

                If Not IsDBNull(table2.Rows(0)("屬工業區或其他分區_說明")) Then
                    TextBox245.Visible = True
                    TextBox245.Text = table2.Rows(0)("屬工業區或其他分區_說明")
                End If
            Else
                DropDownList10.SelectedValue = "無"
            End If

            '持有期間有無居住
            'If Not IsDBNull(table2.Rows(0)("持有期間有無居住")) Then
            '    DropDownList58.SelectedValue = table2.Rows(0)("持有期間有無居住")
            'Else
            '    DropDownList58.SelectedValue = "無"
            'End If

            '使用執照有無備註之注意事項
            If Not IsDBNull(table2.Rows(0)("使用執照有無備註之注意事項")) Then
                DropDownList59.SelectedValue = table2.Rows(0)("使用執照有無備註之注意事項")

                If DropDownList59.SelectedValue = "無" Then
                    Label36.Visible = False
                    TextBox241.Text = ""
                    TextBox241.Visible = False
                Else
                    Label36.Visible = True
                    TextBox241.Visible = True
                    If Not IsDBNull(table2.Rows(0)("使用執照有無備註之注意事項_說明")) Then
                        TextBox241.Text = table2.Rows(0)("使用執照有無備註之注意事項_說明")
                    Else
                        TextBox241.Text = ""
                    End If
                End If
            Else
                DropDownList59.SelectedValue = "無"
            End If

            '有無禁建情事
            If Not IsDBNull(table2.Rows(0)("有無禁建情事")) Then
                DropDownList57.SelectedValue = table2.Rows(0)("有無禁建情事")

                If DropDownList57.SelectedValue = "無" Then
                    Label44.Visible = False
                    TextBox21.Text = ""
                    TextBox21.Visible = False
                Else
                    Label44.Visible = True
                    TextBox21.Visible = True
                    If Not IsDBNull(table2.Rows(0)("有無禁建情事_說明")) Then
                        TextBox21.Text = table2.Rows(0)("有無禁建情事_說明")
                    Else
                        TextBox21.Text = ""
                    End If
                End If
            Else
                DropDownList57.SelectedValue = "無"
            End If

            If Not IsDBNull(table2.Rows(0)("中繼幫浦")) Then
                DropDownList61.SelectedValue = table2.Rows(0)("中繼幫浦")
                If DropDownList61.SelectedValue = "無" Then
                    Label45.Visible = False
                    TextBox249.Text = ""
                    TextBox249.Visible = False
                Else
                    Label45.Visible = True
                    TextBox249.Visible = True
                    If Not IsDBNull(table2.Rows(0)("中繼幫浦_說明")) Then
                        TextBox249.Text = table2.Rows(0)("中繼幫浦_說明")
                    Else
                        TextBox249.Text = ""
                    End If
                End If
            Else
                DropDownList61.SelectedValue = "無"
            End If

            '太陽光電發電設備
            If Not IsDBNull(table2.Rows(0)("太陽光電發電設備")) Then
                DropDownList62.SelectedValue = table2.Rows(0)("太陽光電發電設備").ToString()
                If DropDownList62.SelectedValue = "有" Then
                    LabelSolarPos.Visible = True
                    TextBox250.Visible = True
                    If Not IsDBNull(table2.Rows(0)("太陽光電發電設備_設置位置")) Then
                        TextBox250.Text = table2.Rows(0)("太陽光電發電設備_設置位置").ToString()
                    Else
                        TextBox250.Text = ""
                    End If
                Else
                    LabelSolarPos.Visible = False
                    TextBox250.Visible = False
                    TextBox250.Text = ""
                End If
            Else
                DropDownList62.SelectedValue = "無"
                LabelSolarPos.Visible = False
                TextBox250.Visible = False
                TextBox250.Text = ""
            End If

            '建築能效標示
            If Not IsDBNull(table2.Rows(0)("建築能效標示")) Then
                DropDownList63.SelectedValue = table2.Rows(0)("建築能效標示").ToString()
                If DropDownList63.SelectedValue = "有" Then
                    PlaceHolder86.Visible = True
                    DropDownList64.SelectedValue = If(IsDBNull(table2.Rows(0)("建築能效標示_能效等級")), "", table2.Rows(0)("建築能效標示_能效等級").ToString())
                    TextBox253.Text = If(IsDBNull(table2.Rows(0)("建築能效標示_證書效期")), "", table2.Rows(0)("建築能效標示_證書效期").ToString())
                Else
                    PlaceHolder86.Visible = False
                    DropDownList64.SelectedIndex = 0
                    TextBox253.Text = ""
                End If
            Else
                DropDownList63.SelectedValue = "無"
                PlaceHolder86.Visible = False
                DropDownList64.SelectedIndex = 0
                TextBox253.Text = ""
            End If


            '**************** 屋主現況說明部分*****************
            '頂樓基地台,頂樓基地台_說明, "
            If Not IsDBNull(table2.Rows(0)("頂樓基地台")) Then
                DropDownList2.SelectedValue = table2.Rows(0)("頂樓基地台")
                If table2.Rows(0)("頂樓基地台") = "有" Then
                    Label1.Visible = True
                    TextBox1.Visible = True
                End If
                TextBox1.Text = table2.Rows(0)("頂樓基地台_說明")
            Else
                TextBox1.Text = ""
            End If

            '衛生下水道工程, 衛生下水道工程_選項, 衛生下水道工程_說明,"
            If Not IsDBNull(table2.Rows(0)("衛生下水道工程")) Then
                If table2.Rows(0)("衛生下水道工程") = "無" Then
                    DropDownList3.SelectedValue = table2.Rows(0)("衛生下水道工程")
                    DropDownList4.SelectedValue = table2.Rows(0)("衛生下水道工程_選項")
                    TextBox2.Text = table2.Rows(0)("衛生下水道工程_說明")

                    DropDownList4.Visible = True
                    Label2.Visible = True
                    TextBox2.Visible = True
                ElseIf table2.Rows(0)("衛生下水道工程") = "有" Then
                    DropDownList3.SelectedValue = table2.Rows(0)("衛生下水道工程")
                    DropDownList4.Visible = False
                    Label2.Visible = False
                    TextBox2.Visible = False
                End If

            Else
                TextBox2.Text = ""
            End If

            '共有部分範圍_有無, 共有部分範圍, 專有部分範圍_有無, 專有部分範圍
            If Not IsDBNull(table2.Rows(0)("共有部分範圍_有無")) Then
                DropDownList5.SelectedValue = table2.Rows(0)("共有部分範圍_有無")
                TextBox3.Text = table2.Rows(0)("共有部分範圍")
                If table2.Rows(0)("共有部分範圍_有無") = "有" Then
                    PlaceHolder7.Visible = True
                Else
                    PlaceHolder7.Visible = False
                End If

            Else
                PlaceHolder7.Visible = False
                TextBox3.Text = ""
            End If
            If Not IsDBNull(table2.Rows(0)("專有部分範圍_有無")) Then
                DropDownList5.SelectedValue = table2.Rows(0)("專有部分範圍_有無")
                TextBox4.Text = table2.Rows(0)("專有部分範圍")
                If table2.Rows(0)("專有部分範圍_有無") = "有" Then
                    PlaceHolder77.Visible = True
                Else
                    PlaceHolder77.Visible = False
                End If

            Else
                PlaceHolder77.Visible = False
                TextBox4.Text = ""
            End If
            '規約外特殊使用, 規約外特殊使用_共用說明, 規約外特殊使用_專用說明,"
            'If Not IsDBNull(table2.Rows(0)("規約外特殊使用")) Then
            '    DropDownList5.SelectedValue = table2.Rows(0)("規約外特殊使用")
            '    TextBox3.Text = table2.Rows(0)("規約外特殊使用_共用說明")
            '    TextBox4.Text = table2.Rows(0)("規約外特殊使用_專用說明")
            '    If table2.Rows(0)("規約外特殊使用") = "有" Then
            '        PlaceHolder7.Visible = True
            '    Else
            '        PlaceHolder7.Visible = False
            '    End If

            'Else
            '    PlaceHolder7.Visible = False
            '    TextBox3.Text = ""
            '    TextBox4.Text = ""
            'End If

            '管理公司_有無,管理公司,"
            If Not IsDBNull(table2.Rows(0)("管理公司_有無")) Then
                DropDownList6.SelectedValue = table2.Rows(0)("管理公司_有無")
                TextBox5.Text = table2.Rows(0)("管理公司")

                If table2.Rows(0)("管理公司_有無") = "有" Then
                    Label6.Visible = True
                    TextBox5.Visible = True
                End If

            Else
                TextBox5.Text = ""
            End If

            '專有約定共用,專有約定共用之範圍,專有約定共用之使用方式,"
            If Not IsDBNull(table2.Rows(0)("專有約定共用")) Then
                DropDownList8.SelectedValue = table2.Rows(0)("專有約定共用")
                TextBox6.Text = table2.Rows(0)("專有約定共用之範圍")
                TextBox7.Text = table2.Rows(0)("專有約定共用之使用方式")

                PlaceHolder8.Visible = True

            Else

                PlaceHolder8.Visible = False
                TextBox6.Text = ""
                TextBox7.Text = ""
            End If

            '共有約定專用,共有約定專用之範圍,共有約定專用之使用方式,"
            If Not IsDBNull(table2.Rows(0)("共有約定專用")) Then
                DropDownList9.SelectedValue = table2.Rows(0)("共有約定專用")
                TextBox8.Text = table2.Rows(0)("共有約定專用之範圍")
                TextBox9.Text = table2.Rows(0)("共有約定專用之使用方式")
                PlaceHolder9.Visible = True
            Else
                PlaceHolder9.Visible = False
                TextBox8.Text = ""
                TextBox9.Text = ""
            End If

            '公共基金有無,公共基金數額,公共基金提撥方式,公共基金運用方式,"
            If Not IsDBNull(table2.Rows(0)("公共基金有無")) Then
                DropDownList11.SelectedValue = table2.Rows(0)("公共基金有無")
                If table2.Rows(0)("公共基金有無") = "有" Then
                    PlaceHolder5.Visible = True
                    TextBox10.Text = table2.Rows(0)("公共基金數額")
                    If table2.Rows(0)("公共基金提撥方式") = "管理費" Then
                        DropDownList12.SelectedValue = table2.Rows(0)("公共基金提撥方式")
                    Else
                        DropDownList12.SelectedValue = "其他"
                        TextBox11.Text = table2.Rows(0)("公共基金提撥方式")
                    End If

                    If table2.Rows(0)("公共基金運用方式") = "管委會決議" Then
                        DropDownList13.SelectedValue = table2.Rows(0)("公共基金運用方式")
                    Else
                        DropDownList13.SelectedValue = "其他"
                        TextBox12.Text = table2.Rows(0)("公共基金運用方式")
                    End If
                Else
                    PlaceHolder5.Visible = False
                End If
            Else
                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
            End If

            '管理費使用,管理費或使用費,管理費,管理費繳交方式,"
            If Not IsDBNull(table2.Rows(0)("管理費使用")) Then
                If table2.Rows(0)("管理費使用") = "有" Then
                    PlaceHolder6.Visible = True
                    DropDownList14.SelectedValue = "有"
                    DropDownList15.SelectedValue = table2.Rows(0)("管理費或使用費")
                    TextBox14.Text = table2.Rows(0)("管理費")
                    TextBox15.Text = table2.Rows(0)("管理費繳交方式")
                Else
                    PlaceHolder6.Visible = False
                End If
            Else
                TextBox14.Text = ""
                TextBox15.Text = ""
            End If
            '管理組織_有無,管理組織及方式,"
            If Not IsDBNull(table2.Rows(0)("管理組織_有無")) Then
                If table2.Rows(0)("管理組織_有無") = "有" Then
                    DropDownList17.Visible = True
                    TextBox13.Visible = True

                    DropDownList16.SelectedValue = "有"
                    If table2.Rows(0)("管理組織及方式") = "管委會" Then
                        DropDownList17.SelectedValue = table2.Rows(0)("管理組織及方式")
                    Else
                        DropDownList17.SelectedValue = "其他"
                        TextBox13.Text = table2.Rows(0)("管理組織及方式")
                    End If
                Else

                End If
            Else

            End If
            '住戶規約使用手冊,"
            If Not IsDBNull(table2.Rows(0)("住戶規約使用手冊")) Then
                DropDownList18.SelectedValue = table2.Rows(0)("住戶規約使用手冊")

            End If
            '張貼合格標章,張貼合格標章_說明,"
            If Not IsDBNull(table2.Rows(0)("電梯設備")) Then
                If table2.Rows(0)("電梯設備") = "有" Then
                    DropDownList7.SelectedValue = "有"
                    Label7.Visible = True

                    Label16.Visible = True
                    TextBox16.Visible = True
                    DropDownList19.Visible = True
                End If
            End If
            If Not IsDBNull(table2.Rows(0)("張貼合格標章")) Then
                DropDownList19.SelectedValue = table2.Rows(0)("張貼合格標章")
                If DropDownList19.SelectedValue = "有" Then

                    Label16.Visible = False
                    TextBox16.Visible = False
                Else

                    Label16.Visible = True
                    TextBox16.Visible = True
                End If


                If Not IsDBNull(table2.Rows(0)("張貼合格標章_說明")) Then
                    TextBox16.Text = table2.Rows(0)("張貼合格標章_說明")

                End If
            Else

            End If
            '出租狀況,出租範圍,出租範圍備註,出租情況類型,租金,租期起,租期迄,租約是否公證,押租保證金,出租狀況_說明,"
            If Not IsDBNull(table2.Rows(0)("出租狀況")) Then
                DropDownList53.SelectedValue = table2.Rows(0)("出租狀況")
                If table2.Rows(0)("出租狀況") = "有" Then
                    PlaceHolder1.Visible = True
                    If Not IsDBNull(table2.Rows(0)("出租範圍")) Then
                        DropDownList55.SelectedValue = table2.Rows(0)("出租範圍")
                    End If
                    If Not IsDBNull(table2.Rows(0)("出租範圍備註")) Then
                        TextBox56.Text = table2.Rows(0)("出租範圍備註")
                    End If



                    If Not IsDBNull(table2.Rows(0)("出租情況類型")) And table2.Rows(0)("出租情況類型").ToString.Trim <> "" Then

                        Dim 出租情況類型 As Array = Split(table2.Rows(0)("出租情況類型").ToString, ";")

                        For i = 0 To 出租情況類型.Length - 1
                            If 出租情況類型(i).ToString = "不定期租約" Then
                                RadioButton6.Checked = True
                                TextBox55.Text = table2.Rows(0)("租金")
                            End If
                            If 出租情況類型(i).ToString = "定期租約" Then
                                RadioButton7.Checked = True
                                If table2.Rows(0)("租約是否公證") = "有" Or table2.Rows(0)("租約是否公證") = "是" Then
                                    RadioButtonList3.SelectedIndex = 0
                                Else
                                    RadioButtonList3.SelectedIndex = 1
                                End If
                                TextBox59.Text = table2.Rows(0)("租金")
                                TextBox60.Text = table2.Rows(0)("押租保證金")
                                TextBox61.Text = table2.Rows(0)("租期起")
                                TextBox62.Text = table2.Rows(0)("租期迄")
                            End If
                            If 出租情況類型(i).ToString = "租賃之權利義務隨同移轉" Then
                                RadioButton8.Checked = True
                            End If
                            If 出租情況類型(i).ToString = "屋主終止租約騰空交屋" Then
                                RadioButton9.Checked = True
                            End If
                            If 出租情況類型(i).ToString = "其他" Then
                                RadioButton10.Checked = True
                                TextBox58.Text = table2.Rows(0)("出租狀況_說明")
                            End If
                        Next

                    End If

                Else
                    PlaceHolder1.Visible = False
                End If

            Else

            End If
            '出借狀況,出借範圍,出借範圍備註,出借書面約定與否,借用人姓名,出借起日,出借迄日,出借返還條件,"
            If Not IsDBNull(table2.Rows(0)("出借狀況")) Then
                DropDownList54.SelectedValue = table2.Rows(0)("出借狀況")
                If table2.Rows(0)("出借狀況") = "有" Then
                    PlaceHolder2.Visible = True
                    DropDownList56.SelectedValue = table2.Rows(0)("出借範圍")
                    TextBox56.Text = table2.Rows(0)("出借範圍備註")

                    If table2.Rows(0)("出借書面約定與否") = "無書面約定" Then
                        RadioButton11.Checked = True
                    Else
                        RadioButton12.Checked = True
                    End If
                    TextBox63.Text = table2.Rows(0)("借用人姓名")
                    TextBox64.Text = table2.Rows(0)("出借起日")
                    TextBox65.Text = table2.Rows(0)("出借迄日")
                    TextBox66.Text = table2.Rows(0)("出借返還條件")
                Else
                    PlaceHolder2.Visible = False
                End If

            Else

            End If
            '佔用情形,佔用他人建物土地,佔用他人建物土地_說明,"
            If Not IsDBNull(table2.Rows(0)("佔用情形")) Then
                DropDownList20.SelectedValue = table2.Rows(0)("佔用情形")
                If table2.Rows(0)("佔用情形") = "有" Then
                    PlaceHolder13.Visible = True
                Else
                    PlaceHolder13.Visible = False
                End If
                If table2.Rows(0)("佔用他人建物土地") = "建物占用他人土地" Or table2.Rows(0)("佔用他人建物土地") = "建物被他人占用" Then
                    DropDownList21.SelectedValue = table2.Rows(0)("佔用他人建物土地")
                Else
                    DropDownList21.SelectedValue = "其他"
                    TextBox67.Text = table2.Rows(0)("佔用他人建物土地")
                End If
                TextBox17.Text = table2.Rows(0)("佔用他人建物土地_說明")
            Else

            End If
            '消防設備,消防設備_說明,"
            If Not IsDBNull(table2.Rows(0)("消防設備")) Then
                DropDownList22.SelectedValue = table2.Rows(0)("消防設備")
                If table2.Rows(0)("消防設備") = "有" Then
                    CheckBoxList1.Visible = True
                    TextBox18.Visible = True
                Else
                    CheckBoxList1.Visible = False
                    TextBox18.Visible = False
                End If

                If Not IsDBNull(table2.Rows(0)("消防設備_說明")) Then
                    Dim otherstr As String = table2.Rows(0)("消防設備_說明")
                    For Each li As ListItem In CheckBoxList1.Items
                        If table2.Rows(0)("消防設備_說明").ToString.IndexOf(li.Text) >= 0 Then
                            li.Selected = True
                            otherstr = otherstr.Replace(li.Text, "")
                        End If

                    Next
                    otherstr = otherstr.Replace(",,", "")
                    TextBox18.Text = otherstr
                End If
            End If
            '無障礙設施,無障礙設施_說明,"
            If Not IsDBNull(table2.Rows(0)("無障礙設施")) Then
                If table2.Rows(0)("無障礙設施") = "有" Then
                    PlaceHolder10.Visible = True
                Else
                    PlaceHolder10.Visible = False
                End If
                DropDownList23.SelectedValue = table2.Rows(0)("無障礙設施")
                TextBox19.Text = table2.Rows(0)("無障礙設施_說明")


            End If

            '夾層,夾層面積,夾層面積1,夾層其他,"
            If Not IsDBNull(table2.Rows(0)("夾層")) Then
                DropDownList24.SelectedValue = table2.Rows(0)("夾層")
                If table2.Rows(0)("夾層") = "有" Then
                    PlaceHolder11.Visible = True
                    If Not IsDBNull(table2.Rows(0)("夾層面積")) Then
                        TextBox20.Text = table2.Rows(0)("夾層面積")
                    End If
                    If Not IsDBNull(table2.Rows(0)("夾層面積1")) Then
                        TextBox22.Text = table2.Rows(0)("夾層面積1")
                    End If
                    If Not IsDBNull(table2.Rows(0)("夾層其他")) Then
                        TextBox23.Text = table2.Rows(0)("夾層其他")
                    End If



                Else
                    PlaceHolder11.Visible = False
                End If
            Else
                DropDownList24.SelectedValue = "無"
            End If

            '獨立供水,供水類型,供水是否正常,獨立供水_說明,"
            If Not IsDBNull(table2.Rows(0)("獨立供水")) Then
                DropDownList25.SelectedValue = table2.Rows(0)("獨立供水")

                If table2.Rows(0)("獨立供水") = "有" Then
                    PlaceHolder12.Visible = True
                    DropDownList26.SelectedValue = table2.Rows(0)("供水類型")
                    If Not IsDBNull(table2.Rows(0)("供水是否正常")) Then
                        RadioButtonList1.SelectedValue = table2.Rows(0)("供水是否正常")
                    End If
                    TextBox24.Text = table2.Rows(0)("獨立供水_說明")
                Else
                    PlaceHolder12.Visible = False
                End If
            End If
            '獨立電表,獨立電表_說明,"
            If Not IsDBNull(table2.Rows(0)("獨立電表")) Then
                DropDownList27.SelectedValue = table2.Rows(0)("獨立電表")
                If table2.Rows(0)("獨立電表") = "無" Then
                    TextBox25.Text = table2.Rows(0)("獨立電表_說明")

                    Label21.Visible = True
                    TextBox25.Visible = True
                Else
                    Label21.Visible = False
                    TextBox25.Visible = False
                End If
            End If

            '天然瓦斯,天然瓦斯_說明,天然瓦斯_說明2,"
            If Not IsDBNull(table2.Rows(0)("天然瓦斯")) Then
                DropDownList28.SelectedValue = table2.Rows(0)("天然瓦斯")
                If table2.Rows(0)("天然瓦斯") = "有" Then
                    RadioButtonList2.SelectedValue = table2.Rows(0)("天然瓦斯_說明")
                    TextBox26.Text = IIf(IsDBNull(table2.Rows(0)("天然瓦斯_說明2")), "", table2.Rows(0)("天然瓦斯_說明2"))

                    RadioButtonList2.Visible = True
                    TextBox26.Visible = True
                Else
                    RadioButtonList2.Visible = False
                    TextBox26.Visible = False
                End If
            End If
            '水電管線是否更新,水管更新日期,電線更新日期,"
            If Not IsDBNull(table2.Rows(0)("水電管線是否更新")) Then
                DropDownList30.SelectedValue = table2.Rows(0)("水電管線是否更新")
                If table2.Rows(0)("水電管線是否更新") = "有" Then
                    PlaceHolder14.Visible = True
                    TextBox27.Text = table2.Rows(0)("水管更新日期")
                    TextBox28.Text = table2.Rows(0)("電線更新日期")
                Else
                    PlaceHolder14.Visible = False
                End If
            End If
            '積欠應繳費用,積欠應繳費用_說明,"
            If Not IsDBNull(table2.Rows(0)("積欠應繳費用")) Then
                DropDownList31.SelectedValue = table2.Rows(0)("積欠應繳費用")
                If table2.Rows(0)("積欠應繳費用") = "有" Then
                    Label3.Visible = True
                    TextBox29.Visible = True
                    TextBox29.Text = table2.Rows(0)("積欠應繳費用_說明")
                Else
                    Label3.Visible = False
                    TextBox29.Visible = False
                End If
            End If
            '持有期間有無居住,"
            If Not IsDBNull(table2.Rows(0)("持有期間有無居住")) Then
                DropDownList32.SelectedValue = table2.Rows(0)("持有期間有無居住")

            End If
            '有無公共設施重大修繕,有無公共設施重大修繕_說明,有無公共設施重大修繕_金額,"
            If Not IsDBNull(table2.Rows(0)("有無公共設施重大修繕")) Then
                DropDownList33.SelectedValue = table2.Rows(0)("有無公共設施重大修繕")
                If table2.Rows(0)("有無公共設施重大修繕") = "有" Then
                    TextBox30.Text = table2.Rows(0)("有無公共設施重大修繕_說明")
                    TextBox31.Text = table2.Rows(0)("有無公共設施重大修繕_金額")
                    PlaceHolder15.Visible = True
                Else
                    PlaceHolder15.Visible = False
                End If
            End If
            '混凝土中氯離子含量,混凝土中氯離子含量_說明,"
            If Not IsDBNull(table2.Rows(0)("混凝土中氯離子含量")) Then
                DropDownList34.SelectedValue = table2.Rows(0)("混凝土中氯離子含量")

                If table2.Rows(0)("混凝土中氯離子含量_說明") <> "屋主未檢測" Then
                    DropDownList35.SelectedValue = "其他"
                    TextBox33.Text = table2.Rows(0)("混凝土中氯離子含量_說明")
                Else
                    DropDownList35.SelectedIndex = 0
                End If

            End If
            '輻射檢測,輻射檢測_說明,"
            If Not IsDBNull(table2.Rows(0)("輻射檢測")) Then
                DropDownList37.SelectedValue = table2.Rows(0)("輻射檢測")

                If table2.Rows(0)("輻射檢測") = "無" Then
                    If table2.Rows(0)("輻射檢測_說明") <> "屋主未檢測" Then
                        DropDownList38.SelectedValue = "其他"
                        TextBox33.Text = table2.Rows(0)("輻射檢測_說明")
                    Else
                        DropDownList38.SelectedIndex = 0
                    End If
                Else
                    Label9.Text = "(附檢測結果)"
                    DropDownList38.Visible = False
                    TextBox33.Visible = False
                End If
            End If
            '曾發生火災或其他災害,曾發生火災或其他災害_說明,"
            If Not IsDBNull(table2.Rows(0)("曾發生火災或其他災害")) Then
                DropDownList39.SelectedValue = table2.Rows(0)("曾發生火災或其他災害")
                If table2.Rows(0)("曾發生火災或其他災害") = "有" Then
                    TextBox34.Text = table2.Rows(0)("曾發生火災或其他災害_說明")
                    Label22.Visible = True
                    TextBox34.Visible = True
                Else
                    Label22.Visible = False
                    TextBox34.Visible = False
                End If
            End If
            '因地震被公告為危險建築,因地震被公告為危險建築_說明,"
            If Not IsDBNull(table2.Rows(0)("因地震被公告為危險建築")) Then
                DropDownList40.SelectedValue = table2.Rows(0)("因地震被公告為危險建築")
                If table2.Rows(0)("因地震被公告為危險建築") = "有" Then
                    TextBox35.Text = table2.Rows(0)("因地震被公告為危險建築_說明")

                    Label23.Visible = True
                    TextBox35.Visible = True
                Else

                    Label23.Visible = False
                    TextBox35.Visible = False
                End If
            End If
            '樑柱部分是否有顯見裂痕,樑柱部分是否有顯見裂痕_說明,裂痕長度,間隙寬度,"
            If Not IsDBNull(table2.Rows(0)("樑柱部分是否有顯見裂痕")) Then
                DropDownList42.SelectedValue = table2.Rows(0)("樑柱部分是否有顯見裂痕")
                If table2.Rows(0)("樑柱部分是否有顯見裂痕") = "有" Then
                    TextBox36.Text = table2.Rows(0)("樑柱部分是否有顯見裂痕_說明")
                    TextBox37.Text = table2.Rows(0)("裂痕長度")
                    TextBox38.Text = table2.Rows(0)("間隙寬度")

                    PlaceHolder16.Visible = True
                Else
                    PlaceHolder16.Visible = False
                End If
            End If
            '建物鋼筋裸露,建物鋼筋裸露_說明,"
            If Not IsDBNull(table2.Rows(0)("建物鋼筋裸露")) Then
                DropDownList43.SelectedValue = table2.Rows(0)("建物鋼筋裸露")
                If table2.Rows(0)("建物鋼筋裸露") = "有" Then
                    TextBox39.Text = table2.Rows(0)("建物鋼筋裸露_說明")

                    Label27.Visible = True
                    TextBox39.Visible = True
                Else
                    Label27.Visible = False
                    TextBox39.Visible = False
                End If
            End If

            '是否為兇宅,兇宅發生期間,是否為兇宅_說明,"
            If Not IsDBNull(table2.Rows(0)("是否為兇宅")) Then
                DropDownList44.SelectedValue = table2.Rows(0)("是否為兇宅")
                If table2.Rows(0)("是否為兇宅") = "有" Then
                    If table2.Rows(0)("兇宅發生期間") = "產權持有期間:" Then
                        RadioButton1.Checked = True
                        TextBox40.Text = table2.Rows(0)("是否為兇宅_說明")
                    End If
                    If table2.Rows(0)("兇宅發生期間") = "產權持有期間前,賣方:" Then
                        RadioButton2.Checked = True
                        TextBox41.Text = table2.Rows(0)("是否為兇宅_說明")
                    End If
                    If table2.Rows(0)("兇宅發生期間") = "確認無發生過" Then
                        RadioButton3.Checked = True
                    End If
                    If table2.Rows(0)("兇宅發生期間") = "知道曾發生過" Then
                        RadioButton4.Checked = True
                    End If
                    If table2.Rows(0)("是否為兇宅") = "其他:" Then
                        RadioButton5.Checked = True
                        TextBox42.Text = table2.Rows(0)("是否為兇宅_說明")
                    End If

                    PlaceHolder17.Visible = True
                Else
                    PlaceHolder17.Visible = False
                End If
            End If
            '滲漏水狀態,滲漏水狀態_說明,滲漏水狀態_處理,"
            If Not IsDBNull(table2.Rows(0)("滲漏水狀態")) Then
                DropDownList45.SelectedValue = table2.Rows(0)("滲漏水狀態")
                If table2.Rows(0)("滲漏水狀態") = "有" Then
                    Dim otherStr As String = table2.Rows(0)("滲漏水狀態_說明").ToString
                    For Each lit As ListItem In CheckBoxList2.Items
                        If table2.Rows(0)("滲漏水狀態_說明").ToString.IndexOf(lit.Text) >= 0 Then
                            otherStr = otherStr.Replace(lit.Text & ";", "")
                            lit.Selected = True
                        End If
                    Next
                    TextBox43.Text = otherStr

                    If Not IsDBNull(table2.Rows(0)("滲漏水狀態_處理")) Then
                        If table2.Rows(0)("滲漏水狀態_處理") = "賣方修繕後交屋" Or table2.Rows(0)("滲漏水狀態_處理") = "買方自行修繕" Or table2.Rows(0)("滲漏水狀態_處理") = "減價" Then
                            DropDownList46.SelectedValue = table2.Rows(0)("滲漏水狀態_處理")
                        Else
                            DropDownList46.SelectedValue = "其他"
                            TextBox44.Text = table2.Rows(0)("滲漏水狀態_處理")
                        End If
                    End If


                    PlaceHolder4.Visible = True
                Else
                    PlaceHolder4.Visible = False
                End If
            End If
            '違增建使用權,違增建使用權_說明,違增建_面積,違增建列管情形,違增建列管情形_說明,"
            If Not IsDBNull(table2.Rows(0)("違增建使用權")) Then
                DropDownList47.SelectedValue = table2.Rows(0)("違增建使用權")
                If table2.Rows(0)("違增建使用權") = "有" Then
                    Dim otherStr As String = table2.Rows(0)("違增建使用權_說明").ToString
                    For Each lit As ListItem In CheckBoxList3.Items
                        If table2.Rows(0)("違增建使用權_說明").ToString.IndexOf(lit.Text) >= 0 Then
                            otherStr = otherStr.Replace(lit.Text & ";", "")
                            lit.Selected = True
                        End If
                    Next
                    TextBox45.Text = otherStr

                    If Not IsDBNull(table2.Rows(0)("違增建_面積").ToString) Then
                        TextBox47.Text = table2.Rows(0)("違增建_面積").ToString.Trim
                    Else
                        TextBox47.Text = ""
                    End If

                    If Not IsDBNull(table2.Rows(0)("違增建列管情形")) Then
                        DropDownList48.SelectedValue = table2.Rows(0)("違增建列管情形")
                    Else
                        DropDownList48.SelectedValue = "無"
                    End If
                    If Not IsDBNull(table2.Rows(0)("違增建列管情形_說明")) Then
                        TextBox46.Text = table2.Rows(0)("違增建列管情形_說明")
                    Else
                        TextBox46.Text = ""
                    End If


                    PlaceHolder3.Visible = True
                Else
                    PlaceHolder3.Visible = False
                End If
            End If
            '排水系統,排水系統_說明,"
            If Not IsDBNull(table2.Rows(0)("排水系統")) Then
                DropDownList49.SelectedValue = table2.Rows(0)("排水系統")
                If table2.Rows(0)("排水系統") = "無" Then
                    DropDownList50.SelectedValue = table2.Rows(0)("排水系統_說明")

                    Label4.Visible = True
                    DropDownList50.Visible = True
                Else
                    Label4.Visible = False
                    DropDownList50.Visible = False
                End If

            End If
            '隨附設備有無,隨附設備,"
            If Not IsDBNull(table2.Rows(0)("隨附設備有無")) Then
                DropDownList51.SelectedValue = table2.Rows(0)("隨附設備有無")
                If table2.Rows(0)("隨附設備有無") = "有" Then
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox1.Text) >= 0 Then
                        CheckBox1.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox2.Text) >= 0 Then
                        CheckBox2.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox3.Text) >= 0 Then
                        CheckBox3.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox4.Text) >= 0 Then
                        CheckBox4.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox5.Text) >= 0 Then
                        CheckBox5.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox6.Text) >= 0 Then
                        CheckBox6.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox7.Text) >= 0 Then
                        CheckBox7.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox8.Text) >= 0 Then
                        CheckBox8.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox9.Text) >= 0 Then
                        CheckBox9.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox10.Text) >= 0 Then
                        CheckBox10.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox11.Text) >= 0 Then
                        CheckBox11.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox12.Text) >= 0 Then
                        CheckBox12.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox13.Text) >= 0 Then
                        CheckBox13.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox14.Text) >= 0 Then
                        CheckBox14.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox15.Text) >= 0 Then
                        CheckBox15.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox16.Text) >= 0 Then
                        CheckBox16.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox17.Text) >= 0 Then
                        CheckBox17.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox18.Text) >= 0 Then
                        CheckBox18.Checked = True
                    End If
                    If table2.Rows(0)("隨附設備").ToString.IndexOf(CheckBox19.Text) >= 0 Then
                        CheckBox19.Checked = True
                    End If

                    '沙發數,電視數,冰箱數,冷氣數,洗衣機數,乾衣機數,
                    TextBox48.Text = table2.Rows(0)("沙發數").ToString
                    TextBox49.Text = table2.Rows(0)("電視數").ToString
                    TextBox50.Text = table2.Rows(0)("冰箱數").ToString
                    TextBox51.Text = table2.Rows(0)("冷氣數").ToString
                    TextBox52.Text = table2.Rows(0)("洗衣機數").ToString
                    TextBox53.Text = table2.Rows(0)("乾衣機數").ToString

                    PlaceHolder18.Visible = True
                Else
                    PlaceHolder18.Visible = False
                End If

            End If

            '全棟
            If Not IsDBNull(table2.Rows(0)("全棟")) Then
                If table2.Rows(0)("全棟") = "Y" Then
                    house.Checked = True
                End If
            End If

            '其他重要事項,其他重要事項_說明,"
            If Not IsDBNull(table2.Rows(0)("其他重要事項")) Then
                DropDownList52.SelectedValue = table2.Rows(0)("其他重要事項")
                If table2.Rows(0)("其他重要事項") = "有" Then
                    TextBox54.Text = table2.Rows(0)("其他重要事項_說明")

                    Label28.Visible = True
                    TextBox54.Visible = True
                Else
                    Label28.Visible = False
                    TextBox54.Visible = False
                End If
            End If

        End If

        conn.Close()
        conn.Dispose()
    End Sub

    '新增
    Protected Sub ImageButton1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click

        If DropDownList62.SelectedValue = "有" AndAlso String.IsNullOrWhiteSpace(TextBox250.Text) Then
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v1", "alert('「設置位置」為必填');", True)
            Exit Sub
        End If

        If DropDownList63.SelectedValue = "有" Then
            If String.IsNullOrWhiteSpace(DropDownList64.SelectedValue) Then
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v2", "alert('「能效等級」為必填，且不可選「-請選擇能效等級-」');", True)
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(TextBox253.Text) Then
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v3", "alert('「有效期間」為必填');", True)
                Exit Sub
            End If
        End If

        '如果編號跟店代號跟預設值不一樣，先下列步驟
        If oid.Text <> uoid.Text Or sid.Text <> usid.Text Then
            Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
            Dim sql_upt As String = "update 產調_建物 set 物件編號='" & uoid.Text & "',店代號='" & usid.Text & "' where 物件編號='" & oid.Text & "' and 店代號='" & sid.Text & "'"
            Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
            cmd_upt.CommandType = CommandType.Text

            conn_upt.Open()

            cmd_upt.ExecuteNonQuery()

            conn_upt.Close()
            conn_upt.Dispose()

            'UPDATE已存資料後給予新的值
            '物件編號
            Me.Label11.Text = uoid.Text

            '店代號
            Me.Label12.Text = usid.Text

        End If

        conn = New SqlConnection(EGOUPLOADSqlConnStr)


        Dim sql As String = ""
        sql = "insert into 產調_建物(物件編號,流水號,店代號,"
        sql &= "是否共有,有無分管協議登記,有無分管協議登記_說明,"
        sql &= "限制登記,限制登記_說明,信託登記,信託登記_說明,其他權利,其他權利_說明,"
        sql &= "獎勵容積開放空間提供公共使用,獎勵容積開放空間提供公共使用_說明,"
        sql &= "屬工業區或其他分區, 屬工業區或其他分區_說明, "
        sql &= "使用執照有無備註之注意事項, 使用執照有無備註之注意事項_說明, "
        sql &= "有無禁建情事,有無禁建情事_說明, "
        sql &= "中繼幫浦,中繼幫浦_說明, "
        ''******** 太陽光電 建築能效********
        sql &= "太陽光電發電設備,太陽光電發電設備_設置位置,"
        sql &= "建築能效標示,建築能效標示_能效等級,建築能效標示_證書效期,"
        ''******** 屋主現況說明部分 ********
        sql &= "頂樓基地台,頂樓基地台_說明, "
        sql &= "衛生下水道工程, 衛生下水道工程_選項, 衛生下水道工程_說明,"
        'sql &= "規約外特殊使用, 規約外特殊使用_共用說明, 規約外特殊使用_專用說明," <-- 改用下列
        sql &= "共有部分範圍_有無, 共有部分範圍, 專有部分範圍_有無, 專有部分範圍,"
        sql &= "管理公司_有無,管理公司,"
        sql &= "專有約定共用,專有約定共用之範圍,專有約定共用之使用方式,"
        sql &= "共有約定專用,共有約定專用之範圍,共有約定專用之使用方式,"
        sql &= "公共基金有無,公共基金數額,公共基金提撥方式,公共基金運用方式,"
        sql &= "管理費使用,管理費或使用費,管理費,管理費繳交方式,"
        sql &= "管理組織_有無,管理組織及方式,"
        sql &= "住戶規約使用手冊,"
        sql &= "電梯設備,張貼合格標章,張貼合格標章_說明,"
        sql &= "出租狀況,出租範圍,出租範圍備註,出租情況類型,租金,租期起,租期迄,租約是否公證,押租保證金,出租狀況_說明,"
        sql &= "出借狀況,出借範圍,出借範圍備註,出借書面約定與否,借用人姓名,出借起日,出借迄日,出借返還條件,"
        sql &= "佔用情形,佔用他人建物土地,佔用他人建物土地_說明,"
        sql &= "消防設備,消防設備_說明,"
        sql &= "無障礙設施,無障礙設施_說明,"
        sql &= "夾層,夾層面積,夾層面積1,夾層其他,"
        sql &= "獨立供水,供水類型,供水是否正常,獨立供水_說明,"
        sql &= "獨立電表,獨立電表_說明,"
        sql &= "天然瓦斯,天然瓦斯_說明,天然瓦斯_說明2,"
        sql &= "水電管線是否更新,水管更新日期,電線更新日期,"
        sql &= "積欠應繳費用,積欠應繳費用_說明,"
        sql &= "持有期間有無居住,"
        sql &= "有無公共設施重大修繕,有無公共設施重大修繕_說明,有無公共設施重大修繕_金額,"
        sql &= "混凝土中氯離子含量,混凝土中氯離子含量_說明,"
        sql &= "輻射檢測,輻射檢測_說明,"
        sql &= "曾發生火災或其他災害,曾發生火災或其他災害_說明,"
        sql &= "因地震被公告為危險建築,因地震被公告為危險建築_說明,"
        sql &= "樑柱部分是否有顯見裂痕,樑柱部分是否有顯見裂痕_說明,裂痕長度,間隙寬度,"
        sql &= "建物鋼筋裸露,建物鋼筋裸露_說明,"
        sql &= "是否為兇宅,兇宅發生期間,是否為兇宅_說明,"
        '---------- OK -------
        sql &= "滲漏水狀態,滲漏水狀態_說明,滲漏水狀態_處理,"
        sql &= "違增建使用權,違增建使用權_說明,違增建_面積,違增建列管情形,違增建列管情形_說明,"
        sql &= "排水系統,排水系統_說明,"
        sql &= "隨附設備有無,隨附設備,沙發數,電視數,冰箱數,冷氣數,洗衣機數,乾衣機數,全棟,"
        sql &= "其他重要事項,其他重要事項_說明,"
        sql &= "新增日期,新增者) values "
        sql &= "('" & Label11.Text & "','" & NUM.Text & "','" & Label12.Text & "',"
        sql &= "'" & DropDownList58.SelectedValue & "','" & RadioButtonList4.SelectedValue & "','" & TextBox68.Text & "',"

        sql &= "'" & DropDownList41.SelectedValue & "','" & TextBox246.Text & "','" & DropDownList29.SelectedValue & "','" & TextBox247.Text & "','" & DropDownList1.SelectedValue & "','" & TextBox248.Text & "',"
        '"獎勵容積開放空間提供公共使用,獎勵容積開放空間提供公共使用_說明,"
        sql &= "'" & DropDownList36.SelectedValue & "','" & TextBox236.Text & "',"
        '"屬工業區或其他分區, 屬工業區或其他分區_說明, "
        sql &= "'" & DropDownList10.SelectedValue & "','" & TextBox245.Text & "',"
        '"使用執照有無備註之注意事項, 使用執照有無備註之注意事項_說明, "
        sql &= "'" & DropDownList59.SelectedValue & "','" & TextBox241.Text & "',"
        '"有無禁建情事,有無禁建情事_說明"
        sql &= "'" & DropDownList57.SelectedValue & "','" & TextBox21.Text & "',"
        sql &= "'" & DropDownList61.SelectedValue & "','" & TextBox249.Text & "',"
        ''******** 太陽光電 建築能效********
        sql &= "'" & DropDownList62.SelectedValue & "',"
        If DropDownList62.SelectedValue = "有" Then
            sql &= "'" & TextBox250.Text & "',"
        Else
            sql &= "'',"
        End If

        sql &= "'" & DropDownList63.SelectedValue & "',"
        If DropDownList63.SelectedValue = "有" Then
            sql &= "'" & DropDownList64.SelectedValue & "','" & TextBox253.Text & "',"
        Else
            sql &= "'','',"
        End If
        ''******** 屋主現況說明部分 ********
        '"頂樓基地台,頂樓基地台_說明, "
        sql &= "'" & DropDownList2.SelectedValue & "','" & TextBox1.Text & "',"
        '"衛生下水道工程, 衛生下水道工程_選項, 衛生下水道工程_說明,"
        sql &= "'" & DropDownList3.SelectedValue & "','" & DropDownList4.SelectedValue & "','" & TextBox2.Text & "',"
        '"規約外特殊使用, 規約外特殊使用_共用說明, 規約外特殊使用_專用說明,"
        'sql &= "'" & DropDownList5.SelectedValue & "','" & TextBox3.Text & "','" & TextBox4.Text & "',"
        ''共有部分範圍_有無, 共有部分範圍, 專有部分範圍_有無, 專有部分範圍
        sql &= "'" & DropDownList5.SelectedValue & "','" & TextBox3.Text & "','" & DropDownList60.SelectedValue & "','" & TextBox4.Text & "',"
        '"管理公司_有無,管理公司,"
        sql &= "'" & DropDownList6.SelectedValue & "','" & TextBox5.Text & "',"
        '"專有約定共用,專有約定共用之範圍,專有約定共用之使用方式,"
        sql &= "'" & DropDownList8.SelectedValue & "','" & TextBox6.Text & "','" & TextBox7.Text & "',"
        '"共有約定專用,共有約定專用之範圍,共有約定專用之使用方式,"
        sql &= "'" & DropDownList9.SelectedValue & "','" & TextBox8.Text & "','" & TextBox9.Text & "',"
        '"公共基金有無,"
        sql &= "'" & DropDownList11.SelectedValue & "',"

        If DropDownList11.SelectedValue = "有" Then
            ' 公共基金數額
            sql &= "'" & TextBox10.Text & "',"
            '公共基金提撥方式
            If DropDownList12.SelectedIndex <> 0 Then
                sql &= "'" & TextBox11.Text & "',"
            Else
                sql &= "'" & DropDownList12.SelectedValue & "',"
            End If
            '公共基金運用方式
            If DropDownList13.SelectedIndex <> 0 Then
                sql &= "'" & TextBox12.Text & "',"
            Else
                sql &= "'" & DropDownList13.SelectedValue & "',"
            End If
        Else
            '無公共基金時， 公共基金數額,公共基金提撥方式,公共基金運用方式
            sql &= "'--','無','無',"
        End If
        '"管理費使用,管理費或使用費,管理費,管理費繳交方式,"
        sql &= "'" & DropDownList14.SelectedValue & "',"
        If DropDownList14.SelectedValue = "有" Then
            sql &= "'" & DropDownList15.SelectedValue & "','" & TextBox14.Text & "','" & TextBox15.Text & "',"

        Else
            sql &= "'無','0','無',"
        End If
        '"管理組織_有無,管理組織及方式,"
        sql &= "'" & DropDownList16.SelectedValue & "',"
        If DropDownList16.SelectedValue = "有" Then
            If DropDownList17.SelectedValue <> "其他" Then
                sql &= "'" & DropDownList17.SelectedValue & "',"
            Else
                sql &= "'" & TextBox13.Text & "',"
            End If
        Else
            sql &= "'無',"
        End If

        '住戶規約使用手冊
        sql &= "'" & DropDownList18.SelectedValue & "',"

        '電梯設備,張貼合格標章,張貼合格標章_說明,"
        sql &= "'" & DropDownList7.SelectedValue & "','" & DropDownList19.SelectedValue & "','" & TextBox16.Text & "',"
        '"出租狀況,出租範圍,出租範圍備註,
        sql &= "'" & DropDownList53.SelectedValue & "','" & DropDownList55.SelectedValue & "','" & TextBox57.Text & "',"
        '出租情況類型
        If RadioButton6.Checked = True Then
            sql &= "'不定期租約',"
            ',租金
            sql &= "" & IIf(TextBox55.Text.Length = 0, 0, TextBox55.Text) & ","
            ',租期起,租期迄,租約是否公證,押租保證金"
            sql &= "'','','','',"
        ElseIf RadioButton7.Checked = True Then
            sql &= "'定期租約',"
            ',租金
            sql &= "" & TextBox59.Text & ","
            ',租期起,租期迄,租約是否公證,押租保證金"
            sql &= "'" & TextBox61.Text & "','" & TextBox62.Text & "','" & RadioButtonList3.SelectedValue & "','" & TextBox60.Text & "',"
        ElseIf RadioButton8.Checked = True Then
            sql &= "'" & RadioButton8.Text & "',"
            ',租金
            sql &= "0,"
            ',租期起,租期迄,租約是否公證,押租保證金"
            sql &= "'','','','',"
        ElseIf RadioButton9.Checked = True Then
            sql &= "'" & RadioButton9.Text & "',"
            ',租金
            sql &= "0,"
            ',租期起,租期迄,租約是否公證,押租保證金"
            sql &= "'','','','',"
        ElseIf RadioButton10.Checked = True Then
            sql &= "'其他',"
            ',租金
            sql &= "0,"
            ',租期起,租期迄,租約是否公證,押租保證金"
            sql &= "'','','','',"
        Else
            sql &= "'',0,'','','','',"
        End If
        '出租狀況_說明
        sql &= "'" & TextBox58.Text & "',"

        '"出借狀況,出借範圍,出借範圍備註
        sql &= "'" & DropDownList54.SelectedValue & "','" & DropDownList56.SelectedValue & "','" & TextBox56.Text & "',"
        ',出借書面約定與否,
        If RadioButton11.Checked = True Then
            sql &= "'" & RadioButton11.Text & "',"
            '借用人姓名,出借起日,出借迄日,出借返還條件,"
            sql &= "'','','','',"
        Else
            sql &= "'" & RadioButton12.Text & "',"
            '借用人姓名,出借起日,出借迄日,出借返還條件,"
            sql &= "'" & TextBox63.Text & "','" & TextBox64.Text & "','" & TextBox65.Text & "','" & TextBox66.Text & "',"
        End If
        '"佔用情形
        sql &= "'" & DropDownList20.SelectedValue & "',"
        ',佔用他人建物土地,
        If DropDownList21.SelectedValue <> "其他" Then
            sql &= "'" & DropDownList21.SelectedValue & "',"
        Else
            sql &= "'" & TextBox67.Text & "',"
        End If
        '佔用他人建物土地_說明,"
        sql &= "'" & TextBox17.Text & "',"
        '"消防設備,
        sql &= "'" & DropDownList22.SelectedValue & "',"
        '消防設備_說明,"
        Dim tempdescript As String = ""
        For Each lit As ListItem In CheckBoxList1.Items
            If lit.Selected = True Then
                tempdescript += lit.Text & ","
            End If
        Next
        If TextBox18.Text.Length > 0 Then
            tempdescript += TextBox18.Text & ","
        End If

        sql &= "'" & tempdescript & "',"
        '"無障礙設施,無障礙設施_說明,"
        sql &= "'" & DropDownList23.SelectedValue & "',"
        sql &= "'" & TextBox19.Text & "',"
        '"夾層,夾層面積,夾層面積1,夾層其他,"
        sql &= "'" & DropDownList24.SelectedValue & "',"
        If TextBox20.Text <> "" Then
            sql &= "'" & TextBox20.Text & "',"
        Else
            sql &= "'0',"
        End If
        If TextBox22.Text <> "" Then
            sql &= "'" & TextBox22.Text & "',"
        Else
            sql &= "'0',"
        End If
        '水電管線是否更新,水管更新日期,電線更新日期,
        sql &= "'" & TextBox23.Text & "',"
        '"獨立供水,供水類型,供水是否正常,獨立供水_說明,"
        sql &= "'" & DropDownList25.SelectedValue & "','" & DropDownList26.SelectedValue & "','" & RadioButtonList1.SelectedValue & "','" & TextBox24.Text & "',"
        '"獨立電表,獨立電表_說明,"
        sql &= "'" & DropDownList27.SelectedValue & "','" & TextBox25.Text & "',"
        '天然瓦斯,天然瓦斯_說明,天然瓦斯_說明2,"
        sql &= "'" & DropDownList28.SelectedValue & "','" & RadioButtonList2.SelectedValue & "','" & TextBox26.Text & "',"
        sql &= "'" & DropDownList30.SelectedValue & "','" & TextBox27.Text & "','" & TextBox28.Text & "',"
        '積欠應繳費用,積欠應繳費用_說明
        sql &= "'" & DropDownList31.SelectedValue & "','" & TextBox29.Text & "',"
        '持有期間有無居住
        sql &= "'" & DropDownList32.SelectedValue & "',"
        '有無公共設施重大修繕,有無公共設施重大修繕_說明,有無公共設施重大修繕_金額
        sql &= "'" & DropDownList33.SelectedValue & "','" & TextBox30.Text & "','" & TextBox31.Text & "',"
        '混凝土中氯離子含量,
        sql &= "'" & DropDownList34.SelectedValue & "',"
        '混凝土中氯離子含量_說明
        If DropDownList35.SelectedValue <> "其他" Then
            sql &= "'" & DropDownList35.SelectedValue & "',"
        Else
            sql &= "'" & TextBox32.Text & "',"
        End If
        '"輻射檢測,輻射檢測_說明,"
        sql &= "'" & DropDownList37.SelectedValue & "',"
        If DropDownList38.SelectedValue <> "其他" Then
            sql &= "'" & DropDownList38.SelectedValue & "',"
        Else
            sql &= "'" & TextBox33.Text & "',"
        End If
        '曾發生火災或其他災害,曾發生火災或其他災害_說明
        sql &= "'" & DropDownList39.SelectedValue & "','" & TextBox34.Text & "',"
        '因地震被公告為危險建築,因地震被公告為危險建築_說明
        sql &= "'" & DropDownList40.SelectedValue & "','" & TextBox35.Text & "',"
        '"樑柱部分是否有顯見裂痕,樑柱部分是否有顯見裂痕_說明,裂痕長度,間隙寬度,"
        sql &= "'" & DropDownList42.SelectedValue & "','" & TextBox36.Text & "','" & TextBox37.Text & "','" & TextBox38.Text & "',"
        '"建物鋼筋裸露,建物鋼筋裸露_說明,"
        sql &= "'" & DropDownList43.SelectedValue & "','" & TextBox39.Text & "',"
        '"是否為兇宅,
        sql &= "'" & DropDownList44.SelectedValue & "',"
        '兇宅發生期間,是否為兇宅_說明,"
        If RadioButton1.Checked = True Then
            sql &= "'" & RadioButton1.Text & "','" & TextBox40.Text & "',"
        ElseIf RadioButton2.Checked = True Then
            sql &= "'" & RadioButton2.Text & "','" & TextBox41.Text & "',"
        ElseIf RadioButton3.Checked = True Then
            sql &= "'" & RadioButton3.Text & "','',"
        ElseIf RadioButton4.Checked = True Then
            sql &= "'" & RadioButton4.Text & "','',"
        ElseIf RadioButton5.Checked = True Then
            sql &= "'" & RadioButton5.Text & "','" & TextBox42.Text & "',"
        Else
            sql &= "'','',"
        End If
        '"滲漏水狀態,滲漏水狀態_說明,"
        sql &= "'" & DropDownList45.SelectedValue & "',"
        Dim postemp As String = TextBox43.Text & ";"
        For Each lis As ListItem In CheckBoxList2.Items
            If lis.Selected = True Then
                postemp += lis.Text & ";"
            End If
        Next
        If postemp = ";" Then postemp = ""
        sql &= "'" & postemp & "',"
        '滲漏水狀態_處理
        If DropDownList46.SelectedValue <> "其他" Then
            sql &= "'" & DropDownList46.SelectedValue & "',"
        Else
            sql &= "'" & TextBox44.Text & "',"
        End If

        '"違增建使用權,,
        sql &= "'" & DropDownList47.SelectedValue & "',"
        If DropDownList47.SelectedValue = "有" Then
            '違增建使用權_說明
            postemp = ""
            For Each lis As ListItem In CheckBoxList3.Items
                If lis.Selected = True Then
                    postemp += lis.Text & ";"
                End If
            Next
            If TextBox45.Text.Length > 0 Then
                postemp += TextBox45.Text & ";"
            End If
            sql &= "'" & postemp & "',"
            '違增建_面積,
            sql &= "'" & TextBox47.Text & "',"
            '違增建列管情形,違增建列管情形_說明,"
            sql &= "'" & DropDownList48.SelectedValue & "','" & TextBox46.Text & "',"
        Else
            sql &= "'','','','',"
        End If

        '排水系統,排水系統_說明
        sql &= "'" & DropDownList49.SelectedValue & "',"
        If DropDownList49.SelectedValue = "無" Then
            sql &= "'" & DropDownList50.SelectedValue & "',"
        Else
            sql &= "'',"
        End If
        '隨附設備有無,"隨附設備,
        sql &= "'" & DropDownList51.SelectedValue & "',"
        If DropDownList51.SelectedValue = "有" Then
            Dim 隨附設備 As String = ""
            If CheckBox1.Checked = True Then
                隨附設備 += CheckBox1.Text & ";"
            End If
            If CheckBox2.Checked = True Then
                隨附設備 += CheckBox2.Text & ";"
            End If
            If CheckBox3.Checked = True Then
                隨附設備 += CheckBox3.Text & ";"
            End If
            If CheckBox4.Checked = True Then
                隨附設備 += CheckBox4.Text & ";"
            End If
            If CheckBox5.Checked = True Then
                隨附設備 += CheckBox5.Text & ";"
            End If
            If CheckBox6.Checked = True Then
                隨附設備 += CheckBox6.Text & ";"
            End If
            If CheckBox7.Checked = True Then
                隨附設備 += CheckBox7.Text & ";"
            End If
            If CheckBox8.Checked = True Then
                隨附設備 += CheckBox8.Text & ";"
            End If
            If CheckBox9.Checked = True Then
                隨附設備 += CheckBox9.Text & ";"
            End If
            If CheckBox10.Checked = True Then
                隨附設備 += CheckBox10.Text & ";"
            End If
            If CheckBox11.Checked = True Then
                隨附設備 += CheckBox11.Text & ";"
            End If
            If CheckBox12.Checked = True Then
                隨附設備 += CheckBox12.Text & ";"
            End If
            If CheckBox13.Checked = True Then
                隨附設備 += CheckBox13.Text & ";"
            End If
            If CheckBox14.Checked = True Then
                隨附設備 += CheckBox14.Text & ";"
            End If
            If CheckBox15.Checked = True Then
                隨附設備 += CheckBox15.Text & ";"
            End If
            If CheckBox16.Checked = True Then
                隨附設備 += CheckBox16.Text & ";"
            End If
            If CheckBox17.Checked = True Then
                隨附設備 += CheckBox17.Text & ";"
            End If
            If CheckBox18.Checked = True Then
                隨附設備 += CheckBox18.Text & ";"
            End If
            If CheckBox19.Checked = True Then
                隨附設備 += CheckBox19.Text & ";"
            End If
            sql &= "'" & 隨附設備 & "',"
            '沙發數,電視數,冰箱數,冷氣數,洗衣機數,乾衣機數,"
            sql &= "'" & TextBox48.Text & "','" & TextBox49.Text & "','" & TextBox50.Text & "','" & TextBox51.Text & "','" & TextBox52.Text & "','" & TextBox53.Text & "',"

        Else
            ',隨附設備,沙發數,電視數,冰箱數,冷氣數,洗衣機數,乾衣機數,"
            sql &= "'',0,0,0,0,0,0,"
        End If

        '全棟
        If house.Checked = True Then
            sql &= "'Y', "
        Else
            sql &= "'N', "
        End If

        '其他重要事項,其他重要事項_說明
        sql &= "'" & DropDownList52.SelectedValue & "','" & TextBox54.Text & "',"

        sql &= "'" & sysdate & "','" & Request.Cookies("webfly_empno").Value & "'"
        sql &= ")"
        Dim cmd As New SqlCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        conn.Open()

        Try
            cmd.ExecuteNonQuery()

            Dim script As String = ""
            script += "alert('新增成功!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            ImageButton1.Visible = False
            ImageButton12.Visible = True
            ImageButton19.Visible = True

        Catch ex As Exception
            Response.Write(ex.ToString)

            'If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
            Response.Write(sql & "<br>")
            Response.End()
            'End If

            Dim script As String = ""
            script += "alert('新增失敗!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)

            ImageButton1.Visible = True
            ImageButton12.Visible = False
            ImageButton19.Visible = False
        End Try


        conn.Close()
        conn.Dispose()
    End Sub


    '修改
    Protected Sub ImageButton12_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton12.Click

        If DropDownList62.SelectedValue = "有" AndAlso String.IsNullOrWhiteSpace(TextBox250.Text) Then
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v1", "alert('「設置位置」為必填');", True)
            Exit Sub
        End If

        If DropDownList63.SelectedValue = "有" Then
            If String.IsNullOrWhiteSpace(DropDownList64.SelectedValue) Then
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v2", "alert('「能效等級」為必填，且不可選「-請選擇能效等級-」');", True)
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(TextBox253.Text) Then
                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "v3", "alert('「有效期間」為必填');", True)
                Exit Sub
            End If
        End If

        '如果編號跟店代號跟預設值不一樣，先下列步驟
        If oid.Text <> uoid.Text Or sid.Text <> usid.Text Then
            Dim conn_upt As New SqlConnection(EGOUPLOADSqlConnStr)
            Dim sql_upt As String = "update 產調_建物 set 物件編號='" & uoid.Text & "',店代號='" & usid.Text & "' where 物件編號='" & oid.Text & "' and 店代號='" & sid.Text & "'"
            Dim cmd_upt As New SqlCommand(sql_upt, conn_upt)
            cmd_upt.CommandType = CommandType.Text

            conn_upt.Open()



            Try
                cmd_upt.ExecuteNonQuery()
            Catch ex As Exception
                Response.Write(ex.ToString & "<br>" & sql_upt)
            End Try

            conn_upt.Close()
            conn_upt.Dispose()

            'UPDATE已存資料後給予新的值
            '物件編號
            Me.Label11.Text = uoid.Text

            '店代號
            Me.Label12.Text = usid.Text

        End If

        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        sql = "Update 產調_建物 "
        sql &= " Set "
        sql &= "限制登記='" & DropDownList41.SelectedValue & "',限制登記_說明='" & TextBox246.Text & "',信託登記='" & DropDownList29.SelectedValue & "',信託登記_說明='" & TextBox247.Text & "',其他權利='" & DropDownList1.SelectedValue & "',其他權利_說明='" & TextBox248.Text & "',"
        sql &= "是否共有='" & DropDownList58.SelectedValue & "',有無分管協議登記='" & RadioButtonList4.SelectedValue & "',有無分管協議登記_說明='" & TextBox68.Text & "',"
        sql &= "獎勵容積開放空間提供公共使用='" & DropDownList36.SelectedValue & "',獎勵容積開放空間提供公共使用_說明='" & TextBox236.Text & "',"
        sql &= "屬工業區或其他分區='" & DropDownList10.SelectedValue & "', 屬工業區或其他分區_說明='" & TextBox245.Text & "', 使用執照有無備註之注意事項='" & DropDownList59.SelectedValue & "', 使用執照有無備註之注意事項_說明='" & TextBox241.Text & "',"
        sql &= "有無禁建情事='" & DropDownList57.SelectedValue & "',有無禁建情事_說明='" & TextBox21.Text & "',"
        sql &= "中繼幫浦='" & DropDownList61.SelectedValue & "',中繼幫浦_說明='" & TextBox249.Text & "',"
        ''******** 太陽光電 建築能效********
        sql &= "太陽光電發電設備='" & DropDownList62.SelectedValue & "',太陽光電發電設備_設置位置='" & TextBox250.Text & "',"
        sql &= "建築能效標示='" & DropDownList63.SelectedValue & "',建築能效標示_能效等級='" & DropDownList64.SelectedValue & "',建築能效標示_證書效期='" & TextBox253.Text & "',"
        '******** 屋主現況說明部分 ********
        sql &= "頂樓基地台='" & DropDownList2.SelectedValue & "',頂樓基地台_說明='" & TextBox1.Text & "',"
        sql &= "衛生下水道工程='" & DropDownList3.SelectedValue & "',衛生下水道工程_選項='" & DropDownList4.SelectedValue & "',衛生下水道工程_說明='" & TextBox2.Text & "',"
        'sql &= "規約外特殊使用='" & DropDownList5.SelectedValue & "',規約外特殊使用_共用說明='" & TextBox3.Text & "', 規約外特殊使用_專用說明='" & TextBox4.Text & "',"
        sql &= "共有部分範圍_有無='" & DropDownList5.SelectedValue & "', 共有部分範圍='" & TextBox3.Text & "', 專有部分範圍_有無='" & DropDownList60.SelectedValue & "', 專有部分範圍='" & TextBox4.Text & "',"
        sql &= "管理公司_有無='" & DropDownList6.SelectedValue & "',管理公司='" & TextBox5.Text & "',"
        sql &= "專有約定共用='" & DropDownList8.SelectedValue & "',專有約定共用之範圍='" & TextBox6.Text & "',專有約定共用之使用方式='" & TextBox7.Text & "',"
        sql &= "共有約定專用='" & DropDownList9.SelectedValue & "',共有約定專用之範圍='" & TextBox8.Text & "',共有約定專用之使用方式='" & TextBox9.Text & "',"
        sql &= "公共基金有無='" & DropDownList11.SelectedValue & "',"

        If DropDownList11.SelectedValue = "有" Then
            If Trim(TextBox10.Text) <> "" Then
                sql &= "公共基金數額='" & TextBox10.Text & "',"
            Else
                sql &= "公共基金數額='--',"
            End If
            If DropDownList12.SelectedIndex <> 0 Then
                sql &= "公共基金提撥方式='" & TextBox11.Text & "',"
            Else
                sql &= "公共基金提撥方式='" & DropDownList12.SelectedValue & "',"
            End If
            If DropDownList13.SelectedIndex <> 0 Then
                sql &= "公共基金運用方式='" & TextBox12.Text & "',"
            Else
                sql &= "公共基金運用方式='" & DropDownList13.SelectedValue & "',"
            End If
        Else
            sql &= "公共基金提撥方式='無',公共基金運用方式='無',"
        End If

        sql &= "管理費使用='" & DropDownList14.SelectedValue & "',"

        If DropDownList14.SelectedValue = "有" Then
            sql &= "管理費或使用費='" & DropDownList15.SelectedValue & "',管理費='" & TextBox14.Text & "',管理費繳交方式='" & TextBox15.Text & "',"

        Else
            sql &= "管理費或使用費='無',管理費=0,管理費繳交方式='無',"
        End If

        sql &= "管理組織_有無='" & DropDownList16.SelectedValue & "',"
        If DropDownList16.SelectedValue = "有" Then
            If DropDownList17.SelectedValue <> "其他" Then
                sql &= "管理組織及方式='" & DropDownList17.SelectedValue & "',"
            Else
                sql &= "管理組織及方式='" & TextBox13.Text & "',"
            End If
        Else
            sql &= "管理組織及方式='無',"
        End If


        sql &= "住戶規約使用手冊='" & DropDownList18.SelectedValue & "',"

        sql &= "電梯設備='" & DropDownList7.SelectedValue & "',張貼合格標章='" & DropDownList19.SelectedValue & "',張貼合格標章_說明='" & TextBox16.Text & "',"

        sql &= "出租狀況='" & DropDownList53.SelectedValue & "',出租範圍='" & DropDownList55.SelectedValue & "',出租範圍備註='" & TextBox57.Text & "',"

        Dim 出租情況類型 As String = ""
        If RadioButton6.Checked = True Then
            出租情況類型 &= "不定期租約;"
            'sql &= "出租情況類型='不定期租約',"
            sql &= "租金='" & TextBox55.Text & "',"
            sql &= "租期起='',租期迄='',租約是否公證='',押租保證金='',"

        End If
        If RadioButton7.Checked = True Then
            出租情況類型 &= "定期租約;"
            'sql &= "出租情況類型='定期租約',"
            sql &= "租金='" & TextBox59.Text & "',"
            sql &= "租期起='" & TextBox61.Text & "',租期迄='" & TextBox62.Text & "',租約是否公證='" & RadioButtonList3.SelectedValue & "',押租保證金='" & TextBox60.Text & "',"
        End If
        If RadioButton8.Checked = True Then     '租賃之權利義務隨同移轉
            出租情況類型 &= RadioButton8.Text & ";"
            'sql &= "出租情況類型='" & RadioButton8.Text & "',"
            'sql &= "租金=0,"
            'sql &= "租期起='',租期迄='',租約是否公證='',押租保證金='',"
        End If
        If RadioButton9.Checked = True Then     '屋主終止租約騰空交屋
            出租情況類型 &= RadioButton9.Text & ";"
            'sql &= "出租情況類型='" & RadioButton9.Text & "',"
        End If
        If RadioButton10.Checked = True Then    '其他
            出租情況類型 &= "其他;"
            'sql &= "出租情況類型='" & RadioButton10.Text & "',"
        End If
        sql &= "出租情況類型='" & 出租情況類型 & "',"

        sql &= "出租狀況_說明='" & TextBox58.Text & "',"

        sql &= "出借狀況='" & DropDownList54.SelectedValue & "',出借範圍='" & DropDownList56.SelectedValue & "',出借範圍備註='" & TextBox56.Text & "',"

        If RadioButton11.Checked = True Then
            sql &= "出借書面約定與否='" & RadioButton11.Text & "',"

            sql &= "借用人姓名='',出借起日='',出借迄日='',出借返還條件='',"
        Else
            sql &= "出借書面約定與否='" & RadioButton12.Text & "',"

            sql &= "借用人姓名='" & TextBox63.Text & "',出借起日='" & TextBox64.Text & "',出借迄日='" & TextBox65.Text & "',出借返還條件='" & TextBox66.Text & "',"
        End If

        sql &= "佔用情形='" & DropDownList20.SelectedValue & "',"

        If DropDownList21.SelectedValue <> "其他" Then
            sql &= "佔用他人建物土地='" & DropDownList21.SelectedValue & "',"
        Else
            sql &= "佔用他人建物土地='" & TextBox67.Text & "',"
        End If

        sql &= "佔用他人建物土地_說明='" & TextBox17.Text & "',"

        '消防設備有無
        sql &= "消防設備='" & DropDownList22.SelectedValue & "',"
        '消防設備說明
        Dim tempdescript As String = ""
        For Each lit As ListItem In CheckBoxList1.Items
            If lit.Selected = True Then
                tempdescript += lit.Text & ","
            End If
        Next
        tempdescript += TextBox18.Text & ","
        sql &= "消防設備_說明='" & tempdescript & "',"

        sql &= "無障礙設施='" & DropDownList23.SelectedValue & "',"
        sql &= "無障礙設施_說明='" & TextBox19.Text & "',"

        sql &= "夾層='" & DropDownList24.SelectedValue & "',"
        If TextBox20.Text <> "" Then
            sql &= "夾層面積='" & TextBox20.Text & "',"
        End If
        If TextBox22.Text <> "" Then
            sql &= "夾層面積1='" & TextBox22.Text & "',"
        End If
        sql &= "夾層其他='" & TextBox23.Text & "',"

        sql &= "獨立供水='" & DropDownList25.SelectedValue & "',供水類型='" & DropDownList26.SelectedValue & "',供水是否正常='" & RadioButtonList1.SelectedValue & "',獨立供水_說明='" & TextBox24.Text & "',"

        sql &= "獨立電表='" & DropDownList27.SelectedValue & "',獨立電表_說明='" & TextBox25.Text & "',"
        If DropDownList28.SelectedValue = "有" Then
            sql &= "天然瓦斯='" & DropDownList28.SelectedValue & "',天然瓦斯_說明='" & RadioButtonList2.SelectedValue & "',天然瓦斯_說明2='" & TextBox26.Text & "',"
        Else
            sql &= "天然瓦斯='" & DropDownList28.SelectedValue & "',天然瓦斯_說明='',天然瓦斯_說明2='',"
        End If

        sql &= "水電管線是否更新='" & DropDownList30.SelectedValue & "',水管更新日期='" & TextBox27.Text & "',電線更新日期='" & TextBox28.Text & "',"
        sql &= "積欠應繳費用='" & DropDownList31.SelectedValue & "',積欠應繳費用_說明='" & TextBox29.Text & "',"
        sql &= "持有期間有無居住='" & DropDownList32.SelectedValue & "',"
        sql &= "有無公共設施重大修繕='" & DropDownList33.SelectedValue & "',有無公共設施重大修繕_說明='" & TextBox30.Text & "',有無公共設施重大修繕_金額='" & TextBox31.Text & "',"

        sql &= "混凝土中氯離子含量='" & DropDownList34.SelectedValue & "',"

        If DropDownList35.SelectedValue <> "其他" Then
            sql &= "混凝土中氯離子含量_說明='" & DropDownList35.SelectedValue & "',"
        Else
            sql &= "混凝土中氯離子含量_說明='" & TextBox32.Text & "',"
        End If

        sql &= "輻射檢測='" & DropDownList37.SelectedValue & "',"
        If DropDownList38.SelectedValue <> "其他" Then
            sql &= "輻射檢測_說明='" & DropDownList38.SelectedValue & "',"
        Else
            sql &= "輻射檢測_說明='" & TextBox33.Text & "',"
        End If

        sql &= "曾發生火災或其他災害='" & DropDownList39.SelectedValue & "',曾發生火災或其他災害_說明='" & TextBox34.Text & "',"
        sql &= "因地震被公告為危險建築='" & DropDownList40.SelectedValue & "',因地震被公告為危險建築_說明='" & TextBox35.Text & "',"
        sql &= "樑柱部分是否有顯見裂痕='" & DropDownList42.SelectedValue & "',樑柱部分是否有顯見裂痕_說明='" & TextBox36.Text & "',裂痕長度='" & TextBox37.Text & "',間隙寬度='" & TextBox38.Text & "',"
        sql &= "建物鋼筋裸露='" & DropDownList43.SelectedValue & "',建物鋼筋裸露_說明='" & TextBox39.Text & "',"

        sql &= "是否為兇宅='" & DropDownList44.SelectedValue & "',"

        If RadioButton1.Checked = True Then
            sql &= "兇宅發生期間='" & RadioButton1.Text & "',是否為兇宅_說明='" & TextBox40.Text & "',"
        ElseIf RadioButton2.Checked = True Then
            sql &= "兇宅發生期間='" & RadioButton2.Text & "',是否為兇宅_說明='" & TextBox41.Text & "',"
        ElseIf RadioButton3.Checked = True Then
            sql &= "兇宅發生期間='" & RadioButton3.Text & "',是否為兇宅_說明='',"
        ElseIf RadioButton4.Checked = True Then
            sql &= "兇宅發生期間='" & RadioButton4.Text & "',是否為兇宅_說明='',"
        ElseIf RadioButton5.Checked = True Then
            sql &= "兇宅發生期間='" & RadioButton5.Text & "',是否為兇宅_說明='" & TextBox42.Text & "',"
        Else
            sql &= "兇宅發生期間='',是否為兇宅_說明='',"
        End If

        sql &= "滲漏水狀態='" & DropDownList45.SelectedValue & "',"

        If DropDownList45.SelectedValue = "有" Then
            Dim postemp As String = TextBox43.Text & ";"
            For Each lis As ListItem In CheckBoxList2.Items
                If lis.Selected = True Then
                    postemp += lis.Text & ";"
                End If
            Next
            sql &= "滲漏水狀態_說明='" & postemp & "',"

            If DropDownList46.SelectedValue <> "其他" Then
                sql &= "滲漏水狀態_處理='" & DropDownList46.SelectedValue & "',"
            Else
                sql &= "滲漏水狀態_處理='" & TextBox44.Text & "',"
            End If
        Else
            sql &= "滲漏水狀態_說明='',滲漏水狀態_處理='',"
        End If



        sql &= "違增建使用權='" & DropDownList47.SelectedValue & "',"
        If DropDownList47.SelectedValue = "有" Then
            Dim postemp2 As String = ""
            For Each lis As ListItem In CheckBoxList3.Items
                If lis.Selected = True Then
                    postemp2 += lis.Text & ";"
                End If
            Next
            If TextBox45.Text.Length > 0 Then
                postemp2 += TextBox45.Text & ";"
            End If
            sql &= "違增建使用權_說明='" & postemp2 & "',"

            sql &= "違增建_面積='" & TextBox47.Text & "',"
        Else
            sql &= "違增建使用權_說明='',違增建_面積='',"
        End If

        sql &= "違增建列管情形='" & DropDownList48.SelectedValue & "',違增建列管情形_說明='" & TextBox46.Text & "',"
        sql &= "排水系統='" & DropDownList49.SelectedValue & "',排水系統_說明='" & DropDownList50.SelectedValue & "',"

        sql &= "隨附設備有無='" & DropDownList51.SelectedValue & "',"
        Dim 隨附設備 As String = ""
        If CheckBox1.Checked = True Then
            隨附設備 += CheckBox1.Text & ";"
        End If
        If CheckBox2.Checked = True Then
            隨附設備 += CheckBox2.Text & ";"
        End If
        If CheckBox3.Checked = True Then
            隨附設備 += CheckBox3.Text & ";"
        End If
        If CheckBox4.Checked = True Then
            隨附設備 += CheckBox4.Text & ";"
        End If
        If CheckBox5.Checked = True Then
            隨附設備 += CheckBox5.Text & ";"
        End If
        If CheckBox6.Checked = True Then
            隨附設備 += CheckBox6.Text & ";"
        End If
        If CheckBox7.Checked = True Then
            隨附設備 += CheckBox7.Text & ";"
        End If
        If CheckBox8.Checked = True Then
            隨附設備 += CheckBox8.Text & ";"
        End If
        If CheckBox9.Checked = True Then
            隨附設備 += CheckBox9.Text & ";"
        End If
        If CheckBox10.Checked = True Then
            隨附設備 += CheckBox10.Text & ";"
        End If
        If CheckBox11.Checked = True Then
            隨附設備 += CheckBox11.Text & ";"
        End If
        If CheckBox12.Checked = True Then
            隨附設備 += CheckBox12.Text & ";"
        End If
        If CheckBox13.Checked = True Then
            隨附設備 += CheckBox13.Text & ";"
        End If
        If CheckBox14.Checked = True Then
            隨附設備 += CheckBox14.Text & ";"
        End If
        If CheckBox15.Checked = True Then
            隨附設備 += CheckBox15.Text & ";"
        End If
        If CheckBox16.Checked = True Then
            隨附設備 += CheckBox16.Text & ";"
        End If
        If CheckBox17.Checked = True Then
            隨附設備 += CheckBox17.Text & ";"
        End If
        If CheckBox18.Checked = True Then
            隨附設備 += CheckBox18.Text & ";"
        End If
        If CheckBox19.Checked = True Then
            隨附設備 += CheckBox19.Text & ";"
        End If
        sql &= "隨附設備='" & 隨附設備 & "',"
        sql &= "沙發數='" & TextBox48.Text & "',電視數='" & TextBox49.Text & "',冰箱數='" & TextBox50.Text & "',冷氣數='" & TextBox51.Text & "',洗衣機數='" & TextBox52.Text & "',乾衣機數='" & TextBox53.Text & "',"
        sql &= "其他重要事項='" & DropDownList52.SelectedValue & "',其他重要事項_說明='" & TextBox54.Text & "',"

        If house.Checked = True Then
            sql &= "全棟='Y', "
        Else
            sql &= "全棟='N',"
        End If

        '***********************
        sql &= " 修改日期 = '" & sysdate & "' ,"
        sql &= " 修改者 = '" & Request.Cookies("webfly_empno").Value & "' "
        sql &= " Where 物件編號 = '" & Me.Label11.Text & "' and 流水號='" & NUM.Text & "' and 店代號='" & usid.Text & "'"


        Dim cmd As New SqlCommand(sql, conn)

        Try
            cmd.ExecuteNonQuery()
            Dim script As String = ""
            script += "alert('修改成功!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        Catch ex As Exception
            checkleave.錯誤訊息(Request.Cookies("store_id").Value.ToString, Request.Cookies("webfly_empno").Value.ToString, "產調-建物", Request.Url.ToString, sql, ex.ToString)
            'Response.Write(ex.ToString & "<br>" & sql)
            If Request.Cookies("webfly_empno").Value.ToLower = "92h" Then
                Response.End()
            End If
            Dim script As String = ""
            script += "alert('修改失敗!!');"
            ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
        End Try


        conn.Close()
        conn.Dispose()
    End Sub

    '刪除
    Protected Sub ImageButton19_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImageButton19.Click
        conn = New SqlConnection(EGOUPLOADSqlConnStr)
        conn.Open()

        '刪除產調_土地----------------------------
        sql = "Delete 產調_建物 Where 物件編號 = '" & Me.Label11.Text & "' and 流水號='" & NUM.Text & "' and 店代號='" & usid.Text & "'"

        Dim cmd As New SqlCommand(sql, conn)
        cmd.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

        Dim script As String = ""
        script += "alert('刪除成功!!');"
        ScriptManager.RegisterStartupScript(Me, GetType(String), "", script, True)
    End Sub


    Protected Sub DropDownList62_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList62.SelectedIndexChanged
        If DropDownList62.SelectedValue = "有" Then
            LabelSolarPos.Visible = True
            TextBox250.Visible = True
        Else
            LabelSolarPos.Visible = False
            TextBox250.Visible = False
            TextBox250.Text = ""
        End If
    End Sub

    Protected Sub DropDownList63_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList63.SelectedIndexChanged
        If DropDownList63.SelectedValue = "有" Then
            PlaceHolder86.Visible = True
        Else
            PlaceHolder86.Visible = False
            DropDownList64.SelectedIndex = 0
            TextBox253.Text = ""
        End If
    End Sub




    '---------------------------------------------------------CHANGE事件---------------------------------------------------------
    '限制登記
    Protected Sub DropDownList41_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList41.SelectedIndexChanged
        If DropDownList41.SelectedValue = "無" Then
            Label38.Visible = False
            TextBox246.Text = ""
            TextBox246.Visible = False
        Else
            Label38.Visible = True
            TextBox246.Text = ""
            TextBox246.Visible = True
        End If
    End Sub
    '信託登記
    Protected Sub DropDownList29_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList29.SelectedIndexChanged
        If DropDownList29.SelectedValue = "無" Then
            Label39.Visible = False
            TextBox247.Text = ""
            TextBox247.Visible = False
        Else
            Label39.Visible = True
            TextBox247.Text = ""
            TextBox247.Visible = True
        End If
    End Sub
    '其他事項(其他權利)
    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        If DropDownList1.SelectedValue = "無" Then
            Label40.Visible = False
            TextBox248.Text = ""
            TextBox248.Visible = False
        Else
            Label40.Visible = True
            TextBox248.Text = ""
            TextBox248.Visible = True
        End If
    End Sub

    '是否共有
    Protected Sub DropDownList58_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList58.SelectedIndexChanged
        If DropDownList58.SelectedValue = "無" Then
            PlaceHolder19.Visible = False

        Else
            PlaceHolder19.Visible = True

        End If
    End Sub

    '獎勵容積開放空間
    Protected Sub DropDownList36_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList36.SelectedIndexChanged
        If DropDownList36.SelectedValue = "無" Then
            Label17.Visible = False
            TextBox236.Text = ""
            TextBox236.Visible = False
        Else
            Label17.Visible = True
            TextBox236.Text = ""
            TextBox236.Visible = True
        End If
    End Sub




    '屬工業區或其他分區
    Protected Sub DropDownList10_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList10.SelectedIndexChanged
        If DropDownList10.SelectedValue = "無" Then
            Label37.Visible = False
            TextBox245.Text = ""
            TextBox245.Visible = False
        Else
            Label37.Visible = True
            TextBox245.Visible = True
        End If
    End Sub
    '使用執照有無備註之注意事項
    Protected Sub DropDownList59_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList59.SelectedIndexChanged
        If DropDownList59.SelectedValue = "無" Then
            Label36.Visible = False
            TextBox241.Text = ""
            TextBox241.Visible = False
        Else
            Label36.Visible = True
            TextBox241.Visible = True
        End If
    End Sub

    '有無禁建情事
    Protected Sub DropDownList57_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList57.SelectedIndexChanged
        If DropDownList57.SelectedValue = "無" Then
            Label44.Visible = False
            TextBox21.Text = ""
            TextBox21.Visible = False
        Else
            Label44.Visible = True
            TextBox21.Visible = True
        End If
    End Sub

    Protected Sub DropDownList61_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList61.SelectedIndexChanged
        If DropDownList61.SelectedValue = "無" Then
            Label45.Visible = False
            TextBox249.Text = ""
            TextBox249.Visible = False
        Else
            Label45.Visible = True
            TextBox249.Visible = True
        End If
    End Sub

    '*********** 屋主現況說明 *****************
    '行動電話基地台
    Protected Sub DropDownList2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        If DropDownList2.SelectedValue = "無" Then
            Label1.Visible = False
            TextBox1.Text = ""
            TextBox1.Visible = False
        Else
            Label1.Visible = True
            TextBox1.Visible = True
        End If
    End Sub
    '下水道工程
    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList3.SelectedIndexChanged
        If DropDownList3.SelectedValue = "有" Then
            DropDownList4.Visible = False
            Label2.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        Else
            DropDownList4.Visible = True
            Label2.Visible = True
            TextBox2.Visible = True

        End If
    End Sub
    '共用部分有分管協議
    Protected Sub DropDownList5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList5.SelectedIndexChanged
        If DropDownList5.SelectedValue = "無" Then
            PlaceHolder7.Visible = False
        Else
            PlaceHolder7.Visible = True
        End If
    End Sub

    '使用專有部分有限制
    Protected Sub DropDownList60_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList60.SelectedIndexChanged
        If DropDownList60.SelectedValue = "無" Then
            PlaceHolder77.Visible = False
        Else
            PlaceHolder77.Visible = True
        End If
    End Sub

    '管理維護公司
    Protected Sub DropDownList6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList6.SelectedIndexChanged
        If DropDownList6.SelectedValue = "無" Then
            Label6.Visible = False
            TextBox5.Visible = False
        Else
            Label6.Visible = True
            TextBox5.Visible = True
        End If
    End Sub
    '約定共用
    Protected Sub DropDownList8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList8.SelectedIndexChanged
        If DropDownList8.SelectedValue = "無" Then
            PlaceHolder8.Visible = False
        Else
            PlaceHolder8.Visible = True
        End If
    End Sub
    '約定專用
    Protected Sub DropDownList9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList9.SelectedIndexChanged
        If DropDownList9.SelectedValue = "無" Then
            PlaceHolder9.Visible = False
        Else
            PlaceHolder9.Visible = True
        End If
    End Sub
    '公共基金
    Protected Sub DropDownList11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList11.SelectedIndexChanged
        If DropDownList11.SelectedValue = "無" Then
            PlaceHolder5.Visible = False
        Else
            PlaceHolder5.Visible = True
        End If
    End Sub

    '社區管理費或使用費
    Protected Sub DropDownList14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList14.SelectedIndexChanged
        If DropDownList14.SelectedValue = "無" Then
            PlaceHolder6.Visible = False
            TextBox14.Text = ""
            TextBox15.Text = ""
        Else
            PlaceHolder6.Visible = True
        End If
    End Sub
    '有無電梯設備
    Protected Sub DropDownList7_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DropDownList7.SelectedIndexChanged
        If DropDownList7.SelectedValue = "無" Then
            DropDownList19.Visible = False
            Label7.Visible = False
            Label16.Visible = False
            TextBox16.Visible = False
        Else
            DropDownList19.Visible = True
            Label7.Visible = True
            Label16.Visible = True
            TextBox16.Visible = True
        End If
    End Sub
    '電梯設備有無張貼有效合格認證標章
    Protected Sub DropDownList19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList19.SelectedIndexChanged
        If DropDownList19.SelectedValue = "有" Then
            Label7.Visible = True
            Label16.Visible = False
            TextBox16.Visible = False
        Else
            Label7.Visible = True
            Label16.Visible = True
            TextBox16.Visible = True
        End If
    End Sub
    '建物有無出租情形
    Protected Sub DropDownList53_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList53.SelectedIndexChanged
        If DropDownList53.SelectedValue = "無" Then
            PlaceHolder1.Visible = False
        Else
            PlaceHolder1.Visible = True
        End If
    End Sub
    '管理組織
    Protected Sub DropDownList16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList16.SelectedIndexChanged
        If DropDownList16.SelectedValue = "無" Then
            DropDownList17.Visible = False
            TextBox13.Visible = False
        Else
            DropDownList17.Visible = True
            TextBox13.Visible = True
        End If
    End Sub
    '建物有無有出借情形
    Protected Sub DropDownList54_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList54.SelectedIndexChanged
        If DropDownList54.SelectedValue = "無" Then
            PlaceHolder2.Visible = False
        Else
            PlaceHolder2.Visible = True
        End If
    End Sub
    '建物有無占用情形
    Protected Sub DropDownList20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList20.SelectedIndexChanged
        If DropDownList20.SelectedValue = "無" Then
            PlaceHolder13.Visible = False

        Else
            PlaceHolder13.Visible = True

        End If
    End Sub
    '本標的物及公設內有無消防設施
    Protected Sub DropDownList22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList22.SelectedIndexChanged
        If DropDownList22.SelectedValue = "無" Then
            CheckBoxList1.Visible = False
            TextBox18.Visible = False
        Else
            CheckBoxList1.Visible = True
            TextBox18.Visible = True
        End If
    End Sub
    '本標的物及公設內有無無障礙設施
    Protected Sub DropDownList23_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList23.SelectedIndexChanged
        If DropDownList23.SelectedValue = "無" Then
            PlaceHolder10.Visible = False
        Else
            PlaceHolder10.Visible = True
        End If
    End Sub
    '房屋有無施作夾層
    Protected Sub DropDownList24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList24.SelectedIndexChanged
        If DropDownList24.SelectedValue = "無" Then
            PlaceHolder11.Visible = False
        Else
            PlaceHolder11.Visible = True
        End If
    End Sub
    '本戶有無供水使用
    Protected Sub DropDownList25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList25.SelectedIndexChanged
        If DropDownList25.SelectedValue = "無" Then
            PlaceHolder12.Visible = False
        Else
            PlaceHolder12.Visible = True
        End If
    End Sub
    '本戶有無獨立之電表供電

    Protected Sub DropDownList27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList27.SelectedIndexChanged
        ' 
        If DropDownList27.SelectedValue = "有" Then
            Label21.Visible = False
            TextBox25.Visible = False
        Else
            Label21.Visible = True
            TextBox25.Visible = True
        End If
    End Sub
    '本戶瓦斯供應情形
    Protected Sub DropDownList28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList28.SelectedIndexChanged
        If DropDownList28.SelectedValue = "無" Then
            RadioButtonList2.Visible = False
            TextBox26.Visible = False
        Else
            RadioButtonList2.Visible = True
            TextBox26.Visible = True
        End If
    End Sub
    '水、電管線於產權持有期間有無更新
    Protected Sub DropDownList30_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList30.SelectedIndexChanged
        If DropDownList30.SelectedValue = "無" Then
            PlaceHolder14.Visible = False
        Else
            PlaceHolder14.Visible = True
        End If
    End Sub
    '有無積欠應繳費用情形
    Protected Sub DropDownList31_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList31.SelectedIndexChanged
        If DropDownList31.SelectedValue = "無" Then
            Label3.Visible = False
            TextBox29.Visible = False
        Else
            Label3.Visible = True
            TextBox29.Visible = True
        End If
    End Sub
    '社區有無公共設施重大修繕決議
    Protected Sub DropDownList33_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList33.SelectedIndexChanged
        If DropDownList33.SelectedValue = "無" Then
            PlaceHolder15.Visible = False
        Else
            PlaceHolder15.Visible = True
        End If
    End Sub
    '混凝土中水溶性氯離子含量檢測情事
    'Protected Sub DropDownList34_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList34.SelectedIndexChanged
    '    If DropDownList34.SelectedValue = "無" Then
    '        DropDownList35.Visible = False
    '        TextBox32.Visible = False
    '    Else
    '        DropDownList35.Visible = True
    '        TextBox32.Visible = True
    '    End If
    'End Sub
    '輻射檢測情事
    Protected Sub DropDownList37_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList37.SelectedIndexChanged
        If DropDownList37.SelectedValue = "無" Then
            DropDownList38.Visible = True
            TextBox33.Visible = True
            Label9.Text = "若無,未檢測原因:"
        Else
            Label9.Text = "(附檢測結果)"
            DropDownList38.Visible = False
            TextBox33.Visible = False
        End If
    End Sub
    '建物有無曾經發生火災或其他天然災害或人為破壞造成建築物損害及修繕情形
    Protected Sub DropDownList39_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList39.SelectedIndexChanged
        If DropDownList39.SelectedValue = "無" Then
            Label22.Visible = False
            TextBox34.Visible = False
        Else
            Label22.Visible = True
            TextBox34.Visible = True
        End If
    End Sub
    '建物目前有無因地震被建管單位公告列為危險建築情形
    Protected Sub DropDownList40_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList40.SelectedIndexChanged
        If DropDownList40.SelectedValue = "無" Then
            Label23.Visible = False
            TextBox35.Visible = False
        Else
            Label23.Visible = True
            TextBox35.Visible = True
        End If
    End Sub
    '樑、柱部分有無顯見間隙裂痕
    Protected Sub DropDownList42_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList42.SelectedIndexChanged
        If DropDownList42.SelectedValue = "無" Then
            PlaceHolder16.Visible = False

        Else
            PlaceHolder16.Visible = True

        End If
    End Sub
    '房屋鋼筋有無裸露
    Protected Sub DropDownList43_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList43.SelectedIndexChanged
        If DropDownList43.SelectedValue = "無" Then
            Label27.Visible = False
            TextBox39.Visible = False
        Else
            Label27.Visible = True
            TextBox39.Visible = True
        End If
    End Sub
    '本建物(專有部分)有無發生兇殺、自殺、一氧化碳中毒或其他非自然死亡之情形
    Protected Sub DropDownList44_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList44.SelectedIndexChanged
        If DropDownList44.SelectedValue = "無" Then

            PlaceHolder17.Visible = False
        Else
            PlaceHolder17.Visible = True

        End If
    End Sub
    '有無滲漏水狀況
    Protected Sub DropDownList45_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList45.SelectedIndexChanged
        If DropDownList45.SelectedValue = "無" Then

            PlaceHolder4.Visible = False
        Else
            PlaceHolder4.Visible = True

        End If
    End Sub
    '建物有無違建、增建情事？ 其位置、約略面積及建管機關列管情形？
    Protected Sub DropDownList47_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList47.SelectedIndexChanged
        If DropDownList47.SelectedValue = "無" Then
            PlaceHolder3.Visible = False
        Else
            PlaceHolder3.Visible = True
        End If
    End Sub
    '建物排水系統有無正常
    Protected Sub DropDownList49_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList49.SelectedIndexChanged
        If DropDownList49.SelectedValue = "有" Then
            DropDownList50.Visible = False
            Label4.Visible = False
        Else
            DropDownList50.Visible = True
            Label4.Visible = True
        End If
    End Sub
    '賣方是否有附隨買賣之設備？
    Protected Sub DropDownList51_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList51.SelectedIndexChanged
        If DropDownList51.SelectedValue = "無" Then
            PlaceHolder18.Visible = False

        Else
            PlaceHolder18.Visible = True

        End If
    End Sub
    '其他重要事項
    Protected Sub DropDownList52_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList52.SelectedIndexChanged
        If DropDownList52.SelectedValue = "無" Then
            Label28.Visible = False
            TextBox54.Visible = False
        Else
            Label28.Visible = True
            TextBox54.Visible = True
        End If
    End Sub


End Class
