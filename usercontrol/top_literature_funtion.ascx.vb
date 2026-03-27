Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Partial Class usercontrol_top_literature_funtion
    Inherits System.Web.UI.UserControl
    Dim CheckLeave As checkleave

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fun As String = System.IO.Path.GetFileName(Request.PhysicalPath)

        Dim Direct As String = "N"
        Direct = checkleave.直營(Request.Cookies("store_id").Value.ToString)

        li2.Text = ""
        Select Case fun
            Case "literature_download.aspx"   '東森,分眾下屏廣告
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    'li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_eiptv.aspx"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_download_eiptv.aspx"   '東森,分眾下屏廣告
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li  id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    'li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_eiptv.aspx"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_eiptv.aspx"  '東森,分眾下屏廣告
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_ai.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"

                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_board.aspx", "literature_board2.aspx", "literature_board3.aspx", "literature_board4.aspx", "literature_board5.aspx", "literature_board6.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"

                    If Direct = "N" Then
                        li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_etwarm_logo.aspx", "literature_DM.aspx", "literature_standard_logo.aspx", "literature_paper_logo.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_card.aspx", "literature_card_download.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_plan.aspx", "literature_plan_shopsign.aspx", "literature_plan_pillar", "literature_plan_pillar.aspx", "literature_plan_arcade.aspx", "literature_plan_window.aspx", "literature_plan_distinguish.aspx", "literature_plan_distinguishs.aspx", "literature_plan_certificate.aspx", "literature_plan_light.aspx", "literature_plan_download.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_canvas.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_element.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_vr.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
            Case "literature_Terrier.aspx"
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text = "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download.aspx?types=image"">文宣下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_download_eiptv.aspx?types=image"">分眾下屏廣告</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_ai.aspx"">總部相關廣告</a></li>"
                    If Direct = "N" Then
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board.aspx"">租售看板</a></li>"
                    Else
                        li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_board5.aspx"">租售看板</a></li>"
                    End If
                End If
                li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_etwarm_logo.aspx"">LOGO下載</a></li>"
                If Direct = "N" Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_card.aspx"">制式名片</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_plan.aspx"">店面規劃</a></li>"
                End If
                If InStr(Request.Cookies("store_id").Value.ToUpper, "S") = 0 Then
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_canvas.aspx"">大型帆布下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_element.aspx?types=sticker"">元件素材下載</a></li>"
                    li2.Text &= "<li id=""current""><a href=""..\F_HeadquartersSupport\literature_vr.aspx"">影音/VR素材下載</a></li>"
                    li2.Text &= "<li><a href=""..\F_HeadquartersSupport\literature_Terrier.aspx"">房產Line貼圖下載</a></li>"
                End If
        End Select
        'li2.Text &= "<li id=""current""><a href='#' onClick=""GB_showCenter('使用說明', '../Explanation/use.aspx?id=237',600,1200)"";>使用說明</a></li>"
        'li2.Text &= "<li id=""current""><a href='#' onClick=""GB_showCenter('相關文件', '../Explanation/document.aspx?Class_b_ID=E&Class_s_ID=08&Class_area_ID=00001',600,1200)"";>相關文件</a></li>"
    End Sub
End Class
