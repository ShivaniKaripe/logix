<script runat="server">
' version:7.3.1.138972.Official Build (SUSDAY10202)
  Dim CopientFileName As String = "graphic-reward.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  Dim DeleteBtnDisabled As String = ""
  Dim SummaryOnly As Boolean = False
  Dim IsTemplateOffer As Boolean = False
  Dim FromTemplateOffer As Boolean = False
  Dim RewardDisabled As String = ""
  Dim OfferEditable As Boolean = True
  
  Public Sub SetDeleteBtnDisabled(ByVal disableText As String)
    DeleteBtnDisabled = disableText
  End Sub
  
  Public Sub SetSummaryOnly(ByVal bSummary As Boolean)
    SummaryOnly = bSummary
  End Sub
  
  Public Sub SetEditableByUser(ByVal IsEditable As Boolean)
    OfferEditable = IsEditable
  End Sub
  
  Public Function GetSummaryOnly() As Boolean
    Return SummaryOnly
  End Function
  
  Public Sub RemoveGraphic(ByVal IncentiveID As Long, ByVal DeliverableID As Long)
    Dim MyCommon As New Copient.CommonInc
    'Dim DeliverableID As Long
    Dim Done As Boolean
    Dim rst As DataTable
    Dim row As DataRow
    'Dim IncentiveID As Long
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      'IncentiveID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
      'DeliverableID = MyCommon.Extract_Val(Request.QueryString("deliverableid"))
      MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "delete from CPE_DeliverableROIDs with (RowLock) where DeliverableID=" & DeliverableID & ";"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), StatusFlag=1 where IncentiveID=" & IncentiveID & ";"
      MyCommon.LRT_Execute()
      'clean up any orphans created by deleting the deliverable
      Done = False
      While Not (Done)
        MyCommon.QueryStr = "select RO.RewardOptionID from CPE_RewardOptions as RO with (NoLock) " & _
                            "where RO.IncentiveID=" & IncentiveID & " and RO.TouchResponse=1 and RO.Deleted=0 and RO.RewardOptionID not in " & _
                            "(select RewardOptionID from CPE_DeliverableROIDs with (NoLock) as DR where DR.IncentiveID=" & IncentiveID & " and DR.Deleted=0);"
        'Send ("<!-- querystr=" & QueryStr & " -->")
        'Send ("right here")
        'End
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            'before we delete any deliverables from the orphaned ROID, delete the DeliverableROIDs records for those Deliverables
            MyCommon.QueryStr = "delete from CPE_DeliverableROIDs with (RowLock) where DeliverableID in " & _
                                "(select DeliverableID from CPE_Deliverables with (NoLock) where RewardOptionID=" & MyCommon.NZ(row.Item("RewardOptionID"), 0) & ");"
            MyCommon.LRT_Execute()
            'now delete the Deliverables for this ROID - may create more orphaned ROIDS ... that's why we are looping
            MyCommon.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & MyCommon.NZ(row.Item("RewardOptionID"), 0) & ";"
            MyCommon.LRT_Execute()
            'finally ... delete the orphaned ROID
            MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set Deleted=1, LastUpdate=getdate() where RewardOptionID=" & MyCommon.NZ(row.Item("RewardOptionID"), 0) & ";"
            MyCommon.LRT_Execute()
          Next
        Else
          Done = True
        End If
      End While
    Catch ex As Exception
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
  
  Sub Send_TouchpointRewards(ByVal OfferID As Long, ByVal ROID As Long, ByVal Phase As Integer, ByVal TierLevels As Integer, _
                             Optional ByVal IsTempOffer As Boolean = False, Optional ByVal FromTempOffer As Boolean = False)
    Dim MyCommon As New Copient.CommonInc
    Dim rst As DataTable
    Dim row As DataRow
    Dim RewardType As Long
    Dim DeliverableID As Long
    Dim Details As String
    Dim NumOfCols As Integer = 0
    
    Try
      IsTemplateOffer = IsTempOffer
      FromTemplateOffer = FromTempOffer
      
      MyCommon.AppName = "graphic-reward.aspx"
      
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "dbo.pa_CPE_GetTouchPointRewards"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = ROID
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      rst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          RewardType = MyCommon.NZ(row.Item("RewardType"), 0)
          DeliverableID = MyCommon.NZ(row.Item("PKID"), 0)
          Details = MyCommon.NZ(row.Item("Name"), "")
          If (GetSummaryOnly()) Then
          Else
            Send("<tr class=""shaded"">")
          End If
          Send(GetTpRewardLinks(RewardType, OfferID, DeliverableID, Details, Phase))
          If (GetSummaryOnly()) Then
          Else
            Send("</tr>")
          End If
        Next
      Else
        If (GetSummaryOnly()) Then
          Send("<i>" & Copient.PhraseLib.Lookup("CPE-rew-graphics.notouchpointreward", LanguageID) & "</i>")
        Else
          NumOfCols = IIf(IsTemplateOffer OrElse FromTemplateOffer, TierLevels + 3, TierLevels + 2)
          Send("<tr class=""shaded"">")
          Send("<td></td>")
          Send("<td colspan=""" & NumOfCols & """><i>" & Copient.PhraseLib.Lookup("CPE-rew-graphics.notouchpointreward", LanguageID) & "</i></td>")
          Send("</tr>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LRTsp()
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
  
  Function GetTpRewardLinks(ByVal RewardType As Long, ByVal OfferID As Long, ByVal DeliverableID As Long, ByVal Details As String, ByVal Phase As Integer) As String
    Dim link As String
    Select Case RewardType
      Case 1 ' Graphics
        link = GetTpGraphicLinks(DeliverableID, OfferID, Details, Phase)
      Case 2 ' Discounts
        link = GetTpDiscountLinks(DeliverableID, OfferID, Phase)
      Case 4 ' Printed Messages
        link = GetTpPrintedMsgLinks(DeliverableID, OfferID, Phase)
      Case 5 ' Membership
        link = GetTpMembershipLinks(DeliverableID, OfferID, Phase)
      Case 8 'Points
        link = GetTpPointsLinks(DeliverableID, OfferID, Phase)
      Case 9 ' Cashier Message
        link = GetTpCashierLinks(DeliverableID, OfferID, Phase)
      Case 11 ' Stored Value
        link = GetTpStoredValueLinks(DeliverableID, OfferID, Phase)
      Case Else
        link = "<td></td><td><a href=""#"">" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & " RT: " & RewardType & "</a></td><td></td><td></td>"
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          link += ("<td></td>")
        End If
    End Select
    Return link.ToString
  End Function
  
  Function GetTpGraphicLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Details As String, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc()
    Dim rst As DataTable
    Dim linkBuf As New StringBuilder()
    Dim AdID As Long = 0
    Dim ImageType As Integer = 1, CellID As Integer = 0
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select D.ScreenCellID as CellID, D.DisallowEdit, OSA.OnScreenAdID as AdID, OSA.Width, OSA.Height, OSA.ImageType as ImageType " & _
                          "from CPE_Deliverables D with (NoLock) inner join OnScreenAds OSA with (NoLock) on D.OutputID=OSA.OnScreenAdID " & _
                          "where D.DeliverableID=" & DeliverableID & " and OSA.Deleted=0;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        AdID = MyCommon.NZ(rst.Rows(0).Item("AdID"), 0)
        CellID = MyCommon.NZ(rst.Rows(0).Item("CellID"), 0)
        ImageType = MyCommon.NZ(rst.Rows(0).Item("ImageType"), 0)
        DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
        confirmDelLangId = IIf(Phase = 1, "notification.confirmdelete", "reward.confirmdelete")
        If (Phase = 1) Then
          DeleteURL = "CPEoffer-not.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & DeliverableID
        Else
          DeleteURL = "CPEoffer-rew.aspx?mode=DeleteGraphic&OfferID=" & OfferID & "&deliverableid=" & DeliverableID
        End If
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & ": ")
          linkBuf.Append("<a href=""graphic-edit.aspx?OnScreenAdId=" & AdID & """>" & MyCommon.SplitNonSpacedString(Details, 25) & "</a><br />")
        Else
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-graphic.aspx?OfferID=" & OfferID)
          linkBuf.Append("&ad=" & AdID & "&cellselect=" & CellID & "&imagetype=" & ImageType & "&preview=1&Phase=" & Phase)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.graphic", LanguageID))
          linkBuf.Append("</a><td><a href=""#""></a></td>")
          linkBuf.Append("<td><a href=""graphic-edit.aspx?OnScreenAdId=" & AdID & """>" & MyCommon.SplitNonSpacedString(Details, 25) & "</a> (" & rst.Rows(0).Item("Width") & " x " & rst.Rows(0).Item("Height"))
          If MyCommon.NZ(rst.Rows(0).Item("ImageType"), "") = 1 Then
            linkBuf.Append("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
          ElseIf MyCommon.NZ(rst.Rows(0).Item("ImageType"), "") = 2 Then
            linkBuf.Append("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
          End If
          linkBuf.Append(")</td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked9"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#""></a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    
    Return linkBuf.ToString
  End Function
  
  Function GetTpDiscountLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim rst3 As DataTable
    Dim row3 As DataRow
    Dim Details As StringBuilder
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select DISC.DiscountID, DISC.DiscountTypeID, D.DeliverableID, D.DisallowEdit, PT.Phrase as AmountTypeName, DISC.DiscountAmount, AT.AmountTypeID,  D.RewardOptionID, " & _
                          "DISC.DiscountedProductGroupID as SelectedPG, DISC.ExcludedProductGroupID as ExcludedPG, DISC.UserGroupID, PT2.Phrase as DiscountType " & _
                          "from CPE_Deliverables D with (NoLock) " & _
                          "inner join CPE_Discounts DISC with (NoLock) on D.OutputID=DISC.DiscountID " & _
                          "left join CPE_AmountTypes AT with (NoLock) on AT.AmountTypeID=DISC.AmountTypeID " & _
                          "inner join PhraseText PT with (NoLock) on PT.PhraseID=AT.PhraseID " & _
                          "left join CPE_DiscountTypes DT with (NoLock) on DT.DiscountTypeID=DISC.DiscountTypeID " & _
                          "inner join PhraseText PT2 with (NoLock) on PT2.PhraseID=DT.PhraseID " & _
                          "where D.RewardOptionPhase=" & Phase & " and D.DeliverableID=" & DeliverableID & " and D.DeliverableTypeID=2;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.discount", LanguageID) & ": ")
          Details = New StringBuilder(200)
          Select Case (MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0))
            Case 1, 5
              Details.Append(FormatCurrency(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "")) & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            Case 3
              Details.Append(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "") & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            Case 4
              Details.Append(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
            Case 2, 6
              Details.Append(FormatCurrency(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "")) & "&nbsp;")
            Case Else
              Details.Append(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), ""))
          End Select
          If MyCommon.NZ(rst.Rows(0).Item("SelectedPG"), 0) = 0 Then
            Details.Append(StrConv(Copient.PhraseLib.Lookup("term.nothing", LanguageID), VbStrConv.Lowercase))
          Else
            MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID = " & rst.Rows(0).Item("SelectedPG")
            rst3 = MyCommon.LRT_Select()
            For Each row3 In rst3.Rows
              If rst.Rows(0).Item("SelectedPG") = 1 Then
                Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase) & "&nbsp;")
              Else
                Details.Append("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(rst.Rows(0).Item("SelectedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
              End If
            Next
          End If
          If MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 2 Then
            Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.department", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 3 Then
            Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.basket", LanguageID), VbStrConv.Lowercase) & ")")
          End If
          linkBuf.Append(Details.ToString & "<br />")
        Else
          If (Phase = 1) Then
            DeleteURL = "CPEoffer-not.aspx?mode=DeleteDiscount&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&DiscountID=" & rst.Rows(0).Item("DiscountID")
          Else
            DeleteURL = "CPEoffer-rew.aspx?mode=DeleteDiscount&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&DiscountID=" & rst.Rows(0).Item("DiscountID")
          End If
          confirmDelLangId = IIf(Phase = 1, "notification.confirmdelete", "reward.confirmdelete")
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-discount.aspx?OfferID=" & OfferID & "&RewardID=" & rst.Rows(0).Item("RewardOptionID") & "&DiscountID=" & MyCommon.NZ(rst.Rows(0).Item("DiscountID"), 0) & "&DeliverableID=" & DeliverableID)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.discount", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td>" & MyCommon.NZ(rst.Rows(0).Item("AmountTypeName"), "") & "</td>")
          linkBuf.Append("<td>")
          Details = New StringBuilder(200)
          Select Case (MyCommon.NZ(rst.Rows(0).Item("AmountTypeID"), 0))
            Case 1, 5
              Details.Append(FormatCurrency(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "")) & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            Case 3
              Details.Append(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "") & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            Case 4
              Details.Append(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
            Case 2, 6
              Details.Append(FormatCurrency(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), "")) & "&nbsp;")
            Case Else
              Details.Append(MyCommon.NZ(rst.Rows(0).Item("DiscountAmount"), ""))
          End Select
          If MyCommon.NZ(rst.Rows(0).Item("SelectedPG"), 0) = 0 Then
            Details.Append(StrConv(Copient.PhraseLib.Lookup("term.nothing", LanguageID), VbStrConv.Lowercase))
          Else
            MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID = " & rst.Rows(0).Item("SelectedPG")
            rst3 = MyCommon.LRT_Select()
            For Each row3 In rst3.Rows
              If rst.Rows(0).Item("SelectedPG") = 1 Then
                Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase) & "&nbsp;")
              Else
                Details.Append("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(rst.Rows(0).Item("SelectedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
              End If
            Next
          End If
          If MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 2 Then
            Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.department", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf MyCommon.NZ(rst.Rows(0).Item("DiscountTypeID"), 0) = 3 Then
            Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.basket", LanguageID), VbStrConv.Lowercase) & ")")
          End If
          linkBuf.Append(Details.ToString & "</td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked10"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td><td><a href=""#"">")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID) & " ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.discount", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Function GetTpPrintedMsgLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim UrlTokens As String = ""
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String
    Dim Details As StringBuilder
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID, D.RewardOptionID, D.DisallowEdit " & _
                          "from CPE_Deliverables D with (NoLock) " & _
                          "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                          "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                          "where D.RewardOptionPhase=" & Phase & " and D.DeliverableID=" & DeliverableID & " and D.DeliverableTypeID=4 and PMT.TierLevel=1;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        If (Phase = 1) Then
          DeleteURL = "CPEoffer-not.aspx?mode=DeletePrintedMsg&OfferID=" & OfferID & "&MessageID=" & rst.Rows(0).Item("MessageID") & "&DeliverableID=" & DeliverableID & "&Phase=" & Phase
          confirmDelLangId = "notification.confirmdelete"
        ElseIf (Phase = 2) Then
          confirmDelLangId = "accumulation.confirmdelete"
        Else
          DeleteURL = "CPEoffer-rew.aspx?mode=DeletePrintedMsg&OfferID=" & OfferID & "&MessageID=" & rst.Rows(0).Item("MessageID") & "&DeliverableID=" & DeliverableID & "&Phase=" & Phase
          confirmDelLangId = "reward.confirmdelete"
        End If
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & ": ")
          Details = New StringBuilder(200)
          Details.Append(MyCommon.NZ(rst.Rows(0).Item("BodyText"), ""))
          If (Details.ToString().Length > 30) Then
            Details = Details.Remove(27, (Details.Length - 27))
            Details.Append("...")
          End If
          linkBuf.Append("""" & Details.ToString & """<br />")
        Else
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-pmsg.aspx?OfferID=" & OfferID & "&RewardID=" & rst.Rows(0).Item("RewardOptionID") & "&MessageID=" & rst.Rows(0).Item("MessageID") & "&Phase=" & Phase)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.printedmessage", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td>" & GetMessageTypeName(MyCommon.NZ(rst.Rows(0).Item("MessageTypeID"), 0)) & "</td>")
          Details = New StringBuilder(200)
          Details.Append(MyCommon.NZ(rst.Rows(0).Item("BodyText"), ""))
          If (Details.ToString().Length > 60) Then
            Details = Details.Remove(57, (Details.Length - 57))
            Details.Append("...")
          End If
          linkBuf.Append("<td>""" & MyCommon.SplitNonSpacedString(Details.ToString, 25) & """</td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked11"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#"">" & Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & ": ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Function GetTpMembershipLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim UrlTokens As String = ""
    Dim DeleteURL As String
    Dim confirmDelLangId As String = "reward.confirmdelete"
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select CG.Name as Name, CG.CustomerGroupID, D.DeliverableID, D.DeliverableTypeID, D.RewardOptionID as ROID, D.DisallowEdit " & _
                          "from CustomerGroups as CG with (NoLock) " & _
                          "inner join CPE_Deliverables as D with (NoLock) on D.OutputID=CG.CustomerGroupID " & _
                          "and CG.Deleted=0 and D.DeliverableID=" & DeliverableID & " and D.DeliverableTypeID in (5) and D.RewardOptionPhase=" & Phase & " " & _
                          "order by D.DeliverableTypeID, CG.Name;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        UrlTokens = "?OfferID=" & OfferID & "&RewardID=" & MyCommon.NZ(rst.Rows(0).Item("ROID"), 0) & "&DeliverableID=" & MyCommon.NZ(rst.Rows(0).Item("DeliverableID"), -1) & "&GroupID=" & MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), 0) & "&Phase=" & Phase & "&action=" & MyCommon.NZ(rst.Rows(0).Item("DeliverableTypeID"), 0)
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.membership", LanguageID) & ": ")
          linkBuf.Append("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a><br />")
        Else
          If (Phase = 1) Then
            DeleteURL = "CPEoffer-not.aspx" & UrlTokens & "&mode=DeleteMembership"
          Else
            DeleteURL = "CPEoffer-rew.aspx" & UrlTokens & "&mode=DeleteMembership"
          End If
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-membership.aspx" & UrlTokens)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.membership", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td></td>")
          linkBuf.Append("<td><a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(rst.Rows(0).Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a></td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked12"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#"">" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & ": ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Function GetTpPointsLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim Details As String = ""
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String = "reward.confirmdelete"
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select PP.ProgramName as Name, PP.ProgramID, D.DeliverableID as PKID, D.DisallowEdit, DP.Quantity, D.RewardOptionID " & _
                          "from PointsPrograms as PP with (NoLock) " & _
                          "inner join CPE_DeliverablePoints as DP with (NoLock) on PP.ProgramID=DP.ProgramID and PP.Deleted=0 " & _
                          "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DP.DeliverableID and D.RewardOptionPhase=" & Phase & " " & _
                          "where D.DeliverableID=" & DeliverableID & " order by PP.ProgramName;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        If (Phase = 1) Then
          DeleteURL = "CPEoffer-not.aspx?mode=DeletePoints&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&ProgramID=" & rst.Rows(0).Item("ProgramID")
        Else
          DeleteURL = "CPEoffer-rew.aspx?mode=DeletePoints&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&ProgramID=" & rst.Rows(0).Item("ProgramID")
        End If
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.points", LanguageID) & ": ")
          linkBuf.Append("<a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(rst.Rows(0).Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a> ")
          linkBuf.Append("(" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), "") & ")<br />")
        Else
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-point.aspx?OfferID=" & OfferID & "&RewardID=" & rst.Rows(0).Item("RewardOptionID") & "&ProgramID=" & MyCommon.NZ(rst.Rows(0).Item("ProgramID"), 0) & "&quantity=" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), 0) & "&DeliverableID=" & DeliverableID)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.points", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td></td>")
          linkBuf.Append("<td>" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), "") & " " & Copient.PhraseLib.Lookup("term.awardedinprogram", LanguageID) & " ")
          linkBuf.Append("<a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(rst.Rows(0).Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a></td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked13"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#"">" & Copient.PhraseLib.Lookup("term.points", LanguageID) & ": ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Function GetTpStoredValueLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim Details As String = ""
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String = "reward.confirmdelete"
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select SVP.Name as Name, SVP.SVProgramID, D.DeliverableID as PKID, D.DisallowEdit, DSV.Quantity, D.RewardOptionID " & _
                          "from StoredValuePrograms as SVP with (NoLock) " & _
                          "inner join CPE_DeliverableStoredValue as DSV with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " & _
                          "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DSV.DeliverableID and D.RewardOptionPhase=" & Phase & " " & _
                          "where D.DeliverableID=" & DeliverableID & " order by SVP.Name;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        If (Phase = 1) Then
          DeleteURL = "CPEoffer-not.aspx?mode=DeleteStoredValue&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&ProgramID=" & rst.Rows(0).Item("SVProgramID")
        Else
          DeleteURL = "CPEoffer-rew.aspx?mode=DeleteStoredValue&OfferID=" & OfferID & "&DeliverableID=" & DeliverableID & "&ProgramID=" & rst.Rows(0).Item("SVProgramID")
        End If
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & ": ")
          linkBuf.Append("<a href=""SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a> ")
          linkBuf.Append("(" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), "") & ")<br />")
        Else
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & "');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-sv.aspx?OfferID=" & OfferID & "&RewardID=" & rst.Rows(0).Item("RewardOptionID") & "&ProgramID=" & MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0) & "&quantity=" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), 0) & "&DeliverableID=" & DeliverableID)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td></td>")
          linkBuf.Append("<td>" & MyCommon.NZ(rst.Rows(0).Item("Quantity"), "") & " " & Copient.PhraseLib.Lookup("term.awardedinprogram", LanguageID) & " ")
          linkBuf.Append("<a href=""SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(rst.Rows(0).Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(rst.Rows(0).Item("Name"), 25) & "</a></td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked13"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & ": ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Function GetTpCashierLinks(ByVal DeliverableID As Long, ByVal OfferID As Long, ByVal Phase As Integer) As String
    Dim MyCommon As New Copient.CommonInc
    Dim linkBuf As New StringBuilder()
    Dim rst As DataTable
    Dim Details As String = ""
    Dim DeleteURL As String = ""
    Dim confirmDelLangId As String
    Dim DisallowEdit As Boolean = False
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select D.DeliverableID, D.RewardOptionID, D.DisallowEdit, CM.MessageID, CMT.Line1, CMT.Line2 " & _
                          "from CPE_Deliverables D with (NoLock) " & _
                          "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                          "inner join CPE_CashierMessageTiers CMT with (NoLock) on CM.MessageID=CMT.MessageID " & _
                          "where D.DeliverableID=" & DeliverableID & " and DeliverableTypeID=9 and D.RewardOptionPhase=" & Phase & " and CMT.TierLevel=1;"
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        If (Phase = 1) Then
          DeleteURL = "CPEoffer-not.aspx?mode=DeleteCashierMsg&OfferID=" & OfferID & "&MessageID=" & rst.Rows(0).Item("MessageID") & "&DeliverableID=" & DeliverableID & "&Phase=" & Phase
          confirmDelLangId = "notification.confirmdelete"
        Else
          DeleteURL = "CPEoffer-rew.aspx?mode=DeleteCashierMsg&OfferID=" & OfferID & "&MessageID=" & rst.Rows(0).Item("MessageID") & "&DeliverableID=" & DeliverableID & "&Phase=" & Phase
          confirmDelLangId = "reward.confirmdelete"
        End If
        If (GetSummaryOnly()) Then
          linkBuf.Append("○ " & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & ": ")
          linkBuf.Append("""" & rst.Rows(0).Item("Line1") & "<span class=""grey""> / </span>" & rst.Rows(0).Item("Line2") & """<br />")
        Else
          DisallowEdit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
          RewardDisabled = IIf((OfferEditable And Not (FromTemplateOffer And DisallowEdit)), "", " disabled=""disabled""")
          linkBuf.Append("<td><input type=""button"" class=""ex"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & RewardDisabled & " onClick=""if(confirm('" & Copient.PhraseLib.Lookup(confirmDelLangId, LanguageID) & "')) LoadDocument('" & DeleteURL & " ');"" value=""X"" /></td>")
          linkBuf.Append("<td><a href=""javascript:openPopup('CPEoffer-rew-cmsg.aspx?OfferID=" & OfferID & "&RewardID=" & rst.Rows(0).Item("RewardOptionID") & "&DeliverableID=" & DeliverableID & "&Phase=" & Phase)
          linkBuf.Append("')"">")
          linkBuf.Append(Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID))
          linkBuf.Append("</a></td>")
          linkBuf.Append("<td></td>")
          linkBuf.Append("<td>""" & rst.Rows(0).Item("Line1") & "<br />" & rst.Rows(0).Item("Line2") & """</td>")
          If (IsTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">")
            linkBuf.Append("  <input type=""checkbox"" id=""chkLocked14"" name=""chkLocked"" value=""" & DeliverableID & """" & IIf(DisallowEdit, " checked=""checked""", "") & " onclick=""javascript:updateLocked('lockTP" & DeliverableID & "', this.checked);"" />")
            linkBuf.Append("  <input type=""hidden"" id=""rewTP" & DeliverableID & """ name=""rew"" value=""" & DeliverableID & """ />")
            linkBuf.Append("  <input type=""hidden"" id=""lockTP" & DeliverableID & """ name=""locked"" value=""" & IIf(DisallowEdit, "1", "0") & """ />")
            linkBuf.Append("</td>")
          ElseIf (FromTemplateOffer) Then
            linkBuf.Append("<td class=""templine"">" & IIf(DisallowEdit, "Yes", "No") & "</td>")
          End If
        End If
      Else
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td><a href=""#"">" & Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & ": ")
        linkBuf.Append(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
        linkBuf.Append("</a></td>")
        linkBuf.Append("<td></td>")
        linkBuf.Append("<td></td>")
        If (IsTemplateOffer OrElse FromTemplateOffer) Then
          linkBuf.Append("<td></td>")
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    Return linkBuf.ToString
  End Function
  
  Sub Send_TPRewardOptions(ByVal OfferID As Long, ByVal ROID As Long)
    Dim MyCommon As New Copient.CommonInc
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim TouchResponse As Boolean = False
    Dim ParentROID As Long
    Dim NumSelectedUserGroups As Integer
    Dim TPGraphic As Boolean = False
    Dim TPDiscount As Boolean = False
    Dim TPPrinted As Boolean = False
    Dim TPMembership As Boolean = False
    Dim TPPoints As Boolean = False
    Dim TPCashier As Boolean = False
    Dim TPFranking As Boolean = False
    Dim TPStoredValue As Boolean = False
    Dim TPPassThru As Boolean = False
    Dim EngineSubType As Integer = -1
    
    Try
      MyCommon.AppName = "graphic-reward.aspx"
      MyCommon.Open_LogixRT()
      
      MyCommon.QueryStr = "select RO.TouchResponse from CPE_RewardOptions RO with (NoLock) where RO.RewardOptionID=" & ROID & ";"
      rst = MyCommon.LRT_Select()
      
      If (rst.Rows.Count > 0) Then
        TouchResponse = MyCommon.NZ(rst.Rows(0).Item("TouchResponse"), False)
      End If
      
      If (TouchResponse) Then 'get the parent
        MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          ParentROID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
        End If
      End If
      
      ' Check if there are already touchpoint rewards of any given type
      MyCommon.QueryStr = "dbo.pa_CPE_GetTouchPointRewards"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = ROID
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = 3
      rst2 = MyCommon.LRTsp_select()
      If rst2.Rows.Count > 0 Then
        For Each row2 In rst2.Rows
          If row2.Item("RewardType") = 1 Then TPGraphic = True
          If row2.Item("RewardType") = 2 Then TPDiscount = True
          If row2.Item("RewardType") = 4 Then TPPrinted = True
          If row2.Item("RewardType") = 5 Then TPMembership = True
          If row2.Item("RewardType") = 6 Then TPMembership = True
          If row2.Item("RewardType") = 8 Then TPPoints = True
          If row2.Item("RewardType") = 9 Then TPCashier = True
          If row2.Item("RewardType") = 10 Then TPFranking = True
          If row2.Item("RewardType") = 11 Then TPStoredValue = True
          If row2.Item("RewardType") = 12 Then TPPassThru = True
        Next
      End If
      MyCommon.Close_LRTsp()
      rst2 = Nothing
      
      MyCommon.QueryStr = "select RO.IncentiveID from CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                          "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " & _
                          "where IncentiveID=" & OfferID & " and RO.Deleted=0;"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        TPDiscount = True
      End If
      
      'see if the RewardOption has any assoicated UserGroups
      MyCommon.QueryStr = "select count(*) as NumRecs from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & ParentROID & " and Deleted=0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        NumSelectedUserGroups = MyCommon.NZ(rst.Rows(0).Item("NumRecs"), 0)
      End If
      
      MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives where IncentiveID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If(rst.Rows.Count > 0) Then
        EngineSubType = rst.Rows(0).Item("EngineSubTypeID")
      End If
      
      MyCommon.QueryStr = "SELECT PECT.EngineID, PECT.ComponentTypeID, DT.DeliverableTypeID, DT.Description, DT.PhraseID, PECT.Singular, " & _
                          "  CASE DeliverableTypeID " & _
                          "    WHEN 1 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=1) " & _
                          "    WHEN 2 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=2) " & _
                          "    WHEN 4 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=4) " & _
                          "    WHEN 5 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=5) " & _
                          "    WHEN 6 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=6) " & _
                          "    WHEN 7 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=7) " & _
                          "    WHEN 8 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=8) " & _
                          "    WHEN 9 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=9) " & _
                          "    WHEN 10 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=10) " & _
                          "    WHEN 11 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=11) " & _
                          "    WHEN 12 THEN (SELECT COUNT(*) FROM CPE_Deliverables WITH (NOLOCK) where RewardOptionID=" & ParentROID & " and Deleted=0 and RewardOptionPhase=3 and DeliverableTypeID=12) " & _
                          "    ELSE 0 " & _
                          "  END as ItemCount " & _
                          "FROM PromoEngineComponentTypes AS PECT " & _
                          "INNER JOIN CPE_DeliverableTypes AS DT ON DT.DeliverableTypeID=PECT.LinkID " & _
                          "WHERE EngineID=2 AND PECT.ComponentTypeID=2 AND Enabled=1 AND Touchpoint=1"
      
      If (EngineSubType > -1) Then
        MyCommon.QueryStr &= " AND EngineSubTypeID=" & EngineSubType
      End If                    
      'Impose a few special limits on the query based on various factors:
      If (Not TouchResponse) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>1"
      End If
      If (NumSelectedUserGroups = 0 OrElse TPDiscount) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>2"
      End If
      If (NumSelectedUserGroups = 0 OrElse Not TouchResponse OrElse TPPrinted) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>4"
      End If
      If ((NumSelectedUserGroups = 0) And (Not TouchResponse)) OrElse (TPMembership) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>5"
      End If
      If (NumSelectedUserGroups = 0 OrElse TPPoints) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>8"
      End If
      If (NumSelectedUserGroups = 0 OrElse TPCashier) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>9"
      End If
      If (NumSelectedUserGroups = 0 OrElse TPCashier) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>10"
      End If
      If (NumSelectedUserGroups = 0 OrElse TPStoredValue) Then
        MyCommon.QueryStr &= " AND DeliverableTypeID<>11"
      End If
      MyCommon.QueryStr &= " ORDER BY DisplayOrder;"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
          Send("<option value=""" & row.Item("DeliverableTypeID") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
        Next
      End If
      
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
  
  Function GetMessageTypeName(ByVal MessageTypeID As Long) As String
    Dim MessageTypeName As String = ""
    Dim MyCommon As New Copient.CommonInc
    Dim rst As DataTable
    MyCommon.Open_LogixRT()
    MyCommon.QueryStr = "select TypeID, PhraseID from PrintedMessageTypes with (NoLock) where TypeID = " & MessageTypeID & " order by TypeID;"
    rst = MyCommon.LRT_Select()
    If rst.rows.count > 0 Then
      MessageTypeName = Copient.PhraseLib.Lookup(rst.Rows(0).Item("PhraseID"), LanguageID)
    Else
      MessageTypeName = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
    End If
    MyCommon.Close_LogixRT()
    Return MessageTypeName
  End Function
  
  Function ReplaceTags(ByVal str As String) As String
    Dim i As Integer = 0
    Dim c As String
    Dim startTag As Boolean = False
    Dim endTag As Boolean = False
    Dim strBuilder As New StringBuilder()
    For i = 0 To str.Length - 1
      c = str.Substring(i, 1)
      If (c = "|") Then
        If (startTag) Then
          endTag = True
          startTag = False
          strBuilder.Append(" ")
        Else
          startTag = True
          endTag = False
        End If
      Else
        If Not (startTag) Then
          strBuilder.Append(c)
        End If
      End If
    Next
    Return strBuilder.ToString
  End Function
</script>
