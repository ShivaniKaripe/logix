﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %><%@ Import
    Namespace="Newtonsoft.Json" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%
    Dim CopientFileName As String = "OfferFeeds.aspx"
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
  
    Dim AdminUserID As Integer
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
  
    MyCommon.AppName = "OfferFeeds.aspx"
    CurrentRequest.Resolver.AppName = MyCommon.AppName
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    If (Request.Form("Mode") = "AllowSpecialCharactersCM") Or (Request.Form("Mode") = "AllowSpecialCharactersUE") Then
        Dim sessionvalue As String = Request.Form("OfferDescription")
        Dim offerIdSpecial As String = Request.Form("OfferID")
        sessionvalue = Mid(sessionvalue, 1, 1000)
        If (Request.Form("Mode") = "AllowSpecialCharactersCM") Then
            MyCommon.QueryStr = "update Offers with (RowLock) set Description='" & MyCommon.Parse_Quotes(sessionvalue) & "' where OfferID=" & offerIdSpecial
            MyCommon.LRT_Execute()
        ElseIf (Request.Form("Mode") = "AllowSpecialCharactersUE") Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set Description='" & MyCommon.Parse_Quotes(sessionvalue) & "' where IncentiveID=" & offerIdSpecial
            MyCommon.LRT_Execute()
        End If
    End If

    Response.Expires = 0
    Response.Clear()
    Response.ContentType = "text/html"
    Select Case Request.QueryString("Mode")
        Case "GetBanners"
            SendBanners(MyCommon.Extract_Val(Request.QueryString("OfferID")), AdminUserID)
        Case "SaveBanners"
            SaveBanners(MyCommon.Extract_Val(Request.QueryString("OfferID")), AdminUserID)
        Case "SendPopup"
            SendPopup(MyCommon.Extract_Val(Request.QueryString("OfferID")), AdminUserID, Request.QueryString("PageName"), _
                      Request.QueryString("Tokens"))
        Case "FavoriteForAll"
            FavoriteOfferForAll(MyCommon.Extract_Val(Request.QueryString("OfferID")), AdminUserID)
        Case "ProductGroups"
            Dim Disqualifier As Boolean = False
            If (Request.QueryString("Disqualifier") <> "" AndAlso (Request.QueryString("Disqualifier").ToLower() = "true" Or Request.QueryString("Disqualifier") = "1")) Then Disqualifier = True
            FindProductGroups(Request.QueryString("ProductSearch"), MyCommon.Extract_Val(Request.QueryString("ROID")), Disqualifier, MyCommon.Extract_Val(Request.QueryString("SelectedGroup")), _
                              MyCommon.Extract_Val(Request.QueryString("ExcludedGroup")), Request.QueryString("SearchRadio"))
        Case "DiscountProductGroups"
            DiscountProductGroups(Request.QueryString("ProductSearch"), MyCommon.Extract_Val(Request.QueryString("SelectedGroup")), Request.QueryString("SearchRadio"))
        Case "ConditionCustomerGroups"
            Dim AnyCustomerEnabled As Boolean = False
            If (Request.QueryString("AnyCustomerEnabled") <> "" AndAlso Request.QueryString("AnyCustomerEnabled").ToLower() = "true") Then AnyCustomerEnabled = True
            ConditionCustomerGroups(Request.QueryString("Search"), MyCommon.Extract_Val(Request.QueryString("EngineID")), AnyCustomerEnabled, _
                                    Request.QueryString("SelectedGroups"), Request.QueryString("ExcludedGroups"), MyCommon.Extract_Val(Request.QueryString("OfferID")), Request.QueryString("SearchRadio"))
        Case "GrantMembership"
            Dim strSelectedGroup As String = String.Empty
            If (Request.QueryString("GroupCount") IsNot Nothing) Then
                Dim count As Integer
                count = Convert.ToInt32(Request.QueryString("GroupCount"))
                For index = 1 To count
                    If index < count Then
                        strSelectedGroup = strSelectedGroup & Request.QueryString("Group" & index) & ","
                    Else
                        strSelectedGroup = strSelectedGroup & Request.QueryString("Group" & index)
                    End If
                Next
            End If
            GrantMembershipGroups(Request.QueryString("OfferID"), Request.QueryString("RewardID"), Request.QueryString("Search"), MyCommon.Extract_Val(Request.QueryString("EngineID")), Request.QueryString("SearchRadio"), strSelectedGroup)
        Case "GetProductCollisions"
            GetProductCollisions(MyCommon.Extract_Val(Request.QueryString("OfferID")))
        Case "GetProductCollisionsUE"
            GetProductCollisionsUE(MyCommon.Extract_Val(Request.QueryString("OfferID")), Request.QueryString("DeferDeploy"), ApprovalType:=MyCommon.Extract_Val(Request.QueryString("ApprovalType")))
        Case "GetProductCollisionsUEDetection"
            GetProductCollisionsUE(MyCommon.Extract_Val(Request.QueryString("OfferID")), Request.QueryString("DeferDeploy"), Request.QueryString("CallingLocation"))
        Case "ProductCollisionsBackgroundUEDetection"
            ProcessProductCollisionsBackgroundUE(MyCommon.Extract_Val(Request.QueryString("OfferID")), Request.QueryString("DeferDeploy"), Request.QueryString("CallingLocation"))
        Case "ProductCollisionsBackgroundUE"
            ProcessProductCollisionsBackgroundUE(MyCommon.Extract_Val(Request.QueryString("OfferID")), Request.QueryString("DeferDeploy"), ApprovalType:=MyCommon.Extract_Val(Request.QueryString("ApprovalType")))
        Case "ApproveOffer"
            ApproveOffer(MyCommon.Extract_Val(Request.QueryString("OfferID")), MyCommon.Extract_Val(Request.QueryString("ApprovalType")), MyCommon.Extract_Val(Request.QueryString("OCDEnabled")))
        Case "PredefinedReceiptText"
            GeneratePredefinedRecTextMessages(Request.QueryString("baselangtext"), Request.QueryString("mlClickedID"))
        Case "ProductGroupsCM"
            Dim CallingPage As String = String.Empty
            If Not Request.QueryString("CallingPage") Is Nothing And Not String.IsNullOrWhiteSpace(Request.QueryString("CallingPage")) Then
                CallingPage = Request.QueryString("CallingPage")
            End If
            FindProductGroupsCM(Request.QueryString("ProductSearch"), MyCommon.Extract_Val(Request.QueryString("OfferID")), MyCommon.Extract_Val(Request.QueryString("SelectedGroup")), _
                          MyCommon.Extract_Val(Request.QueryString("ExcludedGroup")), Request.QueryString("SearchRadio"), CallingPage)
        Case "UDFStringValue"
            Dim sUdfValue As String
            If IsNothing(Request.QueryString("UDFValue")) Then
                sUdfValue = GetUDFStringValues(MyCommon.Extract_Val(Request.QueryString("UDFPK")), MyCommon.Extract_Val(Request.QueryString("OfferID")))
                Send(sUdfValue)
            Else
                sUdfValue = Request.Form("UDFValue")
                UpdateUDFStringValues(MyCommon.Extract_Val(Request.QueryString("UDFPK")), MyCommon.Extract_Val(Request.QueryString("OfferID")), sUdfValue)
            End If
        Case "NoOfDuplicateOffers"
            Try
                DuplicateoffersFromTemplete(MyCommon.Extract_Val(Request.QueryString("OfferID")), MyCommon.Extract_Val(Request.QueryString("EngineID")), Request.QueryString("DuplicateCnt"))
            Catch ex As Exception
                Dim err As String = ""
                If ex.Message = "error.couldnot-processoffers" Then
                    err = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
                Else
                    err = ex.Message
                End If
                Send(err)
            End Try
            
        Case "StoredValueProgramsCM"
            FindStoredValueProgramsCM(Request.QueryString("ProgramSearch"), MyCommon.Extract_Val(Request.QueryString("SelectedProgram")), Request.QueryString("SearchRadio"))
        Case "PointsProgramsCM"
            FindPointsProgramsCM(Request.QueryString("ProgramSearch"), MyCommon.Extract_Val(Request.QueryString("SelectedProgram")), Request.QueryString("SearchRadio"))
        Case "IsDeployableOffer"
            Dim offerDeploymentValidator As IOfferDeploymentValidator = CurrentRequest.Resolver.Resolve(Of IOfferDeploymentValidator)()
            Dim Errormsg As String = String.Empty
			Dim DeployType As String = ""
            Dim amsresult As AMSResult(Of Boolean) = New AMSResult(Of Boolean)
            Dim DeploymentSkip As Boolean = False
            Dim DeferDeploy As Boolean = Convert.ToBoolean(Request.QueryString("DeferDeploy"))
			If GetCgiValue("deferdeploy") <> "" Then
                DeployType = "deferdeploy"
            Else
                DeployType = "deploy"
            End If
            amsresult = offerDeploymentValidator.ValidateCPEOffer(Request.QueryString("OfferID"), DeploymentSkip, DeferDeploy, False, False, False, AdminUserID)
            If (amsresult.ResultType = AMSResultType.Success) Then
                If (amsresult.Result = True) Then
                    Send(amsresult.Result)
                Else
                    Dim validPhrase As String = Nothing
                    If MyCommon.Fetch_SystemOption("124") = 1 AndAlso DeploymentSkip = False Then
                        MyCommon.QueryStr = "dbo.pa_ReqTranslationCheck"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.NVarChar, 200).Value = Request.QueryString("OfferID")
                        MyCommon.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.NVarChar, 2000).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        validPhrase = MyCommon.LRTsp.Parameters("@ErrorTerm").Value
                    End If
                    If validPhrase <> "" Then
                        Send(amsresult.MessageString & " <input type=""submit"" class=""regular"" id=""" & DeployType & """ name=""" & DeployType & """ value=""" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & """ onclick=""document.getElementById('deploytransreqskip').value='1';"" />")
                    Else
                        Send(amsresult.MessageString)
                    End If
                End If
            Else
                Send(Copient.PhraseLib.Lookup(amsresult.MessageString, LanguageID))
            End If
        Case "StoreUserPane"
            StoreUserPane(MyCommon.Extract_Val(Request.QueryString("UserID")))
        Case "SaveStoreUserLocations"
            SaveStoreUserLocations(MyCommon.Extract_Val(Request.QueryString("UserID")), Request.QueryString("LocationList"))
        Case "StoreUserLocationSearch"
            StoreUserLocationSearch(MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("UserID"))), Server.HtmlEncode(Request.QueryString("searchterms")), MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchoption"))))
        Case "GetSVProgDescription"
            GetDespForSVProgID(Request.Form("SVProgID"))
        Case "CreateGroupOrProgramFromOffer"
            CreateGroupOrProgramFromOffer(Request.QueryString("CreateType"), Request.QueryString("Name"))
        Case "GetCardTypeForCustomerCondition"
            GetCardTypeForCustomerCondition(MyCommon.Extract_Val(Request.QueryString("conditionID")))
        Case "CGReportCondition"
            CGReportCondition(Request.QueryString("Search"), Request.QueryString("SearchRadio"))
        Case "OffReportCondition"
            OffReportCondition(Request.QueryString("Search"), Request.QueryString("SearchRadio"))
        Case Else
            Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
            Send(Request.RawUrl)
    End Select
  If Response.IsClientConnected() Then
    Response.Flush()
  End if
    Response.End()
%>
<script runat="server">
    Public DefaultLanguageID
    Public MyCommon As New Copient.CommonInc
    Dim iOfferID_CDS As Long, DeferDeployment_CDS As Boolean
    Dim ApprovalType_CDS As Integer

    Function ValidateName(ByVal CreateType As String, ByVal Name As String) As String
        Dim validationMessage As String = ""
        Try
            Select Case CreateType
                Case "CustomerGroup"
                    If (Name = "") Then
                        validationMessage = Copient.PhraseLib.Lookup("cgroup-edit.noname", LanguageID)
                    End If
                Case "ProductGroup"
                    If Name.Length > 190 Then
                        validationMessage = Copient.PhraseLib.Lookup("pgroup-edit.nametoolong", LanguageID)
                    End If
                    If (Name = "") Then
                        validationMessage = Copient.PhraseLib.Lookup("pgroup-edit.noname", LanguageID)
                    End If
                Case "Points"
                    If (Name = "") Then
                        validationMessage = Copient.PhraseLib.Lookup("point-edit.noname", LanguageID)
                    End If
                Case "StoredValue"
                    If (Name = "") Then
                        validationMessage = Copient.PhraseLib.Lookup("sv-no-name", LanguageID)
                    End If
                Case "Location"
                    If (Name = "") Then
                        validationMessage = Copient.PhraseLib.Lookup("lgroup-edit.noname", LanguageID)
                    End If
            End Select
        Catch ex As Exception

        End Try
        Return validationMessage
    End Function
    Function DoesGroupOrProgramExists(ByVal CreateType As String, ByVal Name As String) As Integer
        Dim id As Integer = -1
        Dim dst As DataTable = Nothing
        Try
            Select Case CreateType
                Case "CustomerGroup"
                    MyCommon.QueryStr = "SELECT Name, CustomerGroupID FROM CustomerGroups WITH (NoLock) WHERE Name='" & MyCommon.Parse_Quotes(Name) & "' AND Deleted=0 ;"
                    dst = MyCommon.LRT_Select

                    If (dst.Rows.Count > 0) Then
                        id = MyCommon.NZ(dst.Rows(0).Item("CustomerGroupID"), -1)
                    End If

                Case "ProductGroup"
                    MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = @Name AND Deleted=0"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 200).Value = MyCommon.Parse_Quotes(Name)
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (dst.Rows.Count > 0) Then
                        id = MyCommon.NZ(dst.Rows(0).Item("ProductGroupID"), -1)
                    End If

                Case "Points"
                    MyCommon.QueryStr = "select ProgramID from PointsPrograms with (NoLock) where ProgramName = '" & MyCommon.Parse_Quotes(Name) & "' and Deleted=0;"
                    dst = MyCommon.LRT_Select

                    If (dst.Rows.Count > 0) Then
                        id = MyCommon.NZ(dst.Rows(0).Item("ProgramID"), -1)
                    End If

                Case "StoredValue"
                    MyCommon.QueryStr = "SELECT SVProgramID FROM StoredValuePrograms with (NoLock) WHERE Name='" & MyCommon.Parse_Quotes(Name) & "' AND Deleted=0;"
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        id = MyCommon.NZ(dst.Rows(0).Item("SVProgramID"), -1)
                    End If

                Case "Location"
                    MyCommon.QueryStr = "SELECT LocationGroupID FROM LocationGroups WHERE Name = '" & MyCommon.Parse_Quotes(Name) & "' AND Deleted=0;"
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        id = MyCommon.NZ(dst.Rows(0).Item("LocationGroupID"), -1)
                    End If

            End Select
        Catch ex As Exception

        End Try
        Return id
    End Function
    Sub CreateGroupOrProgramFromOffer(ByVal CreateType As String, ByVal Name As String)
        Dim ResponseText As String = String.Empty
        Dim id As Integer = -1
        Dim Logix As New Copient.LogixInc
        Dim InfoMessage As String = ""
        Try
            MyCommon.Open_LogixRT()
            MyCommon.Open_LogixXS()
            Dim dictionary As Dictionary(Of Integer, String) = Nothing
            Select Case CreateType
                Case "CustomerGroup"
                    Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                    InfoMessage = ValidateName(CreateType, Name)
                    If (InfoMessage.Length > 0) Then
                        ResponseText = "Error~" & InfoMessage
                    Else
                        id = DoesGroupOrProgramExists(CreateType, Name)
                        If (id = -1) Then
                            MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Name
                            MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = Val("")
                            MyCommon.LRTsp.Parameters.Add("@EditControlTypeID", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@RoleID", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.ExecuteNonQuery()
                            id = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
                            MyCommon.Activity_Log(4, id, AdminUserID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))
                            ResponseText = "Ok~" & Name & "|" & id
                        Else

                            ResponseText = "Fail~" & Name & "|" & id & "~" & Copient.PhraseLib.Lookup("term.existing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.customergroup", LanguageID).ToLower() & " : '" & Name & "'  " & Copient.PhraseLib.Lookup("offer.message", LanguageID)
                        End If
                    End If

                Case "ProductGroup"
                    Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                    InfoMessage = ValidateName(CreateType, Name)
                    If (InfoMessage.Length > 0) Then
                        ResponseText = "Error~" & InfoMessage
                    Else
                        id = DoesGroupOrProgramExists(CreateType, Name)
                        If (id = -1) Then
                            MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                            MyCommon.Open_LRTsp()
                            Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Name
                            MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@BuyerId", SqlDbType.Int).Value = DBNull.Value
                            MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                            MyCommon.LRTsp.Parameters.Add("@ProductGroupTypeID", SqlDbType.TinyInt).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            id = MyCommon.LRTsp.Parameters("@ProductGroupID").Value
                            MyCommon.Activity_Log(5, id, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID), -1)
                            ResponseText = "Ok~" & Name & "|" & id
                        Else
                            ResponseText = "Fail~" & Name & "|" & id & "~" & Copient.PhraseLib.Lookup("term.existing", LanguageID) & "  " & Copient.PhraseLib.Lookup("term.productgroup", LanguageID).ToLower() & ":  '" & Name & "'  " & Copient.PhraseLib.Lookup("offer.message", LanguageID)
                        End If
                    End If
                Case "Points"
                    Dim PromoVarID As String = String.Empty
                    Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                    InfoMessage = ValidateName(CreateType, Name)
                    If (InfoMessage.Length > 0) Then
                        ResponseText = "Error~" & InfoMessage
                    Else
                        id = DoesGroupOrProgramExists(CreateType, Name)
                        If (id = -1) Then
                            MyCommon.QueryStr = "dbo.pt_PointsPrograms_Insert"
                            MyCommon.Open_LRTsp()
                            Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                            MyCommon.LRTsp.Parameters.Add("@ProgramName", SqlDbType.NVarChar, 200).Value = Name
                            MyCommon.LRTsp.Parameters.Add("@CAMProgram", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@ExternalProgram", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@AutoDelete", SqlDbType.Bit).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            id = MyCommon.LRTsp.Parameters("@ProgramID").Value
                            MyCommon.Activity_Log(7, id, AdminUserID, Copient.PhraseLib.Lookup("history.point-create", LanguageID))

                            MyCommon.QueryStr = "dbo.pc_PointsVar_Create"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@ProgramID", SqlDbType.BigInt).Value = id
                            MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.ExecuteNonQuery()
                            PromoVarID = MyCommon.LXSsp.Parameters("@VarID").Value
                            MyCommon.Close_LXSsp()

                            MyCommon.QueryStr = "update PointsPrograms with (RowLock) SET " & _
                                                 "PromoVarID=" & PromoVarID & " " & _
                                                 "where ProgramID=" & id & ";"
                            MyCommon.LRT_Execute()

                            ResponseText = "Ok~" & Name & "|" & id
                        Else

                            ResponseText = "Fail~" & Name & "|" & id & "~" & Copient.PhraseLib.Lookup("term.existing", LanguageID) & "  " & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID).ToLower() & ":  '" & Name & "' " & Copient.PhraseLib.Lookup("offer.message", LanguageID)
                        End If
                    End If
                Case "StoredValue"
                    Dim bAllowExpirationExtension As Boolean = False
                    Dim NewExtID As String
                    Dim svTypeID As Integer = 1
                    Dim svExpireType As Integer = 1

                    bAllowExpirationExtension = IIf(MyCommon.Fetch_SystemOption(281) = "1", True, False)

                    Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                    InfoMessage = ValidateName(CreateType, Name)
                    If (InfoMessage.Length > 0) Then
                        ResponseText = "Error~" & InfoMessage
                    Else
                        id = DoesGroupOrProgramExists(CreateType, Name)
                        If (id = -1) Then
                            MyCommon.QueryStr = "dbo.pt_StoredValuePrograms_Insert"
                            MyCommon.Open_LRTsp()
                            Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Name
                            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@ExpirePeriod", SqlDbType.Int).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 200).Value = "1"
                            MyCommon.LRTsp.Parameters.Add("@OneUnitPerRec", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@SVExpireType", SqlDbType.Int).Value = svExpireType
                            MyCommon.LRTsp.Parameters.Add("@SVExpirePeriodType", SqlDbType.Int).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@ExpireTOD", SqlDbType.VarChar, 5).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = Date.Parse("12/31/2025 23:59")
                            MyCommon.LRTsp.Parameters.Add("@ExpireCentralServerTZ", SqlDbType.Bit).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@SVTypeID", SqlDbType.Int).Value = svTypeID
                            MyCommon.LRTsp.Parameters.Add("@UOMLimit", SqlDbType.Int).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@AllowReissue", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@ScorecardID", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@ScorecardDesc", SqlDbType.NVarChar, 100).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@ScorecardBold", SqlDbType.Bit).Value = 1
                            MyCommon.LRTsp.Parameters.Add("@DisallowRedeemInEarnTrans", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@AllowNegativeBal", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@AdjustmentUPC", SqlDbType.NVarChar, 100).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@RedemptionRestrictionID", SqlDbType.Int).Value = 0
                            MyCommon.LRTsp.Parameters.Add("@MemberRedemptionID", SqlDbType.Int).Value = 0

                            ' Save the first 30 characters of the name as the external ID when feature
                            ' is enabled, SV type is Points and Expire Type is fixed date/time
                            If bAllowExpirationExtension AndAlso svTypeID = 1 AndAlso svExpireType = 1 Then
                                NewExtID = Left(Name, 30)
                                MyCommon.LRTsp.Parameters.Add("@ExtProgramID", SqlDbType.NVarChar, 30).Value = NewExtID
                            End If

                            MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            id = MyCommon.LRTsp.Parameters("@SVProgramID").Value
                            MyCommon.Activity_Log(26, id, AdminUserID, Copient.PhraseLib.Lookup("history-sv-create", LanguageID))
                            ResponseText = "Ok~" & Name & "|" & id
                        Else
                            ResponseText = "Fail~" & Name & "|" & id & "~" & Copient.PhraseLib.Lookup("term.existing", LanguageID) & "  " & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID).ToLower() & ":  '" & Name & "' " & Copient.PhraseLib.Lookup("offer.message", LanguageID)
                        End If
                    End If
                Case "Location"
                    Name = MyCommon.Parse_Quotes(Logix.TrimAll(Name))
                    InfoMessage = ValidateName(CreateType, Name)
                    If (InfoMessage.Length > 0) Then
                        ResponseText = "Error~" & InfoMessage
                    Else
                        Dim OfferID As Long = MyCommon.Extract_Val(Request.QueryString("OfferID"))
                        Dim EngineType As Integer = 0
                        Dim dst As DataTable = Nothing
                        id = DoesGroupOrProgramExists(CreateType, Name)
                        If (id = -1) Then
                            Dim BannersEnabled = IIf(MyCommon.Fetch_SystemOption(66) = "1", True, False)
                            Dim BannerID As Integer = 0


                            If (BannersEnabled) Then
                                MyCommon.QueryStr = "select Top 1 BannerID from BannerOffers with (NoLock) where OfferID = " & OfferID & ";"
                                dst = MyCommon.LRT_Select

                                If (dst.Rows.Count > 0) Then
                                    BannerID = MyCommon.NZ(dst.Rows(0).Item("BannerID"), 0)
                                End If

                                If (BannerID > 0) Then
                                    MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=" & BannerID
                                    dst = MyCommon.LRT_Select
                                    If (dst.Rows.Count > 0) Then
                                        EngineType = MyCommon.NZ(dst.Rows(0).Item("EngineID"), 0)
                                    End If
                                End If
                            End If


                            MyCommon.QueryStr = "dbo.pt_LocationGroups_Insert"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Logix.TrimAll(Name)
                            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@ExtGroupId", SqlDbType.NVarChar, 20).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@ExtSeqNum", SqlDbType.NVarChar, 20).Value = ""
                            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
                            If (BannersEnabled AndAlso BannerID > 0) Then
                                MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            End If
                            MyCommon.LRTsp.ExecuteNonQuery()
                            id = MyCommon.LRTsp.Parameters("@LocationGroupId").Value
                            MyCommon.Activity_Log(11, id, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-create", LanguageID))
                        End If

                        If (id > 0) Then
                            'Associating location group to Offer
                            MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where LocationGroupID=1 and OfferID=" & OfferID
                            dst = MyCommon.LRT_Select

                            If (dst.Rows.Count > 0) Then
                                ' All cardholders is already in the selected box, so lose it
                                MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID=1 and OfferID=" & OfferID
                                MyCommon.LRT_Execute()
                                MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
                                MyCommon.LRT_Execute()
                            End If

                            If (id = 1) Then
                                MyCommon.QueryStr = "delete from OfferLocations with (RowLock) where Deleted=1 and OfferID=" & OfferID
                                MyCommon.LRT_Execute()
                                MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, TCRMAStatusFlag=3 where OfferID=" & OfferID
                                MyCommon.LRT_Execute()

                                MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                                MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = id
                                MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
                                MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
                                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                MyCommon.LRTsp.ExecuteNonQuery()
                                MyCommon.Close_LRTsp()

                                MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & id
                                MyCommon.LRT_Execute()

                                MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
                                MyCommon.LRT_Execute()
                                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
                            Else
                                MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                                MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = id
                                MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
                                MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
                                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                MyCommon.LRTsp.ExecuteNonQuery()
                                MyCommon.Close_LRTsp()

                                MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & id
                                MyCommon.LRT_Execute()

                                MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
                                MyCommon.LRT_Execute()
                                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
                            End If
                        End If
                        ResponseText = "Ok~" & Name & "|" & id
                    End If
            End Select
        Catch ex As Exception
            ResponseText = "Error~" & ex.Message
        Finally
            MyCommon.Close_LogixRT()
            MyCommon.Close_LogixXS()
            MyCommon = Nothing
        End Try

        Send(ResponseText)
    End Sub

    Sub GetCardTypeForCustomerCondition(conditionID As Int32)
        Dim dictionary As Dictionary(Of Integer, String) = Nothing
        Dim dtAllCardTypes As DataTable = Nothing
        Dim dtSavedCardTypesForCustCondition As DataTable = Nothing
        Dim distinctCounts As IEnumerable(Of Int32) = Nothing
        Dim responseText As String =""

        Try
            Dim resolverbuilder As New WebRequestResolverBuilder()
            CurrentRequest.Resolver = resolverbuilder.GetResolver()
            resolverbuilder.Build()
            CurrentRequest.Resolver.AppName = "OfferFeeds.aspx"
            Dim customerCondService As ICustomerGroupCondition = CurrentRequest.Resolver.Resolve(Of ICustomerGroupCondition)(CurrentRequest.Resolver.AppName)
            dtAllCardTypes = customerCondService.GetCustomerConditionCardTypes()
            If (dtAllCardTypes.Rows.Count > 0) Then
                dictionary=New Dictionary(Of Integer, String)
                For Each row2 In dtAllCardTypes.Rows
                    dictionary.Add(MyCommon.NZ(row2.Item("CardTypeID"), 0), Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), ""), LanguageID))
                Next
            End If
        Catch ex As Exception
            responseText = "fail~" & ex.Message
        End Try
        If (dictionary.Count > 0) Then
            responseText = "success~" & JsonConvert.SerializeObject(dictionary)
        Else
            responseText = "success~" & "-1"
        End If
        If responseText.Length > 0 Then
            responseText = responseText.Replace("&#39;", "'")
        End If
        Send(responseText)
    End Sub

    Sub SendBanners(ByVal OfferID As Long, ByVal AdminUserID As Integer)
        Dim dt As DataTable
        Dim row As DataRow
        Dim SelectedList As String = ""
        Dim i As Integer
        Dim SelectedBanners, EditableBanners As ArrayList
        Dim IsEditableBanner As Boolean = False
        Dim AllowMultipleBanners As Boolean = False
        Dim BannersEnabled As Boolean = False
        Dim EditableBuf As New StringBuilder("")
        Dim NonEditableBuf As New StringBuilder("")
        Dim TempBuf As New StringBuilder()

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try

            BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
            AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")

            If (BannersEnabled AndAlso AllowMultipleBanners) Then
                Send("<div id=""titlebar""><a href=""javascript:closeResults();""><img src=""../images/close.png"" border=""0"" /></a></div>")
                Send("<div style=""padding-left:300px;"">")
                Send("  <input id=""btn_OF_Save"" name=""btn_OF_Save"" type=""button"" class=""regular"" value=""Save""  onclick=""saveBanners();"" />")
                Send("</div>")
                Send("<div id=""banners"" style=""padding-left:10px;line-height:20px;height:400px;overflow:auto;"">")

                ' get the selected banners and store for later lookup
                MyCommon.QueryStr = "select BAN.BannerID from BannerOffers BO with (NoLock) " & _
                                    "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                                    "where BAN.Deleted=0 and BO.OfferID = " & OfferID
                dt = MyCommon.LRT_Select
                SelectedBanners = New ArrayList(dt.Rows.Count)
                For Each row In dt.Rows
                    SelectedBanners.Add(MyCommon.NZ(row.Item("BannerID"), -1))
                    If (SelectedList <> "") Then SelectedList &= ","
                    SelectedList &= MyCommon.NZ(row.Item("BannerID"), -1)
                Next

                Send("<input type=""hidden"" name=""bannerschanged"" id=""bannerschanged"" value=""false"" />")

                ' get the banners for which this user is permitted to edit
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock)" & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and AUB.AdminUserID =" & AdminUserID & ";"
                dt = MyCommon.LRT_Select
                EditableBanners = New ArrayList(dt.Rows.Count)
                For Each row In dt.Rows
                    EditableBanners.Add(MyCommon.NZ(row.Item("BannerID"), -1))
                Next

                ' get all the assigned banners for CPE
                i = 0
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and BAN.AllBanners=0;"
                dt = MyCommon.LRT_Select()
                For Each row In dt.Rows
                    IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row.Item("BannerID"), -1))
                    TempBuf = New StringBuilder()
                    TempBuf.Append(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""bannerid" & i & """ value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """")
                    TempBuf.Append(IIf(SelectedBanners.Contains(MyCommon.NZ(row.Item("BannerID"), -1)), " checked=""checked""", " "))
                    TempBuf.Append(IIf(IsEditableBanner, " ", " disabled = ""disabled"""))
                    TempBuf.Append(" onClick=""handleBanners(this);""")
                    TempBuf.Append(" />")
                    TempBuf.Append(("<label for=""bannerid" & i & """ title=""" & Copient.PhraseLib.Lookup(IIf(IsEditableBanner, "banners.add-to-offer-note", "banners.not-user-note"), LanguageID) & """"))
                    TempBuf.Append(">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</label><br />" & ControlChars.CrLf)

                    If (IsEditableBanner) Then
                        EditableBuf.Append(TempBuf)
                    Else
                        NonEditableBuf.Append(TempBuf)
                    End If

                    i += 1
                Next

                If (EditableBuf.Length > 0) Then
                    Send("<br class=""half""/><b><u>" & Copient.PhraseLib.Lookup("offer-feeds.AssignedBanners", LanguageID) & "</u></b><br />")
                    Send(EditableBuf.ToString)
                End If

                If (NonEditableBuf.Length > 0) Then
                    Send("<br /><b><u>" & Copient.PhraseLib.Lookup("offer-feeds.UnassignableBanners", LanguageID) & "</u></b><br />")
                    Send(NonEditableBuf.ToString)
                End If

                ' get all the assigned ALL banners for CPE
                i = 0
                MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                    "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                    "where BE.EngineID=2 and BAN.AllBanners=1;"
                dt = MyCommon.LRT_Select()
                If (dt.Rows.Count > 0) Then
                    Send("<br />")
                    Send("<b><u>" & Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & ":</u></b><br />")
                    For Each row In dt.Rows
                        IsEditableBanner = EditableBanners.Contains(MyCommon.NZ(row.Item("BannerID"), -1))
                        Sendb(Space(10) & "<input type=""checkbox"" name=""bannerids"" id=""allbannerid" & i & """ value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """")
                        Sendb(IIf(SelectedBanners.Contains(MyCommon.NZ(row.Item("BannerID"), -1)), " checked=""checked""", " "))
                        Sendb(IIf(IsEditableBanner, " ", " disabled = ""disabled"""))
                        Sendb(" onClick=""handleAllBanners(this);""")
                        Sendb(" />")
                        Sendb("<label for=""allbannerid" & i & """ title=""" & Copient.PhraseLib.Lookup(IIf(IsEditableBanner, "banners.add-to-offer-note", "banners.not-user-note"), LanguageID) & """")
                        Send(">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</label><br />")

                        i += 1
                    Next
                End If
                Send("<br />")
                Send("</div>")
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub SaveBanners(ByVal OfferID As Long, ByVal AdminUserID As Integer)
        Dim i As Integer
        Dim AllowMultipleBanners As Boolean = False
        Dim BannersEnabled As Boolean = False

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            MyCommon.Open_LogixRT()

            BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
            AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")

            If (BannersEnabled AndAlso AllowMultipleBanners AndAlso Request.QueryString("bannerschanged") = "true") Then
                ' first clear out the existing banners
                MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID =" & OfferID & ";"
                MyCommon.LRT_Execute()

                ' add the selected banners
                If (Request.QueryString("bannerids") <> "") Then
                    For i = 0 To Request.QueryString.GetValues("bannerids").GetUpperBound(0)
                        MyCommon.QueryStr = "insert into BannerOffers with (RowLock) (BannerID, OfferID) values (" & MyCommon.Extract_Val(Request.QueryString.GetValues("bannerids")(i)) & "," & OfferID & ");"
                        MyCommon.LRT_Execute()
                    Next i
                End If

            End If

            Send("OK")
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub SendPopup(ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal PageName As String, ByVal Tokens As String)
        Dim uriStr As String = ""
        Dim pos As Integer

        Try
            ' build the URL path to the condition page
            uriStr = Request.Url.AbsoluteUri
            pos = uriStr.LastIndexOf("/OfferFeeds.aspx")

            If (pos > -1) Then
                uriStr = uriStr.Remove(pos + 1)
                uriStr &= PageName & "?OfferID=" & OfferID
                If (Tokens <> "") Then
                    uriStr &= "&" & Tokens.Replace(":", "=").Replace(";", "&")
                End If
            End If
            Send("<div id=""titlebar""><a href=""javascript:closeResults();""><img src=""../images/close.png"" border=""0"" /></a></div>")
            Send("<iframe id=""iframePopup"" src=""" & uriStr & """ style=""height:525px;width:700px;border:0px;"" frameborder=""0"" scrolling=""no"" >")
            Send("</iframe>")
        Catch ex As Exception
            Send(ex.ToString)
        End Try

    End Sub

    Sub FavoriteOfferForAll(ByVal OfferID As Long, ByVal AdminUserID As Long)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Try
            MyCommon.Open_LogixRT()

            ' remove all existing users that have favorited the offer
            MyCommon.QueryStr = "delete AdminUserOffers with (RowLock) where OfferID =" & OfferID
            MyCommon.LRT_Execute()

            ' now favorite the offer for all users
            MyCommon.QueryStr = "insert into AdminUserOffers with (RowLock) (AdminUserID, OfferID, Priority, FavoredBy, FavoredDate) " & _
                                "  select AdminUserID, " & OfferID & ", 1, " & AdminUserID & ", getdate() from AdminUsers AU with (NoLock);"
            MyCommon.LRT_Execute()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-favoritesall", LanguageID))
            Sendb("OK")
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try

    End Sub

    Sub FindProductGroupsCM(ByVal Search As String, ByVal OfferID As Integer, ByVal SelectedGroup As Integer, ByVal ExcludedGroup As Integer, ByVal SearchRadio As String, Optional ByVal CallingPage As String = "")
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim sValidLocIDs As String = ""
        Dim sValidSU As String = ""
        Dim wherestr As String = ""
        Dim sJoin As String = ""
        Dim orderby As String

        If (MyCommon.Fetch_SystemOption(235) = "1") Then
            orderby = " order by AnyProduct desc, Name"
        Else
            orderby = " order by AnyProduct desc, ProductGroupID desc, Name asc"
        End If

        Try
            MyCommon.Open_LogixRT()

            Dim sendString As String = ""

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO) Then
                    nameLike = " and LTRIM(Name) like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = " and LTRIM(Name) like '" & Search & "%' "
                End If
            End If

            Dim RECORD_LIMIT As Integer = GroupRecordLimit
            Dim topGroups As String = ""
            If (RECORD_LIMIT > 0) Then topGroups = " top " & RECORD_LIMIT & " "

            If StoreUser(sValidLocIDs, sValidSU) Then
                sJoin &= "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID "
                wherestr = " and (LocationID in (" & sValidLocIDs & ") or CreatedByAdminID in (" & sValidSU & ")) "
            End If

            'Find products groups using search
            MyCommon.QueryStr = "Select " & topGroups & "  pg.ProductGroupID,Name,PhraseID from ProductGroups pg with (NoLock) " & sJoin & " where pg.ProductGroupID is not null and deleted=0 " & wherestr
            If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 "
            If nameLike <> "" Then MyCommon.QueryStr &= nameLike
            'Selected Product Group
            If (CallingPage.Equals("StoredValue")) Then
                'Selected Product Group
                MyCommon.QueryStr &= "AND pg.ProductGroupID NOT IN (SELECT ProductGroupID from OfferRewards WHERE  Deleted=0 AND OfferID =" & OfferID & ") "
                'Excluded product group
                MyCommon.QueryStr &= "AND pg.ProductGroupID NOT IN (SELECT ExcludedProdGroupID from OfferRewards WHERE  Deleted=0 AND  OfferID = " & OfferID & ")"
            ElseIf (CallingPage.Equals("conProduct")) Then
                'Selected Product Group
                MyCommon.QueryStr &= "AND pg.ProductGroupID NOT IN (SELECT LinkID from OfferConditions WHERE  Deleted=0 AND ConditionTypeID=2 AND OfferID =" & OfferID & ") "
                'Excluded product group
                MyCommon.QueryStr &= "AND pg.ProductGroupID NOT IN (SELECT ExcludedID from OfferConditions WHERE  Deleted=0 AND ConditionTypeID=2 AND OfferID = " & OfferID & ")"
            End If

            'Do not count current selected group.
            MyCommon.QueryStr &= "AND pg.ProductGroupID NOT IN(" & SelectedGroup & "," & ExcludedGroup & ")"

            MyCommon.QueryStr &= orderby
            'If results meet number limit, send back the results
            Dim rst As DataTable = MyCommon.LRT_Select()
            'Build select

            Dim productGroupID As Integer = 0
            For Each row As DataRow In rst.Rows
                productGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                If (productGroupID <> 0 AndAlso productGroupID <> SelectedGroup AndAlso productGroupID <> ExcludedGroup) Then
                    If productGroupID = 1 Then
                        sendString &= "<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>"
                    Else
                        sendString &= "<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>"
                    End If
                End If
            Next
            Send(sendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Function StoreUser(ByRef sValidLocIds As String, ByRef sValidSU As String) As Boolean
        Dim rst As DataTable
        Dim bStoreUser As Boolean = False
        Dim iLen As Integer = 0
        Dim i As Integer = 0

        'Store User
        If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
            MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
            rst = MyCommon.LRT_Select
            iLen = rst.Rows.Count
            If iLen > 0 Then
                bStoreUser = True
                sValidSU = AdminUserID
                For i = 0 To (iLen - 1)
                    If i = 0 Then
                        sValidLocIds = rst.Rows(0).Item("LocationID")
                    Else
                        sValidLocIds &= "," & rst.Rows(i).Item("LocationID")
                    End If
                Next

                MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIds & ") and NOT UserID=" & AdminUserID & ";"
                rst = MyCommon.LRT_Select
                iLen = rst.Rows.Count
                If iLen > 0 Then
                    For i = 0 To (iLen - 1)
                        sValidSU &= "," & rst.Rows(i).Item("UserID")
                    Next
                End If
            End If
        End If
        Return bStoreUser
    End Function

    Sub FindStoredValueProgramsCM(ByVal Search As String, ByVal SelectedProgram As Integer, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "Pfunctionradio1"
        Const CONTAINING_RADIO As String = "Pfunctionradio2"
        Try
            MyCommon.Open_LogixRT()

            Dim sendString As String = ""

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO Or SearchRadio = "functionradio2") Then
                    nameLike = " and LTRIM(Name) like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = " and LTRIM(Name) like '" & Search & "%' "
                End If
            End If

            Dim RECORD_LIMIT As Integer = GroupRecordLimit
            Dim topGroups As String = ""
            If (RECORD_LIMIT > 0) Then topGroups = " top " & RECORD_LIMIT & " "

            'Find stored value programs using search
            MyCommon.QueryStr = "Select " & topGroups & "  SVProgramID, Name, SVTypeID from StoredValuePrograms with (NoLock) where SVProgramID is not null and deleted=0 "

            If nameLike <> "" Then MyCommon.QueryStr &= nameLike

            'Do not count current selected group.
            MyCommon.QueryStr &= "AND SVProgramID NOT IN(" & SelectedProgram & ")"

            MyCommon.QueryStr &= " order by SVProgramID desc, Name asc"
            'If results meet number limit, send back the results
            Dim rst As DataTable = MyCommon.LRT_Select()

            Dim svProgramID As Integer = 0
            For Each row As DataRow In rst.Rows
                svProgramID = MyCommon.NZ(row.Item("SVProgramID"), 0)
                If (svProgramID <> 0 AndAlso svProgramID <> SelectedProgram) Then
                    sendString &= "<option value=""" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>"
                End If
            Next
            Send(sendString)

        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub
    Sub FindPointsProgramsCM(ByVal Search As String, ByVal SelectedProgram As Integer, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "Pfunctionradio1"
        Const CONTAINING_RADIO As String = "Pfunctionradio2"
        Try
            MyCommon.Open_LogixRT()

            Dim sendString As String = ""

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO Or SearchRadio = "functionradio2") Then
                    nameLike = " and LTRIM(ProgramName) like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = " and LTRIM(ProgramName) like '" & Search & "%' "
                End If
            End If

            Dim sTenderPointsId As String = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(37), "")
            Dim RECORD_LIMIT As Integer = GroupRecordLimit
            Dim topGroups As String = ""
            If (RECORD_LIMIT > 0) Then topGroups = " top " & RECORD_LIMIT & " "

            'Find stored value programs using search


            MyCommon.QueryStr = "Select " & topGroups & "  pp.ProgramID, pp.ProgramName from PointsPrograms pp with (NoLock) where "
            MyCommon.QueryStr &= "  pp.ProgramID is not null and pp.Deleted = 0  "
            If nameLike <> "" Then MyCommon.QueryStr &= nameLike
            'Do not count current selected group.
            If Not String.IsNullOrEmpty(sTenderPointsId) Then
                If (Not String.IsNullOrEmpty(SelectedProgram)) Then
                    MyCommon.QueryStr &= " and NOT EXISTS (SELECT e.ProgramID FROM PointsPrograms e WHERE e.ProgramID=pp.ProgramID and  e.ProgramID IN( " & SelectedProgram & "," & sTenderPointsId & ")) "
                Else
                    MyCommon.QueryStr &= " and NOT EXISTS (SELECT e.ProgramID FROM PointsPrograms e WHERE e.ProgramID=pp.ProgramID and  e.ProgramID IN( " & sTenderPointsId & ")) "
                End If

            Else
                If (Not String.IsNullOrEmpty(SelectedProgram)) Then
                    MyCommon.QueryStr &= " and NOT EXISTS (SELECT e.ProgramID FROM PointsPrograms e WHERE e.ProgramID=pp.ProgramID and  e.ProgramID IN( " & SelectedProgram & ")) "
                End If
            End If
            MyCommon.QueryStr &= " order by ProgramID desc, ProgramName asc"
            'If results meet number limit, send back the results
            Dim rst As DataTable = MyCommon.LRT_Select()

            Dim pointsProgramID As Integer = 0
            For Each row As DataRow In rst.Rows
                pointsProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
                If (pointsProgramID <> 0 AndAlso pointsProgramID <> SelectedProgram) Then
                    sendString &= "<option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>"
                End If
            Next
            Send(sendString)

        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub FindProductGroups(ByVal Search As String, ByVal ROID As Integer, ByVal Disqualifier As Boolean, ByVal SelectedGroup As Integer, ByVal ExcludedGroup As Integer, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Try
            MyCommon.Open_LogixRT()

            Dim sendString As String = ""

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO) Then
                    nameLike = " and Name like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = " and Name like '" & Search & "%' "
                End If
            End If

            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            Dim topGroups As String = ""
            If (RECORD_LIMIT > 0) Then topGroups = " top " & RECORD_LIMIT & " "

            'Find products groups using search
            MyCommon.QueryStr = "select " & topGroups & " ProductGroupID, Name from ProductGroups where ProductGroupID is not null " & _
                                         "and Deleted=0 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
            If nameLike <> "" Then MyCommon.QueryStr &= nameLike
            If Disqualifier Then
                MyCommon.QueryStr &= " and ProductGroupID <> 1  and ProductGroupID not in " & _
                                     "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & ROID & " and Disqualifier=0 and ExcludedProducts=0)"
            Else
                MyCommon.QueryStr &= " and ProductGroupID not in " & _
                                     "   (select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & ROID & " and Disqualifier=1 and ExcludedProducts=0)"
            End If

            If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(TranslatedFromOfferID,0) = 0 "

            MyCommon.QueryStr &= " order by AnyProduct desc, ProductGroupID desc, Name asc"
            'If results meet number limit, send back the results

            Dim rst As DataTable = MyCommon.LRT_Select()
            'Build select

            Dim productGroupID As Integer = 0
            For Each row As DataRow In rst.Rows
                productGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                If (productGroupID <> 0 AndAlso productGroupID <> SelectedGroup AndAlso productGroupID <> ExcludedGroup) Then
                    If productGroupID = 1 Then
                        sendString &= "<option value=""1"" style=""color: brown; font-weight: bold;"">" & Copient.PhraseLib.Lookup("term.anyproduct", LanguageID) & "</option>"
                    Else
                        sendString &= "<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & MyCommon.NZ(row.Item("ProductGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>"
                    End If
                End If
            Next
            Send(sendString)

        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub DiscountProductGroups(ByVal Search As String, ByVal SelectedGroup As Integer, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Try
            MyCommon.Open_LogixRT()

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO) Then
                    nameLike = " and Name like '%" & Search & "%' "
                Else
                    nameLike = " and Name like '" & Search & "%' "
                End If
            End If

            Dim topString As String = ""
            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
            MyCommon.QueryStr = "select " & topString & " ProductGroupID, Name from ProductGroups with (NoLock) " & _
                                "where Deleted=0 and ProductGroupID<>1 and NonDiscountableGroup=0 and PointsNotApplyGroup=0 "
            MyCommon.QueryStr &= nameLike & " order by ProductGroupID desc, Name;"
            Dim rst As DataTable = MyCommon.LRT_Select
            Dim sendString As String = ""
            Dim productGroupID As Integer = 0
            For Each row As DataRow In rst.Rows
                productGroupID = MyCommon.NZ(row.Item("ProductGroupID"), 0)
                If (productGroupID > 0 AndAlso Not productGroupID = SelectedGroup) Then sendString &= "<option title=""" & MyCommon.NZ(row.Item("Name"), "") & """ value=""" & productGroupID & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>"
            Next
            Send(sendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub ConditionCustomerGroups(ByVal Search As String, ByVal EngineID As Integer, ByVal AnyCustomerEnabled As Boolean, ByVal SelectedGroups As String, ByVal ExcludedGroups As String, ByVal OfferID As Long, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Dim IsAnyCustomerEnabled As Boolean = False
        Dim IsAnyCardholderEnabled As Boolean = False
        Dim ALLCAM As Integer = -1
        Dim NewCardholdersID As Integer = -1
        Dim SearchStartIndex As Boolean = False
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Try
            MyCommon.Open_LogixRT()

            Dim selectedList As IList = SelectedGroups.Split(",")
            Dim excludedList As IList = ExcludedGroups.Split(",")
            Dim SendString As String = ""

            If (EngineID <> 6) Then
                'add "special" customer groups
                If AnyCustomerEnabled Then
                    'see if the offer conditions/rewards allow us to display AnyCustomer group.  (This is not allowed if conditions/rewards require a known customer ex: Points, Stored Value, etc.)
                    MyCommon.QueryStr = "dbo.pa_Check_AnyCustomer_Violation"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                    Dim dst As DataTable = MyCommon.LRTsp_select
                    MyCommon.Close_LRTsp()
                    If dst.Rows.Count = 0 AndAlso Not selectedList.Contains("1") Then
                        IsAnyCustomerEnabled = True
                    End If
                    dst = Nothing
                End If
                MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where NewCardholders=1 and Deleted=0;"
                Dim newCardDT As DataTable = MyCommon.LRT_Select
                If newCardDT.Rows.Count > 0 Then
                    NewCardholdersID = MyCommon.NZ(newCardDT.Rows(0).Item("CustomerGroupID"), -1)
                End If
                If (Not selectedList.Contains("2")) Then
                    IsAnyCardholderEnabled = True
                End If
            Else
                MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0;"
                Dim camDT As DataTable = MyCommon.LRT_Select
                If camDT.Rows.Count > 0 Then
                    ALLCAM = MyCommon.NZ(camDT.Rows(0).Item("CustomerGroupID"), -1)
                End If
            End If

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO) Then
                    nameLike = "  CG.Name like '%" & Search & "%' "
                Else
                    nameLike = "  CG.Name like '" & Search & "%' "
                    SearchStartIndex = True
                End If
            End If

            Dim topString As String = ""
            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
            If (EngineID <> 6) Then
                If (EngineID = 0) Then
                    'MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where deleted=0 and AnyCustomer<>1 and CustomerGroupID <> 2 and NewCardholders=0 and CustomerGroupID is not null and BannerID is null "
                    MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                                         "FROM CustomerGroups CG With (NOLOCK) " &
                                         "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                                         "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                                         "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                                         "And NewCardholders = 0 AND CG.Deleted = 0 " &
                                         "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                                         "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                                         " And CG.isOptInGroup = 0 "
                Else
                    ' MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup<>1 "
                    'MyCommon.QueryStr = MyCommon.QueryStr & " and isOptInGroup=0 "
                    MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                                        "FROM CustomerGroups CG With (NOLOCK) " &
                                        "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                                        "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                                        "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                                        "And NewCardholders = 0 And CAMCustomerGroup <> 1 AND CG.Deleted = 0 " &
                                        "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                                        "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                                        " And CG.isOptInGroup = 0 "

                End If
            Else
                'MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup=1 and AnyCAMCardholder<>1 "
                MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                                        "FROM CustomerGroups CG With (NOLOCK) " &
                                        "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                                        "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                                        "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                                        "And NewCardholders = 0  and CAMCustomerGroup=1 and AnyCAMCardholder<>1 AND CG.Deleted = 0 " &
                                        "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                                        "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                                        " And CG.isOptInGroup = 0 "
            End If

            'To disable option group for every engine
            'MyCommon.QueryStr = MyCommon.QueryStr & " and isOptInGroup=0 "
            If (nameLike <> String.Empty) Then
                MyCommon.QueryStr &= "and " & nameLike
            End If
            MyCommon.QueryStr &= MyCommon.QueryStr & " order by CG.CustomerGroupID desc, CG.Name;"

            Dim groupDT As DataTable = AddSystemGroups(MyCommon.LRT_Select(), IsAnyCustomerEnabled, IsAnyCardholderEnabled, NewCardholdersID, ALLCAM, Search, SearchStartIndex)
            Dim custGroupId As Int32 = 0
            For Each row As DataRow In groupDT.Rows
                If (Not selectedList.Contains(MyCommon.NZ(row.Item("CustomerGroupID"), 0).ToString()) AndAlso Not excludedList.Contains(MyCommon.NZ(row.Item("CustomerGroupID"), 0).ToString())) Then
                    custGroupId = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                    If (custGroupId = 1 Or custGroupId = 2 Or custGroupId = 3 Or custGroupId = 4) Then
                        SendString &= "<option value=""" & custGroupId & """ style=""color:brown;font-weight:bold;"">" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "</option>"
                    Else
                        SendString &= "<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "</option>"
                    End If
                End If
            Next

            Send(SendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try

    End Sub
    Sub CGReportCondition(ByVal Search As String, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Dim IsAnyCustomerEnabled As Boolean = False
        Dim IsAnyCardholderEnabled As Boolean = False
        Dim ALLCAM As Integer = -1
        Dim NewCardholdersID As Integer = -1
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Try
            MyCommon.Open_LogixRT()
            Dim SendString As String = ""
        
            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO Or SearchRadio = "functionradio2") Then
                    nameLike = " and LTRIM(Name) like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = " and LTRIM(Name) like '" & Search & "%' "
                End If
            End If
            
            Dim topString As String = ""
            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
            
            MyCommon.QueryStr = "Select " & topString & "CustomerGroupID, ExtGroupID, Name from CustomerGroups with (NoLock) " & _
           " where Deleted = 0 And CustomerGroupID <> 1 And CustomerGroupID <> 2" & _
           " And BannerID Is null And NewCardholders = 0 And AnyCAMCardholder = 0"
           
            If nameLike <> "" Then MyCommon.QueryStr &= nameLike
            
            Dim groupDT As DataTable = MyCommon.LRT_Select()
            Dim custGroupId As Int32 = 0
            For Each row As DataRow In groupDT.Rows
                
                custGroupId = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                If (custGroupId <> 0) Then
                    SendString &= "<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "</option>"
                End If
            Next
            Send(SendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    
    End Sub
    
    Sub OffReportCondition(ByVal Search As String, ByVal SearchRadio As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Dim IsAnyCustomerEnabled As Boolean = False
        Dim IsAnyCardholderEnabled As Boolean = False
        Dim ALLCAM As Integer = -1
        Dim NewCardholdersID As Integer = -1
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradioOff1"
        Const CONTAINING_RADIO As String = "functionradioOff2"
        Try
            MyCommon.Open_LogixRT()
            Dim SendString As String = ""
        
            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO Or SearchRadio = "functionradioOff2") Then
                    nameLike = "WHERE LTRIM(Name) like '%" & Search & "%' "
                Else 'Default the search to be starting with
                    nameLike = "WHERE LTRIM(Name) like '" & Search & "%' "
                End If
            End If
            
            Dim topString As String = ""
            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
            
            MyCommon.QueryStr = "select OfferID, Name from" & _
                                "(select OfferID, Name from Offers with (NoLock) where Deleted=0 and IsTemplate=0 " & _
                             " union " & _
                             "select IncentiveID as OfferID, IncentiveName as Name from CPE_Incentives with (NoLock) where Deleted=0 and IsTemplate=0 ) T "
                           
            If nameLike <> "" Then MyCommon.QueryStr &= nameLike
            
            Dim OffersDT As DataTable = MyCommon.LRT_Select()
            Dim OfferId As Int32 = 0
            For Each row As DataRow In OffersDT.Rows
                
                OfferId = MyCommon.NZ(row.Item("OfferId"), 0)
                If (OfferId <> 0) Then
                    SendString &= "<option value=""" & MyCommon.NZ(row.Item("OfferId"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "").ToString().Replace("'", "\'") & "</option>"
                End If
            Next
            Send(SendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    
    End Sub
    Function AddSystemGroups(ByVal dt As DataTable, ByVal IsAnyCustomerEnabled As Boolean, ByVal IsAnyCardholderEnabled As Boolean, ByVal NewCardholdersID As Integer,
                              ByVal ALLCAM As Integer, ByVal search As String, ByVal SearchStartIndex As Boolean) As DataTable
        Dim dt1 = dt.Clone()
        Dim dr As DataRow
        If IsAnyCustomerEnabled = True Then
            dr = dt1.NewRow()
            dr("CustomerGroupID") = 1
            dr("Name") = StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase)
            dt1.Rows.Add(dr)
        End If
        If IsAnyCardholderEnabled = True Then
            dr = dt1.NewRow()
            dr("CustomerGroupID") = 2
            dr("Name") = StrConv(Copient.PhraseLib.Lookup("term.anycardholder", LanguageID), VbStrConv.ProperCase)
            dt1.Rows.Add(dr)
        End If
        If NewCardholdersID > 0 Then
            dr = dt1.NewRow()
            dr("CustomerGroupID") = NewCardholdersID
            dr("Name") = StrConv(Copient.PhraseLib.Lookup("term.newcardholders", LanguageID), VbStrConv.ProperCase)
            dt1.Rows.Add(dr)
        End If
        If ALLCAM > 0 Then
            dr = dt1.NewRow()
            dr("CustomerGroupID") = ALLCAM
            dr("Name") = StrConv(Copient.PhraseLib.Lookup("term.allcam", LanguageID), VbStrConv.ProperCase)
            dt1.Rows.Add(dr)
        End If
        If (search <> String.Empty) Then
            dt1 = FilterSystemGroups(dt1, search, SearchStartIndex)
            dt1.Merge(dt)
            Return dt1
        End If
        dt1.Merge(dt)
        Return dt1
    End Function

    Private Function FilterSystemGroups(ByVal tbl As DataTable, ByVal search As String, ByVal SearchStartIndex As Boolean) As DataTable
        Dim temp = tbl.AsEnumerable().Where(Function(Row, retVal)
                                                retVal = False
                                                If SearchStartIndex AndAlso Row(1).ToString.ToLower.StartsWith(search) Then
                                                    retVal = True
                                                ElseIf Row(1).ToString.ToLower.IndexOf(search) <> -1 AndAlso Not SearchStartIndex Then
                                                    retVal = True
                                                End If
                                                Return retVal
                                            End Function)
        If (temp.Count > 0) Then
            tbl = temp.CopyToDataTable()
        Else
            tbl.Rows.Clear()
        End If
        Return tbl
    End Function

    Sub GrantMembershipGroups(ByVal OfferID As String, ByVal RewardID As String, ByVal Search As String, ByVal EngineID As Integer, ByVal SearchRadio As String, Optional ByVal SelectedGroup As String = "")
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Name of radio buttons from product search
        Const START_WITH_RADIO As String = "functionradio1"
        Const CONTAINING_RADIO As String = "functionradio2"
        Dim dt As DataTable
        Dim row As DataRow
        Dim RewardTypeID As Integer
        Dim NotIn As String = String.Empty
        Try
            MyCommon.Open_LogixRT()

            Dim nameLike As String = ""
            If (Search <> "") Then
                If (SearchRadio = CONTAINING_RADIO) Then
                    nameLike = " and Name like '%" & Search & "%' "
                Else
                    nameLike = " and Name like '" & Search & "%' "
                End If
            End If
            'Exclude already selected group
            If (SelectedGroup <> "") Then
                NotIn = "NOT IN(" & SelectedGroup & ")"
            End If
            Dim RECORD_LIMIT As Integer = GroupRecordLimit 'MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126).Trim()) '500
            Dim topString As String = ""
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) AndAlso RewardID <> String.Empty) Then
                MyCommon.QueryStr = "Select RewardTypeID from OfferRewards as OFR with (NoLock) left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID where RewardID=" & RewardID
                dt = MyCommon.LRT_Select
                For Each row In dt.Rows
                    RewardTypeID = row.Item("RewardTypeID")
                Next
                If RewardTypeID = 5 Then
                    MyCommon.QueryStr = "Select distinct" & topString & " CustomerGroupID,Name from CustomerGroups as CG with (NoLock) Left Outer Join OfferConditions OC ON CG.CustomerGroupID <> OC.LinkID " &
                            "Left join ExtSegmentMap exs on exs.InternalId = cg.CustomerGroupID " &
                            "where CG.AnyCardholder <> 1 and CG.AnyCustomer <> 1 and CG.NewCardholders <> 1 and CG.Deleted = 0 " &
                            "and (exs.ExtSegmentID is null or exs.ExtSegmentID > 0) and (exs.SegmentTypeID is null or exs.SegmentTypeID = 1) " &
                            "and OC.ConditionTypeID = 1 and OC.OfferID = " & OfferID
                Else
                    MyCommon.QueryStr = "select distinct" & topString & " CG.CustomerGroupID, CG.Name from CustomerGroups as CG with (NoLock) " & _
                                 "where Deleted=0 and AnyCustomer<>1 and NewCardholders<>1 and AnyCardholder <> 1"
                End If
            ElseIf (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                MyCommon.QueryStr = "select distinct" & topString & " CG.CustomerGroupID, CG.Name From CustomerGroups as CG " &
                    "Left Outer Join ( Select CGI.CustomerGroupID From CPE_IncentiveCustomerGroups as CGI " &
                    "Inner Join CPE_RewardOptions as RO on CGI.RewardOptionID = RO.RewardOptionID " &
                    "Where RO.IncentiveID = " & OfferID & " and CGI.ExcludedUsers = 0 and CGI.Deleted = 0) as EX on EX.CustomerGroupID = CG.CustomerGroupID " &
                    "Left Join ExtSegmentMap as exs on exs.InternalId = CG.CustomerGroupID " &
                    "Where EX.CustomerGroupID is null and CG.AnyCardholder <> 1 and CG.AnyCustomer <> 1 and CG.NewCardholders <> 1 " &
                    "and CG.Deleted = 0 and CG.CustomerGroupID <> 1 " &
                    "and (exs.SegmentTypeId = 1 or exs.SegmentTypeId is null) " &
                    "and (exs.ExtSegmentID > 0 or exs.ExtSegmentID is null) and (exs.IncentiveID <> 0 or exs.IncentiveID is null) "
            ElseIf (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
                MyCommon.QueryStr = "select distinct" & topString & " CG.CustomerGroupID, CG.Name From CustomerGroups as CG " &
                    "Left Outer Join ( Select CGI.CustomerGroupID From CPE_IncentiveCustomerGroups as CGI " &
                    "Inner Join CPE_RewardOptions as RO on CGI.RewardOptionID = RO.RewardOptionID " &
                    "Where RO.IncentiveID = " & OfferID & " and CGI.ExcludedUsers = 0 and CGI.Deleted = 0) as EX on EX.CustomerGroupID = CG.CustomerGroupID " &
                    "Left join ExtSegmentMap exs on exs.InternalId = cg.CustomerGroupID " &
                    "Where EX.CustomerGroupID is null and CG.AnyCardholder <> 1 and CG.AnyCustomer <> 1 and CG.NewCardholders <> 1 " &
                    "and (exs.ExtSegmentID is null or exs.ExtSegmentID > 0) and (exs.SegmentTypeID is null or exs.SegmentTypeID = 1) " &
                    "and CG.Deleted = 0 and CG.CustomerGroupID <> 1 "
            Else
                MyCommon.QueryStr = "select distinct" & topString & " CG.CustomerGroupID, CG.Name from CustomerGroups as CG with (NoLock) " & _
                    "where Deleted=0 and AnyCustomer<>1 and NewCardholders<>1 and AnyCardholder <> 1 and CustomerGroupID <> 2 "
            End If

            If EngineID = 6 Then
                MyCommon.QueryStr &= "and CG.CAMCustomerGroup=1 "
            Else
                MyCommon.QueryStr &= "and CG.CAMCustomerGroup=0 "
            End If
            If (Not String.IsNullOrEmpty(NotIn)) Then
                MyCommon.QueryStr = MyCommon.QueryStr & "and CG.CustomerGroupID " & NotIn
            End If

            'To disable option group for every engine             
            'MyCommon.QueryStr = MyCommon.QueryStr & " and isOptInGroup=0 "
            MyCommon.QueryStr &= nameLike & " order by CG.CustomerGroupID desc, CG.Name;"
            Dim sendString As String = ""
            Dim groupDT As DataTable = MyCommon.LRT_Select
            For Each groupRow As DataRow In groupDT.Rows
                sendString &= "<option value=""" & MyCommon.NZ(groupRow.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(groupRow.Item("Name"), "") & "</option>"
            Next

            Send(sendString)
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try

    End Sub

    Sub GetProductCollisions(ByVal OfferID As Long)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim row As DataRow
        Dim i As Integer = 0
        Dim CollisionThreshold As Integer = 0
        Dim CollisionScope As Integer = 0
        Dim ROID As Long = 0
        Dim EngineID As Integer = -1
        Dim EngineSubTypeID As Integer = 0
        Dim TargetStartDate As DateTime
        Dim TargetEndDate As DateTime
        Dim HasBundleDiscount As Boolean = False
        Dim Shaded As String = " style=""background-color:#e9e9e9;"""
        Dim ReturnString As New StringBuilder()

        CollisionThreshold = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(143))
        CollisionScope = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(144))

        Try
            MyCommon.Open_LogixRT()

            If OfferID > 0 Then
                MyCommon.QueryStr = "select O.EngineID, O.EngineSubTypeID, RO.RewardOptionID, I.StartDate, I.EndDate " & _
                                    "from OfferIDs as O with (NoLock) " & _
                                    "inner join CPE_RewardOptions as RO on RO.IncentiveID=O.OfferID " & _
                                    "inner join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                                    "where O.OfferID=" & OfferID & " and RO.Deleted=0;"
                dt = MyCommon.LRT_Select()

                If dt.Rows.Count > 0 Then
                    EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
                    EngineSubTypeID = MyCommon.NZ(dt.Rows(0).Item("EngineSubTypeID"), 0)
                    ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
                    TargetStartDate = MyCommon.NZ(dt.Rows(0).Item("StartDate"), "1/1/1900")
                    TargetEndDate = MyCommon.NZ(dt.Rows(0).Item("EndDate"), "1/1/1900")

                    'Determine if the offer has a group-level conditional (bundle) discount
                    MyCommon.QueryStr = "select DiscountID from CPE_Discounts as DIS with (NoLock) " & _
                                        "inner join CPE_Deliverables as DEL on DEL.OutputID=DIS.DiscountID " & _
                                        "where DIS.DiscountTypeID=4 and DEL.RewardOptionID=" & ROID & " and DEL.Deleted=0 and DIS.Deleted=0;"
                    dt2 = MyCommon.LRT_Select
                    If dt2.Rows.Count > 0 Then
                        HasBundleDiscount = True
                    End If

                    If EngineID = 2 And ROID > 0 Then
                        If CollisionScope = 1 Then
                            MyCommon.QueryStr = "select top " & CollisionThreshold & " OP.ProductID, OP.ExtProductID, OP.IncentiveID, OP.IncentiveName from " & _
                                                " (select distinct PGI.ProductID " & _
                                                "  from ProdGroupItems as PGI with (NoLock) " & _
                                                "  inner join CPE_IncentiveProductGroups as IPG with (NoLock) on IPG.ProductGroupID=PGI.ProductGroupID and IPG.Deleted=0 and PGI.Deleted=0 and IPG.ExcludedProducts=0 and IPG.Disqualifier=0 " & _
                                                "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                                                "  where RO.RewardOptionID=" & ROID & ") as DP " & _
                                                "inner join " & _
                                                "  (select distinct PGI.ProductID, P.ExtProductID, I.IncentiveID, I.IncentiveName " & _
                                                "  from ProdGroupItems as PGI with (NoLock) " & _
                                                "  inner join CPE_IncentiveProductGroups as IPG with (NoLock) on IPG.ProductGroupID=PGI.ProductGroupID and IPG.Deleted=0 and PGI.Deleted=0 and IPG.ExcludedProducts=0 and IPG.Disqualifier=0 " & _
                                                "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID and RO.Deleted=0 " & _
                                                "  inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID and I.Deleted=0 " & _
                                                "  inner join Products as P with (NoLock) on P.ProductID=PGI.ProductID " & _
                                                "  where RO.RewardOptionID<>" & ROID & " " & _
                                                "  and DATEADD(d, 1, I.EndDate)>GETDATE() " & _
                                                "  and ( " & _
                                                "      (I.StartDate>='" & TargetStartDate & "' and I.StartDate<='" & TargetEndDate & "') " & _
                                                "      or (I.EndDate>='" & TargetStartDate & "' and I.EndDate<='" & TargetEndDate & "') " & _
                                                "      or (I.StartDate<='" & TargetStartDate & "' and I.EndDate>='" & TargetEndDate & "') " & _
                                                "      ) " & _
                                                "  ) as OP " & _
                                                "on DP.ProductID=OP.ProductID " & _
                                                "order by IncentiveID, ExtProductID;"
                        ElseIf CollisionScope = 2 Then
                            If HasBundleDiscount Then
                                MyCommon.QueryStr = "select top " & CollisionThreshold & " OP.ProductID, OP.ExtProductID, OP.IncentiveID, OP.IncentiveName from " & _
                                                      " (select distinct PGI.ProductID " & _
                                                      "  from ProdGroupItems as PGI with (NoLock) " & _
                                                      "  inner join CPE_IncentiveProductGroups as IPG with (NoLock) on IPG.ProductGroupID=PGI.ProductGroupID and IPG.Deleted=0 and PGI.Deleted=0 and IPG.ExcludedProducts=0 and IPG.Disqualifier=0 " & _
                                                      "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                                                      "  where RO.RewardOptionID=" & ROID & ") as DP " & _
                                                      "inner join " & _
                                                      "  (select distinct PGI.ProductID, P.ExtProductID, I.IncentiveID, I.IncentiveName " & _
                                                      "  from ProdGroupItems as PGI with (NoLock) " & _
                                                      "  inner join CPE_Discounts as DIS with (NoLock) on DIS.DiscountedProductGroupID=PGI.ProductGroupID and DIS.Deleted=0 and PGI.Deleted=0 " & _
                                                      "  inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID=DIS.DiscountID and DEL.Deleted=0 " & _
                                                      "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID and RO.Deleted=0 " & _
                                                      "  inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID and I.Deleted=0 " & _
                                                      "  inner join Products as P with (NoLock) on P.ProductID=PGI.ProductID " & _
                                                      "  where RO.RewardOptionID<>" & ROID & " " & _
                                                      "  and DATEADD(d, 1, I.EndDate)>GETDATE() " & _
                                                      "  and (" & _
                                                      "      (I.StartDate>='" & TargetStartDate & "' and I.StartDate<='" & TargetEndDate & "')" & _
                                                      "      or (I.EndDate>='" & TargetStartDate & "' and I.EndDate<='" & TargetEndDate & "')" & _
                                                      "      or (I.StartDate<='" & TargetStartDate & "' and I.EndDate>='" & TargetEndDate & "')" & _
                                                      "      )" & _
                                                      "  ) as OP " & _
                                                      "on DP.ProductID=OP.ProductID " & _
                                                      "order by IncentiveID, ExtProductID;"
                            Else
                                MyCommon.QueryStr = "select top " & CollisionThreshold & " OP.ProductID, OP.ExtProductID, OP.IncentiveID, OP.IncentiveName from " & _
                                                    "  (select distinct PGI.ProductID " & _
                                                    "  from ProdGroupItems as PGI with (NoLock) " & _
                                                    "  inner join CPE_Discounts as DIS with (NoLock) on DIS.DiscountedProductGroupID=PGI.ProductGroupID and DIS.Deleted=0 and PGI.Deleted=0 " & _
                                                    "  inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID=DIS.DiscountID and DEL.Deleted=0 " & _
                                                    "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID " & _
                                                    "  where RO.RewardOptionID=" & ROID & ") as DP " & _
                                                    "inner join " & _
                                                    "  (select distinct PGI.ProductID, P.ExtProductID, I.IncentiveID, I.IncentiveName " & _
                                                    "  from ProdGroupItems as PGI with (NoLock) " & _
                                                    "  inner join CPE_Discounts as DIS with (NoLock) on DIS.DiscountedProductGroupID=PGI.ProductGroupID and DIS.Deleted=0 and PGI.Deleted=0 " & _
                                                    "  inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID=DIS.DiscountID and DEL.Deleted=0 " & _
                                                    "  inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID and RO.Deleted=0 " & _
                                                    "  inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID and I.Deleted=0 " & _
                                                    "  inner join Products as P with (NoLock) on P.ProductID=PGI.ProductID " & _
                                                    "  where RO.RewardOptionID<>" & ROID & " " & _
                                                    "  and DATEADD(d, 1, I.EndDate)>GETDATE() " & _
                                                    "  and (" & _
                                                    "      (I.StartDate>='" & TargetStartDate & "' and I.StartDate<='" & TargetEndDate & "')" & _
                                                    "      or (I.EndDate>='" & TargetStartDate & "' and I.EndDate<='" & TargetEndDate & "')" & _
                                                    "      or (I.StartDate<='" & TargetStartDate & "' and I.EndDate>='" & TargetEndDate & "')" & _
                                                    "      )" & _
                                                    "  ) as OP " & _
                                                    "on DP.ProductID=OP.ProductID " & _
                                                    "order by IncentiveID, ExtProductID;"
                            End If
                        End If
                        dt = MyCommon.LRT_Select
                        If dt.Rows.Count >= CollisionThreshold Then
                            If CollisionScope = 1 Then
                                ReturnString.Append("  <p>" & Copient.PhraseLib.Detokenize("CPEoffer-gen.FoundCollisions-Conditions", LanguageID, CollisionThreshold) & " " & Copient.PhraseLib.Lookup("CPEoffer-gen.DeployAnyway", LanguageID) & "</p>")
                            ElseIf CollisionScope = 2 Then
                                ReturnString.Append("  <p>" & Copient.PhraseLib.Detokenize("CPEoffer-gen.FoundCollisions-Rewards", LanguageID, CollisionThreshold) & " " & Copient.PhraseLib.Lookup("CPEoffer-gen.DeployAnyway", LanguageID) & "</p>")
                            End If
                            ReturnString.Append("  <table summary=""" & Copient.PhraseLib.Lookup("term.collisions", LanguageID) & """>")
                            ReturnString.Append("    <thead>")
                            ReturnString.Append("      <tr>")
                            ReturnString.Append("        <th style=""min-width:50px;"">" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</th>")
                            ReturnString.Append("        <th>" & Copient.PhraseLib.Lookup("term.offername", LanguageID) & "</th>")
                            ReturnString.Append("        <th>" & Copient.PhraseLib.Lookup("term.productid", LanguageID) & "</th>")
                            ReturnString.Append("      </tr>")
                            ReturnString.Append("    </thead>")
                            ReturnString.Append("    <tbody>")
                            For Each row In dt.Rows
                                ReturnString.Append("      <tr" & Shaded & ">")
                                ReturnString.Append("        <td>" & MyCommon.NZ(row.Item("IncentiveID"), 0) & "</td>")
                                ReturnString.Append("        <td>" & MyCommon.NZ(row.Item("IncentiveName"), "") & "</td>")
                                ReturnString.Append("        <td>" & MyCommon.NZ(row.Item("ExtProductID"), "") & "</td>")
                                ReturnString.Append("      </tr>")
                                If Shaded = "" Then
                                    Shaded = " style=""background-color:#e9e9e9;"""
                                Else
                                    Shaded = ""
                                End If
                            Next
                            ReturnString.Append("    </tbody>")
                            ReturnString.Append("  </table>")

                        End If
                    End If

                End If

            End If

            Send(ReturnString.ToString)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub ProcessProductCollisionsBackgroundUE(ByVal OfferID As Long, ByVal DeferDeploy As Boolean, Optional ByVal CallingLocation As Integer = 0, Optional ByVal ApprovalType As Integer = -1)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Try
            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()
            Threading.ThreadPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf BackgroundCollisionDetectionWorker), New Object() {OfferID, DeferDeploy, collisiondetectionservice, CallingLocation, ApprovalType})
            Sendb("True")
        Catch ex As Exception
            Send(ex.Message)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
        'MyCommon.Write_Log(AcceptanceLogFile, String.Format("Collision Detection Initiated for Offer ID: {0}", LogixID), True)
    End Sub

    Sub ApproveOffer(ByVal OfferID As Long, ByVal ApprovalType As Integer, ByVal OCDEnabled As Integer)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Try
            Dim ocdService As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()
            Dim oawService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)()
            Dim result As AMSResult(Of Boolean) = oawService.ApproveOffer(OfferID, AdminUserID)
            If (result.ResultType = AMSResultType.Success AndAlso result.Result = True) Then
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.offer-approved", LanguageID))
                If OCDEnabled = 1 AndAlso ApprovalType <> 13 Then
                    Try
                        Threading.ThreadPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf BackgroundCollisionDetectionWorker), New Object() {OfferID, 0, ocdService, 2, ApprovalType})
                    Catch ex As Exception
                        Send(ex.Message)
                    Finally
                        MyCommon.Close_LogixRT()
                        MyCommon = Nothing
                    End Try
                Else
                    iOfferID_CDS = OfferID
                    ApprovalType_CDS = ApprovalType
                    DetectOfferCollisionApprovalCallBack(New AMSResult(Of Int32)(0), False)
                End If
                Sendb("True")
            Else
                'Exception while approving offer.
            End If
        Catch ex As Exception
            Send(ex.Message)
        End Try
    End Sub

    Public Sub BackgroundCollisionDetectionWorker(State As Object)
        Dim obj As Object() = State
        Dim OfferID As Int64 = obj(0)
        Dim DeferDeploy As Boolean = obj(1)
        Dim collisiondetectionservice As ICollisionDetectionService = obj(2)
        Dim CallingLocation = obj(3)
        ApprovalType_CDS = obj(4)
        iOfferID_CDS = OfferID
        DeferDeployment_CDS = DeferDeploy
        If CallingLocation = 1 Then
            AddHandler collisiondetectionservice.OnDetectionComplete, AddressOf DetectOfferCollisionDetectionCallBack
            collisiondetectionservice.DetectOfferCollisionAsync(Of Int32)(OfferID, AdminUserID)
        ElseIf CallingLocation = 2 Then
            AddHandler collisiondetectionservice.OnDetectionComplete, AddressOf DetectOfferCollisionApprovalCallBack
            collisiondetectionservice.DetectOfferCollisionAsync(Of Int32)(OfferID, AdminUserID)
        Else
            AddHandler collisiondetectionservice.OnDetectionComplete, AddressOf DetectOfferCollisionCallBack
            collisiondetectionservice.DetectOfferCollisionAsync(Of Int32)(OfferID, AdminUserID)
        End If
    End Sub

    Public Sub DetectOfferCollisionDetectionCallBack(CollisionCount As AMSResult(Of Int32))
        Dim bCloseConn As Boolean = False
        Dim Common As New Copient.CommonInc
        If CollisionCount.ResultType <> AMSResultType.Success Then Exit Sub
        If Common.LRTadoConn.State = ConnectionState.Closed Then
            bCloseConn = True
            Common.Open_LogixRT()
        End If
        If CollisionCount.Result = 0 Then
            Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deploy-nocollisionsfound", LanguageID), iOfferID_CDS))
        Else
            Dim resolverbuilder As New WebRequestResolverBuilder()
            CurrentRequest.Resolver = resolverbuilder.GetResolver()
            resolverbuilder.Build()
            CurrentRequest.Resolver.AppName = "OfferFeeds.aspx"
            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)(CurrentRequest.Resolver.AppName)
            collisiondetectionservice.SendNotifications(iOfferID_CDS)
            Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deployfailed-collisionsfound", LanguageID), iOfferID_CDS))
        End If
        If (Common.LRTadoConn.State <> ConnectionState.Closed AndAlso bCloseConn = True) Then Common.Close_LogixRT()
    End Sub

    Public Sub DetectOfferCollisionApprovalCallBack(CollisionCount As AMSResult(Of Int32), Optional isOCDEnabled As Boolean = True)
        Dim bCloseConn As Boolean = False
        Dim Common As New Copient.CommonInc
        If CollisionCount.ResultType <> AMSResultType.Success Then Exit Sub
        If Common.LRTadoConn.State = ConnectionState.Closed Then
            bCloseConn = True
            Common.Open_LogixRT()
        End If
        If CollisionCount.Result = 0 Then
            If isOCDEnabled Then
                Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deploy-nocollisionsfound", LanguageID), iOfferID_CDS))
            End If
            Dim logText As String = ""
            Common.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
            If (ApprovalType_CDS = 13) Then
                Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
            ElseIf (ApprovalType_CDS = 14) Then
                Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
            ElseIf (ApprovalType_CDS = 15) Then
                Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID = @IncentiveID"
            End If

            Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
        Else
            Dim resolverbuilder As New WebRequestResolverBuilder()
            CurrentRequest.Resolver = resolverbuilder.GetResolver()
            resolverbuilder.Build()
            CurrentRequest.Resolver.AppName = "OfferFeeds.aspx"
            Common.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
            Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
            Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)(CurrentRequest.Resolver.AppName)
            collisiondetectionservice.SendNotifications(iOfferID_CDS)
            Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deployfailed-collisionsfound", LanguageID), iOfferID_CDS))
        End If
        If (Common.LRTadoConn.State <> ConnectionState.Closed AndAlso bCloseConn = True) Then Common.Close_LogixRT()
    End Sub

    Public Sub DetectOfferCollisionCallBack(CollisionCount As AMSResult(Of Int32))
        Dim bCloseConn As Boolean = False
        Dim Common As New Copient.CommonInc
        If CollisionCount.ResultType <> AMSResultType.Success Then Exit Sub
        If Common.LRTadoConn.State = ConnectionState.Closed Then
            bCloseConn = True
            Common.Open_LogixRT()
        End If
        Dim resolverbuilder As New WebRequestResolverBuilder()
        CurrentRequest.Resolver = resolverbuilder.GetResolver()
        resolverbuilder.Build()
        CurrentRequest.Resolver.AppName = "OfferFeeds.aspx"


        If CollisionCount.Result = 0 Then
            Dim m_OAWService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)(CurrentRequest.Resolver.AppName)
            Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deploy-nocollisionsfound", LanguageID), iOfferID_CDS))
            If ApprovalType_CDS <= 0 Then
                Dim m_isOAWEnabled As Boolean = False
                Dim Logix As New Copient.LogixInc
                If (Common.Fetch_SystemOption(66) = "1") Then
                    Dim BannerIds As Integer() = Logix.GetBannersForOffer(iOfferID_CDS)
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIds).Result
                Else
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
                End If
                If Not m_isOAWEnabled Then
                    If DeferDeployment_CDS = False Then
                        Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID = @IncentiveID"
                        Common.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
                    Else
                        Common.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID = @IncentiveID"
                        Common.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = iOfferID_CDS
                    End If
                    Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    If (Common.RowsAffected > 0) Then
                        Common.Activity_Log(3, iOfferID_CDS, 1, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
                    End If
                End If

            Else
                Dim logText As String = ""
                If (ApprovalType_CDS = 13) Then
                    logText = "Offer has been requested for approval."
                ElseIf (ApprovalType_CDS = 14) Then
                    logText = "Offer has been requested for approval and will be deployed once approved."
                ElseIf (ApprovalType_CDS = 15) Then
                    logText = "Offer has been requested for approval and will be defer deployed once approved."
                End If

                Common.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=" & ApprovalType_CDS & ", DeployDeferred=0, LastUpdate=getdate(), " &
                                  "LastUpdatedByAdminID =" & AdminUserID & " where IncentiveID=" & iOfferID_CDS
                Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
                If (Common.RowsAffected > 0) Then

                    m_OAWService.InsertUpdateOfferApprovalRecord(iOfferID_CDS, AdminUserID)
                    Common.Activity_Log(3, iOfferID_CDS, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(logText, LanguageID))
                End If
            End If
        Else
            Dim collisiondetectionservice As ICollisionDetectionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)(CurrentRequest.Resolver.AppName)
            collisiondetectionservice.SendNotifications(iOfferID_CDS)
            Common.Activity_Log(3, iOfferID_CDS, 1, String.Format(Copient.PhraseLib.Lookup("history.offer-deployfailed-collisionsfound", LanguageID), iOfferID_CDS))
        End If
        If (Common.LRTadoConn.State <> ConnectionState.Closed AndAlso bCloseConn = True) Then Common.Close_LogixRT()
    End Sub

    Sub GetProductCollisionsUE(ByVal OfferID As Long, ByVal DeferDeploy As Boolean, Optional ByVal CallingLocation As Integer = 0, Optional ByVal ApprovalType As Integer = -1)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Dim Shaded As String = " style=""background-color:#e9e9e9;"""
        Dim ReturnString As New StringBuilder()
        Dim CollisionProducts As AMSResult(Of Integer)
        Dim collisionService As ICollisionDetectionService

        Try
            collisionService = CurrentRequest.Resolver.Resolve(Of ICollisionDetectionService)()

            Dim collisionServiceURL As String = MyCommon.Fetch_UE_SystemOption(185)
            If (String.IsNullOrWhiteSpace(collisionServiceURL)) Then
                CollisionProducts = New AMSResult(Of Integer)(AMSResultType.ValidationError, Copient.PhraseLib.Lookup("term.undefinedcollisionserviceurl", LanguageID))
            Else
                'calling RestCommAPI to get collision detection count. 
                CollisionProducts = collisionService.DetectOfferCollision(OfferID, AdminUserID)
            End If

            Dim loadingBoxID As String = IIf(Request.QueryString("Mode") = "GetProductCollisionsUEDetection", "loadingDetection", "loading")

            If CollisionProducts.ResultType <> AMSResultType.Success Then
                ReturnString.Append("<p style="" background-color:red; color:white;   font-weight: bold;"">" + CollisionProducts.MessageString + "</p>")
                ReturnString.Append("<br/><br/><br/><br/>")
                ReturnString.Append("        <p style=""text-align:center;"">")
                ReturnString.Append("          <input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ onclick=""javascript:toggleDialog('" & loadingBoxID & "', false);"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ />")
                ReturnString.Append("        </p>")
                Send(ReturnString.ToString)
                Exit Try
            End If
            If (CollisionProducts.Result > 0) Then
                ReturnString.Append("        <br/>")
                ReturnString.Append("        <p>" & Copient.PhraseLib.Lookup("term.collidingproductsfound", LanguageID) & " :<span style=""Color:Red;"">" & CollisionProducts.Result & "</span><p>")
                ReturnString.Append("<p>" & Copient.PhraseLib.Lookup("term.collisionsfound", LanguageID) & "<p>")
                ReturnString.Append("        <p style=""text-align:center;"">")
                ReturnString.Append("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """ />")
                If (CallingLocation <> 1) Then
                    ReturnString.Append("          <input type=""submit"" class=""regular"" id=""viewreport"" name=""viewreport"" value=""" & Copient.PhraseLib.Lookup("term.viewreport", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.viewreport", LanguageID) & """ />")
                    If (ApprovalType = -1) Then
                        If (DeferDeploy) Then
                            ReturnString.Append("          <input type=""submit"" class=""regular"" id=""confirmingDeferDeploy"" name=""deferdeploy"" value=""" & Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & """/>")
                        Else
                            ReturnString.Append("          <input type=""submit"" class=""regular"" id=""collisionDeploy"" name=""deploy"" value=""" & Copient.PhraseLib.Lookup("term.deployoffer", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.deployoffer", LanguageID) & """/>")
                        End If
                        ReturnString.Append("          <input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("term.canceldeployment", LanguageID) & """ onclick=""javascript:toggleDialog('loading', false);"" title=""" & Copient.PhraseLib.Lookup("term.canceldeployment", LanguageID) & """ />")
                    Else
                        If ApprovalType = 13 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmApproval"" name=""reqApproval"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        ElseIf ApprovalType = 14 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmDeployApproval"" name=""reqApprovalWithDeployment"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        ElseIf ApprovalType = 15 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmDeferDeployApproval"" name=""reqApprovalWithDeferDeployment"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        End If
                        ReturnString.Append("          <input type=""button"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:toggleDialog('loading', false);"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ />")
                    End If

                Else
                    ReturnString.Append("          <input type=""submit"" class=""mediumshort"" id=""viewreport"" name=""viewreport"" value=""" & Copient.PhraseLib.Lookup("term.viewreport", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.viewreport", LanguageID) & """ />")
                    ReturnString.Append("          <input type=""button"" class=""mediumshort"" value=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ onclick=""javascript:toggleDialog('loadingDetection', false);"" title=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ />")
                End If
                ReturnString.Append("        </p>")
                Send(ReturnString.ToString)
            Else
                ReturnString.Append("        <p> " & Copient.PhraseLib.Lookup("term.nocollisions", LanguageID) & "</p>")
                ReturnString.Append(" <p style=""text-align: center; padding-top: 50px;"">")
                ReturnString.Append("<input type=""hidden"" name=""OfferID"" value=""" & OfferID & """ />")
                If (CallingLocation <> 1) Then
                    If (ApprovalType = -1) Then
                        If (DeferDeploy) Then
                            ReturnString.Append("          <input type=""submit"" class=""regular"" id=""confirmingDeferDeploy"" name=""deferdeploy"" value=""" & Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & """/>")
                        Else
                            ReturnString.Append("          <input type=""submit"" class=""regular"" id=""collisionDeploy"" name=""deploy"" value=""" & Copient.PhraseLib.Lookup("term.deployoffer", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.deployoffer", LanguageID) & """/>")
                        End If
                        ReturnString.Append("          <span style=""padding:15px""> ")
                        ReturnString.Append("          </span>")
                        ReturnString.Append("<input type=""button"" class=""large""  value=""" & Copient.PhraseLib.Lookup("term.canceldeployment", LanguageID) & """ onclick=""javascript:toggleDialog('loading', false);"" title=""" & Copient.PhraseLib.Lookup("term.canceldeployment", LanguageID) & """ />")
                    Else
                        If ApprovalType = 13 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmApproval"" name=""reqApproval"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        ElseIf ApprovalType = 14 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmDeployApproval"" name=""reqApprovalWithDeployment"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        ElseIf ApprovalType = 15 Then
                            ReturnString.Append("           <input type=""submit"" class=""large"" id=""confirmDeferDeployApproval"" name=""reqApprovalWithDeferDeployment"" title=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.submitapproval", LanguageID) & """ />")
                        End If
                        ReturnString.Append("<input type=""button"" class=""large""  value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:toggleDialog('loading', false);"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ />")
                    End If

                Else
                    ReturnString.Append("          <input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ onclick=""javascript:toggleDialog('loadingDetection', false);"" title=""" & Copient.PhraseLib.Lookup("term.ok", LanguageID) & """ />")
                End If
                ReturnString.Append("</p>")
                Send(ReturnString.ToString)
            End If
        Catch ex As Exception
            ReturnString.Append("<p style="" background-color:red; color:white;   font-weight: bold;"">" + ex.ToString + "</p>")
            ReturnString.Append("<br/><br/><br/><br/>")
            ReturnString.Append("        <p style=""text-align:center;"">")
            If (CallingLocation <> 1) Then
                ReturnString.Append("          <input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ onclick=""javascript:toggleDialog('loading', false);"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ />")
            Else
                ReturnString.Append("          <input type=""button"" class=""large"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ onclick=""javascript:toggleDialog('loadingDetection', false);"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ />")
            End If
            ReturnString.Append("        </p>")
            Send(ReturnString.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub GeneratePredefinedRecTextMessages(ByVal BaseRecTextId As String, ByVal MLID As String)
        Dim rst As DataTable
        Dim row As DataRow
        Dim LangIds As String = ""
        Dim LangRecMsgs As String = ""
        Dim intCnt As Integer = 0
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        Try
            Dim baseLangId As Integer = MyCommon.Fetch_SystemOption(1)
            MyCommon.QueryStr = "Select isnull(ML.ReceiptTextMsg,'')'ReceiptTextMsg',L.MsNetCode " & _
                                "from Languages L  with (NoLock) inner join TransLanguagesCF_CPE TC with (NoLock) " & _
              "on L.LanguageID = TC.LanguageID left outer join " & _
                                "(Select ReceiptTextMsg,Isnull(LanguageID,0) LanguageID from " & _
              "CPE_ReceiptTextMessages with (NoLock) where  isnull(BaseReceiptTextID,0) in " & _
              "(select ReceiptTextID from CPE_ReceiptTextMessages with (NoLock) where " & _
              "ReceiptTextMsg='" & BaseRecTextId & "' and isnull(BaseReceiptTextID,0)=0))ML " & _
              " on L.LanguageID = ML.LanguageID where L.LanguageID <> " & baseLangId.ToString
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                    If LangIds = "" Then
                        LangIds = row.Item("MsNetCode")
                    Else
                        LangIds = LangIds & row.Item("MsNetCode")
                    End If
                    If LangRecMsgs = "" Then
                        LangRecMsgs = row.Item("ReceiptTextMsg")
                    Else
                        LangRecMsgs = LangRecMsgs & row.Item("ReceiptTextMsg")
                    End If
                    If intCnt <> rst.Rows.Count - 1 Then
                        LangIds = LangIds & Chr(20)
                        LangRecMsgs = LangRecMsgs & Chr(20)
                        intCnt += 1
                    End If
                Next
            End If
            Send("<ML>" & MLID & "</ML>")
            Send("<LangIds>" & LangIds & "</LangIds>")
            Send("<LangRecMsgs>" & LangRecMsgs & "</LangRecMsgs>")

        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Function GetUDFStringValues(ByVal UDFPK As Long, ByVal OfferID As Long) As String
        Dim rst As DataTable
        Dim UDFStringValue As String = ""
        Dim bOpenedRTConnection As Boolean = False

        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bOpenedRTConnection = True
            End If
            MyCommon.QueryStr = "select Stringvalue from OfferUDFStringValues where OfferID =" & OfferID & "  and  UDFPK = " & UDFPK
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                UDFStringValue = rst.Rows(0).Item("Stringvalue").ToString()
            Else
                MyCommon.QueryStr = "select Stringvalue from UserDefinedFieldsValues where deleted  = 0 and OfferID =" & OfferID & "  and  UDFPK = " & UDFPK
                Dim rst2 As DataTable = MyCommon.LRT_Select
                If rst2.Rows.Count > 0 Then
                    If Not System.Convert.IsDBNull(rst2.Rows(0).Item("Stringvalue")) Then
                        UDFStringValue = rst2.Rows(0).Item("Stringvalue").ToString()
                    Else
                        MyCommon.QueryStr = "Select value from UserDefinedField_ValidValues where UDFPK = " + Convert.ToString(UDFPK) + " and isDefault=1"
                        Dim df As DataTable = MyCommon.LRT_Select
                        If df.Rows.Count = 1 Then
                            UDFStringValue = Convert.ToString(df.Rows(0).Item("value"))
                        Else
                            UDFStringValue = ""
                        End If
                    End If
                Else
                    UDFStringValue = ""
                End If
            End If
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            If bOpenedRTConnection = True Then
                If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
                    MyCommon.Close_LogixRT()
                End If
            End If
        End Try
        Return UDFStringValue
    End Function

    Sub DuplicateoffersFromTemplete(ByVal OfferID As Long, ByVal EngineId As Integer, ByVal noOfTimes As Integer)
        Dim index As Integer
        Dim bOpenedRTConnection As Boolean = False
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            MyCommon.Open_LogixRT()
            bOpenedRTConnection = True
        End If
        Dim bUseOfferRedemptionThreshold As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(83) = "1", True, False)
        Dim bUseDisplayDates As Boolean = False
        If (EngineId = 9 AndAlso MyCommon.Fetch_UE_SystemOption(143) = "1") Then
            bUseDisplayDates = True
        ElseIf (EngineId = 0 AndAlso MyCommon.Fetch_CM_SystemOption(85) = "1") Then
            bUseDisplayDates = True
        End If

        Dim bUseMultipleProductExclusionGroups As Boolean = True
        Dim SourceOfferID As Long = OfferID

        If EngineId = 0 Then
            Dim bCopyInboundCrmEngineID As Boolean
            If MyCommon.Fetch_CM_SystemOption(107) = "1" Then
                bCopyInboundCrmEngineID = True
            Else
                bCopyInboundCrmEngineID = False
            End If

            For index = 1 To noOfTimes
                MyCommon.QueryStr = "dbo.pc_Create_CM_OfferFromTemplate"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@TemplateID", SqlDbType.NVarChar, 200).Value = SourceOfferID
                MyCommon.LRTsp.Parameters.Add("@CopyInboundCRM", SqlDbType.Bit).Value = IIf(bCopyInboundCrmEngineID, 1, 0)
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferID = MyCommon.LRTsp.Parameters("@OfferID").Value
                MyCommon.Close_LRTsp()

                CreateNewLocalPromotionVariables(OfferID, MyCommon)
                If bUseDisplayDates Then
                    'Updating TemplatePermission table with the Disallow_DisplayDates based on the SystemOption #85
                    UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 85)
                    SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
                End If

                If (bUseOfferRedemptionThreshold) Then
                    'Updating TemplatePermission table with the Disallow_OfferRedempThreshold based on the SystemOption #83
                    UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 83)
                    SaveOfferThresholdPerHourValue(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
                End If
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("offer.createdfromtemplate", LanguageID) & ": " & SourceOfferID)
            Next index
        ElseIf EngineId = 9 Then
            ' dbo.pc_CreateOfferFromTemplate @TemplateID bigint, @OfferID bigint OUTPUT
            For index = 1 To noOfTimes
                MyCommon.QueryStr = "dbo.pc_Create_CPE_OfferFromTemplate"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = SourceOfferID
                MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()

                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                MyCommon.Close_LRTsp()
                Dim m_isOAWEnabled As Boolean = False
                Dim m_OAWService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)(CurrentRequest.Resolver.AppName)
                Dim Logix As New Copient.LogixInc
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    Dim BannerIds As Integer() = Logix.GetBannersForOffer(OfferID)
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIds).Result
                Else
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
                End If

                If OfferID > 0 AndAlso m_isOAWEnabled Then
                    m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)
                End If
                If bUseDisplayDates Then
                    'Updating TemplatePermission table with the Disallow_DisplayDates based on the UE SystemOption #143
                    UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 143)
                    SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
                End If

                If (OfferID > 0 AndAlso bUseMultipleProductExclusionGroups) Then
                    Dim TargetROID As Long = 0
                    MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID  = " & OfferID & ";"
                    Dim dt As New DataTable

                    dt = MyCommon.LRT_Select()
                    If dt.Rows.Count > 0 Then
                        TargetROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
                    End If

                    If (TargetROID > 0) Then
                        MyCommon.QueryStr = "dbo.pc_Duplicate_CPE_ProductCondition_MulitpleExcludedProductCondition"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@TargetROID", SqlDbType.BigInt).Value = TargetROID
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    End If
                End If

                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("offer.createdfromtemplate", LanguageID) & ": " & SourceOfferID)
            Next index
        End If
        If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        If noOfTimes = 1 Then
            Send("OfferID=" & OfferID)
        Else
            Send("OK")
        End If
    End Sub

    Sub SaveOfferThresholdPerHourValue(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)
        Dim OfferRedemptionThresholdperHour As Integer = 0
        Common.QueryStr = "SELECT RedemThresholdPerHour FROM offerAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
        Dim dtRedemption As New DataTable
        dtRedemption = Common.LRT_Select()
        If dtRedemption.Rows.Count > 0 Then
            OfferRedemptionThresholdperHour = Common.NZ(dtRedemption.Rows(0).Item("RedemThresholdPerHour"), 0)
        End If
        Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = DBNull.Value
        Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = DBNull.Value
        'Updating only OfferRedemptionThresholdperHour because it depends on CM SystemOption #83,  'pa_UpdateOfferAccessoryFields' contains logic to insert/update based on the engineID and systemoption passed     
        Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = OfferRedemptionThresholdperHour
        Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId
        Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 83
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    End Sub

    Sub UpdateTemplatePermissions(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal systemOption As Integer)

        Dim dtTempPermission As New DataTable
        Dim Disallow_DisplayDates As Integer = Integer.MinValue
        Dim Disallow_OfferRedempThreshold As Integer = Integer.MinValue

        If (systemOption = 85) Then
            Common.QueryStr = "SELECT Disallow_DisplayDates from TemplatePermissions with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
            dtTempPermission = Common.LRT_Select()
            If dtTempPermission.Rows.Count > 0 Then
                Disallow_DisplayDates = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_DisplayDates"))
            End If
            Common.QueryStr = "UPDATE TemplatePermissions with (RowLock) Set Disallow_DisplayDates=" & Disallow_DisplayDates & " WHERE OfferID = " & OfferID
            Common.LRT_Execute()
        ElseIf (systemOption = 83) Then
            Common.QueryStr = "SELECT Disallow_OfferRedempThreshold from TemplatePermissions  with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
            dtTempPermission = Common.LRT_Select()
            If dtTempPermission.Rows.Count > 0 Then
                Disallow_OfferRedempThreshold = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_OfferRedempThreshold"))
            End If
            Common.QueryStr = "UPDATE TemplatePermissions  with (RowLock) Set Disallow_OfferRedempThreshold=" & Disallow_OfferRedempThreshold & " WHERE OfferID = " & OfferID
            Common.LRT_Execute()
        End If
        If (systemOption = 143) Then
            Common.QueryStr = "SELECT Disallow_DisplayDates from TemplatePermissions with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
            dtTempPermission = Common.LRT_Select()
            If dtTempPermission.Rows.Count > 0 Then
                Disallow_DisplayDates = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_DisplayDates"))
            End If
            Common.QueryStr = "UPDATE TemplatePermissions with (RowLock) Set Disallow_DisplayDates=" & Disallow_DisplayDates & " WHERE OfferID = " & OfferID
            Common.LRT_Execute()
        End If

    End Sub


    Sub SaveOfferDisplayDates(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)
        Common.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
        Dim dtODisp As New DataTable
        Dim startDate As String = ""
        Dim endDate As String = ""
        dtODisp = Common.LRT_Select()
        If dtODisp.Rows.Count > 0 Then
            startDate = Common.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
            endDate = Common.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
        End If
        Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(startDate), DBNull.Value, startDate)
        Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(endDate), DBNull.Value, endDate)

        Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
        Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId

        Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 85
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

    End Sub

    Sub CreateNewLocalPromotionVariables(ByVal lOfferId As Long, ByRef Mycommon As Copient.CommonInc)
        Dim lRewardID As Long
        Dim lPromoVarId As Long
        Dim rst As DataTable
        Dim row As DataRow

        ' create local promotion variables for this new offer
        Mycommon.QueryStr = "select OfferID from Offers with (NoLock) where OfferID=" & lOfferId & _
                            " and DistPeriodLimit > 0.00 and DistPeriod <> 0 and DistPeriodVarID=0;"
        rst = Mycommon.LRT_Select
        For Each row In rst.Rows
            Mycommon.Open_LogixXS()
            Mycommon.QueryStr = "dbo.pc_DistributionVar_Create"
            Mycommon.Open_LXSsp()
            Mycommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferId
            Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            Mycommon.LXSsp.ExecuteNonQuery()
            lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
            Mycommon.Close_LXSsp()
            Mycommon.Close_LogixXS()
            Mycommon.QueryStr = "update Offers with (RowLock) set DistPeriodVarID=" & lPromoVarId & " where OfferID=" & lOfferId & ";"
            Mycommon.LRT_Execute()
        Next

        ' create local promotion variables for this new offer's rewards
        Mycommon.QueryStr = "select RewardID from OfferRewards with (NoLock) where OfferID=" & lOfferId & _
                            " and RewardLimit > 0.00 and RewardDistPeriod <> 0 and RewardDistLimitVarID=0;"
        rst = Mycommon.LRT_Select
        For Each row In rst.Rows
            lRewardID = Mycommon.NZ(row.Item("RewardID"), 0)
            Mycommon.Open_LogixXS()
            Mycommon.QueryStr = "dbo.pc_RewardLimitVar_Create"
            Mycommon.Open_LXSsp()
            Mycommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = lRewardID
            Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            Mycommon.LXSsp.ExecuteNonQuery()
            lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
            Mycommon.Close_LXSsp()
            Mycommon.Close_LogixXS()
            Mycommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & lPromoVarId & " where RewardID=" & lRewardID & ";"
            Mycommon.LRT_Execute()
        Next
    End Sub


    Sub UpdateUDFStringValues(ByVal UDFPK As Long, ByVal OfferID As Long, ByVal UDFStringText As String)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Dim SpecialCharacters As String = MyCommon.Fetch_SystemOption(171)
        'UDFStringText = CleanString(UDFStringText, SpecialCharacters)
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        Try
            MyCommon.QueryStr = "dbo.pt_UDFStringValuesUpdate"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 1000).Value = UDFStringText
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub StoreUserPane(ByVal UserID As Integer)
        Dim dtStoreUser As DataTable
        Dim row As DataRow
        Dim maxStoreUser As Integer = 0
        Try

            If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(131), maxStoreUser) Then
                maxStoreUser = 0
            End If

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "select loc.LocationName, loc.LocationID from Locations loc inner join storeusers su on su.LocationID = loc.LocationID where loc.Deleted = 0 and su.UserID =  " & UserID
            dtStoreUser = MyCommon.LRT_Select

            For Each row In dtStoreUser.Rows
                Send("<option value =""userLoc-" & row.Item("LocationID") & """ >" & MyCommon.NZ(row.Item("LocationName"), "(Unkown)") & "</option>")
            Next

        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub SaveStoreUserLocations(ByVal UserID As Integer, ByVal LocationList As String)
        Dim dtStoreUsers As DataTable
        Dim LocArray() As String = Split(LocationList)
        Dim QueryStr As String
        Dim row As DataRow


        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            If LocationList.Length = 0 Then
                MyCommon.QueryStr = "delete from storeusers where UserID = " & UserID 'no records saved. delete all for this user
                MyCommon.LRT_Execute()
            Else
                LocationList = Left(LocationList, LocationList.Length - 1) ' deletes extra comma at the end
                QueryStr = "Create Table #tempstoreusers (UserID int not null, LocationID int not null); "
                QueryStr += "Insert into #tempstoreusers values " & LocationList & "; "
                QueryStr += "delete from storeusers where locationid not in (Select locationid from #tempstoreusers) and UserID = " & UserID & "; "
                QueryStr += "insert into storeusers (UserID, locationid) select userid, locationid from #tempstoreusers where locationid not in (select locationid from storeusers where UserID = " & UserID & "); "
                QueryStr += "drop table #tempstoreusers "
                QueryStr += "select loc.LocationName, loc.LocationID from Locations loc inner join storeusers su on su.LocationID = loc.LocationID where su.UserID = " & UserID
                MyCommon.QueryStr = QueryStr
                dtStoreUsers = MyCommon.LRT_Select

                For Each row In dtStoreUsers.Rows
                    Send("<tr><td id=""Store-" & row.Item("LocationID") & """ > &#8226; " & row.Item("LocationName") & "</td></tr>")
                Next
            End If


        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try

    End Sub

    Sub StoreUserLocationSearch(ByVal UserID As Integer, ByVal searchTerms As String, ByVal searchOption As Integer)
        Dim dtResults As DataTable
        Dim row As DataRow

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "select Top 100 LocationName, LocationID from locations where Deleted =0 and LocationTypeID = 1 " ' and locationid not in (select locationid from storeusers where adminuserid = " & UserID & ") "


            If searchTerms <> "" Then
                MyCommon.QueryStr += " and LocationName like '" & "@SearchString" & "%';"
                MyCommon.DBParameters.AddWithValue("@SearchString", IIf(searchOption = 2, "%", "") & searchTerms)
            End If

            dtResults = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            For Each row In dtResults.Rows
                Send("<option value =""storeLoc-" & row.Item("LocationID") & """ >" & MyCommon.NZ(row.Item("LocationName"), "(Unkown)") & "</option>")
            Next

        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub

    Sub GetDespForSVProgID(ByVal svProgID As Integer)
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Try
            MyCommon.Open_LogixRT()
            Dim sendString As String = ""
            Dim rst As DataTable = Nothing
            Dim lan As DataTable = Nothing
            Dim DefaultLanguageID As Integer = 0
            Dim bEnablemulLan As Integer = 0
            Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
            bEnablemulLan = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)

            'Get the description of the stored value program for each language
            If (bEnablemulLan) Then
                MyCommon.QueryStr = "SELECT L.LanguageID, L.MSNetCode FROM Languages AS L " & _
                    " LEFT JOIN CPE_DeliverableMonSVTranslations AS T ON T.LanguageID=L.LanguageID AND T.SVProgramID=" & svProgID & _
                    " WHERE L.LanguageID in (SELECT TLV.LanguageID FROM TransLanguagesCF_UE AS TLV) " & _
                    " ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
                lan = MyCommon.LRT_Select()
                If lan.Rows.Count > 0 Then
                    For Each rowL As DataRow In lan.Rows
                        MyCommon.QueryStr = "SELECT Description FROM CPE_DeliverableMonSVTranslations with (NoLock) WHERE SVProgramID = " & svProgID & " and LanguageID = " & MyCommon.NZ(rowL.Item("LanguageID"), 0)
                        rst = MyCommon.LRT_Select()
                        If rst.Rows.Count > 0 Then
                            sendString &= MyCommon.NZ(rowL.Item("MSNetCode"), 0) & "|" & Trim(MyCommon.NZ(rst.Rows(0).Item("Description"), "")) & "~"
                        Else
                            MyCommon.QueryStr = "SELECT Description FROM CPE_DeliverableMonSVTranslations with (NoLock) WHERE SVProgramID = " & svProgID & " "
                            Dim firstSV As DataTable = Nothing
                            firstSV = MyCommon.LRT_Select()
                            If firstSV.Rows.Count = 0 Then
                                'Getting the description from Stored value program table when SV program is assigned.
                                MyCommon.QueryStr = "SELECT Description, Name FROM StoredValuePrograms with (NoLock) WHERE SVProgramID = " & svProgID
                                Dim dtSVDesc As DataTable = Nothing
                                dtSVDesc = MyCommon.LRT_Select()
                                If dtSVDesc.Rows.Count > 0 Then
                                    sendString &= MyCommon.NZ(dtSVDesc.Rows(0).Item("Description"), "")
                                    Exit For
                                End If
                            Else
                                sendString &= MyCommon.NZ(rowL.Item("MSNetCode"), 0) & "|" & "" & "~"
                            End If
                        End If
                    Next
                End If
            Else
                MyCommon.QueryStr = "SELECT Description, Name FROM StoredValuePrograms with (NoLock) WHERE SVProgramID = " & svProgID
                rst = MyCommon.LRT_Select()
                If rst.Rows.Count > 0 Then
                    sendString &= MyCommon.NZ(rst.Rows(0).Item("Description"), "")
                End If
            End If
            If sendString.Length > 0 AndAlso sendString.Contains("~") Then
                sendString.Trim("~")
            End If
            Send(sendString)

        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    End Sub
</script>
