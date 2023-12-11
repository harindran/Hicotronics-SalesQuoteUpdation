Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.Text.RegularExpressions
Imports SAPbobsCOM

Namespace SalesQuoteUpdation
    <FormAttribute("OQUT", "SalesQuote/SalesQuoteUpdate.b1f")>
    Friend Class SalesQuoteUpdate
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("105").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("MtQuote").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("lbuy").Specific, SAPbouiCOM.StaticText)
            Me.StaticText1 = CType(Me.GetItem("lsqnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txtentry").Specific, SAPbouiCOM.EditText)
            Me.EditText2 = CType(Me.GetItem("DocEntry").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lmake").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("CustN").Specific, SAPbouiCOM.EditText)
            Me.EditText5 = CType(Me.GetItem("DueDate").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lduedate").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("cmbmake").Specific, SAPbouiCOM.ComboBox)
            Me.ComboBox1 = CType(Me.GetItem("BuyName").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText4 = CType(Me.GetItem("lcust").Specific, SAPbouiCOM.StaticText)
            Me.Button2 = CType(Me.GetItem("btnGet").Specific, SAPbouiCOM.Button)
            Me.StaticText5 = CType(Me.GetItem("lFrmDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("FromDate").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lTDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("ToDate").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("lSQCateg").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox2 = CType(Me.GetItem("SQCat").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText8 = CType(Me.GetItem("lmpn").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox3 = CType(Me.GetItem("cmbmpn").Specific, SAPbouiCOM.ComboBox)
            Me.EditText6 = CType(Me.GetItem("txtMake").Specific, SAPbouiCOM.EditText)
            Me.EditText7 = CType(Me.GetItem("txtMPN").Specific, SAPbouiCOM.EditText)
            Me.CheckBox1 = CType(Me.GetItem("chkmake").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox2 = CType(Me.GetItem("chkmpn").Specific, SAPbouiCOM.CheckBox)
            Me.Button3 = CType(Me.GetItem("btnMake").Specific, SAPbouiCOM.Button)
            Me.Button4 = CType(Me.GetItem("btnMPN").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Public Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("OQUT", Me.FormCount)
                objform.Freeze(True)
                'objform.ActiveItem = "txtentry"
                SuperUser = ValidateSuperUser()
                Matrix0.AddRow()
                LoadCombo()
                If SuperUser Then
                    ComboBox1.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    ComboBox1.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
                    objform.ActiveItem = "cmbmake"
                    ComboBox1.Item.Enabled = False
                End If
                ComboBox0.Select("-", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                ComboBox3.Select("-", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                EditText6.Item.Visible = False
                Button3.Item.Visible = False
                Button4.Item.Visible = False
                EditText7.Item.Visible = False
                objform.Items.Item("DocEntry").Visible = False
                objaddon.objapplication.Menus.Item("1300").Activate()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Public Function ValidateSuperUser()
            Dim strsql As String
            Dim objrs As SAPbobsCOM.Recordset
            strsql = "select  ""U_SUser"" from ""@BUYERMAKE"" where ifnull(""U_SUser"",'')='Y' and ""U_Buy""='" & objaddon.objapplication.Company.UserName & "' "
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount > 0 Then If objrs.Fields.Item("U_SUser").Value.ToString.ToUpper = "Y" Then Return True
            Return False
        End Function

        Private Sub LoadCombo()
            Try
                Dim objCombo As SAPbouiCOM.Column
                Dim StrQuery As String = ""
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If SuperUser Then
                    StrQuery = "select Distinct ""U_Buy"",""U_BuyName"" from ""@BUYERMAKE"" where ""U_Buy""<>'' and ""U_BuyName""<>''"
                Else
                    StrQuery = "select Top 1 ""U_Buy"",""U_BuyName"" from ""@BUYERMAKE"" where  ""U_Buy""='" & objaddon.objapplication.Company.UserName & "' and ""U_Buy"" <>'' "
                End If
                objRs.DoQuery(StrQuery)

                If objRs.RecordCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("Please Update the Valid User Name for this User in BuyerMake UDT...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                ComboBox1.ValidValues.Add("-", "-")   'BuyerName
                ComboBox1.ValidValues.Add("-1", "All")

                For i As Integer = 0 To objRs.RecordCount - 1
                    ComboBox1.ValidValues.Add(objRs.Fields.Item("U_Buy").Value.ToString, objRs.Fields.Item("U_BuyName").Value.ToString)
                    objRs.MoveNext()
                Next
                If SuperUser Then
                    StrQuery = "select Distinct ""U_Make"" from ""@BUYERMAKE"" where ""U_Make""<>''"
                Else
                    StrQuery = "select Distinct ""U_Make"" from ""@BUYERMAKE""  where  ""U_Buy""='" & objaddon.objapplication.Company.UserName & "' and ""U_Make""<>''"
                End If
                objRs.DoQuery(StrQuery)
                If objRs.RecordCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("Please Update the Make for this User in BuyerMake UDT...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                ComboBox0.ValidValues.Add("-", "-")   'Make
                ComboBox0.ValidValues.Add("-1", "All")
                For i As Integer = 0 To objRs.RecordCount - 1
                    ComboBox0.ValidValues.Add((i + 1), objRs.Fields.Item("U_Make").Value.ToString)
                    objRs.MoveNext()
                Next

                StrQuery = "select Distinct ""U_PrtNo"" from QUT1 where ""U_PrtNo""<>''"
                objRs.DoQuery(StrQuery)
                If objRs.RecordCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("MPN Not found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                ComboBox3.ValidValues.Add("-", "-")   'MPN
                ComboBox3.ValidValues.Add("-1", "All")
                For i As Integer = 0 To objRs.RecordCount - 1
                    ComboBox3.ValidValues.Add((i + 1), objRs.Fields.Item("U_PrtNo").Value.ToString)
                    objRs.MoveNext()
                Next

                StrQuery = "select ""CurrCode"",""CurrName"" from OCRN"
                objRs.DoQuery(StrQuery)
                objCombo = Matrix0.Columns.Item("Curr")
                For i As Integer = 0 To objRs.RecordCount - 1
                    objCombo.ValidValues.Add(objRs.Fields.Item("CurrCode").Value.ToString, objRs.Fields.Item("CurrName").Value.ToString)
                    objRs.MoveNext()
                Next
                objRs = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub EditText1_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText1.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_1")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "DocStatus"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "O"
                ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        

                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText1_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText1.Value = pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value
                    Catch ex As Exception
                        EditText1.Value = pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value
                    End Try
                    Try
                        EditText2.Value = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                    Catch ex As Exception
                        EditText2.Value = pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value
                    End Try
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText1_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.LostFocusAfter
            'Throw New System.NotImplementedException()
            'Matrix0.AddRow()
            'If objform.Items.Item("txtentry").Specific.String <> "" Then
            '    objform.Freeze(True)
            '    LoadSQScreen(objform.Items.Item("txtentry").Specific.String)
            '    objform.Freeze(False)
            'End If
        End Sub


        Private Sub LoadSQScreen(ByVal SQNum As String)
            Dim objRecordset As SAPbobsCOM.Recordset
            Dim odbdsDetails As SAPbouiCOM.DBDataSource
            Dim Str As String = "", Make As String = ""
            'Dim objCurrency As SAPbouiCOM.ComboBox
            Try
                objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Str = "select T1.""LineNum"",T1.""ItemCode"",T1.""Dscription"",T1.""U_Make"" as ""Make"",T1.""Quantity"",T1.""U_PrtNo"" as ""MPN"",T1.""U_SPQ"" as ""SPQ"","
                'Str += "T1.""U_MOQ"" as ""MOQ"",T1.""Price"",T1.""U_SupRefNo"" as ""SupRefNo"",T1.""U_QutValid"" as ""QutValid"",T1.""U_Cost"" as ""Cost"",T1.""U_Remarks"" ""Remarks"",T1.""U_LeadTime"" as ""LeadTime"" from OQUT T0 join QUT1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                'Str += "join OITM T2 on T1.""ItemCode""=T2.""ItemCode"" and T1.""U_Make""=T2.""U_Make"" join ""@BUYERMAKE"" T3 on T1.""U_Make""=T3.""U_Make"""
                'Str += "where  T0.""DocStatus""='O' and T1.""LineStatus""='O' and T3.""U_Buy""='" & objform.Items.Item("BuyName").Specific.String & "' and T0.""DocEntry""='" & SQNum & "' order by T1.""LineNum"""
                If SuperUser Then
                    Str = "Select T0.""DocNum"",T0.""UserSign"",T1.""VisOrder"" as ""LineNum"",TO_VARCHAR(T0.""DocDueDate"" ,'dd/MM/yyyy') as ""DocDueDate"",T0.""CardName"",T4.""GroupName"",T1.""ItemCode"",T1.""Dscription"",T1.""U_Make"" as ""Make"",T1.""U_PrtNo"" as ""MPN"",T1.""U_SPQ"" as ""SPQ"",T1.""U_MOQ"" as ""MOQ"",T1.""Quantity"",T1.""U_CustPrice"","
                    Str += " T1.""U_HMPN"",T1.""U_HDescription"",T1.""U_HMake"",T1.""U_Remarks"" as ""Remarks"",T1.""U_Cost"" as ""Cost"",T1.""U_CostCur"" as ""Currency"",T1.""Price"",T1.""U_LeadTime"" as ""LeadTime"",T1.""U_SupRefNo"" as ""SupRefNo"", TO_VARCHAR(T1.""U_QutValid"",'dd/MM/yyyy') as ""QutValid"","
                    'Str += " Case when T1.""U_QutValid"" <>'' then Left(T1.""U_QutValid"",2)||'/'||Left(Right(T1.""U_QutValid"",4),2)||'/'||Right(T1.""U_QutValid"",2)Else '' END as ""QutValid"","
                    Str += " T1.""U_BuyRem"" as ""BuyerRemarks"",T1.""U_ProjName"" as ""ProjName"",T1.""U_Applictn"" as ""Application"",T1.""U_AnnualVol"" as ""AnnualVolume"",T1.""U_EndCust"" as ""EndCustomer"""
                    Str += " from OQUT T0 Left join QUT1 T1 on T0.""DocEntry""=T1.""DocEntry"" Left join OCRD T2 on T0.""CardCode""=T2.""CardCode"" Left join OCRG T4 on T2.""GroupCode""=T4.""GroupCode"" "
                    Str += " where  T0.""DocStatus""='O' and T1.""LineStatus""='O' "
                    If ComboBox1.Selected.Description = "All" Then
                        Str += ""
                    ElseIf ComboBox1.Selected.Description = "-" Then
                        Str += ""
                    Else
                        Str = ""
                        Str = "Select T0.""DocNum"",T0.""UserSign"",T1.""VisOrder"" as ""LineNum"",TO_VARCHAR(T0.""DocDueDate"" ,'dd/MM/yyyy') as ""DocDueDate"",T0.""CardName"",T4.""GroupName"",T1.""ItemCode"",T1.""Dscription"",T1.""U_Make"" as ""Make"",T1.""U_PrtNo"" as ""MPN"",T1.""U_SPQ"" as ""SPQ"",T1.""U_MOQ"" as ""MOQ"",T1.""Quantity"",T1.""U_CustPrice"","
                        Str += " T1.""U_HMPN"",T1.""U_HDescription"",T1.""U_HMake"",T1.""U_Remarks"" as ""Remarks"",T1.""U_Cost"" as ""Cost"",T1.""U_CostCur"" as ""Currency"",T1.""Price"",T1.""U_LeadTime"" as ""LeadTime"",T1.""U_SupRefNo"" as ""SupRefNo"", TO_VARCHAR(T1.""U_QutValid"",'dd/MM/yyyy') as ""QutValid"","
                        Str += " T1.""U_BuyRem"" as ""BuyerRemarks"",T1.""U_ProjName"" as ""ProjName"",T1.""U_Applictn"" as ""Application"",T1.""U_AnnualVol"" as ""AnnualVolume"",T1.""U_EndCust"" as ""EndCustomer"""
                        Str += " from OQUT T0 Left join QUT1 T1 on T0.""DocEntry""=T1.""DocEntry"" Left join OCRD T2 on T0.""CardCode""=T2.""CardCode"" Left join OCRG T4 on T2.""GroupCode""=T4.""GroupCode"" Left  join ""@BUYERMAKE"" T3 on T1.""U_Make""=T3.""U_Make"" "
                        Str += " where  T0.""DocStatus""='O' and T1.""LineStatus""='O' and T3.""U_BuyName""='" & ComboBox1.Selected.Description & "' "
                    End If
                    If CheckBox1.Checked Then
                        If EditText6.Value <> "" Then
                            Str += " and T1.""U_Make"" in (" & EditText6.Value & ") "
                        End If
                    Else
                        'Not ComboBox2.Selected.Value = "-"
                        If ComboBox0.Selected.Value = "-1" Then
                            Str += " and T1.""U_Make""<>''"
                        ElseIf Not ComboBox0.Selected.Value = "-" Then
                            Str += " and T1.""U_Make""='" & ComboBox0.Selected.Description & "'   "
                        End If
                    End If
                Else
                    Str = "Select T0.""DocNum"",T0.""UserSign"",T1.""VisOrder"" as ""LineNum"",TO_VARCHAR(T0.""DocDueDate"" ,'dd/MM/yyyy') as ""DocDueDate"",T0.""CardName"",T4.""GroupName"",T1.""ItemCode"",T1.""Dscription"",T1.""U_Make"" as ""Make"",T1.""U_PrtNo"" as ""MPN"",T1.""U_SPQ"" as ""SPQ"",T1.""U_MOQ"" as ""MOQ"",T1.""U_PrtNo"" as ""MPN"",T1.""Quantity"",T1.""U_CustPrice"","
                    Str += " T1.""U_HMPN"",T1.""U_HDescription"",T1.""U_HMake"",T1.""U_Remarks"" as ""Remarks"",T1.""U_Cost"" as ""Cost"",T1.""U_CostCur"" as ""Currency"",T1.""Price"",T1.""U_LeadTime"" as ""LeadTime"",T1.""U_SupRefNo"" as ""SupRefNo"",TO_VARCHAR(T1.""U_QutValid"",'dd/MM/yyyy') as ""QutValid"", "
                    ' Str += " Case when T1.""U_QutValid"" <>'' then Left(T1.""U_QutValid"",2)||'/'||Left(Right(T1.""U_QutValid"",4),2)||'/'||Right(T1.""U_QutValid"",2)Else '' END as ""QutValid"","
                    Str += " T1.""U_BuyRem"" as ""BuyerRemarks"",T1.""U_ProjName"" as ""ProjName"",T1.""U_Applictn"" as ""Application"",T1.""U_AnnualVol"" as ""AnnualVolume"",T1.""U_EndCust"" as ""EndCustomer"""
                    Str += " from OQUT T0 Left join QUT1 T1 on T0.""DocEntry""=T1.""DocEntry"" Left join OCRD T2 on T0.""CardCode""=T2.""CardCode"" Left join OCRG T4 on T2.""GroupCode""=T4.""GroupCode"" join ""@BUYERMAKE"" T3 on T1.""U_Make""=T3.""U_Make"""
                    Str += " where  T0.""DocStatus""='O' and T1.""LineStatus""='O' "
                    If CheckBox1.Checked Then
                        If EditText6.Value <> "" Then
                            For i As Integer = 2 To ComboBox0.ValidValues.Count - 1
                                If i = 2 Then
                                    Make = "'" + ComboBox0.ValidValues.Item(i).Description + "'"
                                Else
                                    Make = Make + ",'" + ComboBox0.ValidValues.Item(i).Description + "'"
                                End If
                            Next
                            If Make <> EditText6.Value Then
                                objaddon.objapplication.StatusBar.SetText("Unassigned Make found for this user.Please update the assigned Make...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            Else
                                Str += " and T1.""U_Make"" in (" & EditText6.Value & ") "
                            End If

                        End If
                    Else
                        If ComboBox0.Selected.Value = "-1" Then
                            For i As Integer = 2 To ComboBox0.ValidValues.Count - 1
                                If i = 2 Then
                                    Make = "'" + ComboBox0.ValidValues.Item(i).Description + "'"
                                Else
                                    Make = Make + ",'" + ComboBox0.ValidValues.Item(i).Description + "'"
                                End If
                            Next
                            Str += " and T1.""U_Make""in (" & Make & ")"
                        ElseIf Not ComboBox0.Selected.Value = "-" Then
                            Str += " and T1.""U_Make""='" & ComboBox0.Selected.Description & "'  "
                        End If
                    End If
                    
                    If ComboBox1.Selected.Description = "All" Then
                        Str += ""
                    ElseIf ComboBox1.Selected.Description = "-" Then
                        Str += ""
                    Else
                        Str += " and T3.""U_BuyName""='" & ComboBox1.Selected.Description & "' "
                    End If
                End If
                If SQNum <> "" Then
                    Str += " and T0.""DocEntry""='" & SQNum & "' "
                End If
                If objform.Items.Item("CustN").Specific.String <> "" Then
                    Str += " and T0.""CardCode""='" & objform.Items.Item("CustN").Specific.String & "'  "
                End If
                If objform.Items.Item("DueDate").Specific.String <> "" Then
                    Dim DocDate As Date = Date.ParseExact(EditText5.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Str += " and T0.""DocDueDate""<='" & DocDate.ToString("yyyyMMdd") & "' "
                End If
                If EditText0.Value <> "" Or EditText3.Value <> "" Then
                    If EditText0.Value = "" Or EditText3.Value = "" Then
                        objaddon.objapplication.StatusBar.SetText("Please Select FromDate & ToDate to get the SQ data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Dim ToDate As Date = Date.ParseExact(EditText3.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Str += " and T0.""DocDate"" Between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' "
                End If
                If Not ComboBox2.Selected.Value = "-" Then
                    Str += " and T0.""U_SQCategory""='" & ComboBox2.Selected.Value & "' "
                End If
                If CheckBox2.Checked Then
                    If EditText7.Value <> "" Then
                        Str += " and T1.""U_PrtNo"" in (" & EditText7.Value & ") "
                    End If
                Else
                    If ComboBox3.Selected.Value = "-1" Then
                        Str += " and T1.""U_PrtNo""<>''"
                    ElseIf Not ComboBox3.Selected.Value = "-" Then
                        Str += " and T1.""U_PrtNo""='" & ComboBox3.Selected.Description & "'   "
                    End If
                End If
                Str += "  order by T0.""DocDueDate"" "
                objRecordset.DoQuery(Str)
                'objform.Freeze(True)
                'odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object))
                'odbdsDetails.Clear()
                'Matrix0.LoadFromDataSource()
                'Dim Validity As Date
                If objRecordset.RecordCount = 0 Then objaddon.objapplication.SetStatusBarMessage("No records Found", SAPbouiCOM.BoMessageTime.bmt_Short, True) : objform.Freeze(False) : Exit Sub
                odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object))
                odbdsDetails.Clear()
                Matrix0.Clear()
                odbdsDetails.InsertRecord(odbdsDetails.Size)
                objaddon.objapplication.SetStatusBarMessage("SalesQuotation Loading Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                For i As Integer = 0 To objRecordset.RecordCount - 1
                    'Dim DocDueDate As Date = Date.ParseExact(objRecordset.Fields.Item("DocDueDate").Value.ToString, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    'If objRecordset.Fields.Item("QutValid").Value.ToString <> "" Or IsDBNull(objRecordset.Fields.Item("QutValid").Value.ToString) Then
                    '    Validity = Date.ParseExact(objRecordset.Fields.Item("QutValid").Value.ToString, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    'Else
                    '    Validity = ""
                    'End If
                    odbdsDetails.SetValue("LineId", i, (i + 1))
                    odbdsDetails.SetValue("U_SQNum", i, objRecordset.Fields.Item("DocNum").Value)
                    odbdsDetails.SetValue("U_SQLine", i, objRecordset.Fields.Item("LineNum").Value)
                    odbdsDetails.SetValue("U_DocDueDate", i, objRecordset.Fields.Item("DocDueDate").Value)
                    odbdsDetails.SetValue("U_CustName", i, objRecordset.Fields.Item("CardName").Value)
                    odbdsDetails.SetValue("U_CustType", i, objRecordset.Fields.Item("GroupName").Value)
                    odbdsDetails.SetValue("U_Itemcode", i, objRecordset.Fields.Item("ItemCode").Value)
                    odbdsDetails.SetValue("U_ItemDesc", i, objRecordset.Fields.Item("Dscription").Value)
                    odbdsDetails.SetValue("U_Make", i, objRecordset.Fields.Item("Make").Value)
                    odbdsDetails.SetValue("U_MPN", i, objRecordset.Fields.Item("MPN").Value)
                    odbdsDetails.SetValue("U_SPQ", i, objRecordset.Fields.Item("SPQ").Value)
                    odbdsDetails.SetValue("U_MOQ", i, objRecordset.Fields.Item("MOQ").Value)
                    odbdsDetails.SetValue("U_Quantity", i, objRecordset.Fields.Item("Quantity").Value)
                    odbdsDetails.SetValue("U_TPrice", i, objRecordset.Fields.Item("U_CustPrice").Value)
                    odbdsDetails.SetValue("U_HicoMPN", i, objRecordset.Fields.Item("U_HMPN").Value)
                    odbdsDetails.SetValue("U_HicoDesc", i, objRecordset.Fields.Item("U_HDescription").Value)
                    odbdsDetails.SetValue("U_HicoMake", i, objRecordset.Fields.Item("U_HMake").Value)
                    odbdsDetails.SetValue("U_Remarks", i, objRecordset.Fields.Item("Remarks").Value)
                    odbdsDetails.SetValue("U_Cost", i, objRecordset.Fields.Item("Cost").Value)
                    odbdsDetails.SetValue("U_Currency", i, objRecordset.Fields.Item("Currency").Value)
                    odbdsDetails.SetValue("U_UnitPrice", i, objRecordset.Fields.Item("Price").Value)
                    odbdsDetails.SetValue("U_LeadTime", i, objRecordset.Fields.Item("LeadTime").Value)
                    odbdsDetails.SetValue("U_SupRefNo", i, objRecordset.Fields.Item("SupRefNo").Value)
                    'If IsDate(Validity) Then
                    '    odbdsDetails.SetValue("U_QutValid", i, Validity.ToString("dd/MM/yy")) 'objRecordset.Fields.Item("QutValid").Value)
                    'End If
                    odbdsDetails.SetValue("U_QutValid", i, objRecordset.Fields.Item("QutValid").Value) 'objRecordset.Fields.Item("QutValid").Value)
                    odbdsDetails.SetValue("U_BuyRem", i, objRecordset.Fields.Item("BuyerRemarks").Value)
                    odbdsDetails.SetValue("U_ProjName", i, objRecordset.Fields.Item("ProjName").Value)
                    odbdsDetails.SetValue("U_Applictn", i, objRecordset.Fields.Item("Application").Value)
                    odbdsDetails.SetValue("U_AnnualVol", i, objRecordset.Fields.Item("AnnualVolume").Value)
                    odbdsDetails.SetValue("U_EndCust", i, objRecordset.Fields.Item("EndCustomer").Value)
                    odbdsDetails.SetValue("U_CUser", i, objRecordset.Fields.Item("UserSign").Value)
                    objRecordset.MoveNext()
                    If i <> objRecordset.RecordCount - 1 Then odbdsDetails.InsertRecord(odbdsDetails.Size)
                Next
                Matrix0.LoadFromDataSource()

                'Dim Row As Integer = 0
                'For i As Integer = 0 To objRecordset.RecordCount - 1
                '    Matrix0.AddRow()
                '    Row += 1
                '    Matrix0.Columns.Item("#").Cells.Item(Row).Specific.String = i + 1
                '    Matrix0.Columns.Item("SQNum").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("DocNum").Value
                '    Matrix0.Columns.Item("SQLine").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("LineNum").Value
                '    Matrix0.Columns.Item("SQDue").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("DocDueDate").Value
                '    Matrix0.Columns.Item("CustName").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("CardName").Value
                '    Matrix0.Columns.Item("CustType").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("GroupName").Value
                '    Matrix0.Columns.Item("ItemCode").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("ItemCode").Value
                '    Matrix0.Columns.Item("ItemDesc").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Dscription").Value
                '    Matrix0.Columns.Item("Make").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Make").Value
                '    Matrix0.Columns.Item("MPN").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("MPN").Value
                '    Matrix0.Columns.Item("SPQ").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("SPQ").Value
                '    Matrix0.Columns.Item("MOQ").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("MOQ").Value
                '    Matrix0.Columns.Item("Quant").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Quantity").Value
                '    Matrix0.Columns.Item("CustTP").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("U_CustPrice").Value
                '    Matrix0.Columns.Item("HMPN").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("U_HMPN").Value
                '    Matrix0.Columns.Item("HDesc").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("U_HDescription").Value
                '    Matrix0.Columns.Item("HMake").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("U_HMake").Value
                '    Matrix0.Columns.Item("Remarks").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Remarks").Value
                '    Matrix0.Columns.Item("Cost").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Cost").Value
                '    If objRecordset.Fields.Item("Currency").Value <> "" Then
                '        objCurrency = Matrix0.Columns.Item("Curr").Cells.Item(Row).Specific
                '        Dim Currency As String = objRecordset.Fields.Item("Currency").Value.ToString
                '        objCurrency.Select(Currency, SAPbouiCOM.BoSearchKey.psk_ByValue)
                '    End If
                '    Matrix0.Columns.Item("UnitPr").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Price").Value
                '    Matrix0.Columns.Item("LeadT").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("LeadTime").Value
                '    Matrix0.Columns.Item("SupRef").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("SupRefNo").Value
                '    Matrix0.Columns.Item("Valid").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("QutValid").Value
                '    Matrix0.Columns.Item("BuyRem").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("BuyerRemarks").Value
                '    Matrix0.Columns.Item("ProjName").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("ProjName").Value
                '    Matrix0.Columns.Item("Applicat").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("Application").Value
                '    Matrix0.Columns.Item("AnnualVol").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("AnnualVolume").Value
                '    Matrix0.Columns.Item("EndCust").Cells.Item(Row).Specific.String = objRecordset.Fields.Item("EndCustomer").Value
                '    objRecordset.MoveNext()
                'Next
                Matrix0.CommonSetting.FixedColumnsCount = 7
                objaddon.objapplication.Menus.Item("1300").Activate()
                'For i As Integer = 1 To Matrix0.Columns.Count - 1
                '    Matrix0.Columns.Item(i).TitleObject.Sortable = True
                'Next
                'objform.Freeze(False)
                Dim GetDummyItem As String
                GetDummyItem = objaddon.objglobalmethods.getSingleValue("select Top 1 ""U_ItemCode"" from ""@BUYERMAKE"" where ""U_ItemCode""<>'';")
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("ItemCode").Cells.Item(i).Specific.String = GetDummyItem Then
                        Matrix0.CommonSetting.SetCellEditable(i, 7, True)
                        Matrix0.CommonSetting.SetCellEditable(i, 11, True)
                        Matrix0.CommonSetting.SetCellEditable(i, 12, True)
                    Else
                        Matrix0.CommonSetting.SetCellEditable(i, 7, False)
                        Matrix0.CommonSetting.SetCellEditable(i, 11, False)
                        Matrix0.CommonSetting.SetCellEditable(i, 12, False)
                    End If
                Next
                Matrix0.Columns.Item("#").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objaddon.objapplication.StatusBar.SetText("SalesQuotation Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objform.Freeze(False)
                GC.Collect()
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub UpdateSalesQuotation()
           
            Dim objSalesQuotation As SAPbobsCOM.Documents
            Dim objrs As SAPbobsCOM.Recordset
            Dim SQNum As String, SQEntry As String, HSNCode As String
            'Dim LineNum As String
            Dim Success As Boolean = False
            Try
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Str As String = ""
                Dim objCurrency As SAPbouiCOM.ComboBox
                Dim objCombo As SAPbouiCOM.ComboBox
                'Dim objSQData As SAPbouiCOM.DBDataSource
                objCombo = objform.Items.Item("BuyName").Specific
                Dim Valid As String
                objSalesQuotation = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
                'objSQData = objform.DataSources.DBDataSources.Item("OQUT")
                If Matrix0.VisualRowCount > 0 Then
                    'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    objaddon.objapplication.SetStatusBarMessage("SalesQuotation Updating Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    For RowNo As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("Select").Cells.Item(RowNo).Specific.Checked = True Then
                            SQNum = Trim(Matrix0.Columns.Item("SQNum").Cells.Item(RowNo).Specific.String)
                            Dim DocDueDate As Date = Date.ParseExact(Matrix0.Columns.Item("SQDue").Cells.Item(RowNo).Specific.String, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            Str = "select ""DocEntry"",""DocDate"",""Series"" from OQUT where ""DocNum""='" & SQNum & "' and  ""DocStatus""='O' and ""DocDueDate""='" & DocDueDate.ToString("yyyyMMdd") & "'"
                            objrs.DoQuery(Str)
                            SQEntry = Trim(objrs.Fields.Item("DocEntry").Value.ToString)

                            If objSalesQuotation.GetByKey(SQEntry) Then
                                'LineNum = objSalesQuotation.Lines.VisualOrder 'Trim(Matrix0.Columns.Item("SQLine").Cells.Item(RowNo).Specific.String)
                                'objSalesQuotation.Series = objrs.Fields.Item("Series").Value
                                objSalesQuotation.Lines.SetCurrentLine(Trim(Matrix0.Columns.Item("SQLine").Cells.Item(RowNo).Specific.String))
                                'TaxCode = objaddon.objglobalmethods.getSingleValue("Select T1.""TaxCode"" from OQUT T0 join QUT1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocNum""=" & Trim(Matrix0.Columns.Item("SQNum").Cells.Item(RowNo).Specific.String) & " and T1.""VisOrder""='" & Trim(Matrix0.Columns.Item("SQLine").Cells.Item(RowNo).Specific.String) & "'")
                                objSalesQuotation.Lines.ItemCode = Matrix0.Columns.Item("ItemCode").Cells.Item(RowNo).Specific.String 'objSQData.GetValue("ItemCode", 0)
                                'objSalesQuotation.Lines.ItemDescription = Matrix0.Columns.Item("ItemDesc").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.Quantity = Matrix0.Columns.Item("Quant").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.TaxCode = objSalesQuotation.Lines.TaxCode
                                objSalesQuotation.Lines.UnitPrice = Matrix0.Columns.Item("UnitPr").Cells.Item(RowNo).Specific.String
                                HSNCode = objaddon.objglobalmethods.getSingleValue("select ""ChapterID"" from OITM where ""ItemCode""='" & Matrix0.Columns.Item("ItemCode").Cells.Item(RowNo).Specific.String & "';")
                                If HSNCode <> "" Then
                                    objSalesQuotation.Lines.HSNEntry = HSNCode
                                End If
                                'objSalesQuotation.Lines.LineTotal = CDbl(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String) * CDbl(Matrix0.Columns.Item("UnitPr").Cells.Item(i).Specific.String)
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_CustPrice").Value = Matrix0.Columns.Item("CustTP").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_HMPN").Value = Matrix0.Columns.Item("HMPN").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_HDescription").Value = Matrix0.Columns.Item("HDesc").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_HMake").Value = Matrix0.Columns.Item("HMake").Cells.Item(RowNo).Specific.String
                                'objSalesQuotation.Lines.UserFields.Fields.Item("U_LeadTime").Value = Matrix0.Columns.Item("LeadT").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_SPQ").Value = Matrix0.Columns.Item("SPQ").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_MOQ").Value = Matrix0.Columns.Item("MOQ").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_Remarks").Value = Matrix0.Columns.Item("Remarks").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_Cost").Value = Matrix0.Columns.Item("Cost").Cells.Item(RowNo).Specific.String
                                objCurrency = Matrix0.Columns.Item("Curr").Cells.Item(RowNo).Specific
                                If Not objCurrency.Selected Is Nothing Then
                                    objSalesQuotation.Lines.UserFields.Fields.Item("U_CostCur").Value = objCurrency.Selected.Value 'Matrix0.Columns.Item("Curr").Cells.Item(RowNo).Specific.String
                                End If
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_LeadTime").Value = Matrix0.Columns.Item("LeadT").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_SupRefNo").Value = Matrix0.Columns.Item("SupRef").Cells.Item(RowNo).Specific.String
                                Valid = Matrix0.Columns.Item("Valid").Cells.Item(RowNo).Specific.String
                                Valid = Valid.Replace("/", "")
                                Valid = Right(Valid, 4) & "/" & Left(Right(Valid, 6), 2) & "/" & Left(Valid, 2)
                                If Matrix0.Columns.Item("Valid").Cells.Item(RowNo).Specific.String <> "" Then
                                    'Valid = Date.ParseExact(Matrix0.Columns.Item("Valid").Cells.Item(RowNo).Specific.String, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    objSalesQuotation.Lines.UserFields.Fields.Item("U_QutValid").Value = Valid   'Matrix0.Columns.Item("Valid").Cells.Item(RowNo).Specific.ToString("dd/MM/yy") 
                                End If

                                'objSalesQuotation.Lines.UserFields.Fields.Item("U_PrtNo").Value = Matrix0.Columns.Item("MPN").Cells.Item(RowNo).Specific.String
                                'objSalesQuotation.Lines.UserFields.Fields.Item("U_Make").Value = Matrix0.Columns.Item("Make").Cells.Item(RowNo).Specific.String
                                ' objSalesQuotation.Lines.UserFields.Fields.Item("U_UnitPrice").Value = Matrix0.Columns.Item("UnitPr").Cells.Item(i).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_BuyRem").Value = Matrix0.Columns.Item("BuyRem").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_ProjName").Value = Matrix0.Columns.Item("ProjName").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_Applictn").Value = Matrix0.Columns.Item("Applicat").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_AnnualVol").Value = Matrix0.Columns.Item("AnnualVol").Cells.Item(RowNo).Specific.String
                                objSalesQuotation.Lines.UserFields.Fields.Item("U_EndCust").Value = Matrix0.Columns.Item("EndCust").Cells.Item(RowNo).Specific.String

                                If SuperUser Then
                                    objSalesQuotation.Lines.UserFields.Fields.Item("U_BuyName").Value = objaddon.objglobalmethods.getSingleValue(" select ""U_NAME"" from OUSR where ""USER_CODE""='" & objaddon.objapplication.Company.UserName & "'") 'objCombo.Selected.Description
                                Else
                                    objSalesQuotation.Lines.UserFields.Fields.Item("U_BuyName").Value = objCombo.Selected.Description 'objform.Items.Item("BuyName").Specific.String
                                End If

                                'objSalesQuotation.Lines.Add()
                                If objSalesQuotation.Update() <> 0 Then
                                    'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    objaddon.objapplication.StatusBar.SetText(objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    Success = True
                                End If
                            Else
                                'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objaddon.objapplication.StatusBar.SetText("Sales Quotation Not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    Next
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(objSalesQuotation)
                GC.Collect()

                Try
                    Dim objChkbox As SAPbouiCOM.CheckBox
                    Dim DocEntry, UserName As String
                    Dim DocDate As Date
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        objChkbox = Matrix0.Columns.Item("Select").Cells.Item(i).Specific
                        If objChkbox.Checked = True Then
                            DocDate = Date.ParseExact(Matrix0.Columns.Item("SQDue").Cells.Item(i).Specific.String, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            DocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from OQUT where ""DocNum""='" & Matrix0.Columns.Item("SQNum").Cells.Item(i).Specific.String & "' and ""DocDueDate""='" & DocDate.ToString("yyyyMMdd") & "'")
                            UserName = objaddon.objglobalmethods.getSingleValue("select ""USER_CODE"" from OUSR where ""USERID""='" & Matrix0.Columns.Item("CUser").Cells.Item(i).Specific.String & "'")
                            SendAlertMessageToUsers(UserName, DocEntry, i)
                        End If
                    Next
                Catch ex As Exception
                End Try
                If Success Then
                    Button0.Caption = "OK"
                    objaddon.objapplication.StatusBar.SetText("SalesQuotation Updated & Alert has been sent Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Button0.Caption = "Update"
                End If

            Catch ex As Exception
                ' If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(objaddon.objcompany)
                GC.Collect()
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub Alert()
            Try
                Dim oAlertTemplatesService As AlertManagementService
                oAlertTemplatesService = objaddon.objcompany.GetBusinessService(ServiceTypes.AlertManagementService)
                Dim oAlertTemplate As AlertManagement
                Dim oAlertTemplateParams As AlertManagementParams
                Dim oAlertTemplateRecipients As AlertManagementRecipients
                Dim oAlertRecipient As AlertManagementRecipient
                Dim j As Integer
                Dim Doc, Prior, freq As String
                Doc = AlertManagementDocumentEnum.atd_Quotations
                Prior = AlertManagementPriorityEnum.atp_High
                freq = AlertManagementFrequencyType.atfi_Minutes
                oAlertTemplateParams = oAlertTemplatesService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagement)
                'set system alert code
                oAlertTemplateParams.Code = 1 'txtCode.Text
                'get alert template
                oAlertTemplate = oAlertTemplatesService.GetAlertManagement(oAlertTemplateParams)
                'set alert name
                oAlertTemplate.Name = "Quotation Updation"
                oAlertTemplate.FrequencyType = AlertManagementFrequencyType.atfi_Minutes
                oAlertTemplate.Active = BoYesNoEnum.tYES ' IIf(chkActive.Checked, BoYesNoEnum.tYES, BoYesNoEnum.tNO)
                'set % Discount
                oAlertTemplate.Param = "" 'txtConditions.Text
                'set priority
                oAlertTemplate.Priority = AlertManagementPriorityEnum.atp_Normal ' GetPriority(cboPriority.SelectedItem)
                'update selected document
                For j = 0 To oAlertTemplate.AlertManagementDocuments.Count - 1
                    If oAlertTemplate.AlertManagementDocuments.Item(j).Document = AlertManagementDocumentEnum.atd_Quotations Then ' GetDocumentType(cboDocument.SelectedItem) Then
                        oAlertTemplate.AlertManagementDocuments.Item(j).Active = BoYesNoEnum.tYES
                        Exit For
                    End If
                Next j
                'get recipient collection
                oAlertTemplateRecipients = oAlertTemplate.AlertManagementRecipients
                'add recipient
                oAlertRecipient = oAlertTemplateRecipients.Add()
                'set recipient code(manager=1)
                oAlertRecipient.UserCode = 8
                'set internal message
                oAlertRecipient.SendInternal = BoYesNoEnum.tYES 'IIf(chkSendInternal.Checked, BoYesNoEnum.tYES, BoYesNoEnum.tNO)
                'update alert
                oAlertTemplatesService.UpdateAlertManagement(oAlertTemplate)
            Catch ex As Exception

                MsgBox(ex.Message)

            End Try
        End Sub

        'Private Sub AddAlert(ByVal Usercode As Integer)

        '    Try
        '        Dim oAlertTemplatesService As AlertManagementService
        '        oAlertTemplatesService = objaddon.objcompany.Getd
        '        Dim oAlertTemplate As AlertManagement
        '        Dim oAlertTemplateParams As AlertManagementParams
        '        Dim oAlertTemplateRecipients As AlertManagementRecipients
        '        Dim oAlertRecipient As AlertManagementRecipient
        '        Dim Doc, Prior, freq As String
        '        Doc = AlertManagementDocumentEnum.atd_Quotations
        '        Prior = AlertManagementPriorityEnum.atp_High
        '        freq = AlertManagementFrequencyType.atfi_Minutes
        '        'get alert template
        '        oAlertTemplate = oAlertTemplatesService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagement)
        '        'set alert name
        '        oAlertTemplate.Name = "Test" 'txtName.Text
        '        'set query
        '        'oAlertTemplate.QueryID = "" 'txtQuery.Text
        '        oAlertTemplate.Active = BoYesNoEnum.tYES 'IIf(chkActive.Checked, BoYesNoEnum.tYES, BoYesNoEnum.tNO)
        '        'set priority
        '        oAlertTemplate.Priority = AlertManagementPriorityEnum.atp_High ' GetPriority(cboPriority.SelectedItem)
        '        'set the FrequencyType (minutes,hours...)
        '        oAlertTemplate.FrequencyType = AlertManagementFrequencyType.atfi_Minutes 'GetFrequencyType(cboFreqType.SelectedItem)
        '        'set intervals
        '        oAlertTemplate.FrequencyInterval = 0 'txtFreqIntrvls.Text
        '        'get Recipients collection 
        '        oAlertTemplateRecipients = oAlertTemplate.AlertManagementRecipients
        '        'add recipient
        '        oAlertRecipient = oAlertTemplateRecipients.Add()
        '        'set recipient code(manager=1)
        '        oAlertRecipient.UserCode = Usercode ' txtUserCode.Text
        '        'set internal message
        '        oAlertRecipient.SendInternal = BoYesNoEnum.tYES 'IIf(chkSendInternal.Checked, BoYesNoEnum.tYES, BoYesNoEnum.tNO)
        '        'add alert template
        '        oAlertTemplateParams = oAlertTemplatesService.AddAlertManagement(oAlertTemplate)
        '    Catch ex As Exception

        '        MsgBox(ex.Message)

        '    End Try

        'End Sub

        Private Sub AddAlertNew()

            Try
                Dim oAlertTemplatesService As AlertManagementService
                Dim oAlertTemplate As AlertManagement
                Dim oAlertTemplateParams As AlertManagementParams
                Dim oAlertTemplateRecipients As AlertManagementRecipients
                Dim oAlertRecipient As AlertManagementRecipient
                oAlertTemplatesService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.AlertManagementService)
                Dim Doc, Prior, freq As String
                Doc = AlertManagementDocumentEnum.atd_Quotations
                Prior = AlertManagementPriorityEnum.atp_High
                freq = AlertManagementFrequencyType.atfi_Minutes
                'get alert template
                oAlertTemplate = oAlertTemplatesService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagement)
                'set alert name
                oAlertTemplate.Name = "Test"
                'set query
                'oAlertTemplate.QueryID = ""
                oAlertTemplate.Active = BoYesNoEnum.tYES
                'set priority
                oAlertTemplate.Priority = Prior
                'set the FrequencyType (minutes,hours...)
                oAlertTemplate.FrequencyType = freq
                'set intervals
                oAlertTemplate.FrequencyInterval = 0
                'get Recipients collection 
                oAlertTemplateRecipients = oAlertTemplate.AlertManagementRecipients
                'add recipient
                oAlertRecipient = oAlertTemplateRecipients.Add()
                'set recipient code(manager=1)
                oAlertRecipient.UserCode = 8
                'set internal message
                'oAlertRecipient.SendInternal = BoYesNoEnum.tYES
                'add alert template
                oAlertTemplateParams = oAlertTemplatesService.AddAlertManagement(oAlertTemplate)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End Sub

        Private Sub AlertNew23102020()
            Try
                Dim oAlertTemplatesService As AlertManagementService
                Dim oAlertTemplate As AlertManagement
                Dim oAlertTemplateParams As AlertManagementParams
                Dim oAlertTemplateRecipients As AlertManagementRecipients
                Dim oAlertRecipient As AlertManagementRecipient
                Dim j As Integer
                oAlertTemplatesService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.AlertManagementService)
                'get alert template params
                oAlertTemplateParams = oAlertTemplatesService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagementParams)

                'set system alert code
                oAlertTemplateParams.Code = 2

                'get alert template
                oAlertTemplate = oAlertTemplatesService.GetAlertManagement(oAlertTemplateParams)

                'set alert name
                oAlertTemplate.Name = "Testing"
                oAlertTemplate.QueryID = 305
                oAlertTemplate.Active = BoYesNoEnum.tYES

                'set % Discount
                'oAlertTemplate.Param = txtConditions.Text

                'set priority
                oAlertTemplate.Priority = AlertManagementPriorityEnum.atp_Normal

                'update selected document
                For j = 0 To oAlertTemplate.AlertManagementDocuments.Count - 1
                    If oAlertTemplate.AlertManagementDocuments.Item(j).Document = AlertManagementDocumentEnum.atd_Quotations Then
                        oAlertTemplate.AlertManagementDocuments.Item(j).Active = BoYesNoEnum.tYES
                        Exit For
                    End If
                Next j

                'get recipient collection
                oAlertTemplateRecipients = oAlertTemplate.AlertManagementRecipients

                'add recipient
                oAlertRecipient = oAlertTemplateRecipients.Add()

                'set recipient code(manager=1)
                oAlertRecipient.UserCode = 8

                'set internal message
                oAlertRecipient.SendInternal = BoYesNoEnum.tYES
                'update alert
                'oAlertTemplatesService.UpdateAlertManagement(oAlertTemplate)

            Catch ex As Exception

                MsgBox(ex.Message)

            End Try
        End Sub

        Private Sub testalert()
            Dim msg As SAPbobsCOM.Messages
            msg = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oMessages)
            'msg.MessageText = "This is the content of message"
            msg.Subject = "Hello manager"
            'there are two recipients in this message
            Call msg.Recipients.Add()
            'set values for the first recipients
            'Call msg.Recipients.SetCurentLine(0)
            msg.Recipients.UserCode = "srmprof1"
            msg.Recipients.NameTo = "srmprof1"
            'msg.Recipients.SendEmail = BoYesNoEnum.tYES
            'msg.Recipients.EmailAddress = "manager.Test@hotmail.com"
            msg.Recipients.SendInternal = BoYesNoEnum.tYES
            For i As Integer = 0 To msg.Recipients.Count

            Next
            Call msg.AddDataColumn("Please review", "This is an invoice, pls review it, thx.", BoObjectTypes.oQuotations, "39")
            'add attachment
            'Call msg.Attachments.Add()
            'msg.Attachments.Item(0).FileName = "C:\temp\Object.xml"
            'send the message
            Call msg.Add()
            Dim strCode As String
            Call objaddon.objcompany.GetNewObjectCode(strCode)
            MsgBox("New Object Code = " + strCode)
            'Check Error
            Dim nErr As Long
            Dim errMsg As String
            Call objaddon.objcompany.GetLastError(nErr, errMsg)
            If (0 <> nErr) Then
                MsgBox("Found error:" + errMsg)
            End If
        End Sub

        Private Sub testAlertMessage()
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As MessagesService
            Dim oMessage As Message
            Dim pMessageDataColumns As MessageDataColumns
            Dim pMessageDataColumn As MessageDataColumn
            Dim oLines As MessageDataLines
            Dim oLine As MessageDataLine
            Dim oRecipientCollection As RecipientCollection
            Try

                'get company service
                oCmpSrv = objaddon.objcompany.GetCompanyService

                'get msg service
                oMessageService = oCmpSrv.GetBusinessService(ServiceTypes.MessagesService)

                'get the data interface for the new message
                oMessage = oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage)

                'fill subject
                oMessage.Subject = "My Subject"

                'fill text
                oMessage.Text = "My Text"

                'Add Recipient
                oRecipientCollection = oMessage.RecipientCollection

                'Add new a recipient
                oRecipientCollection.Add()

                'send internal message
                oRecipientCollection.Item(0).SendInternal = BoYesNoEnum.tYES

                'add existing user code
                oRecipientCollection.Item(0).UserCode = "srmprof1"

                'get columns data
                pMessageDataColumns = oMessage.MessageDataColumns
                'get column
                pMessageDataColumn = pMessageDataColumns.Add()
                'set column name
                pMessageDataColumn.ColumnName = "My Column Name"
                'get lines
                oLines = pMessageDataColumn.MessageDataLines()
                'add new line
                oLine = oLines.Add()
                'set the line value
                oLine.Value = "My Value"
                'send the message
                oMessageService.SendMessage(oMessage)

            Catch ex As Exception

            End Try
        End Sub

        Private Sub SendAlertMessageToUsers(ByVal UserName As String, ByVal DocEntry As String, ByVal RowId As Integer)
            Dim oMessage As Message
            Dim pMessageDataColumns As MessageDataColumns
            Dim pMessageDataColumn As MessageDataColumn
            Dim oLines As MessageDataLines
            Dim oLine As MessageDataLine
            Dim oRecipientCollection As RecipientCollection
            Try
                Dim Row As Integer = 0
                Dim oMessageService As MessagesService
                oMessageService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.MessagesService)
                oMessage = oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage)
                oMessage.Subject = "Sales Quotation Updated"
                oRecipientCollection = oMessage.RecipientCollection
                oRecipientCollection.Add()
                'Try
                '    Dim objChkbox As SAPbouiCOM.CheckBox
                '    Dim DocEntry As String
                '    Dim DocDate As Date
                '    For MatrixLine As Integer = 1 To Matrix0.VisualRowCount
                '        objChkbox = Matrix0.Columns.Item("Select").Cells.Item(MatrixLine).Specific
                '        If objChkbox.Checked = True Then
                '            DocDate = Date.ParseExact(Matrix0.Columns.Item("SQDue").Cells.Item(MatrixLine).Specific.String, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                '            DocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from OQUT where ""DocNum""='" & Matrix0.Columns.Item("SQNum").Cells.Item(MatrixLine).Specific.String & "' and ""DocDueDate""='" & DocDate.ToString("yyyyMMdd") & "'")
                '            oRecipientCollection.Item(0).SendInternal = BoYesNoEnum.tYES
                '            oRecipientCollection.Item(0).UserCode = "srmprof1"

                '            pMessageDataColumns = oMessage.MessageDataColumns

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "DocEntry"
                '            pMessageDataColumn.Link = BoYesNoEnum.tYES
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            'set the line value
                '            oLine.Value = DocEntry
                '            oLine.Object = "23"  'Object Type
                '            oLine.ObjectKey = DocEntry

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "DocNum"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            'MsgBox(Matrix0.Columns.Item("SQNum").Cells.Item(RowId).Specific.String)
                '            oLine.Value = Matrix0.Columns.Item("SQNum").Cells.Item(MatrixLine).Specific.String

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "LineId"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            oLine.Value = Matrix0.Columns.Item("SQLine").Cells.Item(MatrixLine).Specific.String

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "Customer Name"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            oLine.Value = Matrix0.Columns.Item("CustName").Cells.Item(MatrixLine).Specific.String

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "ItemCode"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            oLine.Value = Matrix0.Columns.Item("ItemCode").Cells.Item(MatrixLine).Specific.String

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "ItemName"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            oLine.Value = Matrix0.Columns.Item("ItemDesc").Cells.Item(MatrixLine).Specific.String

                '            pMessageDataColumn = pMessageDataColumns.Add()
                '            pMessageDataColumn.ColumnName = "UserName"
                '            oLines = pMessageDataColumn.MessageDataLines()
                '            oLine = oLines.Add()
                '            oLine.Value = objaddon.objcompany.UserName

                '            oMessage.MessageDataColumns.Item(pMessageDataColumns.Add)
                '        End If
                '    Next
                '    oMessageService.SendMessage(oMessage)
                'Catch ex As Exception
                'End Try
                For i As Integer = 0 To oRecipientCollection.Count - 1
                    oRecipientCollection.Item(i).SendInternal = BoYesNoEnum.tYES
                    oRecipientCollection.Item(i).UserCode = UserName '"srmprof1"
                    pMessageDataColumns = oMessage.MessageDataColumns
                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "DocEntry"
                    pMessageDataColumn.Link = BoYesNoEnum.tYES
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    'set the line value
                    oLine.Value = DocEntry
                    oLine.Object = "23"  'Object Type
                    oLine.ObjectKey = DocEntry

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "DocNum"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = Matrix0.Columns.Item("SQNum").Cells.Item(RowId).Specific.String

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "LineId"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = Matrix0.Columns.Item("SQLine").Cells.Item(RowId).Specific.String

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "Customer Name"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = Matrix0.Columns.Item("CustName").Cells.Item(RowId).Specific.String

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "ItemCode"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = Matrix0.Columns.Item("ItemCode").Cells.Item(RowId).Specific.String

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "ItemName"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = Matrix0.Columns.Item("ItemDesc").Cells.Item(RowId).Specific.String

                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "UserName"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = objaddon.objcompany.UserName
                    'send the message
                    oMessageService.SendMessage(oMessage)
                Next

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                UpdateSalesQuotation()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Button1_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button1.ClickBefore

            Try
                objform.Freeze(True)
                EditText1.Value = ""
                EditText4.Value = ""
                EditText5.Value = ""
                EditText2.Value = ""
                EditText0.Value = ""
                EditText3.Value = ""
                EditText6.Value = ""
                EditText7.Value = ""
                If SuperUser Then
                    ComboBox1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
                ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                ComboBox2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                ComboBox3.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Button0.Caption = "Update"
                Matrix0.Clear()
                objform.ActiveItem = "txtentry"
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub EditText1_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.KeyDownAfter
            'Try
            '    If objform.Items.Item("txtentry").Specific.String <> "" Then
            '        objform.Freeze(True)
            '        LoadSQScreen(objform.Items.Item("DocEntry").Specific.String)
            '        objform.Freeze(False)
            '    End If
            'Catch ex As Exception

            'End Try

        End Sub

        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText


        Private Sub EditText4_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText4.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CustCode")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "C"
                ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND        

                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText4_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText4.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    Catch ex As Exception
                        EditText4.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    End Try
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private WithEvents Button2 As SAPbouiCOM.Button


        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                LoadSQScreen(objform.Items.Item("DocEntry").Specific.String)
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub


        Private Sub Matrix0_DoubleClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.DoubleClickAfter
            'Try
            '    Dim objcheck As SAPbouiCOM.CheckBox
            '    Select Case pVal.ColUID
            '        Case "Select"
            '            If pVal.Row = 0 Then
            '                'objform.Freeze(True)
            '                For i As Integer = 1 To Matrix0.RowCount
            '                    objcheck = Matrix0.Columns.Item("Select").Cells.Item(i).Specific
            '                    If objcheck.Checked = False Then
            '                        objcheck.Checked = True
            '                    End If
            '                Next
            '                'objform.Freeze(False)
            '            End If
            '    End Select
            'Catch ex As Exception
            '    objform.Freeze(False)
            '    objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            'Try
            '    Dim Flag As Boolean = False
            '    For i As Integer = 1 To Matrix0.RowCount
            '        If Matrix0.Columns.Item("Select").Cells.Item(i).Specific.Checked = True Then
            '            Flag = True
            '            Exit For
            '        End If
            '    Next
            '    If Flag = False Then
            '        objaddon.objapplication.StatusBar.SetText("Please Select atleast One SQ Line to Update...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        BubbleEvent = False : Exit Sub
            '    End If
            '    ValidateDateField()
            'Catch ex As Exception

            'End Try

        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            'Try
            '    Dim Len As String, Validity As String
            '    Select Case pVal.ColUID
            '        Case "Valid"
            '            If Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Specific.String <> "" Then
            '                Validity = Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Specific.String
            '                Len = Validity.Length
            '                If Len > 8 Or Len < 8 Then
            '                    objaddon.objapplication.StatusBar.SetText("Date Value should be in (ddMMyyyy) Format... Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                    Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
            '                    Exit Sub
            '                End If
            '                If Not IsNumeric(Validity) Then
            '                    objaddon.objapplication.StatusBar.SetText("Validity Field should contains number format...Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '                    Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
            '                    Exit Sub
            '                End If
            '            End If
            '    End Select
            'Catch ex As Exception
            '    objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try
        End Sub

        Private Sub ValidateDateField()
            Try
                Dim Len As String, Validity As String
                For i As Integer = 1 To Matrix0.RowCount
                    If Matrix0.Columns.Item("Select").Cells.Item(i).Specific.Checked = True Then
                        If Matrix0.Columns.Item("Valid").Cells.Item(i).Specific.String <> "" Then
                            Validity = Matrix0.Columns.Item("Valid").Cells.Item(i).Specific.String
                            Len = Validity.Length
                            If Len > 8 Or Len < 8 Then
                                objaddon.objapplication.StatusBar.SetText("Date Value should be in (ddMMyyyy) Format... Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            If Not IsNumeric(Validity) Then
                                objaddon.objapplication.StatusBar.SetText("Validity Field should contains number format...Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    End If
                Next

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub Matrix0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                If pVal.Row <> 0 Then
                    Matrix0.SelectRow(pVal.Row, True, False)
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox2 As SAPbouiCOM.ComboBox


        Private Sub Matrix0_ValidateBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.ValidateBefore
            Try
                Dim Len As String, Validity As String

                Select Case pVal.ColUID
                    Case "Valid"
                        If Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Validity = Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Specific.String
                            Len = Validity.Length
                            If Not IsNumeric(Validity) Then
                                objaddon.objapplication.StatusBar.SetText("Validity Field should contains number format...Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
                                BubbleEvent = False : Exit Sub
                            End If
                            If Len > 8 Or Len < 8 Then
                                objaddon.objapplication.StatusBar.SetText("Date Value should be in (ddMMyyyy) Format... Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
                                BubbleEvent = False : Exit Sub
                            End If
                            If Left(Validity, 2) < 1 Or Left(Validity, 2) > 31 Then
                                objaddon.objapplication.StatusBar.SetText("Number in Date should be less than or equal to 31... Not " & Left(Validity, 2) & ". Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
                                BubbleEvent = False : Exit Sub
                            End If
                            If Right(Left(Validity, 4), 2) < 1 Or Right(Left(Validity, 4), 2) > 12 Then
                                objaddon.objapplication.StatusBar.SetText("Number in Month should be less than or equal to 12... not " & Right(Left(Validity, 4), 2) & ". Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
                                BubbleEvent = False : Exit Sub
                            End If
                            If Right(Validity, 4) < DateTime.Now.Year Or Right(Validity, 4) > DateTime.Now.Year Then
                                objaddon.objapplication.StatusBar.SetText("Number in Year should be Current Year... Not " & Right(Validity, 4) & ". Line " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Matrix0.Columns.Item("Valid").Cells.Item(pVal.Row).Click()
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                End Select
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox3 As SAPbouiCOM.ComboBox
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox2 As SAPbouiCOM.CheckBox

        Private Sub Matrix0_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                If pVal.ColUID = "ItemCode" And pVal.ActionSuccess = True Then
                    Try
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            End Try
                            Try
                                Matrix0.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("ItemDesc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value
                            End Try
                        End If
                        objform.Update()
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents Button4 As SAPbouiCOM.Button

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                formMultiSelect = objaddon.objapplication.Forms.ActiveForm
                FieldName = "Make"
                If Not objaddon.objglobalmethods.FormExist("MULSEL") Then
                    Dim activeform As New FrmMultiSelect
                    activeform.Show()
                End If
              
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button4_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button4.ClickAfter
            Try
                formMultiSelect = objaddon.objapplication.Forms.ActiveForm
                FieldName = "MPN"
                If Not objaddon.objglobalmethods.FormExist("MULSEL") Then
                    Dim activeform As New FrmMultiSelect
                    activeform.Show()
                End If
                
            Catch ex As Exception

            End Try

        End Sub

        Private Sub CheckBox1_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox1.PressedAfter
            'Make
            Try
                If CheckBox1.Checked = True Then
                    objform.Freeze(True)
                    Button3.Item.Visible = True
                    objform.ActiveItem = "BuyName"
                    ComboBox0.Item.Visible = False
                    EditText6.Item.Visible = True
                    EditText6.Item.Top = 10
                    EditText6.Item.Left = 363
                Else
                    Button3.Item.Visible = False
                    ComboBox0.Item.Visible = True
                    objform.ActiveItem = "cmbmake"
                    EditText6.Item.Visible = False
                End If
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub CheckBox2_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox2.PressedAfter
            'MPN
            Try
                If CheckBox2.Checked = True Then
                    objform.Freeze(True)
                    Button4.Item.Visible = True
                    objform.ActiveItem = "BuyName"
                    ComboBox3.Item.Visible = False
                    EditText7.Item.Visible = True
                    EditText7.Item.Top = 28
                    EditText7.Item.Left = 363
                Else
                    Button4.Item.Visible = False
                    ComboBox3.Item.Visible = True
                    objform.ActiveItem = "cmbmpn"
                    EditText7.Item.Visible = False
                End If
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                EditText7.Item.Top = 28
                EditText7.Item.Left = 363
                EditText6.Item.Top = 10
                EditText6.Item.Left = 363
            Catch ex As Exception

            End Try

        End Sub

    End Class
End Namespace
