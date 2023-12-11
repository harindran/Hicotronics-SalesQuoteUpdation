Namespace SalesQuoteUpdation

    Public Class clsTable

        Public Sub FieldCreation()
            SalesQuoteUpdate()
            AddFields("@BUYERMAKE", "BuyName", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@BUYERMAKE", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@BUYERMAKE", "SUser", "Super User", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", , {"Y,Yes", "N,No"})
            AddFields("QUT1", "BuyName", "BuyName", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("QUT1", "SupRefNo", "Sup RefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("QUT1", "QutValid", "Quote Valid", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("QUT1", "Cost", "Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("QUT1", "CustPrice", "Customer TPrice", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("QUT1", "HMPN", "Hico MPN", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("QUT1", "HDescription", "Hico Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("QUT1", "HMake", "Hico Make", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("QUT1", "BuyRem", "Buyer Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("QUT1", "EndCust", "End Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("QUT1", "Remarks1", "Remarks1", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("QUT1", "ProjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("QUT1", "Applictn", "Application", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("QUT1", "AnnualVol", "AnnualVol Name", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("QUT1", "CostCur", "Cost Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddTables("MI_MSEL", "SalesQuote Updation Header", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        End Sub



#Region "Document Data Creation"

        Private Sub SalesQuoteUpdate()
            AddTables("MIPL_OQUT", "SalesQuote Updation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MIPL_QUT1", "SalesQuote Updation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MIPL_OQUT", "UserName", "User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OQUT", "CustName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OQUT", "SQNum", "SQ Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_OQUT", "DocDate", "Doc Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OQUT", "FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OQUT", "ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OQUT", "Make", "Make", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_OQUT", "MPN", "MPN", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OQUT", "SQCat", "SQ Category", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "-", , {"-,-", "Bid,Bid", "Buy,Buy"})
            AddFields("@MIPL_OQUT", "ChkMake", "Make Enable", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_OQUT", "ChkMPN", "MPN Enable", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)

            AddFields("@MIPL_QUT1", "Select", "Select", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_QUT1", "SQNum", "SQ Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_QUT1", "SQLine", "SQ LineNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MIPL_QUT1", "CUser", "Created User", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "DocDueDate", "Doc Due Date", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_QUT1", "CustName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "CustType", "Customer Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_QUT1", "Itemcode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_QUT1", "MPN", "SAP MPN", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "Make", "SAP Make", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "SPQ", "SAP SPQ", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_QUT1", "MOQ", "SAP MOQ", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_QUT1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_QUT1", "TPrice", "Customer TPrice", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_QUT1", "HicoMPN", "Hico MPN", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "HicoDesc", "Hico Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_QUT1", "HicoMake", "Hico Make", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            'AddFields("@MIPL_QUT1", "QutValid", "Quote Validity", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_QUT1", "QutValid", "Quote Validity", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_QUT1", "Cost", "Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_QUT1", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_QUT1", "LeadTIme", "Lead Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "SupRefNo", "Sup RefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "BuyRem", "Buyer Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "EndCust", "End Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "Remarks1", "Remarks1", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_QUT1", "ProjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "Applictn", "Application", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_QUT1", "AnnualVol", "AnnualVol Name", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)



            AddUDO("OQUT", "SalesQuote Updation", SAPbobsCOM.BoUDOObjType.boud_Document, "MIPL_OQUT", {"MIPL_QUT1"}, {"DocEntry", "DocNum", "U_CustName", "U_Make"}, True, True)
        End Sub

#End Region

#Region "Table Creation Common Functions"

        Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Try
                oUserTablesMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                'Adding Table
                If Not oUserTablesMD.GetByKey(strTab) Then
                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType

                    If oUserTablesMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription & strTab)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub AddFields(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoFieldTypes, _
                             Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, _
                              Optional ByVal defaultvalue As String = "", Optional ByVal Yesno As Boolean = False, Optional ByVal Validvalues() As String = Nothing)
            Dim oUserFieldMD1 As SAPbobsCOM.UserFieldsMD
            oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                'If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                '    strTab = "@" + strTab
                'End If
                If Not IsColumnExists(strTab, strCol) Then
                    'If Not oUserFieldMD1 Is Nothing Then
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    'End If
                    'oUserFieldMD1 = Nothing
                    'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc
                    oUserFieldMD1.Name = strCol
                    oUserFieldMD1.Type = nType
                    oUserFieldMD1.SubType = nSubType
                    oUserFieldMD1.TableName = strTab
                    oUserFieldMD1.EditSize = nEditSize
                    oUserFieldMD1.Mandatory = Mandatory
                    oUserFieldMD1.DefaultValue = defaultvalue

                    If Yesno = True Then
                        oUserFieldMD1.ValidValues.Value = "Y"
                        oUserFieldMD1.ValidValues.Description = "Yes"
                        oUserFieldMD1.ValidValues.Add()
                        oUserFieldMD1.ValidValues.Value = "N"
                        oUserFieldMD1.ValidValues.Description = "No"
                        oUserFieldMD1.ValidValues.Add()
                    End If

                    Dim split_char() As String
                    If Not Validvalues Is Nothing Then
                        If Validvalues.Length > 0 Then
                            For i = 0 To Validvalues.Length - 1
                                If Trim(Validvalues(i)) = "" Then Continue For
                                split_char = Validvalues(i).Split(",")
                                If split_char.Length <> 2 Then Continue For
                                oUserFieldMD1.ValidValues.Value = split_char(0)
                                oUserFieldMD1.ValidValues.Description = split_char(1)
                                oUserFieldMD1.ValidValues.Add()
                            Next
                        End If
                    End If
                    Dim val As Integer
                    val = oUserFieldMD1.Add
                    If val <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage(objaddon.objcompany.GetLastErrorDescription & strTab & strCol, True)
                    End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                End If
            Catch ex As Exception
                Throw ex
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                oUserFieldMD1 = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strSQL As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                Else
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
                End If

                oRecordSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Function

        Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
            Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

            Try
                '// The meta-data object must be initialized with a
                '// regular UserKeys object
                oUserKeysMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

                If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                    '// Set the table name and the key name
                    oUserKeysMD.TableName = strTab
                    oUserKeysMD.KeyName = strKey

                    '// Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn
                    oUserKeysMD.Elements.Add()
                    oUserKeysMD.Elements.ColumnAlias = "RentFac"

                    '// Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                    '// Add the key
                    If oUserKeysMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
                oUserKeysMD = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub AddUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String, _
                           Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If

            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub

#End Region

    End Class
End Namespace
