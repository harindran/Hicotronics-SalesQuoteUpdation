Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SalesQuoteUpdation
    <FormAttribute("MULSEL", "SalesQuote/FrmMultiSelect.b1f")>
    Friend Class FrmMultiSelect
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("MtxLoad").Specific, SAPbouiCOM.Matrix)
            Me.Button2 = CType(Me.GetItem("3").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULSEL", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                objform.EnableMenu("1292", True)
                objform.EnableMenu("773", True)
                If FieldName = "Make" Then
                    objform.Title = "Select Make"
                    Matrix0.Columns.Item("Code").TitleObject.Caption = "Make"
                ElseIf FieldName = "MPN" Then
                    objform.Title = "Select MPN"
                    Matrix0.Columns.Item("Code").TitleObject.Caption = "MPN"
                End If
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim Code As String = "'"
                For i As Integer = 1 To Matrix0.VisualRowCount
                    Code = Code + Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String + "','"
                Next
                Code = Code.Remove(Code.Length - 5)
                If FieldName = "Make" Then
                    formMultiSelect.Items.Item("txtMake").Specific.String = ""
                    formMultiSelect.Items.Item("txtMake").Specific.String = Code
                ElseIf FieldName = "MPN" Then
                    formMultiSelect.Items.Item("txtMPN").Specific.String = ""
                    formMultiSelect.Items.Item("txtMPN").Specific.String = Code
                End If

                formMultiSelect = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents Button2 As SAPbouiCOM.Button

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Matrix0.Clear()
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ValidateBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.ValidateBefore
            Try
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try
        End Sub
    End Class
End Namespace
