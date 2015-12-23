Public Class clsItemMaster
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub AddControls(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            AddChooseFromList(aForm)
            oApplication.Utilities.AddControls(aForm, "HRstcomp", "25", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Item Category", 120)
            oApplication.Utilities.AddControls(aForm, "HRedComp", "24", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, , , 145)
            oApplication.Utilities.AddControls(aForm, "HRstcoNa", "52", SAPbouiCOM.BoFormItemTypes.it_STATIC, "RIGHT", 0, 0, , "Category Name", 120)
            oApplication.Utilities.AddControls(aForm, "HRedCoNa", "34", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", 0, 0, , , 120)
            oEditText = aForm.Items.Item("HRedComp").Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_CatCode")
            oEditText.ChooseFromListUID = "CFL_1"
            oEditText.ChooseFromListAlias = "U_CatCode"
            oEditText = aForm.Items.Item("HRedCoNa").Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_CatDesc")
            aForm.Items.Item("HRedComp").Enabled = True
            aForm.Items.Item("HRedCoNa").Visible = False
            aForm.Items.Item("HRstcoNa").Visible = False
            oItem = aForm.Items.Item("HRstcomp")
            oItem.LinkTo = "HRedComp"
            oItem = aForm.Items.Item("HRstcoNa")
            oItem.LinkTo = "HRedCoNa"
        Catch ex As Exception
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub AddChooseFromList(ByVal aform As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = aform.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams


            ' Adding 1 CFL, one for the button and one for the edit text.
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_ITCAT"
            oCFLCreationParams.UniqueID = "CFL_1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ItemMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm.Freeze(True)
                                AddControls(oForm)
                                oForm.Freeze(False)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oRec As SAPbobsCOM.Recordset
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim val2 As Integer
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "HRedComp" Then
                                            val = oDataTable.GetValue("U_CatCode", 0)
                                            val1 = oDataTable.GetValue("U_CatDesc", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "HRedCoNa", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "HRedComp", val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvSO
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
