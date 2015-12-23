Public Class clsLoginSetup
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCheckBox, oCheckBox1 As SAPbouiCOM.CheckBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private oColumn As SAPbouiCOM.Column
    Private InvBase As DocumentType
    Private MatrixId As String
    Private RowtoDelete As Integer
    Private InvBaseDocNo As String
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim oDataSrc_Line, oDataSrc_Line1 As SAPbouiCOM.DBDataSource

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_LogSetup, frm_LogSetup)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "26"
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu("1283", True)
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.PaneLevel = 1
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT1")
        oForm.DataSources.UserDataSources.Add("LineID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oDataSrc_Line1 = oForm.DataSources.DBDataSources.Item("@Z_ITCAT2")
        oForm.DataSources.UserDataSources.Add("LineID1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        For count = 1 To oDataSrc_Line1.Size - 1
            oDataSrc_Line1.SetValue("LineId", count - 1, count)
        Next
        oForm.Items.Item("23").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Freeze(False)
    End Sub


#Region "Add Choose From List"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        oMatrix = aForm.Items.Item("25").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL1"
        oColumn.ChooseFromListAlias = "U_CatCode"

        oMatrix = aForm.Items.Item("28").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL2"
        oColumn.ChooseFromListAlias = "WhsCode"
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Z_ITCAT"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = 64
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_ESSWhs"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region
#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("25").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

    Private Sub AssignLineNo1(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("28").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT2")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region
#Region "Validations"
     Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oRec As SAPbobsCOM.Recordset
        Dim strLoginPassword, strSAPPassword As String
        Dim strcode, strcode1 As String
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            strSAPPassword = oApplication.Utilities.getEdittextvalue(aForm, "10")
            oCombobox = aForm.Items.Item("16").Specific
            If oCombobox.Selected.Value = "S" Then
                If oApplication.Utilities.getEdittextvalue(aForm, "8") = "" Then
                    oApplication.Utilities.Message("SAP UserId missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                End If
                If oApplication.Utilities.getEdittextvalue(aForm, "8") <> "" Then
                    If oApplication.Utilities.getEdittextvalue(aForm, "10") = "" Then
                        oApplication.Utilities.Message("SAP Password missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False
                    Else
                        strSAPPassword = oApplication.Utilities.getEdittextvalue(aForm, "10")
                    End If
                End If
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "12") = "" Then
                oApplication.Utilities.Message("UserId missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "14") = "" Then
                oApplication.Utilities.Message("Password missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            Else
                strLoginPassword = oApplication.Utilities.getEdittextvalue(aForm, "14")
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Employee ID missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If

            oCheckBox = aForm.Items.Item("19").Specific
            If oCheckBox.Checked = False Then
                oMatrix = aForm.Items.Item("25").Specific
                If oMatrix.RowCount = 0 Then
                    oApplication.Utilities.Message("Item Category Row Cannot be Empty...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If

            oCheckBox = aForm.Items.Item("20").Specific
            If oCheckBox.Checked = False Then
                oMatrix = aForm.Items.Item("28").Specific
                If oMatrix.RowCount = 0 Then
                    oApplication.Utilities.Message("Warehouse Row Cannot be Empty...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                End If
            End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRec.DoQuery("Select * from [@Z_DLC_LOGIN] where U_EMPID='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "'")
            Else
                oRec.DoQuery("Select * from [@Z_DLC_LOGIN] where U_EMPID='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and DocEntry <> '" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "'")
            End If
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.Message("Record already exists for this employee...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRec.DoQuery("Select * from [@Z_DLC_LOGIN] where Upper(U_UID)='" & oApplication.Utilities.getEdittextvalue(aForm, "12").ToUpper() & "'")
            Else
                oRec.DoQuery("Select * from [@Z_DLC_LOGIN] where Upper(U_UID)='" & oApplication.Utilities.getEdittextvalue(aForm, "12").ToUpper() & "' and DocEntry <> '" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "'")
            End If
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.Message("Record already exists for this ESS User...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strEncryptText As String = oApplication.Utilities.Encrypt(strLoginPassword, oApplication.Utilities.key)
            oApplication.Utilities.setEdittextvalue(aForm, "14", strEncryptText) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

            Dim strEncryptText1 As String = oApplication.Utilities.Encrypt(strSAPPassword, oApplication.Utilities.key)
            oApplication.Utilities.setEdittextvalue(aForm, "10", strEncryptText1) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

            oMatrix = aForm.Items.Item("25").Specific
            oCheckBox = aForm.Items.Item("19").Specific
            If oCheckBox.Checked = False Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    strcode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                    For intLoop As Integer = intRow + 1 To oMatrix.RowCount
                        strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intLoop)
                        If strcode1 <> "" Then
                            If strcode.ToUpper = strcode1.ToUpper Then
                                oApplication.Utilities.Message("This entry already exists : " & strcode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oMatrix.Columns.Item("V_0").Cells.Item(intLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Return False
                            End If
                        End If
                    Next
                Next
            End If

            oMatrix = aForm.Items.Item("28").Specific
            oCheckBox1 = aForm.Items.Item("20").Specific
            If oCheckBox1.Checked = False Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    strcode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
                    For intLoop As Integer = intRow + 1 To oMatrix.RowCount
                        strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intLoop)
                        If strcode1 <> "" Then
                            If strcode.ToUpper = strcode1.ToUpper Then
                                oApplication.Utilities.Message("This entry already exists : " & strcode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oMatrix.Columns.Item("V_0").Cells.Item(intLoop).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Return False
                            End If
                        End If
                    Next
                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("25").Specific
                    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                        End If

                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDataSrc_Line.Size
                        oDataSrc_Line.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
                Case "2"
                    oMatrix = aForm.Items.Item("28").Specific
                    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT2")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    End If
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    Try
                        If oEditText.Value <> "" Then
                            oMatrix.AddRow()
                            oMatrix.ClearRowData(oMatrix.RowCount)
                        End If

                    Catch ex As Exception
                        aForm.Freeze(False)
                    End Try
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDataSrc_Line.Size
                        oDataSrc_Line.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo1(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId = "25" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT1")
        ElseIf Me.MatrixId = "28" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_ITCAT2")
        End If
        If intSelectedMatrixrow <= 0 Then
            Exit Sub
        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
        End If
    End Sub

#End Region




#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_LogSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "25" And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item("25").Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "25"
                                    frmSourceMatrix = oMatrix
                                End If
                                If pVal.ItemUID = "28" And pVal.Row > 0 Then
                                    oMatrix = oForm.Items.Item("28").Specific
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "28"
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "12"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "13"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "23"
                                        oForm.PaneLevel = 1
                                    Case "24"
                                        oForm.PaneLevel = 2
                                    Case "19"
                                        oCheckBox = oForm.Items.Item("19").Specific
                                        If oCheckBox.Checked = True Then
                                            oMatrix = oForm.Items.Item("25").Specific
                                            oMatrix.Clear()
                                            oForm.Items.Item("25").Enabled = False
                                        Else
                                            oForm.Items.Item("25").Enabled = True
                                        End If
                                    Case "20"
                                        oCheckBox = oForm.Items.Item("20").Specific
                                        If oCheckBox.Checked = True Then
                                            oMatrix = oForm.Items.Item("28").Specific
                                            oMatrix.Clear()
                                            oForm.Items.Item("28").Enabled = False
                                        Else
                                            oForm.Items.Item("28").Enabled = True
                                        End If
                                End Select
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
                                        If pVal.ItemUID = "25" And pVal.ColUID = "V_0" Then
                                            val1 = oDataTable.GetValue("U_CatCode", 0)
                                            val = oDataTable.GetValue("U_CatDesc", 0)
                                            oMatrix = oForm.Items.Item("25").Specific
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val1)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                        If pVal.ItemUID = "28" And pVal.ColUID = "V_0" Then
                                            val1 = oDataTable.GetValue("WhsCode", 0)
                                            val = oDataTable.GetValue("WhsName", 0)
                                            oMatrix = oForm.Items.Item("28").Specific
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val1)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val2 = oDataTable.GetValue("userId", 0)
                                            val1 = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("middleName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val2)
                                            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRec.DoQuery("Select isnull(USER_CODE,'') from OUSR where INTERNAL_K='" & val2 & "'")
                                            If oRec.RecordCount > 0 Then
                                                Try
                                                    oApplication.Utilities.setEdittextvalue(oForm, "8", oRec.Fields.Item(0).Value)
                                                Catch ex As Exception
                                                End Try
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        End If
                                        If pVal.ItemUID = "8" Then
                                            val = oDataTable.GetValue("USER_CODE", 0)
                                            val1 = oDataTable.GetValue("INTERNAL_K", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "8", val)
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
                Case mnu_LogSetup
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("5").Enabled = False
                        oForm.Items.Item("7").Enabled = False
                    End If
                Case mnu_ADD_ROW

                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else
                        'If ValidateDeletion(oForm) = False Then
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("5").Enabled = True
                        oForm.Items.Item("7").Enabled = True
                        oCheckBox = oForm.Items.Item("19").Specific
                        If oCheckBox.Checked = True Then
                            oMatrix = oForm.Items.Item("25").Specific
                            oMatrix.Clear()
                            oForm.Items.Item("25").Enabled = False
                        Else
                            oForm.Items.Item("25").Enabled = True
                        End If

                        oCheckBox = oForm.Items.Item("20").Specific
                        If oCheckBox.Checked = True Then
                            oMatrix = oForm.Items.Item("28").Specific
                            oMatrix.Clear()
                            oForm.Items.Item("28").Enabled = False
                        Else
                            oForm.Items.Item("28").Enabled = True
                        End If
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        oForm.Items.Item("5").Enabled = True
                        oForm.Items.Item("7").Enabled = True
                    End If
                Case "1283"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        Dim strValue As String
                        If oApplication.SBO_Application.MessageBox("Do you want to delete the details?", , "Yes", "No") = 2 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
       Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()

                Dim strEncryptText As String = oApplication.Utilities.getLoginPassword(oApplication.Utilities.getEdittextvalue(oForm, "14"))
                oApplication.Utilities.setEdittextvalue(oForm, "14", strEncryptText) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

                Dim strEncryptText1 As String = oApplication.Utilities.getLoginPassword(oApplication.Utilities.getEdittextvalue(oForm, "10"))
                oApplication.Utilities.setEdittextvalue(oForm, "10", strEncryptText1) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

                oCheckBox = oForm.Items.Item("19").Specific
                If oCheckBox.Checked = True Then
                    oMatrix = oForm.Items.Item("25").Specific
                    oMatrix.Clear()
                    oForm.Items.Item("25").Enabled = False
                Else
                    oForm.Items.Item("25").Enabled = True
                End If

                oCheckBox = oForm.Items.Item("20").Specific
                If oCheckBox.Checked = True Then
                    oMatrix = oForm.Items.Item("28").Specific
                    oMatrix.Clear()
                    oForm.Items.Item("28").Enabled = False
                Else
                    oForm.Items.Item("28").Enabled = True
                End If
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
