Public Class clsChooseFromList
    Inherits clsBase

#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String
    Public Shared frmFromBin As String
    Public Shared frmWarehouse As String
    Public sourceGrid As SAPbouiCOM.Grid
    Public sourceMatrix As SAPbouiCOM.Matrix
    Private ocombo As SAPbouiCOM.ComboBox
    Private oComboBox As SAPbouiCOM.ComboBoxColumn

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditText
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Grid
    Private Const SEPRATOR As String = "~~~"
    Private SelectedRow As Integer
    Private sSearchColumn As String
    Private oItem As SAPbouiCOM.Item
    Public stritemid As SAPbouiCOM.Item
    Private intformmode As SAPbouiCOM.BoFormMode
    Private objGrid As SAPbouiCOM.Grid
    Private objSourcematrix As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private objStatic As SAPbouiCOM.StaticText
    Private inttable As Integer = 0
    Public strformid As String
    Public strStaticValue As String
    Public strSQL As String
    Private strSelectedItem1 As String
    Private strSelectedItem2 As String
    Private strSelectedItem3 As String
    Private strSelectedItem4, strSelectedItem5 As String
    Private oRecSet As SAPbobsCOM.Recordset
    '   Private objSBOAPI As ClsSBO
    '   Dim objTransfer As clsTransfer
#End Region

#Region "New"
    '*****************************************************************
    'Type               : Constructor
    'Name               : New
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create object for classes.
    '******************************************************************
    Public Sub New()
        '   objSBOAPI = New ClsSBO
        MyBase.New()
    End Sub
#End Region

#Region "Bind Data"
    '******************************************************************
    'Type               : Procedure
    'Name               : BindData
    'Parameter          : Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Binding the fields.
    '******************************************************************
    Public Sub databound(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objform.Freeze(True)
            objform.DataSources.DataTables.Add("dtLevel3")
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
            oedit = objform.Items.Item("etFind").Specific
            oedit.DataBind.SetBound(True, "", "dbFind")
            objGrid = objform.Items.Item("mtchoose").Specific
            dtTemp = objform.DataSources.DataTables.Item("dtLevel3")

            Dim oTempRS As SAPbobsCOM.Recordset
            Dim stItemCheck As String
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
           
           If choice = "Dept" Then
                If CFLChoice = "Department" Then
                    objform.Title = "Department Selection"
                    strSQL = "SELECT Code,Remarks from OUDP"
                    dtTemp.ExecuteQuery(strSQL)
                    objGrid.DataTable = dtTemp
                    objGrid.Columns.Item(0).TitleObject.Caption = "Department Code"
                    objGrid.Columns.Item(1).TitleObject.Caption = "Description"
                End If

            End If
            objGrid.AutoResizeColumns()
            objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            If objGrid.Rows.Count > 0 Then
                objGrid.Rows.SelectedRows.Add(0)
            End If
            objform.Freeze(False)
            objform.Update()
            'sSearchList = " "
            'Dim i As Integer = 0
            'While i <= objGrid.DataTable.Rows.Count - 1
            '    sSearchList += Convert.ToString(objGrid.DataTable.GetValue(0, i)) + SEPRATOR + i.ToString + " "
            '    System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
            'End While
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region


#Region "Get Form"
    '******************************************************************
    'Type               : Function
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Get The Form
    '******************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return oApplication.SBO_Application.Forms.Item(FormUID)
    End Function
#End Region

#Region "FormDataEvent"


#End Region

#Region "Class Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

    End Sub
#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        BubbleEvent = True
        If pVal.FormTypeEx = frm_ChoosefromList Then
            If pVal.Before_Action = True Then
                If pVal.ItemUID = "mtchoose" Then
                    Try
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row <> -1 Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item(pVal.ItemUID), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            oMatrix.Rows.SelectedRows.Add(pVal.Row)
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.Row <> -1 Then
                    Try
                        oForm = GetForm(pVal.FormUID)
                        oItem = oForm.Items.Item("mtchoose")
                        oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                        Dim inti As Integer
                        inti = 0
                        inti = 0
                        While inti <= oMatrix.DataTable.Rows.Count - 1
                            If oMatrix.Rows.IsSelected(inti) = True Then
                                intRowId = inti
                            End If
                            System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                        End While
                        If CFLChoice = "Department" Then
                            If intRowId = 0 Then
                                strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                            Else
                                strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                            End If
                            oForm.Close()
                            oForm = GetForm(SourceFormUID)
                            If choice = "Dept" Then
                                oForm.Freeze(True)
                                sourceMatrix = oForm.Items.Item(ItemUID).Specific
                                oApplication.Utilities.SetMatrixValues(sourceMatrix, "V_1", SourceLabel, strSelectedItem2)
                                oApplication.Utilities.SetMatrixValues(sourceMatrix, "V_0", SourceLabel, strSelectedItem1)
                                oForm.Freeze(False)
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If



                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    Try
                        If pVal.ItemUID = "mtchoose" Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            intRowId = pVal.Row - 1
                        End If
                        Dim inti As Integer
                        If pVal.CharPressed = 13 Then
                            inti = 0
                            inti = 0
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)

                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            While inti <= oMatrix.DataTable.Rows.Count - 1
                                If oMatrix.Rows.IsSelected(inti) = True Then
                                    intRowId = inti
                                End If
                                System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                            End While
                            If CFLChoice = "Training" Then
                                If intRowId = 0 Then
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(2, intRowId))
                                    strSelectedItem3 = Convert.ToString(oMatrix.DataTable.GetValue(3, intRowId))
                                    strSelectedItem4 = Convert.ToString(oMatrix.DataTable.GetValue(5, intRowId))
                                Else
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(2, intRowId))
                                    strSelectedItem3 = Convert.ToString(oMatrix.DataTable.GetValue(3, intRowId))
                                    strSelectedItem4 = Convert.ToString(oMatrix.DataTable.GetValue(5, intRowId))
                                End If
                                oForm.Close()
                                oForm = GetForm(SourceFormUID)
                                If choice = "Bin" Then
                                    If CFLChoice = "Training" Then
                                        oForm.Freeze(True)
                                        sourceGrid = oForm.Items.Item(ItemUID).Specific
                                        sourceGrid.DataTable.SetValue(sourceColumID, SourceLabel, strSelectedItem2)
                                        sourceGrid.DataTable.SetValue("U_Z_RefCode", SourceLabel, strSelectedItem1)
                                        sourceGrid.DataTable.SetValue("U_Z_HREmpName", SourceLabel, strSelectedItem3)
                                        sourceGrid.DataTable.SetValue("U_Z_DeptName", SourceLabel, strSelectedItem4)
                                        sourceGrid.DataTable.SetValue("U_Z_Status", SourceLabel, "A")
                                        oForm.Freeze(False)
                                    ElseIf CFLChoice = "Insurance" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, ItemUID, strSelectedItem2)
                                    End If
                                Else
                                    sourceGrid = oForm.Items.Item(ItemUID).Specific
                                    sourceGrid.DataTable.SetValue(sourceColumID, SourceLabel, strSelectedItem2)
                                    sourceGrid.DataTable.SetValue(7, SourceLabel, strSelectedItem2)

                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 And pVal.ItemUID = "btnChoose" Then
                    Try
                        oForm = GetForm(pVal.FormUID)
                        oItem = oForm.Items.Item("mtchoose")
                        oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                        Dim inti As Integer
                        inti = 0
                        inti = 0
                        While inti <= oMatrix.DataTable.Rows.Count - 1
                            If oMatrix.Rows.IsSelected(inti) = True Then
                                intRowId = inti
                            End If
                            System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                        End While
                        If CFLChoice = "Department" Then
                            If intRowId = 0 Then
                                strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                            Else
                                strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                            End If
                            oForm.Close()
                            oForm = GetForm(SourceFormUID)
                            If choice = "Dept" Then
                                oForm.Freeze(True)
                                sourceMatrix = oForm.Items.Item(ItemUID).Specific
                                oApplication.Utilities.SetMatrixValues(sourceMatrix, "V_1", SourceLabel, strSelectedItem2)
                                oApplication.Utilities.SetMatrixValues(sourceMatrix, "V_0", SourceLabel, strSelectedItem1)
                                oForm.Freeze(False)
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If
            Else
                If pVal.BeforeAction = False Then
                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) Then
                        BubbleEvent = False
                        Dim objGrid As SAPbouiCOM.Grid
                        Dim oedit As SAPbouiCOM.EditText
                        If pVal.ItemUID = "etFind" And pVal.CharPressed <> "13" Then
                            Dim i, j As Integer
                            Dim strItem As String
                            oForm = oApplication.SBO_Application.Forms.ActiveForm() 'oApplication.SBO_Application.Forms.GetForm("TWBS_FA_CFL", pVal.FormTypeCount)
                            objGrid = oForm.Items.Item("mtchoose").Specific
                            oedit = oForm.Items.Item("etFind").Specific
                            For i = 0 To objGrid.DataTable.Rows.Count - 1
                                strItem = ""
                                strItem = objGrid.DataTable.GetValue(0, i)
                                For j = 1 To oedit.String.Length
                                    If oedit.String.Length <= strItem.Length Then
                                        If strItem.Substring(0, j).ToUpper = oedit.String.ToUpper Then
                                            objGrid.Rows.SelectedRows.Add(i)
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub
#End Region

End Class
