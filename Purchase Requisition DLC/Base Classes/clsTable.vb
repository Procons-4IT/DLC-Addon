Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
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

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OADM" Or strTab = "OITT" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "OWHS" Or strTab = "OUDP" Or strTab = "OPRQ") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

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
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
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

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                Dim intTables As Integer = 0
                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
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
    Public Function UDOItemCategory(ByVal strUDO As String, _
                       ByVal strDesc As String, _
                           ByVal strTable As String, _
                               ByVal intFind As Integer, _
                                   Optional ByVal strCode As String = "", _
                                       Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_CatCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_CatCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_CatDesc"
                oUserObjects.FormColumns.FormColumnDescription = "U_CatDesc"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddTables("Z_ITCAT", "Item Category", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_ITCAT", "CatCode", "Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_ITCAT", "CatDesc", "Category Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 120)

            addField("OWHS", "ESSWhs", "ESS Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OITM", "CatCode", "Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("OITM", "CatDesc", "Category Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 120)
            AddFields("OUDP", "DefWhs", "Default Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)

            AddTables("Z_DLC_LOGIN", "Login Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_DLC_LOGIN", "UID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_LOGIN", "PWD", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_LOGIN", "EMPID", "Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DLC_LOGIN", "EMPNAME", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_DLC_LOGIN", "ESSLogType", "ESS LoginType", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "E,S", "Employee,Store Keeper", "E")
            AddFields("Z_DLC_LOGIN", "EMPUID", "Emp.User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_LOGIN", "USERPWD", "User Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_LOGIN", "INTID", "Internal ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_DLC_LOGIN", "AllItemCat", "All ItemCategory", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("Z_DLC_LOGIN", "AllWhs", "All Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddTables("Z_ITCAT1", "Item Category", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_ITCAT1", "CatCode", "Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_ITCAT1", "CatDesc", "Category Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 120)

            AddTables("Z_ITCAT2", "Warehouse", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_ITCAT2", "whsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_ITCAT2", "whsDesc", "Warehouse Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 120)


            AddTables("Z_DLC_OAPPT", "Approval Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_DLC_APPT2", "Approval Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_DLC_APPT3", "Department Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("Z_DLC_OAPPT", "Z_Code", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_OAPPT", "Z_Name", "Approval Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_OAPPT", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_OAPPT", "Z_DocDesc", "Document Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_OAPPT", "Z_Active", "Active Template", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("Z_DLC_OAPPT", "Z_AllDept", "All Department", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddFields("Z_DLC_APPT3", "Z_DeptCode", "Department Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APPT3", "Z_DeptName", "Department Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("Z_DLC_APPT2", "Z_AUser", "Authorizer Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APPT2", "Z_AName", "Authorizer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_APPT2", "Z_AMan", "Mandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_DLC_APPT2", "Z_AFinal", "Final Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)


            AddTables("Z_DLC_APHIS", "Approval History", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_DLC_APHIS", "Z_DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APHIS", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APHIS", "Z_EmpId", "Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APHIS", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_DLC_APHIS", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,D,C,L,R,A", "Pending,Approved,Close,Cancelled,DLC Rejected,DLC Approved", "P")
            AddFields("Z_DLC_APHIS", "Z_Remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_DLC_APHIS", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_DLC_APHIS", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_DLC_APHIS", "Z_ADocEntry", "Template DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_DLC_APHIS", "Z_ALineId", "Template LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_DLC_APHIS", "Z_ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_APHIS", "Z_OrdQty", "Order Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_DLC_APHIS", "Z_DelUom", "Delivered UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_DLC_APHIS", "Z_DelUomDesc", "Delivered UoM Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_APHIS", "Z_ItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_DLC_APHIS", "Z_DLineId", "Document LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddTables("Z_OPRQ", "Purchase Request Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPRQ", "Z_EmpID", "Request Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPRQ", "Z_EmpName", "Request Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OPRQ", "Z_DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPRQ", "Z_DeptCode", "Department Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_OPRQ", "Z_DeptName", "Department Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_OPRQ", "Z_Priority", "Priority", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "L,M,H", "Low,Medium,High", "L")
            addField("@Z_OPRQ", "Z_DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_Address, "O,I,C,D,R,L,S,DI", "Open,InProgress,Closed,Draft,DLC Rejected,Cancelled,Confirm,DLC InProgress", "D")
            AddFields("Z_OPRQ", "Z_Destination", "Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_PRQ1", "Purchase Request Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PRQ1", "Z_DocNo", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PRQ1", "Z_ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRQ1", "Z_ItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRQ1", "Z_OrdQty", "Order Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRQ1", "Z_OrdUom", "Order UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PRQ1", "Z_OrdUomDesc", "Order UoM Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRQ1", "Z_AltItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRQ1", "Z_AltItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRQ1", "Z_DeliQty", "Delivered Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRQ1", "Z_DelUom", "Delivered UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PRQ1", "Z_DelUomDesc", "Delivered UoM Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRQ1", "Z_RecQty", "Received Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PRQ1", "Z_RecUom", "Received UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PRQ1", "Z_RecUomDesc", "Received UoM Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_PRQ1", "Z_OrdPatient", "Order Related to Patients", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_PRQ1", "Z_BarCode", "Item BarCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRQ1", "Z_AltBarCode", "Alternate Item BarCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_PRQ1", "Z_LineStatus", "Line Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,D,C,L", "Open,Delivered,Close,Cancelled", "O")
            AddFields("Z_PRQ1", "Z_CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PRQ1", "Z_NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_PRQ1", "Z_AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_PRQ1", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PRQ1", "Z_ReqTime", "Required Time", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PRQ1", "Z_RejRemark", "Rejection Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PRQ1", "Z_ApproveId", "Approve Template Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PRQ1", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PRQ1", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_PRQ1", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,R,A", "Pending,Rejected,Approved", "P")
            addField("@Z_PRQ1", "Z_GoodIssue", "Goods Issued", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            '---- User Defined Object's


            AddTables("Z_DLC_OMAIL", "Email SetUp Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_DLC_OMAIL", "Z_SMTPSERV", "SMTP SERVER", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_OMAIL", "Z_SMTPPORT", "SMTP PORT", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_DLC_OMAIL", "Z_SMTPUSER", "SMTP USER", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_OMAIL", "Z_SMTPPWD", "SMTP PASSWORD", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DLC_OMAIL", "Z_SSL", "SMTP SSL", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            AddTables("Z_ORPD", "Material Return Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ORPD", "Z_EmpID", "Request Employee ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_ORPD", "Z_EmpName", "Request Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_ORPD", "Z_DeptCode", "Department Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_ORPD", "Z_DeptName", "Department Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_ORPD", "Z_DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_ORPD", "Z_DocNo", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            addField("@Z_ORPD", "Z_DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_Address, "O,I,C,D,R,L,S,DI", "Open,InProgress,Closed,Draft,DLC Rejected,Cancelled,Confirm,DLC InProgress", "D")
            AddFields("Z_ORPD", "Z_Destination", "Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_RPD1", "Material Return Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_RPD1", "Z_ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RPD1", "Z_ItemName", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_RPD1", "Z_OrdQty", "Order Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_RPD1", "Z_OrdUom", "Order UoM Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_RPD1", "Z_OrdUomDesc", "Order UoM Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RPD1", "Z_BarCode", "Item BarCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_RPD1", "Z_LineStatus", "Line Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,D,L", "Open,Delivered,Cancelled", "O")
            AddFields("Z_RPD1", "Z_CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_RPD1", "Z_NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_RPD1", "Z_AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_RPD1", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_RPD1", "Z_ReqTime", "Required Time", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_RPD1", "Z_RejRemark", "Rejection Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_RPD1", "Z_ApproveId", "Approve Template Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_RPD1", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_RPD1", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_RPD1", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,R,A", "Pending,Rejected,Approved", "P")
            AddFields("Z_RPD1", "Z_DocRefNo", "Goods Return Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            addField("@Z_RPD1", "Z_NewDoc", "New Document", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_RPD1", "Z_GoodReceipt", "Goods Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            CreateUDO()

            oApplication.Utilities.Message("Database Created Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("Z_OPRQ", "Purchase Requisition", "Z_OPRQ", "U_Z_EmpID", "DocEntry", "Z_PRQ1", , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_ORPD", "Material Return", "Z_ORPD", "U_Z_EmpID", "DocEntry", "Z_RPD1", , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_DLC_LOGIN", "Login Setup", "Z_DLC_LOGIN", "U_EMPID", "DocEntry", "Z_ITCAT1", "Z_ITCAT2", SAPbobsCOM.BoUDOObjType.boud_Document)
            UDOItemCategory("Z_ITCAT", "Item Category", "Z_ITCAT", 1, "U_CatCode", )
            AddUDO("Z_DLC_APHIS", "Approval History", "Z_DLC_APHIS", "DocEntry", "U_Z_DocEntry", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_DLC_OAPPT", "Approval Template", "Z_DLC_OAPPT", "DocEntry", "U_Z_Code", "Z_DLC_APPT2", "Z_DLC_APPT3", SAPbobsCOM.BoUDOObjType.boud_Document)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
