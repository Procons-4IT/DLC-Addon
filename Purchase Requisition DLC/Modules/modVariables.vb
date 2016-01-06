Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String
    Public intSelectedMatrixrow As Integer = 0
    Public frmSourceForm As SAPbouiCOM.Form

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62

    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_Invoice As String = "133"
  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_CloseOrderLines As String = "DABT_910"
    Public Const mnu_InvSO As String = "DABT_911"
    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_StRequest As String = "StRequest.xml"
    Public Const xml_InvSO As String = "frm_InvSO.xml"
   
    Public Const mnu_ItemCat As String = "mnu_ItemCat"
    Public Const frm_ItemCat As String = "frm_ItemCat"
    Public Const xml_ItemCat As String = "frm_ItemCat.xml"

    Public Const mnu_LogSetup As String = "mnu_LogSetup"
    Public Const frm_LogSetup As String = "frm_LogSetup"
    Public Const xml_LogSetup As String = "frm_LogSetup.xml"

    Public Const mnu_AppTemp As String = "mnu_AppTemp"
    Public Const frm_ApproveTemp As String = "frm_ApproveTemp"
    Public Const xml_ApproveTemp As String = "frm_ApproveTemp.xml"

    Public Const frm_ChoosefromList As String = "frm_CFL"

    Public Const frm_ItemMaster As String = "150"


    Public Const mnu_DLC_EmailSetUp As String = "z_mnu_DLC_EmailSetUp"
    Public Const xml_DLC_EmailSetUp As String = "frm_DLC_EmailSetUp.xml"
    Public Const frm_DLC_EmailSetUp As String = "frm_DLC_EmailSetUp"

    Public Const frm_DisRule As String = "frm_DisRule"
    Public Const xml_DisRule As String = "frm_DisRule.xml"
End Module
