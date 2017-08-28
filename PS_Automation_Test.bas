Attribute VB_Name = "PS_Automation_Test"
Option Explicit


Const PS_URI_EXPRESS_PO As String = "/psc/ps/EMPLOYEE/ERP/c/MANAGE_PURCHASE_ORDERS.PURCHASE_ORDER_EXP.GBL"

Private Sub SeleniumVer()


    Dim assy As New SeleniumWrapper.Assembly
    
    Debug.Print assy.GetVersion
    

End Sub

Private Sub test_PO_Edit_DueDate()




    Dim user As String, pass As String
    
    user = InputBox("Enter USWIN:", "")
    pass = InputBoxDK("Enter Password:", "")
    
    If user = "" Or pass = "" Then
        MsgBox "Empty user given. Quitting"
        Exit Sub
    End If
    
    
    Dim session As PeopleSoft_Session
    
    session = PeopleSoft_NewSession(user, pass)
    
    
    Dim poChangeOrder As PeopleSoft_PurchaseOrder_ChangeOrder
    
    
    poChangeOrder.PO_ID = "NNYM090797"

    
    'poChangeOrder.PO_DUE_DATE = #6/1/2014#


    

    Dim i As Integer
    
    poChangeOrder.PO_ChangeOrder_ItemCount = 2
    ReDim poChangeOrder.PO_ChangeOrder_Items(1 To poChangeOrder.PO_ChangeOrder_ItemCount) As PeopleSoft_PurchaseOrder_ChangeOrder_Item
  
    poChangeOrder.PO_ChangeOrder_Items(1).PO_Line = 2
    poChangeOrder.PO_ChangeOrder_Items(1).PO_Schedule = 1
    poChangeOrder.PO_ChangeOrder_Items(2).PO_Line = 3
    poChangeOrder.PO_ChangeOrder_Items(2).PO_Schedule = 1
    'poChangeOrder.PO_ChangeOrder_Items(1).SCH_DUE_DATE = #6/2/2014#


    PeopleSoft_PurchaseOrder_ProcessChangeOrder session, poChangeOrder
    
End Sub

Private Sub test_PO_CFQ()




    Dim user As String, pass As String
    
    user = InputBox("Enter USWIN:", "")
    pass = InputBox("Enter Password:", "")
    
    If user = "" Or pass = "" Then
        MsgBox "Empty user given. Quitting"
        Exit Sub
    End If
    
    
    Dim session As PeopleSoft_Session
    
    session = PeopleSoft_NewSession(user, pass)
    
    Dim poCFQ As PeopleSoft_PurchaseOrder_CreateFromQuoteParams
    
    poCFQ.PO_Fields.PO_BUSINESS_UNIT = "NTNYM"
    poCFQ.PO_Fields.PO_HDR_VENDOR_ID = 1399
    poCFQ.PO_Fields.PO_HDR_BUYER_ID = 9736187
    poCFQ.PO_Fields.PO_HDR_APPROVER_ID = 463474
    poCFQ.PO_Fields.PO_HDR_COMMENTS = "Mark PO on all packages"
    poCFQ.PO_Fields.PO_HDR_PO_REF = "PO Test"
    
    poCFQ.PO_Defaults.SCH_DUE_DATE = #4/15/2014#
    poCFQ.PO_Defaults.DIST_BUSINESS_UNIT_PC = "NNWYK"
    poCFQ.PO_Defaults.DIST_PROJECT_CODE = "20140971463"
    poCFQ.PO_Defaults.SCH_SHIPTO_ID = 160231
    poCFQ.PO_Defaults.DIST_ACTIVITY_ID = "EQUIP"
    
    poCFQ.E_QUOTE_NBR = "VZW14N40015"
    
    If False Then
        poCFQ.PO_LineModCount = 0
        'ReDim poCFQ.PO_LineMods(1 To poCFQ.PO_LineModCount) As PeopleSoft_PurchaseOrder_CreateFromQuote_LineModification
        
        poCFQ.PO_LineMods(1).PO_Line = 53
        poCFQ.PO_LineMods(1).PO_LINE_ITEM_ID = "ENG-LTE-CELL"
        poCFQ.PO_LineMods(1).SCH_DUE_DATE = #10/1/2014#
        poCFQ.PO_LineMods(1).SCH_SHIPTO_ID = 145127
        
        poCFQ.PO_LineMods(2).PO_Line = 54
        poCFQ.PO_LineMods(2).PO_LINE_ITEM_ID = "INSTL4322170443"
        poCFQ.PO_LineMods(2).SCH_DUE_DATE = #10/1/2014#
        poCFQ.PO_LineMods(2).SCH_SHIPTO_ID = 145127
    End If
    
    PeopleSoft_PurchaseOrder_CreateFromQuote session, poCFQ
    
    
End Sub

