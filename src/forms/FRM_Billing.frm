VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Billing 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED (Billing)"
   ClientHeight    =   11020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20390
   OleObjectBlob   =   "FRM_Billing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%      %%%%%%%%%%%%%%      %%%%%%%%%%%    %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%    %%%%%%%%%%%%%%     %%%%%%%%%%%%%   %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %%%   %%%        %%%    %%%              %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %%    %%%        %%%    %%%              %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %     %%%        %%%     %%%%%%%%%%%%    %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%    %%%%%%%%%%%%%%      %%%%%%%%%%%%   %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%%   %%%%%%%%%%%%%%               %%%   %%%  ######
'######  @@@        @@@   @@@              @@@         @@     %%%         %%   %%%        %%%               %%%   %%%  ######
'######  @@@        @@@   @@@@@@@@@@@@@@   @@@         @@     %%%         %%   %%%        %%%     %%%%%%%%%%%%    %%%  ######
'######  @@@        @@@   @@@@@@@@@@@@@@   @@@         @@     %%%         %%   %%%        %%%      %%%%%%%%%%     %%%  ######
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################


Private Sub BTN_Add_Customer_Click()
    
    Dim CustCon As New ADODB.Connection
    Dim CustRec As New ADODB.Recordset
    Dim SqlCust As String
    
    'First Need to check whther user is trying to change the existing customer name or adding new customer and number
    SqlCust = "Select * From TBL_Customer_Detail Where TBL_Customer_Detail.[Customer Mobile Number]= '" & TXT_Mobile_Number & "'"
        Call Connection(CustCon)
        CustRec.Open SqlCust, CustCon, adOpenKeyset, adLockOptimistic
        
            With CustRec
                If .RecordCount = 1 Then
                    .Fields("Customer Name") = TXT_CustName.Text
                    .Update
                    LBL_Customer_Status.Caption = "Customer name updated."
                    BTN_Add_Customer.Visible = False
                    
                End If
                
                If .RecordCount > 1 Then
                    .MoveFirst
                    Do Until .EOF
                        If .Fields("Customer Name") = Cbox_Multiple_Cust_Name.Text Then
                            .Fields("Customer Name") = TXT_CustName.Text
                            .Update
                            LBL_Customer_Status = "Customer Name updated and " & Cbox_Multiple_Cust_Name.Text & " removed and new name updated"
                            
                        End If
                        .MoveNext
                    Loop
                End If
                
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("Customer Name") = TXT_CustName.Text
                    .Fields("Customer Mobile Number") = TXT_Mobile_Number.Text
                    .Update
                    LBL_Customer_Status.Caption = "Number Updated in Database"
                    BTN_Add_Customer.Visible = False
                    GoTo FinalTask
                End If
                
            End With
                
                Application.Wait (Now + TimeValue("00:00:05"))
            
                    Cbox_Multiple_Cust_Name.Clear
                    
                    Call TXT_Mobile_Number_Change
FinalTask:
            CustRec.Close
            CustCon.Close
    
    
    
    
    
End Sub

Private Sub BTN_Add_Item_Click()
    LBox_Ordered_Item.Visible = True
    Dim L As Byte   'For List Count in the ListBox
    Dim i As Byte   'For looping inside the list
    Dim TAmnt As Double
    
    
    
    L = LBox_Ordered_Item.ListCount
    
        If Not CBox_BillingType.Text = "" And Not CBox_Dish.Text = "" And Not Cbox_DishCode.Text = "" And _
        Not Cbox_From_Business.Text = "" And Not TXT_CustName.Text = "" And Not TXT_Mobile_Number.Text = "" And Not Cbox_Variant.Text = "" Then
        
            'Adding dish to the listbox (LBox_Ordered)
            LBox_Ordered_Item.AddItem
            LBox_Ordered_Item.List(L, 0) = CBox_Dish.Text
            LBox_Ordered_Item.List(L, 1) = Cbox_Variant.Text
            LBox_Ordered_Item.List(L, 2) = LBL_Qty
            LBox_Ordered_Item.List(L, 3) = LBL_Amount
            LBox_Ordered_Item.List(L, 4) = Cbox_From_Business
            LBox_Ordered_Item.List(L, 5) = LBL_PrintCode

            LBL_Qty = 1
            Cbox_DishCode.Text = ""
            CBox_Dish.Text = ""
            Cbox_Variant.Text = ""
            'Cbox_From_Business.Text = ""
            
            LBL_Amnt_Desc.Visible = True
            LBL_Total_Amnt.Visible = True
                                    
                        For i = 1 To L
                        
                            TAmnt = TAmnt + CLng(LBox_Ordered_Item.List(i, 3))
                            
                        Next i
                        
                        LBL_Total_Amnt = TAmnt
        Else
                
                    MsgBox "Please Fill all the Details before adding the dish", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        End If

    
End Sub

Private Sub BTN_Back_Click()
    
    Me.Hide
    FRM_DashBoard.Show
    
End Sub


Private Sub BTN_Billing_Proceed_Click()
              
                        If Not TXT_Disc = "" And Cbox_Disc_On = "" Then
                            MsgBox "Disc on restaurant can't be blank", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                            
                            Exit Sub
                            
                        End If
              
              
VarDeclare:
                        Dim MenuCon As New ADODB.Connection
                        Dim MenuRec As New ADODB.Recordset
                        Dim SQL As String
                        Dim tSQL_Ala As String    'For  bill_Total table and counting records for Bill no
                        Dim tSQL_Arab As String
                        Dim UserResult As String       'For Billprint second copy
                        
                        Dim ArabianList As New Collection
                        Dim AlaRasiList As New Collection
                        Dim orderlist As clsMenuList
                        Dim OrderOut As clsMenuList
                        
                        Dim Bill1 As New clsPayment
                        Dim Bill2 As New clsPayment
                        
                        Dim MenuDetail As New ADODB.Connection
                        Dim MenuDetailRec As New ADODB.Recordset
                        
                    Select Case CBox_Payment_Mode1.Text
                     Case Is = "" & Not CBox_BillingType.Text = "Online"
                        
                        MsgBox "Payment Mode1 cannot be empty.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                        Exit Sub
                    
                    Case Is = CBox_Payment_Mode2.Text & Not CBox_BillingType = "Online"
                        
                        MsgBox "Payment Mode1 and Payment Mode2 cannot be same", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                        Exit Sub
                    
                    
                    End Select
                                GoTo BillingType
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

BillingType:
                        Select Case CBox_BillingType.Text
    

                                        
                         Case Is = "Online"
                                        
                                        If Not TXT_CustName.Text = "" And Not TXT_Mobile_Number.Text = "" Then
                                            
                                                                GoSub BillNumber
                                                                GoSub orderlist
                                                                GoSub DiscCal
                                                                GoSub GSTCal
                                                                GoSub Payment
                                                                
                                                              
                                                                
'                                                                SQL = "Select * From TBL_Bill_Detailed"
                                                                
                                                                GoSub Record
                                                                
                                                                GoSub KitchenPrint
                                                                GoSub PrintBill
                                                                GoSub ResetForm
                                                                
                                                                MenuRec.Close
                                                                MenuCon.Close
                                                                
                                                                Exit Sub
                                            
                                        
                                        End If

                        Case Is = "Take Away"
                                                        
                                        If Not CBox_BillingType.Text = "" And Not Cbox_From_Business.Text = "" And LBox_Ordered_Item.ListCount > 1 _
                                        And Not TXT_CustName.Text = "" And Not TXT_Mobile_Number.Text = "" And Not CBox_Payment_Mode1 = "" And Not CBox_Payment_Mode1.Text = CBox_Payment_Mode2.Text Then
                                                                
                                                                
                                                                GoSub BillNumber
                                                                GoSub orderlist
                                                                GoSub DiscCal
                                                                GoSub GSTCal
                                                                GoSub Payment
                                                   
                                                                  
                                                                    
                                                                    
                                                        
'                                                                SQL = "Select * From TBL_Bill_Detailed"
                                                                
                                                                
                                                                'Gosub KOTHOLDRecord
                                                                
                                                                GoSub Record 'Final Bill
                                                                GoSub KitchenPrint
                                                                GoSub PrintBill
                                                                GoSub ResetForm
                                                                
                                                                
                                                                
                                                                Exit Sub
                                                        
                                        Else
                                    
                                                MsgBox "Please fill all the details before proceeding.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                                    
                                        End If
                                    
                    
                          Case Is = "DineIn"
                          
                                        If Not TXT_CustName.Text = "" And Not TXT_Mobile_Number.Text = "" Then
                                           
                                           GoSub BillNumber
                                           
                                           tSQL_Ala = "Select * From TBL_KOT_HOLD_Total Where TBL_KOT_HOLD_Total.From = '" & ShSupport.Range("A4").Value & "'"
                                           tSQL_Arab = "Select * From TBL_HOLD_Total Where TBL_KOT_HOLD_Total.From = '" & ShSupport.Range("A5").Value
                                           
                                           
                                           MenuRec.Open tSQL_Ala, MenuCon, adOpenKeyset, adLockOptimistic
                                           
                                           With MenuRec
                                                Bill1.HoldNo = "KOT/HOLD/" & .RecordCount + 1001
                                           End With
                                               
                                           MenuRec.Close
                                           
                                           MenuRec.Open tSQL_Arab, MenuCon, adOpenKeyset, adLockOptimistic
                                           
                                           With MenuRec
                                                Bill2.HoldNo = "KOT/HOLD/" & .RecordCount + 1001
                                           End With
                                           
                                           MenuRec.Close
                                           
                                           'GoSub HoldOrder
                                           
                                           
                                           
                                           MenuCon.Close
                                           
                                           
                                           
                                           
                        
                                        
                                        End If
                          End Select
                                        
                                        Exit Sub
                    
                    
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    

                    

                       
                        
                
BillNumber:
                        SQL = "Select * From TBL_Restaurant_Selection"
                        tSQL_Ala = "Select * From TBL_Bill_Total Where ((TBL_Bill_Total.[From]) = 'AlaRasi/Manna');"
                        tSQL_Arab = "Select * From TBL_Bill_Total Where ((TBL_Bill_Total.[From]) = 'Arabian Hub');"
                        
                        Call Connection(MenuCon)


                
                                    
                
                        'For Bill No.
                        'MenuCon.Close
                        
                        'For bill no of AlaRasi Bill
                        
                                    MenuRec.Open tSQL_Ala, MenuCon, adOpenKeyset, adLockOptimistic
                            
                            
                                                Bill1.BillNo = "AlaRasi/" & MenuRec.RecordCount + 1001 & "/" & Format(Now, "DD/MM/YYYY")
                            
                                    MenuRec.Close
                            
                        'For Bill no of Araboian Bill
                        
                                    MenuRec.Open tSQL_Arab, MenuCon, adOpenKeyset, adLockOptimistic
                            
                                                Bill2.BillNo = "ArabianHub/" & (MenuRec.RecordCount + 1001) & "/" & Format(Now, "DD/MM/YYYY")
                            
                                    MenuRec.Close
                                    
                                    
                            tSQL_Ala = "Select * From TBL_KOT_HOLD_Total Where((TBL_KOT_HOLD_Total.[From])= 'AlaRasi/Manna');"
                            tSQL_Arab = "Select * From TBL_KOT_HOLD_Total Where((TBL_KOT_HOLD_Total.[From])= 'Arabian Hub');"
                            
                                    MenuRec.Open tSQL_Ala, MenuCon, adOpenKeyset, adLockOptimistic
                                                
                                                Bill1.HoldNo = "Manna/" & (MenuRec.RecordCount + 1001)
                                    
                                    MenuRec.Close
                                    
                                    MenuRec.Open tSQL_Arab, MenuCon, adOpenKeyset, adLockOptimistic
                                    
                                                Bill2.HoldNo = "Arabian/" & (MenuRec.RecordCount + 1001)
                                    
                                    MenuRec.Close
                                    MenuCon.Close
                            
                                        Return          'For returning to gosub statement
                                    
                
orderlist:
                        Dim i As Byte 'For Looping inside the orderlist in LBox_Ordered_Item
                        
                        SQL = "Select * From TBL_KOT_HOLD_DETAIL"
                        tSQL_Arab = "Select * From TBL_KOT_HOLD_Total" 'using it for writing in the KOTHold Total table in access
                        Dim AlaCount As Long    'For Order Counting
                        Dim ArabCount As Long   'For Order Counting
                        AlaCount = 0
                        ArabCount = 0
                        
                        
                        
'                        SQL = "Select * From TBL_Restaurant_Selection"
'                        tSQL_Ala = "Select * From TBL_Bill_Total Where ((TBL_Bill_Total.[From]) = 'AlaRasi/Manna');"
'                        tSQL_Arab = "Select * From TBL_Bill_Total Where ((TBL_Bill_Total.[From]) = 'Arabian Hub');"
                
                            If LBox_Ordered_Item.ListCount < 1 Then                 'If there is no Order in The orderlist then to intimate the user
                                MsgBox "Nothing in the List to proceed." & vbNewLine & "Please add the item in the order list to continue.", , _
                                "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                
                                Exit Sub
                
                            End If
                                
                                Call Connection(MenuCon)
                                MenuRec.Open tSQL_Arab, MenuCon, adOpenKeyset, adLockOptimistic 'Total Table in access
                                MenuDetailRec.Open SQL, MenuCon, adOpenKeyset, adLockOptimistic 'Detailed table in access
                                

                
                
                       For i = 1 To LBox_Ordered_Item.ListCount - 1
                                'order from Arabian will be added in ArabianList collection and hold table in access
                            If LBox_Ordered_Item.List(i, 4) = "Arabian" Then
                                ArabCount = ArabCount + 1
                                
                                Set orderlist = New clsMenuList
                
                                orderlist.Dish = LBox_Ordered_Item.List(i, 0)
                                orderlist.DishVariant = LBox_Ordered_Item.List(i, 1)
                                orderlist.Qty = LBox_Ordered_Item.List(i, 2)
                                orderlist.Amount = LBox_Ordered_Item.List(i, 3)
                                orderlist.BFrom = LBox_Ordered_Item.List(i, 4)
                                orderlist.PrintCode = LBox_Ordered_Item.List(i, 5)
                                
                                Bill2.TAmnt = Bill2.TAmnt + orderlist.Amount
                                
                                With MenuDetailRec
                                    .AddNew
                                    .Fields("KOT Hold Number") = Bill2.HoldNo
                                    .Fields("Item") = orderlist.Dish
                                    .Fields("QTY") = orderlist.Qty
                                    .Fields("Amount") = orderlist.Amount
                                    .Fields("From") = orderlist.BFrom
                                    .Update
                                End With
                                
                                
                                
                                ArabianList.Add orderlist
                
                            End If
                            
                                'order from manna or alarasi will be added in AlaRasiList Collection
                            If LBox_Ordered_Item.List(i, 4) = "AlaRasi" Or LBox_Ordered_Item.List(i, 4) = "Manna" Then
                                AlaCount = AlaCount + 1
                                Set orderlist = New clsMenuList
                
                                orderlist.Dish = LBox_Ordered_Item.List(i, 0)
                                orderlist.DishVariant = LBox_Ordered_Item.List(i, 1)
                                orderlist.Qty = LBox_Ordered_Item.List(i, 2)
                                orderlist.Amount = LBox_Ordered_Item.List(i, 3)
                                orderlist.BFrom = LBox_Ordered_Item.List(i, 4)
                                orderlist.PrintCode = LBox_Ordered_Item.List(i, 5)
                                
                                Bill1.TAmnt = Bill1.TAmnt + orderlist.Amount
                                
                                With MenuDetailRec
                                    .AddNew
                                    .Fields("KOT Hold Number") = Bill1.HoldNo
                                    .Fields("Item") = orderlist.Dish
                                    .Fields("QTY") = orderlist.Qty
                                    .Fields("Amount") = orderlist.Amount
                                    .Fields("From") = orderlist.BFrom
                                    .Update
                                End With
                                
                                
                                
                                AlaRasiList.Add orderlist
                
                
                            End If
                        
                
                       Next i
                            
                            If ArabCount > 0 Then
                                With MenuRec
                                  .AddNew
                                  .Fields("KOT Hold Number") = Bill2.HoldNo
                                  .Fields("Total Amount") = Bill2.TAmnt
                                  .Fields("UserId") = ShSupport.Range("I5").Value
                                  .Fields("From") = "Arabian"
                                  .Fields("Mobile No") = TXT_Mobile_Number.Text
                                  .Update
                                  
                                End With
                                
                            End If
                            
                            If AlaCount > 0 Then
                                With MenuRec
                                    .AddNew
                                    .Fields("KOT Hold Number") = Bill1.HoldNo
                                    .Fields("Total Amount") = Bill1.TAmnt
                                    .Fields("UserId") = ShSupport.Range("I5").Value
                                    .Fields("From") = "Manna"
                                    .Fields("Mobile No") = TXT_Mobile_Number.Text
                                    .Update
                                
                                End With
                                
                            End If
                            
                            MenuRec.Close
                            MenuDetailRec.Close
                            
                    Return          'For Returning back to the go sub statement
                    
                                    
DiscCal:
                                    
                                    'for calculation of the disc
                                        Bill1.DiscAmnt = 0
                                        Bill1.GST = 0
                                        
                                    Select Case LBL_Disc
                                        Case Is = ShMenu.Range("G2").Value 'Disc in Percent
                
                                                    
                                                    'Disc on Manna or both then the calculation of the disc
                                                    
                                            If Cbox_Disc_On.Text = ShSupport.Range("A5").Value Or Cbox_Disc_On.Text = "Both" Then
                                                Bill1.DiscPercent = TXT_Disc.Text & ShMenu.Range("G2").Value
                                                Bill1.DiscAmnt = (Bill1.TAmnt * CLng(TXT_Disc.Text)) / 100
                                            End If
                
                                        Case Is = ShMenu.Range("G3").Value 'Disc in Amount
                
                                            If Cbox_Disc_On.Text = ShSupport.Range("A5").Value Or Cbox_Disc_On.Text = "Both" Then 'Disc on Manna or both then
                                                Bill1.DiscPercent = (CLng(TXT_Disc.Text) * 100) / Bill1.TAmnt
                                                Bill1.DiscAmnt = CLng(TXT_Disc.Text)
                
                                            End If
                
                                     End Select
                                     
                                    
                                    'For calculation of the disc Bill2
                                        Bill2.DiscAmnt = 0
                                        Bill2.GST = 0
                                        
                                        
                                    Select Case LBL_Disc
                                    
                                        Case Is = ShMenu.Range("G2").Value      'Disc in Percent
                                        
                                        
                                            'Disc on Arabian or Both then the calculation of the disc
                                            
                                            If Cbox_Disc_On.Text = ShSupport.Range("A4").Value Or Cbox_Disc_On.Text = "Both" Then
                                                Bill2.DiscPercent = TXT_Disc.Text & ShMenu.Range("G2").Value
                                                Bill2.DiscAmnt = (Bill2.TAmnt * CLng(TXT_Disc.Text)) / 100
                                            
                                            End If
                                        
                                        Case Is = ShMenu.Range("G3").Value      'Disc in Amount
                                        
                                            
                                            If Cbox_Disc_On.Text = ShSupport.Range("A4").Value Or Cbox_Disc_On.Text = "Both" Then
                                                
                                                Bill2.DiscPercent = (CLng(TXT_Disc.Text) * 100) / Bill2.TAmnt
                                                Bill2.DiscAmnt = CLng(TXT_Disc.Text)
                                                
                                            End If
                                        
                                        End Select
                                     
                                     
                                     
                                        Return              'For Returning back to the go sub statement
                
GSTCal:

                                     'Calculation of Gst bill1
                
                                     Select Case UCase(ShSupport.Range("I11").Value) 'GST Applicable or not if yes then CGST and SGST is 9% each total of 18% GST
                
                                        Case Is = "Y"
                                            Bill1.GST = ((Bill1.TAmnt - Bill1.DiscAmnt) * 18) / 100
                
                                        Case Is = "N"
                                            Bill1.GST = 0
                
                                     End Select
                                     
                                     Bill1.CalAmnt = (Bill1.TAmnt - Bill1.DiscAmnt) + Bill1.GST
                                     
                                     
                                    'Calculation of GST bill2
                                    
                                    Select Case UCase(ShSupport.Range("I11").Value)       'GST Applicable or not if yes then CGST and SGST is 9% each total of 18% GST
                                        
                                        Case Is = "Y"
                                            Bill2.GST = ((Bill2.TAmnt - Bill2.DiscAmnt) * 18) / 100
                                            
                                        Case Is = "N"
                                            Bill2.GST = 0
                                        
                                        End Select
                                        
                                        Bill2.CalAmnt = (Bill2.TAmnt - Bill2.DiscAmnt) + Bill2.GST

                                        Return      'For Returning back to the go sub statement
                                            
Payment:
                                            If TXT_Payment2.Text = "" Then
                                                TXT_Payment2.Text = 0
                                                Bill2.PaidAmnt = 0
                                            End If
                                            
                                            If IsNumeric(TXT_Payment1.Text) = True Then
                                                Bill1.PaidAmnt = CDbl(TXT_Payment1.Text)
                                            ElseIf CBox_BillingType = "Online" And TXT_Payment1.Text = "" Then
                                                                                        
                                                Bill1.PaidAmnt = 0
                                                TXT_Payment1.Text = 0
                                            
                                            End If
                                                
                
                                    If (Bill1.CalAmnt + Bill2.CalAmnt) > (CLng(TXT_Payment1.Text) + CLng(TXT_Payment2.Text)) And Not CBox_BillingType = "Online" Then
                                        MsgBox "The total amount paid is less than the due amount"
                                        Exit Sub
                                    ElseIf (Bill1.CalAmnt + Bill2.CalAmnt) < (CLng(TXT_Payment1.Text) + CLng(TXT_Payment2.Text)) And Not CBox_BillingType.Text = "Online" Then
                                        MsgBox "Balance money is-" & ((CLng(TXT_Payment1.Text) + (CLng(TXT_Payment2.Text)) - (Bill1.CalAmnt + Bill2.CalAmnt))), , _
                                        "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                                        
                                    End If
                                    
                                    
                                    Return
                
                
                
                

Record:
                            

                                    
                                    'MenuRec.Close           '"Select * From TBL_KOT_HOLD_Total Where((TBL_KOT_HOLD_Total.[From])= 'Arabian Hub');"
                                    
                                    
                                    
                                    Call Connection(MenuDetail)
                                    SQL = "Select * From TBL_Bill_Detailed"

                                    MenuDetailRec.Open SQL, MenuDetail, adOpenKeyset, adLockOptimistic
'
                                    SQL = "Select * From TBL_KOT_HOLD_DETAIL Where TBL_KOT_HOLD_DETAIL.[KOT Hold Number]= '" & Bill1.HoldNo & "'"

                                    MenuRec.Open SQL, MenuDetail, adOpenKeyset, adLockOptimistic
                                    
                                    
                                    
                                   Call HeaderPrint
                                        
                                        

                                    
                                    Dim CustName As String  'For Customenr Name
                                    Dim Mnumber As String   'For Mobile Number
                                    
                                    CustName = TXT_CustName.Text
                                    Mnumber = TXT_Mobile_Number.Text
                                    
                                    Call CustName_MNumber(CustName, Mnumber)
                                    

                                    Dim HeaderRow As Long   'for Header Row of the bill
                                    HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row
                                    

                                    
                                    
                                    
                                
                                    
                                    
                             'Dim OrderOut As clsMenuList 'for copying the data in the loop
                
                             'If the billing type is TakeAway or Online then
                                     
                               
                
                            If Not AlaRasiList.Count = 0 Then
                                    For i = 1 To AlaRasiList.Count
                
                                        Set OrderOut = New clsMenuList
                                        Set OrderOut = AlaRasiList.Item(i)
                                        'Bill1.TAmnt = CLng(Bill1.TAmnt) + CLng(OrderOut.Amount)        'bill1 amnt is calculated in alarasi collection itself
                                        ShPrint.Cells(HeaderRow + i, 1).Value = OrderOut.PrintCode
                                        ShPrint.Cells(HeaderRow + i, 2).Value = OrderOut.Qty
                                        ShPrint.Cells(HeaderRow + i, 3).Value = OrderOut.Amount
                                            If Not CBox_BillingType = "DineIn" Then
                                                    MenuDetailRec.AddNew
                                                    MenuDetailRec.Fields("Food Item") = OrderOut.Dish & " " & orderlist.DishVariant
                                                    MenuDetailRec.Fields("Amount") = OrderOut.Amount
                                                    MenuDetailRec.Fields("Bill Number") = Bill1.BillNo
                                                    MenuDetailRec.Fields("From") = "Manna"
                                                    MenuDetailRec.Update
                                            End If
                                       
                                        
                                       
                                           
                                        
                                        
                                    Next i
                                    
                                     If CBox_BillingType = "Online" Or CBox_BillingType = "Take Away" Then
                                         With MenuRec
                                            .MoveFirst
                                            
                                            Do Until .EOF
                                                .Fields("Table Number") = "N/A"
                                                .Update
                                            Loop
                                            
                                         End With
                                            
                                        Else
                                            Dim TableNo As Byte
                                            
                                        
                                        End If
                                MenuRec.Close
                                        
                                       
                                    'Writing in TBL_Bill_Total Table (bill1) in Accesss

                                    
                                    
                                    
                                    
                                    SQL = "Select * From TBL_Bill_Total"
                                    
                                
                
                                    MenuRec.Open SQL, MenuCon, adOpenKeyset, adLockOptimistic
                                                    
                            
                                
                                    If CBox_BillingType.Text = "Online" Or CBox_BillingType = "Take Away" Then
                                    
                                        ShPrint.Range("A3").Value = Bill1.BillNo
                                        
                                    ElseIf CBox_BillingType.Text = "DineIn" Then
                                        ShPrint.Range("A3").Value = Bill1.HoldNo
                                        
                                    End If
                                
                                'Sql = "Select * From TBL_Bill_Detailed"
                                    ShPrint.Range("A2", "C2").Merge
                                    ShPrint.Range("A2", "C2").RowHeight = 60
                                    ShPrint.Range("A2", "C2").VerticalAlignment = xlCenter
                                    ShPrint.Range("A2", "C2").HorizontalAlignment = xlCenter
                                    ShPrint.Range("A2", "C2").Font.Size = 12
                                    ShPrint.Range("A2").Value = "Manna South" & Chr(10) & "Indian And Chinese"
                                    ShPrint.Range("A1", "A3").Font.Bold = True
                                    HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                
                             
                                
                
                                        
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                '                        With ShPrint.Range(Cells(HeaderRow, 3)).Borders(xlEdgeTop)
                '                            .LineStyle = xlContinuous0
                '                            .ColorIndex = xlAutomatic
                '                            .TintAndShade = 0
                '                            .Weight = xlThin
                '                        End With
                                        
                                      
                                            
                                            'Only if disc is there then only it will print disc Amount
                                        If Not Bill1.DiscAmnt = 0 Then
                                            ShPrint.Cells(HeaderRow, 1).Value = "Disc"
                                            ShPrint.Cells(HeaderRow, 3).Value = Bill1.DiscAmnt
                                        End If
                                        
                                        
                                        'New row
                                        
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                        'gst on print bill
                                        
                                        If Not Bill1.GST = 0 Then
                                            ShPrint.Cells(HeaderRow, 1).Value = "CGST"
                                            ShPrint.Cells(HeaderRow, 3).Value = Bill1.GST / 2
                                            ShPrint.Cells(HeaderRow + 1, 1).Value = "SGST"
                                            ShPrint.Cells(HeaderRow + 1, 3).Value = Bill1.GST / 2
                                            
                                        Else
                                            
                                            ShPrint.Cells(HeaderRow, 1).Value = "CGST"
                                            ShPrint.Cells(HeaderRow, 3).Value = 0
                                            ShPrint.Cells(HeaderRow + 1, 1).Value = "SGST"
                                            ShPrint.Cells(HeaderRow + 1, 3).Value = 0
                                                                
                                        End If
                                            
                                            ShPrint.Cells(HeaderRow + 2, 1).Value = "Manna Total"
                                            ShPrint.Cells(HeaderRow + 2, 3).Value = Bill1.CalAmnt
                                                
                                                'For Bill1 Entry
                                    If Not CBox_BillingType = "Dine" Then
                                         With MenuRec
                                            .AddNew
                                            .Fields("Bill Number") = Bill1.BillNo
                                            .Fields("KOT Hold Number") = Bill1.HoldNo
                                            .Fields("Disc Amount") = Bill1.DiscAmnt
                                            .Fields("Disc %") = Bill1.DiscPercent
                                            .Fields("GST") = Bill1.GST
                                            .Fields("Total Amount") = Bill1.CalAmnt
                                            .Fields("UserId") = ShSupport.Range("I5").Value
                                            .Fields("From") = "AlaRasi/Manna"
                                            .Fields("Mobile No") = TXT_Mobile_Number.Text
                                            .Fields(CBox_Payment_Mode1.Text) = TXT_Payment1.Text
                                            
                                                If CBox_BillingType = "Online" Then
                                                
                                                    .Fields(CBox_Payment_Mode1.Text) = TXT_Payment1.Text
                                                    .Fields("Reference") = TXT_CustName.Text
                                                    
                                                End If
                                        
                                        If Not CBox_Payment_Mode2.Text = "" And Not CBox_Payment_Mode2.Text = CBox_Payment_Mode1.Text Then
                                            .Fields(CBox_Payment_Mode2.Text) = TXT_Payment2.Text
                                        
                                        ElseIf Not CBox_Payment_Mode2.Text = "" And CBox_Payment_Mode2.Text = CBox_Payment_Mode2.Text Then
                                            .Fields(Cbox_Variant.Text) = CLng(TXT_Payment1.Text) + CLng(TXT_Payment2.Text)
                                        
                                        End If
                                                
                                               
                                            
                                            .Update
                                            
                                        End With
                                        
                                        End If
                
                    
                            End If
                
                                
                
                            If Not ArabianList.Count = 0 Then
                                    
                                    
                                    HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                    
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Merge
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).HorizontalAlignment = xlCenter
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).VerticalAlignment = xlCenter
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Font.Size = 18
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Font.Bold = True
                                    ShPrint.Cells(HeaderRow, 1).Value = "Arabian Hub"
                                    ShPrint.Range(ShPrint.Cells(HeaderRow + 1, 1), ShPrint.Cells(HeaderRow + 1, 3)).Merge
                                    ShPrint.Cells(HeaderRow + 1, 1).Value = Bill2.BillNo
                                    ShPrint.Cells(HeaderRow + 1, 1).Font.Size = 12
                                    ShPrint.Cells(HeaderRow + 1, 1).Font.Bold = True
                                    ShPrint.Range(ShPrint.Cells(HeaderRow + 2, 1), ShPrint.Cells(HeaderRow + 2, 3)).Merge
                                    ShPrint.Range(ShPrint.Cells(HeaderRow + 2, 1), ShPrint.Cells(HeaderRow + 2, 3)).Font.Size = 12
                                    ShPrint.Cells(HeaderRow + 2, 1).Value = TXT_CustName.Text & "-" & TXT_Mobile_Number.Text
                                    HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row
                                    
                                    
                                    
                                    
                                    For i = 1 To ArabianList.Count
                                        Set OrderOut = New clsMenuList
                                        Set OrderOut = ArabianList.Item(i)
                                        'Bill2.TAmnt = CLng(Bill2.TAmnt) + CLng(OrderOut.Amount)        'Bill2 tamount is calculated in arabian list
                                        ShPrint.Cells(HeaderRow + i, 1) = OrderOut.PrintCode
                                        ShPrint.Cells(HeaderRow + i, 2) = OrderOut.Qty
                                        ShPrint.Cells(HeaderRow + i, 3) = OrderOut.Amount
                                        MenuDetailRec.AddNew
                                        MenuDetailRec.Fields("Food Item") = OrderOut.Dish & " " & orderlist.DishVariant
                                        MenuDetailRec.Fields("Amount") = OrderOut.Amount
                                        MenuDetailRec.Fields("Bill Number") = Bill2.BillNo
                                        MenuDetailRec.Fields("From") = "Arabian Hub"
                                        MenuDetailRec.Update
                                        
                                        'need to write on access
                
                                    Next i
                                                
                                    'Wwriting in TBL_Bill_Total Table
                                    
                                        
                                        
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                        ShPrint.Cells(HeaderRow, 1).Value = "Total"
                                        ShPrint.Cells(HeaderRow, 3).Value = Bill2.TAmnt
                                        
                                            'Only if Disc is there then only the disc will print
                                        If Not Bill2.DiscAmnt = 0 Then
                                            ShPrint.Cells(HeaderRow + 1, 1).Value = "Disc"
                                            ShPrint.Cells(HeaderRow + 1, 3).Value = Bill2.DiscAmnt
                                            
                                        End If
                                        
                                        'new header row
                                        
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                        'gst on print bill
                                        
                                        If Not Bill2.GST = 0 Then
                                            ShPrint.Cells(HeaderRow, 1).Value = "CGST"
                                            ShPrint.Cells(HeaderRow, 3).Value = Bill2.GST / 2
                                            ShPrint.Cells(HeaderRow + 1, 1).Value = "SGST"
                                            ShPrint.Cells(HeaderRow + 1, 3).Value = Bill2.GST / 2
                                            
                                        Else
                                            ShPrint.Cells(HeaderRow, 1).Value = "CGST"
                                            ShPrint.Cells(HeaderRow, 3).Value = 0
                                            ShPrint.Cells(HeaderRow + 1, 1).Value = "SGST"
                                            ShPrint.Cells(HeaderRow + 1, 3).Value = 0
                                            
                                        
                                        End If
                                        
                                            ShPrint.Cells(HeaderRow + 2, 1).Value = " Arabian Total"
                                            ShPrint.Cells(HeaderRow + 2, 3).Value = Bill2.CalAmnt
                                        
                                                'For Bill2Entry
                                        
                                        With MenuDetailRec     'menurec is writing in the hold table need to check
                                            .AddNew
                                            .Fields("Bill Number") = Bill2.BillNo
                                            .Fields("KOT Hold Number") = Bill2.HoldNo
                                            .Fields("Disc %") = Bill2.DiscPercent & "%"
                                            .Fields("Disc Amount") = Bill2.DiscAmnt
                                            .Fields("GST") = Bill2.GST
                                            .Fields("Total Amount") = Bill2.CalAmnt
                                            .Fields("UserId") = ShSupport.Range("I5").Value
                                            .Fields("From") = "Arabian Hub"
                                            .Fields("Mobile No") = TXT_Mobile_Number.Text
                                                    If Not CBox_BillingType = "Online" Then
                                                        .Fields(CBox_Payment_Mode1.Text) = TXT_Payment1.Text
                                                        .Fields("Reference") = TXT_CustName.Text
                                                    End If
                                        If Not CBox_Payment_Mode2.Text = "" And Not CBox_Payment_Mode2.Text = CBox_Payment_Mode1.Text Then
                                            .Fields(CBox_Payment_Mode2.Text) = TXT_Payment2.Text
                                        
                                        ElseIf Not CBox_Payment_Mode2.Text = "" And CBox_Payment_Mode2.Text = CBox_Payment_Mode2.Text Then
                                            .Fields(Cbox_Variant.Text) = CLng(TXT_Payment1.Text) + CLng(TXT_Payment2.Text)
                                        
                                        End If
                                            .Update
                                        
                                        End With
                                        
                                        
                                        
                                        
                                        
                                            
                
                
                            End If
                            
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
                                        
                                        ShPrint.Range("A" & HeaderRow).Value = "G Total"
                                        ShPrint.Range("C" & HeaderRow).Value = Bill1.CalAmnt + Bill2.CalAmnt
                                        
                                        
                                        MenuDetailRec.Close
                                        MenuDetail.Close
                                        MenuRec.Close
                                        MenuCon.Close
                                        
                                        
                                        
                                            
                                            Return      'For Returning back to the go sub statement
                                            
PrintBill:
                                        
                                        ShPrint.Range("A1").ColumnWidth = 17
                                        ShPrint.Range("B1").ColumnWidth = 4
                                        ShPrint.Range("C1").ColumnWidth = 7.5
                                        
                                        HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row       'For Finding the last row headerrow is used again
BillRepeat:
                                        ShPrint.Range("A1", "C" & HeaderRow).PrintOut
                                        
                                        UserResult = MsgBox("Do you want print bill again", vbOKCancel, "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED")
                                            If UserResult = vbOK Then
                                                GoTo BillRepeat
                                            End If
                                        
                                        
                                        Return      'For returning to the gosub statement
                                        
                                        
                                        
KitchenPrint:
                                        
                                        HeaderRow = ShPrint.Range("B" & Rows.Count).End(xlUp).Row       'For Finding the last row headerrow is used again
                                        ShPrint.Range("C:C").EntireColumn.Hidden = True
                                        ShPrint.Range("A1").ColumnWidth = 17 + 3.5
                                        ShPrint.Range("A1").EntireRow.Hidden = True
                                        ShPrint.Range("B1").ColumnWidth = 4 + 3
KOTRepeat:
                                        ShPrint.Range("A1", "B" & HeaderRow).PrintOut
                                        
                                        UserResult = MsgBox("Do you want to print KOT again", vbOKCancel, "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED")
                                        If UserResult = vbOK Then
                                            GoTo KOTRepeat
                                        End If
                                            
                                        ShPrint.Range("A1").EntireRow.Hidden = False
                                        
                                        Return      'For returning to the gosub statement
                                        

                                        
                                        
                                        
                                        
                                        
                                        
ResetForm:
                                            
                                            
                                        ShPrint.Range("A:C").Delete
                                        CBox_BillingType.Text = ""
                                        Cbox_Disc_On.Text = ""
                                        CBox_Dish.Text = ""
                                        Cbox_DishCode.Text = ""
                                        Cbox_From_Business.Text = ""
                                        Cbox_Multiple_Cust_Name.Text = ""
                                        CBox_Payment_Mode1.Text = ""
                                        CBox_Payment_Mode2.Text = ""
                                        Cbox_Variant.Text = ""
                                        TXT_CustName.Text = ""
                                        TXT_Disc.Text = ""
                                        TXT_Mobile_Number.Text = ""
                                        TXT_Payment1.Text = ""
                                        TXT_Payment2.Text = ""
                                        LBox_Ordered_Item.Clear
                                        LBL_PrintCode = ""
                                        
                                                                                                    
                                
                    
                    'MenuDetail.Close
                    'MenuRec.Close
                    
                    
                    Return              'For Return to gosub statement
                
                
                '            Select Case CBox_BillingType
                '                Case Is = "Online", "Take Away"
                '
                '
                '
                '            End Select
                
                
                
                
   
End Sub



Private Sub BTN_Hold_Duplicate_Click()
    If ShSupport.Range("I2").Value = "Admin" Then
    
        Dim RestSelection As String
        
        Dim UserResult As String
        UserResult = InputBox("Enter the KOT number", "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED")
        
        Dim SQL As String
        Dim Con As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        SQL = "Select * From TBL_KOT_HOLD_Detail Where TBL_KOT_HOLD_Detail.[KOT Hold Number]= '" & UserResult & "'"
        Call Connection(Con)
        
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
        With Rec
            If Not .RecordCount = 0 Then
                .MoveFirst
                Dim Dup As New clsPayment
                Dim Rest As String
                
                Dup.HoldNo = .Fields("KOT Hold Number")
                Rest = .Fields("From")
                
                
                
            Else
            
                MsgBox "Unable to find the KOT detail. Please check the KOT number and try again", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                
                GoTo ExitStatement
            End If
        End With
        
        Rec.Close
        
    
        Rec.Open SQL, Con, adOpenKeyset, adLockOptimistic
        
                                        Dim HeaderRow As Long
                                        HeaderRow = ShPrint.Range("B" & Rows.Count).End(xlUp).Row
                                        ShPrint.Range("C:C").EntireColumn.Hidden = True
                                        ShPrint.Range("A1").ColumnWidth = 17 + 3.5
                                        ShPrint.Range("A1").EntireRow.Hidden = False
                                        ShPrint.Range("B1").ColumnWidth = 4 + 3
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Merge
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).HorizontalAlignment = xlCenter
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).VerticalAlignment = xlCenter
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Font.Size = 18
                                    ShPrint.Range(ShPrint.Cells(HeaderRow, 1), ShPrint.Cells(HeaderRow, 3)).Font.Bold = True
                                    ShPrint.Cells(HeaderRow, 1).Value = Rest
                                    
                                    HeaderRow = ShPrint.Range("A" & Rows.Count).End(xlUp).Row + 1
        
        With Rec
            .MoveFirst
            Do Until .EOF
                ShPrint.Cells(HeaderRow, 1) = .Fields("Item")
                ShPrint.Cells(HeaderRow, 2) = .Fields("QTY")
                HeaderRow = HeaderRow + 1
                .MoveNext
            Loop
            
            
        End With
                                        
                                        ShPrint.Range("A1", Cells(HeaderRow - 1, 2)).PrintOut
                                        ShPrint.Range("A1", Cells(HeaderRow - 1, 2)).CurrentRegion.Clear
                                        ShPrint.Range("A1", Cells(HeaderRow - 1, 2)).RowHeight = 15
                                        ShPrint.Range("A1", Cells(HeaderRow - 1, 2)).ColumnWidth = 8.43
                                        ShPrint.Range("C:C").EntireColumn.Hidden = False
                                        GoTo ExitStatement
    Else
    
        MsgBox "You are not authorise for this action.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        
        
    End If
    
    Exit Sub
    
ExitStatement:
        Rec.Close
        Con.Close
End Sub

Private Sub BTN_Print_Dish_Code_Click()
 Call DishCode_Print
 Dim R As Long
 R = ShPrint.Range("A" & Rows.Count).End(xlUp).Row
 ShPrint.Range(ShPrint.Cells(1, 1), ShPrint.Cells(R, 3)).PrintOut
 ShPrint.Range(ShPrint.Cells(1, 1), ShPrint.Cells(R, 3)).Clear
 
 
End Sub

Private Sub CBox_BillingType_Change()
'        Dim R As Range
'        Dim C As Range
    Dim Con As New ADODB.Connection
    Dim MenuRec As New ADODB.Recordset
    Dim MenuSql As String
    
      Select Case CBox_BillingType.Text
        
            Case Is = ShMenu.Range("D1").Value    'Online Order
                LBL_Mobile_No = "ID no"
                LBL_CustomerName = "Order From"
                Cbox_Multiple_Cust_Name.Clear
                Cbox_From_Business.Clear
                Cbox_From_Business.Visible = True
                Cbox_Multiple_Cust_Name.Visible = True
                LBL_CustomerName.Visible = True
                TXT_CustName.Visible = True
                LBL_Mobile_No.Visible = True
                TXT_Mobile_Number.Visible = True
                Cbox_DishCode.Visible = True
                LBL_Business_From.Visible = True
                LBL_Dish_Code.Visible = True
                LBL_Rupee_Symbol.Visible = True
                Lbl_Payment1.Visible = False
                CBox_Payment_Mode1.Visible = False
                TXT_Payment1.Visible = False
                Lbl_Payment2.Visible = False
                CBox_Payment_Mode2.Visible = False
                TXT_Payment2.Visible = False
                Lbl_Disc_on.Visible = True
                Cbox_Disc_On.Visible = True
                LBL_Disc_Title.Visible = True
                LBL_Disc.Visible = True
                TXT_Disc.Visible = True
                BTN_Add_Customer.Visible = False
                LBL_Disc = ShMenu.Range("G2").Value
        
                
                MenuSql = "Select * From TBL_Online_Billing_Type"
                    Call Menu_Connection(Con)
                MenuRec.Open MenuSql, Con, adOpenKeyset, adLockOptimistic
                
                'For Adding online business mode (Swiggy, Zomato, Paytm) in future if more option is there then just need to add in the table
                    With MenuRec
                    .MoveFirst
                        Do Until .EOF
                            Cbox_Multiple_Cust_Name.AddItem .Fields("Billing type")
                            .MoveNext
                        Loop
                    End With
                
                
                'Set R = ShMenu.ListObjects("Online_Order").DataBodyRange
                    'For Each C In R
        '                Cbox_Multiple_Cust_Name.AddItem C.Value
                    'Next C
                    
                MenuRec.Close
                 
                MenuSql = "Select * From TBL_Restaurant_Selection"
                    
                MenuRec.Open MenuSql, Con, adOpenKeyset, adLockOptimistic
                
                    With MenuRec
                        .MoveFirst
                            Do Until .EOF
                                Cbox_From_Business.AddItem .Fields("Restaurant name")
                                .MoveNext
                            Loop
                    
                    End With
                    
                    MenuRec.Close
                    Con.Close
                
'                Set R = ShSupport.ListObjects("Income_From").DataBodyRange
'                For Each C In R
'
'                Cbox_From_Business.AddItem C.Value
'
'                Next C
                
        
            Case Is = ShMenu.Range("E1").Value, ShMenu.Range("F1").Value          'TakeAway and DineIn
                        Cbox_DishCode.Visible = True
                        Cbox_From_Business.Clear
                        Cbox_From_Business.Visible = True
                        LBL_CustomerName.Visible = True
                        TXT_CustName.Visible = True
                        LBL_Mobile_No.Visible = True
                        TXT_Mobile_Number.Visible = True
                        
                        
                        LBL_Mobile_No = "Mobile Number"
                        LBL_CustomerName = "Customer Name"
                        Cbox_Multiple_Cust_Name.Clear
                        Cbox_From_Business.Clear
                        Cbox_Multiple_Cust_Name.Visible = False
                        LBL_Business_From.Visible = True
                        LBL_Dish_Code.Visible = True
                        LBL_Rupee_Symbol.Visible = True
                        TXT_Disc.Visible = True
                        Lbl_Payment1.Visible = True
                        CBox_Payment_Mode1.Visible = True
                        TXT_Payment1.Visible = True
                        Lbl_Payment2.Visible = True
                        CBox_Payment_Mode2.Visible = True
                        TXT_Payment2.Visible = True
                        LBL_Disc.Visible = True
                        Lbl_Disc_on.Visible = True
                        Cbox_Disc_On.Visible = True
                        LBL_Disc_Title.Visible = True
                        BTN_Add_Customer.Visible = False
                        
                        Call Menu_Connection(Con)
                        
                        MenuSql = "Select * From TBL_Restaurant_Selection"
                    
                                MenuRec.Open MenuSql, Con, adOpenKeyset, adLockOptimistic
                                
                                    With MenuRec
                                        .MoveFirst
                                            Do Until .EOF
                                                Cbox_From_Business.AddItem .Fields("Restaurant name")
                                                .MoveNext
                                            Loop
                                    
                                    End With
                                    
                                    MenuRec.Close
                                    Con.Close
                        
                        
'                        Set R = ShSupport.ListObjects("Income_From").DataBodyRange
'                        For Each C In R
'
'                        Cbox_From_Business.AddItem C.Value
'
'                        Next C
        
            Case Else 'If CBox_BillingType = "" Then
                        LBox_Ordered_Item.Visible = False
                        Cbox_From_Business.Visible = False
                        Cbox_From_Business.Visible = False
                        Cbox_DishCode.Visible = False
                        Cbox_Multiple_Cust_Name.Visible = False
                        TXT_CustName.Text = ""
                        TXT_Mobile_Number.Text = ""
                        LBL_CustomerName.Visible = False
                        TXT_CustName.Visible = False
                        LBL_Mobile_No.Visible = False
                        TXT_Mobile_Number.Visible = False
                        LBL_Business_From.Visible = True
                        LBL_Dish_Code.Visible = False
                        TXT_Disc.Visible = False
                        LBL_Business_From.Visible = False
                        LBL_Disc.Visible = False
                        LBL_Disc_Title.Visible = False
                        Lbl_Disc_on.Visible = False
                        Cbox_Disc_On.Visible = False
                        Lbl_Payment1.Visible = False
                        CBox_Payment_Mode1.Visible = False
                        TXT_Payment1.Visible = False
                        Lbl_Payment2.Visible = False
                        CBox_Payment_Mode2.Visible = False
                        TXT_Payment2.Visible = False
                        LBL_Disc.Visible = False
                        
                        
                    
                        
                    
        End Select
End Sub



Private Sub CBox_Dish_Change()
    Dim MenuCon As New ADODB.Connection
    Dim MenuRec As New ADODB.Recordset
    Dim SQL As String
    Dim DishVar As String
    Dim Vartype As String
    
    
    
    If Not CBox_Dish.Text = "" Then
    
    SQL = "Select * From TBL_" & Cbox_From_Business.Text & "_Dish_Variant Where TBL_" & Cbox_From_Business.Text & _
            "_Dish_Variant.[Dish Cat]= '" & CBox_Dish.Text & "'"
    
    
        Call Menu_Connection(MenuCon)
        
        MenuRec.Open SQL, MenuCon, adOpenKeyset, adLockOptimistic
                        
                        With MenuRec
                            If .RecordCount = 1 Then
                                .MoveFirst
                                Vartype = .Fields("Variant Type")   'Varient type selecting from the table
                                
                            ElseIf .RecordCount = 0 Then
                            
                                MsgBox "No Record Found, Please Contact Sam", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                                MenuRec.Close
                                MenuCon.Close
                                Exit Sub
                            
                            ElseIf .RecordCount > 1 Then
                                
                                .MoveFirst
                                
                            
                                
                                
                            End If
                                                    
                        End With
                        
        MenuRec.Close
        
        'For Selection of Variant
'
'        If Vartype = "No Variant" Then
'            Cbox_Variant.Visible = True
'            Cbox_Variant.Text = Vartype
'
'        Else
                  
                SQL = "Select * From TBL_Variant_Selection"
                
                MenuRec.Open SQL, MenuCon, adOpenKeyset, adLockOptimistic
                
                Cbox_Variant.Visible = True
                Cbox_Variant.Clear
                Cbox_Variant.ColumnCount = 1
                Cbox_Variant.ColumnWidths = "120"
                
                                With MenuRec
                                        
                                  .MoveFirst
                                    
                                    Do Until .EOF
                                        If .Fields(Vartype) <> "" Then
                                            Cbox_Variant.AddItem .Fields(Vartype)
                                        End If
                                        .MoveNext
                                    Loop
                                
                                End With
                MenuRec.Close
        'End If
        
        
        
        
        
        MenuCon.Close
        
    End If


'    Dim R As Range
'    Dim C As Range
'    Cbox_Variant.Clear
'    Cbox_Variant.ColumnCount = 1
'    Cbox_Variant.ColumnWidths = "140"
'    Dim Dvar As Variant
'    Dim FindR As Range
'    'Dim Res As String
'
'    'Res = CBox_Dish.Text
'        If Not CBox_Dish.Text = "" Then
''    Select Case Cbox_From_Business.Value
''        Case Is = ShSupport.Range("A4").Value
'        Set R = ShMenu.ListObjects(Cbox_From_Business & "_Variant_Selection").DataBodyRange
'
'         Dvar = Application.VLookup(CStr(CBox_Dish.Text), R, 2, 1)
'
'         Cbox_Variant.Value = ""
'
'
'
'
'        Set R = ShMenu.ListObjects(Dvar).DataBodyRange
'
'
'        Cbox_Variant.Visible = True
'        For Each C In R
'            Cbox_Variant.AddItem C.Value
'        Next C
'
'        Else
'            Cbox_Variant.Visible = False
'            Cbox_Variant.Clear
'
'
'        End If
''        Case Is = ShMenu.Range("A5").Value
''
''        Set R = ShMenu.ListObjects("Manna_Variant_Selection").DataBodyRange
''
''            Dvar = Application.VLookup(CStr(CBox_Dish.Text), R, 2, 0)
''
''            Set R = ShMenu.ListObjects(Dvar).DataBodyRange
''
''            Cbox_Variant.Visible = True
''
''
''    End Select
End Sub



Private Sub Cbox_DishCode_Change()
    Dim Mcon As New ADODB.Connection
    Dim Mrec As New ADODB.Recordset
    Dim SQL As String
     
     If Len(Cbox_DishCode.Text) > 1 Then
     
        Call Menu_Connection(Mcon)
        
                If Left(Cbox_DishCode.Text, 1) = "A" Then
                    
                    Cbox_From_Business.Text = "Arabian"
                    SQL = "Select * From TBL_Arabian_Menu Where TBL_Arabian_Menu.[Dish Code]=" & Right(Cbox_DishCode.Text, Len(Cbox_DishCode.Text) - 1)
                    
                ElseIf Left(Cbox_DishCode.Text, 1) = "M" Then
                    
                    Cbox_From_Business.Text = "Manna"
                    SQL = "Select * From TBL_Manna_Menu Where TBL_Manna_Menu.[Dish Code]=" & Right(Cbox_DishCode.Text, Len(Cbox_DishCode.Text) - 1)
                
                ElseIf Left(Cbox_DishCode.Text, 1) = "P" Then
                    
                    Cbox_From_Business.Text = "AlaRasi"
                    SQL = "Select * From TBL_AlaRasi_Menu Where TBL_AlaRasi_Menu.[Dish Code]=" & Right(Cbox_DishCode.Text, Len(Cbox_DishCode.Text) - 1)
                
                Else
                
                    MsgBox "Please Select a Valid Dish Code"
                    
                    Exit Sub
                                        
                End If
            
            Mrec.Open SQL, Mcon, adOpenKeyset, adLockOptimistic
                With Mrec
                    If .RecordCount = 0 Then
                        Exit Sub
                    
                    End If
                           .MoveFirst
                        
                        CBox_Dish.Text = .Fields("Dish Cat")
                        Cbox_Variant.Text = .Fields("Variant")
            
            
                End With
                
        
        Cbox_Variant.Visible = True
    
     End If
    
        
    
    
    
    
    
''    If Cbox_DishCode <> "" Then
''        If Left(Cbox_DishCode.Value, 1)=A Then
''            MsgBox "Please enter the Dish code number only", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
''        Else
'
'
'        Dim Arabian As Long
'        Dim Manna As Long
'        Dim AlaRasi As Long
'        Dim MenuRange As Range
'        Dim DishCoderange As Range
'        Dim PriceColumn As Byte
''        Dim Wrd1 As Byte
''        Dim Wrd2 As Byte
'        Dim CodeNo As Long
'
'        Select Case CBox_BillingType.Text
'
'            Case Is = ShMenu.Range("D1").Value 'Online Price Column
'            PriceColumn = 3
'
'            Case Is = ShMenu.Range("E1").Value  'TakeAway Price Column
'            PriceColumn = 4
'
'            Case Is = ShMenu.Range("F1").Value  'DineIn Price Column
'            PriceColumn = 5
'
'        End Select
'
'        Arabian = ShMenu.ListObjects("Arabian_Menu_List").DataBodyRange.Rows.Count
'        Manna = ShMenu.ListObjects("Manna_Menu_List").DataBodyRange.Rows.Count
'        AlaRasi = ShMenu.ListObjects("AlaRasi_Menu_List").DataBodyRange.Rows.Count
'
'        'Dim VariantType As Variant 'Variant type
'        Dim Dish As String         'Dish Type
'        Dim VariantR As Range
'
'        Select Case Left(Cbox_DishCode.Text, 1)
'
'
'            Case Is = "A" 'For Arabian Hub
'
'            Let CodeNo = Right(Cbox_DishCode.Text, Len(Cbox_DishCode.Text) - 1)
'            Set VariantR = ShMenu.ListObjects("Arabian_Variant_Selection").DataBodyRange
'
'            Set MenuRange = ShMenu.ListObjects("Arabian_Menu_List").DataBodyRange
'
'            Set DishCoderange = ShMenu.Cells(Application.WorksheetFunction.XMatch(Cbox_DishCode.Text, ShMenu.Range("A2", "A" & ShMenu.Range("A2").CurrentRegion.Rows.Count)), 1)
'
'                 Dish = DishCoderange.Offset(, 1).Value
'                 VariantR = DishCoderange.Offset(, 2).Value
'
'
'                 Cbox_From_Business.Text = ShSupport.Range("A4").Value
'
'                    CBox_Dish.Text = Dish
'
''                    Wrd1 = Len(ShMenu.Range("C" & CLng(Cbox_DishCode.Text) + 1).Value)
''                    Wrd2 = Len(CBox_Dish.Text) + 1
'
'                    'VariantType = Application.VLookup(Dish, VariantRange, 2, False)
'
'
'
'                        Cbox_Variant.Text = VariantR
'
'
'
'
'
'
'
'            Case Is = "M"
''            Set VariantR = ShMenu.ListObjects("Manna_Variant_Selection").DataBodyRange
''            Set MenuRange = ShMenu.ListObjects("Manna_Menu_List").DataBodyRange
''            Set DishCoderange = ShMenu.Range("U2", "U" & Manna + 1)
''                    Dish = ShMenu.Range("V" & CLng(Cbox_DishCode.Text) + 1 - Arabian).Value
''
''                Cbox_From_Business.Text = ShSupport.Range("A5").Value
''
''                    CBox_Dish.Text = Dish
'''                    Wrd1 = Len(ShMenu.Range("W" & CLng(Cbox_DishCode.Text) + 1 - Arabian).Value)
'''                    Wrd2 = Len(CBox_Dish.Text) + 1
''                    VariantType = Application.VLookup(Dish, VariantRange, False)
''
''                    If Not VariantType = ShMenu.Range("K2").Value Then
''                    Cbox_Variant.Text = Right(ShMenu.Range("W" & CLng(Cbox_DishCode.Text) + 1 - Arabian), Wrd1 - Wrd2)
''
''                    Else
''                        Cbox_Variant.Text = ShMenu.Range("K2").Value
''                    End If
''
''
''            Case Is = "P"
''            Set VariantR = ShMenu.ListObjects("AlaRasi_Variant_Selection").DataBodyRange
''            Set MenuRange = ShMenu.ListObjects("AlaRasi_Menu_List").DataBodyRange
''            Set DishCoderange = ShMenu.Range("AE2", "AE" & AlaRasi + 1)
''                    Dish = ShMenu.Range("AF" & CLng(Cbox_DishCode.Text) + 1 - Arabian - Manna).Value
''                Cbox_From_Business.Text = ShSupport.Range("A6").Value
''
''                    CBox_Dish.Text = Dish
'''                    Wrd1 = Len(ShMenu.Range("AG" & CLng(Cbox_DishCode.Text) + 1 - Arabian - Manna).Value)
'''                    Wrd2 = Len(CBox_Dish.Text) + 1
''                    VariantType = Application.VLookup(Dish, VariantRange, False)
''
''                    If Not VariantType = ShMenu.Range("K2").Value Then
''                        Cbox_Variant.Text = Right(ShMenu.Range("AG" & CLng(Cbox_DishCode.Text) + 1 - Arabian - Manna), Wrd1 - Wrd2)
''                    Else
''                        Cbox_Variant.Text = ShMenu.Range("K2").Value
''                    End If
''
''            Case Else
''
''                MsgBox "Please Select the correct Dish Code", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
''
'            End Select
''
''        End If
'
'
''    End If
End Sub


Private Sub Cbox_From_Business_Change()
    Dim R As Range
    Dim C As Range
    CBox_Dish.Clear
    CBox_Dish.ColumnCount = 1
    CBox_Dish.ColumnWidths = "140"
    Dim DataCon As New ADODB.Connection
    Dim MenuRec As New ADODB.Recordset
    Dim SQL As String
                
        CBox_Dish.Visible = False
        Cbox_Variant.Visible = False
        LBL_Amount.Visible = False
    
    Select Case Cbox_From_Business.Value
        Case Is = ShSupport.Range("A4").Value 'Arabian Hub
            SQL = "SELECT TBL_Arabian_Dish_Variant.[Dish Cat], TBL_Arabian_Dish_Variant.[Variant Type] From TBL_Arabian_Dish_Variant Order by TBL_Arabian_Dish_Variant.[Dish No] ASC"
            
            Call Menu_Connection(DataCon)
            
            MenuRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            
            With MenuRec
                .MoveFirst
                Do Until .EOF
                    CBox_Dish.AddItem .Fields("Dish Cat")
                    .MoveNext
                Loop
                
            End With
            
            MenuRec.Close
            DataCon.Close
            
            
'            Set R = ShMenu.ListObjects("Arabian_Variant_Selection").ListColumns("Dish").DataBodyRange
'            For Each C In R
'                CBox_Dish.AddItem C.Value
'            Next C
            
            LBL_Dish.Visible = True
            CBox_Dish.Visible = True

        
        Case Is = ShSupport.Range("A5").Value   'Manna
            SQL = "SELECT TBL_Manna_Dish_Variant.[Dish Cat], TBL_Manna_Dish_Variant.[Variant Type] From TBL_Manna_Dish_Variant Order by TBL_Manna_Dish_Variant.[Dish No] ASC"
            
            Call Menu_Connection(DataCon)
            
            MenuRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            
            With MenuRec
                .MoveFirst
                Do Until .EOF
                 CBox_Dish.AddItem .Fields("Dish Cat")
                 .MoveNext
                Loop
        
            End With
            
            MenuRec.Close
            DataCon.Close
        
        
'            Set R = ShMenu.ListObjects("Manna_Variant_Selection").ListColumns("Dish").DataBodyRange
'            For Each C In R
'                CBox_Dish.AddItem C.Value
'            Next C
            
            LBL_Dish.Visible = True
            CBox_Dish.Visible = True

        
        Case Is = ShSupport.Range("A6").Value   'AlaRasi
            SQL = "SELECT TBL_AlaRasi_Dish_Variant.[Dish Cat], TBL_AlaRasi_Dish_Variant.[Variant Type] From TBL_AlaRasi_Dish_Variant Order by TBL_AlaRasi_Dish_Variant.[Dish No] ASC"
            
            Call Menu_Connection(DataCon)
            
            MenuRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            
            With MenuRec
                .MoveFirst
                Do Until .EOF
                    CBox_Dish.AddItem .Fields("Dish Cat")
                    .MoveNext
                Loop
            
            End With
            
            MenuRec.Clone
            DataCon.Close
        
'            Set R = ShMenu.ListObjects("AlaRasi_Variant_Selection").ListColumns("Dish").DataBodyRange
'            For Each C In R
'                CBox_Dish.AddItem C.Value
'            Next C
            
            LBL_Dish.Visible = True
            CBox_Dish.Visible = True

        
        Case Is = ""
            LBL_Dish.Visible = False
            CBox_Dish.Visible = False
            Cbox_Variant.Text = ""
            Cbox_Variant.Visible = False
            LBL_Amount.Visible = False
            Cbox_Variant.Text = ""
            Cbox_DishCode.Text = ""
            
        Case Else
        
            MsgBox "Please Select from the DropDown List", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
    End Select
    
    
End Sub

Private Sub Cbox_Multiple_Cust_Name_Change()

    TXT_CustName.Text = Cbox_Multiple_Cust_Name
    
End Sub



Private Sub Cbox_Variant_Change()
    Dim MenuCon As New ADODB.Connection
    Dim MenuRec As New ADODB.Recordset
    Dim SQL As String
    Dim TBL As String
    Dim CodeT As String
    
    
    
    'for Dishcode
    Select Case Cbox_From_Business.Text
        Case Is = "Arabian"
            CodeT = "A"
        Case Is = "Manna"
            CodeT = "M"
        Case Is = "AlaRasi"
            CodeT = "P"
    End Select
    
    If Not Cbox_Variant.Text = "" And Not Cbox_From_Business.Text = "" And Not CBox_BillingType.Text = "" And Not CBox_Dish.Text = "" Then
    
    TBL = "TBL_" & Cbox_From_Business.Text & "_Menu"
    SQL = "Select * From " & TBL & " Where " & TBL & ".[Dish Cat]='" & CBox_Dish.Text & "' AND " & TBL & ".Variant='" & Cbox_Variant.Text & "'"
    

        Call Menu_Connection(MenuCon)
        MenuRec.Open SQL, MenuCon, adOpenKeyset, adLockOptimistic


            With MenuRec
                If .RecordCount = 1 Then
                    
                    .MoveFirst
                    LBL_Disc.Visible = True
                    TXT_Disc.Visible = True
                    LBL_Disc = ShMenu.Range("G2").Value
                    LBL_Qty.Visible = True
                    LBL_Amount.Visible = True
                    LBL_Rupee_Symbol.Visible = True
                    LBL_Amount = .Fields(CBox_BillingType.Text) * LBL_Qty
                    LBL_Rupee_Symbol = ShMenu.Range("G3").Value
                    
                    
                    Cbox_DishCode.Text = CodeT & .Fields("Dish Code")
                    LBL_PrintCode = .Fields("Print Name")
                    
                ElseIf .RecordCount > 1 Then
                    
                    MsgBox "There are more than one record with the same Dish. Please Update the menu from the Database.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                    
                ElseIf .RecordCount < 1 Then
                    
                    MsgBox "No Record Found please Update the Dish Table in Database.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                
                End If

            End With
    
    Else
    
        LBL_Amount = ""
        LBL_Amount.Visible = False
        LBL_Qty.Visible = False
        LBL_Rupee_Symbol.Visible = False
        

    End If


'    Dim MenuRange As Range
'    Dim MenuVariant As Range
'    'Dim PriceRange As Range
'    Dim Dish As Range   'For looping
'
'
'    Select Case Cbox_From_Business.Text
'
'        Case Is = ShSupport.Range("A4").Value      'Arabian Hub
'            Set MenuRange = ShMenu.ListObjects("Arabian_Menu_list").ListColumns("Cat").DataBodyRange
'
'             For Each Dish In MenuRange
'                If CBox_Dish.Text = Dish.Value & Cbox_Variant.Text = Dish.Offset(, 1).Value Then
'                    LBL_Amount.Visible = True
'                    If CBox_BillingType = ShMenu.Range("D1").Value Then     'For Online
'
'                        LBL_Amount = Dish.Offset(, 2).Value
'
'                    ElseIf CBox_BillingType = ShMenu.Range("E1").Value Then 'for TakeAway
'
'                        LBL_Amount = Dish.Offset(, 3).Value
'
'                    ElseIf CBox_BillingType = ShMenu.Range("F1").Value Then 'for DineIn
'
'                        LBL_Amount = Dish.Offset(, 4).Value
'
'                    End If
'
'                End If
'
'             Next Dish
'
'        Case Is = ShSupport.Range("A5").Value      'Manna
'
'            Set MenuRange = ShMenu.ListObjects("Manna_Menu_list").ListColumns("Cat").DataBodyRange
'
'             For Each Dish In MenuRange
'                If CBox_Dish.Text = Dish.Value & Cbox_Variant.Text = Dish.Offset(, 1).Value Then
'                    LBL_Amount.Visible = True
'                    If CBox_BillingType = ShMenu.Range("D1").Value Then     'For Online
'
'                        LBL_Amount = Dish.Offset(, 2).Value
'
'                    ElseIf CBox_BillingType = ShMenu.Range("E1").Value Then 'for TakeAway
'
'                        LBL_Amount = Dish.Offset(, 3).Value
'
'                    ElseIf CBox_BillingType = ShMenu.Range("F1").Value Then 'for DineIn
'
'                        LBL_Amount = Dish.Offset(, 4).Value
'
'                    End If
'
'                End If
'
'            Next Dish
'
'
'        Case Is = ShSupport.Range("A6").Value      'AlaRasi
'
'            Set MenuRange = ShMenu.ListObjects("Arabian_Menu_list").ListColumns("Cat").DataBodyRange
'
'            For Each Dish In MenuRange
'               If CBox_Dish.Text = Dish.Value & Cbox_Variant.Text = Dish.Offset(, 1).Value Then
'                   LBL_Amount.Visible = True
'                   If CBox_BillingType = ShMenu.Range("D1").Value Then     'For Online
'
'                       LBL_Amount = Dish.Offset(, 2).Value
'
'                   ElseIf CBox_BillingType = ShMenu.Range("E1").Value Then 'for TakeAway
'
'                       LBL_Amount = Dish.Offset(, 3).Value
'
'                   ElseIf CBox_BillingType = ShMenu.Range("F1").Value Then 'for DineIn
'
'                       LBL_Amount = Dish.Offset(, 4).Value
'
'                   End If
'
'               End If
'
'            Next Dish
'
'        Case Else
'
'                MsgBox "Please Select the Proper Billing Type", , "ALARASI FOODS AND BEVERAGES PRIVATE LIMITED"
'
'
'    End Select
'
'
'
'
'
''    'Need to write code
''     If Not Cbox_Variant.Text = "" Then
''        Dim Bfrom As String
''        Dim BillingType As Byte
''        Dim Dish As String
''        Dim DishVariant As String
''        Dim MenuList As Range
''
''        Bfrom = Cbox_From_Business.Text
''
''        Select Case CBox_BillingType.Text
''
''            Case Is = ShMenu.Range("D1").Value      'Online
''                BillingType = 3
''
''            Case Is = ShMenu.Range("E1").Value      'TakeAway
''                BillingType = 4
''            Case Is = ShMenu.Range("F1").Value      'DineIn
''                BillingType = 5
''        End Select
''
''        Select Case Cbox_From_Business.Text
''            Case Is = ShSupport.Range("A4").Value   'Arabian Hub
''                Set MenuList = ShMenu.ListObjects("Arabian_Menu_list").DataBodyRange
''            Case Is = ShSupport.Range("A5").Value   'Manna
''                Set MenuList = ShMenu.ListObjects("Manna_Menu_List").DataBodyRange
''            Case Is = ShSupport.Range("A6").Value   'AlaRAasi
''                Set MenuList = ShMenu.ListObjects("AlaRasi_Menu_List").DataBodyRange
''
''        End Select
''
''
''
''
''     End If
'
End Sub



Private Sub LBL_Disc_Click()
    
    Select Case LBL_Disc
    
    Case Is = ShMenu.Range("G2").Value
        
        LBL_Disc = ShMenu.Range("G3").Value
        
    Case Is = ShMenu.Range("G3").Value
        
        LBL_Disc = ShMenu.Range("G2").Value
    
    End Select
    
End Sub

Private Sub LBL_Qty_Click()
    LBL_Qty = LBL_Qty + 1
    Call Cbox_Variant_Change
End Sub

Private Sub LBL_Qty_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not LBL_Qty = 2 Then
        LBL_Qty = LBL_Qty - 2
        Call Cbox_Variant_Change
    End If
End Sub

Private Sub LBox_Table_Detail_Change()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    Dim No As Byte
    No = 1
            Lbox_Order_Detail.Clear
            Lbox_Order_Detail.ColumnCount = 4
            Lbox_Order_Detail.ColumnWidths = "120,50,120,120"
    
            Lbox_Order_Detail.AddItem
            Lbox_Order_Detail.List(0, 0) = "Item"
            Lbox_Order_Detail.List(0, 1) = "Qty"
            Lbox_Order_Detail.List(0, 2) = "Amount"
            Lbox_Order_Detail.List(0, 3) = "Restaurant"
        
    If IsNumeric(LBox_Table_Detail.Value) Then

            If IsNull(LBox_Table_Detail.List(LBox_Table_Detail.Value, 3)) = False Then
                Dim Hold1 As String
                Hold1 = LBox_Table_Detail.List(LBox_Table_Detail.Value, 3)
                SQL = "Select * From TBL_KOT_HOLD_DETAIL Where TBL_KOT_HOLD_DETAIL.[KOT Hold Number]= '" & Hold1 & "'"
                
                Call Connection(DataCon)
                DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic

                With DataRec
                    .MoveFirst
                    Do Until .EOF
                        
                        Lbox_Order_Detail.AddItem
                        Lbox_Order_Detail.List(No, 0) = .Fields("Item")
                        Lbox_Order_Detail.List(No, 1) = .Fields("QTY")
                        Lbox_Order_Detail.List(No, 2) = .Fields("Amount")
                        Lbox_Order_Detail.List(No, 3) = .Fields("From")
                        No = No + 1
                    
                        .MoveNext
                    Loop


                End With
                
                DataRec.Close
                DataCon.Close
            End If

            If IsNull(LBox_Table_Detail.List(LBox_Table_Detail.Value, 4)) = False Then
                Dim Hold2 As String
                Hold2 = LBox_Table_Detail.List(LBox_Table_Detail.Value, 4)
                SQL = "Select * From TBL_KOT_HOLD_DETAIL Where TBL_KOT_HOLD_DETAIL.[KOT Hold Number]= '" & Hold2 & "'"
                
                Call Connection(DataCon)
                DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic

                With DataRec
                    .MoveFirst
                    Do Until .EOF
                        
                        Lbox_Order_Detail.AddItem
                        Lbox_Order_Detail.List(No, 0) = .Fields("Item")
                        Lbox_Order_Detail.List(No, 1) = .Fields("QTY")
                        Lbox_Order_Detail.List(No, 2) = .Fields("Amount")
                        Lbox_Order_Detail.List(No, 3) = .Fields("From")
                        No = No + 1
                    
                        .MoveNext
                    Loop


                End With
                
                DataRec.Close
                DataCon.Close
                
                
                
                
                
                
                
'                Dim Hold2 As String
'                Hold2 = LBox_Table_Detail.List(LBox_Table_Detail.Value, 4)
'                MsgBox Hold2, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            End If

    End If
        
End Sub




Private Sub multipage_billing_Change()
    Call UserForm_Initialize
End Sub




Private Sub multipage_billing_DblClick(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
                                        ShPrint.Range("A:C").Delete
                                        CBox_BillingType.Text = ""
                                        Cbox_Disc_On.Text = ""
                                        CBox_Dish.Text = ""
                                        Cbox_DishCode.Text = ""
                                        Cbox_From_Business.Text = ""
                                        Cbox_Multiple_Cust_Name.Text = ""
                                        CBox_Payment_Mode1.Text = ""
                                        CBox_Payment_Mode2.Text = ""
                                        Cbox_Variant.Text = ""
                                        TXT_CustName.Text = ""
                                        TXT_Disc.Text = ""
                                        TXT_Mobile_Number.Text = ""
                                        TXT_Payment1.Text = ""
                                        TXT_Payment2.Text = ""
                                        LBox_Ordered_Item.Clear
                                        LBL_PrintCode = ""
                                        Call UserForm_Initialize
End Sub


Private Sub TXT_Disc_Change()
    If Not Cbox_Disc_On = "" Then
                
    
    End If
    
End Sub




Private Sub TXT_Mobile_Number_Change()
    Select Case CBox_BillingType.Value

    Case Is = ShMenu.Range("E1").Value, ShMenu.Range("F1").Value

   If CBox_BillingType.Value = ShMenu.Range("E1").Value Or _
   CBox_BillingType.Value = ShMenu.Range("F1").Value Then
    If Len(TXT_Mobile_Number.Text) = 10 Then
      Dim CustSql As String
      Dim DataCon As New ADODB.Connection
      Dim CustRec As New ADODB.Recordset

      CustSql = "Select * From TBL_Customer_Detail Where TBL_Customer_Detail.[Customer Mobile Number]='" & TXT_Mobile_Number.Text & "'"

      Call Connection(DataCon)
      CustRec.Open CustSql, DataCon, adOpenKeyset, adLockOptimistic

      If CustRec.RecordCount = 1 Then
            LBL_Customer_Status.Visible = False
            Cbox_Multiple_Cust_Name.Visible = False
            CustRec.MoveFirst
            TXT_CustName.Text = CustRec.Fields("Customer Name")
            BTN_Add_Customer.Caption = "Update Customer"
            BTN_Add_Customer.Visible = True
        ElseIf CustRec.RecordCount > 1 Then
            LBL_Customer_Status.Visible = True
            Cbox_Multiple_Cust_Name.Visible = True
            Cbox_Multiple_Cust_Name.ColumnCount = 1

            With CustRec
                .MoveFirst
                Do Until .EOF
                        Cbox_Multiple_Cust_Name.AddItem .Fields("Customer Name")
                    .MoveNext
                Loop


            End With
            BTN_Add_Customer.Caption = "Update Customer"
            BTN_Add_Customer.Visible = True
            LBL_Customer_Status = "More than once the number is Saved"

        Else

            'New Customer
            BTN_Add_Customer.Caption = "Add Customer"
            BTN_Add_Customer.Visible = True
            Cbox_Multiple_Cust_Name.Clear
            Cbox_Multiple_Cust_Name.Visible = False
            LBL_Customer_Status.Visible = True
            LBL_Customer_Status = "10 Digit"
            

        End If

            CustRec.Close
            DataCon.Close
              ElseIf Len(TXT_Mobile_Number.Text) = 0 Then
                  LBL_Customer_Status.Visible = False
                  Cbox_Multiple_Cust_Name.Visible = False
                  BTN_Add_Customer.Visible = False
        Else

        LBL_Customer_Status.Visible = True
        LBL_Customer_Status = "Not 10 digit"
        Cbox_Multiple_Cust_Name.Clear
        Cbox_Multiple_Cust_Name.Visible = False
        BTN_Add_Customer.Visible = False

        End If
    ElseIf CBox_BillingType.Text = ShMenu.Range("D1").Value Then

    End If

    Case Else

    End Select
End Sub




Private Sub UserForm_Activate()
    MultiPage_Billing.Value = 0
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    Dim No As Long
    
    If MultiPage_Billing.Value = 0 Then
    
        No = 0
        
        LBL_Customer_Status.Visible = False
        LBox_Ordered_Item.Visible = False
        CBox_Dish.Visible = False
        LBL_Dish.Visible = False
        Cbox_Multiple_Cust_Name.Visible = False
        Cbox_From_Business.Visible = False
        Cbox_Variant.Visible = False
        LBL_Amount.Visible = False
        Cbox_DishCode.Visible = False
        LBL_CustomerName.Visible = False
        TXT_CustName.Visible = False
        LBL_Mobile_No.Visible = False
        TXT_Mobile_Number.Visible = False
        LBL_Qty.Visible = False
        LBL_Disc.Visible = False
        LBL_Amnt_Desc.Visible = False
        LBL_Total_Amnt.Visible = False
        TXT_Disc.Visible = False
        LBL_Business_From.Visible = False
        LBL_Dish_Code.Visible = False
        LBL_Rupee_Symbol.Visible = False
        Lbl_Payment1.Visible = False
        CBox_Payment_Mode1.Visible = False
        TXT_Payment1.Visible = False
        Lbl_Payment2.Visible = False
        CBox_Payment_Mode2.Visible = False
        TXT_Payment2.Visible = False
        LBL_Disc.Visible = False
        Lbl_Disc_on.Visible = False
        Cbox_Disc_On.Visible = False
        LBL_Disc_Title.Visible = False
        BTN_Add_Customer.Visible = False
        LBL_DineIn_Step = "Step1"
        LBL_DineIn_Step.Visible = False
        
        
        
        
        
        'For adding Header in the listbox (LBox_Ordered)
                LBox_Ordered_Item.Clear
                LBox_Ordered_Item.ColumnCount = 6
                LBox_Ordered_Item.ColumnWidths = "250,150,80,100,120,150"
                LBox_Ordered_Item.AddItem
                LBox_Ordered_Item.List(0, 0) = "Dish"
                LBox_Ordered_Item.List(0, 1) = "Variant"
                LBox_Ordered_Item.List(0, 2) = "QTY"
                LBox_Ordered_Item.List(0, 3) = "Amount"
                LBox_Ordered_Item.List(0, 4) = "From"
                LBox_Ordered_Item.List(0, 5) = "Print Code"

        
    SQL = "Select * From TBL_Billing_Type"
        Call Menu_Connection(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        
        CBox_BillingType.Clear
        CBox_BillingType.ColumnCount = 1
        CBox_BillingType.ColumnWidths = "120"
        
        With DataRec
            .MoveFirst
            Do Until .EOF
                CBox_BillingType.AddItem .Fields("Billing type")
                .MoveNext
            Loop
            
            
        End With
        
        DataRec.Close
        
        CBox_Payment_Mode1.Clear
        CBox_Payment_Mode2.Clear
        CBox_Payment_Mode1.ColumnCount = 1
        CBox_Payment_Mode2.ColumnCount = 1
        CBox_Payment_Mode1.ColumnWidths = "80"
        CBox_Payment_Mode2.ColumnWidths = "80"
        
        Cbox_Disc_On.Clear
        Cbox_Disc_On.ColumnCount = 1
        Cbox_Disc_On.ColumnWidths = "120"
        Cbox_Disc_On.AddItem ShSupport.Range("A4").Value
        Cbox_Disc_On.AddItem ShSupport.Range("A5").Value
        Cbox_Disc_On.AddItem "Both"
        
        
        
        
    SQL = "Select * From TBL_Payment_Mode"
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        With DataRec
            .MoveFirst
            Do Until .EOF
                CBox_Payment_Mode1.AddItem .Fields("Payment Mode")
                CBox_Payment_Mode2.AddItem .Fields("Payment Mode")
                .MoveNext
            Loop
        End With
        
        DataRec.Close
        
    SQL = "Select * From TBL_Arabian_Menu"
    
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        Cbox_DishCode.Clear
        Cbox_DishCode.ColumnCount = 2
        Cbox_DishCode.ColumnWidths = "50,140"
        Cbox_DishCode.ListWidth = "350"
            With DataRec
                .MoveFirst
                Do Until .EOF
                    Cbox_DishCode.AddItem
                    Cbox_DishCode.List(No, 0) = "A" & .Fields("Dish Code")
                    Cbox_DishCode.List(No, 1) = .Fields("Print Name")
                    No = No + 1
                    .MoveNext
                    
                Loop
                
                
                
            End With
            
        DataRec.Close
        
    SQL = "Select * From TBL_Manna_Menu"
    
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
            With DataRec
                .MoveFirst
                Do Until .EOF
                  Cbox_DishCode.AddItem
                  Cbox_DishCode.List(No, 0) = "M" & .Fields("Dish Code")
                  Cbox_DishCode.List(No, 1) = .Fields("Print Name")
                  No = No + 1
                  .MoveNext
                Loop
                
                
                
            End With
        
        DataRec.Close
    
    SQL = "Select * From TBL_AlaRasi_Menu"
        
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            With DataRec
                .MoveFirst
                Do Until .EOF
                    Cbox_DishCode.AddItem
                    Cbox_DishCode.List(No, 0) = "P" & .Fields("Dish Code")
                    Cbox_DishCode.List(No, 1) = .Fields("Print Name")
                    No = No + 1
                    .MoveNext
                Loop
                
                
                
            End With
        
        DataRec.Close
        DataCon.Close
        
    ElseIf MultiPage_Billing.Value = 1 Then
        
        SQL = "Select * From TBL_Table_Status ORDER BY TBL_Table_Status.[Table Number] Asc;"
        Call Connection(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            
            LBox_Table_Detail.Clear
            LBox_Table_Detail.ColumnCount = 5
            LBox_Table_Detail.ColumnWidths = "50,70,120,120,120"
            Lbox_Order_Detail.Clear
            Lbox_Order_Detail.ColumnCount = 4
            Lbox_Order_Detail.ColumnWidths = "120,50,120,120"
            
            'For Adding the header
            
            LBox_Table_Detail.AddItem
            LBox_Table_Detail.List(0, 0) = "T.No."
            LBox_Table_Detail.List(0, 1) = "Status"
            LBox_Table_Detail.List(0, 2) = "Cust Name"
            LBox_Table_Detail.List(0, 3) = "Hold No."
            LBox_Table_Detail.List(0, 4) = "Hold No. 'Bill 2'"
            
            
            
            Lbox_Order_Detail.AddItem
            Lbox_Order_Detail.List(0, 0) = "Item"
            Lbox_Order_Detail.List(0, 1) = "Qty"
            Lbox_Order_Detail.List(0, 2) = "Amount"
            Lbox_Order_Detail.List(0, 3) = "Restaurant"
            
            
            No = 1
            
            
        
        With DataRec
            .MoveFirst
            Do Until .EOF
              LBox_Table_Detail.AddItem
                LBox_Table_Detail.List(No, 0) = .Fields("Table Number")
                LBox_Table_Detail.List(No, 1) = .Fields("Status")
                
                    If IsNull(.Fields("Customer Name")) = False Then
                        LBox_Table_Detail.List(No, 2) = .Fields("Customer Name")
                    End If
                    
                    If IsNull(.Fields("KOT Hold Number")) = False Then
                        LBox_Table_Detail.List(No, 3) = .Fields("KOT Hold Number")
                    End If
                    
                    If IsNull(.Fields("KOT Hold Number2")) = False Then
                        LBox_Table_Detail.List(No, 4) = .Fields("KOT Hold Number2")
                    End If
                
                No = No + 1
                
                .MoveNext
                
            Loop
            
            
        End With
        
        DataRec.Close
        DataCon.Close
        
    End If
End Sub


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
                If KeyCode = vbKeyReturn Then
                    
                
                End If
End Sub
