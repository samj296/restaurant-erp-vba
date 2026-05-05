Attribute VB_Name = "PrintPage"
Option Explicit

Sub HeaderPrint()
      
                                    With ShPrint.Range("A1", "C1")
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                        .Font.Size = 14
                                        .Merge
                                        .Font.Bold = True
                                        .RowHeight = 55
                                        .Value = "ALA RASI FOODS AND" & Chr(10) & "BEVERAGES PRIVATE" & Chr(10) & "LIMITED"
                                    End With
                              
End Sub

Sub CustName_MNumber(Cname As String, Mnumber As String)
    
                                    ShPrint.Range("A3").Font.Size = 12
                                    ShPrint.Range("A4", "C4").Merge
                                    ShPrint.Range("A4", "C4").Font.Size = 12
                                    ShPrint.Range("A4", "C4").HorizontalAlignment = xlLeft
                                    ShPrint.Range("A4", "C4").VerticalAlignment = xlCenter
                                    
                                     
                                    ShPrint.Range("A4").Value = Cname & " - " & Mnumber
End Sub
