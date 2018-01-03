VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'By Matthew Connolly
'Modifed: 3/31/2016
'Purpose: Copy ReCeiver to logs

Private Sub CopyReceiver2X(ws As String, y As String) 'Sub to Copy froM ReCiever to one of the logs

Set ReceiverRange = Range(y) 'Range that is used for the loop

Dim DestRow As Long
Dim n As Long
Dim z As Integer
n = 0 ' Long variable n set to value 0 (counter)
z = Application.WorksheetFunction.CountA(ReceiverRange) ' Counts how Many are filled in the range y
If ws = "CARPET" Then 'goes into this if when ws has the value "CARPET"
    a = 11 'starting row nuMber
    If Sheets("Receiver").Range("B11") = "n/a" Then 'checks to see if B11 on reCeiver has n/a, if so end Sub
        Exit Sub
    End If
    
    
    Do Until n = z 'loops until counter (n) equals how many cells that are filled in the range y

        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available row in CARPET Sheet

        Sheets("Receiver").Range("B" & a & ":E" & a).Copy Sheets(ws).Range("B" & DestRow & ": E" & DestRow) 'Copies section from Receiver to Carpet
        Sheets("Receiver").Range("F" & a & ":G" & a).Copy Sheets(ws).Range("F" & DestRow & ": G" & DestRow) 'Copies section from Receiver to Carpet
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section from ReCeiver to Carpet
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("I" & DestRow) 'Copies section from Receiver to Carpet
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("H" & DestRow) 'Copies section from Receiver to Carpet

        'Debug.Print (n)
        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1

    Loop ' End of loop
ElseIf ws = "CPT.TILE" Then 'Jumps to this code if ws value is "CPT.TILE"
    a = 21 ' Starting row number
    
    If Sheets("Receiver").Range("B21") = "n/a" Then 'If B21 has the value n/a then it will end sub
        Exit Sub
    End If
    
    Do Until n = z 'loops until counter (n) equals how many cells that are filled in the range y
        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available row in CPT.Tile Sheet
        
        Sheets("Receiver").Range("B" & a & ":E" & a).Copy Sheets(ws).Range("B" & DestRow & ": E" & DestRow) 'Copies section from Receiver to CPT.Tile
        Sheets("Receiver").Range("G" & a & ":I" & a).Copy Sheets(ws).Range("F" & DestRow & ": H" & DestRow) 'Copies section from Receiver to CPT.Tile
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section from Receiver to CPT.Tile
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("J" & DestRow) 'Copies section from ReCeiver to CPT.Tile
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("I" & DestRow) 'Copies seCtion from Receiver to CPT.Tile


        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1


    Loop
ElseIf ws = "VINYL" Then
    a = 31
    
    If Sheets("Receiver").Range("B31") = "n/a" Then 'If B31 has the value n/a then it will end sub
        Exit Sub
    End If
    
    Do Until n = z 'loops until counter (n) equals how many cells that are filled in the range y

        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available cell in Vinyl Sheet

        Sheets("Receiver").Range("B" & a & ":E" & a).Copy Sheets(ws).Range("B" & DestRow & ": E" & DestRow) 'Copies section from Receiver to VINYL
        Sheets("Receiver").Range("F" & a & ":G" & a).Copy Sheets(ws).Range("F" & DestRow & ": G" & DestRow) 'Copies section from Receiver to VINYL
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section from Receiver to VINYL
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("I" & DestRow) 'Copies section from Receiver to VINYL
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("H" & DestRow) 'Copies section from Receiver to VINYL

        'Debug.Print (n)
        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1


    Loop
    
ElseIf ws = "TILE" Then
    a = 41
    
    If Sheets("Receiver").Range("B41") = "n/a" Then 'If B41 has the value n/a then it will end sub
        Exit Sub
    End If
    
    Do Until n = z 'loops until counter (n) equals how Many Cells that are filled in the range y
        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available row in TILE Sheet
        
        Sheets("Receiver").Range("B" & a & ":E" & a).Copy Sheets(ws).Range("B" & DestRow & ": E" & DestRow) 'Copies section from Receiver to TILE
        Sheets("Receiver").Range("G" & a & ":I" & a).Copy Sheets(ws).Range("F" & DestRow & ": H" & DestRow) 'Copies section from Receiver to TILE
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section from Receiver to TILE
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("J" & DestRow) 'Copies section from Receiver to TILE
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("I" & DestRow) 'Copies section from ReCeiver to TILE


        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1


    Loop
ElseIf ws = "BASE" Then
    a = 51
    
    If Sheets("Receiver").Range("B51") = "n/a" Then 'If B51 has the value n/a then it will end sub
        Exit Sub
    End If
    
    Do Until n = z 'loops until counter (n) equals how many cells that are filled in the range y
        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available row in BASE Sheet
        
        Sheets("Receiver").Range("B" & a & ":E" & a).Copy Sheets(ws).Range("B" & DestRow & ": E" & DestRow) 'Copies section from Receiver to Base
        Sheets("Receiver").Range("G" & a).Copy Sheets(ws).Range("F" & DestRow) 'Copies section from Receiver to Base
        Sheets("Receiver").Range("I" & a).Copy Sheets(ws).Range("G" & DestRow) 'Copies section from Receiver to Base
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section froM Receiver to Base
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("I" & DestRow) 'Copies section froM ReCeiver to Base
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("H" & DestRow) 'Copies section froM Receiver to Base


        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1


    Loop
    
ElseIf ws = "SUNDRIES" Then
    a = 61
    
    If Sheets("Receiver").Range("B61") = "n/a" Then 'If B61 has the value n/a then it will end sub
        Exit Sub
    End If
    
    Do Until n = z 'loops until counter (n) equals how many cells that are filled in the range y
        DestRow = Sheets(ws).Cells(Rows.Count, "A").End(xlUp).Row + 1 'gets next available row in Sundries Sheet
        
        Sheets("Receiver").Range("H1").Copy Sheets(ws).Range("A" & DestRow) 'Copies section from Receiver to Sundries
        Sheets("Receiver").Range("B" & a & ":D" & a).Copy Sheets(ws).Range("B" & DestRow & ": D" & DestRow) 'Copies section from Receiver to Sundries
        Sheets("Receiver").Range("H" & a & ":I" & a).Copy Sheets(ws).Range("E" & DestRow & ":F" & DestRow) 'Copies section from Receiver to Sundries
        Sheets("Receiver").Range("H2").Copy Sheets(ws).Range("I" & DestRow) 'Copies section from Receiver to Sundries
        Sheets("Receiver").Range("H3").Copy Sheets(ws).Range("H" & DestRow) 'Copies section from Receiver to Sundries
        
        n = n + 1 'Increment n by 1
        a = a + 1 'Increment a by 1

    Loop
End If



End Sub

'By Matthew Connolly
'Modified: 3/31/2016

Private Sub Main()


    Call CopyReceiver2X("CARPET", "B11:B20") 'Copies ReCeiver to Carpet

    Call CopyReceiver2X("CPT.TILE", "B21:B30") 'Copies Receiver to CPT.Tile

    Call CopyReceiver2X("VINYL", "B31:B40") ' Copies Receiver to VINYL

    Call CopyReceiver2X("TILE", "B41:B50") 'Copies Receiver to TILE

    Call CopyReceiver2X("BASE", "B51:B60") ' Copies Receiver to Base

    Call CopyReceiver2X("SUNDRIES", "B61:B73") ' Copies ReCeiver to Sundries

    Debug.Print (MsgBox("Receiver has been copied to the Logs!"))  ' Prints out a Message to the screen saying that the logs were copied over
    ActiveSheet.PrintOut ' sends the active sheet to a printer to print
    Debug.Print (MsgBox("Printing Receiver...")) ' Displays a message saying that the Receiver is printing from printer                                                                                                                                                                                                                                 this code can not be modified in any way shape or form unless notified by Matthew Connolly and approved by Matthew Connolly. This Code Belongs to Matthew Connolly and as all rights to this code. This code can not be coppied to a new sheet or modified on this sheet.
    


End Sub

Private Sub Location()

Debug.Print (Sheets("CPT.TILE").Cells(Rows.Count, "A").End(xlUp).Row)

End Sub

