VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Report 
   Caption         =   "All_Parts    v 1.1.01"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14490
   OleObjectBlob   =   "Report.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim FILENAME
Dim PLISNs As New Collection
Dim HEADER As String
Dim LOADFILE As String
'''Begin the card variants
Dim SAP As New Collection
Dim allparts As New Collection

'===================================================================================================
'GLOBAL STUFF
'===================================================================================================
Function SortParts(list As Variant, combob As Variant)
    'Uses the Sort spreadsheet to sort part lists
    On Error Resume Next
    Dim last As Variant
    Dim part As Variant
    Dim i As Long
    
    'Keeps screen from updating so you don't see this stuff
    Application.ScreenUpdating = False
    Columns("A:A").Clear
    
    'Creates a list big enough to hold anything
    Dim nlist(0 To 50000) As String
    
    'sort sheet select and add list
    Sheets("Sort").Select
    Range("A1:A" & UBound(combob.list)).NumberFormat = "@"
    Range("A1:A" & UBound(combob.list)) = combob.list
    
    'sort everything
    ActiveWorkbook.Worksheets("sort").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sort").sort.SortFields.Add Key:=Range("A:A"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("sort").sort
        .SetRange Range("A:A")
        .HEADER = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'this will make sure no duplicates are added
    last = Cells(1, 1)
    nlist(0) = last
    i = 2
    
    Do While Not IsEmpty(Cells(i, 1))
        'double check this is not a duplicate
        If Not Cells(i, 1) = last Then
            'collect items
            nlist(i - 1) = Cells(i, 1)
        End If
        'no duplicates
        last = Cells(i, 1)
        i = i + 1
    Loop
    
    'Clear column for next time
    Columns("A:A").Clear
    
    'Clear combobox and add items collected
    combob.Clear
    combob.Value = ""
    For Each part In nlist
        'make sure only add non blank items
        If Not part = "" Then
            combob.AddItem (part)
        End If
    Next
    
    'Select the main sheet and let things be seen again
    Sheets("Main").Select
    Application.ScreenUpdating = True
End Function
Function Updatepartslist(TextLine As String, cagePart As String)
    'Creates a collection of unique part numbers and info unique to part number
    Dim ap As New parts
    'Errors exist when the part doesn't exist in the collection already
    On Error GoTo 1
    
    'collect acard stuffs
    If Right(TextLine, 3) = "01A" Then
        Dim part As Variant
        Dim row As Integer
        Dim plisn As String
        
        Set ap = New parts
        plisn = Trim(Mid(TextLine, 7, 5))
        
        'update the part and the count of the part
        allparts(cagePart).count = allparts(cagePart).count + 1 'Error when new part
        'Add this plisn to the list of plisns with this part number
        allparts(cagePart).aplisns.Add Item:=plisn
        'Goto 2 because this aint my first rodeo (all the following info was added the first time)
        GoTo 2
1       If Right(TextLine, 3) = "01A" Then
            ap.cagePart = Trim(Mid(TextLine, 14, 37))
            ap.cage = Trim(Mid(TextLine, 14, 5))
            ap.part = Trim(Mid(TextLine, 19, 32))
            ap.descrip = Trim(Mid(TextLine, 56, 19))
            allparts.Add Item:=ap, Key:=(ap.cagePart)
            allparts(ap.cagePart).count = 1
            ap.aplisns.Add Item:=plisn, Key:=(plisn)
        End If
        
2   'collect bcard stuffs
    ElseIf Right(TextLine, 3) = "01B" Then
        allparts(cagePart).nsn = Trim(Mid(TextLine, 16, 13))
    
    End If
End Function
Function displaylistbox2()
    'Show everything gathered in allparts collection for part numbers
    Dim row As Integer
    Dim part As Variant
    On Error GoTo 1
    
    'add each part in allparts collection o listbox2 on part number page
    For Each part In allparts
        ListBox2.AddItem ""
        ListBox2.Column(0, row) = part.part
        ListBox2.Column(1, row) = part.cage
        ListBox2.Column(2, row) = part.descrip
        ListBox2.Column(3, row) = part.nsn
        ListBox2.Column(4, row) = part.count
        row = row + 1
    Next part
    GoTo 2
1   'Error
2
End Function
Sub RemoveDuplicates(combo As ComboBox, remove As Variant, qty As Integer)
    'This removes duplicates from collection
    Dim part As Variant
    Dim CO As Object
    Set CO = CreateObject("Scripting.Dictionary")
    
    With CO
        For Each part In combo.list
            If Not .exists(part) Then
                .Add part, 1
            End If
        Next part
        combo.Value = ""
        combo.Clear
        combo.list = .keys
    End With
End Sub

Private Sub CommandButton23_Click()
    '
    UpdateCardsCSN
End Sub




Function FixLength(tb As Variant)
    'Fixes the length of a given textbox value to equal what is needed in the 036 report
    Dim temp As String
    
    If Left(tb.Tag, 1) = 0 Then
        temp = Application.WorksheetFunction.Rept("0", tb.MaxLength - Len(tb.Value)) & tb.Value
    ElseIf Trim(Left(tb.Tag, 1)) = "" Then
        temp = tb.Value & Application.WorksheetFunction.Rept(" ", tb.MaxLength - Len(tb.Value))
    End If
    FixLength = temp
End Function
Function eachcard(r As Variant, plisn As String)
    'do while for each plisn to update all cards with new part information
    Dim TL As String
    
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
        
        TL = ListBox1.list(r)
        If Right(TL, 3) = "01A" Then
            ListBox1.list(r) = UCase(Left(TL, 13) & FixLength(TextBox165) & FixLength(TextBox164) & Mid(TL, 51, 5) & _
                FixLength(TextBox166) & Mid(TL, 75, 6))
        ElseIf Right(TL, 3) = "01B" Then
            ListBox1.list(r) = UCase(Left(TL, 12) & FixLength(TextBox172) & FixLength(TextBox171) & _
                FixLength(TextBox170) & FixLength(TextBox169) & FixLength(TextBox168) & _
                FixLength(TextBox167) & FixLength(TextBox174) & FixLength(TextBox173) & Mid(TL, 62, 19))
        ''SAP calculations every time instead
        'ElseIf Right(TL, 3) = "01C" Then
            'If Not plisn = ListBox3.list(0) Then
                'ListBox1.list(r) = Left(TL, 59) & ListBox3.list(0) & application.WorksheetFunction.rept(" ", 5 - Len(ListBox3.list(0))) & Mid(TL, 65, 16)
            'Else
                'ListBox1.list(r) = Left(TL, 59) & "     " & Mid(TL, 65, 16)
            'End If
        ElseIf Right(TL, 3) = "01H" Then
            If TextBox175.Visible = True Then
                ListBox1.list(r) = UCase(Left(TL, 32) & FixLength(TextBox175) & Mid(TL, 78, 3))
            End If
        ElseIf Right(TL, 3) = "01K" Then
            If TextBox181.Visible = True Then
                ListBox1.list(r) = UCase(Left(TL, 23) & FixLength(TextBox181) & Mid(TL, 78, 3))
            End If
        End If
        r = r + 1
    Loop

End Function
Function partupdate(this As Boolean)
    'Updates the part selected after clicking to update all or only that specific
        'plisn on the part number page
    'On Error GoTo 1
    Dim plisn As String
    Dim r As Integer
    Dim temp As String
    Dim a As Variant
    Dim findit As String
    Dim s As Variant
    Dim NHA As String
    Dim cagePart As String
    
    cagePart = ListBox2.Column(1, ListBox2.ListIndex) & ListBox2.Column(0, ListBox2.ListIndex)
   
    r = ListBox1.ListIndex
    
    plisn = ListBox3.list(ListBox3.ListIndex)
    
    If plisn = "" Then Exit Function
    
    If this Then
        'Check the NHA and do this for each of those assemblies
        NHA = checkNHA(ListBox1.ListIndex)
        Dim pl As Variant
        Dim indenture As Integer
        Dim found As Boolean
        
        
        For pl = 0 To ListBox1.ListCount - 1
            If Trim(Mid(ListBox1.list(pl), 14, 37)) = NHA Then
                found = False
                Do While found = False
                    pl = pl + 1
                    If cagePart = Trim(Mid(ListBox1.list(pl), 14, 37)) Then
                        'Update this card at this index
                        plisn = Trim(Mid(ListBox1.list(pl), 7, 5))
                        eachcard pl, plisn
                        found = True
                    End If
                    If Right(ListBox1.list(pl), 3) = "01A" Then
                        If Not indenture < Asc(Mid(ListBox1.list(pl), 13, 1)) Then
                            'Future idea: Wasn't part of assembly, would you like to add it?
                            MsgBox cagePart & " Not found in this assembly " & NHA & "!"
                            GoTo 4
                        End If
                    End If
                Loop
4
            End If
        Next pl
    Else
        For Each a In ListBox3.list
            If IsNull(a) Then GoTo 3
            ComboBox1.Value = a
            eachcard ListBox1.ListIndex, ComboBox1.Value
        Next a
    End If
3
    Update_PN_List
    GoTo 2
1   'Error
2   ComboBox1.Value = plisn
End Function
Function SAP_Calc()
    'Calculate sap after moving and editing part numbers
    Dim r As Variant
    Dim SAP As String
    Dim NHA As String
    Dim plisn As String
    
    For r = 0 To ListBox1.ListCount - 1
        plisn = Trim(Mid(ListBox1.list(r), 7, 5))
        If Right(ListBox1.list(r), 3) = "01A" Then
            SAP = checkPLISNs(r)
            If Not SAP = plisn Then
                Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
                    If r >= ListBox1.ListCount - 1 Then GoTo 1
                    If Right(ListBox1.list(r), 3) = "01C" Then
                        ListBox1.list(r) = Left(ListBox1.list(r), 59) & _
                            SAP & Application.WorksheetFunction.Rept(" ", 5 - Len(SAP)) & _
                            Right(ListBox1.list(r), 16)
                    End If
                    r = r + 1
                Loop
            Else
                ListBox1.list(r) = Left(ListBox1.list(r), 59) & _
                    SAP & "     " & _
                    Right(ListBox1.list(r), 17)
            End If
        End If
    Next r
1
End Function
Function checkPLISNs(r As Variant)
    'Checks if the cage and part number combination is in the same as plisn collection yet
    On Error GoTo 1
    Dim cagePart As String
    Dim plisn As String
    cagePart = Trim(Mid(ListBox1.list(r), 14, 37))
    plisn = Trim(Mid(ListBox1.list(r), 7, 5))
    
    SAP.Add Item:=plisn, Key:=cagePart
    
    GoTo 2
1   'the part existed in the collection
    checkPLISNs = SAP.Item(cagePart)
    GoTo 3
2
    checkPLISNs = ""
    
3
End Function

Private Sub ComboBox10_Change()
    'Available PLISNs on add plisn page
    If Len(ComboBox10.Value) = 4 Or Len(ComboBox10.Value) = 5 Then
        CommandButton39.Caption = "Part as " & ComboBox10.Value
        CommandButton44.Caption = "Down part as " & ComboBox10.Value
    Else:
        CommandButton39.Caption = ""
        CommandButton44.Caption = ""
    End If
End Sub



Private Sub ComboBox11_Change()
    If ComboBox11.Value = "" Then
        TextBox199.Value = ""
        ComboBox10.Clear
        ComboBox10.Value = ""
        Exit Sub
    Else:
        ComboBox3_Change
    End If
    
End Sub

Private Sub ComboBox8_Change()
    If ListBox1.ListCount > 0 Then
        ComboBox2.Value = ComboBox8.Value
        TextBox195.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 56, 19))
    End If
End Sub


Function updateplisngap(box As ComboBox)
    On Error GoTo 1
    Dim r As Variant
    Dim temp As String
    Dim nextplisn As String
    
    ComboBox10.Clear
    ComboBox10.Value = ""
    
    'updates textbox200 to a value if none
    If TextBox200.Value = "" Then TextBox200.Value = 1
    TextBox199.Value = 0
    'error check last item in combobox then
    temp = NumOfPLISNS(box.Value)
    nextplisn = box.list(box.ListIndex + 1)
    TextBox202.Value = nextplisn
    
3   Do While convertplisn(temp) < convertplisn(nextplisn)
        
        ComboBox10.AddItem (temp)
        temp = NumOfPLISNS(temp)
        
        'increment how many there can be
        TextBox199.Value = TextBox199.Value + 1
        
        'after 1000 plisns got to the end because this is most likely the last plisn
        If TextBox199.Value > 1000 Then
            GoTo 2
        End If
    Loop
    GoTo 2
1   nextplisn = "ZZZZ"
    GoTo 3
2
End Function
Function NumOfPLISNS(currplisn As String)
    On Error GoTo 1
    Dim B As Variant
    Dim a As Integer
    Dim upto As Integer
    Dim r As Variant
    Dim plisn As Variant
    Dim numop As Integer
    Dim nextplisn As String
    Dim start As Integer
    
    B = Array(11, 25, 35)
    If ComboBox11.Value = "Alpha-Numeric" Then
        upto = B(2)
        start = 65
    ElseIf ComboBox11.Value = "Alpha" Then
        upto = B(1)
        start = 65
    ElseIf ComboBox11.Value = "Numeric" Then
        upto = B(0)
        start = 90
    End If
    'if numeric then start is 90
    
    'figures out if there are 4 or 5 digits in plisn, usually 4
    If Len(currplisn) = 5 Then
        plisn = Array(Asc(Mid(currplisn, 1, 1)), _
                    Asc(Mid(currplisn, 2, 1)), _
                    Asc(Mid(currplisn, 3, 1)), _
                    Asc(Mid(currplisn, 4, 1)), _
                    Asc(Mid(currplisn, 5, 1)))
        numop = 4
    ElseIf Len(currplisn) = 4 Then
        plisn = Array(Asc(Mid(currplisn, 1, 1)), _
                    Asc(Mid(currplisn, 2, 1)), _
                    Asc(Mid(currplisn, 3, 1)), _
                    Asc(Mid(currplisn, 4, 1)))
        numop = 3
    End If
    
    'adds 43 to all numbers to make numbers worth more than letters
    For r = 0 To numop
        If plisn(r) < start Then
            plisn(r) = plisn(r) + 43
        End If
    Next r
    
    'increments by the gap
    plisn(numop) = plisn(numop) + TextBox200.Value
    r = 0
    For r = 0 To numop - 1
        Do While plisn(numop - r) > upto + start
            plisn(numop - r) = plisn(numop) - upto - 1
            plisn(numop - r - 1) = plisn(numop - r - 1) + 1
        Loop
        If Chr(plisn(numop - r)) = "O" Or Chr(plisn(numop - r)) = "I" Then
            plisn(numop - r) = plisn(numop - r) + 1
        End If
    Next r
    r = 0
    For r = 0 To numop
        If plisn(r) > 90 Then
            plisn(r) = plisn(r) - 43
        End If
        nextplisn = nextplisn + Chr(plisn(r))
    Next r
    
    NumOfPLISNS = nextplisn
1
End Function
Function nextdigitup(digits As Variant)
    Dim count As Integer
    Dim temp As Integer
    
    temp = TextBox200.Value
    
    Do While temp > digits
        temp = temp - digits
        count = count + 1
    Loop
    nextdigitup = temp
End Function
Private Sub CommandButton29_Click()
    'update all plisns with this cage part combo
    partupdate False
    
    MultiPage1.Value = 0
    MultiPage1.Value = 2
End Sub

Private Sub CommandButton32_Click()
    '''Opens the NHA PLISN in another box
    'get error when the plisn is taller than the listbox3 height
    On Error Resume Next
    Dim plisn As String
    Dim nor As Integer
    Dim r As Variant
    If TextBox178.Value = "" Then Exit Sub
    For r = 0 To ListBox1.ListIndex
        If Trim(Mid(ListBox1.list(ListBox1.ListIndex - r), 7, 5)) = TextBox178.Value Then
            plisn = ListBox1.list(ListBox1.ListIndex - r)
            r = r + 1
            nor = 1
            Do While Trim(Mid(ListBox1.list(ListBox1.ListIndex - r), 7, 5)) = TextBox178.Value
                If Not ListBox1.list(ListBox1.ListIndex - r) = "" Then
                    plisn = ListBox1.list(ListBox1.ListIndex - r) & vbCr & plisn
                    nor = nor + 1
                    r = r + 1
                End If
            Loop
            GoTo 1
        End If
    Next r
        
1
    Frame97.Height = nor * 14 + 6
    Frame96.Height = nor * 14 + 21
    Frame96.Caption = TextBox178 & " is the Next Higher Assembly of " & ListBox3.list(ListBox3.ListIndex)
    TextBox184.Height = (nor * 14) - 6
    Frame96.Visible = True
    TextBox184.Value = plisn
    ListBox3.Top = nor * 14 + 8
    ListBox3.Height = (143.95 - nor * 14)
    'ListBox3.TopIndex = ListBox3.LISTINDEX
    ListBox3.SetFocus
    ListBox3.TopIndex = ListBox3.ListIndex
End Sub

Private Sub CommandButton33_Click()
    'update only this plisn on part number page
    Dim plisn As String
    
    partupdate True

    MultiPage1.Value = 0
    MultiPage1.Value = 2
    
End Sub
Private Sub CommandButton34_Click()
    'Closes the NHA window on the part number page
    Frame96.Visible = False
    ListBox3.Top = 4.5
    ListBox3.Height = 143.95
End Sub

Private Sub CommandButton35_Click()
    'goto NHA in the part number pages NHA window after clicking the ? button
    ComboBox1.Value = TextBox178.Value
    MultiPage1.Value = 0
    MultiPage1.Value = 2
    Frame96.Visible = False
    ListBox3.Top = 4.5
    ListBox3.Height = 143.95
End Sub

Private Sub CommandButton36_Click()
    'Raw edit return button from main page
    Frame98.Visible = False
End Sub

Private Sub CommandButton37_Click()
    '
    On Error GoTo 1
    Dim plisn As String
    Dim cards As String
    Dim r As Integer
    Dim currentPLISN As String
    currentPLISN = ComboBox3.Value
    
    r = ListBox1.ListIndex
    plisn = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    If plisn = "" Then
        r = r - 1
    Else
        Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
            If Not ListBox1.list(r) = "" Then
                r = r - 1
            End If
        Loop
    End If
    r = r + 1
    
    Dim lines() As String
    lines = Split(TextBox186.Text, vbCrLf)
    Dim count As Integer
    Dim line As Variant
    For Each line In lines
        count = count + 1
    Next line
    If Not count = TextBox187.Value Then
        MsgBox ("You may not add or delete lines here" & _
            Chr(10) & _
            Chr(10) & "Please edit lines only"), vbInformation
        GoTo 2
    End If
    
    For Each line In lines
        ListBox1.list(r) = line
        r = r + 1
    Next line
    Frame98.Visible = False
    Update_PN_List
    Dim curp As Integer
    curp = MultiPage1.Value
    MultiPage1.Value = 0
    MultiPage1.Value = curp
    
    ComboBox3.Value = ""
    ComboBox3.Value = currentPLISN
    GoTo 2
1   MsgBox ("Please select a line to edit"), vbInformation
2
End Sub

Private Sub CommandButton38_Click()
    'Toggle visibility of frame with the textbox that shows the PLISN selected in listbox3,
        'PLISNs, on the part number page
    If Frame98.Visible = False Then
        CommandButton26_Click
    Else
        Frame98.Visible = False
    End If
    
End Sub

Private Sub CommandButton39_Click()

End Sub





Private Sub CommandButton42_Click()
    'Goto NHA PLISN on Assemblies page
    If Len(CommandButton42.Caption) > 5 Then
        ComboBox3.Value = Right(CommandButton42.Caption, 4)
    End If
    If ListBox4.ListCount > 0 Then
        Label137.Caption = "Indenture " & Left(ListBox4.list(0), 2)
    Else:
        Label137.Caption = ""
    End If
    
End Sub

Function DeleteAssembly(deleteCagePart As String, howMany As Integer)
    'When a assembly is deleted, delete the downparts too
    Dim r As Variant
    Dim indenture As String
    Dim Delete As Boolean
    Dim s As Variant
    Dim start As Integer
    Dim NHACagePart As String
    
    NHACagePart = checkNHA(ListBox1.ListIndex)
    
    If howMany < 2 Then
        start = ListBox1.ListIndex
    End If
    'For s = 0 To checkNHA
        
        'go through listbox1 selected assemblies and deleting them
        For r = start To ListBox1.ListCount - 1
        
            If r > ListBox1.ListCount - 1 Then GoTo 1
            If deleteCagePart = Trim(Mid(ListBox1.list(r), 14, 32)) Then
                If howMany = 2 Then
                    'If this is the part of the same assembly
                    If Not checkNHA(r) = NHACagePart Then
                        GoTo 2
                    End If
                End If
                
                Delete = True
                indenture = Mid(ListBox1.list(r), 13, 1)
                
                Do While Delete = True And r < ListBox1.ListCount - 1
                    ListBox1.RemoveItem (r)
                    If Right(ListBox1.list(r), 3) = "01A" Then
                        If Not Asc(Mid(ListBox1.list(r), 13, 1)) > Asc(indenture) Then
                            Delete = False
                        End If
                    End If
                Loop
                
                If howMany < 2 Then
                    GoTo 1
                End If
2
            End If
        Next r
   'Next s
   MsgBox "what"
1   ListBox4.Clear
    Update_PN_List

End Function
Function checkNHA(r As Variant)
    Dim s As Variant
    Dim NHA As String
    
    For s = r To ListBox1.ListCount - 1
        If Right(ListBox1.list(s), 3) = "01C" Then
            NHA = Trim(Mid(ListBox1.list(s), 13, 5))
            GoTo 1
        End If
    Next s
1
    For s = 0 To r
        If Trim(Mid(ListBox1.list(s), 7, 5)) = NHA Then
            checkNHA = Trim(Mid(ListBox1.list(s), 14, 37))
            Exit Function
        End If
    Next s

End Function
Function CheckNumberOfNHA()
    'When a part is deleted, check its nha
    'Checks if there is more than one of the same assemblies this part belongs to
    Dim NHA As String
    Dim r As Integer
    Dim cagePart As String
    
    NHA = Trim(Right(CommandButton42.Caption, 5))
    
    Do While Not ListBox1.ListIndex - r = 0
        If Trim(Mid(ListBox1.list(r), 7, 5)) = NHA And Right(ListBox1.list(r), 3) = "01A" Then
            cagePart = Trim(Mid(ListBox1.list(r), 14, 37))
            If allparts(cagePart).count > 1 Then
                'This means we have to recalculate the NHA assemblies because we changed something on its downparts
                checkNHA = allparts(cagePart).count
                Exit Function
            End If
        End If
        
        r = r + 1
    Loop
    
End Function

Private Sub CommandButton43_Click()
    'Delete parts or assemblies on the assemblies page
    If OptionButton8.Value = True Then
        DeleteAssembly Trim(Mid(ListBox1.list(ListBox1.ListIndex), 14, 37)), 1
    ElseIf OptionButton9.Value = True Then
        DeleteAssembly Trim(Mid(ListBox1.list(ListBox1.ListIndex), 14, 37)), 2
    ElseIf OptionButton10.Value = True Then
        DeleteAssembly Trim(Mid(ListBox1.list(ListBox1.ListIndex), 14, 37)), 3
    End If
End Sub

Private Sub Label130_Click()

End Sub

Private Sub Label139_Click()

End Sub

Private Sub ListBox2_Change()
    'listbox with part number information on part number page
    'On error incase empty listbox1
    On Error Resume Next
    Dim r As Variant
    ListBox3.Clear
    Dim plisn As Variant
    Dim i As Variant
    For Each plisn In allparts(ListBox2.Column(1) & ListBox2.Column(0)).aplisns
        ListBox3.AddItem (plisn)
    Next plisn
    ListBox3.ListIndex = 0
    
End Sub

Private Sub ListBox3_Change()
    'PLISNs listbox on part number page
    Dim kcard As Boolean
    Dim hcard As String
    Dim boxes As Variant
    Dim tb As Variant
    Dim plisn As String
    
    boxes = Array(TextBox165, TextBox164, TextBox172, TextBox171, TextBox170, TextBox182, _
                TextBox166, TextBox175, TextBox181, TextBox169, TextBox168, TextBox167, _
                TextBox174, TextBox173, TextBox178, TextBox179, TextBox176, TextBox177)
    For Each tb In boxes
        tb.Value = ""
    Next tb
    
    If ListBox3.ListIndex >= 0 Then
        If Not ListBox3.list(ListBox3.ListIndex) = "" Then
            ComboBox1.Value = ListBox3.list(ListBox3.ListIndex)
            'update textboxes
            Dim r As Variant
            For r = ListBox1.TopIndex To ListBox1.ListCount - 1
                If Trim(Mid(ListBox1.list(r), 7, 5)) = ListBox3.list(ListBox3.ListIndex) Then
                    If Right(ListBox1.list(r), 3) = "01A" Then
                        TextBox165 = Trim(Mid(ListBox1.list(r), 14, 5))
                        TextBox164 = Trim(Mid(ListBox1.list(r), 19, 32))
                        TextBox166 = Trim(Mid(ListBox1.list(r), 56, 19))
                    ElseIf Right(ListBox1.list(r), 3) = "01B" Then
                        TextBox172 = Trim(Mid(ListBox1.list(r), 13, 3))
                        TextBox171 = Trim(Mid(ListBox1.list(r), 16, 13))
                        TextBox170 = Trim(Mid(ListBox1.list(r), 29, 4))
                        TextBox169 = Trim(Mid(ListBox1.list(r), 33, 2))
                        TextBox168 = Trim(Mid(ListBox1.list(r), 35, 10))
                        TextBox167 = Trim(Mid(ListBox1.list(r), 45, 2))
                        TextBox174 = Trim(Mid(ListBox1.list(r), 47, 10))
                        TextBox173 = Trim(Mid(ListBox1.list(r), 57, 5))
                        TextBox179 = Trim(Mid(ListBox1.list(r), 65, 5))
                    ElseIf Right(ListBox1.list(r), 3) = "01C" Then
                        TextBox178 = Trim(Mid(ListBox1.list(r), 13, 5))
                        TextBox182 = Trim(Mid(ListBox1.list(r), 60, 5))
                    ElseIf Right(ListBox1.list(r), 3) = "01H" Then
                        hcard = Trim(Mid(ListBox1.list(r), 33, 45))
                        
                    ElseIf Right(ListBox1.list(r), 3) = "01K" Then
                        kcard = True
                        
                        TextBox175.Visible = False
                        TextBox181.Visible = True
                        Label107.Caption = "Provisioning Nomenclature"
                        TextBox181 = Trim(Mid(ListBox1.list(r), 24, 54))
                        
                    ElseIf Right(ListBox1.list(r), 3) = "01J" Then
                        If Not Trim(Mid(ListBox1.list(r), 16, 4)) = "" And Not Trim(Mid(ListBox1.list(r), 20, 4)) = "" Then
                            TextBox176 = Trim(Mid(ListBox1.list(r), 16, 4))
                            TextBox177 = Trim(Mid(ListBox1.list(r), 20, 4))
                        End If
                    ElseIf Right(ListBox1.list(r), 3) = "01K" Then
                        If Not Trim(Mid(ListBox1.list(r), 16, 4)) = "" And Not Trim(Mid(ListBox1.list(r), 20, 4)) = "" Then
                            TextBox176 = Trim(Mid(ListBox1.list(r), 16, 4))
                            TextBox177 = Trim(Mid(ListBox1.list(r), 20, 4))
                        End If
                    End If
                End If
            Next r
        End If
    End If
    If Not kcard Then
        TextBox175.Visible = True
        TextBox181.Visible = False
        Label107.Caption = "Remarks"
        TextBox175 = hcard
    End If
    
End Sub

Private Sub ListBox4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'assemblies page
    On Error GoTo 1
    ComboBox3.Value = Trim(Mid(ListBox4.list(ListBox4.ListIndex), 2, 5))
    ListBox4.Selected(0) = True
    If ListBox4.ListCount > 0 Then
        Label137.Caption = "Indenture " & Left(ListBox4.list(0), 2)
    Else:
        Label137.Caption = ""
    End If
1
End Sub

Private Sub MultiPage1_Change()
    'When the main tab changes do stuffs
    'if no file is loaded then error
    On Error GoTo 1
    Dim s As Integer
    Dim r As Variant
    
    SpinButton8.Visible = False
    Frame42.Visible = False
    If MultiPage1.Value = 1 Then
        'edit cards page
        UpdateCardsCSN
        Frame98.Visible = False
        updateEditor
    ElseIf MultiPage1.Value = 0 Then
        'main page
        SpinButton8.Visible = True
        Frame42.Visible = True
    ElseIf MultiPage1.Value = 2 Then
        'part number page
        Dim findit As String
        findit = ComboBox2.Value
        
        'if the a card isn't selected then locate the a card to get the part number
        s = ListBox1.ListIndex
        Do While Trim(Mid(ListBox1.list(s - 1), 7, 5)) = ComboBox1.Value
            s = s - 1
        Loop
        If Right(ListBox1.list(s), 3) = "01A" Then
            findit = Trim(Mid(ListBox1.list(s), 14, 37))
        Else
            findit = ""
        End If
            
        '''update findit to equal the cage code and the part number
        If Not findit = "" Then
            For r = 0 To ListBox2.ListCount - 1
                If ListBox2.Column(1, r) & ListBox2.Column(0, r) = findit Then
                    ListBox2.ListIndex = r
                    ListBox2.Selected(r) = True
                    Exit Sub
                End If
            Next r
        Else
            ListBox2.ListIndex = -1
            ListBox2.MultiSelect = fmMultiSelectExtended
            ListBox2.MultiSelect = fmMultiSelectSingle
        End If
    ElseIf MultiPage1.Value = 3 Then
        If ListBox4.ListCount > 0 Then
            Label137.Caption = "Indenture " & Left(ListBox4.list(0), 2)
        Else:
            Label137.Caption = ""
        End If
    End If
    GoTo 2
1   'Error
2
End Sub
Function updateEditor()
    'update the edi cards page, only once
    If TextBox20.Value = "01" Then
        UpdateCardsCSN
    Else
        '''making it 01 will update cards
        TextBox20.Value = "01"
    End If
End Function
Private Sub MultiPage2_Change()
    '''if not 01 then make it 01
    updateEditor
End Sub

Private Sub OptionButton6_Click()
    'Update PLISN
    If OptionButton6.Value = True Then
        ComboBox1.Value = ComboBox3.Value
        ComboBox3_Change
    End If
End Sub

Private Sub SpinButton17_SpinUp()
    'Find the next thing in the description search box on the main page
    searchHereOn TextBox66.Value
End Sub
Private Sub SpinButton17_SpinDown()
    'Find the previous thing in the description search box on the main page
    searchHereBack TextBox66.Value
End Sub
Private Sub SpinButton18_SpinDown()
    'Find previous part number on main page
    searchHereBack ComboBox2.Value & Application.WorksheetFunction.Rept(" ", 32 - Len(ComboBox2.Value))
End Sub
Private Sub SpinButton18_SpinUp()
    'Find next part number on main page
    searchHereOn ComboBox2.Value & Application.WorksheetFunction.Rept(" ", 32 - Len(ComboBox2.Value))
End Sub
Private Sub SpinButton19_SpinDown()
    'Search for previous NSN on main page
    searchHereBack ComboBox6.Value & Application.WorksheetFunction.Rept(" ", 13 - Len(ComboBox6.Value))
End Sub

Private Sub SpinButton19_SpinUp()
    'Search for next NSN on main page
    searchHereOn ComboBox6.Value & Application.WorksheetFunction.Rept(" ", 13 - Len(ComboBox6.Value))
End Sub
Private Sub SpinButton20_SpinDown()
    'Find previous part number on main page
    searchHereBack ComboBox2.Value & Application.WorksheetFunction.Rept(" ", 32 - Len(ComboBox2.Value))
End Sub
Private Sub SpinButton20_SpinUp()
    'Find next part number on main page
    searchHereOn ComboBox2.Value & Application.WorksheetFunction.Rept(" ", 32 - Len(ComboBox2.Value))
End Sub
Private Sub SpinButton21_SpinDown()
    'Search for previous NSN on main page
    searchHereBack ComboBox6.Value & Application.WorksheetFunction.Rept(" ", 13 - Len(ComboBox6.Value))
End Sub

Private Sub SpinButton21_SpinUp()
    'Search for next NSN on main page
    searchHereOn ComboBox6.Value & Application.WorksheetFunction.Rept(" ", 13 - Len(ComboBox6.Value))
End Sub
Function searchHereOn(match As String)
     'Searchs listbox1 for a match on the main page from current position,
        'foward, and then loops around if not found
    If ListBox1.ListIndex < 0 Then Exit Function
    'use this check box not to update the description box while we search for things.
    CheckBox22 = True
    
    Do While Not ListBox1.list(ListBox1.ListIndex) = ""
        If ListBox1.ListIndex = ListBox1.ListCount - 1 Then ListBox1.ListIndex = 0
        ListBox1.ListIndex = ListBox1.ListIndex + 1
    Loop
    
    Dim r As Variant
    For r = ListBox1.ListIndex To ListBox1.ListCount - 1
        If InStr(1, ListBox1.list(r), match) > 0 Then
            GoTo 1
        End If
    Next r
    For r = 0 To ListBox1.ListIndex
        If InStr(1, ListBox1.list(r), match) > 0 Then
            GoTo 1
        End If
    Next r
1   ComboBox1.Value = Trim(Mid(ListBox1.list(r), 7, 5))
    ListBox1.TopIndex = r
    CheckBox22 = False
    countselected
    ListBox1.ListIndex = r
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value
        r = r - 1
    Loop
    ListBox1.TopIndex = r + 1
End Function

Function searchHereBack(match As String)
    'Searchs listbox1 for a match on the main page from current position,
        'backward, and then loops around if not found
    Dim r As Variant
    If ListBox1.ListIndex < 0 Then Exit Function
    'use this check box not to update the description box while we search for things.
    CheckBox22 = True
    
    Do While Not ListBox1.list(ListBox1.ListIndex) = ""
        If ListBox1.ListIndex = ListBox1.ListCount - 1 Then ListBox1.ListIndex = 0
        ListBox1.ListIndex = ListBox1.ListIndex - 1
    Loop
    
    
    For r = 0 To ListBox1.ListIndex
        If InStr(1, ListBox1.list(ListBox1.ListIndex - r), match) > 0 Then
            ListBox1.TopIndex = ListBox1.ListIndex - r
            ListBox1.ListIndex = ListBox1.ListIndex - r
            GoTo 1
        End If
    Next r
    For r = 0 To ListBox1.ListCount - 1 - ListBox1.ListIndex
        If InStr(1, ListBox1.list(ListBox1.ListCount - 1 - r), match) > 0 Then
            ListBox1.TopIndex = ListBox1.ListCount - 1 - r
            ListBox1.ListIndex = ListBox1.ListCount - 1 - r
            GoTo 1
        End If
    Next r
1   ComboBox1.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    CheckBox22 = False
    countselected
    r = ListBox1.ListIndex
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value
        r = r - 1
    Loop
    ListBox1.TopIndex = r + 1
End Function


Private Sub SpinButton2_SpinDown()
    ''Changes the CSN textbox on edit card page
    If Not TextBox20.Value = "01" Then
        TextBox20.Value = "0" & SpinButton2.Value
    Else:
        TextBox20.Value = "06"
    End If
    SpinButton2.Value = Int(TextBox20.Value)
End Sub

Private Sub SpinButton2_SpinUp()
    'Changes the CSN textbox on edit card page
    If Not TextBox20.Value = "06" Then
        TextBox20.Value = "0" & SpinButton2.Value
    Else:
        TextBox20.Value = "01"
    End If
    SpinButton2.Value = Int(TextBox20.Value)
End Sub



Private Sub SpinButton22_SpinDown()
    'Previous indenture level, assemblies page
    'Searchs listbox1 for a match on the main page from current position,
        'backward, and then loops around if not found
    Dim r As Variant
    Dim match As String
    Dim temp As String
    If ListBox4.ListCount < 1 Then Exit Sub
     match = Left(ListBox4.list(0), 1)
    If ListBox1.ListIndex < 0 Then Exit Sub
    'use this check box not to update the description box while we search for things.
    CheckBox22 = True
    
    Do While Not ListBox1.list(ListBox1.ListIndex) = ""
        If ListBox1.ListIndex = ListBox1.ListCount - 1 Then ListBox1.ListIndex = 0
        ListBox1.ListIndex = ListBox1.ListIndex - 1
    Loop
    
    For r = 0 To ListBox1.ListIndex
        If Not r = ListBox1.ListIndex Then
            temp = ListBox1.list(ListBox1.ListIndex - 1 - r)
        End If
        If Right(temp, 3) = "01A" Then
            If InStr(1, Mid(temp, 13, 1), match) > 0 Then
                ListBox1.TopIndex = ListBox1.ListIndex - 1 - r
                ListBox1.ListIndex = ListBox1.ListIndex - 1 - r
                GoTo 1
            End If
        End If
    Next r
    For r = 0 To ListBox1.ListCount - 1 - ListBox1.ListIndex
        If Right(ListBox1.list(ListBox1.ListCount - 1 - r), 3) = "01A" Then
            If InStr(1, Mid(ListBox1.list(ListBox1.ListCount - 1 - r), 13, 1), match) > 0 Then
                ListBox1.TopIndex = ListBox1.ListCount - 1 - r
                ListBox1.ListIndex = ListBox1.ListCount - 1 - r
                GoTo 1
            End If
        End If
    Next r
1   ComboBox1.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    CheckBox22 = False
    countselected
    r = ListBox1.ListIndex
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value
        r = r - 1
    Loop
    ListBox1.TopIndex = r + 1
    Label137.Caption = "Indenture " & Left(ListBox4.list(0), 1)
End Sub

Private Sub SpinButton22_SpinUp()
    'Next indenture level, assemblies page
    
    Dim r As Variant
    Dim match As String
    If ListBox4.ListCount < 1 Then Exit Sub
    match = Left(ListBox4.list(0), 1)
    'Searchs listbox1 for a match on the main page from current position,
        'foward, and then loops around if not found
    If ListBox1.ListIndex < 0 Then Exit Sub
    'use this check box not to update the description box while we search for things.
    CheckBox22 = True
    
    Do While Not ListBox1.list(ListBox1.ListIndex) = ""
        If ListBox1.ListIndex = ListBox1.ListCount - 1 Then ListBox1.ListIndex = 0
        ListBox1.ListIndex = ListBox1.ListIndex + 1
    Loop
    
    For r = ListBox1.ListIndex To ListBox1.ListCount - 1
        If Right(ListBox1.list(r), 3) = "01A" Then
            If InStr(1, Mid(ListBox1.list(r), 13, 1), match) > 0 Then
                GoTo 1
            End If
        End If
    Next r
    For r = 0 To ListBox1.ListIndex
        If Right(ListBox1.list(r), 3) = "01A" Then
            If InStr(1, Mid(ListBox1.list(r), 13, 1), match) > 0 Then
                GoTo 1
            End If
        End If
    Next r
1   ComboBox1.Value = Trim(Mid(ListBox1.list(r), 7, 5))
    ListBox1.TopIndex = r
    CheckBox22 = False
    countselected
    ListBox1.ListIndex = r
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value
        r = r - 1
    Loop
    ListBox1.TopIndex = r + 1
    Label137.Caption = "Indenture " & Left(ListBox4.list(0), 1)
End Sub


Private Sub SpinButton5_SpinDown()
    'Closes the top part of the form where the file and header are kept
    Frame24.Top = -33
    MultiPage1.Top = 0
    Report.Height = 375 - 36
    Frame25.Top = -12
    Frame25.Height = 26
    SpinButton5.Top = 12
    Report.Top = Report.Top + 36
    Frame26.Top = Frame25.Top
    Frame42.Top = Frame42.Top - 36
    Frame98.Top = Frame98.Top - 36
End Sub

Private Sub SpinButton5_SpinUp()
    'Opens the top part of the form where the file and header are kept
    SpinButton5.Top = -12
    Frame24.Top = 0
    Report.Top = Report.Top - 36
    MultiPage1.Top = 36
    Report.Height = 375
    Frame25.Top = 29
    Frame26.Top = Frame25.Top
    Frame25.Height = 16
    Frame42.Top = Frame42.Top + 36
    Frame98.Top = Frame98.Top + 36
End Sub
Private Sub SpinButton7_SpinDown()
    'Updates CFI textbox on edit card page
    If TextBox21.Value = "A" Then
        TextBox21.Value = "M"
        GoTo 2
    End If
1   SpinButton7.Value = Asc(TextBox21.Value) - 1
    TextBox21.Value = Chr(SpinButton7.Value)
    If TextBox21 = "I" Or TextBox21 = "L" Then GoTo 1
2   pickpage TextBox21.Value
End Sub

Private Sub SpinButton7_SpinUp()
    'Updates CFI textbox on edit card page
    If TextBox21.Value = "M" Then
        TextBox21.Value = "A"
        GoTo 2
    End If
1   SpinButton7.Value = Asc(TextBox21.Value) + 1
    TextBox21.Value = Chr(SpinButton7.Value)
    If TextBox21 = "I" Or TextBox21 = "L" Then GoTo 1
2   pickpage TextBox21.Value
End Sub



Private Sub SpinButton8_SpinDown()
    'Opens the description and search section on the main page
    Frame40.Visible = True
    Frame41.Visible = True
    ListBox1.Height = 182.05
    ListBox1.Top = 76.75
    SpinButton8.Top = SpinButton8.Top + 16
    Frame42.Top = 100 + Frame24.Top
    Label42.Visible = False
    Frame42.Left = 108
End Sub

Private Sub SpinButton8_SpinUp()
    'Closes the description and search section on the main page
    Frame40.Visible = False
    Frame41.Visible = False
    ListBox1.Top = 44.55
    ListBox1.Height = 214.2
    SpinButton8.Top = SpinButton8.Top - 16
    Frame42.Top = Frame24.Top + 86
    Label42.Visible = True
    Frame42.Left = 178
End Sub
Private Sub TextBox165_Change()
    'cage code textbox on the part number page
    TextBox164_Change
End Sub
Private Sub TextBox164_Change()
    'Reference number textbox on the part number page
    'This updates the buttons to update all or just this plisn on the part number page
    If TextBox164 = "" Then Exit Sub
    If ListBox3.ListIndex < 0 Then Exit Sub
    If Not TextBox165 & TextBox164.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 14, 37)) Then
        CommandButton33.Locked = False
        CommandButton33.BackColor = &HFF&
        CommandButton33.Caption = "Change all " & Mid(checkNHA(ListBox1.ListIndex), 6, 32) & "'s" 'ListBox3.list(ListBox3.ListIndex)
        CommandButton29.Caption = "Change all "
        TextBox185.Value = "Cage: " & Trim(Mid(ListBox1.list(ListBox1.ListIndex), 14, 5)) & " " & _
                                    "Part Number: " & Trim(Mid(ListBox1.list(ListBox1.ListIndex), 19, 32)) & _
                                    " to " & "Cage: " & TextBox165 & " Part Number " & TextBox164.Value
    Else
        CommandButton29.Caption = "Update all "
        TextBox185.Value = TextBox165 & " " & TextBox164.Value & "'s"
        
        CommandButton33.Locked = True
        CommandButton33.BackColor = &HC0C0C0
        CommandButton33.Caption = ""
    End If
End Sub
Private Sub TextBox166_Change()
    'part number page item name
    TextBox166.Value = UCase(TextBox166.Value)
End Sub

Private Sub TextBox175_Change()
    'part number page remarks
    TextBox175.Value = UCase(TextBox175.Value)
End Sub

Private Sub TextBox178_Change()
    'NHA textbox on the part number page
    If TextBox178.Value = "" Then Exit Sub
    If Frame96.Visible = True Then
        CommandButton32_Click
    End If
End Sub

Private Sub TextBox182_Change()
    'Same as PLISN textbox on the part number page
    If Not TextBox182 = "" Then
        If TextBox182.Value = ListBox3.list(0) Then
            TextBox182.BackColor = &H8000000B
        Else
            TextBox182.BackColor = &HC0C0FF
        End If
    End If
End Sub

Private Sub TextBox184_Change()
    'Textbox on the part number page in the NHA window when the ? is clicked
    CommandButton38.Caption = ListBox3.list(ListBox3.ListIndex)
End Sub

Private Sub TextBox186_Change()
    'raw edit textbox
    TextBox186.Value = UCase(TextBox186.Value)
End Sub



Private Sub TextBox200_Change()
    'gap for plisns
    If TextBox200 = "" Then
        TextBox199.Value = ""
        ComboBox10.Clear
        ComboBox10.Value = ""
        Exit Sub
    Else:
        ComboBox3_Change
    End If
    
End Sub

Private Sub TextBox45_Change()
    'Header box
    HEADER = TextBox45.Value
End Sub

Private Sub TextBox46_Change()
    'File name
    'update label above listbox1 that you can see when description is expanded
    Label43.Caption = "   " & TextBox46.Value
End Sub

Private Sub TextBox46_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'File name
    If LOADFILE = "" Then Exit Sub
    MsgBox ("File Name: " & LOADFILE), vbInformation
End Sub

Private Sub TextBox47_Change()
    'FSC/NSN count
    TextBox180.Value = TextBox47.Value
End Sub

Private Sub TextBox49_Change()
    'NHA on C card
    CommandButton42.Caption = "Goto " & TextBox49.Value
End Sub

Private Sub TextBox66_Change()
    'Description or search textbox on main page when you click down
    TextBox66.Value = UCase(TextBox66.Value)
End Sub
Private Sub TextBox80_Change()
    'NSN textbox on edit card page
    TextBox26.Value = Left(TextBox80.Value, 4)
    TextBox28.Value = Mid(TextBox80.Value, 5, 9)
    TextBox171.Value = TextBox80.Value
End Sub


Private Sub UserForm_Initialize()
    'What happens when the Reports userform Initializes
    MultiPage1.Value = 0
    MultiPage2.Value = 0
    SpinButton5_SpinDown
    SpinButton8_SpinUp
    With ComboBox11
        .AddItem ("Alpha-Numeric")
        .AddItem ("Alpha")
        .AddItem ("Numeric")
    End With
End Sub
'===========================================================================
'''Mulitpage1 Tab 1
'===========================================================================

Private Sub CommandButton26_Click()
    '''Raw Edit line
    On Error GoTo 1
    OptionButton6 = True
    If Not ListBox1.Selected(ListBox1.ListIndex) = True Then Exit Sub
    If ListBox1.list(ListBox1.ListIndex) = "" Then
        Frame98.Visible = False
        Exit Sub
    End If
    Dim plisn As String
    Dim cards As String
    Dim r As Integer
    
    r = ListBox1.ListIndex
    plisn = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    If plisn = "" Then
        r = r - 1
    Else
        Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
            If Not ListBox1.list(r) = "" Then
                r = r - 1
            End If
        Loop
    End If
    
    r = r + 1
    cards = ListBox1.list(r)
    TextBox187.Value = 1
    r = r + 1
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
        cards = cards & Chr(10) & ListBox1.list(r)
        r = r + 1
        TextBox187.Value = TextBox187.Value + 1
        If r > ListBox1.ListCount - 1 Then
            GoTo 4
        End If
    Loop
4   TextBox186.Value = cards
    Frame98.Visible = True
    Frame98.Caption = "----------------------- Raw Edit PLISN " & ComboBox1.Value & _
            "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    GoTo 2
1   'Error
2
End Sub
Function convertplisn(plisn As String)
    Dim l As Integer
    Dim number As String
    
    For l = 1 To Len(plisn)
        If Asc(Mid(plisn, l, 1)) < 65 Then
            number = number & Asc(Mid(plisn, l, 1))
        Else
            number = number & Asc(Mid(plisn, l, 1)) - 44
        End If
     Next
     convertplisn = number
End Function
Function LocatePLISN(searchf As String)
    
    Dim r As Variant
    Dim cp As String
    Dim ccp As String
    Dim found As Boolean
    On Error Resume Next
    'convert current position and new plisn for quicker search
    cp = convertplisn(searchf)
    ccp = convertplisn(Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5)))
    'search in some direction
    If cp = "" And ccp = "" Then
        Exit Function
    ElseIf cp = ccp Then
        'Back up until the PLISN doesn't match anymore
        Do While Trim(Mid(ListBox1.list(ListBox1.ListIndex - 1), 7, 5)) = ComboBox1.Value
            ListBox1.ListIndex = ListBox1.ListIndex - 1
        Loop
        populate ListBox1.ListIndex
        found = True
        Exit Function
    
    Else
        For r = 0 To ListBox1.ListCount - 1
            If Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value Then
                populate r
                found = True
                Exit Function
            End If
        Next
    End If
    If found = False Then
        ComboBox2.Value = ""
        ComboBox6.Value = ""
    End If
End Function
Function populate(r As Variant)
    'populates the boxes I need for each plisn
    ListBox1.ListIndex = r
    ListBox1.TopIndex = r
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value
        If Right(ListBox1.list(r), 3) = "01A" Then
            ComboBox2.Value = Trim(Mid(ListBox1.list(r), 19, 32))
            If CheckBox22 = False Then TextBox66.Value = Trim(Mid(ListBox1.list(r), 56, 19))
        ElseIf Right(ListBox1.list(r), 3) = "01B" Then
            ComboBox6.Value = Trim(Mid(ListBox1.list(r), 16, 13))
        ElseIf Right(ListBox1.list(r), 3) = "01C" Then
            CommandButton42.Caption = "Goto " & Trim(Mid(ListBox1.list(r), 13, 5))
        ElseIf Right(ListBox1.list(r), 3) = "01H" Then
            If CheckBox22 = False Then TextBox66.Value = TextBox66.Value '& Trim(Mid(ListBox1.list(r), 33, 44))
            Exit Function
        End If
        If r + 1 < ListBox1.ListCount - 1 Then
            r = r + 1
        Else
            Exit Function
        End If
    Loop
End Function
Function search(s As Integer, d As Integer, box As ComboBox)
    Dim r As Variant
    For r = 0 To ListBox1.ListCount - 1
        If box.Value = "" Then GoTo 1
        If Trim(Mid(ListBox1.list(r), s, d)) = box.Value Then
            If Not Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value Then
                ComboBox1.Value = Trim(Mid(ListBox1.list(r), 7, 5))
                search = Trim(Mid(ListBox1.list(r), 7, 5))
                Exit Function
            End If
            
        End If
    Next
1   search = "Not found"
    
End Function

Private Sub ComboBox1_Change()
    '''PLISN main
    On Error Resume Next
    Dim csn As String
    Dim spaces As String
    
    csn = TextBox20.Value
    If OptionButton5.Value = False Then ComboBox3 = ComboBox1
    If CheckBox19 = True Then Exit Sub
    CheckBox19 = True
    CheckBox19.Visible = True
    If ComboBox1.Value = "" Then
        ComboBox2 = ""
        ComboBox6 = ""
        TextBox66 = ""
        TextBox186 = ""
        CheckBox19 = False
        CheckBox19.Visible = False
        TextBox5.Value = ""
        TextBox47.Value = ""
        Exit Sub
    End If
    
    LocatePLISN ComboBox1.Value
    ListBox4.Clear
    'top index is incorrect, make it run up to not equal to plisn then down to A card
    spaces = Asc(Mid(ListBox1.list(ListBox1.ListIndex), 13, 1)) - 65
    spaces = Application.WorksheetFunction.Rept(" ", spaces * 2)
    'ListBox4.AddItem (Mid(ListBox1.list(ListBox1.TopIndex), 13, 1) & " " & Trim(Mid(ListBox1.list(ListBox1.TopIndex), 7, 5)) & spaces & ComboBox2.Value)
    ListBox4.AddItem ""
    ListBox4.Column(0, 0) = Mid(ListBox1.list(ListBox1.ListIndex), 13, 1) & " " & Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    ListBox4.Column(1, 0) = spaces & Trim(Mid(ListBox1.list(ListBox1.ListIndex), 19, 32))

    GetCount ComboBox2.Value, ComboBox6.Value
    SelectAll
    
    If TextBox20.Value = csn Then
        UpdateCardsCSN
    Else
        TextBox20.Value = csn
        
    End If
    CheckBox19 = False
    CheckBox19.Visible = False
    If Frame98.Visible = True Then
        TextBox186 = ""
        CommandButton26_Click
    End If
    ComboBox4.Value = UCase(ComboBox1.Value)
End Sub
Private Sub CommandButton25_Click()
    On Error GoTo 2
    Dim copyt As String
    Dim DataObj As New MSForms.DataObject
    Dim r As Variant
    
    For r = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(r) = True Then
            If copyt = "" Then
                copyt = ListBox1.list(r)
            Else
                copyt = copyt & _
                vbCr & _
                Chr(10) & ListBox1.list(r)
            End If
        End If
    Next
    
1   With DataObj
        .SetText copyt
        .PutInClipboard
    End With
    GoTo 3
2   'Error
3
End Sub



Private Sub ComboBox6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckBox19 = False
    CheckBox19.Visible = False
End Sub
Private Sub ComboBox6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckBox19 = False
    CheckBox19.Visible = False
End Sub

Private Sub ComboBox6_Change()
    'NSN
    Dim sf As String
    If CheckBox19 = True Then Exit Sub
    If CheckBox19 = False Then
        sf = search(16, 13, ComboBox6)
        If sf = "Not found" Then
            CheckBox19 = True
            ComboBox1 = ""
            ComboBox2 = ""
            TextBox66 = ""
            CheckBox19 = False
        End If
    End If
    Exit Sub
End Sub

Private Sub ComboBox2_Change()
    '''Part Number main
    Dim sf As String
    If OptionButton5.Value = False Then
        Updateaddpage
    End If
    If CheckBox19 = True Then Exit Sub
    If CheckBox19 = False Then
        sf = search(19, 32, ComboBox2)
        If sf = "Not found" Then
            CheckBox19 = True
            ComboBox1 = ""
            ComboBox6 = ""
            TextBox66 = ""
            CheckBox19 = False
        End If
        ComboBox8.Value = ComboBox2.Value
    End If
    
    'Exit Sub
    
End Sub
Function GetCount(number As String, nsn As String)
    'Counts the number of each part number or NSN
    Dim r As Variant
    Dim count As Integer
    Dim countb As Integer
    Dim start As Integer
    Dim qty As Integer
    Dim ASSEMBLY As String
    Dim downparts As Integer
    Dim spaces As String
    ListBox4.BackColor = &H80000005
    
    For r = 0 To ListBox1.ListCount - 1
        If Right(ListBox1.list(r), 3) = "01A" Then
            If Not ASSEMBLY = "done" And Not ASSEMBLY = "" Then
                If Asc(Mid(ListBox1.list(r), 13, 1)) > Asc(ASSEMBLY) Then
                    downparts = downparts + 1
                    spaces = Asc(Mid(ListBox1.list(r), 13, 1)) - 65
                    If spaces - (Asc(Mid(ListBox4.Column(0, ListBox4.ListCount - 1), 1, 1)) - 65) > 1 Then
                        ListBox4.BackColor = &HC0C0FF
                    End If
                    spaces = Application.WorksheetFunction.Rept(" ", spaces * 2)
                    ListBox4.AddItem ""
                    ListBox4.Column(0, ListBox4.ListCount - 1) = Mid(ListBox1.list(r), 13, 1) & " " & Trim(Mid(ListBox1.list(r), 7, 5))
                    ListBox4.Column(1, ListBox4.ListCount - 1) = spaces & Trim(Mid(ListBox1.list(r), 19, 32))
                Else
                    ASSEMBLY = "done"
                End If
                
            End If
            If Trim(Mid(ListBox1.list(r), 19, 32)) = number Then
                If Not number = "" Then
                    count = count + 1
                    If ASSEMBLY = "" Then
                        If Trim(Mid(ListBox4.list(0), 2, 5)) = Trim(Mid(ListBox1.list(r), 7, 5)) Then
                            ASSEMBLY = Mid(ListBox1.list(r), 13, 1)
                        End If
                    End If
                End If
            End If
            'count downparts
            TextBox201.Value = downparts
            
        End If
        If Right(ListBox1.list(r), 3) = "01B" Then
            If Trim(Mid(ListBox1.list(r), 16, 13)) = nsn Then
                If Not nsn = "" Then
                    countb = countb + 1
                End If
            End If
        End If
        
    Next
    TextBox5.Value = count
    TextBox47.Value = countb
End Function
Private Sub CommandButton1_Click()
    '''Load Report
    'On Error GoTo 1
    CheckBox19 = True
    TextBox45.Value = ""
    TextBox46.Value = ""
    TextBox3 = ""
    TextBox47 = ""
    TextBox5 = ""
    TextBox164 = ""
    With ComboBox7
        .Clear
        .AddItem "N"
        .AddItem "C"
        .AddItem "B"
        .AddItem "*"
        .AddItem "F"
        .AddItem "A"
        .AddItem "E"
    End With
    Dim s As Variant
    Dim var As Variant
    Dim boxes As Variant
    boxes = Array(ListBox1, ComboBox1, ComboBox2, ComboBox3, ComboBox4, ComboBox6, ListBox3, _
            ListBox2)
    
    For Each var In boxes
        var.Clear
        var.Value = ""
    Next
    TextBox48 = ""
    TextBox66 = ""
    ListBox1.MultiSelect = fmMultiSelectSingle
    Application.ScreenUpdating = False
3   Open_036
    ListBox1.MultiSelect = fmMultiSelectExtended
    Application.ScreenUpdating = True
    GoTo 2
1   MsgBox ("Invalid file type!"), vbExclamation
    GoTo 3
2
    TextBox45.Value = HEADER
    TextBox46.Value = FILENAME
    
    s = SortParts(ComboBox2.list, ComboBox2)
    s = SortParts(ComboBox6.list, ComboBox6)
    CheckBox19 = False
    ComboBox1 = ""
End Sub

Private Sub CommandButton11_Click()
    '''Select cards

    ListBox1.MultiSelect = fmMultiSelectSingle
    ListBox1.MultiSelect = fmMultiSelectExtended
    Dim r As Variant
    Dim s As Variant
    Dim Arr As Variant
    Dim narr As Variant
    Dim t As Variant
    Dim csn As Variant
    Dim Arr2 As Variant
    Dim count As Long
    
    narr = Array("A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "M")
    Arr = Array("N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N")
    '''Update Arr to Y/N
    If CheckBox1.Value = True Then Arr(0) = "Y"
    If CheckBox2.Value = True Then Arr(1) = "Y"
    If CheckBox3.Value = True Then Arr(2) = "Y"
    If CheckBox4.Value = True Then Arr(3) = "Y"
    If CheckBox5.Value = True Then Arr(4) = "Y"
    If CheckBox6.Value = True Then Arr(5) = "Y"
    If CheckBox7.Value = True Then Arr(6) = "Y"
    If CheckBox8.Value = True Then Arr(7) = "Y"
    If CheckBox9.Value = True Then Arr(8) = "Y"
    If CheckBox10.Value = True Then Arr(9) = "Y"
    If CheckBox11.Value = True Then Arr(10) = "Y"
    

    csn = Array("01", "02", "03", "04", "05", "06")
    Arr2 = Array("N", "N", "N", "N", "N", "N")
    
    If CheckBox13.Value = True Then Arr2(0) = "Y"
    If CheckBox14.Value = True Then Arr2(1) = "Y"
    If CheckBox15.Value = True Then Arr2(2) = "Y"
    If CheckBox16.Value = True Then Arr2(3) = "Y"
    If CheckBox17.Value = True Then Arr2(4) = "Y"
    If CheckBox18.Value = True Then Arr2(5) = "Y"
    
    For r = 0 To ListBox1.ListCount - 1
        For s = 0 To 10
            If Arr(s) = "Y" Then
                If Right(ListBox1.list(r), 1) = narr(s) Then
                    ListBox1.Selected(r) = True
                    count = count + 1
                End If
            End If
        Next
        For s = 0 To 5
            If Arr2(s) = "Y" Then
                If Mid(ListBox1.list(r), 78, 2) = csn(s) Then
                    ListBox1.Selected(r) = True
                    count = count + 1
                End If
            End If
        Next
        'if right(listbox.list(r),1)
    
    Next
    TextBox48.Value = count

End Sub

Private Sub CommandButton12_Click()
    ''' + select all
    CheckBox1.Value = True
    CheckBox2.Value = True
    CheckBox3.Value = True
    CheckBox4.Value = True
    CheckBox5.Value = True
    CheckBox6.Value = True
    CheckBox7.Value = True
    CheckBox8.Value = True
    CheckBox9.Value = True
    CheckBox10.Value = True
    CheckBox11.Value = True
    CheckBox13.Value = True
    CheckBox14.Value = True
    CheckBox15.Value = True
    CheckBox16.Value = True
    CheckBox17.Value = True
    CheckBox18.Value = True

End Sub

Private Sub CommandButton13_Click()
    ''' - select all
CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False
CheckBox4.Value = False
CheckBox5.Value = False
CheckBox6.Value = False
CheckBox7.Value = False
CheckBox8.Value = False
CheckBox9.Value = False
CheckBox10.Value = False
CheckBox11.Value = False
CheckBox13.Value = False
CheckBox14.Value = False
CheckBox15.Value = False
CheckBox16.Value = False
CheckBox17.Value = False
CheckBox18.Value = False

End Sub

Private Sub CommandButton14_Click()
    '''Edit card
    Dim sel As String
    Dim answer As String
    
    If ListBox1.ListIndex = -1 Then
        If ComboBox1.Value = "" Then
            ListBox1.MultiSelect = fmMultiSelectExtended
            Exit Sub
        Else
             CommandButton18_Click
        End If
    End If
    
    'If Mid(ListBox1.list(ListBox1.ListIndex), 78, 2) = "01" Then
    
1       Report.MultiPage1.Value = 1
        sel = ListBox1.list(ListBox1.ListIndex)
        pickpage sel
        
        ListBox1.MultiSelect = fmMultiSelectExtended
        TextBox20.Value = Mid(ListBox1.list(ListBox1.ListIndex), 78, 2)
    'Else:
        'answer = inputboxx("Please edit card", "Edit Card", ListBox1.list(ListBox1.ListIndex))
        'If Not answer = "Cancel" Then ListBox1.list(ListBox1.ListIndex) = answer
    'End If
    
2
   
End Sub
Function pickpage(sel As String)

    If Right(sel, 1) = "A" Then
        Report.MultiPage2.Value = 0
    ElseIf Right(sel, 1) = "B" Then
        Report.MultiPage2.Value = 1
    ElseIf Right(sel, 1) = "C" Then
        Report.MultiPage2.Value = 2
    ElseIf Right(sel, 1) = "D" Then
        Report.MultiPage2.Value = 3
    ElseIf Right(sel, 1) = "E" Then
        Report.MultiPage2.Value = 4
    ElseIf Right(sel, 1) = "F" Then
        Report.MultiPage2.Value = 5
    ElseIf Right(sel, 1) = "G" Then
        Report.MultiPage2.Value = 6
    ElseIf Right(sel, 1) = "H" Then
        Report.MultiPage2.Value = 7
    ElseIf Right(sel, 1) = "J" Then
        Report.MultiPage2.Value = 8
    ElseIf Right(sel, 1) = "K" Then
        Report.MultiPage2.Value = 9
    ElseIf Right(sel, 1) = "M" Then
        Report.MultiPage2.Value = 10
    End If
    
End Function
Private Sub CommandButton17_Click()
    '''Reset to loaded file
On Error Resume Next
Dim s As Variant
Dim var As Variant
Dim l As String
Dim items As Variant
Dim o As Variant
Dim cond As Variant
Dim scond As Variant
Dim att As Variant
Dim count As Integer

OptionButton4 = False
OptionButton3 = False
OptionButton2 = False

ListBox1.Clear

TextBox47 = ""
TextBox48 = ""
TextBox3 = ""
TextBox5 = ""
ComboBox1.Clear
ComboBox2.Clear
ComboBox6.Clear
    
    For Each items In PLISNs
        GoTo 1
       

1           With Report.ListBox1
            l = items.cards()
            For Each o In Split(l, vbCr)
                .AddItem o
            Next
        End With
        With Report.ComboBox1
            .AddItem items.plisn
        End With
        With Report.ComboBox2
            For Each var In Array("01", "02", "03", "04", "05", "06")
                .AddItem Trim(Mid(items.Acard(var), 19, 32))
            Next
        End With
        With Report.ComboBox6
            For Each var In Array("01", "02", "03", "04", "05", "06")
                If Not Trim(Mid(items.Bard(var), 16, 13)) = "" Then
                    .AddItem Trim(Mid(items.Bcard(var), 16, 13))
                End If
            Next
        End With
        
        count = count + 1
    Next

    Report.ListBox1.AddItem ""
    Update_PN_List
    
    RemoveDuplicates ComboBox2, "", 0
    RemoveDuplicates ComboBox6, "", 0
    RemoveDuplicates ComboBox1, "", 0
    
    s = SortParts(ComboBox2.list, ComboBox2)
    s = SortParts(ComboBox6.list, ComboBox6)

    TextBox3.Value = ComboBox1.ListCount
        
End Sub
Private Sub CommandButton18_Click()
    '''Edit/Add PLISN
    Report.MultiPage1.Value = 3
End Sub

Private Sub CommandButton9_Click()
    '''Save this file
    Dim fdlr As String
    Dim file As String
    Dim r As Variant
    update_date

    file = LOADFILE
    If LOADFILE = "" Then
        MsgBox ("I did not know the file to save it as or it is readonly"), vbCritical
        CommandButton10_Click
        GoTo 1
    End If
    If MsgBox("Are you sure you want to overwrite the original file?", vbYesNo) = vbYes Then
    
        Open file For Output Access Write As #1
        Print #1, HEADER
        For r = 0 To ListBox1.ListCount - 1
            If ListBox1.list(r) = "" Then
            Else:
                Print #1, ListBox1.list(r)
            End If
        Next
        Close #1
        
    
    End If
    
1
End Sub
Function update_date()
    If CheckBox21.Value = True Then
        If Len(Trim(HEADER)) = 72 Then
            TextBox45.Value = Left(HEADER, 66) & format(Date, "YYMMDD")
        End If
    End If
End Function
Private Sub CommandButton10_Click()
    '''Save as file
    Dim fdlr As String
    Dim file As String
    Dim r As Variant
    
    update_date
    
    If TextBox4.Value = "" Then
        MsgBox ("Please enter a name to save the file as."), vbCritical
        Exit Sub
    Else:
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = ActiveWorkbook.Path & "\"
            .Show
            If .SelectedItems.count = 0 Then GoTo 1
            fdlr = .SelectedItems(1) & "\"
        End With
    
        file = fdlr & TextBox4.Value & ".036"
        Open file For Output Access Write As #1
        Print #1, HEADER
        For r = 0 To ListBox1.ListCount - 1
            If CheckBox20.Value = True Then
                If ListBox1.Selected(r) = True Then
                    Print #1, ListBox1.list(r)
                End If
            Else
                If ListBox1.list(r) = "" Then
                Else:
                    Print #1, ListBox1.list(r)
                End If
            End If
            
        Next
        Close #1
    End If
    MsgBox ("File " & file & " Saved."), vbInformation
    
1
End Sub

Private Sub CommandButton2_Click()
    '''Delete Cards
On Error Resume Next
Dim count As Integer
Dim qty As Integer
Dim qtyb As Integer
    Dim r As Variant
    Dim nr As Long
    Dim nnr As Integer
    Dim lp As String
    
    
    nr = ListBox1.ListCount - 1
    For r = 0 To ListBox1.ListCount - 1
         nnr = nr - r
        
        If ListBox1.Selected(nnr) = True Then
            ListBox1.RemoveItem (nnr)
        End If
        If Trim(Mid(ListBox1.list(nnr), 7, 5)) = lp Then
        Else:
            
        End If
        lp = Trim(Mid(ListBox1.list(nnr), 7, 5))
    Next r
    qty = TextBox5.Value
    qtyb = TextBox47.Value
    RemoveDuplicates ComboBox2, ComboBox2.Value, qty
    RemoveDuplicates ComboBox6, ComboBox6.Value, qtyb
    RemoveDuplicates ComboBox1, ComboBox1.Value, qty
    
    ComboBox6 = ""
    ComboBox2 = ""
    ComboBox1 = ""

    Update_PN_List
    
    TextBox3.Value = ComboBox1.ListCount

    
End Sub

Function inputboxx(messageJ As String, TitleJ As String, DefaultJ As String)

    Dim answer As String
    ' Display message, title, and default value.
    answer = inputbox(messageJ, TitleJ, DefaultJ)
    
    If answer = "" Then
        'Do something if the text equals something.
    End If
    If StrPtr(answer) = 0 Then
        'Do something if cancel is hit.
        inputboxx = "Cancel"
        Exit Function
    Else
        'Do this after Okay is hit
        
    End If
    inputboxx = UCase(answer)
End Function

Function Addline(Text As String, above As Boolean)
    On Error GoTo 1
    Dim r As Variant
    Dim last As Integer
    Dim nextcard As String
    Dim newtext As String
    Dim cardv As String
    
    'if the list is empty add a line
    If ListBox1.ListIndex < 0 Then
        ListBox1.AddItem ("")
    End If
    
    If Text = "" Then
        cardv = inputboxx("Card information after the PLISN?" & _
            Chr(10) & _
            Chr(10) & "You can enter anything after the PLISN" & _
            Chr(10) & "but at least enter the CSN and Card Letter", "Card?", "A")
        
        If cardv = "Cancel" Then Exit Function
        If Not Len(Trim(Left(ListBox1.list(ListBox1.ListIndex), 11))) = 10 Then
                newtext = Application.WorksheetFunction.Rept(" ", 80 - Len(cardv)) & cardv
        Else:
            newtext = Trim(Left(ListBox1.list(ListBox1.ListIndex), 11)) & _
                    Application.WorksheetFunction.Rept(" ", 70 - Len(cardv)) & cardv
        End If
    Else
        newtext = Text
    End If
    
    ListBox1.AddItem ("")
    last = ListBox1.ListCount - 1
    For r = 0 To ListBox1.ListCount - ListBox1.ListIndex - 1
        If last - r = 0 Then
            
        Else
            ListBox1.list(last - r) = ListBox1.list(last - r - 1)
        End If
    Next
    
    If above = True Then
        ListBox1.list(ListBox1.ListIndex) = newtext
    Else:
        ListBox1.list(ListBox1.ListIndex) = ListBox1.list(ListBox1.ListIndex + 1)
        ListBox1.list(ListBox1.ListIndex + 1) = newtext
    End If
    GoTo 2
1   'Error
2
End Function
Private Sub CommandButton4_Click()
    On Error GoTo 1
    ListBox1.RemoveItem (ListBox1.ListIndex)
    If Right(ListBox1.list(ListBox1.ListIndex), 3) = "01B" Then ComboBox6 = ""
    If Right(ListBox1.list(ListBox1.ListIndex), 3) = "01A" Then ComboBox2 = ""
    
    
    Update_PN_List
1   'Error
2
End Sub
Function GetTextArray()
    '''Textbox array for updating and adding
    Dim Text As Variant
    
    If MultiPage2.Value = 0 Then
        TextBox21.Value = "A"
        Text = Array(TextBox9, TextBox10, TextBox11, TextBox12, TextBox13, _
            TextBox14, TextBox15, TextBox16, TextBox17, TextBox18, _
            TextBox19)
    ElseIf MultiPage2.Value = 1 Then
        TextBox21.Value = "B"
        Text = Array(TextBox25, TextBox80, TextBox42, TextBox39, TextBox40, TextBox41, _
            TextBox38, TextBox29, TextBox30, TextBox31, TextBox32, TextBox33, TextBox34, TextBox35, _
            TextBox43, TextBox44)
    ElseIf MultiPage2.Value = 2 Then
        TextBox21.Value = "C"
        Text = Array(TextBox49, ComboBox7, TextBox50, TextBox51, TextBox55, TextBox52, TextBox53, _
        TextBox54, TextBox65, TextBox56, TextBox57, TextBox58, TextBox59, TextBox60)
    ElseIf MultiPage2.Value = 3 Then
        TextBox21.Value = "D"
        Text = Array(TextBox67, TextBox68, TextBox69, TextBox70, TextBox71, TextBox72, TextBox81, _
        TextBox79, TextBox78, TextBox73, TextBox74, TextBox75, TextBox76, TextBox77)
    ElseIf MultiPage2.Value = 4 Then
        TextBox21.Value = "E"
        Text = Array(TextBox82, TextBox83, TextBox96, TextBox97, TextBox98, TextBox99, TextBox100, _
        TextBox101, TextBox102, TextBox108, TextBox109, TextBox105, TextBox106, TextBox92, _
        TextBox93, TextBox91, TextBox94, TextBox95, TextBox90, TextBox89, TextBox84, TextBox85, _
        TextBox86, TextBox111, TextBox110)
    ElseIf MultiPage2.Value = 5 Then
        TextBox21.Value = "F"
        Text = Array(TextBox112, TextBox113, TextBox114, TextBox117, TextBox123, TextBox118, _
        TextBox119, TextBox120, TextBox121, TextBox122)
    ElseIf MultiPage2.Value = 6 Then
        TextBox21.Value = "G"
        Text = Array(TextBox124, TextBox125, TextBox126, TextBox133)
    ElseIf MultiPage2.Value = 7 Then
        TextBox21.Value = "H"
        Text = Array(TextBox134, TextBox135, TextBox137)
    ElseIf MultiPage2.Value = 8 Then
        TextBox21.Value = "J"
        Text = Array(TextBox138, TextBox139, TextBox140, TextBox141, TextBox147, TextBox142, _
        TextBox143, TextBox144, TextBox145, TextBox146, TextBox148, TextBox149, TextBox150, _
        TextBox151, TextBox152, TextBox153, TextBox154, TextBox155)
    ElseIf MultiPage2.Value = 9 Then
        TextBox21.Value = "K"
        Text = Array(TextBox156, TextBox157, TextBox159, TextBox158)
    ElseIf MultiPage2.Value = 10 Then
        TextBox21.Value = "M"
        Text = Array(TextBox163)
    End If
    
GetTextArray = Text
End Function

Private Sub UpdateCardsCSN()
    '''Update Cards
    Dim Text As Variant
    Dim start As Variant
    Dim r As Variant
    Dim c As Variant
    Dim found As Boolean
    Dim TextLine As String
    Dim plisn As String
    On Error Resume Next
    
    plisn = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    Text = GetTextArray
        
    For r = ListBox1.ListIndex To ListBox1.ListCount - 1
        Do While Trim(Mid(ListBox1.list(r - 1), 7, 5)) = plisn
            If r < 0 Then GoTo 1
            r = r - 1
        Loop
        'If ComboBox4 = "" Then
            'Exit Sub
        'End If
        Do While ComboBox1.Value = Trim(Mid(ListBox1.list(r), 7, 5))
            TextLine = ListBox1.list(r)
            If Not IsEmpty(Text) Then
                If Right(TextLine, 3) = TextBox20.Value & TextBox21.Value Then
                    For c = 0 To UBound(Text)
                        Text(c).Value = Trim(Mid(TextLine, Mid(Text(c).Tag, 2, 2), _
                            Text(c).MaxLength))
                    Next
                    TextBox8.Value = Mid(ListBox1.list(r), 12, 1)
                    found = True
                End If
                If found = True Then
                    Exit Sub
                End If
                If r + 1 < ListBox1.ListCount Then
                    r = r + 1
                Else
                    GoTo 1
                End If
            Else
                GoTo endof:
            End If
        Loop
1       If found = False Then
            For c = 0 To UBound(Text)
                Text(c).Value = ""
            Next
            TextBox8.Value = " "
            Exit Sub
        End If
        
    Next
endof:
    
End Sub
Private Sub SelectAll()
    '''Find/Select All
    On Error Resume Next
    Dim r As Variant
    Dim position As Integer
    Dim copyt As String
    Dim DataObj As New MSForms.DataObject
    
    position = ListBox1.ListIndex
    
    ListBox1.MultiSelect = fmMultiSelectSingle
    ListBox1.MultiSelect = fmMultiSelectExtended
    ListBox1.ListIndex = position
    
    For r = position To ListBox1.ListCount - 1
        If Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value Then
            ListBox1.Selected(r) = True
            If copyt = "" Then
                copyt = ListBox1.list(r)
            Else
                copyt = copyt & _
                vbCr & _
                Chr(10) & ListBox1.list(r)
            End If
            If position = 0 Then position = r
        Else:
            If Not position = 0 Then
                ListBox1.TopIndex = position
                'ListBox1.ListIndex = position
                GoTo 1
            End If
        End If
    Next r
    ListBox1.TopIndex = position
    'ListBox1.ListIndex = ListBox1.TopIndex
1   If MultiPage1.Value = 0 Then
        With DataObj
            .SetText copyt
            .PutInClipboard
        End With
    End If
End Sub

Private Sub CommandButton8_Click()
    '''Filter
    Dim fplisn As String
    Dim cond As String
    Dim scond As String
    Dim att As String
    Dim r As Integer
    Dim last As String
    Dim curplisn() As String
    Dim TextLine As String
    Dim c As Integer
    Dim line As Variant
    Dim keep As Boolean
    Dim start As Integer
    Dim remove As Variant
    Dim count As Integer
    On Error Resume Next
    
    For r = 0 To ListBox1.ListCount - 1
        line = ListBox1.list(r)
        If Not line = "" Then
            If OptionButton4.Value = True Then
            'NSN
                If Right(line, 3) = "01B" Then
                    If Len(Trim(Mid(line, 16, 13))) > 4 And Not _
                        Right(Trim(Mid(line, 16, 13)), 1) = "-" Then
                    Else
                        deleteplisn Trim(Mid(line, 7, 5)), r
                    End If
                
                End If
            ElseIf OptionButton3.Value = True Then
            'CAGE
                If Right(line, 3) = "01A" Then
                    If TextBox1.Value = Mid(line, 14, Len(TextBox1.Value)) Then
                    Else
                        deleteplisn Trim(Mid(line, 7, 5)), r
                    End If
                End If
            
            ElseIf OptionButton2.Value = True Then
            'SMR
                If Right(line, 3) = "01B" Then
                    If TextBox2.Value = Mid(line, 65, Len(TextBox2.Value)) Then
                    Else
                        deleteplisn Trim(Mid(line, 7, 5)), r
                    End If
                End If
            
            End If
        End If
        If r > ListBox1.ListCount - 1 Then GoTo 4
1   Next
    
4   Update_PN_List
    TextBox3.Value = ComboBox1.ListCount
End Sub
Function deleteplisn(plisn As String, r As Integer)
    
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
        r = r - 1
    Loop
    ListBox1.RemoveItem (r)
    Do While Trim(Mid(ListBox1.list(r), 7, 5)) = plisn
        ListBox1.RemoveItem (r)
        If r > ListBox1.ListCount - 1 Then Exit Function
    Loop
End Function
Sub temp()
    If OptionButton2.Value = True Then
        cond = TextBox2.Value
        scond = cond
        att = Left(items.SMR, Len(cond))
    End If
    If OptionButton3.Value = True Then
        cond = TextBox1.Value
        scond = cond
        att = Left(Trim(items.cage), Len(cond))
    End If
    If OptionButton4.Value = True Then
        cond = " 13"
        scond = "---------"
        att = Str(Len(Trim(items.nsn)))
    End If
    
    If att = cond Or att = scond Then
End Sub
Private Sub Image1_Click()
    MsgBox ("MakeCents" & _
    Chr(10) & Report.Caption & _
    Chr(10) & _
    Chr(10) & "By Dave Gillespie" & _
    Chr(10) & "Date: 10/23/2013"), vbInformation
        
End Sub

Sub Open_036()
'On Error Resume Next
'On Error GoTo 1
Dim PCCNs As String
Dim p As New plisn
Dim same As String
Dim csn As String
Dim var As Variant
Dim items As Variant
Dim l As String
Dim o As Variant
Dim indexer As Integer
Dim count As Integer
3
For Each items In PLISNs
    PLISNs.remove (items.plisn)
Next

'Dim PLISNs As New Collection

Const ForReading = 1, ForWriting = 2, ForAppending = 8
' The following line contains constants for the OpenTextFile
' format argument, which is not used in the code below.
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fso, MyFile, FILENAME, TextLine

Set fso = CreateObject("Scripting.FileSystemObject")

    FILENAME = File_Picker
    If FILENAME = "" Then Exit Sub
    LOADFILE = FILENAME
' Open the file for input.
Set MyFile = fso.OpenTextFile(FILENAME, ForReading)

' Read from the file and display the results.
Dim Card As String
    Do While MyFile.AtEndOfStream <> True
        
        TextLine = MyFile.ReadLine
        same = Trim(Mid(TextLine, 7, 5))
        Card = Right(TextLine, 1)
        csn = Mid(TextLine, 78, 2)
        If Card = " " Or Len(TextLine) < 80 Then
            If Not Card = "" Then
                HEADER = UCase(TextLine)
            End If
        Else
            
            If Not p.plisn = same Then
                ListBox1.AddItem ("")
                count = count + 1
                PLISNs.Add Item:=p, Key:=p.plisn
                Set p = New plisn
                p.plisn = Trim(Mid(TextLine, 7, 5))
                p.PCCN = Trim(Left(TextLine, 6))
                If PCCNs = "" Then
                    PCCNs = p.PCCN
                    TextBox6.Value = p.PCCN
                End If
                p.partNumber = Trim(Mid(TextLine, 19, 32))
                p.cage = Trim(Mid(TextLine, 14, 5))
                p.Index = indexer
                indexer = indexer + 1
                
            End If
            ListBox1.AddItem (TextLine)
        End If
        
        Select Case Card
            Case "A"
                p.a.Add Item:=TextLine, Key:=(csn)
            Case "B"
                p.nsn = Mid(TextLine, 16, 13)
                p.SMR = Trim(Mid(TextLine, 64, 6))
                p.B.Add Item:=TextLine, Key:=(csn)
            Case "C"
                p.c.Add Item:=TextLine, Key:=(csn)
            Case "D"
                p.d.Add Item:=TextLine, Key:=(csn)
            Case "E"
                p.E.Add Item:=TextLine, Key:=(csn)
            Case "F"
                p.f.Add Item:=TextLine, Key:=(csn)
            Case "G"
                p.G.Add Item:=TextLine, Key:=(csn)
            Case "H"
                p.H.Add Item:=TextLine, Key:=(csn)
            Case "J"
                p.J.Add Item:=TextLine, Key:=(csn)
            Case "K"
                p.K.Add Item:=TextLine, Key:=(csn)
            Case "M"
                p.m.Add Item:=TextLine, Key:=(csn)
        End Select
        
    Loop
        'CommandButton17_Click
        Update_PN_List
        
    GoTo 2
1   MsgBox ("Error while loading file" & _
            Chr(10) & "Please select another file"), vbExclamation
    GoTo 3
2
    MyFile.Close
    TextBox3.Value = count
    'TextBox6.Value = PLISNs.plisn.PCCN
    ComboBox4.list = ComboBox1.list
    ComboBox3.list = ComboBox1.list
    'ComboBox9.list = ComboBox1.list
    
End Sub

Function File_Picker()
    Dim FNAME As String
    Dim digit As String
    Dim cur As Integer
    With Application.FileDialog(msoFileDialogFilePicker)
        .FILTERs.Clear
        .FILTERs.Add "Text", "*.*", 1
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Show
        If .SelectedItems.count = 0 Then GoTo 1
        File_Picker = .SelectedItems(1)
        
        Do While Not digit = "\"
            FNAME = Mid(File_Picker, Len(File_Picker) - cur, Len(File_Picker))
            digit = Left(Mid(File_Picker, Len(File_Picker) - cur, Len(File_Picker)), 1)
            cur = cur + 1
        Loop
        FILENAME = Mid(FNAME, 2, Len(FNAME))
        
1   End With
End Function

Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '''List box on main
    On Error Resume Next
    ComboBox1.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    Exit Sub
    If Not ListBox1.list(ListBox1.ListIndex) = "" Then
        ComboBox2.Value = PLISNs(Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))).partNumber
    End If
    
    countselected
End Sub
Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '''listbox on main
    On Error GoTo 1
    
    ComboBox1.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
    CheckBox19 = False
    Dim position As Integer
    position = ListBox1.ListIndex
    If Not ListBox1.list(ListBox1.ListIndex) = "" Then
        ComboBox1.Value = Trim(Mid(ListBox1.list(ListBox1.ListIndex), 7, 5))
        ListBox1.ListIndex = position
    End If
    countselected
    GoTo 2
1   'Error
2
End Sub
Function countselected()
    Dim r As Variant
    Dim count As Integer
    For r = ListBox1.TopIndex To ListBox1.ListCount - 1
        If ListBox1.Selected(r) = True Then
            count = count + 1
        End If
    Next
    TextBox48.Value = count
End Function
Private Sub OptionButton2_Change()
    '''SMR
    If OptionButton2.Value = False Then TextBox2.Value = ""
End Sub

Private Sub OptionButton3_Change()
    '''CAGE
    If OptionButton3.Value = False Then TextBox1.Value = ""
End Sub


Private Sub OptionButton4_Click()
'NSN
End Sub

Private Sub OptionButton5_Click()
'Add PLISN
End Sub

Private Sub TextBox1_Change()
'''CAGE code
    TextBox1.Value = UCase(TextBox1.Value)
    If TextBox1.Value = "" Then
    Else:
        OptionButton3.Value = True
    End If
End Sub


Private Sub TextBox17_Change()
    TextBox17.Value = UCase(TextBox17.Value)
    TextBox166.Value = TextBox17.Value
End Sub

Private Sub TextBox2_Change()
'''SMR code
    TextBox2.Value = UCase(TextBox2.Value)
    If TextBox2.Value = "" Then
    Else:
        OptionButton2.Value = True
    End If
End Sub

Private Sub TextBox3_Change()
    '''Total in report
End Sub

Private Sub TextBox4_Change()
    'File save as name
End Sub

Private Sub TextBox5_Change()
    '''count PLISNs with this part number
    TextBox203.Value = TextBox5.Value
End Sub
Private Sub TextBox6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '''PCCN
    TextBox23.Value = "PCCN - "
    TextBox23.Value = TextBox23.Value & Chr(10) & Chr(10) & "1. "

End Sub

Private Sub TextBox8_Change()
    TextBox8.Value = UCase(TextBox8.Value)
End Sub

'===========================================================================
'''Mulitpage1 Tab 2
'===========================================================================
Private Sub Frame8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox23.Value = "Key:" & Chr(10)
    TextBox23.Value = TextBox23.Value & "Above each textbox is the name for the entry field" & Chr(10)
    TextBox23.Value = TextBox23.Value & "Hovering or entering the textbox shows more information, about this entry, in this window" & Chr(10)
End Sub


Private Sub TextBox31_Change()
    '''SMR
    TextBox31.Value = UCase(TextBox31.Value)
End Sub


Private Sub SpinButton3_Change()
    TextBox36.Value = "0" & Trim(Str(SpinButton3.Value))
End Sub

Private Sub CommandButton15_Click()
    On Error GoTo 1
    Update_Card
    Update_PN_List
1
End Sub

Private Sub CommandButton16_Click()
    '''Reset Card
    UpdateCardsCSN
End Sub
Private Sub CommandButton20_Click()
    CommandButtonAdd
    Update_PN_List
End Sub
Private Sub CommandButtonAdd()
    '''Add Card
    On Error GoTo 2
    Dim TextLine As String
    Dim r As Variant
    Dim txbx As Variant
    Dim cardid As String
    Dim found As Boolean
    Dim plisnfix As String
    Dim csnwas As String
    csnwas = TextBox20.Value
    
    cardid = TextBox20.Value & TextBox21.Value
    
    TextLine = TextLine & fixacard
    
    For r = ListBox1.TopIndex To ListBox1.ListCount - 1
        If Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox4.Value Then
            'Begin PLISN and check for card
            If Right(ListBox1.list(r), 1) = TextBox21.Value Then
                If Int(TextBox20.Value) > Int(Mid(ListBox1.list(r), 78, 2)) Then
                    If Right(ListBox1.list(r + 1), 1) = TextBox21.Value Then
                    Else
                        ListBox1.ListIndex = r
                        Addline TextLine, False
                        GoTo 1
                    End If
                Else
                        ListBox1.ListIndex = r
                        Addline TextLine, True
                        GoTo 1
                End If
            ElseIf Int(TextBox20.Value) <= Int(Mid(ListBox1.list(r), 78, 2)) Then
                
                If Asc(TextBox21.Value) < Asc(Right(ListBox1.list(r), 1)) Then
                    
                    ListBox1.ListIndex = r
                    Addline TextLine, True
                    GoTo 1
                ElseIf r + 1 > (ListBox1.ListCount - 1) Or Right(ListBox1.list(r + 1), 1) = "" Then
                    ListBox1.ListIndex = r
                    Addline TextLine, False
                    GoTo 1
                End If
            End If
            
            
        End If
        
        
    Next
1   TextBox23.Value = "Your card has be added"
    GoTo 3
2   'Error
3
End Sub


Sub Update_PN_List()
    Dim r As Variant
    Dim partNumber As String
    Dim s As Variant
    Dim plisn As String
    Dim bplisn As String
    Dim last As String
    partNumber = TextBox11.Value
    Dim part As Variant
    Dim cagePart As String
    
    ComboBox2.Clear
    ComboBox6.Clear
    
    ListBox2.Clear
    For part = 1 To (allparts.count)
        allparts.remove (1)
    Next part
    ComboBox1.Clear
    bplisn = ComboBox1.Value
    For r = 0 To ListBox1.ListCount - 1
        
        If Right(ListBox1.list(r), 1) = "A" Then
            If Not (Trim(Mid(ListBox1.list(r), 19, 32))) = "" Then
                ComboBox2.AddItem (Trim(Mid(ListBox1.list(r), 19, 32)))
                cagePart = (Trim(Mid(ListBox1.list(r), 14, 37)))
                plisn = (Trim(Mid(ListBox1.list(r), 7, 5)))
                Updatepartslist ListBox1.list(r), cagePart
            End If
        End If
        If Right(ListBox1.list(r), 1) = "B" Then
            If Not Trim(Mid(ListBox1.list(r), 19, 32)) = "" Then
                ComboBox6.AddItem (Trim(Mid(ListBox1.list(r), 16, 13)))
                If Not plisn = Trim(Mid(ListBox1.list(r), 7, 5)) Then
                    cagePart = ""
                End If
                Updatepartslist ListBox1.list(r), cagePart
            End If
        End If
        If Not last = (Trim(Mid(ListBox1.list(r), 7, 5))) And Not Trim(Mid(ListBox1.list(r), 7, 5)) = "" Then
            ComboBox1.AddItem (Trim(Mid(ListBox1.list(r), 7, 5)))
        End If
        last = Trim(Mid(ListBox1.list(r), 7, 5))
    Next
    
    s = SortParts(ComboBox2.list, ComboBox2)
    s = SortParts(ComboBox6.list, ComboBox6)
    
    ComboBox1.Value = ""
    ComboBox1.Value = bplisn
    'ComboBox2 = ""
    'ComboBox6 = ""
    
    
    
    Dim clr As Variant
    For Each clr In SAP
        SAP.remove (1)
    Next clr
    'make new collection to update SAP
    SAP_Calc
    
    displaylistbox2
    ComboBox3.list = ComboBox1.list
    ComboBox4.list = ComboBox1.list
    ComboBox8.list = ComboBox2.list
    'ComboBox9.list = ComboBox1.list
End Sub
Private Sub ComboBox4_Change()
    '''PLISN on Add/Edit Cards
    If OptionButton5.Value = False Then
        ComboBox3.Value = UCase(ComboBox4.Value)
    Else
        ComboBox1.Value = ComboBox4.Value
    End If
End Sub
Private Sub Update_Card()
    '''Update A Card
    Dim TextLine As String
    Dim r As Variant
    Dim txbx As Variant
    Dim cardid As String
    Dim found As Boolean
    Dim plisnfix As String
    Dim csnwas As String
    csnwas = TextBox20.Value
    If ComboBox4.Value = "" Then Exit Sub
    cardid = TextBox20.Value & TextBox21.Value
    
    TextLine = TextLine & fixacard
    
    For r = ListBox1.TopIndex To ListBox1.ListCount - 1
        If Trim(Mid(ListBox1.list(r), 7, 5)) = ComboBox1.Value Then
            If Right(ListBox1.list(r), 3) = cardid Then
                found = True
                ListBox1.list(r) = UCase(TextLine)
                'TextBox20.Value = "00"
                'TextBox20.Value = csnwas
                TextBox11.Value = Trim(Mid(TextLine, 19, 32))
                GoTo 1
            End If
            
        Else:
            If found = True Then
                GoTo 1
            ElseIf ListBox1.ListCount - 1 = r Then
                If found = False Then
                    If MsgBox("Card doesn't exist. Would you like to add it?", vbYesNo) = vbYes Then
                        CommandButton20_Click
                    End If
                    GoTo 1
                End If
            End If
        End If
    Next
1   TextBox23.Value = "Your card has been updated"
End Sub
Function fixacard()
    Dim Text As Variant
    Dim start As Variant
    Dim r As Variant
    Dim c As Variant
    Dim TextLine As String

    Text = GetTextArray
   
    
    For r = 0 To UBound(Text)
        If Trim(Left(Text(r).Tag, 1)) = "" Then
            TextLine = TextLine & Trim(Text(r)) & Application.WorksheetFunction.Rept(" ", Text(r).MaxLength - Len(Trim(Text(r))))
        ElseIf Trim(Left(Text(r).Tag, 1)) = 0 Then
            If Not Text(r) = "" Then
                TextLine = TextLine & Application.WorksheetFunction.Rept(0, Text(r).MaxLength - Len(Trim(Text(r)))) & Trim(Text(r))
            Else
                TextLine = TextLine & Application.WorksheetFunction.Rept(" ", Text(r).MaxLength)
            End If
        End If
    Next
    Dim plisnfix As String
    plisnfix = ComboBox4.Value
    Do While Len(plisnfix) < 5
        plisnfix = plisnfix & " "
    Loop
    TextLine = TextBox6.Value & plisnfix & TextBox8.Value & TextLine & TextBox20.Value & TextBox21.Value
    
    
    fixacard = UCase(TextLine)
End Function
Private Sub CommandButton3_Click()
    'insert line above
    Addline "", True
End Sub
Private Sub CommandButton5_Click()
    '''insert line below
    Addline "", False
End Sub
Private Sub Frame9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox23.Value = "The PCCN, PLISN, and TOCC are the same for each card on the 036 report"
End Sub


Private Sub MultiPage2_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBox23.Value = "A card - Indenture Level, CAGE Code, Part Number, RNCC, RVNC, DAC, PPSL, EC, Item Name, SL, SLAC"
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "B card - NSN, UM, UM-Price, UI, UI Price, UI Conversion factor, QUP, SMR, DMIL, PLT, HCI, PSPC, PMIC, ADPEC"
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "C card - NHA PLISN, "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "D card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "E card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "F card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "G card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "H card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "J card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "K card - "
    TextBox23.Value = TextBox23.Value & Chr(10) & "==============================" & Chr(10) & "M card - "
End Sub
Private Sub TextBox11_Change()
    'Part Number
    TextBox11.Value = UCase(TextBox11.Value)
    TextBox24.Value = TextBox5.Value
    TextBox164.Value = TextBox11.Value
End Sub
Private Sub TextBox20_Change()
    '''CSN with spin button
    On Error GoTo 1
    SpinButton2.Value = Int(TextBox20.Value)
    UpdateCardsCSN
1   'Error
2
    
End Sub

'===========================================================================
'''Mulitpage1 Tab 2
'===========================================================================

Private Sub ComboBox3_Change()
    'PLISN box on PLISN add/edit page
    ComboBox3.Value = UCase(ComboBox3.Value)
    If OptionButton5.Value = False Then
        ComboBox1.Value = UCase(ComboBox3.Value)
        GoTo 1
    Else
        'What to do to the raw edit box when you want to add
        Frame98.Visible = False
    End If
1   updateplisngap ComboBox3
    If ListBox4.ListCount > 0 Then
        Label137.Caption = "Indenture " & Left(ListBox4.list(0), 2)
    Else:
        Label137.Caption = ""
    End If
End Sub
Function Updateaddpage()
    ComboBox8.Value = ComboBox2.Value
End Function
'==============================================================================================
'''Info Window Update
'==============================================================================================
Function InfoWindow(box As Variant, Info As String, MorInfo As String)
    TextBox23.Value = box.ControlTipText & ": " & box.Value & Chr(10)
    If Mid(box.Tag, 2, 2) > 0 Then
        TextBox23.Value = TextBox23.Value & "Starts at: " & Mid(box.Tag, 2, 2) & Chr(10)
        TextBox23.Value = TextBox23.Value & "Digit(s) allowed: " & box.MaxLength & Chr(10)
    
    End If
    TextBox23.Value = TextBox23.Value & Info & Chr(10)
    TextBox23.Value = TextBox23.Value & MorInfo & Chr(10)
End Function
Private Sub TextBox9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox9, "", "")
End Sub
Private Sub TextBox10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox10, "", "")
End Sub

Private Sub TextBox11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox11, "", "")
End Sub
Private Sub TextBox12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox12, "", "")
End Sub
Private Sub TextBox13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox13, "", "")
End Sub
Private Sub TextBox14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox14, "", "")
End Sub
Private Sub TextBox15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox15, "", "")
End Sub
Private Sub TextBox16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox16, "", "")
End Sub
Private Sub TextBox17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox17, "", "")
End Sub
Private Sub TextBox18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox18, "", "")
End Sub
Private Sub TextBox19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox19, "", "")
End Sub
Private Sub TextBox20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox20, "", "")
End Sub
Private Sub TextBox21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox21, "", "")
End Sub
Private Sub TextBox22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox22, "", "")
End Sub
Private Sub TextBox24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox24, "", "")
End Sub
Private Sub TextBox25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox25, "", "")
End Sub
Private Sub TextBox26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox26, "", "")
    
End Sub
Private Sub TextBox27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(textbox27, "", "")
End Sub
Private Sub TextBox28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox28, "", "")
End Sub
Private Sub TextBox29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox29, "", "")
End Sub
Private Sub TextBox30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox30, "", "")
End Sub
Private Sub TextBox31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox31, "", "")
End Sub
Private Sub TextBox32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox32, "", "")
End Sub
Private Sub TextBox33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox33, "", "")
End Sub
Private Sub TextBox34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox34, "", "")
End Sub
Private Sub TextBox35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox35, "", "")
End Sub
Private Sub TextBox36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox36, "", "")
End Sub
Private Sub TextBox37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox37, "", "")
End Sub
Private Sub TextBox38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox38, "", "")
End Sub
Private Sub TextBox39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox39, "", "")
End Sub
Private Sub TextBox40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox40, "", "")
End Sub
Private Sub TextBox41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox41, "", "")
End Sub
Private Sub TextBox42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox42, "", "")
End Sub
Private Sub TextBox43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox43, "", "")
End Sub
Private Sub TextBox44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox44, "", "")
End Sub


Private Sub TextBox49_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox49, "", "")
End Sub
Private Sub TextBox50_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox50, "", "")
End Sub
Private Sub TextBox51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox51, "", "")
End Sub
Private Sub TextBox52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox52, "", "")
End Sub
Private Sub TextBox53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox53, "", "")
End Sub
Private Sub TextBox54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox54, "", "")
End Sub
Private Sub TextBox55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox55, "", "")
End Sub
Private Sub TextBox56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox56, "", "")
End Sub
Private Sub TextBox57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox57, "", "")
End Sub
Private Sub TextBox58_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox58, "", "")
End Sub
Private Sub TextBox59_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox59, "", "")
End Sub
Private Sub TextBox60_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox60, "", "")
End Sub
Private Sub TextBox61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox61, "", "")
End Sub
Private Sub TextBox62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox62, "", "")
End Sub
Private Sub TextBox63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox63, "", "")
End Sub
Private Sub TextBox64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox64, "", "")
End Sub
Private Sub TextBox65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox65, "", "")
End Sub

Private Sub TextBox67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox67, "", "")
End Sub
Private Sub TextBox68_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox68, "", "")
End Sub
Private Sub TextBox69_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox69, "", "")
End Sub
Private Sub TextBox70_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox70, "", "")
End Sub
Private Sub TextBox71_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox71, "", "")
End Sub
Private Sub TextBox72_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox72, "", "")
End Sub
Private Sub TextBox73_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox73, "", "")
End Sub
Private Sub TextBox74_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox74, "", "")
End Sub
Private Sub TextBox75_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox75, "", "")
End Sub
Private Sub TextBox76_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox76, "", "")
End Sub
Private Sub TextBox77_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox77, "", "")
End Sub
Private Sub TextBox78_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox78, "", "")
End Sub
Private Sub TextBox79_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox79, "", "")
End Sub
Private Sub TextBox80_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox80, "", "")
End Sub
Private Sub TextBox81_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox81, "", "")
End Sub
Private Sub TextBox82_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox82, "", "")
End Sub
Private Sub TextBox83_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox83, "", "")
End Sub
Private Sub TextBox84_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox84, "", "")
End Sub
Private Sub TextBox85_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox85, "", "")
End Sub
Private Sub TextBox86_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox86, "", "")
End Sub
Private Sub TextBox87_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox87, "", "")
End Sub
Private Sub TextBox88_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox88, "", "")
End Sub
Private Sub TextBox89_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox89, "", "")
End Sub
Private Sub TextBox90_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox90, "", "")
End Sub
Private Sub TextBox91_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox91, "", "")
End Sub
Private Sub TextBox92_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox92, "", "")
End Sub
Private Sub TextBox93_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox93, "", "")
End Sub
Private Sub TextBox94_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox94, "", "")
End Sub
Private Sub TextBox95_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox95, "", "")
End Sub
Private Sub TextBox96_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox96, "", "")
End Sub
Private Sub TextBox97_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox97, "", "")
End Sub
Private Sub TextBox98_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox98, "", "")
End Sub
Private Sub TextBox99_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox99, "", "")
End Sub
Private Sub TextBox100_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox100, "", "")
End Sub
Private Sub TextBox101_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox101, "", "")
End Sub
Private Sub TextBox102_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox102, "", "")
End Sub
Private Sub TextBox103_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox103, "", "")
End Sub
Private Sub TextBox104_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox104, "", "")
End Sub
Private Sub TextBox105_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox105, "", "")
End Sub
Private Sub TextBox106_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox106, "", "")
End Sub
Private Sub TextBox107_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox107, "", "")
End Sub
Private Sub TextBox108_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox108, "", "")
End Sub
Private Sub TextBox109_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox109, "", "")
End Sub
Private Sub TextBox110_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox110, "", "")
End Sub
Private Sub TextBox111_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox111, "", "")
End Sub
Private Sub TextBox112_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox112, "", "")
End Sub
Private Sub TextBox113_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox113, "", "")
End Sub
Private Sub TextBox114_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox114, "", "")
End Sub
Private Sub TextBox115_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox115, "", "")
End Sub
Private Sub TextBox116_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox116, "", "")
End Sub
Private Sub TextBox117_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox117, "", "")
End Sub
Private Sub TextBox118_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox118, "", "")
End Sub
Private Sub TextBox119_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox119, "", "")
End Sub
Private Sub TextBox120_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox120, "", "")
End Sub
Private Sub TextBox121_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox121, "", "")
End Sub
Private Sub TextBox122_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox122, "", "")
End Sub
Private Sub TextBox123_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox123, "", "")
End Sub
Private Sub TextBox124_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox124, "", "")
End Sub
Private Sub TextBox125_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox125, "", "")
End Sub
Private Sub TextBox126_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox126, "", "")
End Sub
Private Sub TextBox127_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox127, "", "")
End Sub
Private Sub TextBox128_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox128, "", "")
End Sub
Private Sub TextBox129_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox129, "", "")
End Sub
Private Sub TextBox130_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox130, "", "")
End Sub
Private Sub TextBox131_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox131, "", "")
End Sub
Private Sub TextBox132_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox132, "", "")
End Sub
Private Sub TextBox133_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox133, "", "")
End Sub
Private Sub TextBox134_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox134, "", "")
End Sub
Private Sub TextBox135_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox135, "", "")
End Sub
Private Sub TextBox136_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox136, "", "")
End Sub
Private Sub TextBox137_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox137, "", "")
End Sub
Private Sub TextBox138_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox138, "", "")
End Sub
Private Sub TextBox139_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox139, "", "")
    
End Sub
Private Sub TextBox140_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox140, "", "")
End Sub
Private Sub TextBox141_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox141, "", "")
End Sub
Private Sub TextBox142_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox142, "", "")
End Sub
Private Sub TextBox143_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox143, "", "")
End Sub
Private Sub TextBox144_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox144, "", "")
End Sub
Private Sub TextBox145_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox145, "", "")
End Sub
Private Sub TextBox146_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox146, "", "")
End Sub
Private Sub TextBox147_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox147, "", "")
End Sub
Private Sub TextBox148_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox148, "", "")
End Sub
Private Sub TextBox149_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox149, "", "")
End Sub
Private Sub TextBox150_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox150, "", "")
End Sub
Private Sub TextBox151_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox151, "", "")
End Sub
Private Sub TextBox152_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox152, "", "")
End Sub
Private Sub TextBox153_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox153, "", "")
End Sub
Private Sub TextBox154_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox154, "", "")
End Sub
Private Sub TextBox155_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox155, "", "")
End Sub
Private Sub TextBox156_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox156, "", "")
End Sub
Private Sub TextBox157_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox157, "", "")
End Sub
Private Sub TextBox158_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox158, "", "")
End Sub
Private Sub TextBox159_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox159, "", "")
End Sub
Private Sub TextBox160_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox160, "", "")
End Sub
Private Sub TextBox161_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox161, "", "")
End Sub
Private Sub TextBox162_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox162, "", "")
End Sub
Private Sub TextBox163_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox163, "", "")
End Sub



Private Sub TextBox178_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox178, "", "")
End Sub

Private Sub TextBox180_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox180, "", "")
End Sub

Private Sub TextBox182_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox182, "", "")
End Sub
Private Sub TextBox183_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox183, "", "")
End Sub


Private Sub TextBox187_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox187, "", "")
End Sub
Private Sub TextBox188_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox188, "", "")
End Sub
Private Sub TextBox189_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox189, "", "")
End Sub
Private Sub TextBox190_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox190, "", "")
End Sub
Private Sub TextBox191_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox191, "", "")
End Sub
Private Sub TextBox192_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox192, "", "")
End Sub
Private Sub TextBox193_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox193, "", "")
End Sub
Private Sub TextBox194_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox194, "", "")
End Sub
Private Sub TextBox195_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox195, "", "")
End Sub
Private Sub TextBox196_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox196, "", "")
End Sub
Private Sub TextBox197_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox197, "", "")
End Sub
Private Sub TextBox198_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(TextBox198, "", "")
End Sub
Private Sub ComboBox7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Variant
    m = InfoWindow(ComboBox7, "", "")
End Sub

