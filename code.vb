'================================= COL/CE/2020/F/056 K.R.MADHUSHANKHA ===================================
'                                       ===========================
'                                           ==================


'Declarations
Dim a, b, c, X, rootValue(1 To 2) As Double
Dim St As String

'When the Form is loading
Public Sub Form_Load()
    SystemStart 'calling for the stating function.
End Sub

'New button action
Private Sub CmdNew_Click()
    Clear   'calling for the clearing function.
End Sub

'About button action
Private Sub CmdAbout_Click()
    About ' Calling for the About function. (Info about the owner and the program)
End Sub

'System Close controls
Private Sub CmdClose_Click()
    SystemClose 'Calling the System exit function.
End Sub

Private Sub LblClose_Click()
    SystemClose 'Calling the System exit function.
End Sub

'Result button action
Private Sub CmdResult_Click()
    'Checking for empty Txt fields.
    If (TxtA.Text) = Empty Or (TxtB.Text) = Empty Or (TxtC.Text) = Empty Then
        LblError.Caption = "Empty text fields found! or Press TAB to go to next TexxBox."
        LblError.ForeColor = vbRed
        TxtA.SetFocus
    Else
        DeltaX
    End If
End Sub

'Get Values From The Text Fields.
Private Sub TxtA_LostFocus()
    a = Val(TxtA.Text) 'Get "a" Value
End Sub

Private Sub TxtB_LostFocus()
    b = Val(TxtB.Text)  'Get "b" Value
End Sub

Private Sub TxtC_LostFocus()
    c = Val(TxtC.Text) 'Get "c" Value
End Sub



'=====================================================================
'==================== Functions ============================================================================
'System Start
Public Function SystemStart()
    LblError.Caption = ""
    LblStatus.BackColor = vbWhite
    St = "1234567890.-+"
End Function

'System Close
Private Function SystemClose()
    Msg = MsgBox("Are You Sure?", vbYesNo, "Exit")
    'System exit when click yes
    If Msg = vbYes Then
        End
    End If
End Function

'Clearing all the text fields and notifications
Public Function Clear()
    Msg = MsgBox("Are you sure to clean the previous values?", vbYesNo, "Clear")
    'Clearing the text fields when click yes
    If Msg = vbYes Then
        TxtA.Text = ""
        TxtB.Text = ""
        TxtC.Text = ""
        LblRoot1.Caption = "- -"
        LblRoot2.Caption = "- -"
        LblError.Caption = "Cleaned!"
        LblError.ForeColor = &HC000&
        LblStatus.Caption = "Status"
        LblStatus.BackColor = vbWhite
        TxtA.SetFocus 'focus on first text field again
    End If
End Function

'Check  b^2 - 4ac
Public Function DeltaX()
    Status = Array("There are 2 roots.", "Roots are equal.", "No real solution.", "Roots are generated successfully!", "")
    X = b ^ 2 - 4 * a * c
    'check X value
    If X > 0 Then
        LblStatus.Caption = Status(0)
        LblStatus.BackColor = &HC000&
        root
        LblError.Caption = Status(3)
        LblError.ForeColor = &HC000&
    ElseIf X = 0 Then
        LblStatus.Caption = Status(1)
        LblStatus.BackColor = &HC000&
        root
        LblError.Caption = Status(3)
        LblError.ForeColor = &HC000&
    ElseIf X < 0 Then
        LblStatus.Caption = Status(2)
        LblStatus.BackColor = vbRed
        LblRoot1.Caption = "- -"
        LblRoot2.Caption = "- -"
        LblError.Caption = Status(4)
        CmdNew.SetFocus
    Else
        LblStatus.Caption = "Error"
    End If
End Function

'find roots
Public Function root()
    rootValue(1) = (-b + Sqr(X)) / (2 * a)
    rootValue(2) = (-b - Sqr(X)) / (2 * a)
    'Roots send to the variables
    LblRoot1.Caption = Round(rootValue(1), 2)
    LblRoot2.Caption = Round(rootValue(2), 2)
    'focus on new button
    CmdNew.SetFocus
    'Listing to the recently data
    Listing
End Function

'About
Public Function About()
    Owner = Array("Created By COLCE2020F056 K.R.MADHUSHANKHA.", "2nd Year 1st Semester.", "Solution of Quadratic Equation Can be determined using this program.", "")
    MsgBox Owner(2) & vbCrLf & Owner(3) & vbCrLf & Owner(3) & vbCrLf & Owner(0) & vbCrLf & Owner(1) & vbCrLf & Owner(3), vbOKOnly, "About"
End Function

'Show previous data as recently
Public Function Listing()
    List1.AddItem a
    List2.AddItem b
    List3.AddItem c
    List4.AddItem Round(rootValue(1), 2)
    List5.AddItem Round(rootValue(2), 2)
End Function

'Check is there any empty text fields
Public Function CheckingEmpty()
    If (TxtA.Text) = Empty Or (TxtB.Text) = Empty Or (TxtC.Text) = Empty Then
        LblError.Caption = "Empty text fields found! or Press TAB to go to next TexxBox."
        LblError.ForeColor = vbRed
    End If
End Function

' -- Functions -----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------


'Validation and Submit
'Validation for a
Private Sub TxtA_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then
        If InStr(St, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            LblError.Caption = "Only numeric values will be allowed"
            LblError.ForeColor = vbRed
        Else
            LblError.Caption = ""
        End If
    End If
    'Submit
    If KeyAscii = 13 Then
        CheckingEmpty
    End If
End Sub

'Validation for b
Private Sub TxtB_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then
        If InStr(St, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            LblError.Caption = "Only numeric values will be allowed"
            LblError.ForeColor = vbRed
        Else
            LblError.Caption = ""
        End If
    End If
    'Submit
    If KeyAscii = 13 Then
        CheckingEmpty
    End If
End Sub

'Validation for c
Private Sub TxtC_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then
        If InStr(St, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            LblError.Caption = "Only numeric values will be allowed"
            LblError.ForeColor = vbRed
        Else
            LblError.Caption = ""
        End If
    End If
    'Submit
    If KeyAscii = 13 Then
        CheckingEmpty
    End If
End Sub
' ==========================================================



















'==================================== Move window ===========================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmMove = True
    DragX = X
    DragY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nx, ny
    If FrmMove Then
        nx = Main.Left + X - DragX
        ny = Main.Top + Y - DragY
        Main.Left = nx
        Main.Top = ny
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nx, ny
    nx = Main.Left + X - DragX
    ny = Main.Top + Y - DragY
    Main.Left = nx
    Main.Top = ny
    FrmMove = False
End Sub
Private Sub PlayerSlim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub
Private Sub PlayerSlim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub
Private Sub PlayerSlim_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub
Private Sub PlayerTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub
Private Sub PlayerTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub
Private Sub PlayerTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

'-----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------


