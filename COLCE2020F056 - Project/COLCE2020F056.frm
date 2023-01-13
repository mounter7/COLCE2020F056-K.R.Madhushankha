VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "COLCE2020F056 - Solution of Quadratic Equation"
   ClientHeight    =   5175
   ClientLeft      =   3870
   ClientTop       =   3225
   ClientWidth     =   12015
   Icon            =   "COLCE2020F056.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Recently Roots"
      Height          =   4575
      Left            =   6960
      TabIndex        =   20
      Top             =   480
      Width           =   4815
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3345
         Left            =   3840
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3345
         Left            =   2880
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3345
         Left            =   1920
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3345
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3345
         ItemData        =   "COLCE2020F056.frx":1542
         Left            =   0
         List            =   "COLCE2020F056.frx":1544
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Recently Roots"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "      a                    b                    c                 Root1            Root2"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   720
         Width           =   4815
      End
   End
   Begin VB.CommandButton CmdAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "About"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox TxtC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton CmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "New"
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton CmdResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find Root(s)"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   160
      X2              =   296
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.Label LblEq 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ax  + bx + c = 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   6735
   End
   Begin VB.Shape Shape4 
      Height          =   5535
      Left            =   12000
      Top             =   -360
      Width           =   15
   End
   Begin VB.Shape Shape3 
      Height          =   135
      Left            =   0
      Top             =   5160
      Width           =   12015
   End
   Begin VB.Shape Shape2 
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Label LblClose 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11400
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quadratic Equation"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   12135
   End
   Begin VB.Label LblError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notifications"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label LblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label LblRoot2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "- -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label LblRoot1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "- -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Root 2 :"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Root 1 :"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "c"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "b"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


