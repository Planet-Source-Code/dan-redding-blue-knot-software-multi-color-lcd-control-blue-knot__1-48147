VERSION 5.00
Begin VB.Form frmUCTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blue Knot LCD Test"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmUCTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1440
      Left            =   -15
      TabIndex        =   15
      Top             =   6120
      Width           =   11430
      Begin VB.CommandButton cmdColor 
         Caption         =   "Custom C&olor Demo"
         Height          =   315
         Left            =   6150
         TabIndex        =   18
         Top             =   960
         Width           =   2025
      End
      Begin VB.CommandButton cmdTimed 
         Caption         =   "&Timed Demo"
         Height          =   315
         Left            =   4530
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.Timer tmrDemo 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5460
         Top             =   960
      End
      Begin VB.CheckBox chkPad 
         Caption         =   "&Pad w/ ""Spaces"""
         Height          =   465
         Left            =   10290
         TabIndex        =   10
         Top             =   480
         Width           =   1035
      End
      Begin VB.HScrollBar hsDelay 
         Height          =   225
         LargeChange     =   50
         Left            =   8880
         Max             =   200
         SmallChange     =   25
         TabIndex        =   9
         Top             =   600
         Width           =   1305
      End
      Begin VB.CommandButton cmdDemo 
         Caption         =   "De&mo"
         Height          =   315
         Left            =   3180
         TabIndex        =   16
         Top             =   960
         Width           =   1305
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10170
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear &Form"
         Height          =   315
         Left            =   1830
         TabIndex        =   13
         Top             =   960
         Width           =   1305
      End
      Begin VB.ComboBox cboBack 
         Height          =   315
         ItemData        =   "frmUCTest.frx":0E42
         Left            =   5670
         List            =   "frmUCTest.frx":0E5B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   1575
      End
      Begin VB.CheckBox chkBold 
         Caption         =   """Bol&d"""
         Height          =   225
         Left            =   2460
         TabIndex        =   4
         Top             =   255
         Width           =   795
      End
      Begin VB.CheckBox chkBright 
         Caption         =   """&Bright"""
         Height          =   225
         Left            =   3300
         TabIndex        =   5
         Top             =   255
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         ItemData        =   "frmUCTest.frx":0E8D
         Left            =   990
         List            =   "frmUCTest.frx":0EB8
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "&Show All Characters"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1665
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display"
         Default         =   -1  'True
         Height          =   315
         Left            =   7230
         TabIndex        =   1
         Top             =   570
         Width           =   1095
      End
      Begin VB.TextBox txtMessage 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Text            =   "Type your message here!"
         Top             =   570
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "All demos work with the current background color."
         Height          =   465
         Left            =   8220
         TabIndex        =   19
         Top             =   930
         Width           =   1965
      End
      Begin VB.Label lblDelay 
         Caption         =   "&Delay"
         Height          =   255
         Left            =   8400
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWarning 
         Caption         =   "Note: Background Color may distort the font color, causing it not to appear 'As Advertised'"
         Height          =   435
         Left            =   7290
         TabIndex        =   11
         Top             =   150
         Width           =   3915
      End
      Begin VB.Label lblBack 
         Caption         =   "Bac&kground Color:"
         Height          =   285
         Left            =   4290
         TabIndex        =   6
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label lblColor 
         Caption         =   "Font &Color:"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   915
      End
   End
   Begin prjLCD.bkLCD LCD 
      Left            =   240
      Top             =   180
      _ExtentX        =   1058
      _ExtentY        =   503
      Bold            =   0   'False
      Hue             =   120
      Bright          =   -1  'True
      Background      =   6
      Color           =   6
   End
End
Attribute VB_Name = "frmUCTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Demo form for the bkLCD control
'2003 Dan Redding / Blue Knot Software
'Free to use, mentions and links to http://www.blueknot.com are appreciated!

'Writes directly to the form, could also write to picture box other control with
'an .hDC & .Width property and .Refresh method

Private Sub cboBack_Click()
    'ListIndex corresponds w/ 0-based property list, so no need to
    'Convert or use ItemData
    LCD.Background = cboBack.ListIndex
    'BackgroundColor returns the actual Long color code selected Background so
    'that form or picturebox can be set to match
    frmUCTest.BackColor = LCD.BackgroundColor
    'Clear screen and start over
    frmUCTest.Cls
End Sub

Private Sub cboColor_Click()
    'ListIndex corresponds w/ 0-based property list, so no need to
    'Convert or use ItemData
    LCD.Color = cboColor.ListIndex
End Sub

Private Sub chkBold_Click()
    LCD.Bold = chkBold.Value = vbChecked
End Sub

Private Sub chkBright_Click()
    LCD.Bright = chkBright.Value = vbChecked
End Sub

Private Sub cmdClear_Click()
    frmUCTest.Cls
End Sub

Private Sub cmdColor_Click()
Dim iLoop As Integer
    EndDemo
    frmUCTest.Cls
    With LCD
        For iLoop = 0 To 26
            .Hue = iLoop * 9
            .Bright = False
            .Bold = False
            WriteOnForm "  regular text  ", False
            .Bold = True
            WriteOnForm " bold text  ", False
            .Bold = False
            .Bright = True
            WriteOnForm " bright text  ", False
            .Bold = True
            WriteOnForm " bold & bright text  "
        Next iLoop
    End With
End Sub

Private Sub cmdDemo_Click()
    'A little play, showing off most of the features of the LCD control (custom Hue
    'not shown, and could have used CharHeight/CharWidth properties and a delay to
    'overprint the writing)
    
    'Works with the selected background but overrides all others
    EndDemo
    With frmUCTest
        .Enabled = False
        .Cls
        .MousePointer = vbHourglass
    End With
    With LCD
        .Color = bkLcd_Cyan
        .Bold = True
        .Bright = True
        WriteOnForm "NEW MULTI-COLOR LCD CONTROL", , 10, True
        WriteOnForm String$(64, 95), , , True
        WriteOnForm ""
        .Color = bkLcd_Red
        .Bright = False
        WriteOnForm "13 pre-defined forecolors, 7 backcolors (but black works best!)", , 25
        WriteOnForm "Red, ", False
        .Color = bkLcd_Orange
        WriteOnForm "Orange, ", False
        .Color = bkLcd_Yellow
        WriteOnForm "Yellow, ", False
        .Color = bkLcd_Lime
        WriteOnForm "Lime, ", False
        .Color = bkLcd_Green
        WriteOnForm "Green, ", False
        .Color = bkLcd_Teal
        WriteOnForm "Teal, ", False
        .Color = bkLcd_Cyan
        WriteOnForm "Cyan, ", False
        .Color = bkLcd_Blue
        WriteOnForm "Blue, "
        .Color = bkLcd_Purple
        WriteOnForm "Purple, ", False
        .Color = bkLcd_Violet
        WriteOnForm "Violet, ", False
        .Color = bkLcd_Magenta
        WriteOnForm "Magenta, ", False
        .Color = bkLcd_Pink
        WriteOnForm "Pink, ", False
        .Color = bkLcd_White
        WriteOnForm " & White"
        .Color = bkLcd_Red
        .Bright = True
        WriteOnForm ""
        WriteOnForm "All Colors also have a 'bright' intensity", , 25
        WriteOnForm "Red, ", False
        .Color = bkLcd_Orange
        WriteOnForm "Orange, ", False
        .Color = bkLcd_Yellow
        WriteOnForm "Yellow, ", False
        .Color = bkLcd_Lime
        WriteOnForm "Lime, ", False
        .Color = bkLcd_Green
        WriteOnForm "Green, ", False
        .Color = bkLcd_Teal
        WriteOnForm "Teal, ", False
        .Color = bkLcd_Cyan
        WriteOnForm "Cyan, ", False
        .Color = bkLcd_Blue
        WriteOnForm "Blue, "
        .Color = bkLcd_Purple
        WriteOnForm "Purple, ", False
        .Color = bkLcd_Violet
        WriteOnForm "Violet, ", False
        .Color = bkLcd_Magenta
        WriteOnForm "Magenta, ", False
        .Color = bkLcd_Pink
        WriteOnForm "Pink, ", False
        .Color = bkLcd_White
        WriteOnForm " & White"
        WriteOnForm ""
        .Color = bkLcd_Green
        .Bold = False
        .Bright = False
        WriteOnForm "2 ""Fonts"": Regular text...", False, 60
        .Bold = True
        WriteOnForm "and bold letters.", , 40
        
        .Color = bkLcd_Violet
        .Bright = True
        WriteOnForm "", , , True
        WriteOnForm "Slow printing and line padding available...", , 50, True
        WriteOnForm "", , , True
        
        .Color = bkLcd_Blue
        .Bold = False
        .Bright = True
        WriteOnForm "lots of available characters:", , , True
        .Bold = True
        WriteOnForm "{SHOWALL}", , 25
        .Bold = False
        WriteOnForm "  ( × = Alt+0215 or Chr$(215), ÷ = Alt+0247 or CHR$(247))"
        .Bold = True
        WriteOnForm ""
        .Color = bkLcd_Pink
        WriteOnForm "Original Graphics from ""Vacuum Fluorescent Display Simulator"""
        WriteOnForm " by Peter Wilson ('Regular' font) February 2003"
        .Bold = False
        WriteOnForm ""
        WriteOnForm "New characters & New BOLD font added by Dan Redding August 2003"
        WriteOnForm "(Added +*'×÷!@#$_"";<> and ? ;  Revised - and D)"
        .Color = bkLcd_Yellow
        .Bright = False
        WriteOnForm ""
        .Bold = True
        WriteOnForm "Send comments/suggestions to Dan@blueknot.com", , 30, True
        .Bold = False
        WriteOnForm "             (and please vote on PSC!)", , 30, True
        .Color = bkLcd_Cyan
    End With
    'Set option controls to match current
    With frmUCTest
        .Enabled = True
        .MousePointer = vbDefault
        .cboColor.ListIndex = bkLcd_Cyan
        .chkBold.Value = vbUnchecked
        .chkBright.Value = vbUnchecked
        .txtMessage.Text = "Well? What do you think?"
        .txtMessage.SetFocus
    End With

End Sub

Private Sub cmdDisplay_Click()
    'Display the message
    EndDemo
    frmUCTest.Enabled = False
    WriteOnForm txtMessage.Text, , hsDelay.Value, chkPad.Value = vbChecked
    frmUCTest.Enabled = True
    SelectAll txtMessage
    txtMessage.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload frmUCTest
End Sub

Private Sub cmdShowAll_Click()
    'Special code that dumps all printable characters.  Only purpose is demo
    WriteOnForm "{SHOWALL}"
End Sub

Private Sub cmdTimed_Click()
    If Left$(cmdTimed.Caption, 4) = "Stop" Then
        EndDemo
    Else
        cmdTimed.Caption = "Stop " & cmdTimed.Caption
        frmUCTest.Cls
        With LCD
            .Color = bkLcd_Red
            .Bright = True
            .Bold = True
            WriteOnForm "Special Chars for Math:"
            WriteOnForm "  10+5=15", False
            WriteOnForm "  10-5=5", False
            WriteOnForm "  10×5=50", False
            WriteOnForm "  10÷5=2", False
            .Bright = False
            WriteOnForm "  10÷0=__?"
            WriteOnForm ""
            .Bold = False
            .Bright = True
            .Color = bkLcd_Green
            WriteOnForm "The current date and time is:"
            .Bold = True
        End With
        tmrDemo.Enabled = True
    End If
End Sub

Private Sub Form_DblClick()
    WriteOnForm "Hello!  You Double-clicked the form!"
End Sub

Private Sub Form_Load()
    'Initialize
    'AutoRedraw property set at design time
    cboColor.ListIndex = LCD.Color
    cboBack.ListIndex = LCD.Background
    frmUCTest.BackColor = LCD.BackgroundColor
End Sub

Private Sub WriteOnForm(sMessage As String, Optional blnLineFeed As Boolean = True, Optional lDelay As Long = 0&, Optional blnPad As Boolean = False)
    'Little wrapper function that prints the message and takes care of the linefeed
    With frmUCTest
        If .CurrentY + LCD.CharHeight > .fraOptions.Top Then
            .Cls
        End If
        LCD.PrintOut frmUCTest, .CurrentX, .CurrentY, sMessage, blnLineFeed, lDelay, blnPad
        .CurrentX = LCD.LastX
        .CurrentY = LCD.LastY
        .Refresh
    End With
End Sub

Private Sub tmrDemo_Timer()
    With LCD
        frmUCTest.CurrentX = 3 * .CharWidth
        frmUCTest.CurrentY = 5 * .CharHeight
        WriteOnForm Format$(Now(), "mmm dd, yyyy - hh:nn:ss")
    End With
End Sub

Private Sub txtMessage_GotFocus()
    SelectAll txtMessage
End Sub

Private Sub txtMessage_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdDisplay.Value = True
End Sub

Private Sub SelectAll(txt As TextBox)
    With txt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub EndDemo()
    If Left$(cmdTimed.Caption, 4) = "Stop" Then
        cmdTimed.Caption = Mid$(cmdTimed.Caption, 6)
        tmrDemo.Enabled = False
        frmUCTest.Cls
        With LCD
            .Background = cboBack.ListIndex
            .Color = cboColor.ListIndex
            .Bold = chkBold.Value = vbChecked
            .Bright = chkBright.Value = vbChecked
        End With
    End If
End Sub
