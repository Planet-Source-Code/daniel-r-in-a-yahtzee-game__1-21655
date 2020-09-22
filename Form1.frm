VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Yahtzee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Yahtzee"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Restart"
      Height          =   375
      Left            =   2520
      TabIndex        =   45
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   465
      Left            =   600
      TabIndex        =   44
      Top             =   480
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   820
      ButtonWidth     =   741
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "4"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "5"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0594
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1650
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   2175
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   0
         TabIndex        =   24
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   0
         TabIndex        =   23
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   0
         TabIndex        =   22
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1790
         TabIndex        =   43
         Top             =   4920
         Width           =   360
      End
      Begin VB.Label Label43 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1790
         TabIndex        =   42
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1790
         TabIndex        =   41
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   40
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   39
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label39 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   38
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   37
         Top             =   3720
         Width           =   360
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   36
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   35
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   34
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   33
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   32
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1790
         TabIndex        =   31
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   30
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   29
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   28
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   27
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   26
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1790
         TabIndex        =   25
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   1665
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper total"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Yahtzee"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   1665
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4200
         Width           =   1665
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Full house"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   1665
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Large straight"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   1665
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Small straight"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Four of a kind"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1665
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Three of a kind"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1665
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Two pairs"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "One pair"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1665
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Sixes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fives"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fours"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1665
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Threes"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Twos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ones"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roll em"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Not just for me but for all the good coders on the planet"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remember: If you like the code: VOTE!!"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Yahtzee v. 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   46
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rolls left : 3"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Yahtzee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dices...
Dim one As Single, two As Single, three As Single, four As Single, five As Single
'other...
Dim rolls As Integer
Dim hold1 As Boolean
Dim hold2 As Boolean
Dim hold3 As Boolean
Dim hold4 As Boolean
Dim hold5 As Boolean
Dim klicked As Boolean

Private Sub Command1_Click()
'sub for declaring the dices
Playsound "dice"
Toolbar1.Enabled = True
klicked = False
rolls = rolls + 1
repeat:
Randomize
If hold1 = False Then
    one = 7 * Rnd(1)
    If one > 7 Or one < 1 Then GoTo repeat
    If Int(one) = 1 Then Toolbar1.Buttons(1).Image = 1
    If Int(one) = 2 Then Toolbar1.Buttons(1).Image = 2
    If Int(one) = 3 Then Toolbar1.Buttons(1).Image = 3
    If Int(one) = 4 Then Toolbar1.Buttons(1).Image = 4
    If Int(one) = 5 Then Toolbar1.Buttons(1).Image = 5
    If Int(one) = 6 Then Toolbar1.Buttons(1).Image = 6
End If

If hold2 = False Then
    two = 7 * Rnd(1)
    If two > 7 Or two < 1 Then GoTo repeat
    If Int(two) = 1 Then Toolbar1.Buttons(3).Image = 1
    If Int(two) = 2 Then Toolbar1.Buttons(3).Image = 2
    If Int(two) = 3 Then Toolbar1.Buttons(3).Image = 3
    If Int(two) = 4 Then Toolbar1.Buttons(3).Image = 4
    If Int(two) = 5 Then Toolbar1.Buttons(3).Image = 5
    If Int(two) = 6 Then Toolbar1.Buttons(3).Image = 6
End If

If hold3 = False Then
    three = 7 * Rnd(1)
    If three > 7 Or three < 1 Then GoTo repeat
    If Int(three) = 1 Then Toolbar1.Buttons(5).Image = 1
    If Int(three) = 2 Then Toolbar1.Buttons(5).Image = 2
    If Int(three) = 3 Then Toolbar1.Buttons(5).Image = 3
    If Int(three) = 4 Then Toolbar1.Buttons(5).Image = 4
    If Int(three) = 5 Then Toolbar1.Buttons(5).Image = 5
    If Int(three) = 6 Then Toolbar1.Buttons(5).Image = 6
End If

If hold4 = False Then
    four = 7 * Rnd(1)
    If four > 7 Or four < 1 Then GoTo repeat
    If Int(four) = 1 Then Toolbar1.Buttons(7).Image = 1
    If Int(four) = 2 Then Toolbar1.Buttons(7).Image = 2
    If Int(four) = 3 Then Toolbar1.Buttons(7).Image = 3
    If Int(four) = 4 Then Toolbar1.Buttons(7).Image = 4
    If Int(four) = 5 Then Toolbar1.Buttons(7).Image = 5
    If Int(four) = 6 Then Toolbar1.Buttons(7).Image = 6
End If

If hold5 = False Then
    five = 7 * Rnd(1)
    If five > 7 Or five < 1 Then GoTo repeat
    If Int(five) = 1 Then Toolbar1.Buttons(9).Image = 1
    If Int(five) = 2 Then Toolbar1.Buttons(9).Image = 2
    If Int(five) = 3 Then Toolbar1.Buttons(9).Image = 3
    If Int(five) = 4 Then Toolbar1.Buttons(9).Image = 4
    If Int(five) = 5 Then Toolbar1.Buttons(9).Image = 5
    If Int(five) = 6 Then Toolbar1.Buttons(9).Image = 6
End If

If hold1 = True Then
    Toolbar1.Buttons.Item(1).Value = tbrPressed
    Else
    Toolbar1.Buttons.Item(1).Value = tbrUnpressed
End If

If hold2 = True Then
    Toolbar1.Buttons.Item(3).Value = tbrPressed
    Else
    Toolbar1.Buttons.Item(3).Value = tbrUnpressed
End If

If hold3 = True Then
    Toolbar1.Buttons.Item(5).Value = tbrPressed
    Else
    Toolbar1.Buttons.Item(5).Value = tbrUnpressed
End If

If hold4 = True Then
    Toolbar1.Buttons.Item(7).Value = tbrPressed
    Else
    Toolbar1.Buttons.Item(7).Value = tbrUnpressed
End If

If hold5 = True Then
    Toolbar1.Buttons.Item(9).Value = tbrPressed
    Else
    Toolbar1.Buttons.Item(9).Value = tbrUnpressed
End If
If Int(one) = Int(two) And Int(one) = Int(three) And Int(one) = Int(four) _
And Int(one) = Int(five) Then Playsound "wahoo"

Label6.Caption = "Rolls left : " & Str(3 - rolls)
If rolls = 3 Then Command1.Enabled = False
End Sub

Private Sub Command2_Click()
restart = MsgBox("Are you sure you want to restart?", vbOKCancel, "Confirm")
If restart = 1 Then
reinitialize
End If
End Sub

Private Sub Form_Load()
Command1_Click
End Sub

Private Sub Label10_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label10.FontBold Then
klicked = True
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
Label10.FontBold = True
Label29.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(3).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(5).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(7).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(9).Image = 4 Then sum = sum + 4
If sum > 0 Then
Label29.Caption = sum
Else
Label29.Caption = "-"
Playsound "doh"
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2
End If
End If
calculate
done
End Sub

Private Sub Label11_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label11.FontBold Then
klicked = True
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
Label11.FontBold = True
Label30.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(3).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(5).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(7).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(9).Image = 5 Then sum = sum + 5
If sum > 0 Then
Label30.Caption = sum
Else
Label30.Caption = "-"
Playsound "doh"
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2
End If
End If
calculate
done
End Sub

Private Sub Label12_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label12.FontBold Then
klicked = True
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
Label12.FontBold = True
Label31.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(3).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(5).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(7).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(9).Image = 6 Then sum = sum + 6
If sum > 0 Then
Label31.Caption = sum
Else
Label31.Caption = "-"
Playsound "doh"
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2
End If
End If
calculate
done
End Sub

Private Sub Label14_Click()
If klicked = False Then
If Not Label14.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label14.FontBold = True
Label33.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(3).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(5).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(7).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(9).Image = 6 Then sum = sum + 6
If sum >= 12 Then sum = 12: GoTo check
If sum < 12 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(3).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(5).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(7).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(9).Image = 5 Then sum = sum + 5
If sum >= 10 Then sum = 10: GoTo check
If sum < 10 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(3).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(5).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(7).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(9).Image = 4 Then sum = sum + 4
If sum >= 8 Then sum = 8: GoTo check
If sum < 8 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(3).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(5).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(7).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(9).Image = 3 Then sum = sum + 3
If sum >= 6 Then sum = 6: GoTo check
If sum < 6 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(3).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(5).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(7).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(9).Image = 2 Then sum = sum + 2
If sum >= 4 Then sum = 4: GoTo check
If sum < 4 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(3).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(5).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(7).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(9).Image = 1 Then sum = sum + 1
If sum >= 2 Then sum = 2: GoTo check
If sum < 2 Then sum = 0
check:
If sum > 0 Then
Label33.Caption = sum
Else
Label33.Caption = "-"
Playsound "doh"
End If
End If
End If
calculate
done
End Sub

Private Sub Label15_Click()
If klicked = False Then
If Not Label15.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label15.FontBold = True
Label34.FontBold = True
If Int(one) = Int(two) Then
chk = chk + 1
End If
If Int(one) = Int(three) Then
chk = chk + 1
End If
If Int(one) = Int(four) Then
chk = chk + 1
End If
If Int(one) = Int(five) Then
chk = chk + 1
End If
If Int(two) = Int(three) Then
chk = chk + 1
End If
If Int(two) = Int(four) Then
chk = chk + 1
End If
If Int(two) = Int(five) Then
chk = chk + 1
End If
If Int(three) = Int(four) Then
chk = chk + 1
End If
If Int(three) = Int(five) Then
chk = chk + 1
End If
If Int(four) = Int(five) Then
chk = chk + 1
End If

If chk = 2 Then
If Int(one) = Int(two) Then
sum = sum + Int(one) + Int(two)
End If
If Int(one) = Int(three) Then
sum = sum + Int(one) + Int(three)
End If
If Int(one) = Int(four) Then
sum = sum + Int(one) + Int(four)
End If
If Int(one) = Int(five) Then
sum = sum + Int(one) + Int(five)
End If
If Int(two) = Int(three) Then
sum = sum + Int(two) + Int(three)
End If
If Int(two) = Int(four) Then
sum = sum + Int(two) + Int(four)
End If
If Int(two) = Int(five) Then
sum = sum + Int(two) + Int(five)
End If
If Int(three) = Int(four) Then
sum = sum + Int(three) + Int(four)
End If
If Int(three) = Int(five) Then
sum = sum + Int(three) + Int(five)
End If
If Int(four) = Int(five) Then
sum = sum + Int(four) + Int(five)
End If
Label34.Caption = sum
Else
Label34.Caption = "-"
Playsound "doh"
End If
'Label34.Caption = sum
End If
End If
calculate
done
End Sub


Private Sub Label16_Click()
If klicked = False Then
If Not Label16.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label16.FontBold = True
Label35.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(3).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(5).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(7).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(9).Image = 6 Then sum = sum + 6
If sum >= 18 Then sum = 18: GoTo check
If sum < 18 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(3).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(5).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(7).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(9).Image = 5 Then sum = sum + 5
If sum >= 15 Then sum = 15: GoTo check
If sum < 15 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(3).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(5).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(7).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(9).Image = 4 Then sum = sum + 4
If sum >= 12 Then sum = 12: GoTo check
If sum < 12 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(3).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(5).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(7).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(9).Image = 3 Then sum = sum + 3
If sum >= 9 Then sum = 9: GoTo check
If sum < 9 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(3).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(5).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(7).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(9).Image = 2 Then sum = sum + 2
If sum >= 6 Then sum = 6: GoTo check
If sum < 6 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(3).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(5).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(7).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(9).Image = 1 Then sum = sum + 1
If sum >= 3 Then sum = 3: GoTo check
If sum < 3 Then sum = 0
check:
If sum > 0 Then
Label35.Caption = sum
Else
Label35.Caption = "-"
Playsound "doh"
End If
End If
End If
calculate
done


End Sub

Private Sub Label17_Click()
If klicked = False Then
If Not Label17.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label17.FontBold = True
Label36.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(3).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(5).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(7).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(9).Image = 6 Then sum = sum + 6
If sum >= 24 Then sum = 24: GoTo check
If sum < 24 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(3).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(5).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(7).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(9).Image = 5 Then sum = sum + 5
If sum >= 20 Then sum = 20: GoTo check
If sum < 20 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(3).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(5).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(7).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(9).Image = 4 Then sum = sum + 4
If sum >= 16 Then sum = 16: GoTo check
If sum < 16 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(3).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(5).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(7).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(9).Image = 3 Then sum = sum + 3
If sum >= 12 Then sum = 12: GoTo check
If sum < 12 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(3).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(5).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(7).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(9).Image = 2 Then sum = sum + 2
If sum >= 8 Then sum = 8: GoTo check
If sum < 8 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(3).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(5).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(7).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(9).Image = 1 Then sum = sum + 1
If sum >= 4 Then sum = 4: GoTo check
If sum < 4 Then sum = 0
check:
If sum > 0 Then
Label36.Caption = sum
Else
Label36.Caption = "-"
Playsound "doh"
End If
End If
End If
calculate
done
End Sub

Private Sub Label18_Click()
If klicked = False Then
If Not Label18.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label18.FontBold = True
Label37.FontBold = True
Label37.Caption = "-"
If Int(one) + Int(two) + Int(three) + Int(four) + Int(five) = 15 Then
    If Int(one) = Int(two) Or Int(one) = Int(three) Or Int(one) _
    = Int(four) Or Int(one) = Int(five) Or Int(two) = Int(three) _
    Or Int(two) = Int(four) Or Int(two) = Int(five) Or Int(three) _
    = Int(four) Or Int(three) = Int(five) Or Int(four) = Int(five) Then
    Label37.Caption = "-"
    'Playsound "doh"
    Else
    Label37.Caption = "15"
    End If
End If
If Not Int(one) + Int(two) + Int(three) + Int(four) + Int(five) = 15 Then Playsound "doh"

End If
End If
done
calculate
End Sub

Private Sub Label19_Click()
If klicked = False Then
If Not Label19.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label19.FontBold = True
Label38.FontBold = True
Label38.Caption = "-"
If Int(one) + Int(two) + Int(three) + Int(four) + Int(five) = 20 Then
    If Int(one) = Int(two) Or Int(one) = Int(three) Or Int(one) _
    = Int(four) Or Int(one) = Int(five) Or Int(two) = Int(three) _
    Or Int(two) = Int(four) Or Int(two) = Int(five) Or Int(three) _
    = Int(four) Or Int(three) = Int(five) Or Int(four) = Int(five) Then
    Label38.Caption = "-"
    'Playsound "doh"
    Else
    Label38.Caption = "20"
    End If
End If
If Not Int(one) + Int(two) + Int(three) + Int(four) + Int(five) = 20 Then Playsound "doh"
End If
End If
calculate
done
End Sub

Private Sub Label20_Click()
If klicked = False Then
If Not Label20.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label20.FontBold = True
Label39.FontBold = True
If Int(one) = Int(two) Then
chk = chk + 1
End If
If Int(one) = Int(three) Then
chk = chk + 1
End If
If Int(one) = Int(four) Then
chk = chk + 1
End If
If Int(one) = Int(five) Then
chk = chk + 1
End If
If Int(two) = Int(three) Then
chk = chk + 1
End If
If Int(two) = Int(four) Then
chk = chk + 1
End If
If Int(two) = Int(five) Then
chk = chk + 1
End If
If Int(three) = Int(four) Then
chk = chk + 1
End If
If Int(three) = Int(five) Then
chk = chk + 1
End If
If Int(four) = Int(five) Then
chk = chk + 1
End If
If chk = 4 Then
    sum = Int(one) + Int(two) + Int(three) + Int(four) + Int(five)
    Label39.Caption = sum
    Else
    Label39.Caption = "-"
    Playsound "doh"
End If
End If
End If
calculate
done
End Sub

Private Sub Label21_Click()
If klicked = False Then
If Not Label21.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label21.FontBold = True
Label40.FontBold = True
sum = Int(one) + Int(two) + Int(three) + Int(four) + Int(five)
Label40.Caption = sum

End If
End If
calculate
done
End Sub

Private Sub Label22_Click()
If klicked = False Then
If Not Label22.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
Label22.FontBold = True
Label41.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(3).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(5).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(7).Image = 6 Then sum = sum + 6
If Toolbar1.Buttons.Item(9).Image = 6 Then sum = sum + 6
If sum >= 30 Then sum = 30: GoTo check
If sum < 30 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(3).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(5).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(7).Image = 5 Then sum = sum + 5
If Toolbar1.Buttons.Item(9).Image = 5 Then sum = sum + 5
If sum >= 25 Then sum = 25: GoTo check
If sum < 25 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(3).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(5).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(7).Image = 4 Then sum = sum + 4
If Toolbar1.Buttons.Item(9).Image = 4 Then sum = sum + 4
If sum >= 20 Then sum = 20: GoTo check
If sum < 20 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(3).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(5).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(7).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(9).Image = 3 Then sum = sum + 3
If sum >= 15 Then sum = 15: GoTo check
If sum < 15 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(3).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(5).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(7).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(9).Image = 2 Then sum = sum + 2
If sum >= 10 Then sum = 10: GoTo check
If sum < 10 Then sum = 0
If Toolbar1.Buttons.Item(1).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(3).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(5).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(7).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(9).Image = 1 Then sum = sum + 1
If sum >= 5 Then sum = 5: GoTo check
If sum < 5 Then sum = 0
check:
If sum > 0 Then
Label41.Caption = "50"
Else
Label41.Caption = "-"
Playsound "doh"
End If
End If
End If
calculate
done
End Sub

Private Sub Label7_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label7.FontBold Then
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True
If Toolbar1.Buttons.Item(1).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(3).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(5).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(7).Image = 1 Then sum = sum + 1
If Toolbar1.Buttons.Item(9).Image = 1 Then sum = sum + 1
Label26.FontBold = True
Label7.FontBold = True
If sum = 0 Then
Label26.Caption = "-"
Playsound "doh"
Else
Label26.Caption = sum
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2
'MsgBox temp2
End If
End If
done
calculate
End Sub

Private Sub Label8_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label8.FontBold Then
klicked = True
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
Label8.FontBold = True
Label27.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(3).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(5).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(7).Image = 2 Then sum = sum + 2
If Toolbar1.Buttons.Item(9).Image = 2 Then sum = sum + 2
If sum > 0 Then
Label27.Caption = sum
Else
Label27.Caption = "-"
Playsound "doh"
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2

'MsgBox temp2
End If
End If
calculate
done
End Sub

Private Sub Label9_Click()
On Error Resume Next
Dim sum As Integer, temp As Integer, temp2 As Integer
If klicked = False Then
If Not Label9.FontBold Then
klicked = True
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
Label9.FontBold = True
Label28.FontBold = True
If Toolbar1.Buttons.Item(1).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(3).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(5).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(7).Image = 3 Then sum = sum + 3
If Toolbar1.Buttons.Item(9).Image = 3 Then sum = sum + 3
If sum > 0 Then
Label28.Caption = sum
Else
Label28.Caption = "-"
Playsound "doh"
End If
total = total + sum
temp = Label26.Caption
temp2 = temp2 + temp
temp = Label27.Caption
temp2 = temp2 + temp
temp = Label28.Caption
temp2 = temp2 + temp
temp = Label29.Caption
temp2 = temp2 + temp
temp = Label30.Caption
temp2 = temp2 + temp
temp = Label31.Caption
temp2 = temp2 + temp

Label42.Caption = temp2
If temp2 > 62 Then Label32.Caption = 50
temp = Label32.Caption
temp2 = temp2 + temp
Label43.Caption = temp2
End If
End If
calculate
done
End Sub

Private Sub Timer1_Timer()
'done
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Tag = 1 Then
    If hold1 = True Then
    hold1 = False
    Toolbar1.Buttons.Item(1).Value = tbrUnpressed
    Else
    hold1 = True
    Toolbar1.Buttons.Item(1).Value = tbrPressed
    End If
End If

If Button.Tag = 2 Then
    If hold2 = True Then
    hold2 = False
    Toolbar1.Buttons.Item(3).Value = tbrUnpressed
    Else
    hold2 = True
    Toolbar1.Buttons.Item(3).Value = tbrPressed
    End If
End If

If Button.Tag = 3 Then
    If hold3 = True Then
    hold3 = False
    Toolbar1.Buttons.Item(5).Value = tbrUnpressed
    Else
    hold3 = True
    Toolbar1.Buttons.Item(5).Value = tbrPressed
    End If
End If

If Button.Tag = 4 Then
    If hold4 = True Then
    hold4 = False
    Toolbar1.Buttons.Item(7).Value = tbrUnpressed
    Else
    hold4 = True
    Toolbar1.Buttons.Item(7).Value = tbrPressed
    End If
End If

If Button.Tag = 5 Then
    If hold5 = True Then
    hold5 = False
    Toolbar1.Buttons.Item(9).Value = tbrUnpressed
    Else
    hold5 = True
    Toolbar1.Buttons.Item(9).Value = tbrPressed
    End If
End If

End Sub
Private Sub calculate()
'Dim tmp As Integer, tmp2 As Integer
Dim tmp2 As Integer

tmp = 0

tmp2 = 0
tmp = Label43.Caption
tmp2 = tmp2 + tmp
tmp = Label33.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label34.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label35.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label36.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label37.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label38.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label39.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label40.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
tmp = Label41.Caption
If tmp = "-" Then tmp = 0
tmp2 = tmp2 + tmp
Label44.Caption = tmp2
Label44.Refresh
End Sub
Private Sub done()
If Label7.FontBold = False Then Exit Sub
If Label8.FontBold = False Then Exit Sub
If Label9.FontBold = False Then Exit Sub
If Label10.FontBold = False Then Exit Sub
If Label11.FontBold = False Then Exit Sub
If Label12.FontBold = False Then Exit Sub
If Label14.FontBold = False Then Exit Sub
If Label15.FontBold = False Then Exit Sub
If Label16.FontBold = False Then Exit Sub
If Label17.FontBold = False Then Exit Sub
If Label18.FontBold = False Then Exit Sub
If Label19.FontBold = False Then Exit Sub
If Label20.FontBold = False Then Exit Sub
If Label21.FontBold = False Then Exit Sub
If Label22.FontBold = False Then Exit Sub
Timer1.Enabled = False
test = MsgBox("All done. You've got " & Label44.Caption & " points." & vbCrLf & "Do you want to play again?", vbOKCancel, "Game over")
If test = 2 Then
End
Else
Timer1.Enabled = True
reinitialize
End If
End Sub
Private Sub reinitialize()
Label7.FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False

Label14.FontBold = False
Label15.FontBold = False
Label16.FontBold = False
Label17.FontBold = False
Label18.FontBold = False
Label19.FontBold = False
Label20.FontBold = False
Label21.FontBold = False
Label22.FontBold = False

Label26.FontBold = False
Label27.FontBold = False
Label28.FontBold = False
Label29.FontBold = False
Label30.FontBold = False
Label31.FontBold = False

Label33.FontBold = False
Label34.FontBold = False
Label35.FontBold = False
Label36.FontBold = False
Label37.FontBold = False
Label38.FontBold = False
Label39.FontBold = False
Label40.FontBold = False
Label41.FontBold = False

Label26.Caption = "0"
Label27.Caption = "0"
Label28.Caption = "0"
Label29.Caption = "0"
Label30.Caption = "0"
Label31.Caption = "0"
Label32.Caption = "0"
Label33.Caption = "0"
Label34.Caption = "0"
Label35.Caption = "0"
Label36.Caption = "0"
Label37.Caption = "0"
Label38.Caption = "0"
Label39.Caption = "0"
Label40.Caption = "0"
Label41.Caption = "0"
Label42.Caption = "0"
Label43.Caption = "0"
Label44.Caption = "0"
Command1.Enabled = True
rolls = 0
hold1 = False
hold2 = False
hold3 = False
hold4 = False
hold5 = False
Toolbar1.Buttons.Item(1).Value = tbrUnpressed
Toolbar1.Buttons.Item(3).Value = tbrUnpressed
Toolbar1.Buttons.Item(5).Value = tbrUnpressed
Toolbar1.Buttons.Item(7).Value = tbrUnpressed
Toolbar1.Buttons.Item(9).Value = tbrUnpressed
Toolbar1.Enabled = False
klicked = True

End Sub
