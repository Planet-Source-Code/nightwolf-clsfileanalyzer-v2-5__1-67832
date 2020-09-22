VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Analyzer v2.5 by NightWolf"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDescription 
      Appearance      =   0  'Flat
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   105
      TabIndex        =   2
      Top             =   1260
      Width           =   6990
      Begin MSComCtl2.FlatScrollBar fsbScroll 
         Height          =   4220
         Left            =   6600
         TabIndex        =   4
         Top             =   300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   7435
         _Version        =   393216
         MousePointer    =   1
         LargeChange     =   50
         Max             =   100
         Orientation     =   1245184
         SmallChange     =   10
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   6500
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Analyzer v2.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   225
      Width           =   3255
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by NightWolf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2580
      TabIndex        =   1
      Top             =   675
      Width           =   1785
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   4320
      Picture         =   "frmMain.frx":058A
      Top             =   105
      Width           =   720
   End
   Begin VB.Shape shpBackground 
      BackStyle       =   1  'Opaque
      Height          =   1200
      Left            =   -15
      Top             =   -15
      Width           =   7230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Form (frmMain.frm)
'************************
'Created By: NightWolf
'Handles: Presentation of the clsFileAnalyzer Class description and functions
'-----------------------------------------------------------------------------------------

'Statements
'^^^^^^^^^^^
Option Explicit 'Force variable declaration

'APIs
'^^^^^
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'Sends a message to the system, in order to subclass a control

'Constants
'^^^^^^^^^^

'Scroll Textbox

Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINESCROLL = &HB6

'FORM
'*****

'Load
'^^^^^
Private Sub Form_Load()

Dim FA As clsFileAnalyzer: Set FA = New clsFileAnalyzer 'Create a File Analyzer instance

txtDescription.Text = FA.GetFileContents(FA.FixPath(App.Path) & "readme.txt") 'Read the contents of the 'readme.txt' file located in the application's path

Debug.Print FA.GetBaseName("C:\boot")

End Sub

'FLAT SCROLL BARS
'*****************

'Textbox Flat Scrollbar
'***********************

'Change
'^^^^^^^
Private Sub fsbScroll_Change()

Dim lngCurTopLine As Long, lngLinesToScroll As Long
    
lngCurTopLine = SendMessage(txtDescription.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0) 'Retrieve the index of the first visible line
    
lngLinesToScroll = ((fsbScroll.Value / 2) + (fsbScroll.Value / 4)) - lngCurTopLine 'Work out how many lines to scroll
    
Call SendMessage(txtDescription.hwnd, EM_LINESCROLL, 0&, ByVal lngLinesToScroll) 'Scroll the specified amount of lines

End Sub

'Scroll
'^^^^^^^
Private Sub fsbScroll_Scroll()

Call fsbScroll_Change 'Perform the same operation as in Change()

End Sub

'TEXTBOXES
'**********

'Change
'^^^^^^^
Private Sub txtDescription_Change()
'
Dim lngCount As Long, lngLine As Long

lngCount = SendMessage(txtDescription.hwnd, EM_GETLINECOUNT, 0, 0) 'Retrieve the amount of lines in the textbox

lngLine = SendMessage(txtDescription.hwnd, EM_LINEFROMCHAR, txtDescription.SelStart, 0) 'Retrieve the line the cursor is upon

If lngCount < 100 Then fsbScroll.Max = 100 Else fsbScroll.Max = lngCount 'Set maximum scroll value

fsbScroll.Value = lngLine 'Set scrollbar value

End Sub
