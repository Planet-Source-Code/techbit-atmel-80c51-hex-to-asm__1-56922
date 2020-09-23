VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATMEL 8051 - HEX To ASM Converter | Autor:_Leandro Gastón Vacirca - RAC Coder as Agent Neo"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12300
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstHEXCode 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "clmA"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "clmB"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "clmC"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "clmD"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "clmE"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "clmF"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "clmH"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "clmI"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "clmJ"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "clmK"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "clmL"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   688
      EndProperty
   End
   Begin MSComctlLib.ListView lstASMCode 
      Height          =   5055
      Left            =   5445
      TabIndex        =   3
      Top             =   840
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Mnemonic"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Operands"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Comments"
         Object.Width           =   3836
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "OPCode"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "NOB"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Oscillator Period"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6405
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ATMEL_80C51_HexToAsm.XPButton cmdConvert 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   661
      Caption         =   "Convert HEX To ASM"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      X1              =   -8
      X2              =   824
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open File"
      End
      Begin VB.Menu mnuSplitLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------
'                            ATMEL 80C51 HEX TO ASM CONVERTER
' ----------------------------------------------------------------------------------

' This code has been implemented to convert
' a 80C51 hexadecimal code to a 80C51 assembler code.
' There is a file inside the project folder, that
' contains the instruction set of the uP, saved as binary format.
'
' Please if you found this program useful think in my expend time
' in it for you and think if I really deserve a vote.
' If you vote me I will still writing more source code for the "programming world".

'           |---------------------------------------------------------|
'           |------------>>> ¡ PLEASE VOTE THIS CODE ! <<<------------|
'           |---------------------------------------------------------|

' ----------------------------------------------------------------------------------
'                                 ADDITIONAL INFORMATION
' ----------------------------------------------------------------------------------
' CODER:                       <<<LEANDRO GASTÓN VACIRCA>>>
'
'           _AGENT NEO in Rent a Coder
'           _LEANDRO V. in Planet Source Code
'
' - lgvacirca@hotmail.com
' - lgvacirca@yahoo.com.ar
' ----------------------------------------------------------------------------------


' ----------------------------------------------------------------------------------
'                              GUI FUNCTIONS & COMMANDS
' ----------------------------------------------------------------------------------

' This Sub will open a file
Private Sub mnuOpenFile_Click()
On Error GoTo Solution
    With cd1
        ' Prepare the Commmon Dialog
        .CancelError = True
        .Filter = "HEX Files (*.hex)|*.hex"
        .Flags = cdlOFNFileMustExist
    
        ' Show Common Dialog
        .ShowOpen
        
        ' Save FileName
        HexFileName = .FileName
        
        ' GUI & Control Commands
        lstHEXCode.ListItems.Clear
        lstASMCode.ListItems.Clear
        StatusBar.SimpleText = .FileName
        Call ConvertHexFormat(.FileName): Call LoadHEX(lstHEXCode)
        Call Sleep(1000): cmdConvert.Enabled = True
    End With
    ' Exit here
    Exit Sub
Solution:
    If Err.Number <> cdlCancel Then
        MsgBox "There was an error while opening the file.", vbCritical
    End If
End Sub

' Call to Convert Function from HEX to ASM
Private Sub cmdConvert_Click()
    Dim ASMFileName As String
    
    ' Check if the file has been loaded
    If HexFileName <> "" Then
        ' Allows to the user select the destination to the ASM File
        ASMFileName = Examined(Me, "Select the Destination Folder to save the ASM File Code")
        ' Check if the user cancel
        If ASMFileName = "" Then Exit Sub
        ' Call to the Function that converts the Hex to Asm Code
        cmdConvert.Enabled = False: Call ConvertHexToAsm(lstHEXCode, lstASMCode, ASMFileName)
    Else
        ' Indicating the error
        Call MsgBox("There isn't loaded any file." & vbCr & "You should load a file before to make a click over this button.", vbCritical)
    End If
End Sub

' Exit here
Private Sub mnuExit_Click()
    Unload frmMain
    End
End Sub





