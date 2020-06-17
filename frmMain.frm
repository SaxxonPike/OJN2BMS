VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OJN2BMS v1.0"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Options and Actions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5655
      Begin VB.CheckBox chkErrors 
         Caption         =   "Don't display errors"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkFreeze 
         Caption         =   "Convert freeze notes (ZZ ends)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Convert"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton cmdBrowseFolder 
         Caption         =   "Browse..."
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "OJN file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdBrowseOJN 
         Caption         =   "Browse..."
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtOJN 
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Const IDLE_PRIORITY_CLASS = &H40

Private OJN As New clsOJNFile

Private Sub cmdBrowseFolder_Click()
    Dim FName As String
    CD_ShowFolder Me.hWnd, _
        BIF_EDITBOX + BIF_VALIDATE, _
        "", FName, "Select Destination Folder. Folders will be created within the target folder for each song."
    If FName <> "" Then
        txtTarget = FName
    End If
End Sub

Private Sub cmdBrowseOJN_Click()
    Dim FName As String
    CD_ShowOpen_Save Me.hWnd, _
        OFN_FILEMUSTEXIST + OFN_EXPLORER, _
        FName, "", "*.ojn", "Open O2Jam Music File", "O2Jam Chart Files (*.OJN)|*.OJN|All Files (*.*)|*.*", "", True
    If FName <> "" Then
        txtOJN = FName
    End If
End Sub

Private Sub Command1_Click()
    If OJN.LoadOJN(txtOJN, txtTarget, (chkFreeze.Value = 1), (chkErrors.Value = 0)) Then
        MsgBox "Finished"
    End If
End Sub

Private Sub Form_Load()
    SetPriorityClass GetCurrentProcess, IDLE_PRIORITY_CLASS
End Sub

Private Sub txtOJN_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim x2 As Long
    If Data.Files.Count = 1 Then
        txtOJN = Data.Files(1)
    Else
        If txtTarget = "" Then
            MsgBox "Before " + "attempting to convert multiple files" + ", please define a target folder."
            Exit Sub
        End If
        If MsgBox("You are " + "attempting to convert multiple files" + "." + vbCrLf + "Continue?", vbYesNo, "Batch Conversion") = vbYes Then
            Command1.Enabled = False
            MsgBox "Please note that this process may take a long time, so" + vbCrLf + "please be patient. Click OK to continue.", vbInformation
            For x2 = 1 To Data.Files.Count
                Command1.Caption = Data.Files(x2)
                DoEvents
                If OJN.LoadOJN(Data.Files(x2), txtTarget, (chkFreeze.Value = 1), (chkErrors.Value = 0)) = False Then
                    If chkErrors.Value = 0 Then
                        If MsgBox("There was an error while converting:" + vbCrLf + Data.Files(x2) + vbCrLf + vbCrLf + "Keep going?", vbYesNo) = vbNo Then
                            Exit For
                        End If
                    End If
                End If
            Next x2
            Command1.Enabled = True
            Command1.Caption = "Convert"
            MsgBox "Finished"
        End If
    End If
End Sub

