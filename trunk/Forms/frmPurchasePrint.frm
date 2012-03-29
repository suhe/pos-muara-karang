VERSION 5.00
Begin VB.Form frmPurchasePrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Purchase"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmPurchasePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton Option2 
         Caption         =   "Print Details Obat"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Print Summary Faktur"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.ComboBox cmbOperation 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmPurchasePrint.frx":038A
         Left            =   360
         List            =   "frmPurchasePrint.frx":0394
         TabIndex        =   1
         Text            =   "Ascending"
         Top             =   840
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmPurchasePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If is_empty(cmbOperation(1), True) = True Then Exit Sub
    If Option1.Value = True Then
        Unload Me
        Call printPurchaseSummary
    ElseIf Option2.Value = True Then
        Unload Me
        Call printPurchaseDetails
    Else
        MsgBox "Invalid Operation !", vbCritical + vbInformation
    End If
End Sub
