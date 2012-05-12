VERSION 5.00
Begin VB.Form frmStockOpnamePrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Opname"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   Icon            =   "frmStockOpnamePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Only Print"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Print && Remove History"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmStockOpnamePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    If Option1.Value = True Then
        Unload Me
        tbl.TABLE_FLAG_OPNAME = 0
        Call printStockOpname
    ElseIf Option2.Value = True Then
        Unload Me
        tbl.TABLE_FLAG_OPNAME = 1
        Call printStockOpname
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStockOpnamePrint = Nothing
End Sub
