VERSION 5.00
Begin VB.Form frmCategoriesAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategoriesAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ctrlLiner1 
      Height          =   45
      Left            =   -525
      ScaleHeight     =   45
      ScaleWidth      =   5790
      TabIndex        =   6
      Top             =   2475
      Width           =   5790
   End
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   1
      Left            =   1575
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Tag             =   "Description"
      Top             =   525
      Width           =   3405
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1575
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Category Name"
      Top             =   150
      Width           =   3390
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3690
      TabIndex        =   3
      Top             =   2625
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Top             =   2625
      Width           =   1335
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   5
      Top             =   525
      Width           =   1290
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category Name"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmCategoriesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    With rs
        txtEntry(0).Text = .Fields("nm_kategori")
        txtEntry(1).Text = .Fields("desk_kategori")
    End With
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If State = adStateAddMode Then
        With rs
            .AddNew
            .Fields("tgl_input") = Now
            .Fields("id_pengguna") = CurrUser.USER_PK
            .Fields("nm_kategori") = txtEntry(0).Text
            .Fields("desk_kategori") = txtEntry(1).Text
            .Update
        End With
    Else
        sql = "UPDATE tbl_kategori "
        sql = sql + "SET "
        sql = sql + " nm_kategori='" & txtEntry(0).Text & "', "
        sql = sql + " desk_kategori='" & txtEntry(1).Text & "' "
        sql = sql + " WHERE id_kategori=" & PK
        CN.Execute sql
    End If
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            frmCategories.RefreshRecords
            ResetFields
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_kategori WHERE id_kategori = " & PK, CN, adOpenStatic, adLockOptimistic
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        'PK = getIndex("tbl_kategori")
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmCategories.RefreshRecords
        End If
    End If
    Set frmCategoriesAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub
