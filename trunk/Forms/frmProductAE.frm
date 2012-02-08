VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProductAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicine AE"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbKemasan 
      Height          =   315
      Left            =   1560
      TabIndex        =   21
      Text            =   "cmbKemasan"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   19
      Tag             =   "Ilmiah"
      Top             =   1200
      Width           =   3915
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Tag             =   "Stock"
      Text            =   "0"
      Top             =   3960
      Width           =   795
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Tag             =   "Stock"
      Text            =   "0"
      Top             =   3600
      Width           =   795
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   99
      TabIndex        =   17
      Tag             =   "Code"
      Top             =   480
      Width           =   1080
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Tag             =   "Sell Price"
      Text            =   "0"
      Top             =   3195
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Tag             =   "Buy Price "
      Text            =   "0"
      Top             =   2820
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   825
      Width           =   3915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DCCategory 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "Category"
      Top             =   1575
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Ilmiah"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock Min"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Tag             =   "Unit"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock"
      Height          =   240
      Index           =   19
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   14
      Top             =   150
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Sell Price"
      Height          =   240
      Index           =   7
      Left            =   165
      TabIndex        =   13
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Buy Price"
      Height          =   240
      Index           =   6
      Left            =   165
      TabIndex        =   12
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pricing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   3
      Left            =   105
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Merk"
      Height          =   240
      Index           =   1
      Left            =   105
      TabIndex        =   9
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Code"
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   450
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2520
      Width           =   2865
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   150
      Width           =   3915
   End
End
Attribute VB_Name = "frmProductAE"
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
        txtEntry(0).Text = .Fields("kd_obat")
        txtEntry(1).Text = .Fields("nm_obat")
        DCCategory.BoundText = .Fields("id_kategori")
        DCCategory.Text = .Fields("nm_kategori")
        txtEntry(2).Text = .Fields("harga_beli")
        txtEntry(3).Text = .Fields("harga_jual")
        txtEntry(4).Text = .Fields("stok")
        txtEntry(5).Text = .Fields("stok_min")
        txtEntry(6).Text = .Fields("nm_ilmiah")
        cmbKemasan.Text = .Fields("kemasan")
        If (CurrUser.USER_ISADMIN = False) Then
            txtEntry(4).Enabled = False
            txtEntry(5).Enabled = False
        End If
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
    txtEntry(1).SetFocus
End Sub

Private Function getPK(ByVal srcTable As String) As String
    On Error GoTo err
    Dim rsKode As New Recordset
    Dim RI As Integer
    Dim pl, kode, rip As String
    rsKode.CursorLocation = adUseClient
    rsKode.Open "SELECT * FROM tbl_obat WHERE nm_obat<>'Bebas' AND pl_obat = '" & srcTable & "' ORDER BY ABS(pk_obat) DESC ", CN, adOpenStatic, adLockOptimistic
    If (rsKode.RecordCount > 0) Then
        pl = DCCategory.BoundText
        RI = Val(rsKode.Fields("pk_obat")) + 1
    Else
        pl = srcTable
        RI = 1
    End If
    
    If (pl <= 9) Then
        pl = "0" & pl
    Else
        pl = pl
    End If
    
    If (RI <= 9) Then
        rip = "000" & RI
    ElseIf (RI <= 99) Then
        rip = "00" & RI
    ElseIf (RI <= 999) Then
        rip = "0" & RI
    ElseIf (RI <= 9999) Then
        rip = RI
    End If
    
    'MsgBox RI
    
    getPK = pl & rip
    Set rsKode = Nothing
    Exit Function
err:
        If err.Number = 94 Then getPK = 1: Resume Next
End Function

Private Sub cmdSave_Click()
    'On Error Resume Next
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If is_empty(txtEntry(3), True) = True Then Exit Sub
    If is_empty(txtEntry(4), True) = True Then Exit Sub
    If is_empty(txtEntry(5), True) = True Then Exit Sub
    If is_empty(txtEntry(6), True) = True Then Exit Sub
    'If is_zero(txtEntry(2), True) = True Then Exit Sub
    'If is_zero(txtEntry(3), True) = True Then Exit Sub
    
    'MsgBox Val(txtEntry(2).Text)
    If State = adStateAddMode Or State = adStatePopupMode Then
        'With rs
        '    .AddNew
        '    .Fields("tgl_input") = Now
        '    .Fields("id_pengguna") = CurrUser.USER_PK
        '    .Fields("kd_obat") = txtEntry(0).Text
        '    .Fields("nm_obat") = txtEntry(1).Text
        '    .Fields("id_kategori") = DCCategory.BoundText
        '    .Fields("pk_obat") = Mid(txtEntry(0).Text, 2, 5)
        '    .Fields("pl_obat") = Left(txtEntry(0).Text, 1)
        '    .Fields("kemasan") = cmbKemasan.Text
        '    .Fields("harga_beli") = txtEntry(2).Text
        '    .Fields("harga_jual") = txtEntry(3).Text
        '    .Fields("stok") = txtEntry(4).Text
        '    .Fields("stok_min") = txtEntry(5).Text
        '    .Fields("nm_ilmiah") = txtEntry(6).Text
        '    .Update
            'MsgBox "Mantap"
            sql = "INSERT INTO tbl_obat(tgl_input,kd_obat,nm_obat,nm_ilmiah,id_kategori,pk_obat,pl_obat,kemasan,harga_beli,harga_jual,stok,stok_min,id_pengguna) "
            sql = sql + "VALUES( "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd h:m:s") & "', "
            sql = sql + " '" & txtEntry(0).Text & "', "
            sql = sql + " '" & txtEntry(1).Text & "', "
            sql = sql + " '" & txtEntry(6).Text & "', "
            sql = sql + " " & DCCategory.BoundText & ", "
            sql = sql + " '" & Mid(txtEntry(0).Text, 3, 5) & "', "
            sql = sql + " " & DCCategory.BoundText & ", "
            sql = sql + " '" & cmbKemasan.Text & "', "
            sql = sql + " " & txtEntry(2).Text & ", "
            sql = sql + " " & txtEntry(3).Text & ", "
            sql = sql + " " & txtEntry(4).Text & ", "
            sql = sql + " " & txtEntry(5).Text & ", "
            sql = sql + " " & CurrUser.USER_PK & " "
            sql = sql + " ) "
            'MsgBox sql
            CN.Execute sql
        'End With
    Else
            sql = "UPDATE tbl_obat "
            sql = sql + "SET "
            sql = sql + " nm_obat='" & txtEntry(1).Text & "', "
            sql = sql + " harga_beli=" & txtEntry(2).Text & ", "
            sql = sql + " harga_jual=" & txtEntry(3).Text & ", "
            sql = sql + " stok=" & txtEntry(4).Text & ", "
            sql = sql + " stok_min=" & txtEntry(5).Text & ", "
            sql = sql + " nm_ilmiah='" & txtEntry(6).Text & "', "
            sql = sql + " id_kategori=" & DCCategory.BoundText & ", "
            sql = sql + " kemasan='" & cmbKemasan.Text & "' "
            sql = sql + " WHERE id_obat=" & PK
            CN.Execute sql
    End If
    
    HaveAction = True
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            frmProduct.RefreshRecords
            ResetFields
            txtEntry(1).SetFocus
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub


Private Sub DCCategory_Change()
     On Error Resume Next
    'MsgBox "s"
    'If Index = 1 Then
      If State = adStateAddMode Or State = adStatePopupMode Then
        txtEntry(0).Text = getPK(DCCategory.BoundText)
        'txtEntry(0).Text = "xx"
      End If
    'End If
    
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    sql = "SELECT *,o.kemasan,o.stok,o.tgl_input,o.id_pengguna,o.id_kategori FROM tbl_obat o LEFT JOIN tbl_kategori k ON k.id_kategori=o.id_kategori WHERE o.id_obat = " & PK
    'MsgBox sql
    rs.Open sql, CN, adOpenStatic, adLockOptimistic, adCmdText
    bind_dc "SELECT * FROM tbl_kategori", "nm_kategori", DCCategory, "id_kategori", True
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        DCCategory.Enabled = False
        'MsgBox "Categori Pada Edit Menu "
    End If
    With cmbKemasan
        .AddItem "Tablet"
        .AddItem "Pot"
        .AddItem "Sirup"
        .AddItem "Tube"
    End With
    'MsgBox Format(Now, "YYYY-mm-dd h:m:s")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then frmProduct.RefreshRecords
        'MDIMainMenu.UpdateInfoMsg
    End If
    Set frmProductAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 3 And Index < 5 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub


