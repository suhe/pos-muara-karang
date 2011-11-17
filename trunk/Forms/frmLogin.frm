VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Login"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   4740
      TabIndex        =   7
      Top             =   750
      Width           =   4740
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log-in"
      Default         =   -1  'True
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1110
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -600
      ScaleHeight     =   30
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   1800
      Width           =   6315
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   975
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   1350
      Width           =   1515
   End
   Begin MSDataListLib.DataCombo dcUser 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   975
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcGroup 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   1800
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Group"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your username and enter your password in the space provided bellow."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   750
      TabIndex        =   8
      Top             =   150
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   765
      Left            =   675
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   0
      Picture         =   "frmLogin.frx":038A
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   240
      Index           =   18
      Left            =   -300
      TabIndex        =   4
      Top             =   975
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    MDIMainMenu.CloseMe = True
    Unload Me
End Sub

Private Sub cmdLog_Click()
    'Verify
    Dim rsGroupBusinnes As New Recordset
    Dim rsCash As New Recordset
    Dim kas As Double
    
    If dcUser.Text = "" Then dcUser.SetFocus: Exit Sub
    If txtPass.Text = "" Then txtPass.SetFocus: Exit Sub
    If dcGroup.Enabled = True Then
        If dcGroup.Text = "" Then dcGroup.SetFocus: Exit Sub
    End If
    Dim strPass As String
    Dim total, admin, manager, user, cash As Byte
    total = getRecordCount("id", "tbl_pengguna", "WHERE nm_pengguna ='" & dcUser.Text & "' AND password='" & txtPass.Text & "'")
    admin = getRecordCount("id", "tbl_pengguna", "WHERE nm_pengguna ='" & dcUser.Text & "' AND password='" & txtPass.Text & "' AND level='Administrator' ")
    With CurrUser
        .USER_PK = dcUser.BoundText
        .USER_NAME = dcUser.Text
    End With
    If (dcGroup.Enabled = True) Then
        manager = getRecordCount("id", "tbl_pengguna", "WHERE nm_pengguna ='" & dcUser.Text & "' AND password='" & txtPass.Text & "' AND level='Manager' " & " And user_cabang = " & dcGroup.BoundText & "")
        user = getRecordCount("id", "tbl_pengguna", "WHERE nm_pengguna ='" & dcUser.Text & "' AND password='" & txtPass.Text & "' AND level='User'" & " And user_cabang = " & dcGroup.BoundText & " ")
        If rsGroupBusinnes.State = 1 Then rsGroupBusinnes.Close
        sql = "SELECT kd_cabang,plafon_default FROM tbl_cabang WHERE id_cabang=" & dcGroup.BoundText
        'MsgBox sql
        rsGroupBusinnes.Open sql, CN, adOpenStatic, adLockReadOnly
        If (rsGroupBusinnes.RecordCount > 0) Then
            'MsgBox rsGroupBusinnes.Fields("kd_cabang")
            tbl.TABLE_GROUP = Trim(UCase(rsGroupBusinnes.Fields("kd_cabang")))
            CurrBiz.BUSINNES_PLAFON = rsGroupBusinnes.Fields("plafon_default")
            rsGroupBusinnes.Close
        Else
            tbl.TABLE_GROUP = "AA"
        End If
    Else
        manager = 0
        tbl.TABLE_GROUP = "AA"
        user = 0
    End If
    
    If (total > 0) Then
        If (admin > 0) Then
            CurrUser.USER_ISADMIN = True
            CurrUser.USER_ISMANAGER = False
            CurrUser.USER_ISCASHIER = False
            CurrBiz.BUSINNES_GROUP = 0
             Unload Me
             LoadForm frmShortcuts
        ElseIf (manager > 0) Then
           CurrUser.USER_ISADMIN = False
            CurrUser.USER_ISMANAGER = True
            CurrUser.USER_ISCASHIER = False
            CurrBiz.BUSINNES_GROUP = dcGroup.BoundText
             Unload Me
             LoadForm frmShortcuts
        ElseIf (user > 0) Then
            CurrUser.USER_ISADMIN = False
            CurrUser.USER_ISMANAGER = False
            CurrUser.USER_ISCASHIER = True
            CurrBiz.BUSINNES_GROUP = dcGroup.BoundText
            Unload Me
            LoadForm frmShortcuts
        Else
            MsgBox "Invalid Group.Please try again!", vbExclamation
        End If
        
        With MDIMainMenu
            If CurrUser.USER_ISADMIN = False Then
                    .m_Master.Enabled = True
                    .m_transaction.Enabled = True
                    .m_trans_view.Enabled = True
                    .m_setting.Enabled = True
            ElseIf CurrUser.USER_ISMANAGER = True Then
                    .m_Master.Enabled = True
                    .m_transaction.Enabled = True
                    .m_trans_view.Enabled = True
                    .m_setting.Enabled = False
            ElseIf CurrUser.USER_ISCASHIER Then
                    .m_Master.Enabled = True
                    .m_transaction.Enabled = True
                    .m_cashier.Enabled = True
                    .m_purchasing.Enabled = False
                    .m_trans_view.Enabled = True
                    .m_setting.Enabled = False
            End If
        End With
        
        cash = getRecordCount("id", "tbl_cash", "WHERE tgl_cash ='" & Format(Now, "YYYY-MM-DD") & "' ")
        If (cash < 1) Then
            If rsCash.State = 1 Then rsCash.Close
            'sql = "SELECT kas_total FROM vw_cash_flow WHERE tgl_cash=DATE_ADD(CURDATE(), INTERVAL - 1 DAY) "
            sql = "SELECT (kas_total + (retur) ) as kas_total FROM vw_cash_flow ORDER BY tgl_cash DESC "
            rsCash.Open sql, CN, adOpenStatic, adLockReadOnly
            On Error Resume Next
            If (rsCash.RecordCount > 0) Then
                kas = Val(rsCash.Fields("kas_total"))
                rsCash.Close
            Else
                total = 0
            End If
            sql = "INSERT INTO tbl_cash(money_cash,tgl_cash,cash,tgl_input,id_pengguna) "
            sql = sql + "VALUES( "
            sql = sql + " " & kas & ", "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd") & "', "
            sql = sql + " 0, "
            sql = sql + " '" & Format(Now, "YYYY-mm-dd") & "', "
            sql = sql + " " & CurrUser.USER_PK & " "
            sql = sql + " ) "
            'MsgBox sql
            CN.Execute sql
        End If
        
        If (CurrBiz.BUSINNES_GROUP <> "") Then
            sql = "INSERT INTO tbl_log(tgl_akses,id_pengguna,id_cabang) "
            sql = sql + "VALUES("
            sql = sql + "'" & Format(Now, "YYYY-mm-dd hh:mm:ss") & "',"
            sql = sql + "" & CurrUser.USER_PK & ", "
            sql = sql + "" & CurrBiz.BUSINNES_GROUP & ""
            sql = sql + ")"
            'MsgBox sql
            CN.Execute sql
        End If
        'Dim root As Byte
        'root = getRecordCount("id_departement", "tbl_departement", "WHERE id_departement =0")
        'If (root < 1) Then
        '    sql = "INSERT INTO tbl_departement(id_departement,nm_departement,tgl_input,id_pengguna) "
        '    sql = sql + "VALUES( "
        '    sql = sql + " 0, "
        '    sql = sql + " 'Root Level', "
        '    sql = sql + " '" & Format(Now, "YYYY-mm-dd") & "', "
        '    sql = sql + " " & CurrUser.USER_PK & " "
        '    sql = sql + " ) "
        '    MsgBox sql
        '    CN.Execute sql
        'End If
    Else
        MsgBox "Invalid password.Please try again!", vbExclamation
        txtPass.SetFocus
    End If
    total = vbNullString
End Sub

Private Sub dcUser_Click(Area As Integer)
    On Error Resume Next
    Dim admin As Byte
    admin = getRecordCount("id", "tbl_pengguna", "WHERE nm_pengguna ='" & dcUser.Text & "' AND level='Administrator' ")
    If (admin) Then
        dcGroup.Enabled = False
    Else
        dcGroup.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM tbl_pengguna WHERE id<>0", "nm_pengguna", dcUser, "id"
    bind_dc "SELECT * FROM tbl_cabang", "nm_cabang", dcGroup, "id_cabang"
End Sub

Private Sub txtPass_Change()
    txtPass.SelStart = Len(txtPass.Text)
End Sub

Private Sub txtPass_GotFocus()
    HLText txtPass
End Sub
