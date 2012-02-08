Attribute VB_Name = "modVarType"
'Variable structure for user
Public Type USER_INFO
    USER_PK As Long
    USER_NAME As String
    USER_ISADMIN As Boolean
    USER_ISMANAGER As Boolean
    USER_ISCASHIER As Boolean
    User_DSN As String
End Type

'Enumerator for form state
Public Enum FormState
    adStateAddMode = 0
    adStateEditMode = 1
    adStatePopupMode = 2
End Enum

Public Type BUSINESS_INFO
    BUSINNES_NAME As String
    BUSINESS_ADDRESS As String
    BUSINESS_CONTACT_INFO As String
    BUSINNES_CITY As String
    BUSINNES_NOTE As String
    BUSINNES_NPWP As String
    BUSINNES_GROUP As String
    BUSINNES_BANK As Integer
    BUSINNES_NEW As Byte
    BUSINNES_RECEPT As Byte
    BUSINNES_SALE As Byte
    BUSINNES_PLAFON As Double
End Type

Public Type TABLE_INFO
    TABLE_NO_FAK As String
    TABLE_GROUP As String
    TABLE_PAYMENT As Double
    TABLE_ID_PASIEN As Integer
    TABLE_KD_PASIEN As String
    TABLE_NM_PASIEN As String
    TABLE_UMUR_PASIEN As String
    TABLE_TLP_PASIEN As String
    TABLE_TOTAL_PASIEN As String
    TABLE_RELASI As String
    TABLE_TOTAL_OBAT As Double
    TABLE_ID_DEPT As String
    TABLE_KD_DEPT As String
    TABLE_NM_DEPT As String
    TABLE_ID_KREDITUR As String
    TABLE_NM_KREDITUR As String
    TABLE_ALMT_KREDITUR As String
    TABLE_KOTA_KREDITUR As String
    TABLE_CP_KREDITOR As String
    TABLE_TLP_KREDITOR As String
    TABLE_TANGGAL As String
    TABLE_ID_OBAT As String
    TABLE_KD_OBAT As String
    TABLE_NM_OBAT As String
    TABLE_RETUR_OBAT As String
    TABLE_SISA_OBAT As String
    TABLE_SISA_RETUR As String
    TABLE_TANGGAL_AWAL As String
    TABLE_TANGGAL_AKHIR As String
    TABLE_TYPE As String
    TABLE_PAY_TYPE As String
    TABLE_TOTAL As Double
    TABLE_KOMISI As Double
    TABLE_MONEY As Double
    TABLE_CBACK As Double
    TABLE_FLAG_OPNAME As Byte
    TABLE_FLAG_DEPT As String
    TABLE_FLAG_KREDITOR As String
    TABLE_FLAG_SUPPLIER As String
    TABLE_ID_SUPPLIER As String
    TABLE_NM_SUPPLIER As String
    TABLE_TLP_SUPPLIER As String
    TABLE_ALMT_SUPPLIER As String
    TABLE_KOTA_SUPPLIER As String
    TABLE_SEARCH As Byte
    TABLE_SEARCH2 As Byte
    TABLE_SEARCH3 As Byte
    TABLE_SEARCH_KREDITOR As String
    TABLE_SEARCH_SUPPLIER As String
    TABLE_SEARCH_DEP As String
    TABLE_SEARCH_FLAG As String
    TABLE_SEARCH_FLAG_2 As String
    TABLE_SEARCH_FLAG_3 As String
    TABLE_SEARCH_TANGGAL As String
    TABLE_SEARCH_TANGGAL_2 As String
    TABLE_SEARCH_TANGGAL_3 As String
    TABLE_BN As Double
    TABLE_AN As Double
    TABLE_PN As Double
    TABLE_VN As Double
    TABLE_TGL_CASH As String
    TABLE_LABA_BERSIH As String
    TABLE_TRANSFER As String
    TABLE_KAS_SISA As String
End Type

Public gsngXpos As Single
Public gsngYpos As Single

'public array to pass data values and labels to frmZoomChart
Public gsngZoomData() As Variant

'public array to pass series labels to frmZoomChart
Public gsSeriesLabels() As String


Public Sub NumberOnly(ByRef KeyAscii As Integer)
       If ((KeyAscii < 48 And KeyAscii <> 8 And KeyAscii <> 13) Or KeyAscii > 57) Then
           KeyAscii = 0
       End If
End Sub
