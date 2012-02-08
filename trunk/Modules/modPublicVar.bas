Attribute VB_Name = "modPublicVar"
Option Explicit


Public CurrUser                     As USER_INFO
Public DBPath                       As String
Public Enc                          As New clsBlowfish
Public CurrBiz                      As BUSINESS_INFO
Public tbl                          As TABLE_INFO
Public CN                           As New Connection
Public sql                          As String


