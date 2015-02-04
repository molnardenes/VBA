VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exportUserForm 
   Caption         =   "Export"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3570
   OleObjectBlob   =   "exportUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exportUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub exportButton_Click()
    Dim selectedSheetName As String
    selectedSheetName = exportUserForm.courseComboBox.SelText
    Call exportToIcs(Worksheets(selectedSheetName))
End Sub
