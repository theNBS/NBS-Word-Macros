VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKeyNoteText 
   Caption         =   "Keynote Text"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   OleObjectBlob   =   "frmKeyNoteText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKeyNoteText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SetText(value As String)

    txtOutputText.Text = value
    
    With txtOutputText
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub



Private Sub cmdClose_Click()
    Hide
End Sub
