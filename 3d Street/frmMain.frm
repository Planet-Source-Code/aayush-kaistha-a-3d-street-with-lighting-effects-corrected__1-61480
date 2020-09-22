VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################################
'
'       Walk through an exotic 3d locality
'
'       Written By: Aayush Kaistha
'       Place:      UIET, Panjab University, Chandigarh
'       Contact:    aayushk_007@yahoo.com
'
'   Special thanx 2 Jack Hoxley (externalweb.exhedra.com/directx4vb)
'   for his gr8 tutorials
'
'########################################################

'3d objects were created using 3d studio max and exported
'in .3ds format which were then converted into .x format
'using the program conv3ds.exe available with direct x sdk

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = True
If KeyCode = vbKeyDown Then DownKey = True
If KeyCode = vbKeyLeft Then LeftKey = True
If KeyCode = vbKeyRight Then RightKey = True

If KeyCode = vbKeyS Then SKey = True
If KeyCode = vbKeyW Then WKey = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = False
If KeyCode = vbKeyDown Then DownKey = False
If KeyCode = vbKeyLeft Then LeftKey = False
If KeyCode = vbKeyRight Then RightKey = False
If KeyCode = vbKeyEscape Then bRunning = False

If KeyCode = vbKeyS Then SKey = False
If KeyCode = vbKeyW Then WKey = False

End Sub

Private Sub Form_Click()
bRunning = False
End Sub

