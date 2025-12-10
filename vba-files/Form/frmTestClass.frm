VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "UserForm1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10605
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsBI As clsButtonIcon
Attribute clsBI.VB_VarHelpID = -1

Private Sub chbEnabled_Click()
    clsBI.Enabled = chbEnabled.Value
End Sub

Private Sub chbHover_Click()
    clsBI.HoverOn = chbHover.Value
End Sub

Private Sub chbVisible_Click()
    clsBI.Visible = chbVisible.Value
End Sub

Private Sub clsBI_Click(mLabelOut As MSForms.Label, mLabelIn As MSForms.Label, Value As Boolean)
    lbValue.Caption = "Value: " & Value
End Sub

Private Sub cmbCodeIconIn_Change()
    clsBI.IconCodeInOff = cmbCodeIconIn.Value
End Sub

Private Sub cmbCodeIconInOn_Change()
    clsBI.IconCodeInOn = cmbCodeIconInOn.Value
End Sub

Private Sub optColor_1_change()
    clsBI.ColorOutOff = optColor_1.BackColor
End Sub

Private Sub optColor_2_Click()
    clsBI.ColorOutOff = optColor_2.BackColor
End Sub

Private Sub optColor_3_Click()
    clsBI.ColorOutOff = optColor_3.BackColor
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .left = Application.left + 0.5 * (Application.width - .width)
        .top = Application.top + 0.5 * (Application.height - .height)
    End With
    
    cmbCodeIconIn.addItem VBA.ChrW$(59962)
    cmbCodeIconIn.addItem VBA.ChrW$(59963)
    cmbCodeIconIn.addItem VBA.ChrW$(59145)
    cmbCodeIconIn.addItem VBA.ChrW$(60236)
    cmbCodeIconIn.addItem VBA.ChrW$(59642)
    cmbCodeIconIn.addItem VBA.ChrW$(60392)
    
    cmbCodeIconInOn.List = cmbCodeIconIn.List

    Set clsBI = New clsButtonIcon
    Call clsBI.Initialize(Label1, VBA.ChrW$(59962), _
            VBA.ChrW$(59963), _
            rgbAzure, _
            vbRed, _
            VBA.ChrW$(59145), _
            VBA.ChrW$(60236), _
            rgbOrchid, _
            rgbOrchid, _
            True)
            
    lbValue.Caption = "Value: " & clsBI.Value
    lbVersion.Caption = clsBI.Version(enAll)
End Sub