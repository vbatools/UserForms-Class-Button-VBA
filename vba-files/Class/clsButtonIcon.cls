VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButtonIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class: clsButtonIcon
' Author: VBATools
' Version: 1.0.1
' Creation Date: 10.12.2025 11:41
' Last Update Date: 10.12.2025 12:15
' License: Apache License
'========================================================================================

Private Const FONT_NAME_ICON As String = "Segoe MDL2 Assets"

Private WithEvents mLabelMain As MSForms.Label
Attribute mLabelMain.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1
Private mLabelOut   As MSForms.Label
Private mLabelIn    As MSForms.Label

Public Event Click(ByRef mLabelOut As MSForms.Label, ByRef mLabelIn As MSForms.Label, ByRef Value As Boolean)


Private mItems      As Collection
Private mName       As String

Private mIconCodeOutOn As String
Private mIconCodeOutOff As String
Private mIconCodeInOn As String
Private mIconCodeInOff As String

Private msnSizeOut  As Single
Private msnSizeIn   As Single

Private mColorInOff As XlRgbColor
Private mColorInOn  As XlRgbColor
Private mColorOutOff As XlRgbColor
Private mColorOutOn As XlRgbColor

Private mVisible    As Boolean
Private mEnabled    As Boolean
Private mValue      As Boolean
Private mHoverOn    As Boolean

Public Enum enumParametrVersion
    enName = 0
    enAuthor
    enVersion
    enLicense
    enDateOfCreation
    enDateOfUpdate
    enDescription
    enAll
    [_First] = enName
    [_Last] = enAll
End Enum

Public Function Version(ByVal Parametr As enumParametrVersion) As String
    Dim sRes As String
    Select Case Parametr
        Case enumParametrVersion.enAll:
            Dim i   As Byte
            For i = enumParametrVersion.[_First] To enumParametrVersion.[_Last] - 1
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                sRes = sRes & VersionItem(i)
            Next i
        Case Else:
            sRes = VersionItem(Parametr)
    End Select
    Version = sRes
End Function

Private Function VersionItem(ByVal Parametr As enumParametrVersion) As String
    Select Case Parametr
        Case enumParametrVersion.enName:
            VersionItem = "Class: clsButtonIcon"
        Case enumParametrVersion.enAuthor:
            VersionItem = "Author: VBATools"
        Case enumParametrVersion.enVersion:
            VersionItem = "Version: 1.0.1"
        Case enumParametrVersion.enLicense:
            VersionItem = "License: Apache License"
        Case enumParametrVersion.enDateOfCreation:
            VersionItem = "Date of creation: 10.12.2025 11:41"
        Case enumParametrVersion.enDateOfUpdate:
            VersionItem = "Date of update: 10.12.2025 12:15"
        Case enumParametrVersion.enDescription:
            VersionItem = "Description: Сlass that implements an interactive double icon button for use in Excel VBA UserForms. The class manages two Label controls (outer and inner), creating a unique visual element with the ability to change state and color on mouse hover."
    End Select
End Function

Public Property Get HoverOn() As Boolean
    Call errRaise
    HoverOn = mHoverOn
End Property

Public Property Let HoverOn(ByVal Value As Boolean)
    Call errRaise
    mHoverOn = Value
End Property


Public Property Get Name() As String
    Call errRaise
    Name = mName
End Property

Public Property Let Name(ByVal Name As String)
    Call errRaise
    mName = Name
End Property

Public Property Get IconCodeOutOn() As String
    Call errRaise
    IconCodeOutOn = mIconCodeOutOn
End Property

Public Property Let IconCodeOutOn(ByVal Code As String)
    Call errRaise
    If Len(Code) = 0 Then
        Call Err.Raise(Number:=vbObjectError + 105, Description:="IconCodeOutOn cannot be empty. Please provide a valid icon code.", Source:=TypeName(Me))
    End If
    mIconCodeOutOn = Code
    Call ChangeStaet(Me)
End Property

Public Property Get IconCodeOutOff() As String
    Call errRaise
    IconCodeOutOff = mIconCodeOutOff
End Property

Public Property Let IconCodeOutOff(ByVal Code As String)
    Call errRaise
    If Len(Code) = 0 Then
        Call Err.Raise(Number:=vbObjectError + 106, Description:="IconCodeOutOff cannot be empty. Please provide a valid icon code.", Source:=TypeName(Me))
    End If
    mIconCodeOutOff = Code
    Call ChangeStaet(Me)
End Property

Public Property Get IconCodeInOn() As String
    Call errRaise
    IconCodeInOn = mIconCodeInOn
End Property

Public Property Let IconCodeInOn(ByVal Code As String)
    Call errRaise
    If Len(Code) = 0 Then
        Call Err.Raise(Number:=vbObjectError + 107, Description:="IconCodeInOn cannot be empty. Please provide a valid icon code.", Source:=TypeName(Me))
    End If
    mIconCodeInOn = Code
    Call ChangeStaet(Me)
End Property

Public Property Get IconCodeInOff() As String
    Call errRaise
    IconCodeInOff = mIconCodeInOff
End Property

Public Property Let IconCodeInOff(ByVal Code As String)
    Call errRaise
    If Len(Code) = 0 Then
        Call Err.Raise(Number:=vbObjectError + 108, Description:="IconCodeInOff cannot be empty. Please provide a valid icon code.", Source:=TypeName(Me))
    End If
    mIconCodeInOff = Code
    Call ChangeStaet(Me)
End Property

Public Property Get SizeOut() As Single
    Call errRaise
    SizeOut = msnSizeOut
End Property

Public Property Let SizeOut(ByVal Size As Single)
    Call errRaise
    If Size <= 0 Then
        Call Err.Raise(Number:=vbObjectError + 109, Description:="SizeOut must be greater than 0. Current value: " & Size, Source:=TypeName(Me))
    End If
    msnSizeOut = Size
End Property

Public Property Get SizeIn() As Single
    Call errRaise
    SizeIn = msnSizeIn
End Property

Public Property Let SizeIn(ByVal Size As Single)
    Call errRaise
    If Size <= 0 Then
        Call Err.Raise(Number:=vbObjectError + 110, Description:="SizeIn must be greater than 0. Current value: " & Size, Source:=TypeName(Me))
    End If
    msnSizeIn = Size
End Property

Public Property Get LabelMain() As MSForms.Label
    Call errRaise
    Set LabelMain = mLabelMain
End Property

Public Property Set LabelMain(ByRef Label As MSForms.Label)
    Set mLabelMain = Label
End Property

Public Property Get LabelOut() As MSForms.Label
    Call errRaise
    Set LabelOut = mLabelOut
End Property

Public Property Set LabelOut(ByRef Label As MSForms.Label)
    Call errRaise
    Set mLabelOut = Label
End Property

Public Property Get LabelIn() As MSForms.Label
    Call errRaise
    Set LabelIn = mLabelIn
End Property

Public Property Set LabelIn(ByRef Label As MSForms.Label)
    Call errRaise
    Set mLabelIn = Label
End Property

Public Property Get ColorOutOn() As XlRgbColor
    Call errRaise
    ColorOutOn = mColorOutOn
End Property

Public Property Let ColorOutOn(ByVal Color As XlRgbColor)
    Call errRaise
    mColorOutOn = Color
End Property

Public Property Get ColorOutOff() As XlRgbColor
    Call errRaise
    ColorOutOff = mColorOutOff
End Property

Public Property Let ColorOutOff(ByVal Color As XlRgbColor)
    Call errRaise
    mColorOutOff = Color
End Property

Public Property Get ColorInOn() As XlRgbColor
    Call errRaise
    ColorInOn = mColorInOn
End Property

Public Property Let ColorInOn(ByVal Color As XlRgbColor)
    Call errRaise
    mColorInOn = Color
End Property

Public Property Get ColorInOff() As XlRgbColor
    Call errRaise
    ColorInOff = mColorInOff
End Property

Public Property Let ColorInOff(ByVal Color As XlRgbColor)
    Call errRaise
    mColorInOff = Color
End Property

Public Property Get Value() As Boolean
    Call errRaise
    Value = mValue
End Property

Public Property Let Value(ByVal Value As Boolean)
    Call errRaise
    mValue = Value
End Property

Public Property Get Visible() As Boolean
    Call errRaise
    Visible = mVisible
End Property

Public Property Let Visible(ByVal Visible As Boolean)
    Call errRaise
    mVisible = Visible
    mLabelMain.Visible = Visible
    mLabelOut.Visible = Visible
    mLabelIn.Visible = Visible
End Property

Public Property Get Enabled() As Boolean
    Call errRaise
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal Enabled As Boolean)
    Call errRaise
    mEnabled = Enabled
    mLabelMain.Enabled = Enabled
    mLabelOut.Enabled = Enabled
    mLabelIn.Enabled = Enabled
End Property

Private Sub errRaise()
    If mLabelMain Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 100, Description:="Button item is not set. Please initialize the button using the Initialize method before using its properties or methods.", Source:=TypeName(Me))
    End If
End Sub

Public Sub Initialize(ByRef Label As MSForms.Label, _
        ByVal OutIconCodeOn As String, _
        ByVal OutIconCodeOff As String, _
        ByVal OutColorOff As XlRgbColor, _
        ByVal OutColorOn As XlRgbColor, _
        ByVal InIconCodeOn As String, _
        ByVal InIconCodeOff As String, _
        ByVal InColorOff As XlRgbColor, _
        ByVal InColorOn As XlRgbColor, _
        Optional Value As Boolean = True, _
        Optional HoverOn As Boolean = True)
    
    ' Проверка корректности входных данных
    If Label Is Nothing Then
        Call Err.Raise(Number:=vbObjectError + 101, Description:="Label parameter cannot be Nothing. Please provide a valid MSForms.Label control.", Source:=TypeName(Me))
    End If
    
    If Label.height <= 0 Then
        Call Err.Raise(Number:=vbObjectError + 103, Description:="Label height must be greater than 0. Current height: " & Label.height, Source:=TypeName(Me))
    End If
    
    If Label.width <= 0 Then
        Call Err.Raise(Number:=vbObjectError + 104, Description:="Label width must be greater than 0. Current width: " & Label.width, Source:=TypeName(Me))
    End If
    
    ' Проверка, что шрифт иконки доступен
    ' Note: Проверка доступности шрифта может быть сложной в VBA, поэтому пропускаем эту проверку
    
    Set mLabelMain = Label
    Set mParent = Label.Parent
    mName = mLabelMain.Name
    msnSizeOut = Label.height
    msnSizeIn = msnSizeOut * 0.75
    mIconCodeOutOn = OutIconCodeOn
    mIconCodeOutOff = OutIconCodeOff
    mColorOutOn = OutColorOn
    mColorOutOff = OutColorOff
    mIconCodeInOn = InIconCodeOn
    mIconCodeInOff = InIconCodeOff
    mColorInOn = InColorOn
    mColorInOff = InColorOff
    mValue = Value
    mHoverOn = HoverOn

    With mLabelMain
        Set mLabelOut = .Parent.Controls.Add("Forms.Label.1", .Name & "_OUT")
        Set mLabelIn = .Parent.Controls.Add("Forms.Label.1", .Name & "_IN")
        .width = msnSizeOut
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .Caption = vbNullString
    End With
    With mLabelOut
        .top = mLabelMain.top
        .left = mLabelMain.left
        .width = msnSizeOut
        .height = msnSizeOut
        .Font.Size = msnSizeOut * 0.9
        .Font.Name = FONT_NAME_ICON
        .Caption = mIconCodeOutOn
        .TextAlign = fmTextAlignCenter
        .ForeColor = OutColorOn
        '.BorderStyle = fmBorderStyleSingle
    End With
    With mLabelIn
        .height = msnSizeIn
        .width = msnSizeIn
        .Font.Size = msnSizeIn * 0.75
        .Font.Name = FONT_NAME_ICON
        .Caption = mIconCodeInOn
        .TextAlign = fmTextAlignCenter
        .ForeColor = InColorOn
        .BackStyle = fmBackStyleTransparent
        '.BorderStyle = fmBorderStyleSingle

        .top = mLabelMain.top + (msnSizeOut - msnSizeIn) / 2 + 0.5
        .left = mLabelMain.left + (msnSizeOut - msnSizeIn) / 2
    End With

    mLabelMain.ZOrder 0
    mVisible = mLabelOut.Visible
    mEnabled = mLabelOut.Enabled
End Sub

Private Sub mLabelMain_Click()
    Call ChangeStaet(Me)
    RaiseEvent Click(mLabelOut, mLabelIn, mValue)
End Sub

Private Sub ChangeStaet(ByRef Item As clsButtonIcon)
    With Item
        With .LabelOut
            If .Caption <> Item.IconCodeOutOn Then
                .Caption = Item.IconCodeOutOn
            Else
                .Caption = Item.IconCodeOutOff
            End If
            If .ForeColor <> Item.ColorOutOn Then
                .ForeColor = Item.ColorOutOn
            Else
                .ForeColor = Item.ColorOutOff
            End If
        End With
        With Item.LabelIn
            If .Caption <> Item.IconCodeInOn Then
                .Caption = Item.IconCodeInOn
            Else
                .Caption = Item.IconCodeInOff
            End If
            If .ForeColor <> Item.ColorInOn Then
                .ForeColor = Item.ColorInOn
            Else
                .ForeColor = Item.ColorInOff
            End If
        End With
        .Value = Not .Value
    End With
End Sub

Private Sub mLabelMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not mHoverOn Then Exit Sub
    Call hoverColor(mLabelOut, mValue, mColorOutOn, mColorOutOff, 150)
    Call hoverColor(mLabelIn, mValue, mColorInOn, mColorInOff, 150)
End Sub

Private Sub mParent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not mHoverOn Then Exit Sub
    Call hoverColor(mLabelOut, mValue, mColorOutOn, mColorOutOff, 0)
    Call hoverColor(mLabelIn, mValue, mColorInOn, mColorInOff, 0)
End Sub

' Function to blend colors for hover effect
Private Function BlendColor(ByVal color1 As Long, color2 As Long, factor As Double) As Long
    Dim r1 As Byte, g1 As Byte, b1 As Byte
    Dim r2 As Byte, g2 As Byte, b2 As Byte
    Dim r As Byte, g As Byte, b As Byte

    ' Extract BGR components from the first color (VBA uses BGR format)
    r1 = color1 Mod 256
    g1 = (color1 \ 256) Mod 256
    b1 = (color1 \ 65536) Mod 256

    ' Extract BGR components from the second color (VBA uses BGR format)
    r2 = color2 Mod 256
    g2 = (color2 \ 256) Mod 256
    b2 = (color2 \ 65536) Mod 256

    ' Blend the colors
    r = r1 + (r2 - r1) * factor
    g = g1 + (g2 - g1) * factor
    b = b1 + (b2 - b1) * factor

    ' Return the blended color
    BlendColor = RGB(r, g, b)
End Function

Private Sub hoverColor(ByRef Label As MSForms.Label, ByVal Value As Boolean, ByVal ColorOn As XlRgbColor, ByVal ColorOff As XlRgbColor, ByVal shade As Byte)
    If Label Is Nothing Then Exit Sub
    With Label
        If Value Then
            .ForeColor = BlendColor(ColorOn, RGB(255, 255, 255), shade / 255)
        Else
            .ForeColor = BlendColor(ColorOff, RGB(255, 255, 255), shade / 255)
        End If
    End With
End Sub

Private Sub Class_Terminate()
    Set mLabelMain = Nothing
    Set mLabelOut = Nothing
    Set mLabelIn = Nothing
    Set mParent = Nothing
End Sub