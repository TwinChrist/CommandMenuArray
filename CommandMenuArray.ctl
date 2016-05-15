VERSION 5.00
Begin VB.UserControl CommandMenuArray 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   PropertyPages   =   "CommandMenuArray.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   3465
   Begin VB.VScrollBar VScroll 
      Height          =   1215
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox pnl_Search 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   3465
      TabIndex        =   3
      Top             =   0
      Width           =   3465
      Begin VB.TextBox txt_SearchMenu 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_SearchSubMenu 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.PictureBox pnl_Buttons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   3105
      TabIndex        =   1
      Top             =   735
      Width           =   3105
      Begin NCommandMenuArray.CommandMenu cnu_commandMenu 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "CommandMenuArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Const CHeight = 375
Private gTop As Double
Private gCount As Long

Public Event BeforeClicked(vCommandIndex As Integer, vCommandSender As Object, vLabelIndex As Integer, vLabelSender As Object)
Public Event AfterClicked(vCommandIndex As Integer, vCommandSender As Object, vLabelIndex As Integer, vLabelSender As Object)
Public Event BeforeOpened(vCommandMenuIndex As Integer, vCommandMenuSender As Object)
Public Event AfterOpened(vCommandMenuIndex As Integer, vCommandMenuSender As Object)
Public Event BeforeSearch(vSearchSender As Object)
Public Event AfterSearch(vSearchSender As Object)

Public Property Let Caption(vIndex As Integer, vCaption As String)
    cnu_commandMenu(vIndex).Caption = vCaption
End Property
Public Property Get Caption(vIndex As Integer) As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "CommandMenuPPage"
    Caption = cnu_commandMenu(vIndex).Caption
End Property

Private Sub resize()
    Dim mTop As Long: mTop = 0
    Dim mHeight As Long: mHeight = 0
    Dim mFilter As String
    With txt_SearchMenu
        .Top = 0
        .Left = 0
        .Width = Width
        .Visible = True
        mTop = mTop + .Height + .Top
    End With
    With txt_SearchSubMenu
        .Top = mTop
        .Left = 0
        .Width = Width
        .Visible = True
        mTop = .Height + .Top
    End With
    With pnl_Search
        .Height = txt_SearchMenu.Height + txt_SearchSubMenu.Height
    End With
    mTop = gTop
    Dim i As Long
    For i = 0 To cnu_commandMenu.count - 1
        With cnu_commandMenu(i)
            mFilter = Trim(txt_SearchMenu.Text)
            .Filter = Trim(txt_SearchSubMenu.Text)
            If mFilter <> "" Then
                    If InStr(1, LCase(Trim(.Caption)), LCase(Trim(mFilter))) = 0 Then
                        .Visible = False
                        GoTo lout
                    End If
            End If
lFilter:
            .Top = mTop
            .Left = 0
            .Width = pnl_Buttons.Width
            .Visible = True
            mTop = .Height + .Top
            mHeight = mHeight + .Height

lout:
        End With
    Next
    
    mTop = pnl_Search.Height
    With pnl_Buttons
        .Top = mTop
        .Left = VScroll.Width
        .Width = Width - VScroll.Width
        .Height = mHeight
    End With
    With VScroll
        .Top = mTop
        .Left = 0
        If Height >= 2 * pnl_Search.Height Then
            .Height = Height + 165 - 2 * pnl_Search.Height
        End If
        .Min = 0
        If mHeight <= .Height Then
            .Max = 0
        Else
            .Max = mHeight / 2500
        End If
'        Debug.Print ".max " & .Max & " " & mHeight
    End With
End Sub

Private Sub chk_FilterMenus_Click()
    resize
End Sub

Private Sub chk_FilterMenu_Click()
    resize
End Sub

Private Sub chk_FilterSubMenu_Click()
    resize
End Sub

Private Sub cnu_commandMenu_AfterClicked(vCommandIndex As Integer, sender As Object, vLabelIndex As Integer)
    RaiseEvent BeforeClicked(vCommandIndex, cnu_commandMenu(vCommandIndex), vLabelIndex, sender)
    resize
    RaiseEvent AfterClicked(vCommandIndex, cnu_commandMenu(vCommandIndex), vLabelIndex, sender)
End Sub

Private Sub cnu_commandMenu_AfterOpened(Index As Integer, sender As Object)
    RaiseEvent BeforeOpened(Index, sender)
    resize
    RaiseEvent AfterOpened(Index, sender)
End Sub

Private Sub txt_SearchMenu_Change()
    RaiseEvent BeforeSearch(txt_SearchMenu)
    resize
    RaiseEvent AfterSearch(txt_SearchMenu)
End Sub

Private Sub txt_SearchSubMenu_Change()
    RaiseEvent BeforeSearch(txt_SearchMenu)
    resize
    RaiseEvent AfterSearch(txt_SearchMenu)
End Sub

Private Sub UserControl_Initialize()
    gCount = 0
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    resize
End Sub

Private Sub UserControl_Resize()
    resize
End Sub

Public Sub Add(vIndex As Integer, vCaption As String, vCommandMenu As CommandMenu)
    If gCount = 0 Then
        Set vCommandMenu = cnu_commandMenu(gCount)
        gCount = 1
    Else
        gCount = cnu_commandMenu.count
        Load cnu_commandMenu(gCount)
        Set vCommandMenu = cnu_commandMenu(gCount)
    End If
    vCommandMenu.Caption = vCaption
    resize
End Sub

Private Sub VScroll_Change()
    Dim mFactor As Double
    If VScroll.Value = 0 Then
        gTop = 0
    Else
        mFactor = VScroll.Max / VScroll.Value
        gTop = -(pnl_Buttons.Height / mFactor - (Height - 2 * pnl_Search.Height) / mFactor)
    End If
    pnl_Buttons.Top = gTop
    'Debug.Print gTop
    resize
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub
