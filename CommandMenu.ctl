VERSION 5.00
Begin VB.UserControl CommandMenu 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "CommandMenu.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   2775
   Begin VB.CommandButton cmd_Menu 
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lbl_SubMenu 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "CommandMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Const CHeight = 375
Private gCount As Long
Private gChecked As Boolean
Private gCaption As String
Private gFilter As String

Public Event BeforeClicked(sender As Object, vIndex As Integer)
Public Event AfterClicked(sender As Object, vIndex As Integer)
Public Event BeforeOpened(sender As Object)
Public Event AfterOpened(sender As Object)

Public Property Let Caption(vCaption As String)
    gCaption = vCaption
    cmd_Menu.Caption = gCaption
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Caption = gCaption
End Property

Public Property Let Filter(vFilter As String)
    gFilter = vFilter
    resize
End Property
Public Property Get Filter() As String
    Filter = gFilter
End Property

Private Sub resize()
    Const mSepHGap = 125
    Const mSepHght = 50
    Dim mHeight As Long: mHeight = CHeight
    Dim lbl As Label
    With cmd_Menu
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = CHeight
        .Visible = True
    End With
    
    If Not gChecked Then
        Height = CHeight
        For Each lbl In lbl_SubMenu
            lbl.Visible = False
        Next
    Else
        For Each lbl In lbl_SubMenu
            With lbl
                If .Caption = "" Then
                    .Left = mSepHGap
                    .Height = mSepHght
                    .Width = Width - mSepHGap * 2
                    .Visible = True
                Else
                    If Trim(gFilter) <> "" Then
                        If InStr(1, LCase(Trim(.Caption)), LCase(Trim(gFilter))) = 0 Then
                            .Visible = False
                            .Height = 0
                        Else
                            .Left = 0
                            .Height = CHeight
                            .Width = Width
                            .Visible = True
                        End If
                    Else
                        .Left = 0
                        .Height = CHeight
                        .Width = Width
                        .Visible = True
                    End If
                End If
                .Top = mHeight
                mHeight = mHeight + .Height
            End With
        Next
        Height = mHeight
    End If
End Sub

Private Sub lbl_SubMenu_Click(vLabelIndex As Integer)
    RaiseEvent BeforeClicked(lbl_SubMenu(vLabelIndex), vLabelIndex)
    RaiseEvent AfterClicked(lbl_SubMenu(vLabelIndex), vLabelIndex)
End Sub

Private Sub lbl_Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lbl As Label
    For Each lbl In lbl_SubMenu
        With lbl
            If .Index = Index Then
                .FontUnderline = True
                .ForeColor = vbBlue
            Else
                .FontUnderline = False
                .ForeColor = vbBlack
            End If
        End With
    Next
End Sub

Private Sub UserControl_Initialize()
    gCount = 0
    gChecked = False
End Sub

Private Sub cmd_Menu_Click()
    RaiseEvent BeforeOpened(cmd_Menu)
    gChecked = Not gChecked
    resize
    RaiseEvent AfterOpened(cmd_Menu)
End Sub

Private Sub UserControl_Resize()
    resize
End Sub

Public Sub Add(vCaption As String)
    Dim lbl As Label
    If gCount = 0 Then
        Set lbl = lbl_SubMenu(gCount)
        gCount = 1
    Else
        gCount = lbl_SubMenu.count
        Load lbl_SubMenu(gCount)
        Set lbl = lbl_SubMenu(gCount)
    End If
    With lbl
        If Trim(vCaption) = "" Then
            .Appearance = 1
            .BackStyle = 1
            .BorderStyle = 1
            .Alignment = vbRightJustify
            .Caption = ""
            .ForeColor = vbWhite
            '.BackColor = vbWhite
            .RightToLeft = True
        Else
            .Appearance = 1
            .BackStyle = 1
            .Alignment = vbRightJustify
            .Caption = "  " & vCaption & "  "
            .ForeColor = vbBlack
            '.BackColor = vbWhite
            .RightToLeft = True
        End If
    End With
    resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", gCaption, ""
End Sub

Public Property Get checked() As Boolean
    checked = gChecked
End Property

Public Property Get count() As Integer
    count = gCount
End Property
