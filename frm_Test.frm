VERSION 5.00
Object = "{688DDD0E-F569-4785-AB66-03885F2B828B}#1.0#0"; "NCommandMenuArray.ocx"
Begin VB.Form frm_Test 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin NCommandMenuArray.CommandMenuArray cma_CommandMenuArray 
      Align           =   1  'Align Top
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _extentx        =   8043
      _extenty        =   6376
   End
End
Attribute VB_Name = "frm_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cma_CommandMenuArray_AfterClicked(vCommandIndex As Integer, vCommandSender As Object, vLabelIndex As Integer, vLabelSender As Object)
'    MsgBox (vCommandIndex & " - " & vCommandSender.Caption & " : " & vLabelIndex & " - " & vLabelSender.Caption)
End Sub

Private Sub cma_CommandMenuArray_AfterOpened(vCommandMenuIndex As Integer, commandMenuSender As Object)
'    MsgBox (vCommandMenuIndex & " - " & commandMenuSender.Caption)
End Sub

Private Sub Form_Load()
    Dim cnu As Object
    Dim i As Integer
    For i = 0 To 100
        cma_CommandMenuArray.Add 0, "«”‰«œ Õ”«»œ«—Ì" & i, cnu
        cnu.Add "«ÌÃ«œ ”‰œ"
        cnu.Add "«’·«Õ ”‰œ"
        cnu.Add ""
        cnu.Add "Õ–› ”‰œ"
    Next i
    
    cma_CommandMenuArray.Add 1, "Test", cnu
    cnu.Add "test 01"
    cnu.Add "test 02"

    cma_CommandMenuArray.Add 2, "Testing " & 2, cnu
    cnu.Add "Testing  01"
    cnu.Add "Testing  02"
End Sub

Private Sub Form_Resize()
    cma_CommandMenuArray.Height = Height
End Sub
