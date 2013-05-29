VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin Segunda.FImage FImage1 
      Height          =   1815
      Left            =   1080
      TabIndex        =   2
      Top             =   4560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
   End
   Begin VB.ListBox List2 
      Height          =   3375
      ItemData        =   "Form2.frx":0000
      Left            =   3600
      List            =   "Form2.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "Form2.frx":0004
      Left            =   720
      List            =   "Form2.frx":0006
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwmilliseconds As Long)
Dim i As Long
Dim II As Long
Dim YY As Long
Dim A As String
Dim idee
Dim B As Long
Dim E As String
Dim Aspen1 As String
Dim aspo As String
Dim aspen2 As String
Dim capa As Long
Dim capa2 As Long
Dim cap As Long
Private Sub Form_Load()
'On Error GoTo finales
If Dir(App.Path & "\Corte.txt") <> "" Then
    Open App.Path & "\Corte.txt" For Input As #2
    Do While Not EOF(2)
        Line Input #2, A
        List1.AddItem A
    Loop
    Close #2
End If
If Dir(App.Path & "\list2.txt") <> "" Then
    Open App.Path & "\list2.txt" For Input As #4
    Do While Not EOF(4)
        Line Input #4, Aspen1
        List2.AddItem Aspen1
    Loop
    Close #4
End If
If Dir(App.Path & "\Palo.txt") <> "" Then
    Open App.Path & "\Palo.txt" For Input As #3
        Line Input #3, E
    Close #3
End If
If Dir(App.Path & "\numero.txt") <> "" Then
    Open App.Path & "\numero.txt" For Input As #1
        Line Input #1, aspo
        i = Val(aspo)
    Close #1
End If
List1.ListIndex = E
B = List1.Text
For II = (i + 1) To B
    If Dir(App.Path & "\images\" & II & ".tif") <> "" Then
        FImage1.FILoad (App.Path & "\images\" & II & ".tif")
    ElseIf Dir(App.Path & "\images\Hoja-000" & II & ".tif") <> "" Then
        FImage1.FILoad (App.Path & "\images\Hoja-000" & II & ".tif")
    End If
    Printer.PaintPicture FImage1.Image, 0, 0, Printer.width, Printer.height
    If II < B Then
        Printer.NewPage
    End If
    aspo = Val(aspo) + 1
Next II
Open App.Path & "\numero.txt" For Output As #1
    Print #1, aspo
Close #1
If Dir(App.Path & "\Nombre.txt") <> "" Then
    Open App.Path & "\Nombre.txt" For Input As #5
        Line Input #5, aspen2
    Close #5
End If
Sleep 30000
If Dir("C:\Temp\Proyecto.pdf") <> "" Then
    List2.ListIndex = aspen2
    YY = 0
    Do Until YY = 1
          If FileLen("C:\Temp\Proyecto.pdf") > 0 Then
            YY = 1
            Exit Do
          End If
    Loop
    List2.ListIndex = aspen2
    FileCopy "C:\Temp\Proyecto.pdf", (App.Path & "\unidos\" & List2.Text)
    Kill "C:\Temp\Proyecto.pdf"
End If
aspen2 = Val(aspen2) + 1
Open App.Path & "\Nombre.txt" For Output As #5
    Print #5, aspen2
Close #5
E = E + 1
Open App.Path & "\Palo.txt" For Output As #3
    Print #3, E
Close #3
idee = ShellExecute(Me.hwnd, "Open", App.Path & "\Primera.exe", "", "c:\windows", 1)
End
'Exit Sub
'finales:
'    Open App.Path & "\Dondeaborto.txt" For Output As #1
'        Print #1, aspo
'    Close #1
'    MsgBox "Aborto el programa"
End Sub


