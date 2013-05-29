VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Genera PDF Ver 2.0"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13545
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tiempos"
      Height          =   3975
      Left            =   5640
      TabIndex        =   11
      Top             =   2400
      Width           =   2415
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   855
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Calibrar Tiempos"
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   12120
      TabIndex        =   7
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continuar PDF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      ItemData        =   "Form1.frx":058A
      Left            =   8160
      List            =   "Form1.frx":058C
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generar PDF encadenado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin Primera.FImage FImage1 
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   6840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      ItemData        =   "Form1.frx":058E
      Left            =   3000
      List            =   "Form1.frx":0590
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      ItemData        =   "Form1.frx":0592
      Left            =   240
      List            =   "Form1.frx":0594
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   11760
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   11760
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lbl_Cantidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwmilliseconds As Long)

Sub CreaPDFCarpetas()
    For ZZ = 0 To List1.ListCount - 1
        List1.ListIndex = ZZ
        lbl_Cantidad.Caption = lbl_Cantidad.Caption + 1
        lbl_Cantidad.Refresh
        B = Val(List1.Text)
        Imprimetotal = 0
        For II = (Val(aspen2) + 1) To B
            Set objFSO = New Scripting.FileSystemObject
            If Dir(strImages) <> "" Then
                FImage1.FILoad (strImages)
                Set objFile = objFSO.GetFile(strImages)
            ElseIf Dir(strHojas) <> "" Then
                FImage1.FILoad (strHojas)
                Set objFile = objFSO.GetFile(strHojas)
            End If
            ImgPeso = ImgPeso + (objFile.size) / 1000
            Set objFile = Nothing
            Set objFSO = Nothing
            Printer.PaintPicture FImage1.Image, 0, 0, Printer.width, Printer.height
            If II < B Then Printer.NewPage
            aspen2 = Val(aspen2) + 1
            Imprimetotal = Imprimetotal + 1
        Next II
        Printer.EndDoc
        
        YY = 0
        Do Until YY = 1
            If Dir("C:\Temp\Proyecto.pdf") <> "" Then
                List2.ListIndex = ZZ
                YY = 0
                Do Until YY = 1
                      If FileLen("C:\Temp\Proyecto.pdf") > 20 Then
                        YY = 1
                        Exit Do
                      End If
                Loop
                FileCopy "C:\Temp\Proyecto.pdf", (App.Path & "\" & List3.Text & "\unidos\" & List2.Text)
            End If
        Loop
    Next ZZ
End Sub

Private Sub Command3_Click()
    Call CargarCarpeta(List3)
    If Existe_Carpetas <> False Then
        lbl_Cantidad.Caption = 0
        For ZZ1 = 0 To List3.ListCount - 1
            Call CargarCorte(List1, List3, lbl_Cantidad)
            Call BuscarRepetidos(List2)
            If Dir((strUnidos), vbDirectory) = "" Then MkDir strUnidos
            msg = MsgBox("¿Desea comenzar el proceso?", vbOKCancel, "Pdf Creador")
            If msg = vbOK Then Call CreaPDFCarpetas
        Next ZZ1
    End If
End Sub

Private Sub Command4_Click()
continuardesde = ""
continuardesde = InputBox("Ingresa el archivo anterior al que abortó", "Continua")
 If continuardesde <> "" Then
    'On Error GoTo finales
        i = 0
        aspen2 = 0
        lbl_Cantidad.Caption = "0"
    For ZZ = 0 To List1.ListCount - 1
        List1.ListIndex = ZZ
        List2.ListIndex = ZZ
        lbl_Cantidad.Caption = lbl_Cantidad.Caption + 1
        lbl_Cantidad.Refresh
     If Trim(continuardesde) <> Trim(List2.Text) And Trim(continuardesde) <> "Continuar" Then
                aspen2 = List1.Text
     ElseIf Trim(continuardesde) = Trim(List2.Text) Or Trim(continuardesde) = "Continuar" Then
        B = Val(List1.Text)
        Imprimetotal = 0
        For II = (Val(aspen2) + 1) To B
            If Dir(App.Path & "\images\" & II & ".tif") <> "" Then
                FImage1.FILoad (App.Path & "\images\" & II & ".tif")
            ElseIf Dir(App.Path & "\images\Hoja-000" & II & ".tif") <> "" Then
                FImage1.FILoad (App.Path & "\images\Hoja-000" & II & ".tif")
            End If
            Printer.PaintPicture FImage1.Image, 0, 0, Printer.width, Printer.height
            If II < B Then
                Printer.NewPage
            End If
            aspen2 = Val(aspen2) + 1
            Imprimetotal = Imprimetotal + 1
        Next II
            Printer.EndDoc
        If Imprimetotal < 7 Then
            Sleep 30000
        ElseIf Imprimetotal < 16 Then
            Sleep 60000
        ElseIf Imprimetotal < 20 Then
            Sleep 90000
        ElseIf Imprimetotal < 28 Then
            Sleep 120000
        ElseIf Imprimetotal < 35 Then
            Sleep 150000
        ElseIf Imprimetotal < 42 Then
            Sleep 180000
        ElseIf Imprimetotal < 49 Then
            Sleep 210000
        Else
            Sleep 240000
        End If
            YY = 0
            Do Until YY = 1
                If Dir("C:\Temp\Proyecto.pdf") <> "" Then
                    List2.ListIndex = ZZ
                    YY = 0
                    Do Until YY = 1
                          If FileLen("C:\Temp\Proyecto.pdf") > 20 Then
                            YY = 1
                            Exit Do
                          End If
                    Loop
                    FileCopy "C:\Temp\Proyecto.pdf", (App.Path & "\unidos\" & List2.Text)
    '                Kill "C:\Temp\Proyecto.pdf"
                End If
            Loop
        continuardesde = "Continuar"
     End If
    Next ZZ
        MsgBox "Final de Generación de PDF", vbInformation, "Final"
 End If
End Sub

Private Sub Command5_Click()
Dim objFSO As Scripting.FileSystemObject
Dim objFile As File

Set objFSO = New Scripting.FileSystemObject
Set objFile = objFSO.GetFile("C:\SumaPDF_SIZE\ROLLO963Prueba\Images\1.tif")
Label3.Caption = Int((objFile.size) / 1000)
Set objFile = Nothing
Set objFSO = Nothing
End Sub

Private Sub Command6_Click()
    Frame1.Visible = True
    Command6.Enabled = False
End Sub

Private Sub Form_Load()
    strRollo = (App.Path & "\" & List3.Text & "\Corte.txt")
    strCarpetas = (App.Path & "\Carpetas.txt")
    strRuta = (App.Path & "\" & List3.Text & "\List2.txt")
    strUnidos = (App.Path & "\" & List3.Text & "\unidos")
    strImages = (App.Path & "\" & List3.Text & "\Images\" & II & ".tif")
    strHojas = (App.Path & "\" & List3.Text & "\Images\Hoja-000" & II & ".tif")
    i = 0
    aspen2 = 0
End Sub
