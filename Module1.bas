Attribute VB_Name = "Module1"
Global Const LB_FINDSTRING = &H18F, LB_FINDSTRINGEXACT As Long = &H1A2, CB_ERR = (-1)
Public msg As String, strRollo As String, strCarpetas As String, strRuta As String, strUnidos As String, CB As Long, Aspen1 As String
Public FINDSTRING As String, A As String 'Variable que toma el texto de TXT
Public Rep As Integer, NoRep As Integer, Existe_Carpetas As Boolean, strHojas As String, strImages As String
Public Tiempo(1 To 10) As Long, Cantidad(1 To 10) As Long
Public i As Long
Public II As Long
Public YY As Long
Public ZZ As Long
Public ZZ1 As Long
Public idee
Public B As Long
Public continuardesde As String
Public E As String
Public aspo As String
Public aspen2 As String
Public capa As Long
Public capa2 As Long
Public cap As Long
Public Imprimetotal As Long
Public objFSO As Scripting.FileSystemObject 'Declaracion de FSO para calcular peso
Public objFile As File 'Archivo para medir el peso
Public ImgPeso As Integer

'Api para validar repetidos
#If Win32 Then
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
Const CB_SHOWDROPDOWN = &H14F
#Else
    Declare Function SendMessage Lib "User" _
        (ByVal hWnd As Integer, ByVal wMsg As Integer, _
         ByVal wParam As Integer, lParam As Any) As Long
Const WM_USER = &H400
Const CB_SHOWDROPDOWN = (WM_USER + 15)
#End If

Public Sub CalcularTiempo()
        If Imprimetotal = 1 Then
            Sleep 15000
        ElseIf Imprimetotal = 2 Then
            Sleep 20000
        ElseIf Imprimetotal < 7 Then
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
End Sub

Public Sub CargarCorte(List1 As ListBox, List3 As ListBox, lbl_Cantidad As Label)
        List3.ListIndex = ZZ1
        lbl_Cantidad.Caption = lbl_Cantidad.Caption + 1
        lbl_Cantidad.Refresh
        List1.Clear
        If Dir(strRollo) <> "" Then
            Open strRollo For Input As #1
            Do While Not EOF(1)
                Line Input #1, A
                List1.AddItem A
            Loop
            Close #1
        End If
End Sub

Public Sub CargarCarpeta(List3 As ListBox)
    If Dir(strCarpetas) <> "" Then
        Open strCarpetas For Input As #1
        Do While Not EOF(1)
            Line Input #1, A
            List3.AddItem A
        Loop
        Close #1
        Existe_Carpetas = True
    Else
        msg = MsgBox("Asegurese de que exista el archivo 'Carpetas.txt' en la ruta de la aplicacion y que contenga informacion", vbCritical, "SumaPDF")
        Existe_Carpetas = False
    End If
End Sub

Public Sub BuscarRepetidos(List2 As ListBox)
    Rep = 1
    List2.Clear
    If Dir(strRuta) <> "" Then
        Open strRuta For Input As #2
        Do While Not EOF(2)
            Line Input #2, Aspen1
            FINDSTRING = Trim$(Aspen1)
            CB = SendMessage(List2.hWnd, LB_FINDSTRINGEXACT, -1, ByVal FINDSTRING)
            If CB = CB_ERR Then
                List2.AddItem FINDSTRING
            Else
                repcount = Rep
                For Rep = 1 To repcount
                    FINDSTRING = Trim$("Repetido" & Rep & "_" & Aspen1)
                    CB = SendMessage(List2.hWnd, LB_FINDSTRINGEXACT, -1, ByVal FINDSTRING)
                    If CB = CB_ERR Then
                        List2.AddItem FINDSTRING
                        Rep = repcount
                    Else
                        repcount = repcount + 1
                    End If
                Next Rep
            End If
        Loop
        Close #2
    End If
End Sub
