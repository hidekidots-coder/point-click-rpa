Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
#Else
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
#End If

' ============================================================
' POINT CLICK - RPA AUTOMATION
' ============================================================
'
' Purpose:
'   Automate data entry in a Citrix application by simulating
'   mouse clicks and keyboard input based on Excel data.
'
' Technologies:
'   VBA
'   Windows API (mouse_event / SetCursorPos)
'   Microsoft Excel
'
' ============================================================

' ===== DELAY =====
Private Sub Esperar(ByVal segundos As Double)
    Dim t As Double
    t = Timer + segundos
    Do While Timer < t
        DoEvents
    Loop
End Sub

' ===== SINGLE CLICK =====
Private Sub Clique(ByVal X As Long, ByVal Y As Long, ByVal velocidade As Double)
    SetCursorPos X, Y
    Esperar velocidade
    mouse_event 2, 0, 0, 0, 0   ' Mouse down
    mouse_event 4, 0, 0, 0, 0   ' Mouse up
    Esperar velocidade
End Sub

' ===== DOUBLE CLICK =====
Private Sub DuploClique(ByVal X As Long, ByVal Y As Long, ByVal velocidade As Double)
    SetCursorPos X, Y
    Esperar velocidade
    
    mouse_event 2, 0, 0, 0, 0
    mouse_event 4, 0, 0, 0, 0
    
    Esperar velocidade / 2
    
    mouse_event 2, 0, 0, 0, 0
    mouse_event 4, 0, 0, 0, 0
    
    Esperar velocidade
End Sub

' ============================================================
' MAIN PROCEDURE
' ============================================================
Public Sub Point_Click()

    On Error GoTo ErroHandler

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Dim rapido As Double
    Dim lento As Double
    
    Dim strB As String
    Dim strE As String
    Dim strG As String
    Dim strH As String
    Dim strI As String
    Dim strJ As String
    
    Dim tentativas As Integer

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Set ws = ActiveSheet
    
    rapido = 0.7
    lento = 0.8
    
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' --------------------------------------------------------
    ' Start
    ' --------------------------------------------------------
    
    MsgBox "Clique no Citrix e não mexa no mouse/teclado!", vbInformation
    Esperar 2
    
    ' --------------------------------------------------------
    ' Main Loop
    ' --------------------------------------------------------
    
    For i = 1 To ultimaLinha
        
        strB = Trim(CStr(ws.Cells(i, "B").Value))
        strE = Trim(CStr(ws.Cells(i, "E").Value))
        strG = Trim(CStr(ws.Cells(i, "G").Value))
        strH = Trim(CStr(ws.Cells(i, "H").Value))
        strI = Trim(CStr(ws.Cells(i, "I").Value))
        strJ = Trim(CStr(ws.Cells(i, "J").Value))
        
        ' ===== RÁPIDO =====
        Clique 1776, 97, rapido
        DuploClique 1986, 315, rapido
        
        ' ===== HAWB =====
        Clique 2195, 398, lento
        SendKeys strB, True
        
        ' ===== FLUXO =====
        Clique 2184, 573, lento
        Clique 2150, 643, lento
        Clique 2152, 600, lento
        Clique 2165, 761, lento
        
        Clique 2134, 679, lento
        Clique 2148, 772, lento
        Clique 2140, 702, lento
        Clique 2174, 727, lento
        
        ' ===== DR =====
        Clique 2333, 811, lento
        SendKeys strE, True
        
        ' ===== ETD =====
        Clique 2463, 811, lento
        SendKeys strG, True
        
        ' ===== ETA =====
        Clique 2621, 811, lento
        SendKeys strH, True
        
        ' ===== USUÁRIO =====
        Clique 2434, 645, lento
        SendKeys "w", True
        SendKeys "{ENTER}", True
        
        ' ===== DESTINO =====
        Clique 2432, 577, lento
        SendKeys "v", True
        SendKeys "{ENTER}", True
        
        ' ===== REF CLIENTE =====
        Clique 2407, 525, lento
        SendKeys strI, True
        
        ' ===== DESPACHANTE =====
        Clique 2536, 475, lento
        Clique 2333, 554, lento
        Clique 1857, 423, lento
        Clique 1976, 594, lento
        
        ' ===== AGENTE =====
        Clique 2446, 427, lento
        
        SendKeys "^a", True
        SendKeys "{DEL}", True
        Application.Wait Now + TimeValue("0:00:01")
        
        SendKeys strJ, True
        Application.Wait Now + TimeValue("0:00:02")
        
        ' Caixinha
        Clique 2534, 422, lento
        ' Processo
        Clique 1997, 425, lento
        ' OK
        Clique 1973, 592, lento
        
        ' ===== CRIAR PROCESSO =====
        Clique 1974, 375, lento
        
        ' ===== PEGAR PROCESSO =====
        Clique 2154, 327, lento
        Esperar 0.5
        SendKeys "^a", True
        Esperar 0.3
        SendKeys "^c", True
        Esperar 0.5
        
        ' ===== COLAR NO EXCEL =====
        ws.Cells(i, "C").Select
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
        
        On Error Resume Next
        ActiveSheet.Paste
        If Err.Number <> 0 Then
            Err.Clear
            SendKeys "^v", True
        End If
        On Error GoTo 0
        
        Application.Wait Now + TimeValue("0:00:01")
        Application.CutCopyMode = False
        
        ' ===== EVITAR LOOP INFINITO =====
        tentativas = 0
        
        Do While IsEmpty(ws.Cells(i, "C").Value) And tentativas < 5
            Application.Wait Now + TimeValue("0:00:01")
            tentativas = tentativas + 1
        Loop
        
        ' ===== FECHAR =====
        Clique 2825, 250, lento
        Clique 2814, 280, lento
        
    Next i
    
    MsgBox "Processo finalizado!", vbInformation

Saida:
    Set ws = Nothing
    Exit Sub

ErroHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
    Resume Saida
    
End Sub
