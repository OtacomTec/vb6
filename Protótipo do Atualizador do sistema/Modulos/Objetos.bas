Attribute VB_Name = "Objetos"
Dim J As Integer
Dim I As Integer
Dim Constante As Integer

Public Function Verifica_TXT(Form As Form, Optional Nome_TXT_Verificar As String) As Boolean

  Dim strTamanho As String
  Dim strMatriz_Nome_TXT() As String
  Dim strString As String
  
  If Nome_TXT_Verificar <> Empty Then
     strMatriz_Nome_TXT() = Split(Nome_TXT_Verificar, ",")
     For I = 0 To Form.Count - 1
         If TypeOf Form.Controls(I) Is TextBox Then
            strString = Form.Controls(I).Name
            J = 0
            Do While strMatriz_Nome_TXT(J) <> Empty
               If strMatriz_Nome_TXT(J) = strString Then
                  strTamanho = Len(strMatriz_Nome_TXT(I))
                  strTamanho = strTamanho - 3
                  strMatriz_Nome_TXT(I) = Right(strMatriz_Nome_TXT(I), strTamanho)
                  Constante = Constante + 1
                  Exit Do
               Else
                    J = J + 1
               End If
            Loop
         End If
      Next I
  End If
    
  For I = 0 To Form.Count - 1
      On Error Resume Next
      If TypeOf Form.Controls(I) Is TextBox Then
         If strMatriz_Nome_TXT(0) = Empty Then
            strMatriz_Nome_TXT(I) = Form.Controls(I).Name
            strTamanho = Len(strMatriz_Nome_TXT(I))
            strTamanho = strTamanho - 3
            strMatriz_Nome_TXT(I) = Right(strMatriz_Nome_TXT(I), strTamanho)
         Else
            strTamanho = Len(strMatriz_Nome_TXT(I))
            strTamanho = strTamanho - 3
            strMatriz_Nome_TXT(I) = Right(strMatriz_Nome_TXT(I), strTamanho)
            Constante = Constante + 1
         End If
      End If
  Next I
  
  I = 0
    
  For I = 0 To strMatriz_Nome_TXT(Constante)
      MsgBox "O Campo " & strMatriz_Nome_TXT(I) & " não foi digitado.", vbInformation, Form.Caption
  Next I
  
End Function
Public Function Limpa_TXT(Form As Form) As String

  For I = 0 To Form.Count - 1
      On Error Resume Next
      If TypeOf Form.Controls(I) Is TextBox Then
         If Form.Controls(I).Text <> Empty Then
            Form.Controls(I).Text = Empty
         End If
      End If
  Next I
   
End Function
Public Function Maiusculo_TXT(Form As Form) As String

  For I = 0 To Form.Count - 1
      On Error Resume Next
      If TypeOf Form.Controls(I) Is TextBox Then
         If Form.Controls(I).Text <> Empty Then
            Form.Controls(I).Text = UCase(Form.Controls(I).Text)
         End If
      End If
  Next I
   
End Function
Public Function Desabilita_TXT(Form As Form, Optional Nome_TXT_Desabilita As String, Optional Numero_TXT_Desabilita As Integer, Optional Nome_TXT_Habilita As String, Optional Numero_TXT_Habilita As Integer) As String

  Dim strMatriz_Desabilita_TXT() As String
  Dim strMatriz_Habilita_TXT() As String
  
  If Nome_TXT_Desabilita <> Empty Then
     strMatriz_Desabilita_TXT() = Split(Nome_TXT_Desabilita, ",")
  End If
    
  If Nome_TXT_Habilita <> Empty Then
     strMatriz_Habilita_TXT() = Split(Nome_TXT_Habilita, ",")
  End If
  
  For I = 0 To Form.Count - 1
      On Error Resume Next
      If TypeOf Form.Controls(I) Is TextBox Or TypeOf Form.Controls(I) Is DataCombo Then
         If Nome_TXT_Desabilita = Empty And Nome_TXT_Habilita = Empty Then
            If Form.Controls(I).Enabled = True Then
               Form.Controls(I).Enabled = False
            End If
         End If
         If Nome_TXT_Desabilita <> Empty Then
            For J = 0 To (Numero_TXT_Desabilita - 1)
                If Form.Controls(I).Name = strMatriz_Desabilita_TXT(J) Then
                   If Form.Controls(I).Enabled = True Then
                      Form.Controls(I).Enabled = False
                   End If
                End If
            Next J
         End If
         If Nome_TXT_Habilita <> Empty Then
            For J = 0 To (Numero_TXT_Habilita - 1)
                If Form.Controls(I).Name = strMatriz_Habilita_TXT(J) Then
                   If Form.Controls(I).Enabled = False Then
                      Form.Controls(I).Enabled = True
                   End If
                End If
            Next J
         End If
      End If
  Next I
   
End Function
Public Function Desabilita_Botoes(Form As Form, Optional Nome_Botoes_Desabilita As String, Optional Numero_Botoes_Desabilita As Integer, Optional Nome_Botoes_Habilita As String, Optional Numero_Botoes_Habilita As Integer) As String
  
  Dim strMatriz_Desabilita_Botoes() As String
  Dim strMatriz_Habilita_Botoes() As String
  
  If Nome_Botoes_Desabilita <> Empty Then
     strMatriz_Desabilita_Botoes() = Split(Nome_Botoes_Desabilita, ",")
  End If
    
  If Nome_Botoes_Habilita <> Empty Then
     strMatriz_Habilita_Botoes() = Split(Nome_Botoes_Habilita, ",")
  End If
  
  For I = 0 To Form.Count - 1
      On Error Resume Next
      If TypeOf Form.Controls(I) Is CommandButton Then
         If Nome_Botoes_Desabilita = Empty And Nome_Botoes_Habilita = Empty Then
            If Form.Controls(I).Enabled = True Then
               Form.Controls(I).Enabled = False
            End If
         End If
         If Nome_Botoes_Desabilita <> Empty Then
            For J = 0 To (Numero_Botoes_Desabilita - 1)
                If Form.Controls(I).Name = strMatriz_Desabilita_Botoes(J) Then
                   If Form.Controls(I).Enabled = True Then
                      Form.Controls(I).Enabled = False
                   End If
                End If
            Next J
         End If
         If Nome_Botoes_Habilita <> Empty Then
            For J = 0 To (Numero_Botoes_Habilita - 1)
                If Form.Controls(I).Name = strMatriz_Habilita_Botoes(J) Then
                   If Form.Controls(I).Enabled = False Then
                      Form.Controls(I).Enabled = True
                   End If
                End If
            Next J
         End If
      End If
  Next I
  
End Function
Public Function Limpa_Variaveis_Memoria(Form As Form)
    
    
    Set Form.strTamanho = Nothing
    Set Form.strNomes = Nothing
    Set Form.strCombo = Nothing
    Set Form.strConsulta = Nothing
    
     
End Function
