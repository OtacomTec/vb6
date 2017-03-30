Attribute VB_Name = "Verifica_txt"
Public Function Verifica_TXT(Form As Form) As Boolean

  Dim I As Integer
  Dim TXT As textbox
  Dim strNome As String
  Dim strTamanho As String
  Dim strMatrizNome(100) As String
  
  For I = 0 To Form.Count - 1
      On Error Resume Next
      Set TXT = Form.Controls(I)
      If Form.Controls(TXT.Name)(I) <> Empty Then
         If TXT.Text = Empty Then
            strNome = TXT.Name
            strTamanho = Len(strNome)
            strTamanho = strTamanho - 3
            strNome = Right(strNome, strTamanho)
            strMatrizNome(I) = strNome
            Verifica_TXT = True
         End If
      End If
  Next I
    
  For I = 0 To 100
      If strMatrizNome(I) <> Empty Then
         MsgBox "O Campo " & strMatrizNome(I) & " não foi digitado.", vbInformation, Form.Caption
      End If
  Next I
  
End Function
    
Public Function Limpa_TXT(Form As Form) As String

  Dim I As Integer
  Dim TXT As textbox
   
  For I = 0 To Form.Count - 1
      On Error Resume Next
      Set TXT = Form.Controls(I)
      If Form.Controls(TXT.Name)(I) <> Empty Then
         If TXT.Text <> Empty Then
            TXT.Text = Empty
         End If
      End If
  Next I
   
End Function
