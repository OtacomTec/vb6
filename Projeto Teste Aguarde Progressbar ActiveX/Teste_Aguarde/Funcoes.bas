Attribute VB_Name = "Funcoes"
Option Explicit

Public Sub Incrementar_Progresso(Optional ByVal Incremento As Long = 1)

    If frmTeste_Aguarde.pgbProgresso.Visible Then
    
        If (frmTeste_Aguarde.pgbProgresso.Value + Incremento) <= frmTeste_Aguarde.pgbProgresso.Max Then
            frmTeste_Aguarde.pgbProgresso.Value = Incremento
        End If
        
    End If

End Sub

Public Sub Destruir()
    End
End Sub
