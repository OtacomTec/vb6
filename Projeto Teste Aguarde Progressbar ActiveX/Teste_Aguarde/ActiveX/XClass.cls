VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AbrirInterface()
    XForm.Show
End Sub

Public Sub Incrementar_Progresso(Optional ByVal Incremento As Long = 1)

    If XForm.pgbProgresso.Visible Then
    
        If CLng(XForm.pgbProgresso.Value + Incremento) <= XForm.pgbProgresso.Max Then
            XForm.ZOrder 0
            If Incremento > 1 Then
                XForm.pgbProgresso.Value = CLng(Incremento)
            Else
                XForm.pgbProgresso.Value = XForm.pgbProgresso.Value + 1
            End If
            XForm.lblAndamento.Caption = XForm.pgbProgresso.Value & " / " & XForm.pgbProgresso.Max
        End If
        
    End If
    
End Sub

Public Sub setMax_Progresso(ByVal lonMax As Long)
    XForm.pgbProgresso.Max = lonMax
End Sub

Public Sub setMin_Progresso(ByVal lonMin As Long)
    XForm.pgbProgresso.Min = lonMin
End Sub

Public Sub Destruir()
'    Call Exibir_Barra_Progresso
'    XForm.setAltura_Padrao_Form
    Unload XForm
End Sub

'Public Sub Exibir_Barra_Progresso(Optional ByVal booExibir As Boolean = True)
'    XForm.lblAndamento.Visible = booExibir
'    XForm.pgbProgresso.Visible = booExibir
'End Sub
