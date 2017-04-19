Attribute VB_Name = "mdlTrataRegistroWindows"
Option Explicit

Public Sub sbSalvaValorNoRegistroDoWindows(ByVal Nome As String, _
                                           ByVal Secao As String, _
                                           ByVal Chave As String, _
                                           ByVal Valor As String)
                                           
    Call VBA.SaveSetting(Nome, Secao, Chave, Valor)

End Sub

Public Function fcRetornaValorDoRegistroDoWindows(ByVal Nome As String, _
                                                  ByVal Secao As String, _
                                                  ByVal Chave As String) As String
    
    fcRetornaValorDoRegistroDoWindows = VBA.GetSetting(Nome, Secao, Chave)
    
End Function

Public Sub sbDeletaValorNoRegistroDoWindows(ByVal Nome As String, _
                                            ByVal Secao As String, _
                                            ByVal Chave As String)
                                            
    Call VBA.DeleteSetting(Nome, Secao, Chave)
    
End Sub
