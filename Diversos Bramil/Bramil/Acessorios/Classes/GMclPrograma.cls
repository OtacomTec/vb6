VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GM_clPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrSQL As String

Function EmpresasCadastradas()
    mstrSQL = "SELECT tEmpresasDef.strNomeResumidotEmpDef, " & _
                     "tEmpresasDef.bCodFilialSigatEmpDef, " & _
                     "tEmpresasDef.bTabPadraoEstoquetEmpDef, " & _
                     "tEmpresasDef.bTabPadraoGirotEmpDef, " & _
                     "tEmpresasDef.bTabTipoEstoquetEmpDef " & _
                "FROM tEmpresasDef " & _
          "INNER JOIN tProgramastEmpresas " & _
                  "ON tEmpresasDef.iCodEmpresatEmpDef = tProgramastEmpresas.iCodEmpresatPrgEmp " & _
               "Where (((tProgramastEmpresas.strCodProgramatPrgEmp) = '" & pstrCodPrograma & "')) " & _
            "ORDER BY tEmpresasDef.strNomeResumidotEmpDef;" 'strNomeFantasiatEmpDef;"
    ADO_Consulta pdbConfus, mstrSQL, prsEmpresas
    
End Function

Function ADO_Consulta(pbdConnection As ADODB.Connection, _
                      pstrQuery As String, _
                      prsRecordSet As ADODB.Recordset, _
                      Optional pstrCursor As String) As ADODB.Recordset

    On Error GoTo Erro_QY
    'Dim prsRecordSet As ADODB.Recordset
    Set prsRecordSet = New ADODB.Recordset
    
'    Screen.MousePointer = 11
'    ADO_Consulta = False
    If pbdConnection.State = adStateOpen Then      'Testa se o banco esta conectado
        If prsRecordSet.State = adStateOpen Then    'Testa se existe uma recordset ativa
            prsRecordSet.Close                      'Se existir fecha a recordset e prepara uma nova.
            Set prsRecordSet = Nothing
            Set prsRecordSet = New ADODB.Recordset
        End If
        If pstrCursor = "S" Then
            prsRecordSet.CursorLocation = adUseServer
            prsRecordSet.Open pstrQuery, pbdConnection, adOpenForwardOnly
        Else
            prsRecordSet.CursorLocation = adUseClient
            prsRecordSet.Open pstrQuery, pbdConnection, adOpenStatic
        End If
        'prsRecordSet.Open pstrQuery, pbdConnection, adOpenDynamic
    End If
    On Error GoTo 0
    
    If Not prsRecordSet.EOF Then Set ADO_Consulta = prsRecordSet
    
    Exit Function
    
Erro_QY:
    'If Err.Number = -2147217865 Or Err.Number = -2147217900 Then Exit Function
    Screen.MousePointer = 0
    If MsgBox("Ocorreu o erro número: " & Err.Number & vbCr & vbCr & "Descrição: " & Err.Description & vbCr & vbCr & "Deseja tentar novamente ?", vbCritical + vbYesNo, "pfboQuery") = vbYes Then
        Resume
        Exit Function
    End If
    On Error GoTo 0
End Function
