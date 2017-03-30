Attribute VB_Name = "frmExplorerModulo"
Public SourceNode As Object
Sub NovoExplorerLocalização(Nome As String)
    'dim fmainform.Name
    Set fMainForm = New frmExplorer

    fMainForm.Caption = Nome
    lJanela = lJanela + 1
    
    fMainForm.Tag = fMainForm.Name & lJanela
    'fMainForm.Caption = Nome
    fMainForm.Show
End Sub


Sub tvAdicionarItem(tvw As TreeView, _
                    KeyDoParente As String, _
                    KeyDoItem As String, _
                    RótuloDoItem As String, _
                    ÍconeDoItem As String, _
                    ÍconeSelecionado As String)
               
    Dim tv As TreeView
    Dim nodX As Node
    Set tv = tvw
    
    If KeyDoParente = "" Then
        Set nodX = tv.Nodes.Add(, tvwChild, KeyDoItem, RótuloDoItem, ÍconeDoItem, ÍconeSelecionado)
    Else
        Set nodX = tv.Nodes.Add(KeyDoParente, tvwChild, KeyDoItem, RótuloDoItem, ÍconeDoItem, ÍconeSelecionado)
    End If
    'Set nodX = tv.Nodes.Add(tvTreeView.SelectedItem.Text, tvwChild, "Hardware", "Hardware", 2)
    
End Sub
Sub tvRemoverItem(tvw As TreeView, _
                    KeyDoParente As String, _
                    KeyDoItem As String)
               
    Dim tv As TreeView
    Dim nodX As Node
    Set tv = tvw
    
    If KeyDoParente = "" Then
        Set nodX = tv.Nodes.Remove(tvwChild, KeyDoItem, RótuloDoItem, ÍconeDoItem, ÍconeSelecionado)
    Else
        Set nodX = tv.Nodes.Add(KeyDoParente, tvwChild, KeyDoItem, RótuloDoItem, ÍconeDoItem, ÍconeSelecionado)
    End If
    'Set nodX = tv.Nodes.Add(tvTreeView.SelectedItem.Text, tvwChild, "Hardware", "Hardware", 2)
    
End Sub



