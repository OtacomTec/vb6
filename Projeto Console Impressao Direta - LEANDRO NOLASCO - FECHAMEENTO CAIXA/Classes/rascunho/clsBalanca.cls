VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsBalanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'local variable(s) to hold property value(s)
Private mvarnumero_balanca As Integer 'local copy
Private mvarmarca_balanca As String 'local copy
Private mvarcapacidade_balanca As String 'local copy
Private mvarnumero_inmetro_balanca As String 'local copy
Private mvarcodigo_empresa As String 'local copy
Public Property Let codigo_empresa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.codigo_empresa = 5
    mvarcodigo_empresa = vData
End Property


Public Property Get codigo_empresa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.codigo_empresa
    codigo_empresa = mvarcodigo_empresa
End Property



Public Property Let numero_inmetro_balanca(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.numero_inmetro_balanca = 5
    mvarnumero_inmetro_balanca = vData
End Property


Public Property Get numero_inmetro_balanca() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.numero_inmetro_balanca
    numero_inmetro_balanca = mvarnumero_inmetro_balanca
End Property



Public Property Let capacidade_balanca(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.capacidade_balanca = 5
    mvarcapacidade_balanca = vData
End Property


Public Property Get capacidade_balanca() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.capacidade_balanca
    capacidade_balanca = mvarcapacidade_balanca
End Property



Public Property Let marca_balanca(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.marca_balanca = 5
    mvarmarca_balanca = vData
End Property


Public Property Get marca_balanca() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.marca_balanca
    marca_balanca = mvarmarca_balanca
End Property



Public Property Let numero_balanca(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.numero_balanca = 5
    mvarnumero_balanca = vData
End Property


Public Property Get numero_balanca() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.numero_balanca
    numero_balanca = mvarnumero_balanca
End Property




