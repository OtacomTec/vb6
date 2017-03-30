VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'GrFinger Sample
'(c) 2005 Griaule Tecnologia Ltda.
'http://www.griaule.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Tecnologia makes no representations concerning either the merchantability
'of this software or the suitability of this sample for any particular purpose.
'
'THIS SAMPLE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR
'IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
'IN NO EVENT SHALL GRIAULE BE LIABLE FOR ANY DIRECT, INDIRECT,
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'You can download the free version of GrFinger directly from Griaule website.
'
'These notices must be retained in any copies of any part of this
'documentation and/or sample.
'
'-------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Database routines
' -----------------------------------------------------------------------------------

Option Explicit
Option Base 0

' the database we'll be connecting to
'Const DBFile = "../../../GrFingerSample.mdb"
'nova tunix
Const DBFile = "GrFingerSample.mdb"
Const ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

' the connection object
Dim connection As New ADODB.connection

' Open connection
Public Function OpenDB() As Boolean
    On Error GoTo cantopen
    connection.Open (ConnectionString & CurrentProject.Path & "\" & DBFile)
    OpenDB = True
    Exit Function
cantopen:
    On Error GoTo 0
    OpenDB = False
End Function

' Close conection
Public Sub closeDB()
    If connection.State = adStateOpen Then connection.Close
End Sub

' Clear database
Public Sub clearDB()
    ' run "clear" query
    connection.Execute "DELETE FROM enroll"
End Sub

' Add template to database. Returns added template ID.
Public Function AddTemplate(ByRef template() As Byte) As Long
    Dim rs As New ADODB.Recordset
    rs.Open "enroll", connection, adOpenStatic, adLockOptimistic
    ' Add a new row.
    rs.AddNew
    rs("template") = template
    
    rs.Update
    ' return ID
    AddTemplate = rs("ID")
End Function

' Returns a ADODB.recordset with all enrolled templates from database.
Public Function getTemplates() As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "select * from enroll", connection, adOpenDynamic, adLockReadOnly
    Set getTemplates = rs
End Function

' returns template with the supplied ID.
Public Function getTemplate(ID As Long) As Byte()
    Dim rs As New ADODB.Recordset
    ' Get query response
    rs.Open "select * from enroll where ID = " & ID, connection, adOpenDynamic, adLockReadOnly
    ' No results?
    If rs.EOF And rs.BOF Then
        Dim tptVazio(0) As Byte
        rs.Close
        getTemplate = tptVazio
    Else
        ' Deserialize template and return it
        Dim tpt() As Byte
        tpt = rs("template")
        rs.Close
        getTemplate = tpt
    End If
End Function


