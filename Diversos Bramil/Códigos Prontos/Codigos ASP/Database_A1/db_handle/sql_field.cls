<%
' Modulo classe Field
' 
' Parametri d'ingresso
' clsObjField.SetField - campo Field
' clsObjField.SetValue - imposta il valore del campo
' clsObjField.SetMode - se a True accetta in Parser i valore impostati in MyValue
'
' Parametri d'uscita
' clsObjField.NewValue - valore con separatori
' clsObjField.NewParser - valore nuova casella di controllo
' clsObjField.Updatable - campo aggiornabile
'

CLASS clsObjField

Private MyAttributes
Private MyType
Private MyValue
Private MyName
Private MyDefinedSize
Private MyMode
Private MySmart

Private FUNCTION IsEqual (a, b)
if a = b then
	IsEqual = True
else
	IsEqual = False
end if
End FUNCTION

Private FUNCTION IsTrue (MyNum)
if MyNum > 0 then
	IsTrue = True
else
	IsTrue = False
end if
End FUNCTION

Private FUNCTION Min (a, b)
Min = a
if a > b then
	Min = b
end if
End FUNCTION

Private FUNCTION IsUpdatable

if IsEqual(MyAttributes, &hFFFFFFFF) then
	IsUpdatable = FALSE
else
	if 	IsTrue(MyAttributes And &h8) then
		IsUpdatable = FALSE
	else
		Tipo = MyType
		if ((MyType > 1) AND (MyType < 7)) Then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		elseif ((MyType > 199) AND (MyType < 204)) Then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		elseif MyType=11 then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		elseif MyType=17 then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		elseif MyType=130 then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		elseif MyType=135 then
			IsUpdatable = IsTrue(MyAttributes And &h4)
		else
			IsUpdatable = FALSE
		end if
	end if
end if

END FUNCTION


FUNCTION ParserCampo

Dim MyCmp, Nome, Valore, CheckOFF, CheckON

MyCmp = MyType
Nome = MyName
Valore = MyValue

if MySmart then

	uscita = Valore
	if MyCmp=130 or MyCmp=200 or MyCmp=201 or MyCmp=202 or MyCmp=203 then
		if Not IsNull(Valore) then
			if Len(Valore) > 15 then
				uscita = Left(Valore, 12) & "..."
			end if
		end if
	elseif MyCmp = 11 then
		CheckON=""
		if Valore then
			uscita = "<center>SI</center>"
			uscita = "<center><input type='checkbox' name'" & Nome & "' value='ON' checked></center>"
		else
			uscita = "<center>NO</center>"
			uscita = "<center><input type='checkbox' name'" & Nome & "' value='OFF'></center>"
			'CheckON=" checked"
		end if
		'uscita = "<center><input type='radio' value='True' name='" & Nome & "'" & CheckON & "></center>"
	end if
else

	SELECT CASE MyCmp

	Case 130
	'testo o memo ACCESS 2000
	if MyDefinedSize > 255 then
		'memo - ok
			if Not IsNull(Valore) then
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'>" & Server.HTMLEncode(Valore) & "</textarea>"
			else
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'></textarea>"
			end if
	else
		'testo - ok
			uscita = "<input type='text' name='" & Nome & "' size='" & Min(MyDefinedSize,40) & "' maxlength='" & MyDefinedSize & "' value='" & Valore & "'>"
	end if

	Case 202
	'testo o memo ACCESS 2000 - MSDAC 2.5 o superiore
	if MyDefinedSize > 255 then
		'memo - ok
			if Not IsNull(Valore) then
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'>" & Server.HTMLEncode(Valore) & "</textarea>"
			else
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'></textarea>"
			end if
	else
		'testo - ok
			uscita = "<input type='text' name='" & Nome & "' size='" & Min(MyDefinedSize,40) & "' maxlength='" & MyDefinedSize & "' value='" & Valore & "'>"
	end if

	Case 203
	'testo o memo ACCESS 2000 - MSDAC 2.5 o superiore
	if MyDefinedSize > 255 then
		'memo - ok
			if Not IsNull(Valore) then
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'>" & Server.HTMLEncode(Valore) & "</textarea>"
			else
				uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'></textarea>"
			end if
	else
		'testo - ok
			uscita = "<input type='text' name='" & Nome & "' size='" & Min(MyDefinedSize,40) & "' maxlength='" & MyDefinedSize & "' value='" & Valore & "'>"
	end if

	Case 200
	'testo - ok
		uscita = "<input type='text' name='" & Nome & "' size='" & Min(MyDefinedSize,40) & "' maxlength='" & MyDefinedSize & "' value='" & Valore & "'>"

	Case 201
	'memo - ok
		uscita = "<!--webbot bot='Validation' I-Maximum-Length='65535' --><textarea rows='4' name='" & Nome & "' cols='40'>" & Valore & "</textarea>"

	Case 2
	'Intero
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Integer' S-Number-Separators='x' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 3
	'intero lungo
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Integer' S-Number-Separators='x' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 4
	'precisione singola
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Number' S-Number-Separators='x,' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 5
	'precisione doppia
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Number' S-Number-Separators='x,' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 6
	'valuta
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Number' S-Number-Separators='x,' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 17
	'intero byte
		uscita = "<!--webbot bot='Validation' startspan S-Data-Type='Integer' S-Number-Separators='x,' --><!--webbot bot='Validation' endspan --><input type='text' name='" & Nome & "' size='15' value='" & Valore & "'>"

	Case 72
	'ID Replica
		uscita = "<CENTER>REPLICA - n.a.</CENTER>"

	Case 135
	'ora e data
		uscita = "<input type='text' name='" & Nome & "' size='20' value='" & Valore & "'>"

	Case 11
	'logico
		CheckOFF=" checked"
		CheckON=""
		if Valore then
			CheckON=" checked"
			CheckOFF=""
		end if
		uscita = "<center>Yes<input type='radio' value='True' name='" & Nome & "'" & CheckON & "> No<input type='radio' value='False' name='" & Nome & "'" & CheckOFF & "></center>"

	Case 205
	'Oggetto OLE
		uscita = "<CENTER>FILE - n.a.</CENTER>"

	Case Else
	'caso nullo
		uscita = "<CENTER>FIELD - n.a.</CENTER>"

	End Select
end if

' debug
'ParserCampo = MyCmp & " " & uscita
ParserCampo = uscita

Set MyCmp=Nothing
Set Nome=Nothing
Set Valore=Nothing
Set CheckOFF=Nothing
Set CheckON=Nothing

END FUNCTION


FUNCTION FieldEval

Dim MyCmp, MyValore, SeparatoreCampo

MyCmp = MyType
MyValore = MyValue

SELECT CASE MyCmp

Case 130
'testo e memo Access 2000
	SeparatoreCampo = "'"
Case 200
'testo - ok
	SeparatoreCampo = "'"
Case 201
'memo - ok
	SeparatoreCampo = "'"
Case 202
'memo - ok
	SeparatoreCampo = "'"
Case 203
'memo - ok
	SeparatoreCampo = "'"
Case 2
'Intero
	SeparatoreCampo = ""
Case 3
'intero lungo
	SeparatoreCampo = ""
Case 4
'precisione singola
	SeparatoreCampo = ""
Case 5
'precisione doppia
	SeparatoreCampo = ""
Case 6
'valuta
	SeparatoreCampo = ""
Case 17
'intero byte
	SeparatoreCampo = ""
Case 72
'ID Replica
	SeparatoreCampo = ""
	if Not IsNull(MyValore) then
		MyValore=Null
	end if
Case 135
'ora e data
	SeparatoreCampo = "#"
Case 11
'logico
	SeparatoreCampo = ""
	if MyValore then
		MyValore = -1
	else
		MyValore = 0
	end if
Case 205
'Oggetto OLE
	SeparatoreCampo = ""
	if Not IsNull(MyValore) then
		MyValore=Null
	end if
Case Else
'caso nullo
	SeparatoreCampo = ""
	if Not IsNull(MyValore) then
		MyValore=Null
	end if
End Select
FieldEval = SeparatoreCampo & MyValore & SeparatoreCampo
Set MyCmp=Nothing
Set MyValore=Nothing
Set SeparatoreCampo=Nothing
END FUNCTION

' proprieta' di estrazione della voce Updatable
Public Property Get Updatable
Updatable = IsUpdatable
End Property

' proprieta' di estrazione della voce NewField
Public Property Get NewValue
NewValue = FieldEval
End Property

' proprieta' di estrazione della voce NewParser
Public Property Get NewParser
NewParser = ParserCampo
End Property

' proprieta' di impostazione della modalita' di visualizzazione del campo
Public Property Let SetMode(NewMode)
if NewMode then
	MyMode=True
else
	MyMode=False
end if
End Property

' proprieta' di impostazione del campo
Public Property Let SetField(NewField)
MySmart=False
MyAttributes=NewField.Attributes
MyType=NewField.Type
if MyMode then
	MyValue=NewField.Value
else
	MyValue=Null
end if
MyName=NewField.Name
MyDefinedSize=NewField.DefinedSize
End Property

' proprieta' di impostazione del valore
Public Property Let SetValue(NewValue)
MyValue=NewValue
End Property

' proprieta' di impostazione del valore
Public Property Let SetSmartRef(NewValue)
if NewValue then
	MySmart=True
else
	MySmart=False
end if
End Property

END CLASS
%>
