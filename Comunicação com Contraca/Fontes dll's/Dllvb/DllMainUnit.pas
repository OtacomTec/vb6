unit DllMainUnit;

interface

uses
  Windows, Forms, SysUtils, Dialogs,
  Classes, Controls, StdCtrls, ExtCtrls, PBJustOne, VaClasses,
  VaComm;

type
  TDllMainForm = class(TForm)
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    RadioGroup3: TRadioGroup;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    TimerPooling: TTimer;
    TimerTimeOutResponse: TTimer;
    VaComm1: TVaComm;
    procedure RadioGroup1Click(Sender: TObject);
    procedure RadioGroup2Click(Sender: TObject);
    procedure RadioGroup3Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure TimerPoolingTimer(Sender: TObject);
    procedure TimerTimeOutResponseTimer(Sender: TObject);
    procedure VaComm1RxFlag(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
	cSTX = #$2;
  cETX = #$3;
  cMAXTERMINAL = 5;
  cMAXSTRBUFFER = 4096;
  cCRC = '    ';
	CRCError = '1';
	NoResponseError = '2';
  TerminalTimeOutError = '3';
  cSETDATETIME = '01';
	cSENDMESSAGE = '02';
  cSETPARAMETERS = '03';
  cQUESTION = '04';
  cANSWER = '05';
  cTERMINALTIMEOUT = '01';
  cCONDITIONAFTERTIMEOUT = '02';
	cENABLED = '1';
  cDISABLED = '0';

var
  DllMainForm: TDllMainForm;
  mTerminals: array [1..5] of string;
  mActualTerminal: Integer = 1;
  mBufferIn, mBufferOut: array [1..cMAXSTRBUFFER] of string;
  mPtrIn1: Integer = 1;
  mPtrIn2: Integer = 1;
  mPtrOut1: Integer = 1;
  mPtrOut2: Integer = 1;
	mIntervalTime, mTerminalResponseTime: Integer;
  mBaudRate: Integer = 4800;
	mActualCommandAndTerminal: string;
	mErrorOnComunication: string;
  mF: TextFile;

 	function ActiveDll: LongInt; stdcall;                                                            // Ativa a dll
	function DeactiveDll: LongInt; stdcall;                                                 // Desativa a Dll
	function SetComm(CommPort: LongInt): LongInt; stdcall;																	// Escolhe a porta serial
	function SetBaudRate(BaudRate: LongInt): LongInt; stdcall;															// Escolhe a velocidade de comunicação
	function OpenComm: LongInt; stdcall;																										// Abre a porta serial
	function CloseComm: LongInt; stdcall; 																									// Fecha a porta serial
	function InsertTerminal (Terminal: LongInt): LongInt; stdcall;													// Registra uma catraca na lista de pooling
	function DeleteTerminal (Terminal: LongInt): LongInt; stdcall;													// Retira uma catraca da lista de pooling
	function EnableTerminal (Terminal: LongInt): LongInt; stdcall;													// Liga o pooling para uma catraca
	function DisableTerminal (Terminal: LongInt): LongInt; stdcall;   											// Desliga o pooling para uma catraca
	function SetPoolingIntervalTime(IntervalTime: LongInt): LongInt; stdcall; 							// Tempo do ciclo de execução do pooling
	function SetTerminalResponseTime(Time: LongInt): LongInt; stdcall;											// Tempo de espera para completar comunicação
 	function SetTerminalTimeOut (TerminalTimeOut: PChar): LongInt; stdcall;   							// Envia a catraca o tempo de espera de uma resposta
  function SetConditionAfterTimeOut (TerminalCondition: PChar): LongInt; stdcall;	      	// Envia a catraca a resolução se decorrido o timeout
	function SetDateTime (TerminalCurrentDateTime: PChar): LongInt; stdcall;		    				// Atualiza data e hora da catraca
  function SendMessage (TerminalTimeMessagePersonalMessage: PChar): LongInt; stdcall;     // Envia uma mensagem a catraca
  function SetParameters (TerminalVariableValue: PChar): LongInt; stdcall                 // Envia parâmetro a catraca
	function StartPooling: LongInt; stdcall;                                                // Inicializa o pooling
	function StopPooling: LongInt; stdcall;                                                 // Interrompe o pooling
	function SetCommShow: LongInt; stdcall;																									// Mostra a form de propriedades e inicializa a porta serial
	function Question: PChar; stdcall;																											// Verifica se há crachá para tratamento
  function	Answer(TerminalBadgePositionStatusTimeMessagePersonalMessage: PChar):         // Informa condição do crachá
    LongInt; stdcall;
 	function  _fCalculateCRCString(_vStr: string): string; stdcall;													// Caucula o crc da mensagem
	function  _fSendStr(_vStr: string): LongInt; stdcall;										                // Envia uma string a catraca
	function  _fSetDateTime(_vStr: string): LongInt; stdcall;																// Acerta o calendário da catraca
	function  _fSetParameters(_vStr: string): LongInt; stdcall;															// Envia parâmetro a catraca
	function  _fQuestion(_vStr: string): LongInt; stdcall;																	// Pergunta a catraca se há crachá
	function  _fAnswer(_vStr: string): LongInt; stdcall;																		// Libera a catraca para passagem
	function	_fSendMessage(_vStr: string): LongInt; stdcall;																// Envia uma mensagem a catraca

implementation

{$R *.DFM}

//MWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMW
// Funcoes e procedimentos exportados
//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM

function ActiveDll: LongInt;
begin { ActiveDll }
  try
		DllMainForm := TDllMainForm.Create(Application);
    Result := 1;
  except
    Result := 0;
  end;
end; { ActiveDll }

function DeactiveDll: LongInt;
begin { DeactiveDll }
  try
  	DllMainForm.Free;
    Result := 1;
  except
    Result := 0;
  end;
end; { DeactiveDll }

function InsertTerminal(Terminal: LongInt): LongInt;
var
	mI,mPosition: Integer;
  mExist: Boolean;
begin { InsertTerminal }
	mPosition := 0;
  mExist := False;
  for mI := 1 to cMAXTERMINAL do
  begin
  	if Format('%.2d',[Terminal])= Copy(mTerminals[mI],1,2) then
    begin
    	mExist := True;
      Break;
    end
    else
    begin
    	if (mTerminals[mI] = '') and (mPosition = 0) then mPosition := mI;
    end;
  end;
	if mExist then
  begin
		Result := 0;
  end
	else
  begin
  	if mPosition <> 0 then
    begin
    	Result := 1;
			mTerminals[mPosition] := Format('%.2d,0',[Terminal]);
    end
    else
    begin
			Result := 2;
    end;
  end;
end; { InsertTerminal }

function DeleteTerminal(Terminal: LongInt): LongInt;
var
	mI: Integer;
begin { DeleteTerminal }
	Result := 0;
  mI := 1;
  repeat
	  if (mTerminals[mI] = Format('%.2d,1',[Terminal])) or (mTerminals[mI] = Format('%.2d,0',[Terminal])) then
	  begin
		  mTerminals[mI] := '';
	    Result := 1;
      Break;
  	end;
   	Inc(mI);
  until mI > cMAXTERMINAL;
end; { DeleteTerminal }

function EnableTerminal(Terminal: LongInt): LongInt;
var
	mI: Integer;
begin { EnableTerminal }
	Result := 0;
  mI := 1;
  repeat
	  if mTerminals[mI] = Format('%.2d,0',[Terminal]) then
  	begin
			mTerminals[mI] := Format('%.2d,1',[Terminal]);
  	  Result := 1;
    	Break;
	  end;
    Inc(mI);
  until mI > cMAXTERMINAL;
end;  { EnableTerminal }

function DisableTerminal(Terminal: LongInt): LongInt;
var
	mI: Integer;
begin { DisableTerminal }
	Result := 0;
  mI := 1;
  repeat
	  if mTerminals[mI] = Format('%.2d,1',[Terminal]) then
  	begin
			mTerminals[mI] := Format('%.2d,0',[Terminal]);
      Result := 1;
    	Break;
	  end;
    Inc(mI);
  until mI > cMAXTERMINAL;
end;  { DisableTerminal }

function SetPoolingIntervalTime(IntervalTime: LongInt): LongInt;
begin { SetPoolingIntervalTime }
  try
  	mIntervalTime := IntervalTime;
	  DllMainForm.TimerPooling.Interval := mIntervalTime;
    Result := 1;
  except
    Result := 0;
  end;
end;  { SetPoolingIntervalTime }

function SetTerminalResponseTime(Time: LongInt): LongInt;
begin  { SetTimeTerminalResponse }
  try
  	mTerminalResponseTime := Time;
    DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
    Result := 1;
  except
    Result := 0;
  end;
end;  { SetTimeTerminalResponse }

function StartPooling: LongInt;
begin { StartPooling }
  try
  	DllMainForm.TimerPooling.Enabled := True;
    Result := 1;
  except
    Result := 0;
  end;
end;  { StartPooling }

function StopPooling: LongInt;
begin { StopPooling }
  try
  	DllMainForm.TimerPooling.Enabled := False;
    Result := 1;
  except
    Result := 0;
  end;
end;  { StopPooling }

function SetComm(CommPort: LongInt): LongInt;
begin { SetComm }
  try
  	DllMainForm.VaComm1.PortNum := CommPort;
    Result := 1;
  except
    Result := 0;
  end;
end; { SetComm }

function SetBaudRate(BaudRate: LongInt): LongInt;
begin { SetBaudRate }
  if BaudRate = 4800 then
  begin
		DllMainForm.VaComm1.BaudRate := br4800;
    Result := 1;
  end
  else
  begin
    if Baudrate = 57600 then
    begin
  		DllMainForm.VaComm1.BaudRate := br57600;
      Result := 1;
    end
    else
    begin
      Result := 0;
    end;
  end;
end; { SetBaudRate }

function SetCommShow: LongInt;
begin { SetCommShow }
  try
    DllMainForm.Visible := True;
    Result := 1;
  except
    Result := 0;
  end;
end; { SetCommShow }

function OpenComm: LongInt;
begin { OpenComm }
	try
  	DllMainForm.VaComm1.Open;
		Result := 1;
	except
    Result := 0;
  end;
end; { OpenComm }

function CloseComm: LongInt;
begin { CloseComm }
	try
    DllMainForm.TimerPooling.Enabled := False;
	  DllMainForm.VaComm1.Close;
    Result := 1;
  except
    Result := 0;
  end;
end; { CloseComm }

//          1         2
// 1234567890123456789012
// NT,DD/MM/YYYY,HH:MM:SS
function SetDateTime(TerminalCurrentDateTime: PChar): LongInt;
begin { SetDateTime }
  mBufferOut[mPtrOut1] := cSETDATETIME+','+
    StrPas(TerminalCurrentDateTime);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end;  { SetDateTime }

//
// 12345
// NT,TO
function SetTerminalTimeOut(TerminalTimeOut: PChar): LongInt;
begin { SetTerminalTimeOut }
  mBufferOut[mPtrOut1] := cSETPARAMETERS+','+
    Copy(StrPas(TerminalTimeOut),1,2)+','+
    cTERMINALTIMEOUT+','+
    Copy(StrPas(TerminalTimeOut),4,2);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end;  { SetTerminalTimeOut }

//
// 1234
// NT,C
function SetConditionAfterTimeOut (TerminalCondition: PChar): LongInt;
begin { SetConditionAfterTimeOut }
  if Copy(StrPas(TerminalCondition),4,1) = 'B' then
    mBufferOut[mPtrOut1] := cSETPARAMETERS+','+
      Copy(StrPas(TerminalCondition),1,2)+','+
      cCONDITIONAFTERTIMEOUT+',00'
  else
    mBufferOut[mPtrOut1] := cSETPARAMETERS+','+
      Copy(StrPas(TerminalCondition),1,2)+','+
      cCONDITIONAFTERTIMEOUT+',01';
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end; { SetConditionAfterTimeOut }

//          1         2         3         4
// 12345678901234567890123456789012345678901
// NT,TEMP,MENSAGEMLINHA1__,MENSAGEMLINHA2__
function SendMessage(TerminalTimeMessagePersonalMessage: PChar): LongInt;
begin { SendMessage }
  mBufferOut[mPtrOut1] := cSENDMESSAGE+','+
    Copy(StrPas(TerminalTimeMessagePersonalMessage),1,2)+','+
    Copy(StrPas(TerminalTimeMessagePersonalMessage),4,2)+','+
    Copy(StrPas(TerminalTimeMessagePersonalMessage),9,16)+
    Copy(StrPas(TerminalTimeMessagePersonalMessage),26,16);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end;  { SendMessage }

function SetParameters(TerminalVariableValue: PChar): LongInt;
begin { SetParameters }
  mBufferOut[mPtrOut1] := cSETPARAMETERS+','+StrPas(TerminalVariableValue);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end;  { SetParameters }

function Question: PChar; stdcall;
begin { Question }
  if mPtrIn1 <> mPtrIn2 then
  begin
		Result := PChar(mBufferIn[mPtrIn2]);
    mPtrIn2 := mPtrIn2 + 1;
    if mPtrIn2 > cMAXSTRBUFFER then mPtrIn2 := 1;
  end
  else
  begin
		Result := nil;
  end;
end;  { Question }

//          1         2         3         4         5
// 12345678901234567890123456789012345678901234567890123456
// NT,XXXXXXXXXX,S,C,TEMP,MENSAGEMLINHA1__,MENSAGEMLINHA2__
function	Answer(TerminalBadgePositionStatusTimeMessagePersonalMessage: PChar): LongInt;
begin { Answer }
  if Length(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage)) = 17 then
  begin
    if Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),17,1) = 'B' then
      mBufferOut[mPtrOut1] := cANSWER+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),1,16)+'N'
    else
      mBufferOut[mPtrOut1] := cANSWER+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),1,16)+'S';
  end
  else
  begin
    if Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),17,1) = 'B' then
      mBufferOut[mPtrOut1] := cANSWER+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),1,16)+'N'+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),19,2)+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),24,16)+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),41,16)
    else
      mBufferOut[mPtrOut1] := cANSWER+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),1,16)+'S'+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),19,2)+','+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),24,16)+
      Copy(StrPas(TerminalBadgePositionStatusTimeMessagePersonalMessage),41,16);
  end;
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
  Result := 1;
end;  { Answer }

//MWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMW
//	Eventos da form principal
//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM

procedure TDllMainForm.ComboBox1Change(Sender: TObject);
begin { ComboBox1Change }
	DllMainForm.VaComm1.DeviceName := ComboBox1.Text;
end;  { ComboBox1Change }

procedure TDllMainForm.ComboBox2Change(Sender: TObject);
begin { ComboBox2Change }
  case ComboBox2.ItemIndex of
    0:	VaComm1.BaudRate := br110;
		1:	VaComm1.BaudRate := br300;
		2:	VaComm1.BaudRate := br600;
		3:	VaComm1.BaudRate := br1200;
		4:	VaComm1.BaudRate := br2400;
		5:	VaComm1.BaudRate := br4800;
		6:	VaComm1.BaudRate := br9600;
		7:	VaComm1.BaudRate := br14400;
		8:	VaComm1.BaudRate := br19200;
    9:	VaComm1.BaudRate := br38400;
		10:	VaComm1.BaudRate := br56000;
		11:	VaComm1.BaudRate := br57600;
		12:	VaComm1.BaudRate := br115200;
		13:	VaComm1.BaudRate := br128000;
		14:	VaComm1.BaudRate := br256000;
	end;
end;  { ComboBox2Change }

procedure TDllMainForm.RadioGroup1Click(Sender: TObject);
begin { RadioGroup1Click }
	case RadioGroup1.ItemIndex of
		0: VaComm1.Stopbits := sb1;
		1: VaComm1.Stopbits := sb15;
		2: VaComm1.Stopbits := sb2;
	end;
end;  { RadioGroup1Click }

procedure TDllMainForm.RadioGroup2Click(Sender: TObject);
begin { RadioGroup2Click }
	case RadioGroup2.ItemIndex of
		1: VaComm1.DataBits := db4;
		2: VaComm1.DataBits := db5;
		3: VaComm1.DataBits := db6;
		4: VaComm1.DataBits := db7;
		5: VaComm1.DataBits := db8;
	end;
end;  { RadioGroup2Click }

procedure TDllMainForm.RadioGroup3Click(Sender: TObject);
begin { RadioGroup3Click }
	case RadioGroup3.ItemIndex of
		0: VaComm1.Parity := paNone;
  	1: VaComm1.Parity := paEven;
    2: VaComm1.Parity := paOdd;
	  3: VaComm1.Parity := paMark;
  	4: VaComm1.Parity := paSpace;
	end;
end;  { RadioGroup3Click }

procedure TDllMainForm.TimerPoolingTimer(Sender: TObject);
var
	_I: Integer;
begin { TimerPoolingTimer }
  TimerPooling.Enabled := False;
	if not TimerTimeOutResponse.Enabled then
	begin
	  if mPtrOut2 <> mPtrOut1 then
  	begin
			if Copy(mBufferOut[mPtrOut2],1,2) = cSETDATETIME then
  	  begin
    		_fSetDateTime(mBufferOut[mPtrOut2]);
        TimerTimeOutResponse.Enabled := True;
				mPtrOut2 := mPtrOut2 + 1;
	      if mPtrOut2 > cMAXSTRBUFFER then mPtrOut2 := 1;
  	  end
    	else
	    begin
				if Copy(mBufferOut[mPtrOut2],1,2) = cSENDMESSAGE then
	  	  begin
					_fSendMessage(mBufferOut[mPtrOut2]);
          TimerTimeOutResponse.Enabled := True;
					mPtrOut2 := mPtrOut2 + 1;
		      if mPtrOut2 > cMAXSTRBUFFER then mPtrOut2 := 1;
			  end
 				else
   			begin
  				if Copy(mBufferOut[mPtrOut2],1,2) = cSETPARAMETERS then
	  	    begin
		  			_fSetParameters(mBufferOut[mPtrOut2]);
            TimerTimeOutResponse.Enabled := True;
 			  		mPtrOut2 := mPtrOut2 + 1;
  	      	if mPtrOut2 > cMAXSTRBUFFER then mPtrOut2 := 1;
	   		  end
  	   		else
    	   	begin
				  	if Copy(mBufferOut[mPtrOut2],1,2) = cANSWER then
   		    	begin
	  					_fAnswer(mBufferOut[mPtrOut2]);
              TimerTimeOutResponse.Enabled := True;
		   				mPtrOut2 := mPtrOut2 + 1;
			        if mPtrOut2 > cMAXSTRBUFFER then mPtrOut2 := 1;
    		    end;
   	    	end;
	    	end;
  	  end;
	  end
	  else
		begin
	  	for _I := 1 to cMAXTERMINAL do
 			begin
				if Copy(mTerminals[mActualTerminal],4,1) = cENABLED then
  	  	begin
					_fQuestion(mTerminals[mActualTerminal]);
          TimerTimeOutResponse.Enabled := True;
	   			Break;
		    end;
			end;
		  repeat
			  mActualTerminal := mActualTerminal + 1;
				if Copy(mTerminals[mActualTerminal],4,1) = cENABLED then break;
	  	until mActualTerminal > cMAXTERMINAL;
			if mActualTerminal = cMAXTERMINAL + 1 then mActualTerminal := 1;
    end;
  end;
	TimerPooling.Enabled := True;
end;  { TimerPoolingTimer }

procedure TDllMainForm.TimerTimeOutResponseTimer(Sender: TObject);
begin { TimerTerminalTimeOutResponseTimer }
	TimerTimeOutResponse.Enabled := False;
	mErrorOnComunication := NoResponseError+','+mActualCommandAndTerminal;
end;  { TimerTerminalTimeOutResponseTimer }

procedure TDllMainForm.VaComm1RxFlag(Sender: TObject);
var
  _mS,_mS1 : string;
begin { VaComm1RxFlag }
	TimerTimeOutResponse.Enabled := False;
	_mS := VaComm1.ReadText;
	_mS1 := _fCalculateCRCString(_mS);
  if Copy(_mS,Length(_mS)-4,4) = Copy(_mS1,Length(_mS1)-4,4) then
 	begin
		if (Copy(_mS,2,2) = cQUESTION) and (Length(_mS) = 21) then
    begin
      if (Copy(_mS,16,1) = 'E') or (Copy(_mS,16,1) = 'S') then
      begin
       	mBufferIn[mPtrIn1] := Copy(_mS,4,2)+','+Copy(_mS,6,10)+','+Copy(_ms,16,1);
        mPtrIn1 := mPtrIn1 + 1;
        if mPtrIn1 > cMAXSTRBUFFER then mPtrIn1 := 1;
      end;
    end
    else
    begin
			if (Copy(_mS,2,2) = cANSWER) and (Copy(_mS,16,1) = 'N') then
      begin
				mErrorOnComunication := TerminalTimeOutError+','+mActualCommandAndTerminal;
      end;
    end;
  end
	else
  begin
		mErrorOnComunication := CRCError+','+mActualCommandAndTerminal;
  end;
end; { VaComm1RxFlag }

//MWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMW
// Funcoes e procedimentos de uso geral
//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM

function _fCalculateCRCString(_vStr: string): string;
var
  _mCRCByte: Word;
  __mPtrMsg: Integer;
	function _fCalculateCRCByte(_vActualCRC: Word; _vB: Byte): Word;
	const
		_cPOLYNOMIAL: Word = $8404;
	var
		_mI: Word;
	begin	{ _fCalculateCRCByte }
 		_vActualCRC := _vB xor _vActualCRC;
	  for _mI := 1 to 8 do
 		begin
	   	if (_vActualCRC and 1) = 1 then _vActualCRC := (_vActualCRC shr 1) xor _cPOLYNOMIAL
  	  else _vActualCRC := _vActualCRC shr 1;
	 	end;
  	Result := _vActualCRC;
	end;	{ _CalculateCRCByte }
	function _fWordToHex(_vInValue: Word): string;
	const
	 	_cHEXTABLE: array [0..15] of Char = ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');
	var
		_mOutValue : string;
	begin	{ _fWordToHex }
		_mOutValue := _cHEXTABLE [Hi(_vInValue) div 16];
		_mOutValue := _mOutValue + _cHEXTABLE [Hi(_vInValue) mod 16];
		_mOutValue := _mOutValue + _cHEXTABLE [Lo(_vInValue) div 16];
		_mOutValue := _mOutValue + _cHEXTABLE [Lo(_vInValue) mod 16];
	 	Result := _mOutValue;
	end;	{ _fWordToHex }
begin	{ _fCalculateCRCString }
	_mCRCByte := 0;
  for __mPtrMsg := 2 to Length(_vStr)-5 do _mCRCByte := _fCalculateCRCByte(_mCRCByte,Ord(_vStr[__mPtrMsg]));
  Result := Copy(_vStr,1,Length(_vStr)-5)+_fWordToHex(_mCRCByte)+Copy(_vStr,Length(_vStr),1);
end;	{ _fCalculateCRCString }

//MWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMW
// Funcoes e procedimentos internos
//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM

function _fSendStr(_vStr: string): LongInt;
begin { _fSendStr }
  mActualCommandAndTerminal := Copy(_vStr,2,2)+','+Copy(_vStr,4,2);
  if DllMainForm.VaComm1.WriteText(_vStr) then
	  Result := 1
  else
  	Result := 0;
end; { _fSendStr }

function _fSetDateTime(_vStr: string): LongInt;
var
	mDados: string;
begin { _fSetDate }
  DllMainForm.TimerPooling.Interval := 1000;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
	mDados := cSTX+cSETDATETIME+Copy(_vStr,4,2)+Copy(_vStr,7,2)+Copy(_vStr,10,2)+Copy(_vStr,13,4)+Copy(_vStr,18,2)+Copy(_vStr,21,2)+cCRC+cETX;
 	mDados := _fCalculateCRCString(mDados);
 	Result := _fSendStr(mDados);
end; { _fSetDate }

function _fSetParameters(_vStr: string): LongInt;
var
	mDados: string;
begin { _fSetParameters }
  DllMainForm.TimerPooling.Interval := mIntervalTime;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
  mDados := cSTX+cSETPARAMETERS+Copy(_vStr,4,2)+Copy(_vStr,7,2)+Copy(_vStr,10,2)+cCRC+cETX;
 	mDados := _fCalculateCRCString(mDados);
 	Result := _fSendStr(mDados);
end; { _fSetParameters }

function _fQuestion(_vStr: string): LongInt;
var
	mDados: string;
begin { _fQuestion }
  DllMainForm.TimerPooling.Interval := mIntervalTime;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
  mDados := cSTX+cQUESTION+Copy(_vStr,1,2)+cCRC+cETX;
 	mDados := _fCalculateCRCString(mDados);
 	Result := _fSendStr(mDados);
end; { _fQuestion }

//          1         2         3         4         5
// 12345678901234567890123456789012345678901234567890123456
// ..,NT,XXXXXXXXXX,P,S,XX,XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
function _fAnswer(_vStr: string): LongInt;
var
	mDados: string;
begin { _fAnswer }
  DllMainForm.TimerPooling.Interval := mIntervalTime;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
  if Length(_vStr) = 20 then
    mDados := cSTX+cANSWER+Copy(_vStr,4,2)+Copy(_vStr,7,10)+Copy(_vStr,18,1)+Copy(_vStr,20,1)+cCRC+cETX
  else
    mDados := cSTX+cANSWER+Copy(_vStr,4,2)+Copy(_vStr,7,10)+Copy(_vStr,18,1)+Copy(_vStr,20,1)+Copy(_vStr,22,2)+Copy(_vStr,25,32)+cCRC+cETX;
  mDados := _fCalculateCRCString(mDados);
 	Result:= _fSendStr(mDados);
end; { _fAnswer }

function _fSendMessage(_vStr: string): LongInt;
var
	mDados: string;
begin { _fSendMessage }
  DllMainForm.TimerPooling.Interval := mIntervalTime;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
	mDados := cSTX+cSENDMESSAGE+Copy(_vStr,4,2)+Copy(_vStr,7,2)+Copy(_vStr,10,32)+cCRC+cETX;
  mDados := _fCalculateCRCString(mDados);
 	Result:= _fSendStr(mDados);
end; { _fSendMessage }

end.

