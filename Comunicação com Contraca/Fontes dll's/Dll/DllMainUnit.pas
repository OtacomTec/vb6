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
    PBJustOne1: TPBJustOne;
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
  cMAXTERMINAL = 32;
  cMAXSTRBUFFER = 1000;
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
  mTerminals: array [1..cMAXTERMINAL] of string;
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

 	procedure ActiveDll; stdcall;
	procedure DeactiveDll; stdcall;
	function  InsertTerminal (Terminal: LongInt): LongInt; stdcall;													// Registra uma catraca na lista de pooling
	function  DeleteTerminal (Terminal: LongInt): LongInt; stdcall;													// Retira uma catraca da lista de pooling
	function  EnableTerminal (Terminal: LongInt): LongInt; stdcall;													// Liga o pooling para uma catraca
	function  DisableTerminal (Terminal: LongInt): LongInt; stdcall;												// Desliga o pooling para uma catraca
	procedure SetPoolingIntervalTime(IntervalTime: LongInt); stdcall;												// Tempo do ciclo de execução do pooling
	procedure SetTerminalResponseTime(Time: LongInt); stdcall;															// Tempo de espera para completar comunicação
	procedure StartPooling; stdcall;
  procedure StopPooling; stdcall;
	procedure	SetComm(CommPort: LongInt); stdcall;																					// Escolhe a porta serial
	procedure	SetBaudRate(BaudRate: LongInt); stdcall;																			// Escolhe a velocidade de comunicação
	procedure SetCommShow; stdcall;																													// Mostra a form de propriedades e inicializa a porta serial
	function	OpenComm: LongInt; stdcall;																										// Abre a porta serial
	procedure CloseComm; stdcall;																														// Fecha a porta serial
	procedure SetDateTime (Terminal: LongInt; CurrentDateTime: PChar); stdcall;							// Atualiza data e hora da catraca
	procedure SetTerminalTimeOut (Terminal, TimeOut: LongInt); stdcall;											// Envia a catraca o tempo de espera de uma resposta
  procedure SetConditionAfterTimeOut (Terminal, Condition: LongInt); stdcall;							// Envia a catraca a resolução se decorrido o timeout
  procedure SendMessage (Terminal, TimeMessage: LongInt; PersonalMessage: PChar); stdcall;// Envia uma mensagem a catraca
  procedure SetParameters (Terminal: LongInt; Variable: LongInt; Value: LongInt); stdcall // Envia parâmetro a catraca

	function  Question: PChar; stdcall;																											// Verifica se há crachá para tratamento
                                                                                          // Informa condição do crachá
  procedure	Answer(Terminal: LongInt; Badge, Position, Status: PChar; TimeMessage: LongInt; PersonalMessage: PChar); stdcall;

 	function  _fCalculateCRCString(_vStr: string): string; stdcall;													// Caucula o crc da mensagem
	function  _fSendStr(_vStr: string): LongInt; stdcall;										// Envia uma string a catraca
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

procedure ActiveDll;
begin { ActiveDll }
  try
		DllMainForm := TDllMainForm.Create(Application);
  except on E: Exception do
		MessageBox(0,PChar('Erro na dll: '+E.message),PChar('Erro:'),MB_ICONERROR+MB_OK);
  end;
  try
	  AssignFile(mF, 'online.log');
  	Rewrite(mF);
	  except on Exception do
		begin
			MessageBox(0,PChar('Erro na abertura do arquivo de log. Impossível continuar!'),PChar('Erro:'),MB_ICONERROR+MB_OK);
	   	DllMainForm.Free;
  	end;
  end;
end; { ActiveDll }

procedure DeactiveDll;
begin { DeactiveDll }
  CloseFile(mF);
	DllMainForm.Free;
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
			mTerminals[mPosition] := Format('%.2d,1',[Terminal]);
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
  mI 		 := 1;
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

procedure SetPoolingIntervalTime(IntervalTime: LongInt);
begin { SetPoolingIntervalTime }
	mIntervalTime := IntervalTime;
	DllMainForm.TimerPooling.Interval := mIntervalTime;
end;  { SetPoolingIntervalTime }

procedure SetTerminalResponseTime(Time: LongInt);
begin  { SetTimeTerminalResponse }
	mTerminalResponseTime := Time;
  DllMainForm.TimerTimeOutResponse.Interval := mTerminalResponseTime;
end;  { SetTimeTerminalResponse }

procedure StartPooling;
begin { StartPooling }
	DllMainForm.TimerPooling.Enabled := True;
end;  { StartPooling }

procedure StopPooling;
begin { StopPooling }
	DllMainForm.TimerPooling.Enabled := False;
end;  { StopPooling }

procedure SetComm(CommPort: LongInt); // Revisada em 04/01...
begin { SetComm }
	DllMainForm.VaComm1.PortNum := CommPort;
end; { SetComm }

procedure SetBaudRate(BaudRate: LongInt);
begin { SetBaudRate }
  case BaudRate of
    110: DllMainForm.VaComm1.BaudRate := br110;
		300: DllMainForm.VaComm1.BaudRate := br300;
		600: DllMainForm.VaComm1.BaudRate := br600;
		1200:	DllMainForm.VaComm1.BaudRate := br1200;
		2400:	DllMainForm.VaComm1.BaudRate := br2400;
		4800:	DllMainForm.VaComm1.BaudRate := br4800;
		9600:	DllMainForm.VaComm1.BaudRate := br9600;
		14400: DllMainForm.VaComm1.BaudRate := br14400;
		19200: DllMainForm.VaComm1.BaudRate := br19200;
    38400: DllMainForm.VaComm1.BaudRate := br38400;
		56000: DllMainForm.VaComm1.BaudRate := br56000;
		57600: DllMainForm.VaComm1.BaudRate := br57600;
		115200:	DllMainForm.VaComm1.BaudRate := br115200;
		128000:	DllMainForm.VaComm1.BaudRate := br128000;
		256000:	DllMainForm.VaComm1.BaudRate := br256000;
    else
    	DllMainForm.VaComm1.BaudRate := br4800;
	end;
end; { SetBaudRate }

procedure SetCommShow;
begin { SetCommShow }
  DllMainForm.Visible := True;
end; { SetCommShow }

function OpenComm: LongInt;
begin { OpenComm }
	try
  	DllMainForm.VaComm1.Open;
		Result := 1;
	  except on exception do Result := 0;
  end;
end; { OpenComm }

procedure CloseComm;
begin { CloseComm }
	DllMainForm.TimerPooling.Enabled := False;
	DllMainForm.VaComm1.Close;
end; { CloseComm }

procedure SetDateTime(Terminal: LongInt; CurrentDateTime: PChar);
begin { SetDateTime }
  mBufferOut[mPtrOut1] := cSETDATETIME+Format(',%.2d,',[Terminal])+CurrentDateTime;
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
end;  { SetDateTime }

procedure SetTerminalTimeOut(Terminal, TimeOut: LongInt);
begin { SetTerminalTimeOut }
  mBufferOut[mPtrOut1] := cSETPARAMETERS+Format(',%.2d,',[Terminal])+cTERMINALTIMEOUT+Format(',%.2d',[TimeOut div 100]);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
end;  { SetTerminalTimeOut }

procedure SetConditionAfterTimeOut (Terminal, Condition: LongInt);
begin { SetConditionAfterTimeOut }
  mBufferOut[mPtrOut1] := cSETPARAMETERS+Format(',%.2d,',[Terminal])+cCONDITIONAFTERTIMEOUT+Format(',%.2d',[Condition]);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
end; { SetConditionAfterTimeOut }

procedure	SendMessage(Terminal, TimeMessage: LongInt; PersonalMessage: PChar);
begin { SendMessage }
  mBufferOut[mPtrOut1] := cSENDMESSAGE+Format(',%.2d',[Terminal])+Format(',%.2d,',[TimeMessage div 100])+PersonalMessage;
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
end;  { SendMessage }

procedure SetParameters(Terminal: LongInt; Variable: LongInt; Value: LongInt);
begin { SetParameters }
  mBufferOut[mPtrOut1] := cSETPARAMETERS+Format(',%.2d',[Terminal])+Format(',%.2d',[Variable])+Format(',%.2d',[Value]);
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
end;  { SetParameters }

function Question: PChar;
begin { Question }
	if mPtrIn1 <> mPtrIn2 then
  begin
		Result := PChar(mBufferIn[mPtrIn2]);
    mPtrIn2 := mPtrIn2 + 1;
    if mPtrIn2 > cMAXSTRBUFFER then mPtrIn2 := 1;
  end
  else
  begin
		Result := PChar('');
  end;
end;  { Question }

procedure	Answer(Terminal: LongInt; Badge, Position, Status: PChar; TimeMessage: LongInt; PersonalMessage: PChar);
begin { Answer }
	if StrPas(PersonalMessage) = '' then mBufferOut[mPtrOut1] := cANSWER+Format(',%.2d,',[Terminal])+Badge+','+Position+','+Status+''
  else mBufferOut[mPtrOut1] := cANSWER+Format(',%.2d,',[Terminal])+Badge+','+Position+','+Status+Format(',%.2d,',[TimeMessage div 100])+PersonalMessage;
  mPtrOut1 := mPtrOut1 + 1;
  if mPtrOut1 > cMAXSTRBUFFER then mPtrOut1 := 1;
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
//  WriteLn(mF,FormatDateTime('dd/mm/yyyy,hh:mm:ss,',Now)+mErrorOnComunication);
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
			  WriteLn(mF,FormatDateTime('dd/mm/yyyy,hh:mm:ss,',Now)+mErrorOnComunication);
      end;
    end;
  end
	else
  begin
		mErrorOnComunication := CRCError+','+mActualCommandAndTerminal;
	  WriteLn(mF,FormatDateTime('dd/mm/yyyy,hh:mm:ss,',Now)+mErrorOnComunication);
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

