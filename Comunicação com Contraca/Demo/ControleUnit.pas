unit ControleUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, DBCtrls, Grids, DBGrids, DB, DBTables, StdCtrls, Buttons;

type
  TControleForm = class(TForm)
    TableCracha: TTable;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    StartPoolingButton: TButton;
    Add5000SpeedButton: TSpeedButton;
    Add100SpeedButton: TSpeedButton;
    Timer1: TTimer;
    Memo1: TMemo;
    TableCrachaNmerodoCrach: TStringField;
    TableCrachaCatracaPermitida: TIntegerField;
    TableCrachaHabilitado: TBooleanField;
    StopPoolingButton: TButton;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Add5000SpeedButtonClick(Sender: TObject);
    procedure StartPoolingButtonClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure StopPoolingButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ControleForm: TControleForm;

implementation

{$R *.DFM}

procedure ActiveDll; stdcall; external 'Online.dll';
procedure DeactiveDll; stdcall; external 'Online.dll';
function  InsertTerminal (Terminal: LongInt): LongInt; stdcall; external 'Online.dll';
function  DeleteTerminal (Terminal: LongInt): LongInt; stdcall; external 'Online.dll';
function  EnableTerminal (Terminal: LongInt): LongInt; stdcall; external 'Online.dll';
function  DisableTerminal (Terminal: LongInt): LongInt; stdcall; external 'Online.dll';
procedure SetPoolingIntervalTime(IntervalTime: LongInt); stdcall; external 'Online.dll';
procedure SetTerminalResponseTime(Time: LongInt); stdcall; external 'Online.dll';
procedure StartPooling; stdcall; external 'Online.dll';
procedure StopPooling; stdcall; external 'Online.dll';
procedure	SetComm(CommPort: LongInt); stdcall; external 'Online.dll';
procedure	SetBaudRate(BaudRate: LongInt); stdcall; external 'Online.dll';
procedure SetCommShow; stdcall; external 'Online.dll';
function	OpenComm: LongInt; stdcall; external 'Online.dll';
procedure CloseComm; stdcall; external 'Online.dll';
procedure SetDateTime (Terminal: LongInt; CurrentDateTime: PChar); stdcall; external 'Online.dll';
procedure SetTerminalTimeOut (Terminal, TimeOut: LongInt); stdcall; external 'Online.dll';
procedure SetConditionAfterTimeOut (Terminal, Condition: LongInt); stdcall; external 'Online.dll';
procedure SendMessage (Terminal, TimeMessage: LongInt; PersonalMessage: PChar); stdcall; external 'Online.dll';
function  Question: PChar; stdcall; external 'Online.dll';
procedure	Answer(Terminal: LongInt; Badge, Position, Status: PChar; TimeMessage: LongInt; PersonalMessage: PChar); stdcall; external 'Online.dll';

procedure TControleForm.FormActivate(Sender: TObject);
begin
  TableCracha.Open;
 	ActiveDll; // ativa a DLL
	SetComm(2); // configura a porta serial e a velocidade de comunicação
	SetBaudRate(4800);
  OpenComm; // abre a porta serial
	SetPoolingIntervalTime(100); // configura o intervalo do pooling = 50ms (milisegundos)
  SetTerminalResponseTime(500); // configura o tempo de aguardo da resposta pelo computador

 	InsertTerminal(1); // insere o terminal 1
	DisableTerminal(1); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(2); // insere o terminal 1
//	DisableTerminal(2); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(3); // insere o terminal 1
//	DisableTerminal(3); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(4); // insere o terminal 1
//	DisableTerminal(4); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(5); // insere o terminal 1
//	DisableTerminal(5); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(6); // insere o terminal 1
//	DisableTerminal(6); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(7); // insere o terminal 1
//	DisableTerminal(7); // desabilita momentaneamente o terminal recentemente inserido
// 	InsertTerminal(8); // insere o terminal 1
//	DisableTerminal(8); // desabilita momentaneamente o terminal recentemente inserido
	EnableTerminal(1); // habilita o terminal
//	EnableTerminal(2); // habilita o terminal
//	EnableTerminal(3); // habilita o terminal
//	EnableTerminal(4); // habilita o terminal
//	EnableTerminal(5); // habilita o terminal
//	EnableTerminal(6); // habilita o terminal
//	EnableTerminal(7); // habilita o terminal
//	EnableTerminal(8); // habilita o terminal

  Randomize;
  Label2.Caption := IntToStr(TableCracha.RecordCount);
end;

procedure TControleForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TableCracha.Close;
	CloseComm;
 	DeactiveDll;
end;

procedure TControleForm.Add5000SpeedButtonClick(Sender: TObject);
var
  mI,mQuantidade: Integer;
begin
  TableCracha.DisableControls;
	if Sender = Add5000SpeedButton then
  	mQuantidade := 5000
  else
  	mQuantidade := 100;
  for mI := 1 to mQuantidade do
  begin
    TableCracha.Append;
    TableCrachaNMERODOCRACH.Value := Format('%.10d',[Random(1000000)]);
    TableCrachaCATRACAPERMITIDA.Value := Random(6);
    TableCrachaHABILITADO.Value := Random(2) = Random(2);
		try
	    TableCracha.Post;
  	  except on exception do
		  TableCracha.Cancel;
    end;
  end;
  TableCracha.EnableControls;
  Label2.Caption := IntToStr(TableCracha.RecordCount);
end;

procedure TControleForm.StartPoolingButtonClick(Sender: TObject);
begin
  SetTerminalTimeOut(1,2000); // configura o tempo de aguardo da resposta do terminal
	SetConditionAfterTimeOut(1,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
  SetDateTime(1,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
  SendMessage(1,1000,PChar('Terminal 1      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling

//  SetTerminalTimeOut(2,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(2,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(2,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(2,1000,PChar('Terminal 2      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling

//  SetTerminalTimeOut(3,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(3,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(3,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(3,1000,PChar('Terminal 3      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling
//  SetTerminalTimeOut(4,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(4,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(4,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(4,1000,PChar('Terminal 4      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling
//  SetTerminalTimeOut(5,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(5,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(5,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(5,1000,PChar('Terminal 5      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling
//  SetTerminalTimeOut(6,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(6,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(6,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(6,1000,PChar('Terminal 6     inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling
//  SetTerminalTimeOut(7,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(7,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(7,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(7,1000,PChar('Terminal 7      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling
//  SetTerminalTimeOut(8,2000); // configura o tempo de aguardo da resposta do terminal
//	SetConditionAfterTimeOut(8,0); // configura o procedimento a ser tomado pelo terminal após decorrido o tempo de aguardo
//  SetDateTime(8,PChar(FormatDateTime('dd/mm/yyyy hh:nn:ss',Now))); // configura o calendário e o relógio do terminal
//  SendMessage(8,1000,PChar('Terminal 8      inicializado... ')); // envia uma mensagem ao terminal afim de confirmar sua participaçãp no pooling

	Timer1.Enabled := True;
	StartPooling;
end;

procedure TControleForm.StopPoolingButtonClick(Sender: TObject);
begin
	Timer1.Enabled := False;
	StopPooling;
end;

procedure TControleForm.Timer1Timer(Sender: TObject);
var
	mCatraca: Integer;
  mAA, mDado, mEstado, mCracha, mMsg: string;
begin
  Timer1.Enabled := False;
	mDado := StrPas(Question);
  if mDado <> '' then
  begin
	  mCatraca := StrToInt(Copy(mDado,1,2));
  	mCracha  := Copy(mDado,4,10);
    if (TableCracha.FindKey([mCracha])) and (TableCrachaHABILITADO.Value) then
    begin
    	mEstado := 'S';
 	    mMsg    := 'Liberado';
			mAA := Copy(mDado,1,2)+' '+Copy(mDado,4,10)+'  '+Copy(mDado,15,1)+'Liberado        ';
    end
    else
    begin
    	mEstado := 'N';
			mMsg    := 'Não liberado';
			mAA := Copy(mDado,1,2)+' '+Copy(mDado,4,10)+'  '+Copy(mDado,15,1)+'Bloqueado       ';
    end;
    Answer(mCatraca, PChar(mCracha), PChar(Copy(mDado,15,1)), PChar(mEstado), 2000, PChar(mAA));
    Memo1.Lines.Add('Catraca:'+Copy(mDado,1,2)+' Crachá:'+ Copy(mDado,4,10)+'Sentido: '+Copy(mDado,15,1)+' -> '+mMsg);
  end;
  Timer1.Enabled := True;
end;

end.
