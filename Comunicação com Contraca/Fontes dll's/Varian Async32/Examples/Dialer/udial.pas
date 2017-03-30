unit udial;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  VaClasses, VaSource, VaComm, VaSystem, StdCtrls, ExtCtrls, VaModem;

type
  TForm1 = class(TForm)
    VaComm: TVaComm;
    VaDataSource: TVaDataSource;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    Edit3: TEdit;
    DialButton: TButton;
    HangupButton: TButton;
    Label4: TLabel;
    Memo1: TMemo;
    Bevel1: TBevel;
    VaModem1: TVaModem;
    Label5: TLabel;
    Memo2: TMemo;
    procedure VaCommRxBuf(Sender: TObject; const Buf: PChar;
      Count: Integer);
    procedure HangupButtonClick(Sender: TObject);
    procedure DialButtonClick(Sender: TObject);
    procedure VaModem1Timeout(Sender: TObject);
    procedure VaModem1Response(Sender: TObject; Event: TVaModemEventType);
    procedure FormCreate(Sender: TObject);
  private
    Retries: Integer;
    Cancel: Boolean;
    Connected: Boolean;
    MemoIndex: Integer;
    procedure Dial;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  VaUtils;

{$R *.DFM}

procedure TForm1.FormCreate(Sender: TObject);
begin
  MemoIndex := Memo2.Lines.Add('');
end;

procedure TForm1.VaCommRxBuf(Sender: TObject; const Buf: PChar;
  Count: Integer);
var
  I: Integer;
  Tmp: string;
begin
  Tmp := Buf;
  for I := 1 to Length(Tmp) do
    case Tmp[I] of
      #10:;
      #13: MemoIndex := Memo2.Lines.Add('');
      else
        Memo2.Lines[MemoIndex] := Memo2.Lines[MemoIndex] + Tmp[I];
    end;
end;

procedure TForm1.HangupButtonClick(Sender: TObject);
begin
  Cancel := True;
  VaModem1.Cancel; //abort pending process
  VaComm.SetDTRState(false); //drop DTR
  Memo1.Lines.Add('HANGUP');
end;

procedure TForm1.Dial;
begin
  if not Cancel then
  begin
    Inc(Retries);
    if Retries > StrToInt(Edit2.Text) then
      Memo1.Lines.Add('Max. retries reached.')
    else
    begin
      if Retries > 1 then //not first attempt
        SysDelay(StrToInt(Edit3.Text) * 1000, True);
      VaModem1.Dial(Edit1.Text);
    end;
  end;
end;

procedure TForm1.DialButtonClick(Sender: TObject);
begin
  Retries := 0;
  Cancel := false;
  Dial;
end;

procedure TForm1.VaModem1Timeout(Sender: TObject);
begin
  Memo1.Lines.Add('TIMEOUT');
  Dial;
end;

procedure TForm1.VaModem1Response(Sender: TObject;
  Event: TVaModemEventType);
begin
  case Event of
    metOK:
      Memo1.Lines.Add('OK');
    metConnect:
      begin
        Connected := True;
        Memo1.Lines.Add('CONNECT');
      end;
    metBusy:
      Memo1.Lines.Add('BUSY');
    metVoice:
      Memo1.Lines.Add('VOICE');
    metNoCarrier:
      begin
        Memo1.Lines.Add('NO CARRIER');
        if Connected then
        begin
          Connected := false;
          Exit;
        end;
      end;
    metNoDialTone:
      Memo1.Lines.Add('NO DIALTONE');
    metError:
      Memo1.Lines.Add('ERROR');
  end;

  if (Event <> metOK) and
     (Event <> metConnect) then Dial; //redial
end;


end.
