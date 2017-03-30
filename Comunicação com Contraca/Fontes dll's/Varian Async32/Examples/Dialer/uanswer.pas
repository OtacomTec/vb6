unit uanswer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  VaClasses, VaSource, VaComm, StdCtrls, VaModem;

type
  TForm1 = class(TForm)
    VaComm: TVaComm;
    VaDataSource1: TVaDataSource;
    VaModem1: TVaModem;
    Button1: TButton;
    Memo1: TMemo;
    Button2: TButton;
    Memo2: TMemo;
    procedure VaCommRxBuf(Sender: TObject; const Buf: PChar;
      Count: Integer);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure VaModem1Answer(Sender: TObject);
    procedure VaModem1Ring(Sender: TObject; var AcceptCall: Boolean);
    procedure FormCreate(Sender: TObject);
  private
    MemoIndex: Integer;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.FormCreate(Sender: TObject);
begin
  MemoIndex := Memo1.Lines.Add('');
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
      #13: MemoIndex := Memo1.Lines.Add('');
      else
        Memo1.Lines[MemoIndex] := Memo1.Lines[MemoIndex] + Tmp[I];
    end;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  VaModem1.WaitCall(5);
  Memo2.Lines.Add('Waiting for call...');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  VaModem1.Cancel;
  VaComm.SetDTRState(false); //hangup
  VaModem1.Hangup;
end;

procedure TForm1.VaModem1Answer(Sender: TObject);
begin
  Memo2.Lines.Add('Answering incoming call...');
end;

procedure TForm1.VaModem1Ring(Sender: TObject; var AcceptCall: Boolean);
begin
  Memo2.Lines.Add('RING: ' + IntToStr(VaModem1.RingCount));
end;


end.
