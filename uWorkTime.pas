unit uWorkTime;

interface

uses
  Classes;

type
  TWorkTime = class(TThread)
  private
    FstartTime : TDateTime;
    { Private declarations }
    procedure GetWorkTime();
  protected
    procedure Execute; override;
  public
    property StartTime : TDateTime read FstartTime write FstartTime;
  end;

implementation

uses SysUtils, Unit1;

{ WorkTime }

procedure TWorkTime.Execute;
begin
  repeat
    synchronize(GetWorkTime);
    Sleep(1000);
  Until 1 < 0;
  { Place thread code here }
end;

procedure TWorkTime.GetWorkTime;
var
  DeltaTime:TDateTime;
begin
  DeltaTime := Now - startTime;
  Form1.Label7.Caption := Format('������� ���: %s',[TimeToStr(DeltaTime)]);
end;

end.
