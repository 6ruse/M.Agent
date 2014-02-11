unit uAbout;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, pngimage, utils, ShellAPI;

type
  TFrmAbout = class(TForm)
    Panel1: TPanel;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    OKButton: TButton;
    procedure FormCreate(Sender: TObject);
    procedure CopyrightClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmAbout: TFrmAbout;

implementation

{$R *.dfm}

procedure TFrmAbout.FormCreate(Sender: TObject);
begin
  Version.Caption := Format('Version %s',[GetSelfVersion]);
end;

procedure TFrmAbout.CopyrightClick(Sender: TObject);
begin
  ShellExecute(Application.Handle, nil, 'http://koder.kz/', nil, nil,SW_SHOWNOACTIVATE);
end;

end.
 
