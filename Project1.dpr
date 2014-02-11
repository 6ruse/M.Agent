program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  uLicense in 'uLicense.pas' {frmLicense},
  utils in 'utils.pas',
  uEncrypt in 'uEncrypt.pas',
  uAbout in 'uAbout.pas' {FrmAbout},
  uSettings in 'uSettings.pas' {FrmSettings},
  uWorkTime in 'uWorkTime.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TFrmSettings, FrmSettings);
  Application.Run;
end.
