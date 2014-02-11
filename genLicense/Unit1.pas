unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, uEncrypt, ExtCtrls;

type
  TForm1 = class(TForm)
    EdtKeyKlient: TEdit;
    DWork: TDateTimePicker;
    EdtLicense: TEdit;
    BtnGenerateLicense: TButton;
    Label1: TLabel;
    Label2: TLabel;
    BtnClose: TButton;
    procedure BtnGenerateLicenseClick(Sender: TObject);
    procedure BtnCloseClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
const
 MAgent_key = 1425;
 MyMir_key  = 8898;

implementation

{$R *.dfm}


procedure TForm1.BtnGenerateLicenseClick(Sender: TObject);
  var S : string;
  key   : Integer;
begin
  s := format('%s-%s',[EdtKeyKlient.Text, DateToStr(DWork.Date)]);
  EdtLicense.Text := Encrypt(s, MAgent_key);
//  Edit3.Text := Decrypt(Edit2.Text,my_key);
end;

procedure TForm1.BtnCloseClick(Sender: TObject);
begin
  close
end;

end.
