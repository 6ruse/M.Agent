unit uLicense;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, PropFilerEh, PropStorageEh, IdComponent,
  IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdBaseComponent,
  IdMessage, IdHTTP, Registry, idSNTP, IdGlobal, pngimage, ExtCtrls;

type
  TfrmLicense = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    EdtEmail: TEdit;
    PropStorageEh1: TPropStorageEh;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    EdtName: TEdit;
    EdtPhone: TEdit;
    IdHTTP21: TIdHTTP;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    EdtLicense: TEdit;
    BtnActivate: TButton;
    BtnSelectLicense: TButton;
    Panel3: TPanel;
    Image1: TImage;
    Label6: TLabel;
    GroupBox3: TGroupBox;
    EdtLicNumber: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    EdtEmailkoder: TEdit;
    Label9: TLabel;
    procedure BtnSelectLicenseClick(Sender: TObject);
    procedure BtnActivateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    FidComp      : string;
    Fversion     : string;
    FDWork       : TDate; // �� ������ ����� ��������
    function  GetHdd(): string ;
    function  GetLeftSerialNumber() : String;
    function  GetVersion() : string;
    function  GetDWork() : TDate;
    function  GetLocalHost() : Boolean;

    property  IdComp : string read FidComp write FidComp;
    property  Version : string read FVersion write FVersion;
    property  DWork : TDate read FDWork write FDWork;
  public
    { Public declarations }
    function  LicenseActivate(): Boolean;
  end;

var
  frmLicense: TfrmLicense;

implementation

uses Unit1, uEncrypt, utils;

{$R *.dfm}

{ TfrmLicense }

function TfrmLicense.GetHdd: string;
var
  VolumeSerialNumber : DWORD;
  MaximumComponentLength : DWORD;
  FileSystemFlags : DWORD;
begin
  GetVolumeInformation('C:\',  nil, 0, @VolumeSerialNumber,MaximumComponentLength, FileSystemFlags,  nil,  0);
  result := IntToHex(HiWord(VolumeSerialNumber), 4)+IntToHex(LoWord(VolumeSerialNumber), 4);
end;

function TfrmLicense.GetLeftSerialNumber: String;
  var
    Hdd : string ;
    s,s1,s2 : string;
    _ecx,_eax,_edx : longword;
begin
  asm
    mov eax,1
    db $0F,$A2
    mov _eax,eax
  end;
  asm
    mov eax,3
    db $0F,$A2
    mov _ecx,ecx
    mov _edx,edx
  end;

  s := IntToHex(_eax, 8);
  s1 := IntToHex(_edx, 8);
  s2 := IntToHex(_ecx, 8);

  Hdd := GetHdd;
  Result := Format('%s%s%s%s',[Hdd, s, s1, s2]);
end;

procedure TfrmLicense.BtnSelectLicenseClick(Sender: TObject);
  var
    TextBody : string;
    License  : string;
begin
  if not IsConnectedToInternet then
  begin
    Application.MessageBox('�� ���� ����������� � ���������, ��� �� koder.kz ����� �������','Error', MB_ICONERROR);
    exit;
  end;

  License  := GetLeftSerialNumber;
  TextBody := format('key = %s from %s, E-mail %s, phone %s',
                                 [License, EdtName.Text,
                                  EdtEmail.Text, EdtPhone.Text]);
  if EdtEmail.Text = '' then
  begin
    MessageBox(Handle,'���� E-Mail ����������� ��� ����������, ��� :(','Error',MB_ICONERROR);
    exit ; 
  end;

  if not SendMsgMail('smtp.mail.ru', 'BootMailRu buy', '6ruse_as@mail.ru', 'wars_angel@mail.ru', TextBody, '', 'wars_angel','pkjytlhtvktn', '25') then // ������ �� �����������
  begin
    ClientHeight := 504;
    EdtLicNumber.Text := License ;
    //MessageBox(Handle,'���� ������ �� ����������, ���������� �������� ���������','Error',MB_ICONERROR);
  end
  else
  begin
    SendMsgMail('smtp.mail.ru', 'BootMailRu buy', 'bknovikov@gmail.com', 'wars_angel@mail.ru', TextBody, '', 'wars_angel','pkjytlhtvktn', '25');
    MessageBox(Handle,'���� ������ �������, ���������� ���������� �� ��������� ��������� ������ � ��� �� �����','���������',MB_ICONINFORMATION);
    ModalResult := mrCancel;
  end;
end;

function TfrmLicense.GetDWork: TDate;
var
  http: TIdHTTP;
  Page: UTF8String;
  k, l, m: Integer;
  day: string;

  ms: TMemoryStream;
  sl: TStringList;
  dateStr : string;
begin

  http := TIdHTTP.Create(nil);
  ms := TMemoryStream.Create;
  http.Get(CNT_SITE_DATE,ms);
  ms.Position := 0;
  sl := TStringList.Create;
  sl.LoadFromStream(ms);
  ms.free;
  page := UTF8Decode(sl.Text);
  sl.Free;
  dateStr := copy(page, pos('�������:',page)+18,10);
  delete(dateStr, pos(',',dateStr),1);

  result := StrToDateTime(dateStr);
  http.Free;
//  result := now(); Internet
end;

function TfrmLicense.LicenseActivate: Boolean;
  const
    my_key = 1425;

  var
   TmpStr    : string;
   License   : string;
   Dlicense  : TDate;
   MyLicense : string;
   MyVersion,
   ServerVersion : string;
begin
  result := true ;
  try
    MyVersion     := GetSelfVersion;
    ServerVersion := GetVersion;
    //� �� �������� �� ��� ��������?
    if GetLocalHost then
    begin
      Application.MessageBox('��������� ����� ������� ��� ���������� ������ - ��� �������!','Error', MB_ICONERROR);
      result := false;
      Application.Terminate;
      Exit;
    end;
    // � �� �� ��� ������ ���������?
    if ServerVersion <> MyVersion then
    begin
      Application.MessageBox('��������� ��������� � ����������, ��� :(','Error', MB_ICONERROR);
      result := false;
      Exit;
    end;

    // � ����� ��� ������ ������?
    if EdtLicense.Text = '' then
    begin
      result := false;
      exit;
    end;

    // ��������� ��������
    MyLicense := GetLeftSerialNumber;
    DWork   := GetDWork;
    TmpStr  := Decrypt(EdtLicense.Text, my_key);

    //������������, ������ ��������� �� �����
    License := copy(TmpStr,1,pos('-',TmpStr)-1);
    Dlicense := StrToDate(copy(TmpStr,pos('-',TmpStr)+1,Length(TmpStr)));
    if DWork > Dlicense then
    begin
      Application.MessageBox('���� �������� �������� ��������, ��� :(','Error', MB_ICONERROR);
      result := false;
    end;
    if MyLicense <> License then
    begin
      Application.MessageBox('�������� �� ��������, ��� :(','Error', MB_ICONERROR);
      result := false;
    end;
  except on E: Exception do
  begin
    Application.MessageBox('bad ��������, ��� :(','Error', MB_ICONERROR);
    result := false;
  end;
  end;
end;

procedure TfrmLicense.BtnActivateClick(Sender: TObject);
begin
  if EdtLicense.Text = '' then
  begin
    Application.MessageBox('������� ������������ �����, ��� :(','Error', MB_ICONERROR);
    Exit;
  end;
  //��������� ��������
  if LicenseActivate then
    ModalResult := mrOk
  else
    ModalResult := mrCancel;
end;

procedure TfrmLicense.FormCreate(Sender: TObject);
begin
  ModalResult := mrCancel;
  ClientHeight := 358;
end;

procedure TfrmLicense.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then
   modalResult := mrCancel;
end;

function TfrmLicense.GetVersion: string;
var
  http: TIdHTTP;
  Page: string;
  k, l, m: Integer;
  day: string;

  ms: TMemoryStream;
  sl: TStringList;
  dateStr : string;
begin
  // ������ ������� ���������
  http := TIdHTTP.Create(nil);
  ms := TMemoryStream.Create;
  http.Get(CNT_LICENSE,ms);
  ms.Position := 0;
  sl := TStringList.Create;
  sl.LoadFromStream(ms);
  ms.free;
  page := sl.Text;
  sl.Free;
  dateStr := copy(page,1,7);
  result := dateStr;
  http.Free;
end;

function TfrmLicense.GetLocalHost: Boolean;
  var
    WinDir   : array [0..255] of char;
    HostFile :string;
    DataHost : TStringList;
    TmpStr   : string;
    i        : integer;
begin
  // ��������� ��� �� ��������������� �� ��������� ���� � ����� ��� ��� � ���
  result := false;

  GetWindowsDirectory(WinDir, 255);
  HostFile := Format('%s\System32\drivers\etc\hosts',[WinDir]);
  DataHost := TStringList.Create();
  try
    DataHost.LoadFromFile(HostFile);
    for i := 0 to DataHost.Count -1 do
    begin
      TmpStr := DataHost[i];
      if Pos(CNT_SITE_DATE,TmpStr) > 0 then
      begin
        result := true;
        exit;
      end;
    end;
  finally
    DataHost.Free;
  end;
end;

end.
