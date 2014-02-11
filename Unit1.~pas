unit UNIT1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Sockets, client, Buttons, ExtCtrls, ComCtrls,
  Menus, ImgList,Wininet, jpeg, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdHTTP,mmsystem, Spin, ActnList,
  XPStyleActnCtrls, ActnMan, PropStorageEh, IdMessageClient, IdSMTP,
  pngimage, PropFilerEh, CheckLst, ToolWin, uWorkTime;

type
  TForm1 = class(TForm)
    TcpClient: TTcpClient;
    StatusBar1: TStatusBar;
    MailClient: TMailClient;
    PopupMenu1: TPopupMenu;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    ActionManager1: TActionManager;
    actOnline: TAction;
    actSaveLogFile: TAction;
    SaveDialog1: TSaveDialog;
    RegPropStorageManEh1: TRegPropStorageManEh;
    Panel3: TPanel;
    Image1: TImage;
    Label5: TLabel;
    PropStorageEh1: TPropStorageEh;
    actAbout: TAction;
    N2: TMenuItem;
    actClose: TAction;
    actSetting: TAction;
    N3: TMenuItem;
    PropStorageEh2: TPropStorageEh;
    PopupMenu2: TPopupMenu;
    N8: TMenuItem;
    N9: TMenuItem;
    actPause: TAction;
    N13: TMenuItem;
    actConnect: TAction;
    N14: TMenuItem;
    actCancelMessage: TAction;
    N15: TMenuItem;
    actSendMessage: TAction;
    N16: TMenuItem;
    actFindContact: TAction;
    N17: TMenuItem;
    actFindAndSend: TAction;
    N18: TMenuItem;
    GroupBox3: TGroupBox;
    ChckLstKontact: TCheckListBox;
    Panel9: TPanel;
    ChckKontackt: TCheckBox;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    Label4: TLabel;
    Panel4: TPanel;
    GroupBox2: TGroupBox;
    RichEdit1: TRichEdit;
    GroupBox1: TGroupBox;
    RichEdit2: TRichEdit;
    Panel12: TPanel;
    Panel11: TPanel;
    Label8: TLabel;
    CtnHideLog: TButton;
    RedtLog: TRichEdit;
    GroupBox6: TGroupBox;
    Label7: TLabel;
    BtnConnect: TButton;
    BtnOnline: TButton;
    BtnCancelSend: TButton;
    BtnAll: TButton;
    BtnPause: TButton;
    Panel5: TPanel;
    Panel6: TPanel;
    Label2: TLabel;
    Panel7: TPanel;
    Label1: TLabel;
    Panel8: TPanel;
    Label3: TLabel;
    Panel10: TPanel;
    Label6: TLabel;
    ProgressBar1: TProgressBar;
    procedure MailClientConnect(Sender: TObject);
    procedure MailClientHello(Sender: TObject);
    procedure MailClientRequestHost(Sender: TObject);
    procedure MailClientRecievedHost(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RichEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure MailClientMessage(Sender: TObject; MsgID: Cardinal; From,
      Text: String; IsRTF: Boolean);
    procedure MailClientMailBoxStatus(Sender: TObject; Reason: Cardinal);
    procedure MailClientMailBoxStatusNew(Sender: TObject; MsgNum: Cardinal;
      MailSender, Subject: String; TimeStamp: Cardinal);
    procedure RichEdit2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MailClientOfflineMessage(Sender: TObject; MsgID: Cardinal;
      From, Text: String);
    procedure MailClientSystemMessage(Sender: TObject; MsgID: Cardinal;
      From, Text: String);
    procedure MailClientContactBeginPrint(Sender: TObject; MsgID: Cardinal;
      From, Text: String);
    procedure MailClientContactEndPrint(Sender: TObject; MsgID: Cardinal;
      From, Text: String);
    procedure MailClientContact(Sender: TObject; GroupID: Cardinal; EMail,
      Nick: String; Status: TStatusClient; InvisState: TInvisStateClient);
      function READ_status(STATUS:TStatusClient):string;
    procedure ListBox11Click(Sender: TObject);
    Procedure Get_Avy;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure MailClientContactAuthorize(Sender: TObject; MsgID: Cardinal;
      From, Text: String);
    procedure MailClientAddContactError(Sender: TObject; GroupID,
      UserID: Cardinal; EMail, Nick: String; Reason: Cardinal);
    procedure N11Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure actOnlineExecute(Sender: TObject);
    procedure actSaveLogFileExecute(Sender: TObject);
    procedure actAboutExecute(Sender: TObject);
    procedure actCloseExecute(Sender: TObject);
    procedure actSettingExecute(Sender: TObject);
    procedure ChckKontacktClick(Sender: TObject);
    procedure actPauseExecute(Sender: TObject);
    procedure actConnectExecute(Sender: TObject);
    procedure actCancelMessageExecute(Sender: TObject);
    procedure actSendMessageExecute(Sender: TObject);
    procedure actFindContactExecute(Sender: TObject);
    procedure actFindAndSendExecute(Sender: TObject);
    procedure MailClientErrorAuthorize(Sender: TObject; Reason: String);
    procedure CtnHideLogClick(Sender: TObject);
  private
    { Private declarations }
    GetSend  : Boolean;
    GetPause : Boolean;
    TakeMess : Integer;
    StartWork : Boolean;// что бы не добавлял не кого больше в контакты
    function FinKontackt : Integer;
    procedure Restart();//перезапуск в случае чаго

  public
    { Public declarations }
    Login        : string;
    Pass         : string;
    SendSleep    : Integer;
    KeySleep     : Integer;
    SmileCount   : Integer;
    RestartMess  : Integer;
    StopError    : Boolean;
    CheckContact : Boolean;
    CheckSmile   : Boolean;
    CheckSelect  : Boolean;// начинать с указанного контакта
    CheckAutoSav : Boolean;
    ChckInternet : Boolean;
    ChckEnter    : Boolean;
    procedure GetSettings;
  end;

var
  Form1     : TForm1;
  em, dom   : String;
  g         : integer;
  hMutex    : THandle;
  StrtI     : Integer;
  WorkTime  : TWorkTime;
  
 implementation

uses uLicense, utils, uAbout, uSettings;

{$R *.dfm}

procedure RE_AlignText( RichEdit: TRichEdit; Alignment: TAlignment );
begin
  RichEdit.Paragraph.Alignment := Alignment;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  FrmLicense := TFrmLicense.Create(self);
  try
    FrmLicense.Show;
    FrmLicense.Hide;
    if not FrmLicense.LicenseActivate then
      if FrmLicense.ShowModal = mrCancel then
      begin
        application.Terminate;
      end;
  finally
    FrmLicense.Free;
  end;


  if mailclient.Connected then
  begin
    mailclient.Status:=OnLine;
  end;
end;

procedure TForm1.MailClientConnect(Sender: TObject);
begin
  MailClient.Hello;
end;

procedure TForm1.MailClientHello(Sender: TObject);
begin
  try
  MailClient.Authorize;
  em:=mailclient.Mail;
  dom:='';
  If em <> '' then
    for g:=1 to Length(em) do
      if em[g] = '@' then
      begin
        dom:=copy(em,g+1,Length(em)-g-3);
        break;
       end;
  except on E : Exception do
  end ;
end;

procedure TForm1.MailClientRequestHost(Sender: TObject);
begin
  MailClient.Connect;
end;

procedure TForm1.RichEdit2KeyPress(Sender: TObject; var Key: Char);
begin
  case key of
  #13:
  begin
    if not ChckEnter then exit;
    if ChckInternet then
      if not IsConnectedToInternet then
      begin
        Application.MessageBox('Не могу подключится к интернету, иди на koder.kz читай мануалы','Error', MB_ICONERROR);
        exit;
      end;

    if ChckLstKontact.ItemIndex = -1 then
    begin
      Application.MessageBox('Для рассылки не выбран не один контакт.'+ #13#10 +'Выберите хотя-бы один контакт (установите галочку)','Error', MB_ICONERROR);
      Exit;
    end;
    if not GetSend then
      If Mailclient.Connected then
      begin
        if CheckSmile then
          RichEdit2.Text := AddSmile(RichEdit2.Text, SmileCount);
          
        mailclient.SendMessage(ChckLstKontact.Items[ChckLstKontact.ItemIndex],RichEdit2.Text);
        RE_AlignText( RichEdit1, taLeftJustify );
        With richedit1.SelAttributes do
        begin
          Color:=clRed;
          RichEdit1.Lines.Add('Я написал '+TimeToStr(Time));
          Color:=clblack;
          RichEdit1.Lines.Add(RichEdit2.Text);
          RichEdit1.Lines.Add('  ');
          RichEdit2.Clear;
        end;
        Label4.Caption := Format('Посл. получил сооб. %s',[ChckLstKontact.Items[ChckLstKontact.ItemIndex]]);
      end;
  end;
  end;
end;

procedure TForm1.MailClientMessage(Sender: TObject; MsgID: Cardinal; From,
  Text: String; IsRTF: Boolean);
begin
  RE_AlignText( RichEdit1, taRightJustify );
  With richedit1.SelAttributes do
  begin
    Color:=clblue;
    RichEdit1.Lines.Add(TimeToStr(time)+' от '+ From);
    Color:=clblack;
    RichEdit1.Lines.Add('        '+TExt);
  end;
  Inc(TakeMess); 
  label2.Caption := Format('Сообщений %d',[TakeMess]);
end;

procedure TForm1.MailClientMailBoxStatus(Sender: TObject;
  Reason: Cardinal);
begin
  StatusBar1.Panels[1].Text:='Непочитаных писем'+ IntToStr(Reason);
end;

procedure TForm1.MailClientMailBoxStatusNew(Sender: TObject;
  MsgNum: Cardinal; MailSender, Subject: String; TimeStamp: Cardinal);
begin
  RedtLog.Lines.Add('MailClientMailBoxStatusNew '+ MailSender+#13+#10+Subject+#13+#10+IntToStr(MsgNum)+#13+#10+IntToStr(TimeStamp));
  if StopError then
    actPauseExecute(Sender);
end;

procedure TForm1.RichEdit2Change(Sender: TObject);
begin
  if mailclient.Connected then
  begin
    if ChckLstKontact.ItemIndex <> -1 then
    begin
      Statusbar1.Panels[2].Text := format('Печатаю %s',[ChckLstKontact.Items[ChckLstKontact.ItemIndex]]);
      mailclient.SendPrintMe(ChckLstKontact.Items[ChckLstKontact.ItemIndex]);
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
   i: Integer;
     MrimS: string;
begin

  GetPause := false ;
  TakeMess := 0 ;

  if WaitForSingleObject(hMutex, 0)<>0 then
  begin
   Application.MessageBox('Приложение уже запущено, увы :(','Error', MB_ICONERROR);
    Application.Terminate;
  end;

  TCPClient.Active:=True;
  if TCPClient.Connected then
    MrimS:=TCPClient.Receiveln(#$A); //Получили адрес и порт
    TCPClient.Disconnect;
  if MrimS <> '' then
    for i:=1 to Length(MrimS) do
      if MrimS[i] = ':' then
        begin
          MailClient.Host:=Copy(MrimS, 0, i-1);
          MailClient.Port:=StrToInt(Copy(MrimS, i+1, Length(MrimS)-i));
          form1.StatusBar1.Panels[0].Text:='Host '+ Mailclient.host+'  Port '+IntToStr(MailClient.Port);
          break;
        end;

  if MailClient.HostInit then
  begin
    MailClient.Connect;  //Соеденяемься
  end
  else
    MailClient.RequestHost;
    
  mailclient.Status:=OnLine;
  Form1.Caption := Format('%s ver. %s',['M.Agent',GetSelfVersion]);

  WorkTime := TWorkTime.Create(true);
  WorkTime.FreeOnTerminate := true;
  WorkTime.StartTime := Now;
  WorkTime.Resume;
end;

procedure TForm1.MailClientRecievedHost(Sender: TObject);
begin
  mailclient.Connect
end;

procedure TForm1.MailClientOfflineMessage(Sender: TObject; MsgID: Cardinal;
  From, Text: String);
begin
  RedtLog.Lines.Add('MailClientOfflineMessage '+ 'Offline  '+From+'  '+Text);
  if StopError then
    actPauseExecute(Sender);
end;

procedure TForm1.MailClientSystemMessage(Sender: TObject; MsgID: Cardinal;
  From, Text: String);
begin
  RedtLog.Lines.Add('SystemMessage '+From+'  '+Text);
  if StopError then
    actPauseExecute(Sender);
end;

procedure TForm1.MailClientContactBeginPrint(Sender: TObject;
  MsgID: Cardinal; From, Text: String);
begin
  StatusBar1.Panels[2].Text:='Собиседник вам пишет .......'
end;

procedure TForm1.MailClientContactEndPrint(Sender: TObject;
  MsgID: Cardinal; From, Text: String);
begin
  StatusBar1.Panels[2].Text:=''
end;


procedure TForm1.MailClientContact(Sender: TObject; GroupID: Cardinal;
  EMail, Nick: String; Status: TStatusClient;
  InvisState: TInvisStateClient);
  var
    I : Integer;
begin
  if not StartWork then
  begin
    if CheckContact then
    begin
      if Status <> offline then
      begin
         ChckLstKontact.Items.Add(Email)
      end;
      StatusBar1.Panels[1].Text := Format('онлайн сейчас %d',[ChckLstKontact.Count]);
    end
    else
    begin
      with ChckLstKontact do
      begin
        ChckLstKontact.Items.Add(Email);
      end;
      StatusBar1.Panels[1].Text := Format('всего сейчас %d',[ChckLstKontact.Count]);
    end;

    for I := 0 to ChckLstKontact.Count -1 do
      ChckLstKontact.Checked[i] := ChckKontackt.Checked;
    
    BtnConnect.Enabled := true;
    BtnOnline.Enabled := true;
    Label1.Caption := Format('Контактов %d ',[ChckLstKontact.Count]);
  end;
end;

function TForm1.READ_status(STATUS: TStatusClient): string;
begin
  if status = OffLine  then
    Read_status:='OffLine';
  if status = OnLine  then
    Read_status:='OnLine';
  if status =  Away  then
    Read_status:=' Away';
  if status = Invisible  then
    Read_status:='Invisible';
end;

procedure TForm1.ListBox11Click(Sender: TObject);
begin
  if ChckLstKontact.Items.Strings[ChckLstKontact.ItemIndex] <>'' then
  begin
    Statusbar1.Panels[2].Text := ChckLstKontact.Items.Strings[ChckLstKontact.ItemIndex];
    Get_Avy
  end ;
end;

procedure TForm1.Get_Avy;
begin
  em:=Statusbar1.Panels[2].Text ;
  If em = 'support@corp.mail.ru' then
    em:='';
  If em <> '' then
  begin
  for g:=1 to Length(em) do
    if em[g] = '@' then
    begin
      dom:=copy(em,g+1,Length(em)-g-3);
      break;
    end;
  end;
end;



procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  mailclient.Disconnect;
end;

procedure TForm1.MailClientContactAuthorize(Sender: TObject;
  MsgID: Cardinal; From, Text: String);
begin
  mailclient.ContactAuthorize(from);
end;

procedure TForm1.MailClientAddContactError(Sender: TObject; GroupID,
  UserID: Cardinal; EMail, Nick: String; Reason: Cardinal);
begin
  RedtLog.Lines.Add('Error');
  if StopError then
    actPauseExecute(Sender);
end;

procedure TForm1.N11Click(Sender: TObject);
begin
  mailclient.AurgorizeMe(ChckLstKontact.Items.Strings[ChckLstKontact.ItemIndex],'777777');
end;

procedure TForm1.N10Click(Sender: TObject);
begin
  mailclient.DeleteContact(ChckLstKontact.Items.Strings[ChckLstKontact.ItemIndex],'',0,0);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  form1.MailClient.Status := OnLine;
end;

procedure TForm1.N6Click(Sender: TObject);
begin
  mailClient.Connect;
  mailclient.status := away;
end;

procedure TForm1.N7Click(Sender: TObject);
begin
  mailclient.status := OffLine
end;

procedure TForm1.actOnlineExecute(Sender: TObject);
begin
  mailclient.Status:=Online;
  mailClient.Connect;
  BtnAll.Enabled := true;
  StartWork := true;
end;

procedure TForm1.actSaveLogFileExecute(Sender: TObject);
begin
  if SaveDialog1.Execute then
    RedtLog.Lines.SaveToFile(SaveDialog1.FileName);
end;

procedure TForm1.actAboutExecute(Sender: TObject);
begin
  FrmAbout := TFrmAbout.Create(Self);
  try
    FrmAbout.ShowModal ;
  finally
    FrmAbout.free;
  end;
end;

procedure TForm1.actCloseExecute(Sender: TObject);
begin
  Close;
end;

procedure TForm1.GetSettings;
begin
  if Assigned (FrmSettings) then
  begin
    FrmSettings.PropStorageEh1.LoadProperties;
    
    Login        := FrmSettings.EdtLogin.Text ;
    Pass         := FrmSettings.EdtPass.Text ;
    SendSleep    := FrmSettings.SedtSleep.Value;
    KeySleep     := FrmSettings.SpinEdit1.Value;
    CheckContact := FrmSettings.CheckContact.Checked;
    CheckSmile   := FrmSettings.CheckSmile.Checked;
    CheckSelect  := FrmSettings.CheckSelect.Checked;
    CheckAutoSav := FrmSettings.ChckAutoSave.Checked;
    ChckInternet := FrmSettings.ChckInternet.Checked;
    ChckEnter    := FrmSettings.ChckEnter.Checked;
    SmileCount   := FrmSettings.EdtCntSmile.Value;
    RestartMess  := FrmSettings.SphEdtRestart.Value;
    StopError    := FrmSettings.ChckStopError.Checked;
  end;
end;

procedure TForm1.actSettingExecute(Sender: TObject);
begin
//показать настройки программы
  if FrmSettings.ShowModal = mrOk then
    GetSettings;
end;

procedure TForm1.ChckKontacktClick(Sender: TObject);
  var
    i : Integer;
begin
  for i := 0 to ChckLstKontact.Count -1 do
    ChckLstKontact.Checked[i] := ChckKontackt.Checked ;
end;

procedure TForm1.actPauseExecute(Sender: TObject);
begin
  GetPause := true ;
end;

procedure TForm1.actConnectExecute(Sender: TObject);
begin
  if not IsConnectedToInternet then
  begin
    Application.MessageBox('Не могу подключится к интернету, иди на koder.kz читай мануалы','Error', MB_ICONERROR);
    exit;
  end;
  StartWork := false;
  BtnConnect.Enabled := false;
  Form1.TcpClient.Active:=false;
  GetSettings;
  If (Login <> '') or (Pass <> '') then
  begin
    ChckLstKontact.Clear;
    form1.MailClient.Disconnect;
    form1.MailClient.Mail := Login;
    form1.MailClient.PassWord := Pass;
    form1.TcpClient.Active := true;
    form1.MailClient.Connect;
  end
  else
    actSettingExecute(Sender);
  BtnConnect.Enabled := true;
  TCPclient.Active:=true;
end;

procedure TForm1.actCancelMessageExecute(Sender: TObject);
begin
  GetSend := false;
end;

procedure TForm1.actSendMessageExecute(Sender: TObject);
var
  i       : integer;
  ItemSlc : integer;
  MySend  : string;
  CountM  : Integer;
  MesNmb  : Integer;
  ChckCnt : Boolean;

procedure GetMessageKontakt(IndexContack : Integer; Messages : string; CountMess, NumberMess : Integer);
  var
    tmpStr  : string;
    iSleep  : integer;
    tmpI    : integer;

begin
  ChckLstKontact.Selected[IndexContack] := true ;
  Label3.Caption := Format('Отправляю %d Осталось %d',[NumberMess, CountMess]);

  if ChckLstKontact.Items[IndexContack] <> 'support@corp.mail.ru' then
  begin
    if CheckSmile then
      tmpStr := AddSmile(Messages, SmileCount);

    RichEdit2.Clear;

    for iSleep := 0 to SendSleep do
    begin // наша задержка в секундах, ПО будет зависать меньше по идее
      Sleep(1000);
      Label6.Caption := Format('До след. отпр. осталось: %d',[SendSleep-iSleep]);
      Application.ProcessMessages;

      if not GetSend then
      begin
        Label3.Caption := 'Отправлено/Осталось';
        Label6.Caption := 'До след. отпр. осталось';
        ProgressBar1.Position := 0;
        Exit;
      end;

      if GetPause then
      begin
        RichEdit2.Clear;
        RichEdit2.Lines.Add(Messages);
        StrtI := I ;
        BtnAll.Caption := 'Продолжить' ;
        Exit ;
      end;
    end;

    for tmpI := 1 to Length(tmpStr) do
    begin
      RichEdit2.Text := RichEdit2.Text + tmpStr[tmpI];
      Sleep(random(KeySleep));
      Application.ProcessMessages;

      if not GetSend then
      begin
        Label3.Caption := 'Отправлено/Осталось';
        Label6.Caption := 'До след. отпр. осталось';
        ProgressBar1.Position := 0;
        Exit;
      end;


      if GetPause then
      begin
        RichEdit2.Clear;
        RichEdit2.Lines.Add(Messages);
        StrtI := I ;
        BtnAll.Caption := 'Продолжить' ;
        Exit ;
      end;
    end ;

    If Mailclient.Connected then
    begin
      mailclient.SendMessage(ChckLstKontact.Items[IndexContack],RichEdit2.Text);
      Application.ProcessMessages;
      RE_AlignText( RichEdit1, taLeftJustify );
      With richedit1.SelAttributes do
      begin
        Color:=clRed;
        RichEdit1.Lines.Add(format('Я написал %s, для %s',[TimeToStr(Time), ChckLstKontact.Items[IndexContack]]));
        Color:=clblack;
        RichEdit1.Lines.Add('  '+RichEdit2.Text);
      end;
    end
    else
    begin
      RedtLog.Lines.Add('no connect, Email '+ ChckLstKontact.Items[IndexContack]);
      mailclient.Disconnect;
      mailclient.Status:=Online;
      mailClient.Connect;
      actPauseExecute(Sender);
    end;
  end;
end;

begin
  CountM  := 0 ;
  MesNmb  := 1 ;
  ChckCnt := false;

  for i := 0 to ChckLstKontact.Count -1 do
    if ChckLstKontact.Checked[i] then
    begin
      Inc(CountM);
      ChckCnt := true ;
    end;

  if not ChckCnt then
  begin
    Application.MessageBox('Для рассылки не выбран не один контакт.'+ #13#10 +'Выберите хотя-бы один контакт (установите галочку)','Error', MB_ICONERROR);
    Exit;
  end;

  if ChckInternet then
    if not IsConnectedToInternet then
    begin
      Application.MessageBox('Не могу подключится к интернету, иди на koder.kz читай мануалы','Error', MB_ICONERROR);
      exit;
    end;

  if (CheckSelect) and (not GetPause) then  // начинать с указанного контакта
  begin
    if ChckLstKontact.ItemIndex = -1 then
      ChckLstKontact.Selected[0] := true ;
    StrtI := ChckLstKontact.ItemIndex;
  end
  else if not GetPause then
    StrtI := 0;

  MySend := RichEdit2.Text;
  TakeMess := 0 ;
  GetSend  := true ;
  GetPause := false ;


  ItemSlc := StrtI;

  BtnAll.Caption := 'Отправить всем';
  ProgressBar1.Max := ChckLstKontact.Count - MesNmb - ItemSlc ;
  ProgressBar1.Position := 1 ;
  randomize;

  Label3.Caption := Format('Отправляю %d Осталось %d',[MesNmb, ChckLstKontact.Count - MesNmb - ItemSlc]);

  for i := StrtI to ChckLstKontact.Count -1 do
  begin
    if ((RestartMess+12 = MesNmb) and (RestartMess <> 0)) then
      Restart;

    if not GetSend then
    begin
      Label3.Caption := 'Отправлено/Осталось';
      Label6.Caption := 'До след. отпр. осталось';
      ProgressBar1.Position := 0;
      Exit;
    end;

    if GetPause then Exit ;

    if (ChckLstKontact.Checked[i])then
    begin
      GetMessageKontakt(i, MySend, ChckLstKontact.Count - MesNmb - ItemSlc, MesNmb);
      Inc(MesNmb);
      if CheckAutoSav then
        PropStorageEh2.SaveProperties;

      Label4.Caption := Format('Посл. получил сооб. %s',[ChckLstKontact.Items[I]]);
    end;

    ProgressBar1.Position := ProgressBar1.Position + 1 ;
    Application.ProcessMessages;
  end;

end;

procedure TForm1.actFindContactExecute(Sender: TObject);
  var
    NmbLst : integer;
begin
// найти последнего получателя
  NmbLst := FinKontackt;
  if NmbLst >= 0 then
    ChckLstKontact.Selected[NmbLst] := true ;
end;

function TForm1.FinKontackt: Integer;
  var
    TmpStr : string;
    NmbLst : integer;
begin
  TmpStr := copy(Label4.Caption,20,Length(Label4.Caption));
  TmpStr := trim(TmpStr);
  NmbLst := ChckLstKontact.Items.IndexOf(TmpStr);
  result := NmbLst;
end;

procedure TForm1.actFindAndSendExecute(Sender: TObject);
  var
    LastCnt : Integer;
begin
//Начать отправку с последнего
  LastCnt := FinKontackt;
  if LastCnt = -1 then
  begin
    Application.MessageBox('Не могу найти последний контакт,'+ #13#10 +'его нет в контакт листе, увы :(','Error', MB_ICONERROR);
    exit ;
  end;
  CheckSelect := true ;
  ChckLstKontact.ItemIndex := LastCnt;
  actSendMessageExecute(Sender);
  GetSettings;// вернем настройки
end;

procedure TForm1.MailClientErrorAuthorize(Sender: TObject; Reason: String);
begin
  Application.MessageBox('Не могу подключится с этой учетной записью, увы :(','Error', MB_ICONERROR);
end;

procedure TForm1.CtnHideLogClick(Sender: TObject);
begin
  if RedtLog.Visible then
  begin
    RedtLog.Visible := false;
    CtnHideLog.Caption := 'Показать';
    Panel12.Height     := 18;
  end
  else
  begin
    RedtLog.Visible    := true;
    CtnHideLog.Caption := 'Скрыть';
    Panel12.Height     := 118;
  end;
end;

procedure TForm1.Restart;
  var
    I : Integer;
begin
// перезапуск в случае ошибки или просто так
  actPauseExecute(Nil);
// отключаемся от сервера
  mailclient.status := OffLine;
  mailclient.Disconnect;
  sleep(5000);
  actConnectExecute(Nil);
  for i := 1 to 100 do
  begin
    application.ProcessMessages;
    sleep(100);
    application.ProcessMessages;
  end;
  actOnlineExecute(Nil);
  sleep(1000);
  GetPause := false;
  actFindAndSendExecute(Nil);
end;

initialization

  hMutex := CreateMutex(nil, True, PChar(ExtractFileName(Application.ExeName)));

finalization

  CloseHandle(hMutex);

end.
