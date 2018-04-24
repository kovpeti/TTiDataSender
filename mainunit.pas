unit mainunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, TTi1604DsplPanel, TTi1604comm, synaser, Forms,
  Controls, Graphics, Dialogs, StdCtrls, Buttons, ActnList, ComCtrls, ExtCtrls,
  Variants, ComObj, ClipBrd, XMLPropStorage, MouseAndKeyInput;

type

  { TMainForm }

  TMainForm = class(TForm)
    SendRightAction: TAction;
    SendLeftAction: TAction;
    SendDownAction: TAction;
    SendUpAction: TAction;
    StatusBar1: TStatusBar;
    StayOnTopAction: TAction;
    Label1: TLabel;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    DividerToolButton: TToolButton;
    ToolButton8: TToolButton;
    SOTButton: TToolButton;
    SendTextAction: TAction;
    PrgInfoAction: TAction;
    SendEnterAction: TAction;
    TextSendEdit: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SendButton: TBitBtn;
    SendAction: TAction;
    ExitAction: TAction;
    StopAction: TAction;
    StartAction: TAction;
    SettingsAction: TAction;
    ActionList1: TActionList;
    ImageList1: TImageList;
    TTi1604Comm1: TTTi1604Comm;
    TTi1604DsplPanel1: TTTi1604DsplPanel;
    XMLPropStorage1: TXMLPropStorage;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure SendButtonKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SendDownActionExecute(Sender: TObject);
    procedure SendLeftActionExecute(Sender: TObject);
    procedure SendRightActionExecute(Sender: TObject);
    procedure SendUpActionExecute(Sender: TObject);
    procedure StayOnTopActionExecute(Sender: TObject);
    procedure StopActionExecute(Sender: TObject);
    procedure ExitActionExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PrgInfoActionExecute(Sender: TObject);
    procedure SendActionExecute(Sender: TObject);
    procedure SendButtonKeyPress(Sender: TObject; var Key: char);
    procedure SendEnterActionExecute(Sender: TObject);
    procedure SendTextActionExecute(Sender: TObject);
    procedure SettingsActionExecute(Sender: TObject);
    procedure StartActionExecute(Sender: TObject);
    procedure TextSendEditEnter(Sender: TObject);
    procedure TextSendEditKeyPress(Sender: TObject; var Key: char);
    procedure XMLPropStorage1RestoreProperties(Sender: TObject);
    procedure XMLPropStorage1SavingProperties(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    PressedKey:char;
    procedure UpdateStatusBar;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

{ TMainForm }

uses JwaWindows, jwatlhelp32, lcltype,math,setupunit,aboutunit;//,MouseAndKeyInput;

var TargetSoftware:string;
    Digits:integer;
    SendEnter:boolean;
    MyHandle:HWND;
    MyPID:DWORD;
    TargetHandle:HWND;
    TargetPID:DWORD;
    MsgSystem:integer;  {0:WinMessaging,1:ClipBoard,2:UNO}
    TargetSystem:integer; {0:MS Office,1:Open/Libre Office, 2:Other}
    MoveCursor:integer;
    SaveSettings:boolean;

function GetProcessMainWindow(const PID: DWORD; const wFirstHandle: HWND): HWND;
var
  wHandle: HWND;
  ProcID: DWord;
begin
  Result := 0;
  wHandle := GetWindow(wFirstHandle, GW_HWNDFIRST);
  while wHandle > 0 do
  begin
    GetWindowThreadProcessID(wHandle, @ProcID);
    if (ProcID = PID) and (GetParent(wHandle) = 0) then //and (IsWindowVisible(wHandle)) then
    begin
      Result := wHandle;
      Break;
    end;
    wHandle := GetWindow(wHandle, GW_HWNDNEXT);
  end;
end;

function FindProcessByName(const ProcName:string): DWORD;
var
  Proc: TPROCESSENTRY32;
  hSnap: HWND;
  Looper: BOOL;
begin
  Result := 0;
  Proc.dwSize := SizeOf(Proc);
  hSnap := CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
  Looper := Process32First(hSnap, Proc);
  while Integer(Looper) <> 0 do
  begin
    if UpperCase(ExtractFileName(proc.szExeFile)) = UpperCase(ProcName) then
    begin
      Result:=proc.th32ProcessID;
      Break;
    end;
    Looper := Process32Next(hSnap, Proc);
  end;
  CloseHandle(hSnap);
end;

function round2(const Number: extended; const Places: longint): extended;
var t: extended;
begin
   t := power(10, places);
   round2 := round(Number*t)/t;
end;

procedure GetMyHandle;
begin
  MyPID := FindProcessByName(ApplicationName+'.exe');
  if MyPID <> 0 then
     MyHandle := GetProcessMainWindow(MyPID, MainForm.Handle);
end;

procedure GetTargetHandle;
begin
  TargetPID := FindProcessByName(TargetSoftware);
  if TargetPID <> 0 then
     TargetHandle := GetProcessMainWindow(TargetPID, MainForm.Handle);
end;

procedure SwitchBack;
begin
  if MyHandle=0 then GetMyHandle;
  if MyHandle<>0 then begin
     ShowWindow(MyHandle, SW_RESTORE);
     SetWindowPos(MyHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
  end;
  //MainForm.SendButton.SetFocus; { TODO : Without this app window may not show - Depends on Win version?? }
end;

procedure SendControlKey(SendKey:integer);
begin
     if MainForm.StartAction.Enabled then exit; //If measure stopped then do nothing
     if TargetHandle = 0 then GetTargetHandle;
     if TargetHandle <> 0 then
     begin
       ShowWindow(TargetHandle, SW_RESTORE);
       SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
       if MsgSystem=0 then
            begin //WinMsgSystem
                   PostMessage(TargetHandle,WM_KEYDOWN,SendKey,0);
                   PostMessage(TargetHandle,WM_KEYUP,SendKey,0);
            end else //Every other case
                   KeyInput.Press(SendKey);
       SetWindowPos(TargetHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
     end;
   //Switch back
   SwitchBack;
end;

procedure SendStringToExcel(SendStr:string);
var
  i:integer;
//  c:integer;
var
  Server, desktop, dispatcher: variant;

  FUNCTION variantArray(): Variant;
  BEGIN
    variantArray:= VarArrayCreate([0, -1], varVariant);
  END;
begin
     if MainForm.StartAction.Enabled then exit; //If measure stopped then do nothing
     if TargetHandle = 0 then GetTargetHandle;
     if TargetHandle <> 0 then
        ShowWindow(TargetHandle, SW_RESTORE);
        SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
       case MsgSystem of
            0:begin {The Windows Message System is available from Win7 only}
              for i := 1 to Length(SendStr) do begin
                  PostMessage(TargetHandle, WM_CHAR, Word(SendStr[i]), 0);
              end;
              if SendEnter then begin
                PostMessage(TargetHandle,WM_KEYDOWN,VK_RETURN,0);
                PostMessage(TargetHandle,WM_KEYUP,VK_RETURN,0);
              end;
              case MoveCursor of
                   0:begin PostMessage(TargetHandle,WM_KEYDOWN,VK_UP,0);PostMessage(TargetHandle,WM_KEYUP,VK_UP,0);end;
                   1:begin PostMessage(TargetHandle,WM_KEYDOWN,VK_DOWN,0);PostMessage(TargetHandle,WM_KEYUP,VK_DOWN,0);end;
                   2:begin PostMessage(TargetHandle,WM_KEYDOWN,VK_LEFT,0);PostMessage(TargetHandle,WM_KEYUP,VK_LEFT,0);end;
                   3:begin PostMessage(TargetHandle,WM_KEYDOWN,VK_RIGHT,0);PostMessage(TargetHandle,WM_KEYUP,VK_RIGHT,0);end;
              end;
            end;
            1:begin {Use ClipBoard}
              Clipboard.Clear;
              Clipboard.AsText:=SendStr;
              //PasteClipBoard;
              KeyInput.Apply([ssCtrl]);
              sleep(10);
              KeyInput.Press(VK_V);
              sleep(10);
              KeyInput.Unapply([ssCtrl]);
              sleep(10);
              if SendEnter then
                 KeyInput.Press(VK_RETURN);
              case MoveCursor of
                   0:KeyInput.Press(VK_UP);
                   1:KeyInput.Press(VK_DOWN);
                   2:KeyInput.Press(VK_LEFT);
                   3:KeyInput.Press(VK_RIGHT);
              end;
            end;
            2:begin {Use Open/Libre Office UNO}
              Server := CreateOleObject('com.sun.star.ServiceManager');
              desktop := Server.createInstance('com.sun.star.frame.Desktop');
              dispatcher := Server.createInstance('com.sun.star.frame.DispatchHelper');
              dispatcher.executeDispatch(desktop.CurrentFrame, '.uno:Paste', '', 0, variantArray());
              dispatcher:= unassigned;
              desktop:= unassigned;
              Server:= unassigned;
              if SendEnter then
                 KeyInput.Press(VK_RETURN);
              case MoveCursor of
                   0:KeyInput.Press(VK_UP);
                   1:KeyInput.Press(VK_DOWN);
                   2:KeyInput.Press(VK_LEFT);
                   3:KeyInput.Press(VK_RIGHT);
              end;
            end;
       end;
     //Switch back
     SwitchBack;
end;

procedure TMainForm.UpdateStatusBar;
begin
     StatusBar1.Panels[0].Text:=TTi1604Comm1.Port;
     if TTi1604Comm1.Active then StatusBar1.Panels[1].Text:='Connected'
        else StatusBar1.Panels[1].Text:='Idle';
     StatusBar1.Panels[2].Text:=TargetSoftware;
end;

procedure TMainForm.ExitActionExecute(Sender: TObject);
begin
     StopActionExecute(self);
     Application.Terminate;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
     DividerToolButton.Width:=212;  //Depends on Windows version, this control loose width property value
     {Check saved settings}
     if FileExists(ApplicationName+'.xml') then
        XMLPropStorage1.Active:=true
        else begin
          {Basic settings}
          TargetSoftware:='EXCEL.EXE';
          SendEnter:=true;
          Digits:=2;
          MsgSystem:=0;
          TTi1604Comm1.Port:='COM1';
          TTi1604Comm1.Active:=false;
          MoveCursor:=3;
          SettingsActionExecute(nil);
          if SaveSettings then XMLPropStorage1.Active:=true
             else XMLPropStorage1.Active:=false;
        end;
     UpdateStatusBar;
end;

procedure TMainForm.PrgInfoActionExecute(Sender: TObject);
var AForm:TAboutForm;
begin
     {Open About dialog}
     AForm:=TAboutForm.Create(self);
     try
        AForm.ShowModal;
     finally
       AForm.Free;
     end;
end;

procedure TMainForm.FormActivate(Sender: TObject);
begin
     SendButton.SetFocus;
end;

procedure TMainForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
     ExitActionExecute(self);
end;

procedure TMainForm.SendButtonKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if Key=vk_UP then SendUpActionExecute(self) else
     if Key=vk_DOWN then SendDownActionExecute(self) else
     if Key=vk_LEFT then SendLeftActionExecute(self) else
     if Key=vk_RIGHT then SendRightActionExecute(self);
end;

procedure TMainForm.SendDownActionExecute(Sender: TObject);
begin
  SendControlKey(vk_DOWN);
//  SendButton.SetFocus;
end;

procedure TMainForm.SendLeftActionExecute(Sender: TObject);
begin
  SendControlKey(vk_LEFT);
//  SendButton.SetFocus;
end;

procedure TMainForm.SendRightActionExecute(Sender: TObject);
begin
  SendControlKey(vk_RIGHT);
//  SendButton.SetFocus;
end;

procedure TMainForm.SendUpActionExecute(Sender: TObject);
begin
     SendControlKey(vk_UP);
//     SendButton.SetFocus;
end;

procedure TMainForm.StayOnTopActionExecute(Sender: TObject);
begin
     if SOTButton.Down then MainForm.FormStyle:=fsSystemStayOnTop
        else MainForm.FormStyle:=fsNormal;
end;

procedure TMainForm.StopActionExecute(Sender: TObject);
begin
  TTi1604Comm1.Active:=false;
  StartAction.Enabled:=true;
  StopAction.Enabled:=False;
  //SendAction.Enabled:=false;
  //SendTextAction.Enabled:=false;
  //SendEnterAction.Enabled:=false;
  //TextSendEdit.Enabled:=false;
  UpdateStatusBar;
end;

procedure TMainForm.SendActionExecute(Sender: TObject);
//var Val:double;
//    s:string;
begin
//     Val:=TTi1604DsplPanel1.Value;
//     s:=floattostrf(Val,ffFixed,8,Digits);
     SendStringToExcel(floattostrf(TTi1604DsplPanel1.Value,ffFixed,8,Digits));
end;

procedure TMainForm.SendButtonKeyPress(Sender: TObject; var Key: char);
begin
     {If any button pressed apart from Enter, send focus to TextSendEdit.}
     if (Key<>#13) then begin
        TextSendEdit.Text:=Key;
        TextSendEdit.SetFocus;
     end;
end;

procedure TMainForm.SendEnterActionExecute(Sender: TObject);
begin
     SendControlKey(vk_RETURN);
//     SendButton.SetFocus;
end;

procedure TMainForm.SendTextActionExecute(Sender: TObject);
begin
  SendStringToExcel(TextSendEdit.Text);
  TextSendEdit.Clear;
  //SendButton.SetFocus;
end;

procedure TMainForm.SettingsActionExecute(Sender: TObject);
var SForm:TSetupForm;
    SList:TStringList;
begin
     SForm:=TSetupForm.Create(self);
     SList := TStringList.Create;
     try
        SForm.CommPortsComboBox.Items.Clear;
        SList.CommaText := GetSerialPortNames();
        SForm.CommPortsComboBox.Items.AddStrings(SList);
        SForm.SendEnterCheckBox.Checked:=SendEnter;
        SForm.RoundSpinEdit.Value:=Digits;
        SForm.MsgSystemRadioGroup.ItemIndex:=MsgSystem;
        SForm.TargetSwRadioGroup.ItemIndex:=TargetSystem;
        SForm.MoveCursorComboBox.ItemIndex:=MoveCursor;
        SForm.SaveSettingsCheckBox.Checked:=SaveSettings;
        SForm.TSwComboBox.Text:=TargetSoftware; { TODO : Text property will be changed even if correct here }
        SForm.CommPortsComboBox.Text:=TTi1604Comm1.Port;

        if SForm.ShowModal=mrOK then begin
           TTi1604Comm1.Port:=SForm.CommPortsComboBox.Text;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
           TargetSoftware:=SForm.TSwComboBox.Text;
           TargetPID := FindProcessByName(TargetSoftware);
           if TargetPID <> 0 then
             TargetHandle := GetProcessMainWindow(TargetPID, MainForm.Handle);
           Digits:=SForm.RoundSpinEdit.Value;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
           MsgSystem:=SForm.MsgSystemRadioGroup.ItemIndex;
           TargetSystem:=SForm.TargetSwRadioGroup.ItemIndex;
           MoveCursor:=SForm.MoveCursorComboBox.ItemIndex;
           SaveSettings:=SForm.SaveSettingsCheckBox.Checked;
        end;
     finally
       SForm.Free;
       SList.Free;
     end;
     UpdateStatusBar;
end;

procedure TMainForm.StartActionExecute(Sender: TObject);
begin
     try
        TTi1604Comm1.Active:=true;
        StartAction.Enabled:=false;
        StopAction.Enabled:=true;
     except
       on E: Exception do begin
          ShowMessage(E.Message);
          StartAction.Enabled:=true;
          StopAction.Enabled:=false;
       end;
     end;
     UpdateStatusBar;
end;

procedure TMainForm.TextSendEditEnter(Sender: TObject);
begin
     TextSendEdit.SelStart:=1;
end;

procedure TMainForm.TextSendEditKeyPress(Sender: TObject; var Key: char);
begin
     if Key=#13 then begin
        SendTextActionExecute(self);
        SendButton.SetFocus;
     end;
end;

procedure TMainForm.XMLPropStorage1RestoreProperties(Sender: TObject);
begin
  if FileExists(ApplicationName+'.xml') then begin
     TTi1604Comm1.Port:=XMLPropStorage1.StoredValue['ComPort'];
     Digits:=strtoint(XMLPropStorage1.StoredValue['Round']);
     MsgSystem:=strtoint(XMLPropStorage1.StoredValue['MsgSystem']);
     TargetSystem:=strtoint(XMLPropStorage1.StoredValue['TargetSystem']);
     TargetSoftware:=XMLPropStorage1.StoredValue['TargetSoftware'];
     SendEnter:=strtobool(XMLPropStorage1.StoredValue['SendEnter']);
     MoveCursor:=strtoint(XMLPropStorage1.StoredValue['MoveCursor']);
     SaveSettings:=strtobool(XMLPropStorage1.StoredValue['SaveSettings']);
     UpdateStatusBar;
  end;
end;

procedure TMainForm.XMLPropStorage1SavingProperties(Sender: TObject);
begin
     XMLPropStorage1.StoredValue['ComPort'] := TTi1604Comm1.Port;
     XMLPropStorage1.StoredValue['Round'] := IntToStr(Digits);
     XMLPropStorage1.StoredValue['MsgSystem'] := inttostr(MsgSystem);
     XMLPropStorage1.StoredValue['TargetSystem'] := inttostr(TargetSystem);
     XMLPropStorage1.StoredValue['TargetSoftware'] := TargetSoftware;
     XMLPropStorage1.StoredValue['SendEnter'] := booltostr(SendEnter);
     XMLPropStorage1.StoredValue['MoveCursor'] := inttostr(MoveCursor);
     XMLPropStorage1.StoredValue['SaveSettings'] := BoolToStr(SaveSettings);
end;

end.

