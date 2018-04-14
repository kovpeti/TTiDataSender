unit mainunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, TTi1604DsplPanel, TTi1604comm, synaser,
  Forms, Controls, Graphics, Dialogs, StdCtrls, Buttons, ActnList,
  ComCtrls, ExtCtrls;

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
    //PortName:string;
    //PortSpeed:integer;
    SendEnter:boolean;
    MyHandle:HWND;
    MyPID:DWORD;
    TargetHandle:HWND;
    TargetPID:DWORD;

const
     MyProgramName='TTi1604DataSender.exe';

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
  MyPID := FindProcessByName(MyProgramName);
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
  //MainForm.SendButton.SetFocus; //Without this app window may not show - Depends on wWin version??
end;

procedure SendControlKey(SendKey:integer);
begin
     if MainForm.StartAction.Enabled then exit; //If measure stopped then do nothing
     if TargetHandle = 0 then GetTargetHandle;
     if TargetHandle <> 0 then
     begin
       ShowWindow(TargetHandle, SW_RESTORE);
       SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
       //For Vista and up
       PostMessage(TargetHandle,WM_KEYDOWN,SendKey,0);
       PostMessage(TargetHandle,WM_KEYUP,SendKey,0);
       //For XP
       //KeyInput.Press(SendKey);
       SetWindowPos(TargetHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
     end;
   //Switch back
   SwitchBack;
end;

procedure SendStringToExcel(SendStr:string);
var
  i:integer;
//  c:integer;
begin
     if MainForm.StartAction.Enabled then exit; //If measure stopped then do nothing
     if TargetHandle = 0 then GetTargetHandle;
     if TargetHandle <> 0 then
     begin
       ShowWindow(TargetHandle, SW_RESTORE);
       SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
         for i := 1 to Length(SendStr) do begin
           //For Wista and up
           PostMessage(TargetHandle, WM_CHAR, Word(SendStr[i]), 0);
           {
           -- This solution causes repeated key press --
           c:=Word(SendStr[i]);
           PostMessage(TargetHandle,WM_KEYDOWN,c,0);
           PostMessage(TargetHandle,WM_KEYUP,c,0);}
           //For XP
      { TODO : Only standard keys are supported.
Shifted characters eg !,",% are not.
The problem is the different keyboard layouts.
Need to test and clarify!
7/3/18 - still not working on XP, repeat keys on Win7 }
{                c:=Word(SendStr[i]);
           if c=$2E then c:=VK_DECIMAL else           //$2E is the decimal point ASCII representative
           if (c>=97) and (c<=122) then c:=c-32 else  //a..z
           if (c>=65) and (c<=90) then begin         //A..Z
              c:=c;
               KeyInput.Apply([ssShift]);
             end else
           if (c>=48) and (c<=57) then c:=c else      //0..9
           c:=0;                                       //Unsupported button
           if c>=0 then KeyInput.Press(c);                // This will simulate press of F1 function key.
           KeyInput.UnApply([ssShift]);                   //for capital letters
}
         end;
       if SendEnter then begin
         //Wista and up
         PostMessage(TargetHandle,WM_KEYDOWN,VK_RETURN,0);
         PostMessage(TargetHandle,WM_KEYUP,VK_RETURN,0);
         //XP
         //KeyInput.Press(VK_RETURN);
       end;
       SetWindowPos(TargetHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
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
     //Basic settings if Cancel pressed on Settings dialog
     TargetSoftware:='EXCEL.EXE';
     SendEnter:=true;
     Digits:=2;
     //RxString:='';
     TTi1604Comm1.Port:='COM1';
     TTi1604Comm1.Active:=false;
     //TTi1604msgOFF
     SettingsActionExecute(nil);
     UpdateStatusBar;
end;

procedure TMainForm.PrgInfoActionExecute(Sender: TObject);
var AForm:TAboutForm;
begin
     //Open About dialog
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
begin
     SendStringToExcel(floattostrf(TTi1604DsplPanel1.Value,ffGeneral,8,Digits));
end;

procedure TMainForm.SendButtonKeyPress(Sender: TObject; var Key: char);
begin
     //If any button pressed apart from Enter, send focus to TextSendEdit.
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
        if SForm.CommPortsComboBox.Items.Count > 0 then
           SForm.CommPortsComboBox.ItemIndex := 0
           else
           SForm.CommPortsComboBox.Text := '';
        if SForm.ShowModal=mrOK then begin
           TTi1604Comm1.Port:=SForm.CommPortsComboBox.Text;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
           TargetSoftware:=SForm.TSwComboBox.Text;
           TargetPID := FindProcessByName(TargetSoftware);
           if TargetPID <> 0 then
             TargetHandle := GetProcessMainWindow(TargetPID, MainForm.Handle);
           Digits:=SForm.RoundSpinEdit.Value;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
        end;
     finally
       SForm.Free;
       SList.Free;
     end;
     UpdateStatusBar;
end;

procedure TMainForm.StartActionExecute(Sender: TObject);
var Ok:boolean;
    b:byte;
    Counter:integer;
begin
     try
        TTi1604Comm1.Active:=true;
     except
       on E: Exception do
          ShowMessage(E.Message);
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

end.

