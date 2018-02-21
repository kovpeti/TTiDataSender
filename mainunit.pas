unit mainunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, synaser,LedNumber, Forms, Controls,
  Graphics, Dialogs, StdCtrls, Buttons, ActnList, ComCtrls, ExtCtrls;

type

  { TMainForm }

  TMainForm = class(TForm)
    SendRightAction: TAction;
    SendLeftAction: TAction;
    SendDownAction: TAction;
    SendUpAction: TAction;
    StatusBar1: TStatusBar;
    StayOnTopAction: TAction;
    NegLabel: TLabel;
    ContLabel: TLabel;
    Label1: TLabel;
    mVLabel: TLabel;
    mALabel: TLabel;
    AVLabel: TLabel;
    kRLabel: TLabel;
    HzLabel: TLabel;
    autoLabel: TLabel;
    DiodeLabel: TLabel;
    RLabel: TLabel;
    NullLabel: TLabel;
    HoldLabel: TLabel;
    MinLabel: TLabel;
    MaxLabel: TLabel;
    DCLabel: TLabel;
    ACLabel: TLabel;
    DataTimer: TTimer;
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
    ToolButton6: TToolButton;
    DividerToolButton: TToolButton;
    ToolButton8: TToolButton;
    SOTButton: TToolButton;
    VLabel: TLabel;
    NumberLEDNumber: TLEDNumber;
    DisplayPanel: TPanel;
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
    THldLabel: TLabel;
    procedure AVLabelClick(Sender: TObject);
    procedure CoolBar1Change(Sender: TObject);
    procedure DataTimerTimer(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormPaint(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SendButtonKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SendDownActionExecute(Sender: TObject);
    procedure SendLeftActionExecute(Sender: TObject);
    procedure SendRightActionExecute(Sender: TObject);
    procedure SendUpActionExecute(Sender: TObject);
    procedure StayOnTopActionExecute(Sender: TObject);
    procedure StopActionExecute(Sender: TObject);
    procedure TTiDataPortSerialDataAppear(Sender: TObject);
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
    procedure DisplayData;
    procedure UpdateStatusBar;
    procedure HideAllOnDisplay;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

{ TMainForm }

uses JwaWindows, jwatlhelp32, lcltype,math,setupunit,aboutunit;//,MouseAndKeyInput;
type TSerPort= record
       Name:string;
       Speed:integer;
       DataBit:integer;
       Parity:char;
       StopBit:integer;
       end;


var RxData:array[0..9] of byte;
    TargetSoftware:string;
    Digits:integer;
    MeasuredValue:double;
    //PortName:string;
    //PortSpeed:integer;
    SendEnter:boolean;
    MyHandle:HWND;
    MyPID:DWORD;
    TargetHandle:HWND;
    TargetPID:DWORD;
    SerPort:TBlockSerial;
    SerPortConf:TSerPort;

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
     MainForm.SendButton.SetFocus; //Without this app window may not show
  end;
end;

procedure SendControlKey(SendKey:integer);
begin
     if TargetHandle = 0 then GetTargetHandle;
     if TargetHandle <> 0 then
     begin
       ShowWindow(TargetHandle, SW_RESTORE);
       SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
       PostMessage(TargetHandle,WM_KEYDOWN,SendKey,0);
       PostMessage(TargetHandle,WM_KEYUP,SendKey,0);
       SetWindowPos(TargetHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
     end;
   //Switch back
   SwitchBack;
end;

procedure SendStringToExcel(SendStr:string);
var
  i:integer;
  begin
       if TargetHandle = 0 then GetTargetHandle;
       if TargetHandle <> 0 then
       begin
         ShowWindow(TargetHandle, SW_RESTORE);
         SetWindowPos(TargetHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
           for i := 1 to Length(SendStr) do begin
                  PostMessage(TargetHandle, WM_CHAR, Word(SendStr[i]), 0);
           end;
         if SendEnter then begin
           PostMessage(TargetHandle,WM_KEYDOWN,VK_RETURN,0);
           PostMessage(TargetHandle,WM_KEYUP,VK_RETURN,0);
         end;
         SetWindowPos(TargetHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
       end;
     //Switch back
     SwitchBack;
end;

function SetSerPort:boolean;
begin
     SerPort.CloseSocket;
     SerPort.Connect(SerPortConf.Name);
     if SerPort.LastError<>sOK then begin
        Application.MessageBox(pChar(SerPort.GetErrorDesc(SerPort.LastError)),'Error');
        SetSerPort:=false;
        exit;
     end;
     Sleep(200);
     SerPort.Config(SerPortConf.Speed,SerPortConf.DataBit,SerPortConf.Parity,SerPortConf.StopBit,false,false);
     //SerPort.Config(9600,8,'N',SB1,false,false);
     if SerPort.LastError<>sOK then begin
        Application.MessageBox(pChar(SerPort.GetErrorDesc(SerPort.LastError)),'Error');
        SetSerPort:=false;
        exit;
     end;
     SerPort.Purge;
     SetSerPort:=true;
end;

procedure TMainForm.UpdateStatusBar;
begin
     StatusBar1.Panels[0].Text:=SerPortConf.Name;// PortName;
     if SerPort.InstanceActive then StatusBar1.Panels[1].Text:='Connected'
        else StatusBar1.Panels[1].Text:='Idle';
     StatusBar1.Panels[2].Text:=TargetSoftware;
end;

procedure TMainForm.HideAllOnDisplay;
var i:integer;
begin
  for i:=0 to DisplayPanel.ControlCount-1 do
      DisplayPanel.Controls[i].Visible:=false;
  NumberLEDNumber.Visible:=true;
end;
procedure TMainForm.DisplayData;
var Display:string;
    i,j:integer;
    Units:byte;
    Value:double;
begin
     if RxData[0] <>13 then exit;
     Display:='';
     Units:=(RxData[1] and %00000111);
     //Make invisible all controls with tag<500 on DisplayPanel
     for i:=0 to DisplayPanel.ControlCount-1 do
         if DisplayPanel.Controls[i].Tag<500 then DisplayPanel.Controls[i].Visible:=false;
     //Make visible current unit
     for i:=0 to DisplayPanel.ControlCount-1 do
         if DisplayPanel.Controls[i].Tag=Units then DisplayPanel.Controls[i].Visible:=true;
     if (Units<5) then
        if ((RxData[1] and %00001000)>0) then begin
           AcLabel.Visible:=true;
           DCLabel.Visible:=false;
           end else begin
               AcLabel.Visible:=false;
               DCLabel.Visible:=true;
               end
        else begin
             AcLabel.Visible:=false;
             DCLabel.Visible:=false;
             end;
     case (RxData[1] and %01110000) of
          0:begin
                 ;
            end;
          16:begin
                 if Units=5 then begin
                   kRLabel.Visible:=true;
                   RLabel.Visible:=false;
                 end;
            end;
          32:begin
                 if Units=5 then begin
                   kRLabel.Visible:=true;
                   RLabel.Visible:=false;
                 end;
            end;
          48:begin
                 if Units=5 then begin
                   kRLabel.Visible:=true;
                   RLabel.Visible:=false;
                 end;
            end;
          64:begin
                 if Units=5 then begin
                   kRLabel.Visible:=true;
                   RLabel.Visible:=false;
                 end;
            end;
          80:begin
                 if Units=5 then begin
                   kRLabel.Visible:=true;
                   RLabel.Visible:=false;
                 end;
            end;
          end;
     //Function information
     if (RxData[2] and %00000001)>0 then THldLabel.Visible:=true else THldLabel.Visible:=false;
     if (RxData[2] and %00000010)>0 then begin
        MinLabel.Visible:=true;
        MaxLabel.Visible:=true;
        end else begin
               MinLabel.Visible:=false;
               MaxLabel.Visible:=false;
               end;
     if (RxData[2] and %00010000)>0 then HzLabel.Visible:=true else HzLabel.Visible:=false;
     if (RxData[2] and %00100000)>0 then NullLabel.Visible:=true else NullLabel.Visible:=false;
     if (RxData[2] and %01000000)>0 then AutoLabel.Visible:=true else AutoLabel.Visible:=false;
     //Sign
     if (RxData[3] and %00000010)>0 then NegLabel.Visible:=true else NegLabel.Visible:=false;// Display:='-' else Display:=' ';
     //Numbers and decimal dot
     j:=2;
     for i:=4 to 8 do begin
       case RxData[i] of
            252:Display:=Display+'0';
            253:begin Display:=Display+'0'; inc(j); Display:=Display+'.'; end;
            96:Display:=Display+'1';
            97:begin Display:=Display+'1'; inc(j); Display:=Display+'.'; end;
            218:Display:=Display+'2';
            219:begin Display:=Display+'2'; inc(j); Display:=Display+'.'; end;
            242:Display:=Display+'3';
            243:begin Display:=Display+'3'; inc(j); Display:=Display+'.'; end;
            102:Display:=Display+'4';
            103:begin Display:=Display+'4'; inc(j); Display:=Display+'.'; end;
            182:Display:=Display+'5';
            183:begin Display:=Display+'5'; inc(j); Display:=Display+'.'; end;
            190:Display:=Display+'6';
            191:begin Display:=Display+'6'; inc(j); Display:=Display+'.'; end;
            224:Display:=Display+'7';
            225:begin Display:=Display+'7'; inc(j); Display:=Display+'.'; end;
            254:Display:=Display+'8';
            255:begin Display:=Display+'8'; inc(j); Display:=Display+'.'; end;
            230:Display:=Display+'9';
            231:begin Display:=Display+'9'; inc(j); Display:=Display+'.'; end;
       end;
       inc(j);
     end;
     NumberLEDNumber.Caption:=Display;
     if trystrtofloat(Display,Value) then

     MeasuredValue:=round2(Value,Digits) else MeasuredValue:=0;
     if NegLabel.Visible then MeasuredValue:=MeasuredValue*-1;
end;

procedure TMainForm.ExitActionExecute(Sender: TObject);
begin
     StopActionExecute(self);
     if SerPort<>nil then SerPort.Free;
     Application.Terminate;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
     DividerToolButton.Width:=202;
     //Basic settings
     TargetSoftware:='excel.exe';
     SendEnter:=true;
     Digits:=2;
     //RxString:='';
     SerPortConf.Name:='COM1';
     SerPortConf.Speed:=9600;
     SerportConf.DataBit:=8;
     SerPortConf.Parity:='N';
     SerPortConf.StopBit:=0;
     //PortName:='COM1';
     SerPort:=TBlockSerial.Create;
     //SerPort.Connect(PortName);
     //SerPort.Config(9600,8,'n',1,false,false);
     //SetSerPort;
     NumberLEDNumber.Caption:=' Off ';
     MeasuredValue:=123.456;
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

procedure TMainForm.TTiDataPortSerialDataAppear(Sender: TObject);
begin
end;

procedure TMainForm.AVLabelClick(Sender: TObject);
begin

end;

procedure TMainForm.CoolBar1Change(Sender: TObject);
begin

end;

procedure TMainForm.DataTimerTimer(Sender: TObject);
var i:integer;
    c:byte;
begin
     if SerPort.CanRead(0) then begin
       repeat
         c:=SerPort.RecvByte(0);
       until c=13;
       RxData[0]:=13;
       for i:=1 to 9 do begin
           RxData[i]:=SerPort.RecvByte(10);
       end;
     DisplayData;
     end;
end;

procedure TMainForm.FormActivate(Sender: TObject);
begin
     if SendAction.Enabled then SendButton.SetFocus;
end;

procedure TMainForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
     ExitActionExecute(self);
end;

procedure TMainForm.FormPaint(Sender: TObject);
begin
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
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
  HideAllOnDisplay;
  NumberLEDNumber.Caption:=' Off ';
  SerPort.SendString('v');  //Local mode constant
  //Send more just in case
  SerPort.SendString('v');  //Local mode constant
  SerPort.SendString('v');  //Local mode constant
  Sleep(1000);
  Serport.CloseSocket;
  StartAction.Enabled:=true;
  StopAction.Enabled:=False;
  DataTimer.Enabled:=false;
  SendAction.Enabled:=false;
  SendTextAction.Enabled:=false;
  SendEnterAction.Enabled:=false;
  TextSendEdit.Enabled:=false;
  UpdateStatusBar;
end;

procedure TMainForm.SendActionExecute(Sender: TObject);
begin
     SendStringToExcel(floattostrf(MeasuredValue,ffGeneral,8,Digits));
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
           SerPortConf.Name:=SForm.CommPortsComboBox.Text;;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
           TargetSoftware:=SForm.TSwComboBox.Text;
           TargetPID := FindProcessByName(TargetSoftware);
           if TargetPID <> 0 then
             TargetHandle := GetProcessMainWindow(TargetPID, MainForm.Handle);
           Digits:=SForm.RoundSpinEdit.Value;
           SendEnter:=SForm.SendEnterCheckBox.Checked;
           SerPortConf.Speed:=strtoint(SForm.CommSpeedComboBox.Text);
           SetSerPort;
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
     NumberLEDNumber.Caption:='Wait ';
     SetSerPort;
     //SerPort.Connect(SerPortConf.Name);
     Sleep(200);
     SerPort.RTS:=false;
     SerPort.DTR:=true;
     if SerPort.LastError<>sOK then begin
        Application.MessageBox(pChar(SerPort.GetErrorDesc(SerPort.LastError)),'Error');
        NumberLEDNumber.Caption:=' Off ';
        exit;
     end;
     Sleep(1000);
     Ok:=false;
     Counter:=0;
     repeat
           //SerPort.SendString('u');  //Remote mode constant
           SerPort.SendByte($75);      //'u' in ASCII
           b:=SerPort.RecvByte(100);
           //Check answer
           if (SerPort.LastError=0) and (b=$75) then Ok:=true
              else begin
                SerPort.Purge;
                inc(Counter);
                SerPort.SendString('v');  //Local mode constant
              end;
     until Ok or (Counter>15);
     if not Ok then begin
        NumberLEDNumber.Caption:=' Off ';
        Application.MessageBox('No answer from the instrument','Communication error',0);
        Serport.CloseSocket;
        end else begin
             Sleep(500);
             SerPort.Purge;
             NumberLEDNumber.Caption:='Ready';
             StartAction.Enabled:=false;
             StopAction.Enabled:=true;
             DataTimer.Enabled:=true;
             SendAction.Enabled:=true;
             SendTextAction.Enabled:=true;
             SendEnterAction.Enabled:=true;
             TextSendEdit.Enabled:=true;
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

