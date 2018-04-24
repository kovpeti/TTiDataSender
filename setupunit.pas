unit setupunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Buttons, ExtCtrls, ComCtrls, Spin;

type

  { TSetupForm }

  TSetupForm = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    SaveSettingsCheckBox: TCheckBox;
    Label4: TLabel;
    RoundGroupBox: TGroupBox;
    Label5: TLabel;
    MoveCursorComboBox: TComboBox;
    CommPortsComboBox: TComboBox;
    MoveCursorGroupBox: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    MsgSystemRadioGroup: TRadioGroup;
    RoundSpinEdit: TSpinEdit;
    SendEnterCheckBox: TCheckBox;
    TabSheet2: TTabSheet;
    TargetSwRadioGroup: TRadioGroup;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TSwComboBox: TComboBox;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MsgSystemRadioGroupClick(Sender: TObject);
    procedure TargetSwRadioGroupClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  SetupForm: TSetupForm;

implementation

uses JwaWindows, jwatlhelp32;
{$R *.lfm}

{ TSetupForm }

procedure TSetupForm.FormCreate(Sender: TObject);
var
  Proc: TPROCESSENTRY32;
  hSnap: HWND;
  Looper: BOOL;
begin
     TSwComboBox.Items.Clear;
     Proc.dwSize := SizeOf(Proc);
     hSnap := CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
     Looper := Process32First(hSnap, Proc);
     while Integer(Looper) <> 0 do
     begin
          TSwComboBox.Items.Add(ExtractFileName(proc.szExeFile));
          Looper := Process32Next(hSnap, Proc);
     end;
     CloseHandle(hSnap);
     //TSwComboBox.ItemIndex:=0;
     //TargetSwRadioGroupClick(self);
end;

procedure TSetupForm.MsgSystemRadioGroupClick(Sender: TObject);
begin
     if (MsgSystemRadioGroup.ItemIndex=2) and (TargetSwRadioGroup.ItemIndex=0) then
        MsgSystemRadioGroup.ItemIndex:=0;
end;

procedure TSetupForm.FormActivate(Sender: TObject);
begin
     TargetSwRadioGroupClick(self);
end;

procedure TSetupForm.TargetSwRadioGroupClick(Sender: TObject);
begin
     case TargetSwRadioGroup.ItemIndex of
          0:begin
                 TSwComboBox.Enabled:=false;
                 //TSwComboBox.Text:='EXCEL.EXE';
                 TSwComboBox.ItemIndex:=TSwComboBox.Items.IndexOf('EXCEL.EXE');
                 if TSwComboBox.ItemIndex=-1 then TargetSwRadioGroup.ItemIndex:=2;
          end;
          1:begin
                 TSwComboBox.Enabled:=false;
                 //TSwComboBox.Text:='soffice.bin';
                 TSwComboBox.ItemIndex:=TSwComboBox.Items.IndexOf('soffice.bin');
                 if TSwComboBox.ItemIndex=-1 then TargetSwRadioGroup.ItemIndex:=2;
          end;
          else begin
                 TSwComboBox.Enabled:=true;
                 //TargetSwRadioGroup.ItemIndex:=2;
          end;
     end;
end;

end.

