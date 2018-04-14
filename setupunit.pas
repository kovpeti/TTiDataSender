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
    CommPortsComboBox: TComboBox;
    Label2: TLabel;
    Label5: TLabel;
    TargetSwRadioGroup: TRadioGroup;
    SendEnterCheckBox: TCheckBox;
    Label3: TLabel;
    Label4: TLabel;
    PageControl1: TPageControl;
    RoundSpinEdit: TSpinEdit;
    TabSheet1: TTabSheet;
    TSwComboBox: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure Label2Click(Sender: TObject);
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
     TSwComboBox.ItemIndex:=0;
     TargetSwRadioGroupClick(self);
end;

procedure TSetupForm.Label2Click(Sender: TObject);
begin

end;

procedure TSetupForm.TargetSwRadioGroupClick(Sender: TObject);
begin
     case TargetSwRadioGroup.ItemIndex of
          0:begin
                 TSwComboBox.Enabled:=false;
                 TSwComboBox.Text:='EXCEL.EXE';
          end;
          1:begin
                 TSwComboBox.Enabled:=false;
                 TSwComboBox.Text:='soffice.bin';
          end;
          2:begin
                 TSwComboBox.Enabled:=true;
          end;
     end;
end;

end.

