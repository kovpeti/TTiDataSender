unit aboutunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Buttons,
  StdCtrls;

type

  { TAboutForm }

  TAboutForm = class(TForm)
    BitBtn1: TBitBtn;
    DocumentationURLLabel: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    VersionLabel: TLabel;
    Label5: TLabel;
    procedure BitBtn1Click(Sender: TObject);
    procedure DocumentationURLLabelMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure DocumentationURLLabelMouseEnter(Sender: TObject);
    procedure DocumentationURLLabelMouseLeave(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  AboutForm: TAboutForm;

const
  Version='0.1.6';

implementation

uses LCLIntf;
{$R *.lfm}

{ TAboutForm }

procedure TAboutForm.BitBtn1Click(Sender: TObject);
begin
     close;
end;

procedure TAboutForm.DocumentationURLLabelMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
     OpenURL(TLabel(Sender).Caption);
end;

procedure TAboutForm.DocumentationURLLabelMouseEnter(Sender: TObject);
begin
     TLabel(Sender).Font.Style := [fsUnderLine];
     TLabel(Sender).Font.Color := clRed;
     TLabel(Sender).Cursor := crHandPoint;
end;

procedure TAboutForm.DocumentationURLLabelMouseLeave(Sender: TObject);
begin
     TLabel(Sender).Font.Style := [];
     TLabel(Sender).Font.Color := clBlue;
     TLabel(Sender).Cursor := crDefault;
end;

procedure TAboutForm.FormCreate(Sender: TObject);
begin
     VersionLabel.Caption:=Version;
end;

end.

