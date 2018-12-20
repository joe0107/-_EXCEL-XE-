unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, JcFileUtils, cxClasses,
  cxShellBrowserDialog, ComCtrls, XLSSheetData5, XLSReadWriteII5, cxPropertiesStore, OleAuto, ExcelXP, JclStrings;

type
  TfmMain = class(TForm)
    EditSrc: TEdit;
    btnSrc: TButton;
    btnDoc: TButton;
    EditDoc: TEdit;
    btnOuput: TButton;
    EditOutputFolder: TEdit;
    cxShellBrowserDialog1: TcxShellBrowserDialog;
    btnExec: TButton;
    ProgressBar1: TProgressBar;
    XLSReadWriteII5: TXLSReadWriteII5;
    cxPropertiesStore1: TcxPropertiesStore;
    Label1: TLabel;
    DateTimePicker_Assign: TDateTimePicker;
    procedure btnSrcClick(Sender: TObject);
    procedure btnDocClick(Sender: TObject);
    procedure btnOuputClick(Sender: TObject);
    procedure btnExecClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FCol_CustNo: Integer;
    FCol_CustName: Integer;
    FCol_CustFullName: Integer;
    FCol_TaxNo: Integer;
    FCol_Contact: Integer;
    FCol_Addr: Integer;
    FCol_CustPhone: Integer;
    FCol_CustFax: Integer;
    FCol_TE: Integer;
    procedure ClearNdx_XLS_COL;
    procedure ParseNdx_XLS_COL(ASheet: TXLSWorksheet);
  private
    FDocNdx: Integer;
    procedure Exec;
    procedure OpenSrc;
    procedure CheckOutputFolder;
    function  GetCustNo(ASheet: TXLSWorksheet; ASrcRow: Integer): string;
    function  GetOutputFileName(ACustNo: string): string;
    procedure MergeDoc(ASheet: TXLSWorksheet; ASrcRow: Integer; ASrcFileName, ADstFileName: string); overload;
    procedure MergeDoc(AExcelApp: Variant; ASheet: TXLSWorksheet; ASrcRow: Integer); overload;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

procedure TfmMain.btnDocClick(Sender: TObject);
var
  aFileName: string;
begin
  if JcExecOpenDialog_XLS(aFileName) then
    EditDoc.Text := aFileName;
end;

procedure TfmMain.btnExecClick(Sender: TObject);
begin
  Exec;
  ShowMessage('Done!');
end;

procedure TfmMain.btnOuputClick(Sender: TObject);
begin
  if cxShellBrowserDialog1.Execute then
    EditOutputFolder.Text := cxShellBrowserDialog1.Path;
end;

procedure TfmMain.btnSrcClick(Sender: TObject);
var
  aFileName: string;
begin
  if JcExecOpenDialog_XLS(aFileName) then
    EditSrc.Text := aFileName;
end;

procedure TfmMain.CheckOutputFolder;
begin
  if not DirectoryExists(EditOutputFolder.Text) then
    ForceDirectories(EditOutputFolder.Text);
end;

procedure TfmMain.ClearNdx_XLS_COL;
begin
  FCol_CustNo := -1;
  FCol_CustName := -1;
  FCol_CustFullName := -1;
  FCol_TaxNo := -1;
  FCol_Contact := -1;
  FCol_Addr := -1;
  FCol_CustPhone := -1;
  FCol_CustFax := -1;
  FCol_TE := -1;
end;

procedure TfmMain.Exec;
var
  aSheet: TXLSWorksheet;
  i: Integer;
  aCustNo, aFileName: string;
begin
  OpenSrc;
  CheckOutputFolder;
  aSheet := XLSReadWriteII5.Sheets[0];
  ClearNdx_XLS_COL;
  ParseNdx_XLS_COL(aSheet);
  ProgressBar1.Position := 0;
  ProgressBar1.Max := aSheet.LastRow - 1;
  FDocNdx := 0;

  for i := aSheet.FirstRow + 1 to aSheet.LastRow do
  begin
    ProgressBar1.Position := ProgressBar1.Position + 1;
    ProgressBar1.Position := ProgressBar1.Position - 1;
    ProgressBar1.Position := ProgressBar1.Position + 1;
    Application.ProcessMessages;

    Inc(FDocNdx);
    aCustNo := GetCustNo(ASheet, i);
    aFileName := GetOutputFileName(aCustNo);
    MergeDoc(aSheet, i, EditDoc.Text, aFileName);
  end;
end;

procedure TfmMain.FormCreate(Sender: TObject);
begin
  DateTimePicker_Assign.Date := Date;
end;

function TfmMain.GetCustNo(ASheet: TXLSWorksheet; ASrcRow: Integer): string;
begin
  if (FCol_CustNo <>-1) then
    Result := ASheet.AsString[FCol_CustNo, ASrcRow]
  else
    Result := '';
end;

function TfmMain.GetOutputFileName(ACustNo: string): string;
var
  aFileName, aFileExt: string;
  i: Integer;
begin
  aFileName := ExtractFileName(EditDoc.Text);
  i := Pos('.', aFileName);
  aFileName := Copy(aFileName, 1, i-1);
  aFileExt := ExtractFileExt(EditDoc.Text);
  Result := IncludeTrailingPathDelimiter(EditOutputFolder.Text) + Format('%s_%s%s', [aFileName, ACustNo, aFileExt]);
end;

procedure TfmMain.MergeDoc(AExcelApp: Variant; ASheet: TXLSWorksheet; ASrcRow: Integer);
var
  aText: string;
  aDate: TDateTime;

  procedure MergeData(ATagName: string; ACol: Integer);
  var
    aText: string;
  begin
    if ACol = -1 then Exit;
    aText := ASheet.AsString[ACol, ASrcRow];
    AExcelApp.Cells.Replace(ATagName, aText, xlPart, xlByRows, False, False);
  end;
begin
  //�m��[�ȥN]
  MergeData('<<�Ȥ�N��>>', FCol_CustNo);
  //�m��[�Ȥ�W��]
  MergeData('<<�Ȥ�W��>>', FCol_CustName);
  //�m��[�Ȥ����]
  MergeData('<<�Ȥ����>>', FCol_CustFullName);
  //�m��[�Τ@�s��]
  MergeData('<<�Τ@�s��>>', FCol_TaxNo);
  //�m��[�p���H]
  MergeData('<<�p���H>>', FCol_Contact);
  MergeData('<<�s���H>>', FCol_Contact);
  //�m��[�a�}]
  MergeData('<<�a�}>>', FCol_Addr);
  //�m��[�q��]
  MergeData('<<�q��>>', FCol_CustPhone);
  //�m��[�ǯu]
  MergeData('<<�ǯu>>', FCol_CustFax);
  //�m��[�V�m�v]
  MergeData('<<�V�m�v>>', FCol_TE);
  //�m��[���Ѥ��]
  AExcelApp.Cells.Replace('<<���Ѥ��>>', FormatDateTime('YYYY.MM.DD', Date), xlPart, xlByRows, False, False);
  //�m��[���w���]
  AExcelApp.Cells.Replace('<<���w���>>', FormatDateTime('YYYY.MM.DD', DateTimePicker_Assign.Date), xlPart, xlByRows, False, False);
  //�m��[�~���y����]
  aDate := EncodeDate(2018, 10, 15);
  aText := '''' + FormatDateTime('YYYYMMDD', aDate) + StrPadLeft(IntToStr(FDocNdx), 3, '0');
  AExcelApp.Cells.Replace('<<�~���y����>>', aText, xlPart, xlByRows, False, False);
  AExcelApp.Cells.Replace('<<���Ѥ���y����>>', aText, xlPart, xlByRows, False, False);
  //�m��[���w�~���y����]
  aDate := DateTimePicker_Assign.Date;
  aText := '''' + FormatDateTime('YYYYMMDD', aDate) + StrPadLeft(IntToStr(FDocNdx), 3, '0');
  AExcelApp.Cells.Replace('<<���w����y����>>', aText, xlPart, xlByRows, False, False);
end;

procedure TfmMain.MergeDoc(ASheet: TXLSWorksheet; ASrcRow: Integer; ASrcFileName, ADstFileName: string);
var
  aExcelApp: Variant;
begin
  aExcelApp := CreateOleObject('Excel.Application');
  aExcelApp.Visible := False;
  aExcelApp.WorkBooks.Open(ASrcFileName);
  aExcelApp.WorkSheets.Item[1].Activate;
  aExcelApp.DisplayAlerts := False;

  try
    MergeDoc(aExcelApp, ASheet, ASrcRow);
    aExcelApp.ActiveWorkBook.SaveAs(ADstFileName);
    aExcelApp.ActiveWorkBook.Close;
  finally
    aExcelApp.Quit;
  end;
end;

procedure TfmMain.OpenSrc;
var
  aSheet: TXLSWorksheet;
begin
  with XLSReadWriteII5 do
  begin
    Filename := EditSrc.Text;
    Read;
    aSheet := Sheets[0];
    //ShowMessage(IntToStr(aSheet.LastRow));
  end;
end;

procedure TfmMain.ParseNdx_XLS_COL(ASheet: TXLSWorksheet);
var
  i: Integer;
  aText: string;
begin
  with ASheet do
  begin
    for i := FirstCol to LastCol do
    begin
      aText := AsString[i, FirstRow];

      if (aText = '�Ȥ�N��') then
        FCol_CustNo := i
      else if (aText = '�Ȥ�W��') or (aText = '���q�W��') then
        FCol_CustName := i
      else if (aText = '�Ȥ����') or (aText = '���q�W��') then
        FCol_CustFullName := i
      else if (aText = '�Τ@�s��') then
        FCol_TaxNo := i
      else if (aText = '�p���H') or (aText = '�s���H') then
        FCol_Contact := i
      else if (aText = '�a�}') or (aText = '���q�a�}') then
        FCol_Addr := i
      else if (aText = '�q��') or (aText = '�p���q��') or (aText = '�s���q��') then
        FCol_CustPhone := i
      else if (aText = '�ǯu') then
        FCol_CustFax := i
      else if (aText = '�V�m�v') then
        FCol_TE := i
    end;
  end;
end;

end.
