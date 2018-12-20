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
  //置換[客代]
  MergeData('<<客戶代號>>', FCol_CustNo);
  //置換[客戶名稱]
  MergeData('<<客戶名稱>>', FCol_CustName);
  //置換[客戶全稱]
  MergeData('<<客戶全稱>>', FCol_CustFullName);
  //置換[統一編號]
  MergeData('<<統一編號>>', FCol_TaxNo);
  //置換[聯絡人]
  MergeData('<<聯絡人>>', FCol_Contact);
  MergeData('<<連絡人>>', FCol_Contact);
  //置換[地址]
  MergeData('<<地址>>', FCol_Addr);
  //置換[電話]
  MergeData('<<電話>>', FCol_CustPhone);
  //置換[傳真]
  MergeData('<<傳真>>', FCol_CustFax);
  //置換[訓練師]
  MergeData('<<訓練師>>', FCol_TE);
  //置換[今天日期]
  AExcelApp.Cells.Replace('<<今天日期>>', FormatDateTime('YYYY.MM.DD', Date), xlPart, xlByRows, False, False);
  //置換[指定日期]
  AExcelApp.Cells.Replace('<<指定日期>>', FormatDateTime('YYYY.MM.DD', DateTimePicker_Assign.Date), xlPart, xlByRows, False, False);
  //置換[年月日流水號]
  aDate := EncodeDate(2018, 10, 15);
  aText := '''' + FormatDateTime('YYYYMMDD', aDate) + StrPadLeft(IntToStr(FDocNdx), 3, '0');
  AExcelApp.Cells.Replace('<<年月日流水號>>', aText, xlPart, xlByRows, False, False);
  AExcelApp.Cells.Replace('<<今天日期流水號>>', aText, xlPart, xlByRows, False, False);
  //置換[指定年月日流水號]
  aDate := DateTimePicker_Assign.Date;
  aText := '''' + FormatDateTime('YYYYMMDD', aDate) + StrPadLeft(IntToStr(FDocNdx), 3, '0');
  AExcelApp.Cells.Replace('<<指定日期流水號>>', aText, xlPart, xlByRows, False, False);
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

      if (aText = '客戶代號') then
        FCol_CustNo := i
      else if (aText = '客戶名稱') or (aText = '公司名稱') then
        FCol_CustName := i
      else if (aText = '客戶全稱') or (aText = '公司名稱') then
        FCol_CustFullName := i
      else if (aText = '統一編號') then
        FCol_TaxNo := i
      else if (aText = '聯絡人') or (aText = '連絡人') then
        FCol_Contact := i
      else if (aText = '地址') or (aText = '公司地址') then
        FCol_Addr := i
      else if (aText = '電話') or (aText = '聯絡電話') or (aText = '連絡電話') then
        FCol_CustPhone := i
      else if (aText = '傳真') then
        FCol_CustFax := i
      else if (aText = '訓練師') then
        FCol_TE := i
    end;
  end;
end;

end.
