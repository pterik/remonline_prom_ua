{$A8,B-,C+,D+,E-,F-,G+,H+,I+,J-,K-,L+,M-,N-,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{$WARN SYMBOL_DEPRECATED ON}
{$WARN SYMBOL_LIBRARY ON}
{$WARN SYMBOL_PLATFORM ON}
{$WARN SYMBOL_EXPERIMENTAL ON}
{$WARN UNIT_LIBRARY ON}
{$WARN UNIT_PLATFORM ON}
{$WARN UNIT_DEPRECATED ON}
{$WARN UNIT_EXPERIMENTAL ON}
{$WARN HRESULT_COMPAT ON}
{$WARN HIDING_MEMBER ON}
{$WARN HIDDEN_VIRTUAL ON}
{$WARN GARBAGE ON}
{$WARN BOUNDS_ERROR ON}
{$WARN ZERO_NIL_COMPAT ON}
{$WARN STRING_CONST_TRUNCED ON}
{$WARN FOR_LOOP_VAR_VARPAR ON}
{$WARN TYPED_CONST_VARPAR ON}
{$WARN ASG_TO_TYPED_CONST ON}
{$WARN CASE_LABEL_RANGE ON}
{$WARN FOR_VARIABLE ON}
{$WARN CONSTRUCTING_ABSTRACT ON}
{$WARN COMPARISON_FALSE ON}
{$WARN COMPARISON_TRUE ON}
{$WARN COMPARING_SIGNED_UNSIGNED ON}
{$WARN COMBINING_SIGNED_UNSIGNED ON}
{$WARN UNSUPPORTED_CONSTRUCT ON}
{$WARN FILE_OPEN ON}
{$WARN FILE_OPEN_UNITSRC ON}
{$WARN BAD_GLOBAL_SYMBOL ON}
{$WARN DUPLICATE_CTOR_DTOR ON}
{$WARN INVALID_DIRECTIVE ON}
{$WARN PACKAGE_NO_LINK ON}
{$WARN PACKAGED_THREADVAR ON}
{$WARN IMPLICIT_IMPORT ON}
{$WARN HPPEMIT_IGNORED ON}
{$WARN NO_RETVAL ON}
{$WARN USE_BEFORE_DEF ON}
{$WARN FOR_LOOP_VAR_UNDEF ON}
{$WARN UNIT_NAME_MISMATCH ON}
{$WARN NO_CFG_FILE_FOUND ON}
{$WARN IMPLICIT_VARIANTS ON}
{$WARN UNICODE_TO_LOCALE ON}
{$WARN LOCALE_TO_UNICODE ON}
{$WARN IMAGEBASE_MULTIPLE ON}
{$WARN SUSPICIOUS_TYPECAST ON}
{$WARN PRIVATE_PROPACCESSOR ON}
{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_CAST OFF}
{$WARN OPTION_TRUNCATED ON}
{$WARN WIDECHAR_REDUCED ON}
{$WARN DUPLICATES_IGNORED ON}
{$WARN UNIT_INIT_SEQ ON}
{$WARN LOCAL_PINVOKE ON}
{$WARN MESSAGE_DIRECTIVE ON}
{$WARN TYPEINFO_IMPLICITLY_ADDED ON}
{$WARN RLINK_WARNING ON}
{$WARN IMPLICIT_STRING_CAST ON}
{$WARN IMPLICIT_STRING_CAST_LOSS ON}
{$WARN EXPLICIT_STRING_CAST OFF}
{$WARN EXPLICIT_STRING_CAST_LOSS OFF}
{$WARN CVT_WCHAR_TO_ACHAR ON}
{$WARN CVT_NARROWING_STRING_LOST ON}
{$WARN CVT_ACHAR_TO_WCHAR ON}
{$WARN CVT_WIDENING_STRING_LOST ON}
{$WARN NON_PORTABLE_TYPECAST ON}
{$WARN XML_WHITESPACE_NOT_ALLOWED ON}
{$WARN XML_UNKNOWN_ENTITY ON}
{$WARN XML_INVALID_NAME_START ON}
{$WARN XML_INVALID_NAME ON}
{$WARN XML_EXPECTED_CHARACTER ON}
{$WARN XML_CREF_NO_RESOLVE ON}
{$WARN XML_NO_PARM ON}
{$WARN XML_NO_MATCHING_PARM ON}
{$WARN IMMUTABLE_STRINGS OFF}
unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, ComObj,
  Vcl.ComCtrls;

const RemontkaHeader : array [1..13] of string = ('���','�������','�����-���','������������','�������','���������','��������',
                           '����������� ������','���������� ����','�������','���� � ���������', '���������','���������');
PromHeader : array [1..22] of string = (
      '���_������','��������_�������','��������_�����','��������','���_������','����',
      '������','�������_���������','�����������_�����_������','�������_����','�����������_�����_���','������_�����������',
      '�������','������','�������������','������_�������������','�����_������','�����_����������',
      '�������������_������','����������_�������������','�������������_����������','�������������_������'
);
FileSeparator:char=chr(9);

type
  Mapping_rec = record
    RemontkaName, PromName:string;
    RemontkaNumber:integer;
    Quoted:boolean;
end;

type PriceRec= array [1..22] of string;
type
  TFormMain = class(TForm)
    MemoTxt: TMemo;
    BitBtnXLS: TBitBtn;
    BitBtnClose: TBitBtn;
    FileOpenDialog1: TFileOpenDialog;
    MemoLog: TMemo;
    BitBtnCSV: TBitBtn;
    PB: TProgressBar;
    procedure BitBtnCloseClick(Sender: TObject);
    procedure BitBtnXLSClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtnCSVClick(Sender: TObject);
    procedure FormDblClick(Sender: TObject);
  private
    { Private declarations }
    Mapping:array [1..22] of Mapping_rec;
    RemontkaText:array[1..22] of string;

    function isRemontkaHeaderCorrect(Where:integer; Value:string):boolean;
    function WritePromHeaders:string;
    function WriteRemontkaHeader: string;
    function CaseNumber(k:integer):string;
    procedure FillMapping;
    function PrintPromText(pRemText:array of string):string;
    function PlusQuotes(Str:string; isQuoted:boolean):string;
    procedure CopyMemoToXLS(FileName:string; Lines:integer);

  public
    { Public declarations }
  end;
var
  FormMain: TFormMain;

implementation

{$R *.dfm}

procedure TFormMain.BitBtnCloseClick(Sender: TObject);
begin
Close;
end;

procedure TFormMain.BitBtnXLSClick(Sender: TObject);
var FileName:string;
ExcelIn: Variant;
//ExcelOut:Variant;
FString:string;
PrintText:string;
FileName1, FileName2:string;
IsEmptyLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
I: Integer;
begin
MemoLog.Clear;
if FileOpenDialog1.Execute then
begin
    FileName:=FileOpenDialog1.FileName;
    MemoLog.Lines.Add('�������������� ���� '+FileName);
    try
      try
       ExcelIn := GetActiveOleObject('Excel.Application');
       except on EOLESysError do
       ExcelIn := CreateOleObject('Excel.Application');
      end;
      ExcelIn.Visible := False;
      ExcelIn.WindowState := -4140;
      ExcelIn.DisplayAlerts := False;
      ExcelIn.WorkBooks.Open(FileOpenDialog1.FileName, 0 , true);
      ExcelIn.WorkSheets[1].Activate;
      MemoTxt.Clear;
      MemoTxt.Lines.Add(WritePromHeaders);
      LineNumber:=1;
      PB.Position:=0;
      PB.Min:=1;
      PB.Max:=300;
      PB.Step:=1;
      PB.StepIt;
      for I := 1 to 13 do
      begin
        CellRow:=caseNumber(i);
        CellNum:='1';
        CellText:=Trim(ExcelIn.Range[CellRow+CellNum]);
        if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoLog.Lines.Add('�������� ��������� �����, ��������� �������� "������� �� ������.xls" �� remonline ��� ���     '
                                +CellRow+CellNum+'!'+'!'+CellText);
            ShowMessage('�������� ���� ���������, �� ������ �������� �� ������ "������� �����"'+chr(10)+chr(13)
                          +'������� �� ���� remonline ��� ��� � ��������� ���� �������� � ������� "����������"'+chr(10)+chr(13)
                          +'�������� ������� "�����", ���������(��� �������) ��������� ����� ������ "�������"');
            exit;
          end;
      end;
      LineNumber:=2;
      isEmptyLine:=false;
      while not IsEmptyLine do
      begin
      PB.StepIt;
      for I := 1 to 13 do
        begin
        CellRow:=caseNumber(i);
        CellNum:=IntToStr(LineNumber);
        CellText:=trim(ExcelIn.Range[CellRow+CellNum]);
        if (CellRow='D') and (length(CellText)=0) then IsEmptyLine:=true;
        RemontkaText[i]:=CellText;
        if LineNumber>50000 then IsEmptyLine:=true;  //������� ���� 50(00) ����� ����� �� ���� ������������
      end;
      if not IsEmptyLine then
      begin
      PrintText:=PrintPromText(RemontkaText);
      if PrintText<>'' then MemoTxt.Lines.Add(PrintText);
      end;
       inc(LineNumber);
      end;
    finally
      ExcelIn.ActiveWorkbook.Close;
      ExcelIn.Application.Quit;
    end;

  if LowerCase(ExtractFileExt(FileName))='.xls' then FileName:=LowerCase(FileName+'x');
  CopyMemoToXLS(ExtractFilePath(FileName)+'prom_'+ExtractFileName(FileName), LineNumber);
  MemoLog.Lines.Add('����� ����� ��� ��������: '+ExtractFilePath(FileName)+'prom_'+ExtractFileName(FileName));
  Pb.Position:=PB.Max;
  end;
end;

function TFormMain.CaseNumber(k: integer): string;
begin
case k of
      1: Result:='A';
      2: Result:='B';
      3: Result:='C';
      4: Result:='D';
      5: Result:='E';
      6: Result:='F';
      7: Result:='G';
      8: Result:='H';
      9: Result:='I';
      10: Result:='J';
      11: Result:='K';
      12: Result:='L';
      13: Result:='M';
      14: Result:='N';
      15: Result:='O';
      16: Result:='P';
      17: Result:='Q';
      18: Result:='R';
      19: Result:='S';
      20: Result:='T';
      21: Result:='U';
      22: Result:='V';
      23: Result:='W';
      24: Result:='X';
      25: Result:='Y';
      26: Result:='Z';
      27: Result:='AA';
      28: Result:='AB';
      29: Result:='AC';
      30: Result:='AD';
      31: Result:='AE';
      32: Result:='AF';
      33: Result:='AG';
      34: Result:='AH';
      35: Result:='AI';
      36: Result:='AJ';
      37: Result:='AK';
      38: Result:='AL';
      39: Result:='AM';
      40: Result:='AN';
      else Result:='ZZ';
end;
end;


procedure TFormMain.CopyMemoToXLS(FileName:string; Lines:integer);
var i:integer;
LineStr:string;
ItemsCntr, where, LineNumber :integer;
CellText, CellNum, CellRow:string;
ExcelOut:Variant;
begin
//PB.Max:=Lines*2;
//PB.Position:=Lines;
PB.StepIt;
try
    //���������, ��� �� ����������� Excel
    try
    ExcelOut := GetActiveOleObject('Excel.Application');
    except
    //���� ���, �� ���������
    on EOLESysError do
      ExcelOut := CreateOleObject('Excel.Application');
    end;
    ExcelOut.Visible := False;
    //��������� Excel �� ������ �����
    ExcelOut.WindowState := -4140;  //-4137
    //�� ���������� ��������������� ���������
    ExcelOut.DisplayAlerts := False;
    ExcelOut.WorkBooks.Add;
    ExcelOut.WorkSheets[1].Activate;
    ExcelOut.WorkBooks[1].WorkSheets[1].Name:='Export Products Sheet';
    MemoLog.Lines.Add(IntToStr(MemoTxt.Lines.Count)+' ����� ������������ � XLS');
    LineNumber:=1;
    CellNum:='';
    CellRow:='';
    for i:=0 to MemoTxt.Lines.Count-1 do
      begin
      LineStr:=MemoTxt.Lines[i];
      CellNum:=IntToStr(i+1);
      ItemsCntr:=0;
      while Pos(FileSeparator,LineStr)>0 do
        begin
        inc(ItemsCntr);
        if ItemsCntr>50 then break;
        CellRow:=caseNumber(ItemsCntr);
        where:= Pos(FileSeparator,LineStr);
        CellText:=Copy(LineStr,1,where-1);
        LineStr:=Copy(LineStr,where+1,length(LineStr));
        // MemoLog.Lines.Add(CellText+'"+"'+LineStr);
        // MemoLog.Lines.Add('Cell="'+CellRow+CellNum+'" ,value="'+CellText+'"');
        ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(i+1),ItemsCntr].Value:=CellText;
        end;
    PB.StepIt;
    end;
    try
     ExcelOut.WorkBooks[1].SaveAs(FileName);
    except on E:EFCreateError do
    MessageDlg('�����-���� '+FileName+' ������ � ��������� Excel, ��� ���� ������ ���� ���������.'+chr(10)+chr(13)+' �������� Excel ���� ������������� ���������', mtError, [mbOK],0);
    end;
  finally
  ExcelOut.ActiveWorkbook.Close;
  ExcelOut.Application.Quit;
  end;
end;

procedure TFormMain.BitBtnCSVClick(Sender: TObject);
var F:TextFile;
FileName:string;
CsvFileName:string;
Excel: Variant;
FString:string;
PrintText:string;
IsEmptyLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
I: Integer;
begin
MemoLog.Clear;
if FileOpenDialog1.Execute then
  begin
    FileName:=FileOpenDialog1.FileName;
    MemoLog.Lines.Add('�������������� ���� '+FileName);
    try
    try
     Excel := GetActiveOleObject('Excel.Application');
     except
    on EOLESysError do
      Excel := CreateOleObject('Excel.Application');
    end;
    Excel.Visible := false;
    Excel.WindowState := -4140;  //-4137
    Excel.DisplayAlerts := False;
    Excel.WorkBooks.Open(FileName, 0 , true);
    Excel.WorkSheets[1].Activate;
    PB.Position:=0;
    PB.Min:=1;
    PB.Max:=300;
    PB.Step:=1;
    PB.StepIt;
    MemoTxt.Clear;
    MemoTxt.Lines.Add(WritePromHeaders);
    LineNumber:=1;
    for I := 1 to 13 do
    begin
      CellRow:=caseNumber(i);
      CellNum:='1';
      CellText:=Trim(Excel.Range[CellRow+CellNum]);
      if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoTxt.Lines.Add('�������� ��������� �����, ��������� �������� "������� �� ������.xls" �� remonline ��� ��� '
                                +CellRow+CellNum+'!'+'!'+CellText);
            ShowMessage('�������� ���� ���������, �� ������ �������� �� ������ "������� �����"'+chr(10)+chr(13)
                          +'������� �� ���� remonline ��� ��� � ��������� ���� �������� � ������� "����������"'+chr(10)+chr(13)
                          +'�������� ������� "�����", ���������(��� �������) ��������� ����� ������ "�������"');
            Excel.ActiveWorkbook.Close;
            Excel.Application.Quit;
            break;
          end;
    end;
    LineNumber:=2;
    PB.StepIt;
    isEmptyLine:=false;
    while not IsEmptyLine do
      begin
      PB.StepIt;
      PrintText:='';
      for I := 1 to 13 do
        begin
        CellRow:=caseNumber(i);
        CellNum:=IntToStr(LineNumber);
        CellText:=trim(Excel.Range[CellRow+CellNum]);
        if (CellRow='D') and (length(CellText)=0) then IsEmptyLine:=true;
        RemontkaText[i]:=CellText;
        if LineNumber>5000 then IsEmptyLine:=true;  //������� ���� 50(00) ����� ����� �� ���� ������������
      end;
      if not IsEmptyLine then MemoTxt.Lines.Add(PrintPromText(RemontkaText));
      inc(LineNumber);
    end;
    // MemoLog.Lines.Add(ExtractFileExt(FileName));
    if LowerCase(ExtractFileExt(FileName))='.xls' then CsvFileName:=ExtractFilePath(FileName)+'prom_'+Copy(ExtractFileName(FileName),1,length(ExtractFileName(FileName))-3)+'csv';
    if LowerCase(ExtractFileExt(FileName))='.xlsx' then CsvFileName:=ExtractFilePath(FileName)+'prom_'+Copy(ExtractFileName(FileName),1,length(ExtractFileName(FileName))-4)+'csv';
    MemoLog.Lines.Add(IntToStr(MemoTxt.Lines.Count)+' ����� ���������� � CSV ���� '+CsvFileName);
    Pb.Position:=PB.Max;
    finally
    Excel.ActiveWorkbook.Close;
    Excel.Application.Quit;
    end;
    try
    MemoTxt.Lines.SaveToFile(CsvFileName);
    except on E:EFCreateError do
      begin
      MessageDlg('�����-���� ������ � ��������� Excel, ��� ���� ������ ���� ���������. �������� Excel ���� ������������� ���������', mtError, [mbOK],0);
      end;
    end;
  end;
end;

procedure TFormMain.FillMapping;
var i:integer;
begin
Mapping[1].PromName:= '���_������';
Mapping[1].Quoted:= false;
Mapping[2].PromName:= '��������_�������';
Mapping[2].Quoted:= false;
Mapping[3].PromName:= '��������_�����';
Mapping[3].Quoted:= false;
Mapping[4].PromName:= '��������';
Mapping[4].Quoted:= false;
Mapping[5].PromName:= '���_������';
Mapping[5].Quoted:= false;
Mapping[6].PromName:= '����';
Mapping[6].Quoted:= false;
Mapping[7].PromName:= '������';
Mapping[7].Quoted:= false;
Mapping[8].PromName:= '�������_���������';
Mapping[8].Quoted:= false;
Mapping[9].PromName:= '�����������_�����_������';
Mapping[9].Quoted:= false;
Mapping[10].PromName:= '�������_����';
Mapping[10].Quoted:= false;
Mapping[11].PromName:= '�����������_�����_���';
Mapping[11].Quoted:= false;
Mapping[12].PromName:= '������_�����������';
Mapping[12].Quoted:= true;
Mapping[13].PromName:= '�������';
Mapping[13].Quoted:= false;
Mapping[14].PromName:= '������';
Mapping[14].Quoted:= false;
Mapping[15].PromName:= '�������������';
Mapping[15].Quoted:= false;
Mapping[16].PromName:= '������_�������������';
Mapping[16].Quoted:= false;
Mapping[17].PromName:= '�����_������';
Mapping[17].Quoted:= true;
Mapping[18].PromName:= '�����_����������';
Mapping[18].Quoted:= false;
Mapping[19].PromName:= '�������������_������';
Mapping[19].Quoted:= false;
Mapping[20].PromName:= '����������_�������������';
Mapping[20].Quoted:= false;
Mapping[21].PromName:= '�������������_����������';
Mapping[21].Quoted:= false;
Mapping[22].PromName:= '�������������_������';
Mapping[22].Quoted:= false;
Mapping[1].RemontkaName:= '���';
Mapping[2].RemontkaName:= '������������';
Mapping[3].RemontkaName:= '';
Mapping[4].RemontkaName:= '������������';
Mapping[5].RemontkaName:= '';
Mapping[6].RemontkaName:= '���������';
Mapping[7].RemontkaName:= '������ UAH';
Mapping[8].RemontkaName:= '';
Mapping[9].RemontkaName:= '';
Mapping[10].RemontkaName:= '���������';
Mapping[11].RemontkaName:= '';
Mapping[12].RemontkaName:= '';
Mapping[13].RemontkaName:= '�������';
Mapping[14].RemontkaName:= '';
Mapping[15].RemontkaName:= '';
Mapping[16].RemontkaName:= '';
Mapping[17].RemontkaName:= '';
Mapping[18].RemontkaName:= '';
Mapping[19].RemontkaName:= '';
Mapping[20].RemontkaName:= '';
Mapping[21].RemontkaName:= '';
Mapping[22].RemontkaName:= '';
for i:=1 to 22 do Mapping[i].RemontkaNumber:=-999;
Mapping[1].RemontkaNumber:=1-1;
Mapping[2].RemontkaNumber:=4-1;
Mapping[4].RemontkaNumber:=4-1;
Mapping[5].RemontkaNumber:=-5;
Mapping[6].RemontkaNumber:=12-1;
Mapping[7].RemontkaNumber:=-7;
Mapping[8].RemontkaNumber:=-8;
Mapping[9].RemontkaNumber:=-9;
Mapping[10].RemontkaNumber:=11-1;
Mapping[11].RemontkaNumber:=-11;
Mapping[13].RemontkaNumber:=5-1;
Mapping[19].RemontkaNumber:=2-1;
//Mapping[22].RemontkaNumber:=6-1;
Mapping[22].RemontkaNumber:=-999;//�� ������������ ������������� ������
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
FillMapping;
end;

procedure TFormMain.FormDblClick(Sender: TObject);
begin
MemoTxt.Visible:=true;
end;

function TFormMain.isRemontkaHeaderCorrect(Where:integer; Value: string): boolean;
 begin
 if trim(Value) = RemontkaHeader[where] then Result:=true else Result:=false;
 if (where=1) and (Value <>'���') then Result:=false;
 if (where=2) and (Value <>'�������') then Result:=false;
end;

function TFormMain.PlusQuotes(Str: string; isQuoted: boolean): string;
begin
if isQuoted then Result:='"'+str+'"' else Result:=Str;
end;

function TFormMain.PrintPromText(pRemText: array of string): string;
var i, RemNumber:integer;
Price:Extended;
begin
Result:='';
if (Mapping[1].RemontkaNumber>=0) then Result:=PlusQuotes(pRemText[Mapping[1].RemontkaNumber],Mapping[i].Quoted);
for I := 2 to 22 do
  begin
    Result:=Result+FileSeparator;
    RemNumber:=Mapping[i].RemontkaNumber;
     case RemNumber of
    -999:;
    -5: Result:=Result+PlusQuotes('u',Mapping[i].Quoted);
    -7: Result:=Result+PlusQuotes('UAH',Mapping[i].Quoted);
    -8: Result:=Result+PlusQuotes('��.',Mapping[i].Quoted);
    -9: Result:=Result+PlusQuotes('1',Mapping[i].Quoted);
    -11: Result:=Result+PlusQuotes('2',Mapping[i].Quoted);
    4: if (StrToIntDef(pRemText[Mapping[i].RemontkaNumber],0)>0)
          then Result:=Result+PlusQuotes('+',Mapping[i].Quoted)
          else Result:=Result+PlusQuotes('-',Mapping[i].Quoted);
         //�������� ���������� �� ������� + ��� -
    else Result:=Result+PlusQuotes(pRemText[Mapping[i].RemontkaNumber],Mapping[i].Quoted);
        //�������� ���� 0 �� ����  0.00001
    if (i=6) then
      begin
      Price:=StrToFloatDef(pRemText[Mapping[i].RemontkaNumber],-1);
      if (Price = 0) then
        begin
        Result:='';
        MemoLog.Lines.Add('����� � ����� "'+pRemText[1]+'", �������� '+pRemText[2]);
        MemoLog.Lines.Add('����� �����������, ������� ����');
        exit;
        end;
      if (Price = -1) then
        begin
        Result:='';
        MemoLog.Lines.Add('����� � ����� "'+pRemText[1]+'", �������� '+pRemText[2]);
        MemoLog.Lines.Add('����� �����������, ������� ����������� ���� '+pRemText[6]+'.�������� ������������.');
        exit;
        end;
      end;
    end;
  end;
 end;

function TFormMain.WriteRemontkaHeader: string;
  var i:integer;
begin
Result:=RemontkaHeader[1];
for I := 2 to 13 do
  begin
     Result:= Result+FileSeparator+ RemontkaHeader[i];
  end;

end;

function TFormMain.WritePromHeaders: string;
var i:integer;
begin
Result:=PromHeader[1];
for I := 2 to 22 do
  begin
    Result:=Result+FileSeparator+PromHeader[i];
  end;

end;

end.