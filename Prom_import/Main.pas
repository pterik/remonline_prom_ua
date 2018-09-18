unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, ComObj;

const RemontkaHeader : array [1..12] of string = ('Код','Артикул','Штрих-код','Наименование','Остаток','Категория','Гарантия',
                           'Гарантийный период','Закупочная цена','Нулевая','Ремонтная','Розничная');
PromHeader : array [1..22] of string = (
      'Код_товара','Название_позиции','Ключевые_слова','Описание','Тип_товара','Цена',
      'Валюта','Единица_измерения','Минимальный_объем_заказа','Оптовая_цена','Минимальный_заказ_опт','Ссылка_изображения',
      'Наличие','Скидка','Производитель','Страна_производитель','Номер_группы','Адрес_подраздела',
      'Идентификатор_товара','Уникальный_идентификатор','Идентификатор_подраздела','Идентификатор_группы'
);
FileSeparator:char=chr(9);

type
  Mapping_rec = record
    RemontkaName, PromName:string;
    RemontkaNumber:integer;
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
    procedure BitBtnCloseClick(Sender: TObject);
    procedure BitBtnXLSClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtnCSVClick(Sender: TObject);
  private
    { Private declarations }
    Mapping:array [1..22] of Mapping_rec;
    RemontkaText:array[1..22] of string;

    function isRemontkaHeaderCorrect(Where:integer; Value:string):boolean;
    function WritePromHeaders:string;
    function WriteRemontkaHeader: string;
    function CaseNumber(k:integer):char;
    procedure FillMapping;
    function PrintPromText(pRemText:array of string):string;
    procedure CopyMemoToXLS;
    procedure CopyLinetoXLS(var LineStr:string);

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
var F:TextFile;
Excel: Variant;
ExcelOut:Variant;
FString:string;
PrintText, PriceLine:string;
FileName1, FileName2:string;
IsEmptyLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
I, PriceCntr: Integer;
Price:array of PriceRec;
begin
//Ветка создания XLS

//if FileOpenDialog1.Execute then
//  begin
    try
    try
    //проверяем, нет ли запущенного Excel
    Excel := GetActiveOleObject('Excel.Application');
    except
    //если нет, то запускаем
    on EOLESysError do
      Excel := CreateOleObject('Excel.Application');
    end;
    Excel.Visible := True;
    //Открывать Excel на полный экран
    Excel.WindowState := -4137;
    //не показывать предупреждающие сообщения
    Excel.DisplayAlerts := False;
    //Открываем рабочую книгу
    //Excel.WorkBooks.Open(FileOpenDialog1.FileName, 0 , true);
    Excel.WorkBooks.Open('D:\ost.xls', 0 , true);
    //Excel.Visible := False;
    Excel.WorkSheets[1].Activate;

    MemoTxt.Clear;
    MemoTxt.Lines.Add(WritePromHeaders);
    PriceCntr:=1;
    SetLength(Price,PriceCntr+1);
    PriceLine:='';
    for I := 1 to Length(PromHeader) do
      begin
       Price[1,i]:=PromHeader[i];
       PriceLine:=PriceLine+Price[1,i]+FileSeparator;
      end;
    LineNumber:=1;
    for I := 1 to length(RemontkaHeader)-1 do
    begin
      CellRow:=caseNumber(i);
      CellNum:='1';
      CellText:=Trim(Excel.Range[CellRow+CellNum]);
      if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoTxt.Lines.Add('Неверный заголовок файла, проведите выгрузку "Остатки на складе.xls" из remonline ещё раз '
                                +CellRow+CellNum+'!'+'!'+CellText);
            Excel.ActiveWorkbook.Close;
            Excel.Application.Quit;
            break;
          end;
    end;
    //CelNum:='2';
    //CellText:=Trim(Excel.Range['A'+CelNum]);
    LineNumber:=2;
    isEmptyLine:=false;
    while not IsEmptyLine do
      begin
      PrintText:='';
      inc(PriceCntr);
      SetLength(Price,PriceCntr+1);
      for I := 1 to 22 do Price[PriceCntr, i]:='';
      for I := 1 to 12 do
        begin
        CellRow:=caseNumber(i);
        CellNum:=IntToStr(LineNumber);
        CellText:=trim(Excel.Range[CellRow+CellNum]);
        if (CellRow='A') and (length(CellText)=0) then IsEmptyLine:=true;
        RemontkaText[i]:=CellText;
        Price[PriceCntr, i]:=CellText;
        if LineNumber>10 then IsEmptyLine:=true;  //Выходим если 50(00) строк чтобы не было зацикливания
      end;
      if not IsEmptyLine then
      begin
      MemoTxt.Lines.Add(PrintPromText(RemontkaText));
      //PriceLine:='';
      //for i:=1 to 22 do PriceLine:=PriceLine+Price[PriceCntr, i]+FileSeparator;

      end;
      //MemoTxt.Lines.Add(PriceLine);
       inc(LineNumber);
    end;
    ;
    try
    //ExcelOut.SaveAs('D:\prom_ost.xls');
    except on E:EFCreateError do
     begin
     MessageDlg('Прайс-лист открыт в программе Excel, или сбой работы этой программы. Закройте Excel либо перезагрузите компьютер', mtError, [mbOK],0);
    end;
    end;
    finally
      Excel.ActiveWorkbook.Close;
      Excel.Application.Quit;
    end;
    try
    try
    //проверяем, нет ли запущенного Excel
    ExcelOut:= GetActiveOleObject('Excel.Application');
    except
    //если нет, то запускаем
    on EOLESysError do
      ExcelOut:= CreateOleObject('Excel.Application');
    end;
    ExcelOut.Visible := True;
    //Открывать Excel на полный экран
    ExcelOut.WindowState := -4137;
    //не показывать предупреждающие сообщения
    ExcelOut.DisplayAlerts := False;
    //Открываем рабочую книгу
    ExcelOut.WorkBooks.Add;
    ExcelOut.WorkSheets[1].Activate;
    CopyMemoToXLS;
    finally
     ExcelOut.ActiveWorkbook.Close;
     ExcelOut.Application.Quit;
    end;
//  end;
end;

function TFormMain.CaseNumber(k: integer): char;
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
end;
end;

procedure TFormMain.CopyLinetoXLS(var LineStr:string);
var ItemsCntr, where:integer;
CellStr:string;
begin
ItemsCntr:=0;
while Pos(FileSeparator,LineStr)>0 do
begin
  inc(ItemsCntr);
  if ItemsCntr>100 then break;
  where:= Pos(FileSeparator,LineStr);
  CellStr:=Copy(LineStr,1,where-1);
  LineStr:=Copy(LineStr,where+1,length(LineStr));
  MemoLog.Lines.Add(CellStr+'"+"'+LineStr);
end;
end;

procedure TFormMain.CopyMemoToXLS;
var LinesCount:integer;
LineStr:string;
begin
LinesCount:=MemoTxt.Lines.Count;
MemoLog.Lines.Add(IntToStr(LinesCount));
LineStr:=MemoTxt.Lines[0];
CopyLineToXLS(LineStr);
end;

procedure TFormMain.BitBtnCSVClick(Sender: TObject);
var F:TextFile;
Excel: Variant;
FString:string;
PrintText:string;
IsEmptyLine:boolean;
CellText, CelNum, CellRow:string;
LineNumber:integer;
I: Integer;
begin
//if FileOpenDialog1.Execute then
//begin
    try
    //проверяем, нет ли запущенного Excel
    Excel := GetActiveOleObject('Excel.Application');
    except
    //если нет, то запускаем
    on EOLESysError do
      Excel := CreateOleObject('Excel.Application');
    end;
    Excel.Visible := True;
    //Открывать Excel на полный экран
    Excel.WindowState := -4137;
    //не показывать предупреждающие сообщения
    Excel.DisplayAlerts := False;
    //Открываем рабочую книгу
    //Excel.WorkBooks.Open(FileOpenDialog1.FileName, 0 , true);
    Excel.WorkBooks.Open('D:\ost.xls', 0 , true);
    //Excel.Visible := False;
    Excel.WorkSheets[1].Activate;
    MemoTxt.Clear;
    MemoTxt.Lines.Add(WritePromHeaders);
    LineNumber:=1;
    for I := 1 to length(RemontkaHeader)-1 do
    begin
      CellRow:=caseNumber(i);
      CelNum:='1';
      CellText:=Trim(Excel.Range[CellRow+CelNum]);
      if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoTxt.Lines.Add('Неверный заголовок файла, проведите выгрузку "Остатки на складе.xls" из remonline ещё раз '
                                +CellRow+CelNum+'!'+'!'+CellText);
            Excel.ActiveWorkbook.Close;
            Excel.Application.Quit;
            break;
          end;
    end;
    LineNumber:=2;
    isEmptyLine:=false;
    while not IsEmptyLine do
      begin
      PrintText:='';
      for I := 1 to 12 do
        begin
        CellRow:=caseNumber(i);
        CelNum:=IntToStr(LineNumber);
        CellText:=trim(Excel.Range[CellRow+CelNum]);
        if (CellRow='A') and (length(CellText)=0) then IsEmptyLine:=true;
        RemontkaText[i]:=CellText;
        if LineNumber>50 then IsEmptyLine:=true;  //Выходим если 50(00) строк чтобы не было зацикливания
      end;
    if not IsEmptyLine then MemoTxt.Lines.Add(PrintPromText(RemontkaText));
    inc(LineNumber);
    end;
    Excel.ActiveWorkbook.Close;
    Excel.Application.Quit;
    try
    MemoTxt.Lines.SaveToFile('D:\ost.csv');
    except on E:EFCreateError do
     begin
     MessageDlg('Прайс-лист открыт в программе Excel, или сбой работы этой программы. Закройте Excel либо перезагрузите компьютер', mtError, [mbOK],0);
    end;
    end;
//end;
end;

procedure TFormMain.FillMapping;
var i:integer;
begin
Mapping[1].PromName:= 'Код_товара';
Mapping[2].PromName:= 'Название_позиции';
Mapping[3].PromName:= 'Ключевые_слова';
Mapping[4].PromName:= 'Описание';
Mapping[5].PromName:= 'Тип_товара';
Mapping[6].PromName:= 'Цена';
Mapping[7].PromName:= 'Валюта';
Mapping[8].PromName:= 'Единица_измерения';
Mapping[9].PromName:= 'Минимальный_объем_заказа';
Mapping[10].PromName:= 'Оптовая_цена';
Mapping[11].PromName:= 'Минимальный_заказ_опт';
Mapping[12].PromName:= 'Ссылка_изображения';
Mapping[13].PromName:= 'Наличие';
Mapping[14].PromName:= 'Скидка';
Mapping[15].PromName:= 'Производитель';
Mapping[16].PromName:= 'Страна_производитель';
Mapping[17].PromName:= 'Номер_группы';
Mapping[18].PromName:= 'Адрес_подраздела';
Mapping[19].PromName:= 'Идентификатор_товара';
Mapping[20].PromName:= 'Уникальный_идентификатор';
Mapping[21].PromName:= 'Идентификатор_подраздела';
Mapping[22].PromName:= 'Идентификатор_группы';
Mapping[1].RemontkaName:= 'Код';
Mapping[2].RemontkaName:= 'Наименование';
Mapping[3].RemontkaName:= '';
Mapping[4].RemontkaName:= 'Наименование';
Mapping[5].RemontkaName:= '';
Mapping[6].RemontkaName:= 'Розничная';
Mapping[7].RemontkaName:= 'Всегда UAH';
Mapping[8].RemontkaName:= '';
Mapping[9].RemontkaName:= '';
Mapping[10].RemontkaName:= 'Ремонтная';
Mapping[11].RemontkaName:= '';
Mapping[12].RemontkaName:= '';
Mapping[13].RemontkaName:= 'Остаток';
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
Mapping[19].RemontkaNumber:=1-1;
//Mapping[22].RemontkaNumber:=6-1;
Mapping[22].RemontkaNumber:=-999;//Не обрабатываем идентификатор группы
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
FillMapping;
end;

function TFormMain.isRemontkaHeaderCorrect(Where:integer; Value: string): boolean;
 begin
 if trim(Value) = RemontkaHeader[where] then Result:=true else Result:=false;
end;

function TFormMain.PrintPromText(pRemText: array of string): string;
var i, RemNumber:integer;
begin
Result:='';
if (Mapping[1].RemontkaNumber>=0) then Result:=pRemText[Mapping[1].RemontkaNumber];
for I := 2 to 22 do
  begin
    Result:=Result+FileSeparator;
    RemNumber:=Mapping[i].RemontkaNumber;
    case RemNumber of
    -999:;
    -5: Result:=Result+'u';
    -7: Result:=Result+'UAH';
    -8: Result:=Result+'шт.';
    -9: Result:=Result+'1';
    -11: Result:=Result+'2';
    4: begin
       try
       if (StrToInt(pRemText[Mapping[i].RemontkaNumber])>0) then Result:=Result+'+' else Result:=Result+'-' ;
         //Заменяем количество на Наличие + или -
         except on E:EConvertError do Result:=Result+'-';
       end;
    end
    else Result:=Result+pRemText[Mapping[i].RemontkaNumber];
    end;
  end;
 end;

function TFormMain.WriteRemontkaHeader: string;
  var i:integer;
begin
Result:=RemontkaHeader[1];
for I := 2 to Length(RemontkaHeader) do
  begin
     Result:= Result+FileSeparator+ RemontkaHeader[i];
  end;

end;

function TFormMain.WritePromHeaders: string;
var i:integer;
begin
Result:=PromHeader[1];
for I := 2 to Length(PromHeader) do
  begin
    Result:=Result+FileSeparator+PromHeader[i];
  end;

end;

end.
