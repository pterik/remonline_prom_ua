unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, ComObj,
  Vcl.ComCtrls, SQLiteTable3, jpeg, System.UITypes, Vcl.ExtCtrls;

const RemontkaHeader : array [1..13] of string = ('Код','Артикул','Штрих-код','Наименование','Остаток','Категория','Гарантия',
                           'Гарантийный период','Закупочная цена','Нулевая','цена в интернете', 'Розничная','Ремонтная');
PromHeader : array [1..23] of string = (
      'Код_товара','Название_позиции','Ключевые_слова','Описание','Тип_товара','Цена',
      'Валюта','Единица_измерения','Минимальный_объем_заказа','Оптовая_цена','Минимальный_заказ_опт','Ссылка_изображения',
      'Наличие','Количество', 'Скидка','Производитель','Страна_производитель','Номер_группы','Адрес_подраздела',
      'Идентификатор_товара','Уникальный_идентификатор','Идентификатор_подраздела','Идентификатор_группы'
);
PromExpandHeader : array [1..42] of string = (
      'Код_товара','Название_позиции','Ключевые_слова','Описание','Тип_товара',
      'Цена','Валюта','Единица_измерения','Минимальный_объем_заказа','Оптовая_цена',
      'Минимальный_заказ_опт','Ссылка_изображения','Наличие', 'Количество','Номер_группы',
      'Название_группы','Адрес_подраздела','Возможность_поставки','Срок_поставки', 'Способ_упаковки',
      'Уникальный_идентификатор','Идентификатор_товара','Идентификатор_подраздела', 'Идентификатор_группы','Производитель',
      'Страна_производитель','Скидка','ID_группы_разновидностей', 'Метки','Продукт_на_сайте',
      'Название_Характеристики','Измерение_Характеристики','Значение_Характеристики',
      'Название_Характеристики','Измерение_Характеристики','Значение_Характеристики',
      'Название_Характеристики','Измерение_Характеристики','Значение_Характеристики',
      'Название_Характеристики','Измерение_Характеристики','Значение_Характеристики'
);


FileSeparator:char=';';

type
  Mapping_rec = record
    RemontkaName, PromName:string;
    RemontkaNumber:integer;
    Quoted:boolean;
end;

type PriceRec= array [1..23] of string;
type
  TFormMain = class(TForm)
    MemoTxt: TMemo;
    BitBtnXLS: TBitBtn;
    BitBtnClose: TBitBtn;
    FileOpenDialog1: TFileOpenDialog;
    MemoLog: TMemo;
    BitBtnCSV: TBitBtn;
    PB: TProgressBar;
    CheckBoxZeroPrice: TCheckBox;
    CheckBoxZeroOstatki: TCheckBox;
    FileOpenDialog2: TFileOpenDialog;
    procedure BitBtnCloseClick(Sender: TObject);
    procedure BitBtnXLSClick(Sender: TObject);
    procedure BitBtnCSVClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    Mapping:array [1..23] of Mapping_rec;
    DBName:string;
    //sltb: TSQLIteTable;
    Flags: TReplaceFlags;
    function  isRemontkaHeaderCorrect(Where:integer; Value:string):boolean;
    function  isPromHeaderCorrect(Where: integer; Value: string): boolean;
    function  isPromExpandHeaderCorrect(Where: integer; Value: string): boolean;
    function  WriteRemontkaHeader: string;
    function  CaseNumber(k:integer):string;
    procedure FillMapping;
    function  PrintPromText(pPromText:array of string):string;
    function  PlusQuotes(Str:string; isQuoted:boolean):string;
    function  TrimSeparator(const Str:string):string;
    function  QuotesForSQL(const Str:string):string;
    procedure CopyMemoToXLS(FileName:string; Lines:integer);
    procedure CopySQLiteToXLS(FileName:string);
    procedure FormDblClick(Sender: TObject);
    procedure EmptySQLite;
    procedure SavePromToSQLite(pPromArray:array of string);
    procedure SaveRemontkaTextToSQLite(pRemArray:array of string);
    procedure LoadRemontkaToSQLite;
    procedure LoadPromToSQLite;
    function  LogText(const PText: array of string): string;
    function WritePromExpandHeaders: string;
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
var
FileName:string;
ExcelIn: Variant;
Price:Extended;
FString:string;
PrintText:string;
FileName1, FileName2:string;
IsEmptyLine, IsExcludedLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
Amount, I: Integer;
begin
MemoLog.Clear;
if not (FileOpenDialog1.Execute) then exit;
EmptySQLite;
FileName:=FileOpenDialog1.FileName;
MemoLog.Lines.Add('Обрабатывается файл '+FileName);
LoadRemontkaToSQLite;
if LowerCase(ExtractFileExt(FileName))='.xls' then FileName:=LowerCase(FileName+'x');
Pb.Position:=PB.Max div 2;
MemoLog.Lines.Add('Остатки обработаны, выберите файл prom.ua для загрузки. (С именем export*.xlsx)');
if not FileOpenDialog2.Execute then exit;
MemoLog.Lines.Add('Выбран файл для загрузки :'+FileOpenDialog2.FileName);
LoadPromToSQLite;
//CopyMemoToXLS(ExtractFilePath(FileName)+'prom_'+ExtractFileName(FileName), LineNumber);
CopySQLiteToXLS(ExtractFilePath(FileName)+'DB_prom_'+ExtractFileName(FileName));
MemoLog.Lines.Add('Остатки обработаны, файл создан '+ExtractFilePath(FileName)+'DB_prom_'+ExtractFileName(FileName));
Pb.Position:=PB.Max;
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
      41: Result:='AO';
      42: Result:='AP';
      43: Result:='AQ';
      44: Result:='AR';
      45: Result:='AS';
      46: Result:='AT';
      47: Result:='AU';
      48: Result:='AV';
      49: Result:='AW';
      50: Result:='AX';
      51: Result:='AY';
      52: Result:='AZ';
      53: Result:='BA';
      54: Result:='BB';
      55: Result:='BC';
      56: Result:='BD';
      57: Result:='BE';
      58: Result:='BF';
      59: Result:='BG';
      60: Result:='BH';
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
    //проверяем, нет ли запущенного Excel
    try
    ExcelOut := GetActiveOleObject('Excel.Application');
    except
    //если нет, то запускаем
    on EOLESysError do
      ExcelOut := CreateOleObject('Excel.Application');
    end;
    ExcelOut.Visible := False;
    //Открывать Excel на полный экран
    ExcelOut.WindowState := -4140;  //-4137
    //не показывать предупреждающие сообщения
    ExcelOut.DisplayAlerts := False;
    ExcelOut.WorkBooks.Add;
    ExcelOut.WorkSheets[1].Activate;
    ExcelOut.WorkBooks[1].WorkSheets[1].Name:='Export Products Sheet';
    MemoLog.Lines.Add(IntToStr(MemoTxt.Lines.Count)+' строк экспортируем в XLS');
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
    MessageDlg('Прайс-лист '+FileName+' открыт в программе Excel, или сбой работы этой программы.'+chr(10)+chr(13)+' Закройте Excel либо перезагрузите компьютер', mtError, [mbOK],0);
    end;
  finally
  ExcelOut.ActiveWorkbook.Close;
  ExcelOut.Application.Quit;
  end;
end;

procedure TFormMain.CopySQLiteToXLS(FileName: string);
var i, LineNumber:integer;
str, Value:string;
ExcelOut:Variant;
S3DB:TSQLiteDatabase;
STBL: TSQLIteTable;
DValue:double;
begin
PB.StepIt;
try
    try
    ExcelOut := GetActiveOleObject('Excel.Application');
    except
    on EOLESysError do
      ExcelOut := CreateOleObject('Excel.Application');
    end;
    ExcelOut.Visible := False;
    ExcelOut.WindowState := -4140;
    ExcelOut.DisplayAlerts := False;
    ExcelOut.WorkBooks.Add;
    ExcelOut.WorkSheets[1].Activate;
    ExcelOut.WorkBooks[1].WorkSheets[1].Name:='Export Products Sheet';
    S3DB := TSQLiteDatabase.Create(DBName);
    STBL := S3DB.GetTable('SELECT * FROM vw_Items');
    MemoLog.Lines.Add(IntToStr(STBL.Count)+' строк экспортируем в XLS');
    for i:=1 to length(PromExpandHeader) do ExcelOut.WorkBooks[1].WorkSheets[1].Cells[1,i].Value:=PromExpandHeader[i];
    LineNumber:=1;
    while not STBL.EOF do
      begin
      inc(LineNumber);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),1].Value:=''''+STBL.FieldAsString(STBL.FieldIndex['Product_Code']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),2].Value:=STBL.FieldAsString(STBL.FieldIndex['Position_Name']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),3].Value:=STBL.FieldAsString(STBL.FieldIndex['Keywords']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),4].Value:=STBL.FieldAsString(STBL.FieldIndex['Description']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),5].Value:=STBL.FieldAsString(STBL.FieldIndex['Product_type']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),6].Value:=STBL.FieldAsString(STBL.FieldIndex['price']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),7].Value:=STBL.FieldAsString(STBL.FieldIndex['Currency']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),8].Value:=STBL.FieldAsString(STBL.FieldIndex['Unit_of_measurement']);

      if (STBL.FieldAsDouble(STBL.FieldIndex['Minimum_size_Order'])>0.00001)
        then Value:=STBL.FieldAsString(STBL.FieldIndex['Minimum_size_Order'])
        else Value:='';
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),9].Value:=Value;

      if (STBL.FieldAsDouble(STBL.FieldIndex['Wholesale_price'])>0.00001)
      then ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),10].Value:=STBL.FieldAsString(STBL.FieldIndex['Wholesale_price'])
      else ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),10].Value:=STBL.FieldAsString(STBL.FieldIndex['price']);

      if (STBL.FieldAsDouble(STBL.FieldIndex['Min_Order_Opt'])>0.00001)
        then Value:=STBL.FieldAsString(STBL.FieldIndex['Min_Order_Opt'])
        else Value:='''2';
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),11].Value:=Value;

      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),12].Value:=STBL.FieldAsString(STBL.FieldIndex['Image_Link']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),13].Value:=STBL.FieldAsString(STBL.FieldIndex['Availability']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),14].Value:=STBL.FieldAsString(STBL.FieldIndex['Amount']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),15].Value:=STBL.FieldAsString(STBL.FieldIndex['Group_number']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),16].Value:=STBL.FieldAsString(STBL.FieldIndex['Group_name']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),17].Value:=STBL.FieldAsString(STBL.FieldIndex['Division_Address']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),18].Value:=STBL.FieldAsString(STBL.FieldIndex['Possibility_of_delivery']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),19].Value:=STBL.FieldAsString(STBL.FieldIndex['Delivery_period']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),20].Value:=STBL.FieldAsString(STBL.FieldIndex['Packing_Mode']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),21].Value:=STBL.FieldAsString(STBL.FieldIndex['Unique_identificator']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),22].Value:=STBL.FieldAsString(STBL.FieldIndex['Product_id']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),23].Value:=STBL.FieldAsString(STBL.FieldIndex['Subdivision_id']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),24].Value:=STBL.FieldAsString(STBL.FieldIndex['Group_id']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),25].Value:=STBL.FieldAsString(STBL.FieldIndex['Manufacturer']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),26].Value:=STBL.FieldAsString(STBL.FieldIndex['Producing_country']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),27].Value:=STBL.FieldAsString(STBL.FieldIndex['Discount']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),28].Value:=STBL.FieldAsString(STBL.FieldIndex['Species_Group_ID']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),29].Value:=STBL.FieldAsString(STBL.FieldIndex['Tags']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),30].Value:=STBL.FieldAsString(STBL.FieldIndex['Product_on_Site']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),31].Value:=STBL.FieldAsString(STBL.FieldIndex['Name1_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),32].Value:=STBL.FieldAsString(STBL.FieldIndex['Measurement1_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),33].Value:=STBL.FieldAsString(STBL.FieldIndex['Value1_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),34].Value:=STBL.FieldAsString(STBL.FieldIndex['Name2_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),35].Value:=STBL.FieldAsString(STBL.FieldIndex['Measurement2_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),36].Value:=STBL.FieldAsString(STBL.FieldIndex['Value2_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),37].Value:=STBL.FieldAsString(STBL.FieldIndex['Name3_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),38].Value:=STBL.FieldAsString(STBL.FieldIndex['Measurement3_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),39].Value:=STBL.FieldAsString(STBL.FieldIndex['Value3_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),40].Value:=STBL.FieldAsString(STBL.FieldIndex['Name4_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),41].Value:=STBL.FieldAsString(STBL.FieldIndex['Measurement4_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Cells[IntToStr(LineNumber),42].Value:=STBL.FieldAsString(STBL.FieldIndex['Value4_Characteristics']);
      ExcelOut.WorkBooks[1].WorkSheets[1].Rows[IntToStr(LineNumber)].RowHeight:=16;
      STBL.Next;
      PB.StepIt;
    end;
    try
     ExcelOut.WorkBooks[1].SaveAs(FileName);
    except on E:EFCreateError do
    MessageDlg('Прайс-лист '+FileName+' открыт в программе Excel, или сбой работы этой программы.'+chr(10)+chr(13)+' Закройте Excel либо перезагрузите компьютер', mtError, [mbOK],0);
    end;
  finally
  ExcelOut.ActiveWorkbook.Close;
  ExcelOut.Application.Quit;
  STBL.Free;
  S3DB.Free;
  end;
end;

procedure TFormMain.BitBtnCSVClick(Sender: TObject);
var
RemontkaText: array[1..23] of string;
FileName:string;
CsvFileName:string;
Excel: Variant;
FString:string;
Price:real;
Amount:integer;
IsEmptyLine, isExcludedLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
I: Integer;
begin
MemoLog.Clear;
if FileOpenDialog1.Execute then
  begin
    FileName:=FileOpenDialog1.FileName;
    MemoLog.Lines.Add('Обрабатывается файл '+FileName);
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
    //MemoTxt.Lines.Add(WritePromHeaders);
    LineNumber:=1;
    for I := 1 to 13 do
    begin
      CellRow:=caseNumber(i);
      CellNum:='1';
      CellText:=Trim(Excel.Range[CellRow+CellNum]);
      if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoTxt.Lines.Add('Неверный заголовок файла, проведите выгрузку "Остатки на складе.xls" из remonline ещё раз '
                                +CellRow+CellNum+'!'+'!'+CellText);
            ShowMessage('Неверный файл остатоков, он создан нажатием на кнопку "Создать отчёт"'+chr(10)+chr(13)
                          +'Зайдите на сайт remonline ещё раз и выгрузите файл остатков с помощью "бутерброда"'+chr(10)+chr(13)
                          +'Выберите вкладку "Склад", бутерброд(три полоски) находится возле Строки "Наличие"');
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
    isExcludedLine:=false;
    PB.StepIt;
    for I := 1 to 13 do
      begin
      CellRow:=caseNumber(i);
      CellNum:=IntToStr(LineNumber);
      CellText:=trim(Excel.Range[CellRow+CellNum]);
      RemontkaText[i]:=TrimSeparator(CellText);
      if LineNumber>5000 then IsEmptyLine:=true;  //Выходим если 50(00) строк чтобы не было зацикливания
      end;
    if  (length(RemontkaText[1])=0)and(length(RemontkaText[2])=0)
           and (length(RemontkaText[3])=0)and(length(RemontkaText[4])=0)
      then
        begin
        //LogRemText(RemontkaText);
        //MemoLog.Lines.Add('Найдена пустая строка');
        IsEmptyLine:=true;
        Continue;
        end;
    Amount:=StrToIntDef(RemontkaText[5],-1);
    if (Amount = 0) then
      begin
      if not CheckBoxZeroOstatki.Checked then LogText(RemontkaText);
      if not CheckBoxZeroOstatki.Checked then MemoLog.Lines.Add('Товар исключается, нулевое количество. Код "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
      isExcludedLine:=true;
      end;
    if (Amount = -1) then
      begin
      LogText(RemontkaText);
      MemoLog.Lines.Add('Товар с кодом "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
      MemoLog.Lines.Add('Товар исключается, количество="'+RemontkaText[5]+'" не является числом. Сообщите разработчику.');
      IsExcludedLine:=true;
      end;
    Price:=StrToFloatDef(RemontkaText[11],-1);
    if (Price = 0) then
      begin
      if not CheckBoxZeroPrice.Checked then LogText(RemontkaText);
      if not CheckBoxZeroPrice.Checked then MemoLog.Lines.Add('Товар исключается, нулевая цена. Код "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
      IsExcludedLine:=true;
      end;
    if (Price = -1) then
      begin
      LogText(RemontkaText);
      MemoLog.Lines.Add('Товар с кодом "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
      MemoLog.Lines.Add('Товар исключается, цена ="'+RemontkaText[11]+'" отображается неверно. Сообщите разработчику.');
      isExcludedLine:=true;
      end;
    if not IsEmptyLine and not isExcludedLine then MemoTxt.Lines.Add(PrintPromText(RemontkaText));
    inc(LineNumber);
    end;
    if LowerCase(ExtractFileExt(FileName))='.xls' then CsvFileName:=ExtractFilePath(FileName)+'prom_'+Copy(ExtractFileName(FileName),1,length(ExtractFileName(FileName))-3)+'csv';
    if LowerCase(ExtractFileExt(FileName))='.xlsx' then CsvFileName:=ExtractFilePath(FileName)+'prom_'+Copy(ExtractFileName(FileName),1,length(ExtractFileName(FileName))-4)+'csv';
    MemoLog.Lines.Add(IntToStr(MemoTxt.Lines.Count)+' строк перенесено в CSV файл '+CsvFileName);
    Pb.Position:=PB.Max;
    finally
    Excel.ActiveWorkbook.Close;
    Excel.Application.Quit;
    end;
    try
    MemoTxt.Lines.SaveToFile(CsvFileName, TEncoding.UTF8);
    except on E:EFCreateError do
      begin
      MessageDlg('Прайс-лист открыт в программе Excel, или сбой работы этой программы. Закройте Excel либо перезагрузите компьютер', mtError, [mbOK],0);
      end;
    end;
  end;
end;

procedure TFormMain.FillMapping;
var i:integer;
begin
Mapping[1].PromName:= 'Код_товара';
Mapping[1].Quoted:= false;
Mapping[2].PromName:= 'Название_позиции';
Mapping[2].Quoted:= false;
Mapping[3].PromName:= 'Ключевые_слова';
Mapping[3].Quoted:= false;
Mapping[4].PromName:= 'Описание';
Mapping[4].Quoted:= false;
Mapping[5].PromName:= 'Тип_товара';
Mapping[5].Quoted:= false;
Mapping[6].PromName:= 'Цена';
Mapping[6].Quoted:= false;
Mapping[7].PromName:= 'Валюта';
Mapping[7].Quoted:= false;
Mapping[8].PromName:= 'Единица_измерения';
Mapping[8].Quoted:= false;
Mapping[9].PromName:= 'Минимальный_объем_заказа';
Mapping[9].Quoted:= false;
Mapping[10].PromName:= 'Оптовая_цена';
Mapping[10].Quoted:= false;
Mapping[11].PromName:= 'Минимальный_заказ_опт';
Mapping[11].Quoted:= false;
Mapping[12].PromName:= 'Ссылка_изображения';
Mapping[12].Quoted:= true;
Mapping[13].PromName:= 'Наличие';
Mapping[13].Quoted:= false;
Mapping[14].PromName:= 'Количество';
Mapping[14].Quoted:= false;
Mapping[15].PromName:= 'Скидка';
Mapping[15].Quoted:= false;
Mapping[16].PromName:= 'Производитель';
Mapping[16].Quoted:= false;
Mapping[17].PromName:= 'Страна_производитель';
Mapping[17].Quoted:= false;
Mapping[18].PromName:= 'Номер_группы';
Mapping[18].Quoted:= true;
Mapping[19].PromName:= 'Адрес_подраздела';
Mapping[19].Quoted:= false;
Mapping[20].PromName:= 'Идентификатор_товара';
Mapping[20].Quoted:= false;
Mapping[21].PromName:= 'Уникальный_идентификатор';
Mapping[21].Quoted:= false;
Mapping[22].PromName:= 'Идентификатор_подраздела';
Mapping[22].Quoted:= false;
Mapping[23].PromName:= 'Идентификатор_группы';
Mapping[23].Quoted:= false;
Mapping[1].RemontkaName:= 'Код';
Mapping[2].RemontkaName:= 'Артикул';
Mapping[3].RemontkaName:= 'Штрих-код';
Mapping[4].RemontkaName:= 'Наименование';
Mapping[5].RemontkaName:= 'Остаток';
Mapping[6].RemontkaName:= 'Категория';
Mapping[7].RemontkaName:= 'Гарантия';
Mapping[8].RemontkaName:= 'Гарантийный период';
Mapping[9].RemontkaName:= 'Закупочная цена';
Mapping[10].RemontkaName:= 'Нулевая';
Mapping[11].RemontkaName:= 'Цена в Интернете';
Mapping[12].RemontkaName:= 'Розничная';
Mapping[13].RemontkaName:= 'Ремонтная';
Mapping[14].RemontkaName:= '';
Mapping[15].RemontkaName:= '';
Mapping[16].RemontkaName:= '';
Mapping[17].RemontkaName:= '';
Mapping[18].RemontkaName:= '';
Mapping[19].RemontkaName:= '';
Mapping[20].RemontkaName:= '';
Mapping[21].RemontkaName:= '';
Mapping[22].RemontkaName:= '';
Mapping[23].RemontkaName:= '';
for i:=1 to 23 do Mapping[i].RemontkaNumber:=-999;
Mapping[1].RemontkaNumber:=1-1;
Mapping[2].RemontkaNumber:=4-1;
Mapping[4].RemontkaNumber:=4-1;
Mapping[5].RemontkaNumber:=-5;
Mapping[6].RemontkaNumber:=11-1;
Mapping[7].RemontkaNumber:=-7;
Mapping[8].RemontkaNumber:=-8;
Mapping[9].RemontkaNumber:=-9;
Mapping[10].RemontkaNumber:=11-1;//Исключаем до выяснения с Заказчиком
Mapping[11].RemontkaNumber:=-11;
Mapping[13].RemontkaNumber:=5-1;
Mapping[14].RemontkaNumber:=5-1;
Mapping[20].RemontkaNumber:=2-1;
//Mapping[22].RemontkaNumber:=6-1;
Mapping[23].RemontkaNumber:=-999;//Не обрабатываем идентификатор группы
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
DBName:=ExtractFilepath(application.exename) + 'database.sqlite3';
FillMapping;
end;

procedure TFormMain.FormDblClick(Sender: TObject);
begin
MemoTxt.Visible:=true;
end;

function TFormMain.isRemontkaHeaderCorrect(Where:integer; Value: string): boolean;
 begin
 if trim(Value) = RemontkaHeader[where] then Result:=true else Result:=false;
 if (where=1) and (Value <>'Код') then Result:=false;
 if (where=2) and (Value <>'Артикул') then Result:=false;
end;

function TFormMain.isPromExpandHeaderCorrect(Where: integer;
  Value: string): boolean;
begin
 if trim(Value) = PromExpandHeader[where] then Result:=true else Result:=false;
 if (where=1) and (Value <>'Код_товара') then Result:=false;
 if (where=2) and (Value <>'Название_позиции') then Result:=false;
end;

function TFormMain.isPromHeaderCorrect(Where:integer; Value: string): boolean;
 begin
 if trim(Value) = PromHeader[where] then Result:=true else Result:=false;
 if (where=1) and (Value <>'Название_позиции') then Result:=false;
 if (where=2) and (Value <>'Ключевые_слова') then Result:=false;
end;

procedure TFormMain.LoadPromToSQLite;
var
PromExpandText: array[1..42] of string;
FileName, PromFileName:string;
ExcelIn: Variant;
Price:Extended;
FString:string;
PrintText:string;
FileName1, FileName2:string;
IsEmptyLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
Amount, I: Integer;
begin
PromFileName:=FileOpenDialog2.FileName;
try
try
  ExcelIn := GetActiveOleObject('Excel.Application');
       except on EOLESysError do
       ExcelIn := CreateOleObject('Excel.Application');
      end;
      ExcelIn.Visible := False;
      ExcelIn.WindowState := -4140;
      ExcelIn.DisplayAlerts := False;
      ExcelIn.WorkBooks.Open(PromFileName, 0 , true);
      ExcelIn.WorkSheets[1].Activate;
      //MemoTxt.Clear;
      //MemoTxt.Lines.Add(WritePromHeaders);
      LineNumber:=1;
      PB.Position:=0;
      PB.Min:=1;
      PB.Max:=300;
      PB.Step:=1;
      PB.StepIt;
      for I := 1 to 42 do
      begin
        CellRow:=caseNumber(i);
        CellNum:='1';
        CellText:=Trim(ExcelIn.Range[CellRow+CellNum]);
        if not isPromExpandHeaderCorrect(i, CellText) then
          begin
            MemoLog.Lines.Add('Неверный заголовок файла '+ExtractFileName(PromFileName)+', найдите файл export*.xls, вместо знака * будут цифры');
            MemoLog.Lines.Add('Зайдите на сайт prom.ua и выберите "Товары и услуги", затем кнопка "Экспорт" в правом верхнем углу');
            MemoLog.Lines.Add('Робот отправит письмо на Ваш ящик, и Вы сможете скачать файл export*.xlsx по ссылке, указанной в письме');
            MemoLog.Lines.Add('Ошибка найдена в колонке: "'+CellText+'" номер '+Cellrow);
            ShowMessage('Неверный заголовок файла '+ExtractFileName(PromFileName)+', найдите именно файл export*.xls, вместо знака * будут цифры.'+chr(10)+chr(13)
                          +'Зайдите на сайт prom.ua и выберите "Товары и услуги", затем нажмите кнопку "Экспорт" в правом верхнем углу'+chr(10)+chr(13)
                          +'Робот отправит письмо на Ваш ящик, и Вы сможете скачать файл export*.xlsx по ссылке, указанной в письме');
            exit;
          end;
      end;
      LineNumber:=2;
      isEmptyLine:=false;
      while not IsEmptyLine do
      begin
      PB.StepIt;
      for I := 1 to 42 do
        begin
        CellRow:=caseNumber(i);
        CellNum:=IntToStr(LineNumber);
        CellText:=trim(ExcelIn.Range[CellRow+CellNum]);
        PromExpandText[i]:=CellText;
        if LineNumber>50000 then IsEmptyLine:=true;
        end;
      if  (length(PromExpandText[1])=0)and(length(PromExpandText[2])=0)
           and (length(PromExpandText[3])=0)and(length(PromExpandText[4])=0)
      then
        begin
        //MemoLog.Lines.Add('Найдена пустая строка');
        IsEmptyLine:=true;
        Continue;
        end;
      if not IsEmptyLine then
        begin
        SavePromToSQLite(PromExpandText);
        end;
      inc(LineNumber);
      end;
    finally
      ExcelIn.ActiveWorkbook.Close;
      ExcelIn.Application.Quit;
    end;
end;

procedure TFormMain.LoadRemontkaToSQLite;
var
RemontkaText: array[1..23] of string;
FileName:string;
ExcelIn: Variant;
Price:Extended;
FString:string;
PrintText:string;
FileName1, FileName2:string;
IsEmptyLine, IsExcludedLine:boolean;
CellText, CellNum, CellRow:string;
LineNumber:integer;
Amount, I: Integer;
begin
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
      //MemoTxt.Clear;
      //MemoTxt.Lines.Add(WritePromHeaders);
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
        CellText:=TrimSeparator(CellText);
        if not isRemontkaHeaderCorrect(i, CellText) then
          begin
            MemoLog.Lines.Add('Неверный заголовок файла, проведите выгрузку "Остатки на складе.xls" из remonline ещё раз     '
                                +CellRow+CellNum+'!'+'!'+CellText);
            ShowMessage('Неверный файл остатоков, он создан нажатием на кнопку "Создать отчёт"'+chr(10)+chr(13)
                          +'Зайдите на сайт remonline ещё раз и выгрузите файл остатков с помощью "бутерброда"'+chr(10)+chr(13)
                          +'Выберите вкладку "Склад", бутерброд(три полоски) находится возле Строки "Наличие"');
            exit;
          end;
      end;
      LineNumber:=2;
      isEmptyLine:=false;
      while not IsEmptyLine do
      begin
      isExcludedLine:=false;
      PB.StepIt;
      for I := 1 to 13 do
        begin
        CellRow:=caseNumber(i);
        CellNum:=IntToStr(LineNumber);
        CellText:=trim(ExcelIn.Range[CellRow+CellNum]);
        RemontkaText[i]:=TrimSeparator(CellText);
        //if (i=1) and (length(RemontkaText[i])>0) then RemontkaText[i]:=''''+RemontkaText[i];
        if LineNumber>50000 then IsEmptyLine:=true;
        //Выходим если 50(00) строк чтобы не было зацикливания
        end;
      if  (length(RemontkaText[1])=0)and(length(RemontkaText[2])=0)
           and (length(RemontkaText[3])=0)and(length(RemontkaText[4])=0)
      then
        begin
        //LogText(RemontkaText);
        //MemoLog.Lines.Add('Найдена пустая строка');
        IsEmptyLine:=true;
        Continue;
        end;
      Amount:=StrToIntDef(RemontkaText[5],-1);
      if (Amount = 0) then
        begin
        if not CheckBoxZeroOstatki.Checked then LogText(RemontkaText);
        if not CheckBoxZeroOstatki.Checked then MemoLog.Lines.Add('Товар исключается, нулевое количество. Код "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
        isExcludedLine:=true;
        end;
      if (Amount = -1) then
        begin
        LogText(RemontkaText);
        MemoLog.Lines.Add('Товар с кодом "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
        MemoLog.Lines.Add('Товар исключается, количество="'+RemontkaText[5]+'" не является числом. Сообщите разработчику.');
        IsExcludedLine:=true;
        end;
      Price:=StrToFloatDef(RemontkaText[11],-1);
      if (Price = 0) then
        begin
        if not CheckBoxZeroPrice.Checked then LogText(RemontkaText);
        if not CheckBoxZeroPrice.Checked then MemoLog.Lines.Add('Товар исключается, нулевая цена. Код "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
        IsExcludedLine:=true;
        end;
      if (Price = -1) then
        begin
        LogText(RemontkaText);
        MemoLog.Lines.Add('Товар с кодом "'+RemontkaText[1]+'", Название "'+RemontkaText[4]+'"');
        MemoLog.Lines.Add('Товар исключается, цена ="'+RemontkaText[11]+'" отображается неверно. Сообщите разработчику.');
        isExcludedLine:=true;
        end;
      if not IsEmptyLine and not IsExcludedLine then
        begin
        PrintText:=PrintPromText(RemontkaText);
        SaveRemontkaTextToSQLite(RemontkaText);
        if PrintText<>'' then MemoTxt.Lines.Add(PrintText);
        end;
      inc(LineNumber);
      end;
    finally
      ExcelIn.ActiveWorkbook.Close;
      ExcelIn.Application.Quit;
    end;
end;

function TFormMain.LogText(const PText: array of string): string;
var i:integer;
begin
Result:=PText[1];
for I := 2 to length(Ptext) do Result:=Result+'|'+PText[i];
MemoLog.Lines.Add(Result);
end;

function TFormMain.PlusQuotes(Str: string; isQuoted: boolean): string;
begin
if isQuoted then Result:='"'+str+'"' else Result:=Str;
end;

function TFormMain.PrintPromText(pPromText: array of string): string;
var i, RemNumber:integer;
Price:Extended;
Ostatki:integer;
Nalichie:string;
begin
Result:='';
if (Mapping[1].RemontkaNumber>=0) then Result:=PlusQuotes(pPromText[Mapping[1].RemontkaNumber],Mapping[1].Quoted);
for I := 2 to 23 do
  begin
    Result:=Result+FileSeparator;
    RemNumber:=Mapping[i].RemontkaNumber;
    case RemNumber of
    -999:;
    -5: Result:=Result+PlusQuotes('r',Mapping[i].Quoted);
    -7: Result:=Result+PlusQuotes('UAH',Mapping[i].Quoted);
    -8: Result:=Result+PlusQuotes('шт.',Mapping[i].Quoted);
    -9: Result:=Result+PlusQuotes('1',Mapping[i].Quoted);
    -11: Result:=Result+PlusQuotes('2',Mapping[i].Quoted);
    4: begin
        Ostatki:=StrToIntDef(pPromText[Mapping[i].RemontkaNumber],0);
        if (Ostatki>0)
          then Nalichie:=PlusQuotes('+',Mapping[i].Quoted)
          else Nalichie:=PlusQuotes('-',Mapping[i].Quoted);
        if I=13 then Result:=Result+Nalichie;
        if i=14 then Result:=Result+IntToStr(Ostatki);
        end;
    else Result:=Result+PlusQuotes(pPromText[Mapping[i].RemontkaNumber],Mapping[i].Quoted);
    if (i=6) then
      begin
      Price:=StrToFloatDef(pPromText[Mapping[i].RemontkaNumber],-1);
      if (Price = 0) then
        begin
        Result:='';
        MemoLog.Lines.Add('Товар с кодом "'+pPromText[1]+'", Название "'+pPromText[2]+'"');
        MemoLog.Lines.Add('Товар исключается, нулевая цена');
        exit;
        end;
      if (Price = -1) then
        begin
        Result:='';
        MemoLog.Lines.Add('Товар с кодом "'+pPromText[1]+'", Название "'+pPromText[2]+'"');
        MemoLog.Lines.Add('Товар исключается, неверно выгрузилась цена '+pPromText[6]+'.Сообщите разработчику.');
        exit;
        end;
      end;
    end;
  end;
 end;

function TFormMain.QuotesForSQL(const Str: string): string;
var
Flags: TReplaceFlags;
begin
Flags:= [rfReplaceAll, rfIgnoreCase];
Result:=StringReplace(Str,'"','""',Flags);
end;

procedure TFormMain.SavePromToSQLite(pPromArray:array of string);
var
strSQL: String;
S3DB:TSQLiteDatabase;
S3Tbl: TSQLIteTable;
Code:string;
begin
  try
  S3DB := TSQLiteDatabase.Create(DBName);
  S3DB.BeginTransaction;
  code:=pPromArray[0];
  strSQL := 'INSERT INTO Prom_items(Product_code, Position_Name, Keywords, Description, Product_type, '
  +' Price, Currency, Unit_of_measurement, Minimum_size_Order, Wholesale_price, '
  +' Min_Order_Opt, Image_Link, Availability, Amount, Group_number, '
  +' Group_name, Division_Address, Possibility_of_delivery, Delivery_period, Packing_Mode, '
  +' Unique_identificator, Product_id, Subdivision_id, Group_id, Manufacturer, '
  +' Producing_country, Discount, Species_Group_ID, Tags, Product_on_Site, '
  +' Name1_Characteristics, Measurement1_Characteristics, Value1_Characteristics, '
  +' Name2_Characteristics, Measurement2_Characteristics, Value2_Characteristics, '
  +' Name3_Characteristics, Measurement3_Characteristics, Value3_Characteristics, '
  +' Name4_Characteristics, Measurement4_Characteristics, Value4_Characteristics ) VALUES ("'
    +QuotesForSQL(pPromArray[0])+'" , "'
    +QuotesForSQL(pPromArray[1])+'" , "'
    +QuotesForSQL(pPromArray[2])+'" , "'
    +QuotesForSQL(pPromArray[3])+'" , "'
    +QuotesForSQL(pPromArray[4])+'" , "'
    +QuotesForSQL(pPromArray[5])+'" , "'
    +QuotesForSQL(pPromArray[6])+'" , "'
    +QuotesForSQL(pPromArray[7])+'" , "'
    +QuotesForSQL(pPromArray[8])+'" , "'
    +QuotesForSQL(pPromArray[9])+'" , "'
    +QuotesForSQL(pPromArray[10])+'" , "'
    +QuotesForSQL(pPromArray[11])+'" , "'
    +QuotesForSQL(pPromArray[12])+'" , "'
    +QuotesForSQL(pPromArray[13])+'" , "'
    +QuotesForSQL(pPromArray[14])+'" , "'
    +QuotesForSQL(pPromArray[15])+'" , "'
    +QuotesForSQL(pPromArray[16])+'" , "'
    +QuotesForSQL(pPromArray[17])+'" , "'
    +QuotesForSQL(pPromArray[18])+'" , "'
    +QuotesForSQL(pPromArray[19])+'" , "'
    +QuotesForSQL(pPromArray[20])+'" , "'
    +QuotesForSQL(pPromArray[21])+'" , "'
    +QuotesForSQL(pPromArray[22])+'" , "'
    +QuotesForSQL(pPromArray[23])+'" , "'
    +QuotesForSQL(pPromArray[24])+'" , "'
    +QuotesForSQL(pPromArray[25])+'" , "'
    +QuotesForSQL(pPromArray[26])+'" , "'
    +QuotesForSQL(pPromArray[27])+'" , "'
    +QuotesForSQL(pPromArray[28])+'" , "'
    +QuotesForSQL(pPromArray[29])+'" , "'
    +QuotesForSQL(pPromArray[30])+'" , "'
    +QuotesForSQL(pPromArray[31])+'" , "'
    +QuotesForSQL(pPromArray[32])+'" , "'
    +QuotesForSQL(pPromArray[33])+'" , "'
    +QuotesForSQL(pPromArray[34])+'" , "'
    +QuotesForSQL(pPromArray[35])+'" , "'
    +QuotesForSQL(pPromArray[36])+'" , "'
    +QuotesForSQL(pPromArray[37])+'" , "'
    +QuotesForSQL(pPromArray[38])+'" , "'
    +QuotesForSQL(pPromArray[39])+'" , "'
    +QuotesForSQL(pPromArray[40])+'" , "'
    +QuotesForSQL(pPromArray[41])
    +'" );';
  //MemoLog.Lines.Add(strSQL);
  S3DB.ExecSQL(strSQL);
  S3DB.Commit;
  finally
  S3DB.Free;
  end;
end;

procedure TFormMain.SaveRemontkaTextToSQLite(pRemArray: array of string);
var
strSQL: String;
S3DB:TSQLiteDatabase;
S3Tbl: TSQLIteTable;
Code:string;
Flags: TReplaceFlags;
begin
Flags:= [rfReplaceAll, rfIgnoreCase];
  try
  S3DB := TSQLiteDatabase.Create(DBName);
  S3DB.BeginTransaction;
  code:=pRemArray[0];
  //if Pos('''',Code)>0 then Code:=StringReplace(Code,'''','',Flags);
  strSQL := 'INSERT INTO Remontka_items(Code, Artikul, Barcode, Name, Amount, Category, Warranty, WarrantyPeriod, PurchasePrice, ZeroPrice, InternetPrice, RepairPrice, RetailPrice, RepairPrice) VALUES ("'
    +QuotesForSQL(pRemArray[0])+'" , "'
    +QuotesForSQL(pRemArray[1])+'" , "'
    +QuotesForSQL(pRemArray[2])+'" , "'
    +QuotesForSQL(pRemArray[3])+'" , "'
    +QuotesForSQL(pRemArray[4])+'" , "'
    +QuotesForSQL(pRemArray[5])+'" , "'
    +QuotesForSQL(pRemArray[6])+'" , "'
    +QuotesForSQL(pRemArray[7])+'" , "'
    +QuotesForSQL(pRemArray[8])+'" , "'
    +QuotesForSQL(pRemArray[9])+'" , "'
    +QuotesForSQL(pRemArray[10])+'" , "'
    +QuotesForSQL(pRemArray[11])+'" , "'
    +QuotesForSQL(pRemArray[12])+'" , "'
    +QuotesForSQL(pRemArray[13])
    +'" );';
  //MemoLog.Lines.Add(strSQL);
  S3DB.ExecSQL(strSQL);
  S3DB.Commit;
  finally
  S3DB.Free;
  end;
end;

procedure TFormMain.EmptySQLite;
var
strSQL: String;
S3DB:TSQLiteDatabase;
S3Tbl: TSQLIteTable;
begin
  S3DB := TSQLiteDatabase.Create(DBName);
  try
  S3DB.BeginTransaction;
  strSQL := 'DELETE FROM Remontka_items;';
  S3DB.ExecSQL(strSQL);
  S3DB.Commit;
  S3DB.BeginTransaction;
  strSQL := 'DELETE FROM Prom_items;';
  S3DB.ExecSQL(strSQL);
  S3DB.Commit;
  finally
  S3DB.Free;
  end;
end;

function TFormMain.TrimSeparator(const Str: string): string;
var where:integer;
Local:string;
begin
Local:=Str;
if Pos(FileSeparator, Str)=0 then Result:=Local
else
  while Pos(FileSeparator, Local)>0 do
  begin
  where:=Pos(FileSeparator, Local);
  Local:=Copy(Local, 1, where-1) + Copy(Local, where+1, length(Local));
  end;
  Result:=Local;
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

function TFormMain.WritePromExpandHeaders: string;
var i:integer;
begin
Result:=PromExpandHeader[1];
for I := 2 to length(PromExpandHeader) do
  begin
    Result:=Result+FileSeparator+PromExpandHeader[i];
  end;
//Исправить позже. Неверно показывается последняя колонка для XLS
Result:=Result+FileSeparator;
end;

end.

