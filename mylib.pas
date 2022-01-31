{*************************************************************}
{ Biblioteca de Funções para Delphi                           }
{ Márcio Henrique da Silva                                    }
{ Adaptada para Delphi 2009/2010/XE por Alexssandro Marcelino }
{*************************************************************}

unit mylib;

interface

uses
  Windows, Dialogs, Messages, SysUtils, DateUtils, Classes, Controls, StdCtrls, IdUri,
  FileCtrl,Graphics, shellapi, Printers, Winsock, Registry, IniFiles, Forms, ClipBrd,
  System.Variants, FireDAC.Stan.Intf, FireDAC.Stan.Option, JvCipher,  uhardwareid,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait,
  FireDAC.Comp.Client, JVDBCombobox, JvCaptionPanel, DataSnap.DBClient,
  Grids, DBGrids, db, ComObj, MDBGrid, ComCtrls, SqlExpr, SHDocVw, mshtml, ActiveX,
  RegularExpressions, ADODB, CheckLst, Vcl.ExtCtrls, IdCoder, IdCoder3to4, IdCoderMIME,
  WinApi.TLHelp32, PsAPI;

{$Region 'Declarações'}
type
  TEspessuraRosca = (espFina, espMedia, espGrossa);

   function InverteStr (wStr1: String): String;
   function DecToBin(Valor: Integer) : String;
   function BinToDec(Valor: String): Cardinal;
   function TamArq(const FileName: String): LongInt;
   function GetFileDate(TheFileName: string): string;
   function FileDate(Arquivo: String): String;
   function FillDir(Const AMask: string): TStringList;
   function RecycleBin(sFileName : string ) : boolean;
   function NumLinhasArq(Arqtexto:String): integer;
   function FileCopy(source,dest: String): Boolean;
   function ExtractName(const Filename: String): String;
   function FileTypeName(const aFile: String): String;
   procedure CopyFile( Const sourcefilename, targetfilename: String );
   procedure ZapFiles(vMasc:String);
   function PrintImage(Origem: String):Boolean;
   function Stod(Texto : String) : Boolean;
   function Stodt(Texto : String) : Boolean;
   function StringToFloat(Texto : String) : Boolean;
   function StringToInt(Texto : String) : Boolean;
   function GetIP:string;
   function GetHardId:string;
   function GetHardIdCompleto:Ansistring;
   function CalculateCRC32(AStream : TStream) : Cardinal;
   function NomeComputador : String;
   function UserName : String;
   function DV_BASE10G(mDIG : String) : String;
   function DV_BASE10V(mDIG : String) : Boolean;
   function Crypt(Action, Src, Key : String) : String;
   procedure Del_Row_str(StrGrid : TStringGrid; Linha : Integer);
   procedure Ins_Row_str(StrGrid : TStringGrid; Linha : Integer);
   function IntDeData(data : TDatetime) : Integer;
   function Padr(texto : String;Tam : Integer):String;
   function Padl(texto : String;Tam : Integer):String;
   function TipoDeCampo(Campo : TFieldType) : string;
   function TestaCpf( xCPF:String ):Boolean;
   function TestaCgc(xCGC: String):Boolean;
   function AjustaCpf(xCPF: String) : String;
   function SoNumero(xTexto: String) : String;
   function FormataDinheiro(Texto: String) : String;
   function Centraliza(texto : String; Tam : Integer) : String;
   function Replicate(texto : String; Quant : Integer) : String;
   function TiraAcento(Texto : String) : String;
   function Idade(dtini,dtFim : TDateTime) : String;
   function UltimoDiaDoMes(dt : TDateTime) : Word;
   function StrZero(Zeros:string;Quant:integer):String;
   function CripSenha(Texto : String) : String; overload;
   function CripSenha(Texto, Chave : String) : String; overload;
   function DecripSenha(Texto : String) : String; overload;
   function DecripSenha(Texto, Chave : String) : String; overload;
   procedure CorDoGrid(const Grid: TdbGrid; Rect: TRect; Field: TField; State: TGridDrawState; Cor1 : TColor = $00DEECD7; Cor2 : TColor = $00F3FEFD);
   procedure CorDoGridMITEC(const Grid: TMdbGrid; Rect: TRect; Field: TField; State: TGridDrawState; Cor1 : TColor = $00DEECD7; Cor2 : TColor = $00F3FEFD; CorSelec: TColor = clNavy; CorSelecFnt: TColor = clWhite);
   procedure CorDoGridSel(const Grid: TdbGrid; Rect: TRect; Field: TField; State: TGridDrawState);
   procedure CorDoTitulo(Grid: TdbGrid;Campo: String);
   function BuscaEntre2(Msg, txtAntes, txtDepois : String) : String;
   function PegaNumCombo(Texto :String) : Integer;
   procedure AddLinhas(Var Onde : TStringList; Qtd:Integer);
   function FormataCgcCpf(Texto : String;Valida : Boolean = False ; Alerta : Boolean = False) : String;
   function FloatPonto(Texto : String) : String;
   function FloatVirgula(Texto : String) : String;
   function PosR(OQue: String; Onde: String) : Integer;
   function DivideLin(Texto : String;lin : Integer;larg : Integer = 50) : String;
   function Se(Expressao : Boolean; Valor1, Valor2 : Variant) :Variant;
   function CheckCC(c: string): Integer;
   function SoLetra(Texto: String): String;
   function Sem_Espacos(Texto : String) : String;
   function ValidaSequencia(Texto : String;Valor : Boolean = False;Ordena : Boolean = False) : String;
   function OrdenaString(Lista : TStringList) : TStringList;
   function PegaSeq(Texto : String; posicao : Integer;sep : Char = #44) : String;
   function ContaSeq(Texto : String; Sep: Char = #44) : Integer;
   function AchaPosSeq(Texto,Oque : String;sep : Char = #44) : Integer;
   function ApagaSeq(Texto : String;indice : Integer;sep : Char = #44) : String;
   function Ajusta(Texto,posicoes : String) : String;
   function ValidaEmail(Texto : String) : String;
   function EmailValido(Texto : string): Boolean;
   function FormataPreco(precobase,valentrada : Currency; nparcelas : word;Indice : Real;ComEntrada : Boolean ;var entrada : Currency;Var parcela : Currency) : String;
   function ArredondaPreco(Valor : Currency; Inteiro : Boolean) : Currency;
   function GeraArqSeq(Pasta,IniArq,Extensao : String; qtd : Integer) : String;
   function PrimeirasMaiusculas(Texto: string): string;
   function SeqNumLetra(texto : String) : String;
   function DizMes(NumMes : Word) : String;
   function PontoDir(Texto : String;Tam : Smallint) : String;
   function Passwd(Texto : String) : String;
   function Ajustar(S: String; T: Integer; D: String): String;
   function SoNumLeft(Valor :Extended;Tam,Dec : Integer) : String;
   function Preenche(S: String; T: Integer; D: String; Chr: String): String;
   function FormataCep (sValue:String): String;
   function StrTempo2Min(S: String): Integer;
   function ValiData(Data: String):Boolean;
   function ValiHora(Hora: String):Boolean;
   function ValiInteiro(Valor : String): Boolean;
   function ValiFlutuante(Valor : Variant): Boolean;
   function TraduzMes(texto :String) : String;
   function ExecutaPrograma(Programa: String): String;
   function UltDataDoMes(Data: TDateTime): TDate;
   function PrimeiroDoMes(Data : TDateTime) : TDate;
   function DatacomBarras(Texto: String) : String;
   function DatacomBarrasBrt(Texto: String) : String;
   function Arredondavisa(Valor : Real) : Real;
   function Extenso(Valor: Real; Reais: Boolean; Masculino : Boolean): String;
   function iif(Condicao: Boolean; rVerdade: Variant; rFalso: Variant): Variant;
   function ParamTimestamp(Tempo : Tdatetime) : String;
   function ParamDate(Data : TDateTime) : String;
   procedure GravaIni(usuario,chave,valor : String);
   function LeIni(usuario,chave:String) : String;
   function DivideParcelas(Valor : Currency; QtdParc, NumParc : Extended) : Currency;
   procedure SendToOpenOffice(aDataSet: TDataSet);
   function e64Bits : Boolean;
   procedure SomaMtecGrid(var Grid : TMdbGrid; textoCampos : String; Tipos : String = '0');
   procedure ExportarDadosParaExcel(Qry: TDataSet; RealComoInteiro: Boolean = True);
   procedure PreencheComboBoxEx(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TSQLConnection; PrimeiroItem : String);
   procedure PreencheComboBoxExFD(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TFDConnection; PrimeiroItem : String);
   procedure PreencheComboBoxExADO(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TADOConnection; PrimeiroItem : String);
   procedure PreencheJvDbComboBox(Combo : TJvDbComboBox; FieldItems, FieldValues, sSQL : String; Conexao : TSQLConnection; PrimeiroItem : String = '');
   procedure PreencheCheckListbox(CheckList: TCheckListBox; Resultado,Comando: string; Conexao: TSQLConnection; Checados:Boolean = False);
   function GetAveCharSize(Canvas: TCanvas): TPoint;
   function InputQueryPT(const ACaption, APrompt: string; var Value: string; cap1 : string = '&OK'; cap2 : string = '&Cancelar'): Boolean;
   function Win64 : Boolean;
   function CustomStrToDate(ADate: String): TDate;
   function VersaoWin: string;
   function VersaoArquivo(const NomeArq: string) : String; overload;
   procedure VersaoArquivo(const NomeArq: string; var iVersao: array of integer); overload;
   procedure VersaoArquivo(const NomeArq: string; var sVersao: String; bBuild : Boolean); overload;
//   function VersaoArquivoString(const NomeArq: string): string;
   procedure DeletaDiretorio(const Dir: string);
   procedure DeletaDiretorioRecursivo(const Diretorio : string);
   function TruncTo(Valor: Double; CasasDecimais: Integer) : Double;
   function ArredondarParaCima(Valor : Real; CasasDec : Integer): Real;
   function MinParaHora(Minuto: integer): string;
   function RemoveZeros(S: string): string;
   function FormataTipi(sValue:String): String;
   function ConvertBitmapToGrayscale(const Bitmap: TBitmap): TBitmap;
   procedure ClonarComboBoxEx(Origem, Destino : TComboBoxEx);
   function DialogPersonalizado(Msg: string; AType: TMsgDlgType; AButtons: TMsgDlgButtons;
      IndiceHelp: LongInt; DefButton: TMOdalResult = mrNone;
      sSim : string = '&Sim'; sNao: string = '&Não'; sCancelar:string = '&Canclear';
      sAbortar: string = '&Abortar'; sRepetir: string = '&Repetir'; sIgnorar: string = '&Ignorar';
      sTodos:string = '&Todos'; sAjuda: string = 'A&juda'): Word;
   procedure AjustaVisualDBGrid(DataSource : TDataSource; Propriedades : String);
   function Alinhamento(Sigla : String) : TAlignment;
   function FormatMinToHour(Min: LongInt): String;
   function FormatDecHourToHour(Hours : Extended) : String;
   function DistinctSeq(Lista: String; Sep: Char = #44): String;
   function IfThenString(AValue: Boolean; const ATrue: string; const AFalse: string = ''): string; overload; inline;
   function FormataDataRelatorio(dtIni, dtFin : TDate): string;
   function EnderecoMAC : string;
   procedure DesabilitaDeleteGrid(Sender: TObject; var Key: Word; Shift: TShiftState);
   procedure RetornaRosca(Percentual: Double; CorPerc, CorResto, CorFundo: TColor; Tamanho: Smallint; Destino: TImage; Espessura: TEspessuraRosca; Negrito : Boolean = False; CorNegat: TColor = clMaroon; ExibePerc: Boolean = True);
   procedure QualityResizeBitmap(bmpOrig, bmpDest: TBitmap; newWidth, newHeight: Integer);
   function RecortarImagem(Imagem : TImage; NewWidth, NewHeigth : Integer) : TImage;
   function base64Encode(Texto : AnsiString):AnsiString;
   function base64Decode(Texto : AnsiString):AnsiString;
   function ConverteEncodingXML(var sPathArquivoXML : String; Salvar : Boolean; PaiTemp : TWinControl) : Boolean;
   function Hex2Dec(texto : string) : string;
   function GerarStringRandom(Size : Integer; Tipo : Integer = 1) : String;
   function PosStringInArray(Texto : String; Vetor : Array of String) : Integer;
   procedure NaoEscondeJvCaptionPanel(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
   procedure CentralizarJvCaptionPanel(Panel : TJvCaptionPanel; Referencia : TObject; mkpH : Integer = 4; mkpW : Integer = 2);
   procedure EscondeSheets(pcGeral : TPageControl);
   procedure WBLoadHTML(WebBrowser: TWebBrowser; HTMLCode: string);
   procedure RetornaDifHorasMaior24(DataHoraA, DataHoraB : TDateTime; var Horas: integer; var Minutos: integer; var Segundos: integer; var Milissegundos: integer);
   procedure ExportaCsv(cdsExpor: TClientDataSet; sSeparador: String = ';'; bAutoExecutar: Boolean = False; sNomeArq: string = '');
   function FileTimeToDTime(FTime: TFileTime): TDateTime;
   function RetornaDataArquivo(NomeArquivo, TipoData : String) : TDateTime;
   function VerificaEXE(NomeEXE: String) : Boolean;
   procedure LimparMemoriaResidual;
   procedure PiscaTela(iHandle : Cardinal; iQuantidade : Cardinal = 10; iIntervalo : Cardinal = 500);

   //procedure PreencheComboBoxExFD(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TSQLConnection; PrimeiroItem : String);

type
  TTipoBallon = (tbNenhum , tbInfo , tbWarning , tbError);

type
   TEditBalloonTip = packed record
   cbStruct: DWORD ;
   pszTitle: LPCWSTR ;
   pszText: LPCWSTR;
   ttiIcon: Integer;
  const
    ECM_FIRST = $1500;
    EM_SHOWBALLOONTIP = (ECM_FIRST + 3);
    EM_HIDEBALLOONTIP = (ECM_FIRST + 4);
end;
type TMeuBallonHint = class
  public
   constructor ShowBallon(Window : HWnd;Texto, Titulo : PWideChar; Tipo : TTipoBallon);
   destructor  HideBallon(Window : HWnd);
end;

{$endregion}


implementation

uses StrUtils, Math;

const
  CRC32Table:  ARRAY[0..255] OF DWORD =
   ($00000000, $77073096, $EE0E612C, $990951BA,
    $076DC419, $706AF48F, $E963A535, $9E6495A3,
    $0EDB8832, $79DCB8A4, $E0D5E91E, $97D2D988,
    $09B64C2B, $7EB17CBD, $E7B82D07, $90BF1D91,
    $1DB71064, $6AB020F2, $F3B97148, $84BE41DE,
    $1ADAD47D, $6DDDE4EB, $F4D4B551, $83D385C7,
    $136C9856, $646BA8C0, $FD62F97A, $8A65C9EC,
    $14015C4F, $63066CD9, $FA0F3D63, $8D080DF5,
    $3B6E20C8, $4C69105E, $D56041E4, $A2677172,
    $3C03E4D1, $4B04D447, $D20D85FD, $A50AB56B,
    $35B5A8FA, $42B2986C, $DBBBC9D6, $ACBCF940,
    $32D86CE3, $45DF5C75, $DCD60DCF, $ABD13D59,
    $26D930AC, $51DE003A, $C8D75180, $BFD06116,
    $21B4F4B5, $56B3C423, $CFBA9599, $B8BDA50F,
    $2802B89E, $5F058808, $C60CD9B2, $B10BE924,
    $2F6F7C87, $58684C11, $C1611DAB, $B6662D3D,
    $76DC4190, $01DB7106, $98D220BC, $EFD5102A,
    $71B18589, $06B6B51F, $9FBFE4A5, $E8B8D433,
    $7807C9A2, $0F00F934, $9609A88E, $E10E9818,
    $7F6A0DBB, $086D3D2D, $91646C97, $E6635C01,
    $6B6B51F4, $1C6C6162, $856530D8, $F262004E,
    $6C0695ED, $1B01A57B, $8208F4C1, $F50FC457,
    $65B0D9C6, $12B7E950, $8BBEB8EA, $FCB9887C,
    $62DD1DDF, $15DA2D49, $8CD37CF3, $FBD44C65,
    $4DB26158, $3AB551CE, $A3BC0074, $D4BB30E2,
    $4ADFA541, $3DD895D7, $A4D1C46D, $D3D6F4FB,
    $4369E96A, $346ED9FC, $AD678846, $DA60B8D0,
    $44042D73, $33031DE5, $AA0A4C5F, $DD0D7CC9,
    $5005713C, $270241AA, $BE0B1010, $C90C2086,
    $5768B525, $206F85B3, $B966D409, $CE61E49F,
    $5EDEF90E, $29D9C998, $B0D09822, $C7D7A8B4,
    $59B33D17, $2EB40D81, $B7BD5C3B, $C0BA6CAD,
    $EDB88320, $9ABFB3B6, $03B6E20C, $74B1D29A,
    $EAD54739, $9DD277AF, $04DB2615, $73DC1683,
    $E3630B12, $94643B84, $0D6D6A3E, $7A6A5AA8,
    $E40ECF0B, $9309FF9D, $0A00AE27, $7D079EB1,
    $F00F9344, $8708A3D2, $1E01F268, $6906C2FE,
    $F762575D, $806567CB, $196C3671, $6E6B06E7,
    $FED41B76, $89D32BE0, $10DA7A5A, $67DD4ACC,
    $F9B9DF6F, $8EBEEFF9, $17B7BE43, $60B08ED5,
    $D6D6A3E8, $A1D1937E, $38D8C2C4, $4FDFF252,
    $D1BB67F1, $A6BC5767, $3FB506DD, $48B2364B,
    $D80D2BDA, $AF0A1B4C, $36034AF6, $41047A60,
    $DF60EFC3, $A867DF55, $316E8EEF, $4669BE79,
    $CB61B38C, $BC66831A, $256FD2A0, $5268E236,
    $CC0C7795, $BB0B4703, $220216B9, $5505262F,
    $C5BA3BBE, $B2BD0B28, $2BB45A92, $5CB36A04,
    $C2D7FFA7, $B5D0CF31, $2CD99E8B, $5BDEAE1D,
    $9B64C2B0, $EC63F226, $756AA39C, $026D930A,
    $9C0906A9, $EB0E363F, $72076785, $05005713,
    $95BF4A82, $E2B87A14, $7BB12BAE, $0CB61B38,
    $92D28E9B, $E5D5BE0D, $7CDCEFB7, $0BDBDF21,
    $86D3D2D4, $F1D4E242, $68DDB3F8, $1FDA836E,
    $81BE16CD, $F6B9265B, $6FB077E1, $18B74777,
    $88085AE6, $FF0F6A70, $66063BCA, $11010B5C,
    $8F659EFF, $F862AE69, $616BFFD3, $166CCF45,
    $A00AE278, $D70DD2EE, $4E048354, $3903B3C2,
    $A7672661, $D06016F7, $4969474D, $3E6E77DB,
    $AED16A4A, $D9D65ADC, $40DF0B66, $37D83BF0,
    $A9BCAE53, $DEBB9EC5, $47B2CF7F, $30B5FFE9,
    $BDBDF21C, $CABAC28A, $53B39330, $24B4A3A6,
    $BAD03605, $CDD70693, $54DE5729, $23D967BF,
    $B3667A2E, $C4614AB8, $5D681B02, $2A6F2B94,
    $B40BBE37, $C30C8EA1, $5A05DF1B, $2D02EF8D);



procedure DesabilitaDeleteGrid(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if ((Shift = [ssCtrl]) and (Key = VK_DELETE)) then Abort;
   if ((Shift = [ssCtrl,ssShift]) and (Key = VK_DELETE)) then Abort;
end;


function DecToBin(Valor: Integer) : String;
var
  Binario: String;
begin
  while (valor >= 1) do begin
    Binario := IntToStr(valor mod 2) + Binario;
    Valor := (Valor div 2);
  end;
  Result := Binario;
end;

function BinToDec(Valor: String): Cardinal;
var
  Decimal : Real;
  x, y : Integer;
begin
  decimal := 0;
  y := 0;
  for x := Length(Valor) downTo 1 Do
  begin
    Decimal := Decimal + (StrToFloat(Valor[x])) * Exp(y * Ln(2));
    y := y + 1;
  end;
  Result := Round(Decimal);
end;

function InverteStr(wStr1: String): String;
var
   i: Integer;
begin
  Result := '';
  for i := Length( wStr1 ) downto 1 do Result := Result + Copy(wStr1,i,1 );
end;

function RemoveZeros(S: string): string;
var
   I, J : Integer;
begin
   I := Length(S);
   while (I > 0) and (S[I] <= ' ') do begin
      Dec(I);
   end;
   J := 1;
   while (J < I) and ((S[J] <= ' ') or (S[J] = '0')) do begin
      Inc(J);
   end;
   Result := Copy(S, J, (I-J)+1);
end;

function TamArq(const FileName: String): LongInt;
{Retorna o tamanho de um arquivo}
var
  SearchRec       : TSearchRec;
begin                                       { !Win32! -> GetFileSize }
  if FindFirst(FileName,faAnyFile,SearchRec)=0
    then Result:=SearchRec.Size
    else Result:=0;
  FindClose(SearchRec);
end;


function GetFileDate(TheFileName: string): string;
var
   FHandle: integer;
begin
   FHandle := FileOpen(TheFileName, 0);
   result  := DateToStr((FileDateToDateTime(FileGetDate(FHandle))));
   FileClose(FHandle);
end;


function FileDate(Arquivo: String): String;
{Retorna a data e a hora de um arquivo}
var
   FHandle: integer;
begin
if not fileexists(Arquivo) then
   begin
      Result := 'Nome de Arquivo Inválido';
   end
else
   begin
      FHandle := FileOpen(Arquivo, 0);
      try
        Result := DateTimeToStr(FileDateToDateTime(FileGetDate(FHandle)));
      finally
        FileClose(FHandle);
      end;
   end;
end;


Procedure ZapFiles(vMasc:String);
{Apaga arquivos usando mascaras tipo: *.zip, *.* }
Var Dir : TsearchRec;
    Erro: Integer;
Begin
   Erro := FindFirst(vMasc,faArchive,Dir);
   While Erro = 0 do Begin
      DeleteFile( ExtractFilePAth(vMasc)+Dir.Name );
      Erro := FindNext(Dir);
   end;
   FindClose(Dir);
end;


function FillDir(Const AMask: string): TStringList;
{Retorna uma TStringlist de todos os arquivos localizados
 no path corrente , Esta função trabalha com mascaras}
var
  SearchRec  : TSearchRec;
  intControl : integer;
begin
  Result := TStringList.create;
  intControl := FindFirst( AMask, faAnyFile, SearchRec );
  if intControl = 0 then
     begin
     while (intControl = 0) do
           begin
           Result.Add( SearchRec.Name );
           intControl := FindNext( SearchRec );
           end;
     FindClose( SearchRec );
     end;
end;


Function RecycleBin(sFileName : string ) : boolean;
// Envia um arquivo para a lixeira ( requer a unit Shellapi.pas)
var
fos : TSHFileOpStruct;
Begin
FillChar( fos, SizeOf( fos ), 0 );
With fos do
begin
wFunc := FO_DELETE;
pFrom := PChar( sFileName );
fFlags := FOF_ALLOWUNDO
or FOF_NOCONFIRMATION
or FOF_SILENT;
end;
Result := (0 = ShFileOperation(fos));
end;

function NumLinhasArq(Arqtexto:String): integer;
// Retorna o número de linhas que um arquivo possui
Var
   f: Textfile;
   cont:integer;
Begin
cont := 0;
AssignFile(f,Arqtexto);
Reset(f);
While not eof(f) Do
      begin
      ReadLn(f);
      Cont := Cont + 1;
      end;
Closefile(f);
result := cont;
end;


function FileCopy(source,dest: String): Boolean;
{copia um arquivo de um lugar para outro. Retornando falso em caso de erro}
var
fSrc,fDst,len: Integer;
size: Longint;
buffer: packed array [0..2047] of Byte;
begin
if source <> dest then
   begin
   fSrc := FileOpen(source,fmOpenRead);
   if fSrc >= 0 then
      begin
      size := FileSeek(fSrc,0,2);
      FileSeek(fSrc,0,0);
      fDst := FileCreate(dest);
      if fDst >= 0 then
         begin
         while size > 0 do
               begin
               len := FileRead(fSrc,buffer,sizeof(buffer));
               FileWrite(fDst,buffer,len);
               size := size - len;
               end;
         FileSetDate(fDst,FileGetDate(fSrc));
         FileClose(fDst);
         FileSetAttr(dest,FileGetAttr(source));
         Result := True;
         end
      else
         begin
         Result := False;
         end;
      FileClose(fSrc);
      end;
   end;
end;

Procedure CopyFile( Const sourcefilename, targetfilename: String );
{Copia um arquivo de um lugar para outro}
Var
  S, T: TFileStream;
Begin
  S := TFileStream.Create( sourcefilename, fmOpenRead );
  try
    T := TFileStream.Create( targetfilename, fmOpenWrite or fmCreate );
    try
      T.CopyFrom(S, S.Size ) ;
    finally
      T.Free;
    end;
  finally
    S.Free;
  end;
end;


function ExtractName(const Filename: String): String;
{Retorna o nome do Arquivo sem extensão}
var
aExt : String;
aPos : Integer;
begin
aExt := ExtractFileExt(Filename);
Result := ExtractFileName(Filename);
if aExt <> '' then
   begin
   aPos := Pos(aExt,Result);
   if aPos > 0 then
      begin
      Delete(Result,aPos,Length(aExt));
      end;
   end;
end;


function  FileTypeName(const aFile: String): String;
{Retorna descrição do tipo do arquivo. Requer a unit ShellApi}
var
  aInfo: TSHFileInfo;
begin
  if SHGetFileInfo(PChar(aFile),0,aInfo,Sizeof(aInfo),SHGFI_TYPENAME)<>0 then
     Result := StrPas(aInfo.szTypeName)
  else begin
     Result := ExtractFileExt(aFile);
     Delete(Result,1,1);
     Result := Result +' File';
  end;
end;


function PrintImage(Origem: String):Boolean;
// imprime um bitmap selecionado retornando falso em caso negativo
// requer as units Graphics e printers declaradas na clausula Uses
var
Imagem: TBitmap;
begin
if fileExists(Origem) then
   begin
   Imagem := TBitmap.Create;
   Imagem.LoadFromFile(Origem);
   with Printer do
        begin
        BeginDoc;
        Canvas.Draw((PageWidth - Imagem.Width) div 2,(PageHeight - Imagem.Height) div 2,Imagem);
        endDoc;
        end;
   Imagem.Free;
   Result := True;
   end
else
   begin
   Result := False;
   end;
end;

function stod(Texto : String) : Boolean;
// Retorna True se uma strin pode ser convertida em data
Begin
   try
      strtodate(Texto);
      result := True;
   except
      result := False;
   end;
end;

function stodt(Texto : String) : Boolean;
begin
   try
      strtodatetime(Texto);
      Result := True;
   Except
      Result := False;
   end;
end;

function stringtofloat(Texto : String) : Boolean;
var
   Aux  : Boolean;
   Aux2 : Extended;
Begin
   Aux := True;
   Try
      Aux2 := strtofloat(Texto);
   except
      Aux := False;
   end;
   Result := Aux;
end;

function stringtoint(Texto : String) : Boolean;
var
   Aux  : Boolean;
   Aux2 : integer;
Begin
   Aux := True;
   Try
      Aux2 := strtoint(Texto);
   except
      Aux := False;
   end;
   Result := Aux;
end;


function GetIP:string;
//
// Retorna o IP de sua máquina no momento em que
// você está conectado
//
// Declare a Winsock na clausula uses da unit
//
var
   WSAData: TWSAData;
   HostEnt: PHostEnt;
   Name:string;
begin
   WSAStartup(2, WSAData);
   SetLength(Name, 255);
   Gethostname(PAnsiChar (Name), 255);
   SetLength(Name, StrLen(PChar(Name)));
   HostEnt := gethostbyname(PAnsiChar(Name));
   with HostEnt^  do
        begin
        Result := Format('%d.%d.%d.%d',[Byte(h_addr^[0]),Byte(h_addr^[1]),Byte(h_addr^[2]),Byte(h_addr^[3])]);
        end;
   WSACleanup;
   if FileExists('c:\mhs.cfg') then result := '192.168.0.37';
end;

function GetHardId:string;
// Retorna o Hardware ID de sua máquina no momento em que
// você está conectado
var
   HWID  : THardwareId;
   sHardId  : AnsiString;
   aStringStream : TStringStream;
begin
    try
       HWID:=THardwareId.Create(False);
       try
          HWID.GenerateHardwareId;
          sHardId := (Format('%s %s',['',HWID.HardwareIdHex]));
          aStringStream := TStringStream.Create(sHardId);
          try
             result := IntToHex(CalculateCRC32(aStringStream), 8);
          finally
             aStringStream.Free;
          end;
       finally
        HWID.Free;
       end;
    except
       on E:Exception do
       begin
          ShowMessage('Erro ao consultar Hardware Id.');
       end;
     end;
end;

function GetHardIdCompleto:Ansistring;
// Retorna o Hardware ID de sua máquina no momento em que
// você está conectado
var
   HWID  : THardwareId;
   sHardId  : AnsiString;
begin
    try
       HWID:=THardwareId.Create(False);
       try
          HWID.GenerateHardwareId;
          sHardId := HWID.HardwareIdHex;
          result := sHardId;
       finally
        HWID.Free;
       end;
    except
       on E:Exception do
       begin
          ShowMessage('Erro ao consultar Hardware Id.');
       end;
     end;
end;

function CalculateCRC32(AStream : TStream) : Cardinal;
var aMemStream : TMemoryStream;
    aValue : Byte;
begin
  aMemStream := TMemoryStream.Create;
  try
    Result := $FFFFFFFF;
    while AStream.Position < AStream.Size do begin
      // Read a chunk of data...
      aMemStream.Seek(0, soFromBeginning);
      if AStream.Size - AStream.Position >= 1024*1024
        then aMemStream.CopyFrom(AStream, 1024*1024)
        else begin
          aMemStream.Clear;
          aMemStream.CopyFrom(AStream, AStream.Size-AStream.Position);
        end;
      // Beräkna CRC:n för blocket...
      aMemStream.Seek(0, soFromBeginning);
      while aMemStream.Position < aMemStream.Size do begin
        aMemStream.ReadBuffer(aValue, 1);
        Result := (Result shr 8) xor CRC32Table[aValue xor (Result and $000000FF)];
      end;
    end;
    Result := not Result;
  finally
    aMemStream.Free;
  end;
end;

function NomeComputador : String;
// Retorna o nome do computador
var
  lpBuffer : PChar;
  nSize    : DWord;
const
  Buff_Size = MAX_COMPUTERNAME_LENGTH + 1;
begin
   nSize := Buff_Size;
   lpBuffer := StrAlloc(Buff_Size);
   GetComputerName(lpBuffer,nSize);
   Result := String(lpBuffer);
   StrDispose(lpBuffer);
end;

function UserName : String;
// Retorna o usuário que está logado na rede
// Esta função funciona tanto no Win9x quanto no NT
var
lpBuffer : Array[0..20] of Char;
nSize    : dWord;
Achou    : boolean;
erro     : dWord;
begin
nSize      := 120;
Achou      := GetUserName(lpBuffer,nSize);
if (Achou) then
   begin
   result   := lpBuffer;
   end
else
   begin
   Erro   :=GetLastError();
   result :=IntToStr(Erro);
   end;
end;


function DV_BASE10V(mDIG : String) : Boolean;
var
   i, mDV, mDVA, mMULT : Integer;
   mSOMA : String;
begin
   if Length(mDIG) < 2 Then begin
      Result := False;
      Exit;
   end;
   mDVA  := strtoint( copy(mDIG,length(mDIG),1));
   mDIG  := copy(mDIG,1,Length(mDIG) - 1);
   mMULT := 2;
   mSOMA := '';
   For i := Length(mDIG) downto 1 do begin
      mSOMA := Trim(inttoStr(strtoint(copy(mDIG, i, 1)) * mMULT)) + mSOMA;
      If mMULT = 1 Then mMULT := 2 else mMULT := 1;
   end;

   mMULT := 0;

   For i := 1 to Length(mSOMA) do mMULT := mMULT + strtoint(copy(mSOMA, i, 1));

   mDV := 10 - (mMULT Mod 10);

   If mDV = 10 Then mDV := 0;

   If mDV = mDVA Then Result := True else Result := false;

end;

function DV_BASE10G(mDIG : String) : String;
var
  i, mDV, mMULT : Integer;
  mSOMA : String;
begin
   if Length(mDIG) < 1 Then begin
      DV_BASE10G := '';
      Exit;
   end;
   mMULT := 2;
   mSOMA := '';
   For i := Length(mDIG) downto 1 do begin
      mSOMA := Trim(inttoStr(strtoint(copy(mDIG, i, 1)) * mMULT)) + mSOMA;
      If mMULT = 1 Then mMULT := 2 Else mMULT := 1;
   end;

   mMULT := 0;

   For i := 1 To Length(mSOMA) do mMULT := mMULT + strtoint(copy(mSOMA, i, 1));

   mDV := 10 - (mMULT Mod 10);

   If mDV = 10 Then mDV := 0;
   Result := Trim(inttoStr(mDV));

end;

function Crypt(Action, Src, Key : String) : String;
var
   KeyLen    : Integer;
   KeyPos    : Integer;
   OffSet    : Integer;
   Dest      : string;
   SrcPos    : Integer;
   SrcAsc    : Integer;
   TmpSrcAsc : Integer;
   Range     : Integer;
begin
     Dest   := '';
     KeyLen := Length(Key);
     KeyPos := 0;
     SrcPos := 0;
     SrcAsc := 0;
     Range  := 256;
     if Action = UpperCase('E') then
     begin
          Randomize;
          offset:=Random(Range);
          dest:=format('%1.2x',[offset]);
          for SrcPos := 1 to Length(Src) do
          begin
               SrcAsc:=(Ord(Src[SrcPos]) + offset) MOD 255;
               if KeyPos < KeyLen then KeyPos:= KeyPos + 1 else KeyPos:=1;
               SrcAsc:= SrcAsc xor Ord(Key[KeyPos]);
               dest:=dest + format('%1.2x',[SrcAsc]);
               offset:=SrcAsc;
          end;
     end;
     if Action = UpperCase('D') then
     begin
          offset:=StrToInt('$'+ copy(src,1,2));
          SrcPos:=3;
          repeat
                SrcAsc:=StrToInt('$'+ copy(src,SrcPos,2));
                if KeyPos < KeyLen Then KeyPos := KeyPos + 1 else KeyPos := 1;
                TmpSrcAsc := SrcAsc xor Ord(Key[KeyPos]);
                if TmpSrcAsc <= offset then
                     TmpSrcAsc := 255 + TmpSrcAsc - offset
                else
                     TmpSrcAsc := TmpSrcAsc - offset;
                dest := dest + chr(TmpSrcAsc);
                offset:=srcAsc;
                SrcPos:=SrcPos + 2;
          until SrcPos >= Length(Src);
     end;
     Crypt:=dest;
end;

procedure Del_Row_str(StrGrid : TStringGrid; Linha : Integer);
Var
   Resto,Cols : Integer;
Begin
   For Resto := Linha to StrGrid.RowCount -1 do begin
      for Cols := 0 to StrGrid.ColCount -1 do
          StrGrid.Cells[Cols,Resto] := StrGrid.Cells[Cols,Resto + 1]
   end;
   StrGrid.RowCount := StrGrid.RowCount -1;
end;

procedure Ins_Row_str(StrGrid : TStringGrid; Linha : Integer);
Var
   Resto,Cols : Integer;
Begin
   StrGrid.RowCount := StrGrid.RowCount +1;
   For Resto := StrGrid.RowCount -1 downto Linha do begin
      for Cols := 0 to StrGrid.ColCount -1 do
          StrGrid.Cells[Cols,Resto+1] := StrGrid.Cells[Cols,Resto]
   end;
   strGrid.Rows[linha].Clear;
end;

function intdedata(data : TDatetime) : Integer;
var
  Year, Month, Day: Word;
begin
  DecodeDate(Data, Year, Month, Day);
  Result := Day + Month * 31 + Year * 272;
end;

function padr(texto: String; Tam: Integer): String;
begin
   if length(texto) > Tam then Texto := copy (texto,1,tam);
   while length(texto) < tam do
      Texto := texto + ' ';
   Result := Texto;
end;

function padl(texto : String;Tam : Integer):String;
begin
   if length(texto) > Tam then Texto := copy (texto,1,tam);
   while length(texto) < tam do
      Texto := ' ' + texto;
   Result := Texto;
end;

function TipoDeCampo(Campo: TFieldType): string;
var
   Tipo : string;
begin
  case Campo of
    ftUnknown:     tipo := 'Desconhecido';
    ftString:      Tipo := 'String';
    ftSmallint:    Tipo := 'Inteiro 16-bit';
    ftInteger:     Tipo := 'Inteiro 32-bit';
    ftWord:	       Tipo := 'Interio 16-bit positivos';
    ftBoolean:     Tipo := 'Boleano';
    ftFloat:       Tipo := 'Flutuante';
    ftCurrency:    Tipo := 'Monetário';
    ftBCD:         Tipo := 'Binário decimal';
    ftDate:        Tipo := 'Data';
    ftTime:        Tipo := 'Hora';
    ftDateTime:    Tipo := 'Data/Hora';
    ftBytes:       Tipo := 'Binário fixo';
    ftVarBytes:    Tipo := 'Binário variável';
    ftAutoInc:     Tipo := 'Int. Auto Inc. 32-bit';
    ftBlob:        Tipo := 'Binário longo';
    ftMemo:        Tipo := 'Memorando';
    ftGraphic:     Tipo := 'Bitmap';
    ftFmtMemo:     Tipo := 'Texto formatado';
    ftParadoxOle:  Tipo := 'Paradox OLE';
    ftDBaseOle:    Tipo := 'dBASE OLE';
    ftTypedBinary: Tipo := 'Binário';
    ftCursor:      Tipo := 'Oracle cursor de saída';
    ftFixedChar:   Tipo := 'Caracter fixo';
    ftWideString:  Tipo := 'Wide String';
    ftLargeInt:    Tipo := 'Inteiro longo';
    ftADT:         Tipo := 'Abstrato';
    ftArray:       Tipo := 'Array';
    ftReference:   Tipo := 'REF';
    ftDataSet:     Tipo := 'DataSet';
    ftOraBlob:     Tipo := 'BLOB Oracle 8';
    ftOraClob:     Tipo := 'CLOB Oracle 8';
    ftVariant:     Tipo := 'Variante';
    ftInterface:   Tipo := 'Referência de interfaces';
    ftIDispatch:   Tipo := 'Ref. de interfaces despachada';
    ftGuid:        Tipo := 'GUID';
  end;
  Result := Tipo;
end;

function TestaCpf( xCPF:String ):Boolean;
{Testa se o CPF é válido ou não}
Var
d1,d4,xx,nCount,resto,digito1,digito2 : Integer;
Check : String;
Begin
xCPF := sonumero(xCPF);
if xCPF = '' then begin
   result := true;
   Exit;
end;
d1 := 0; d4 := 0; xx := 1;
for nCount := 1 to Length( xCPF )-2 do
    begin
    if Pos( Copy( xCPF, nCount, 1 ), '/-.' ) = 0 then
       begin
       d1 := d1 + ( 11 - xx ) * StrToInt( Copy( xCPF, nCount, 1 ) );
       d4 := d4 + ( 12 - xx ) * StrToInt( Copy( xCPF, nCount, 1 ) );
       xx := xx+1;
       end;
    end;
resto := (d1 mod 11);
if resto < 2 then
   begin
   digito1 := 0;
   end
else
   begin
   digito1 := 11 - resto;
   end;
d4 := d4 + 2 * digito1;
resto := (d4 mod 11);
if resto < 2 then
   begin
   digito2 := 0;
   end
else
   begin
   digito2 := 11 - resto;
   end;
Check := IntToStr(Digito1) + IntToStr(Digito2);
if Check <> copy(xCPF,succ(length(xCPF)-2),2) then
   begin
   Result := False;
   end
else
   begin
   Result := True;
   end;
end;

function TestaCgc(xCGC: String):Boolean;
{Testa se o CGC é válido ou não}
Var
d1,d4,xx,nCount,fator,resto,digito1,digito2 : Integer;
Check : String;
begin
xCGC := sonumero(xCGC);
d1 := 0;
d4 := 0;
xx := 1;
if xCGC = '' then begin
   result := true;
   Exit;
end;
for nCount := 1 to Length( xCGC )-2 do
    begin
    if Pos( Copy( xCGC, nCount, 1 ), '/-.' ) = 0 then
       begin
       if xx < 5 then
          begin
          fator := 6 - xx;
          end
       else
          begin
          fator := 14 - xx;
          end;
       d1 := d1 + StrToInt( Copy( xCGC, nCount, 1 ) ) * fator;
       if xx < 6 then
          begin
          fator := 7 - xx;
          end
       else
          begin
          fator := 15 - xx;
          end;
       d4 := d4 + StrToInt( Copy( xCGC, nCount, 1 ) ) * fator;
       xx := xx+1;
       end;
    end;
    resto := (d1 mod 11);
    if resto < 2 then
       begin
       digito1 := 0;
       end
   else
       begin
       digito1 := 11 - resto;
       end;
   d4 := d4 + 2 * digito1;
   resto := (d4 mod 11);
   if resto < 2 then
      begin
      digito2 := 0;
      end
   else
      begin
      digito2 := 11 - resto;
      end;
   Check := IntToStr(Digito1) + IntToStr(Digito2);
   if Check <> copy(xCGC,succ(length(xCGC)-2),2) then
      begin
      Result := False;
      end
   else
      begin
      Result := True;
      end;
end;

function AjustaCpf(xCPF: String) : String;
Var Aux : String;
begin
   Aux := sonumero(xCPF);
   Result := Aux;
   if length(Aux) = 11 then
      Result := copy(Aux,1,3) + '.' + copy(Aux,4,3) + '.' + copy(Aux,7,3) + '-' + copy(Aux,10,2);
end;

function Sonumero(xTexto: String) : String;
Var
   i : Integer;
begin
   Result := '';
   for i := 1 to length(xTexto) do if xTexto[i] in ['0'..'9'] then Result := Result + xTexto[i];
end;

function formatadinheiro(Texto: String) : String;
Var
  i : integer;
begin
   // Nova versão mais eficiente 07/01/2004
   result := '';
   for i := length(Texto) downto 1 do
      if Texto[i] in ['0'..'9'] then result := Texto[i] + Result
      else if (Texto[i] in [chr(44),chr(46)]) and (pos(FormatSettings.decimalseparator,Result)=0) then Result := FormatSettings.decimalseparator + Result;
   if result = '' then result := '0' + FormatSettings.decimalseparator + '00';
   if copy(Texto,1,1) = '-' then Result := '-' + Result;
   Result := formatfloat('0.00',strtofloat(result));
end;

function Centraliza(texto : String; Tam : Integer) : String;
Var
   qtd,i : Integer;
   Esq : String;
begin
   Esq := '';
   qtd := (Tam - Length(Texto)) Div 2;
   For i := 1 to qtd do Esq := Esq + ' ';
   Result := Esq + Texto;
end;

function replicate(texto : String; Quant : Integer) : String;
Var i : Integer;
Begin
   Result := '';
   for i := 1 to Quant do Result := Result + texto;
end;

function tiraacento(Texto : String) : String;
Var i,posicao : Integer;
   Acento,Normal: String;
Begin
   Acento := 'áéíóúãõçâêîôûÁÉÍÓÚÃÕÂÊÎÔÛÇ';
   Normal := 'aeiouaocaeiouAEIOUAOAEIOUC';
   Result := '';
   For i := 1 to length(Texto) do begin
      posicao := pos(Texto[i],Acento);
      if posicao <> 0 then Result := Result + Normal[Posicao]
      else Result := Result + Texto[i];
   end;
end;

function Idade(dtini,dtFim : TDateTime) : String;
Var
   ano,dia,mes : Integer;
Begin
   Result := '';
   Ano := yearof(dtFim) - yearof(dtIni);
   Mes := monthof(dtFim) - monthof(dtIni);
   if Mes < 0 then begin
      Mes := 12 - (monthof(dtini) - monthof(dtFim));
   end;
   Dia := dayof(dtFim) - dayof(dtini);
   if Dia < 0 then begin
      Dia := ultimodiadomes(dtIni) - dayof(dtini) + dayof(dtFim);
      if mes > 0 then Dec(Mes);
   end;
   if (((monthof(dtFim) = monthof(dtIni)) and (dayof(dtFim) < dayof(dtIni))) or (monthof(dtFim) < monthof(dtIni))) then dec(Ano);
   if Ano > 0 then begin if Ano = 1 then Result := '1 ano' else Result := inttostr(Ano) + ' anos' end;
   if (Ano > 0) and (Mes > 0) and (Dia > 0) then Result := Result + ', ';
   if ((Ano > 0) and (Mes = 0) and (Dia > 0)) or ((Ano > 0) and (Mes > 0) and (Dia = 0)) then Result := Result + ' e ';
   if Mes > 0 then begin if Mes = 1 then Result := Result + '1 mês' else Result := Result + inttostr(Mes) + ' meses' end;
   if (Mes > 0) and (dia > 0) then Result := Result + ' e ';
   if Dia > 0 then begin if Dia = 1 then Result := Result + '1 dia' else Result := Result + inttostr(Dia) + ' dias' end;
   Result := trim(Result);
end;

function ultimodiadomes(dt : TDateTime) : Word;
Var Aux : TDatetime;
begin
   if monthof(dt) = 12 then
      Aux := strtodatetime('01/01/' + inttostr(yearof(dt)+1))
   Else Aux := strtodatetime('01/' + inttostr(monthof(dt)+1) + '/' + inttostr(yearof(dt)));
   Result := dayof(Aux-1);
end;

function StrZero(Zeros:string;Quant:integer):String;
{Insere Zeros à frente de uma string}
var
  i : integer;
begin
  Zeros := trim(zeros);
  for I:=1 to Quant - length(Zeros) do Zeros:= '0' + Zeros;
  Result := Zeros;
end;

function CripSenha(Texto : String) : String;
Var
  i : integer;
  Temp : String;
Begin
   for i := length(Texto) Downto 1 do
      Temp := Temp + Texto[i];
   Texto := '';
   for i := 1 to length(Temp) Do
      Texto := Texto + chr(ord(Temp[i]) + i);
   Result := Texto;
end;

function DecripSenha(Texto : String) : String;
Var
  i : integer;
  Temp : String;
Begin
   Temp  := texto;
   Texto := '';
   for i := length(Temp) Downto 1 do
      Texto := Texto + chr(ord(Temp[i]) - i);
   temp := '';   
   for i := 1 to length(Texto) do
      Temp := Temp + Texto[i];
   Result := Temp;
end;

function CripSenha(Texto, Chave : String) : String; overload;
var cCipher : TJvCaesarCipher;
begin
   try
      cCipher  := TJvCaesarCipher.Create(nil);
      Result   := cCipher.EncodeString(Chave,Texto);
   finally
      FreeAndNil(cCipher);
   end;
end;

function DecripSenha(Texto, Chave : String) : String; overload;
var cCipher : TJvCaesarCipher;
begin
   try
      cCipher  := TJvCaesarCipher.Create(nil);
      Result   := cCipher.DecodeString(Chave,Texto);
   finally
      FreeAndNil(cCipher);
   end;
end;






procedure cordogrid(const Grid: TdbGrid; Rect: TRect; Field: TField; State: TGridDrawState; Cor1 : TColor = $00DEECD7; Cor2 : TColor = $00F3FEFD);
begin
  if not (gdSelected in State) then begin
     if (Field.DataSet.RecNo mod 2) = 0 then begin
        Grid.Canvas.Brush.Color := Cor1;
        Grid.Canvas.Font.Color  := clBlack;
     end else begin
        Grid.Canvas.Brush.Color := Cor2;
        Grid.Canvas.Font.Color:= clBlack;
     end;
  end else begin
     Grid.Canvas.Brush.Color := clNavy;
     Grid.Canvas.Font.Color:= clWhite;
  end;
  Grid.Canvas.FillRect(Rect);
  Grid.DefaultDrawDataCell(Rect,Field,State);
end;

procedure cordogridMITEC(const Grid: TMdbGrid; Rect: TRect; Field: TField; State: TGridDrawState;
   Cor1 : TColor = $00DEECD7; Cor2 : TColor = $00F3FEFD; CorSelec: TColor = clNavy; CorSelecFnt: TColor = clWhite);
begin
   if not (gdSelected in State) then begin
      if (Field.DataSet.RecNo mod 2) = 0 then begin
         Grid.Canvas.Brush.Color := Cor1;
         Grid.Canvas.Font.Color  := clBlack;
      end else begin
         Grid.Canvas.Brush.Color := Cor2;
         Grid.Canvas.Font.Color:= clBlack;
      end;
   end else begin
      Grid.Canvas.Brush.Color := CorSelec;
      Grid.Canvas.Font.Color:= CorSelecFnt;
   end;
   Grid.Canvas.FillRect(Rect);
   Grid.DefaultDrawDataCell(Rect,Field,State);
end;

procedure cordogridsel(const Grid: TdbGrid; Rect: TRect; Field: TField; State: TGridDrawState);
begin
  if not grid.SelectedRows.CurrentRowSelected then begin
     if (Field.DataSet.RecNo mod 2) = 0 then begin
        Grid.Canvas.Brush.Color := clMoneyGreen;
        Grid.Canvas.Font.Color:= clBlack;
     end else begin
        Grid.Canvas.Brush.Color := clCream;
        Grid.Canvas.Font.Color:= clBlack;
     end;
  end else begin
     Grid.Canvas.Brush.Color := clNavy;
     Grid.Canvas.Font.Color:= clWhite;
  end;
  Grid.Canvas.FillRect(Rect);
  Grid.DefaultDrawDataCell(Rect,Field,State);
end;

function buscaentre2(Msg, txtAntes, txtDepois : String) : String;
Var
   posAntes,posDepois,tam : Integer;
begin
   Result := '';
   posAntes  := pos(txtAntes,Msg);
   posDepois := pos(txtDepois,Msg);
   if posAntes > 0 then begin
      posAntes := posAntes + length(txtAntes);
      tam := posDepois - posAntes;
      if tam > 0 then Result := copy(Msg,posAntes,tam);
   end;
end;

procedure cordotitulo(Grid: TdbGrid;Campo :String);
Var
   i: Integer;
Begin
   For i := 0 to Grid.Columns.Count -1 do
       if uppercase(Grid.Columns[i].Field.FieldName) = uppercase(Campo) then
          Grid.Columns[i].Title.Color := clSkyBlue
       else Grid.Columns[i].Title.Color := clBtnFace;
end;

function PegaNumCombo(Texto :String) : Integer;
begin
   result := strtoint(copy(Texto,1,pos(' - ',Texto)-1));
end;

Procedure addlinhas(Var Onde : TStringList; Qtd:Integer);
Var
  i : Integer;
begin
  for i := 1 to qtd do Onde.Add('');
end;

function formatacgccpf(Texto : String;Valida : Boolean = False ; Alerta : Boolean = False) : String;
Var
   Aux : String;
begin
   Aux := Sonumero(Texto);
   if (Length(Aux) = 14) and (copy(Aux,1,3) = '000') and not TestaCgc(Aux) then
       Aux := copy(aux,4,11);
   if Length(Aux) = 11 then begin
      if Valida and not TestaCpf(Aux) then Aux := '';
      if Alerta and (Aux = '') then ShowMessage('C.P.F. informado é inválido.');
      if Aux <> '' then Aux := copy(aux,1,3) + '.' + copy(Aux,4,3) + '.'  + copy(Aux,7,3) + '-' + copy(Aux,10,2);
   end else if Length(Aux) = 14 then begin
      if Valida and not TestaCgc(Aux) then Aux := '';
      if Alerta and (Aux = '') then ShowMessage('C.N.P.J. informado é inválido.');
      if Aux <> '' then Aux := copy(aux,1,2) + '.' + copy(Aux,3,3) + '.'  + copy(Aux,6,3) + '/' + copy(Aux,9,4) + '-' + copy(Aux,13,2);
   end Else begin
      Aux := Texto;
      if Alerta then ShowMessage('O documento informado não é CNPJ nem CPF válido.');
      if Valida then Aux := '';
   end;
   Result := Aux;
end;

function floatponto(Texto : String) : String;
// Transforma uma string com ponto e/ou virgula em uma string flutuante com ponto
Var
   pponto,pvirg,onde : integer;
   aux : string;
begin
   if pos('-',texto) > 0 then aux := '-';
   pponto := posr('.',Texto);
   pvirg  := posr(',',Texto);
   onde   := pponto;
   if Onde < pvirg then Onde := pvirg;
   if Onde > 0 then begin
      Texto := sonumero(copy(Texto,1,Onde-1)) + '.' + sonumero(copy(Texto,Onde + 1,10));
   end;
   if (Texto = '.') or (Texto = '') then texto := '0';
   Aux := Aux+Texto;
   if pos('--',Aux) > 0 then Aux := Copy(Aux,2,20);
   Result := Aux;
end;

function floatvirgula(Texto : String) : String;
// Transforma uma string com ponto e/ou virgula em uma string flutuante com vírgula
Var
   pponto,pvirg,onde : integer;
   aux : string;
begin
   if pos('-',texto) > 0 then aux := '-';
   pponto := posr('.',Texto);
   pvirg  := posr(',',Texto);
   onde   := pponto;
   if Onde < pvirg then Onde := pvirg;
   if Onde > 0 then begin
      Texto := sonumero(copy(Texto,1,Onde-1)) + ',' + sonumero(copy(Texto,Onde + 1,10));
   end;
   if (Texto = '.') or (Texto = '') then texto := '0';
   if (pos('-',texto) = 0) and (aux <> '') then
      Result := Aux+Texto
   else Result := Texto;
end;

Function posr(OQue: String; Onde: String) : Integer;
//  Procura uma string dentro de outra, da direita para esquerda
//  Retorna a posição onde foi encontrada ou 0 caso não seja encontrada
var
   Pos   : Integer;
   Tam1  : Integer;
   Tam2  : Integer;
   Achou : Boolean;
begin
   Tam1   := Length(OQue);
   Tam2   := Length(Onde);
   Pos    := Tam2-Tam1+1;
   Achou  := False;
   while (Pos >= 1) and not Achou do
         begin
         if Copy(Onde, Pos, Tam1) = OQue then
            begin
            Achou := True
            end
         else
            begin
            Pos := Pos - 1;
            end;
         end;
   Result := Pos;
end;

function dividelin(Texto: String; lin: Integer;larg : Integer = 50): String;
Var
   i     : integer;
   lista : TStringList;
   linha : String;
begin
   lista := TStringList.Create;
   linha := '';
   try
      while Texto <> '' do begin
         if Length(Texto) <= larg then begin
            linha := Texto;
            Texto := '';
            lista.Add(linha);
         end else begin
            linha := copy(Texto,1,larg);
            i     := Larg;
            while (copy(linha,i,1) <> ' ') and (copy(linha,i,1) <> '') and (i>0) do Dec(i);
            if i = 0 then i := larg;
            linha := copy(linha,1,i);
            Texto := copy(texto,length(linha)+1,length(Texto));
            lista.Add(linha);
         end;
      end;
      if lin <= lista.Count then Result := lista.Strings[lin-1];
   Finally
      lista.Free;
   end;
end;

function se(Expressao : Boolean; Valor1, Valor2 : Variant) :Variant;
begin
   if Expressao then Result := Valor1 else Result := Valor2;
end;

function CheckCC(c: string): Integer;
   var
   card: string[21];
   Vcard: array[0..21] of Byte absolute card;
   Xcard: Integer;
   Cstr: string[21];
   y, x: Integer;
begin
   Cstr := '';
   FillChar(Vcard, 22, #0);
   card := c;
   for x := 1 to 20 do
      if (Vcard[x] in [48..57]) then
         Cstr := Cstr + chr(Vcard[x]);
         card := '';
         card := Cstr;
         Xcard := 0;
         if not odd(Length(card)) then
            for x := (Length(card) - 1) downto 1 do
               begin
                  if odd(x) then
                     y := ((Vcard[x] - 48) * 2)
                  else
                     y := (Vcard[x] - 48);
                     if (y >= 10) then
                        y := ((y - 10) + 1);
                        Xcard := (Xcard + y)
                  end
                  else
                     for x := (Length(card) - 1) downto 1 do
                        begin
                           if odd(x) then
                              y := (Vcard[x] - 48)
                           else
                              y := ((Vcard[x] - 48) * 2);
                           if (y >= 10) then
                              y := ((y - 10) + 1);
                              Xcard := (Xcard + y)
                           end;
                           x := (10 - (Xcard mod 10));
                           if (x = 10) then
                              x := 0;
                           if (x = (Vcard[Length(card)] - 48)) then
                              Result := Ord(Cstr[1]) - Ord('2')
                           else
                              Result := 0
end;

function soletra(Texto: String): String;
var
   aux : string;
   i   : integer;
Begin
   aux   := '';
   Texto := Uppercase(Texto);
   for i := 1 to length(Texto) do begin;
       if Texto[i] in ['A'..'Z'] then Aux := Aux + copy(Texto,i,1);
   end;
   soletra := Aux;
end;

function sem_espacos(Texto : String) : String;
Var
   i : Integer;
Begin
   Result := '';
   For i := 1 to length(Texto) do if Texto[i] <> ' ' then Result := Result + copy(Texto,i,1);
end;

function validasequencia(Texto : String;Valor : Boolean = False;Ordena : Boolean = False) : String;
// Recebe um texto com dados separados por vírgula e devolve uma sequencia válida
// Ordena opcionalmente. Retorna '' se a sequencia for inválida
// Por Márcio Henrique - 2004
Var
   aux : TStringList;
   i   : Integer;
   tmp : String;
begin
   aux    := TStringList.Create;
   Result := '';
   try
      while length(Texto) > 0 do begin
         if pos(',',Texto) > 0 then begin
            tmp   := copy(Texto,1,pos(',',Texto)-1);
            if valor then if stringtofloat(tmp) then tmp := FloatToStr(strtofloat(tmp))
               else tmp := '';
            Texto := copy(Texto,pos(',',texto)+1,length(Texto));
         end else begin
            tmp   := Texto;
            if valor then if stringtofloat(tmp) then tmp := FloatToStr(strtofloat(tmp))
               else tmp := '';
            Texto := '';
         end;
         if tmp <> '' then aux.Add(tmp)
      end;
      if Ordena then Aux := OrdenaString(Aux);
      for i := 0 to aux.Count - 1 do
         Result := Result + IfThen(i=0,'',#44) + Aux.Strings[i];
   finally
      aux.Free;
   end;
end;

function OrdenaString(Lista : TStringList) : TStringList;
// Ordena uma Stringlist
// Por Márcio Henrique - 2004
// ---------------------------
   function pos_desordem(l : TStringList) : Integer;
   Var ini   : integer;
       maior : Boolean;
   begin
      Result := -1;
      for ini := 0 to l.Count -2 do begin
         if StringToFloat(l.Strings[ini]) and StringToFloat(l.Strings[ini+1]) then
            maior := StrToFloat(l.Strings[ini]) > StrToFloat(l.Strings[ini+1])
         else maior := UpperCase(tiraacento(l.Strings[ini])) > UpperCase(tiraacento(l.Strings[ini+1]));
         if  maior  then begin
             Result := ini;
             break;
         end;
      end;
   end;
Var
   i : Integer;
   a : String;
begin
   if lista.Count > 0 then begin
      i := pos_desordem(Lista);
      while i <> -1 do begin
         a := Lista.Strings[i];
         Lista.Strings[i]   := lista.Strings[i+1];
         Lista.Strings[i+1] := a;
         i := pos_desordem(Lista);
      end;
   end;
   Result := Lista;
end;

function pegaseq(Texto : String ; posicao : Integer; sep : Char = #44) : String;
// Retorna a string n de uma sequencia do tipo: 23,78,58,90 ou 10|25|52|58
// Exemplo: pegaseq(tipo,2) = 78
Var             
   conta : Integer;
   tmp   : String;
begin
   Result := '';
   conta  := 1;
   while length(Texto) > 0 do begin
      if pos(sep,Texto) > 0 then begin
         tmp   := copy(Texto,1,pos(sep,Texto)-1);
         Texto := copy(Texto,pos(sep,texto)+1,length(Texto));
      end else begin
         tmp   := texto;
         Texto := '';
      end;
      if conta = posicao then begin
         result := tmp;
         texto  := '';
      end;
      inc(conta);
   end;
end;

function ContaSeq(Texto : String; sep : Char = #44) : Integer;
// Retorna a quantidade de dados em uma sequencia do tipo: 23,78,58,90
// Exemplo: contaseq(tipo) = 4
Var
   conta : Integer;
   tmp   : String;
begin
   conta  := 0;
   while length(Texto) > 0 do begin
      if pos(sep,Texto) > 0 then begin
         tmp   := copy(Texto,1,pos(sep,Texto)-1);
         Texto := copy(Texto,pos(sep,texto)+1,length(Texto));
         inc(conta);
      end else begin
         tmp := texto;
         inc(conta);
         texto := '';
      end;
   end;
   Result := conta;
end;

function apagaseq(Texto : String;indice : Integer;sep : Char = #44) : String;
// Retorna uma sequencia excluindo o item informado
// Exemplo: apagaseq('1,2,3,4',2) = '1,3,4'
Var
   i   : Integer;
   aux : String;
begin
   Result := '';
   for i := 1 to contaseq(Texto,Sep) do begin
      Aux := pegaseq(Texto,i,sep);
      if i <> Indice then begin
         if Result <> '' then Result := Result + Sep;
         Result := Result + Aux;
      end;
   end;
end;

function achaposseq(Texto,Oque : String;sep : Char = #44) : Integer;
// Retorna a posicao de um item em uma sequencia
// Exemplo: achaposseq('1,ab,3,4','ab') = 2
Var
   i   : Integer;
   aux : String;
begin
   Result := 0;
   for i := 1 to contaseq(Texto,Sep) do begin
      if pegaseq(Texto,i,sep) = Oque then begin
         Result := i;
         break;
      end;
   end;
end;
function ajusta(Texto,posicoes : String) : String;
// Ajusta um texto para impressão
// Ex.: Ajusta('7878|Produto teste|52,00','10,20,30');
Var
   i,q,j,k : Integer;
   Lin,aux1,aux2 : String;
begin
   i   := 1;
   q   := strtoint(pegaseq(posicoes,1,'|'));
   lin := replicate(' ',q);
   while pegaseq(posicoes,i,'|') <> '' do begin
      Aux1 := pegaseq(Texto,i,'|');
      q    := strtoint(pegaseq(posicoes,i,'|'));
      inc(i);
      aux2 := pegaseq(posicoes,i,'|');
      if aux2 = '' then aux2 := '0';
      j    := strtoint(aux2);
      if j > 0 then Aux1 := padr(Aux1,j-q);
      lin  := Lin + Aux1;
   end;
   Result := Lin;
end;
function EmailValido(Texto : string): Boolean;
begin
   if TregEx.Ismatch(Texto,'^((?>[a-zA-Z\d!#$%&''*+\-/=?^_`{|}~]+\x20*' +
          '|"((?=[\x01-\x7f])[^"\\]|\\[\x01-\x7f])*"\' +
          'x20*)*(?<angle><))?((?!\.)(?>\.?[a-zA-Z\d!' +
          '#$%&''*+\-/=?^_`{|}~]+)+|"((?=[\x01-\x7f])' +
          '[^"\\]|\\[\x01-\x7f])*")@(((?!-)[a-zA-Z\d\' +
          '-]+(?<!-)\.)+[a-zA-Z]{2,}|\[(((?(?<!\[)\.)' +
          '(25[0-5]|2[0-4]\d|[01]?\d?\d)){4}|[a-zA-Z\' +
          'd\-]*[a-zA-Z\d]:((?=[\x01-\x7f])[^\\\[\]]|' +
          '\\[\x01-\x7f])+)\])(?(angle)>)$') then result := true
   else result := false;
end;

function validaemail(Texto : String) : String;
Var
    i : Integer;
   a1 : String;
Const Validos = '@0123456789zxcvbnmlkjhgfdsaqwertyuiop-_.';
begin
   texto := lowercase(Texto);
   for i := 1 to length(Texto) do begin
      if pos(Texto[i],Validos) > 0 then a1 := A1 + Texto[i];
   end;
   if pos('.@',A1) > 0 then a1 := '';
   if pos('@.',A1) > 0 then a1 := '';
   if a1= '@' then a1 := '';
   result := a1;
end;

function formatapreco(precobase,valentrada : Currency; nparcelas : word;Indice : Real;ComEntrada : Boolean ;var entrada : Currency;Var parcela : Currency) : String;
Var
   Ent,Qtd,Valor : String;
   vEnt,vValor   : Currency;
   vQtd          : Word;
begin
   if ComEntrada then begin
      if valentrada = 0 then begin
         vEnt   := arredondapreco((precobase * Indice) / nparcelas,nparcelas = 1);
         vValor := vEnt;
         vQtd   := nparcelas - 1;
         Ent    := FormatFloat('0.00',vEnt) + ' + ';
         Valor  := FormatFloat('0.00',vValor);
         Qtd    := inttostr(vQtd) + ' x ';
         if vQtd = 0 then begin
            vValor := 0;
            Valor  := '';
            Qtd    := '';
            Ent    := FormatFloat('0.00',vEnt);
         end;
      end else begin
         vEnt   := valentrada;
         vQtd   := nparcelas - 1;
         vValor := arredondapreco((precobase * Indice - vEnt) / vQtd,vQtd = 1);
         Ent    := FormatFloat('0.00',vEnt) + ' + ';
         Valor  := FormatFloat('0.00',vValor);
         Qtd    := inttostr(vQtd) + ' x ';
         if vQtd = 0 then begin
            vValor := 0;
            Valor  := '';
            Qtd    := '';
         end;
      end;
   end else begin
      Ent    := '0,00 + ';
      vValor := arredondapreco((precobase * Indice) / nparcelas,nparcelas = 1);
      vQtd   := nparcelas;
      Qtd    := inttostr(vQtd) + ' x ';
      Valor  := FormatFloat('0.00',vValor);
      if vQtd = 0 then begin
         Valor  := '';
         Qtd    := '';
         vValor := 0;
      end;
   end;
   entrada := vEnt;
   parcela := vValor;
   if Qtd = '1 x ' then qtd := '';
   Result  := Ent + Qtd + Valor;
end;

function arredondapreco(Valor : Currency; Inteiro : Boolean) : Currency;
Var
   dif : Real;
begin
   if inteiro then Result := int(Valor)
   else begin
      dif := Valor - int(valor);
      if      dif > 0.75  then dif := 1
      else if dif > 0.5   then dif := 0.5
      else if dif > 0.25  then dif := 0.5
      else if dif > 0     then dif := 0.0
      else dif := 0;
      result := int(Valor) + dif;
   end;
end;

function geraarqseq(Pasta,IniArq,Extensao : String; qtd : Integer) : String;
Var
   i,Maisv    : Integer;
   Nome       : String;
   UltData,dd : TDatetime;
begin
   i       := 0;
   Maisv   := 1;
   UltData := now + 1000;
   Repeat
      inc(i);
      Nome := IniArq + '_' + strzero(inttostr(i),3) + '.' + Extensao;
      if FileExists(Pasta + Nome) then begin
         dd := strtodatetime(FileDate(Pasta + Nome));
         if dd < UltData then begin
            ultdata := dd;
            Maisv   := i;
         end;
      end else begin
         Maisv := i;
         i     := qtd;
      end;
   Until i = Qtd;
   Result := Pasta + IniArq + '_' + strzero(inttostr(Maisv),3) + '.' + Extensao;
end;

{Converte a primeira letra para maiúscula e o resto para minúsculas}
{Alterada por Alexs e Marcio}
function PrimeirasMaiusculas(Texto: string): string;
var
  Index: Integer;
  Espaco,
  Palavra: string;
begin
  Result := '';
  Espaco := '';
  if Texto <> '' then begin
    Texto := AnsiLowerCase(Texto);
    Index := 1;
    while (Index <= Length(Texto)) do begin
      Palavra := '';
      if Texto[Index] = ' ' then Inc(Index)
        else begin
           while (Index <= Length(Texto)) and (Texto[Index] <> ' ') do begin
              Palavra := Palavra + Texto[Index];
              Inc(Index);
           end;
        end;

        if pos(Palavra,'di|du|das|dos|von|der') = 0 then
           Palavra := AnsiUpperCase(Copy(Palavra,1,1)) + AnsiLowerCase(Copy(Palavra, 2, Length(Palavra)));

        Result := Result + Espaco + Palavra;
        inc(index);
        Espaco := ' ';
    end;
  end;
end;

function SeqNumLetra(texto : String) : String;
Var
   aux : String;
   i   : Integer;
begin
   if texto = '' then exit;
   i := length(texto);
   repeat
      if texto[i] in ['0'..'9'] then aux := copy(texto,i,1) + aux;
      dec(i);
   Until (i > 0) and not (texto[i] in ['0'..'9']);
   if aux <> '' then aux := inttostr(strtoint(aux)+1);
   Result := copy(texto,1,i) + aux;
end;

function dizmes(NumMes : Word) : String;
// Retorna o Mês por extenso a partir do nº inteiro
var
   aux : string;
begin
   aux := 'Janeiro  FevereiroMarço    Abril    Maio     Junho    Julho    Agosto   Setembro Outubro  Novembro Dezembro';
   result := copy(Aux,NumMes*9-8,9);
end;

function pontodir(Texto : String;Tam : Smallint) : String;
Var
   posini : Smallint;
begin
   posini := tam - length(trim(Texto));
   while Length(Texto) < tam do Texto := Texto + '.';
   pontodir := Texto;
end;

function passwd(Texto : String) : String;
   Var Car1,Car2,Car3,Bck,i : Smallint;

begin

   Bck    := 0;
   Result := '';

   for i := 1 TO length(Texto) do begin
      Car1 := ord(Texto[i]);
      if (i < length(Texto)) then car2 := ord(Texto[i+1]) else Car2 := 0;
      Car3 := Car1 + Car2 + Bck;
      if Car3 > 255 then Car3 := Car3 - 255;
      Bck  := Car3;
      Result := Result + CHR(Car3);
   end;

end;

{*****************************************************************************
* Funcao Nome : Ajustar(<String>,<Integer>,<Direcao>)
* Objetivo    : Ajustar comprimento da string no tamanho especificado
* Autor       : Jorge Henrique
* Uso         : Ajusta(S, T, D)
* Retorno     : String <S> com <T> comprimento na direcao <D>.
******************************************************************************}
function Ajustar(S: String; T: Integer; D: String): String;
var I: Integer;
begin
  if D='D' then if Length(S) < T then For I:=length(S) to T-1 do S:=S+' ';
  if D='E' then if Length(S) < T then For I:=length(S) to T-1 do S:=' '+S;
  if D='C' then if Length(S) < T then
  begin
    For I:=1 to T do
    begin
      if Length(S) < T then S:=' '+S;
      if Length(S) < T then S:=S+' ';
    end;
  end;
  Result:=S;
end;

function sonumleft(Valor :Extended;Tam,Dec : Integer) : String;
Var
   i   : Integer;
   fmt : String;
begin
   for i := 1 to tam - dec do fmt := fmt + '0';
   if dec > 0 then fmt := fmt + '.';
   for i := 1 to dec do fmt := fmt + '0';
   result := sonumero(FormatFloat(fmt,Valor));
end;

{*****************************************************************************
* Funcao Nome : Preench(<String>,<Integer>,<Direcao><Chr>)
* Objetivo    : Ajustar comprimento da string no tamanho especificado com chr
* Autor       : Jorge Henrique ex-careca
* Uso         : Preench(S, T, D, Chr)
* Retorno     : String <S> com <T> comprimento na direcao <D> preenchendo c/chr
******************************************************************************}
function Preenche(S: String; T: Integer; D: String; Chr: String): String;
var I: Integer;
begin
  Chr:=Copy(Chr,1,1);
  if D='D' then if Length(S) < T then For I:=length(S) to T-1 do S:=S+Chr;
  if D='E' then if Length(S) < T then For I:=length(S) to T-1 do S:=Chr+S;
  if D='C' then if Length(S) < T then
  begin
    For I:=1 to T do
    begin
      if Length(S) < T then S:=Chr+S;
      if Length(S) < T then S:=S+Chr;
    end;
  end;
  Result:=S;
end;

function StrTempo2Min(S: String): Integer;
Var
   letra, aux : String;
begin
   letra  := LowerCase(soletra(s));
   aux    := Sonumero(s);
   if aux = '' then aux := '0';
   result := strtoint(aux);
   if (letra = 'h') or (letra='hora') then Result := Result * 60;
   if (letra = 'dt') or (letra='dia trab') then Result := Result * 60 * 8;
   if (letra = 'd') or (letra='dia') then Result := Result * 60 * 24;
   if (letra = 's') or (letra='sem') then Result := Result * 60 * 24 * 7;
   if (letra='mes') or (letra='meses') then Result := Result * 60 * 24 * 30;
end;

//Validação de Data
function ValiData(Data: String):Boolean;
var
   TesteData: TDateTime;
begin
   Result := True;
   try
      TesteData := StrToDate(Data);
   except
      Result := False;
   end;
end;

function ValiHora(Hora: String):Boolean;
var
   TesteHora: TTime;
begin
   Result := True;
   try
      TesteHora := StrToTime(Hora);
   except
      Result := False;
   end;
end;

function traduzmes(texto :String) : String;
Const Meses = 'JAN01FEB02MAR03APR04MAI05JUN06JUL07AUG08SEP09OCT10NOV11DEC12';
begin
   Result := '00';
   if pos(texto,Meses) > 0 then result := copy(Meses,pos(texto,Meses)+3,2);
end;

function ExecutaPrograma(Programa: String): String;
begin
   Case ShellExecute(Application.Handle,'Open',PChar (Programa), Nil,Nil,SW_SHOWNORMAL) Of
      0:Result:=('Não há recursos suficientes para executar operação!');
      ERROR_FILE_NOT_FOUND:Result:=('Arquivo não encontrado!');
      ERROR_PATH_NOT_FOUND:Result:=('PATH Não encontrado!');
      ERROR_BAD_FORMAT:Result:=('Formato de arquivo EXE inválido!');
      SE_ERR_ACCESSDENIED:Result:=('Não foi possível acessar arquivo especificado!');
      SE_ERR_ASSOCINCOMPLETE:Result:=('Associação entre arquivo e executável inválido!');
      SE_ERR_DDEBUSY:Result:=('Não foi possível executar operação pois uma outra operação DDE esta sendo executada!');
      SE_ERR_DDEFAIL:Result:=('Transação DDE falhou!');
      SE_ERR_DDETIMEOUT:Result:=('TimeOut em transação DDE!');
      SE_ERR_DLLNOTFOUND:Result:=('Não foi possível encontrar DLL necessária para a operação!');
      SE_ERR_NOASSOC:Result:=('Associação entre arquivo e executável inválido!');
      SE_ERR_OOM:Result:=('Não há memória disponível para executar operação!');
      SE_ERR_SHARE:Result:=('Violação no compartilhamento do arquivo');
   else
      Result:=('Comando sendo executado!')
   end;
end;

function UltDataDoMes(Data: TDateTime): TDate;
var
   d,m,a: Word;
   dt: TDateTime;
begin
   DecodeDate(Data, a,m,d);
   Inc(m);
   if m = 13 then
      begin
      m := 1;
      end;
   dt := EncodeDate(a,m,1);
   dt := dt - 1;
   DecodeDate(dt, a,m,d);
   Result := dt;
end;

Function PrimeiroDoMes(Data : TDateTime) : TDate;
var Ano, Mes, Dia : word;
begin
   DecodeDate (Data, Ano, Mes, Dia);
   Dia := 1;
   Result := EncodeDate (Ano, Mes, Dia);
end;

Function DatacomBarras(Texto: String) : String;
// Esta função devolve uma string de data validada para texto no formato YYYYMMDD ou YYMMDD
// Em caso de texto inválido retorna 01/01/1900
Var
   Aux : String;
   dia,Mes,Ano,pini : Word;
begin
   pini := 0;
   dia  := 1;
   mes  := 1;
   ano  := 1900;
   if Length(Texto) = 8 then pini := 2;
   aux  := copy(Texto,5+pini,2);
   try dia := strtoint(aux); except end;
   aux := copy(Texto,3+pini,2);
   try mes := strtoint(aux); except end;
   aux := copy(Texto,1,2+pini);
   if pini = 0 then if strtoint(aux) < 20 then aux := '20' + aux else aux := '19' + aux;
   try ano := strtoint(aux); except end;
   result  := formatdatetime('DD/MM/YYYY',EncodeDate(ano,mes,dia));
end;

Function DatacomBarrasBrt(Texto: String) : String;
// Esta função devolve uma string de data validada para texto no formato DDMMYY ou DDMMYY
// Em caso de texto inválido retorna 01/01/1900
// Formato Britânico Dia/Mes/Ano
Var
   Aux : String;
   dia,Mes,Ano : Word;
begin
   dia  := 1;
   mes  := 1;
   ano  := 1900;
   aux  := copy(Texto,1,2);
   try dia := strtoint(aux); except end;
   aux := copy(Texto,3,2);
   try mes := strtoint(aux); except end;
   aux := copy(Texto,5,4);
   if Length(aux) =2 then if strtoint(aux) < 20 then aux := '20' + aux else aux := '19' + aux;
   try ano := strtoint(aux); except end;
   result  := formatdatetime('DD/MM/YYYY',EncodeDate(ano,mes,dia));
end;

Function arredondavisa(Valor : Real) : Real;
begin
   result := SimpleRoundTo(valor,-2);
end;


function Extenso(Valor: Real; Reais: Boolean; Masculino : Boolean): String;
var Centavos, Centena, Milhar, Milhao, Texto, Msg : String;
Const
  UnidadesM : Array[1..9] of String = (' Um',' Dois',' Três',' Quatro',' Cinco',' Seis',' Sete',' Oito',' Nove');
  UnidadesF : Array[1..9] of String = (' Uma',' Duas',' Três',' Quatro',' Cinco',' Seis',' Sete',' Oito',' Nove');
  Dez       : Array[1..9] of String = (' Onze',' Doze',' Treze',' Quatorze',' Quinze',' Dezesseis',' Dezessete',' Dezoito',' Dezenove');
  Dezenas   : Array[1..9] of String = (' Dez',' Vinte',' Trinta',' Quarenta',' Cinquenta',' Sessenta',' Setenta',' Oitenta',' Noventa');
  Centenas  : Array[1..9] of String = (' Cento',' Duzentos',' Trezentos',' Quatrocentos',' Quinhentos',' Seiscentos',' Setecentos',' Oitocentos',' Novecentos');

  function Ifs(Expressao : Boolean; Casoverdadeiro, CasoFalso: String): String;
  begin
    if Expressao then Result:=CasoVerdadeiro else Result:=CasoFalso;
  end;

  function MiniExtenso(Trio: String): String;
  var Unidade, Dezena, Centena : String;
  begin
    Unidade:=''; Dezena:=''; Centena:='';
    if (Trio[2]='1') and (Trio[3] <> '0') then
    begin
      Unidade:=Dez[StrToInt(Trio[3])];
      Dezena:='';
    end
    else begin
      if Trio[2] <> '0' then Dezena:=Dezenas[StrToInt(Trio[2])];

      if Masculino then begin
         if Trio[3] <> '0' then Unidade:=UnidadesM[StrToInt(Trio[3])];
      end else if Trio[3] <> '0' then Unidade:=UnidadesF[StrToInt(Trio[3])];
    end;
    if (Copy(Trio,1,1) ='1') and (Unidade='') and (Dezena='') then Centena:='Cem'
    else if Copy(Trio,1,1) <>'0' then Centena:=Centenas[StrToInt(Copy(Trio,1,1))] else Centena:='';
    Result:=Centena+ifs((Centena<>'') and ((Dezena<>'') or (Unidade<>'')),' e','')+Dezena+ifs((Dezena<>'') and (Unidade<>''),' e','')+Unidade;
  end;

begin
  if (Valor > 999999.99) or (Valor < 0) then
  begin
    Msg:='O valor está fora do intervalo permitido.';
    Msg:=Msg+'O número deve ser maior ou igual a zero e menor ou igual a 999.999,99.';
    ShowMessage(Msg);
    Result:='';
    Exit;
  end;
  if Valor = 0 then
  begin
    Result:='';
    Exit;
  end;
  Texto:=FormatFloat('000000.00',Valor);
  Milhar:=MiniExtenso(Copy(Texto,1,3));
  Centena:=MiniExtenso(Copy(Texto,4,3));
  Centavos:=MiniExtenso('0'+Copy(Texto,8,2));
  Result:=Milhar;
  If Milhar <> '' then if Copy(Texto,4,3)='000' then Result:=Result+' Mil'+Iif(Reais,' Reais','') else Result:=Result+' Mil';
  if (((Copy(Texto,4,2)='00') and (Milhar<>'') and (Copy(Texto,6,1)<>'0')) or (Centavos='')) and (Centena<>'') then if Result<>'' then Result:=Result+' e';
  if (Milhar+Centena <> '') then Result:=Result+Centena;
  if (Milhar = '') and (Copy(Texto,4,3)='001') then Result:=Result+' Real' else if (Copy(Texto,4,3)<> '000') then Result:=Result+Iif(Reais,' Reais','');
  Result:=Trim(Result);
  if Centavos='' then
  begin
    Exit
  end else begin
    if Milhar+Centena='' then Result:=Centavos else Result:=Result+' e '+Centavos;
    if (Copy(Texto,8,2)='01') and (Centavos<>'') then Result:=Result+Iif(Reais,' Centavo','') else Result:=Result+Iif(Reais,' Centavos','');
    Result:=Trim(Result);
  end;
end;

function iif(Condicao: Boolean; rVerdade: Variant; rFalso: Variant): Variant;
begin
  if Condicao then Result:=rVerdade else Result:=rFalso;
end;

function FormataCep (sValue:String): String;
begin
  sValue := StringReplace(sValue,'.','',[rfReplaceAll, rfIgnoreCase]);
  sValue := StringReplace(sValue,'-','',[rfReplaceAll, rfIgnoreCase]);
  sValue := StrZero (sValue,8);
  sValue := copy(sValue,1,5) + '-' + copy (sValue,6,3);
  Result := sValue;
end;

function ParamTimestamp(Tempo : TDateTime) : String;
begin
   Result := 'CAST(' + QuotedStr(IntToStr(YearOf(Tempo)) + '.' +
             IntToStr(MonthOf(Tempo)) + '.' +
             IntToStr(DayOf(Tempo)) + ' ' +
             IntToStr(HourOf(Tempo)) + ':' +
             IntToStr(MinuteOf(Tempo))) +' AS TIMESTAMP)';
end;

function ParamDate(Data : TDateTime) : String;
begin
   Result := 'CAST(' + QuotedStr(IntToStr(YearOf(Data)) + '.' +
             IntToStr(MonthOf(Data)) + '.' +
             IntToStr(DayOf(Data))) + ' AS DATE)';
end;

procedure gravaini(usuario,chave,valor : String);
var
   Ini : TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
  try
     ini.WriteString(usuario,chave,valor);
  finally
    ini.Free;
  end;
end;

function leini(usuario,chave:String) : String;
Var
   Ini : TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
  try
     Result := ini.ReadString(usuario,chave,'');
  finally
    ini.Free;
  end;
end;

function divideparcelas(Valor : Currency; QtdParc, NumParc : Extended) : Currency;
Var
   ValParc,Valtot,Dif : Currency;
begin
   ValParc := strtofloat(formatfloat('0.00',Valor / QtdParc));
   Valtot  := ValParc * QtdParc;
   Dif     := Valor - Valtot;
   if NumParc = 1 then Result := ValParc + Dif else Result := ValParc;
end;

procedure SendToOpenOffice(aDataSet: TDataSet);
const ooBold: integer = 150; //150 = com.sun.star.awt.FontWeight.BOLD
var
   OpenDesktop, Calc, Sheets, Sheet: Variant;
   Connect, OpenOffice : Variant;
   i : Integer; // Coluna
   lin : Integer; // Linha
   col : Integer; // Coluna Visivel
   s:string;
begin
   Screen.Cursor := crSQLWait;
   try
      aDataset.DisableControls;
      aDataset.Last;
      // Cria o link OLE com o OpenOffice
      if VarIsEmpty(OpenOffice) then OpenOffice := CreateOleObject('com.sun.star.ServiceManager');
      Connect := not (VarIsEmpty(OpenOffice) or VarIsNull(OpenOffice));
      // Inicia o Calc
      OpenDesktop := OpenOffice.CreateInstance('com.sun.star.frame.Desktop');
      Calc        := OpenDesktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, VarArrayCreate([0, - 1], varVariant));
      Sheets      := Calc.Sheets;
      Sheet       := Sheets.getByIndex(0);
      // Cria linha de cabeçalho
      i := 0;
      col := 0;
      while i <= aDataset.FieldCount - 1 do begin
         if aDataset.Fields[i].Visible then begin
            Sheet.getCellByPosition(col,0).setString(aDataset.Fields[i].DisplayName);
            Sheet.getCellByPosition(col,0).getText.createTextCursor.CharWeight:= ooBold;
            inc(col);
         end;
         inc(i);
      end;
      // Preenche a planilha
      lin := 1;
      aDataset.First;
      while not aDataset.Eof do begin
         i := 0;
         col := 0;
         while i <= aDataset.FieldCount - 1 do begin
            if aDataset.Fields[i].Visible then begin
               if (aDataset.Fields[i].Value <> Null) then begin
                  if aDataset.Fields[i].DataType in [ftDate, ftTime, ftDateTime, ftTimestamp] then begin
                     if ((DateToStr(aDataset.Fields[i].Value) <> Null) and (DateToStr(aDataset.Fields[i].Value) <> ''))
                        then Sheet.getCellByPosition(col,lin).SetString(aDataset.Fields[i].AsDateTime);
                  end else if aDataset.Fields[i].DataType in [ftSmallint, ftInteger, ftLargeint] then begin
                     if ((IntToStr(aDataset.Fields[i].Value) <> Null) and (IntToStr(aDataset.Fields[i].AsInteger) <> ''))
                        then Sheet.getCellByPosition(col,lin).SetValue(aDataset.Fields[i].Value);
                  end else if aDataset.Fields[i].DataType in [ftFloat, ftCurrency, ftBCD, ftFMTBcd] then begin
                     if ((FloatToStr(aDataset.Fields[i].Value) <> Null) and (FloatToStr(aDataset.Fields[i].Value) <> ''))
                        then Sheet.getCellByPosition(col,lin).SetValue(aDataset.Fields[i].AsCurrency);
                  end else begin
                     if ((aDataset.Fields[i].Value <> Null) and (aDataset.Fields[i].Value <> ''))
                        then Sheet.getCellByPosition(col,lin).SetString(aDataset.Fields[i].AsString)
                        else Sheet.getCellByPosition(col,lin).SetString('');
                  end;
               end else Sheet.getCellByPosition(col,lin).SetString('');
               inc(col);
            end;
            inc(i);
         end;
         aDataset.Next;
         inc(lin);
      end;
      Sheet.getColumns.OptimalWidth := True;
   except
      on e:exception do showmessage(e.Message);
   end;
   aDataSet.First;
   aDataSet.EnableControls;
   OpenOffice := Unassigned;
   Screen.Cursor := crArrow;
end;

function e64Bits: Boolean;
type
  TIsWow64Process = function(Handle:THandle; var IsWow64 : BOOL) : BOOL; stdcall;
var
   hKernel32 : Integer;
   IsWow64Process : TIsWow64Process;
   IsWow64 : BOOL;
begin
   // http://msdn.microsoft.com/en-us/library/ms684139%28VS.85%29.aspx
   Result := False;
   hKernel32 := LoadLibrary('kernel32.dll');
   if (hKernel32 = 0) then RaiseLastOSError;
   @IsWow64Process := GetProcAddress(hkernel32, 'IsWow64Process');
   if Assigned(IsWow64Process) then begin
      IsWow64 := False;
      if (IsWow64Process(GetCurrentProcess, IsWow64)) then Result := IsWow64
      else RaiseLastOSError;
   end;
   FreeLibrary(hKernel32);
end;

procedure somamtecgrid(var Grid : TMdbGrid; textoCampos : String; Tipos : String = '0');
var
   soma : array of Real;
   i,y  : Integer;
   aux,mask : String;
   {Cacula a soma, a média ou a quantidade de acordo com a coluna do Mtecgrid e preenche o rodapé.
    O parâmetro textoCampos deve ser separado com | e pode ter 2 informaçoes:
    <i1> = Conteudo do rodape (use '=' para contar, '/' para média)
    <i2> = Mascara de formatação do valor:
          (use 'D' para alinhar ' direita a coluna e rodape
          ou 'd' para alinhar à direita apenas o rodape
          quando o campo não for de valor numérico) }
begin
   if not Grid.DataSource.DataSet.Active then exit;
   Grid.DataSource.DataSet.DisableControls;
   SetLength(soma, grid.Columns.Count);
   with grid.DataSource.DataSet do begin
      First;
      i := 0;
      while not Eof do begin
         for i := 0 to Grid.Columns.Count -1 do begin
             if Grid.Columns[i].Field.DataType in [ftSmallint, ftInteger, ftWord, ftFloat, ftCurrency, ftBCD,ftLargeint,ftFMTBcd] then
                if not grid.Columns[i].Field.IsNull then
                soma[i] := soma[i] + StrToFloat(FloatVirgula(Grid.Columns[i].Field.Value));
         end;
         inc(y);
         Next;
      end;
      First;
      for i := 0 to grid.Columns.Count -1 do begin
         aux  := pegaseq(textoCampos,i+1,'|');
         if aux = 'X' then continue;
         mask := '';
         if pos(';',aux) > 0 then begin
            mask := copy(aux,pos(';',aux)+1,length(aux));
            aux  := copy(aux,1,pos(';',aux)-1);
         end;
         if aux = '/' then begin
            if y > 0 then soma[i] := soma[i] / y;
            aux  := '';
         end;
         if aux = '=' then begin
            soma[i] := y;
            aux     := IntToStr(y);
            if Grid.Columns[i].Field.DataType in [ftSmallint, ftInteger, ftWord, ftFloat, ftCurrency, ftBCD, ftLargeint,ftFMTBcd] then
            aux := '';
         end;
         if mask = '' then mask := '#,##0.00 ';
         if Tipos = '0' then begin
            if Grid.Columns[i].Field.DataType in [ftSmallint, ftInteger, ftWord, ftFloat, ftCurrency, ftBCD, ftLargeint,ftFMTBcd] then begin
               Grid.Columns[i].Footer.Text       := aux + ' ' + FormatFloat(mask,soma[i]);
               Grid.Columns[i].Footer.Alignment  := taRightJustify;
               Grid.Columns[i].Title.Alignment   := taRightJustify;
            end else begin
               Grid.Columns[i].Footer.Text := Aux;
               if (mask = 'D') or (mask = 'd') then Grid.Columns[i].Footer.Alignment  := taRightJustify;
               if (mask = 'D')                 then Grid.Columns[i].Alignment         := taRightJustify;
            end;
         end;
         if Tipos = '1' then begin
            if Grid.Columns[i].Field.DataType in [ftWord, ftFloat, ftCurrency, ftBCD, ftFMTBcd] then begin
               Grid.Columns[i].Footer.Text       := aux + ' ' + FormatFloat(mask,soma[i]);
               Grid.Columns[i].Footer.Alignment  := taRightJustify;
               Grid.Columns[i].Title.Alignment   := taRightJustify;
            end;
         end;
      end;
   end;
   Grid.DataSource.DataSet.EnableControls;
end;

// Função para exportar Dados de uma IBQuery para planilha do Excel
// Por: Ricardo Scache Belardinuci - ri-taqua@ig.com.br
// Testada na Versão 2003 do Excel
// Ajustada por: Alexssandro de Souza Marcelino em 09/01/12
// IMPORTANTE: Declare a unit ComObj na 1º cláusula Uses - EXEMPLO:
procedure ExportarDadosParaExcel(Qry: TDataSet; RealComoInteiro: Boolean = True);
var
   Linha, Coluna, ValorCampoI, NumRegistros: integer;
   Planilha: variant;
   ValorCampoS: string;
begin
   if not Qry.IsEmpty then begin
      try
         Planilha:= CreateOleObject('Excel.Application');
         Planilha.Workbooks.Add(1);
         Planilha.Caption:='Exportação de Dados Para o Excel';
         Planilha.Visible:=True;
      except
         MessageBox(Application.Handle,'Provavelmente o Microsoft Excel não está instalado nessa máquina.','Atenção!',MB_OK+MB_ICONERROR);
         Abort;
      end;
      with Qry do begin
         DisableControls;
         Last;
         First;
         NumRegistros:= RecordCount;
         for Linha:=0 to  NumRegistros - 1 do begin
            for Coluna:=1 to FieldCount do begin
               // Para não exportar um determinado campo, atribua a Tag para -1
               if Fields[Coluna-1].Tag = 0 then begin
                  ValorCampoS:= Fields[Coluna-1].AsString;
                  // Se Desejar que os valores fracionados sejam exportados no formato original,
                  // então, ao chamar a função, passe no 2º parâmetro o valor False
                  // Chamada: ExportarDadosParaExcel(IBQuery1, False);
                  // Nesse caso, por exemplo, o valor |12,3| será exportado como |12,3|
                  if (Fields[Coluna-1] is TStringField) or (not RealComoInteiro) or (Trim(Fields[Coluna-1].AsString) ='') then begin

                     if Fields[Coluna-1].DataType in [ftFloat, ftCurrency, ftBCD, ftFMTBcd] then Planilha.Cells[Linha+2,Coluna]:= StrToCurr(ValorCampoS) else
                     if Fields[Coluna-1].DataType in [ftDate, ftTime, ftDateTime,ftTimestamp] then Planilha.Cells[Linha+2,Coluna]:= StrToDateTime(ValorCampoS) else
                     if Fields[Coluna-1].DataType in [ftSmallint, ftInteger, ftLargeint] then Planilha.Cells[Linha+2,Coluna]:= StrToInt(ValorCampoS)
                     else Planilha.Cells[Linha+2,Coluna]:= ValorCampoS;

                  end else begin
                     // Nesse caso, os valores fracionados serão exportados como inteiros
                     // e a função deverá ser chamada simplismente assim: ExportarDadosParaExcel(IBQuery1);
                     // Exemplo: o Valor |12,3| será exportado como |12|
                     ValorCampoI:= Fields[Coluna-1].AsInteger;
                     Planilha.Cells[Linha+2,Coluna]:=ValorCampoI;
                  end;
               end;
            end;
            Next;
         end;
         //Cabeçalho das Colunas
         for Coluna:=1 to FieldCount do begin
            // Para não exportar um determinado campo, abra o Fields Editor da
            // Query, selecione o Campo e altere o valor da Tag para -1
            if Fields[Coluna-1].Tag = 0 then begin
               ValorCampoS:= Fields[Coluna-1].DisplayLabel;
               Planilha.Cells[1,Coluna]:=ValorCampoS;
            end;
         end;
         First;
         EnableControls;
      end;
      Planilha.Columns.AutoFit;
   end else Application.MessageBox('Não há dados para serem exportados para o Excel!','Atenção',MB_OK+MB_ICONINFORMATION);
end;

function DialogPersonalizado(Msg: string; AType: TMsgDlgType; AButtons: TMsgDlgButtons;
   IndiceHelp: LongInt; DefButton: TMOdalResult = mrNone;
   sSim : string = '&Sim'; sNao: string = '&Não'; sCancelar:string = '&Canclear';
   sAbortar: string = '&Abortar'; sRepetir: string = '&Repetir'; sIgnorar: string = '&Ignorar';
   sTodos:string = '&Todos'; sAjuda: string = 'A&juda'): Word;
var
   I: Integer;
   Mensagem: TForm;
begin
   Mensagem := CreateMessageDialog(Msg, AType, Abuttons);
   Mensagem.HelpContext := IndiceHelp;
   with Mensagem do begin
      for i := 0 to ComponentCount - 1 do begin
         if (Components[i] is TButton) then begin
            if (TButton(Components[i]).ModalResult = DefButton) then begin
               ActiveControl := TWincontrol(Components[i]);
            end;
         end;
      end;
      if Atype = mtConfirmation then Caption := 'Confirmação'
      else if AType = mtWarning then Caption := 'Aviso'
      else if AType = mtError then Caption := 'Erro'
      else if AType = mtInformation then Caption := 'Informação';
   end;
   TButton(Mensagem.FindComponent('YES')).Caption := sSim;
   TButton(Mensagem.FindComponent('NO')).Caption := sNao;
   TButton(Mensagem.FindComponent('CANCEL')).Caption := sCancelar;
   TButton(Mensagem.FindComponent('ABORT')).Caption := sAbortar;
   TButton(Mensagem.FindComponent('RETRY')).Caption := sRepetir;
   TButton(Mensagem.FindComponent('IGNORE')).Caption := sIgnorar;
   TButton(Mensagem.FindComponent('ALL')).Caption := sTodos;
   TButton(Mensagem.FindComponent('HELP')).Caption := sAjuda;
   Result := Mensagem.ShowModal;
   Mensagem.Free;
end;

procedure PreencheJvDbComboBox(Combo : TJvDbComboBox; FieldItems, FieldValues, sSQL : String; Conexao : TSQLConnection; PrimeiroItem : String = '');
var Query : TSQLQuery;
begin
   try
      Query := TSQLQuery.Create(Nil);
      Query.SQLConnection := Conexao;
      Query.SQL.Text := sSQL;
      Query.Open;
      Combo.Items.Clear;
      Combo.Values.Clear;
      if PrimeiroItem <> EmptyStr then begin
         Combo.Items.Add(PrimeiroItem);
         Combo.Values.Add('-1');
      end;
      while not Query.Eof do begin
         Combo.Items.Add( Query.FieldByName(FieldItems).AsString );
         Combo.Values.Add(Query.FieldByName(FieldValues).AsString);
         Query.Next;
      end;
   finally
      FreeAndNil(query);
      if Combo.Items.Count > 0 then Combo.ItemIndex := 0;      
   end;
end;

procedure PreencheComboBoxEx(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TSQLConnection; PrimeiroItem : String);
var Query : TSQLQuery;
begin
   Query := TSQLQuery.Create(Nil);
   Query.SQLConnection := Conexao;
   Query.SQL.Text := Comando;
   Query.Open;
   Combo.Clear;
   if PrimeiroItem <> '' then Combo.Items.Add(PrimeiroItem);
   while not Query.Eof do begin
      Combo.ItemsEx.AddItem(Query.FieldByName(Resultado).AsString,Query.FieldByName(Chave).AsInteger,-1,-1,-1,Nil);
      Query.Next;
   end;
   Query.Close;
   FreeAndNil(Query);
   if Combo.Items.Count > 0 then Combo.ItemIndex := 0;
end;

procedure PreencheComboBoxExFD(Combo : TComboBoxEx; Chave : String; Resultado : String; Comando : String; Conexao : TFDConnection; PrimeiroItem : String);
var
   Query : TFDQuery;
begin
   Query := TFDQuery.Create(Nil);
   Query.Connection := Conexao;
   Query.SQL.Text := Comando;
   Query.Open;
   Combo.Clear;
   if PrimeiroItem <> '' then Combo.Items.Add(PrimeiroItem);
   while not Query.Eof do begin
      Combo.ItemsEx.AddItem(Query.FieldByName(Resultado).AsString,Query.FieldByName(Chave).AsInteger,-1,-1,-1,Nil);
      Query.Next;
   end;
   Query.Close;
   FreeAndNil(Query);
   if Combo.Items.Count > 0 then Combo.ItemIndex := 0;
end;

procedure PreencheComboBoxExADO(Combo: TComboBoxEx; Chave, Resultado, Comando: String; Conexao: TADOConnection; PrimeiroItem: String);
var Query : TADOQuery;
begin
   Query := TADOQuery.Create(Nil);
   Query.Connection := Conexao;
   Query.SQL.Text := Comando;
   Query.Open;
   Combo.Clear;
   if PrimeiroItem <> '' then Combo.Items.Add(PrimeiroItem);
   while not Query.Eof do begin
      Combo.ItemsEx.AddItem(Query.FieldByName(Resultado).AsString,Query.FieldByName(Chave).AsInteger,-1,-1,-1,nil);
      Query.Next;
   end;
   Query.Close;
   FreeAndNil(Query);
   if Combo.Items.Count > 0 then Combo.ItemIndex := 0;
end;

procedure PreencheCheckListbox(CheckList: TCheckListBox; Resultado,Comando: string; Conexao: TSQLConnection; Checados:Boolean = False);
var Query: TSQLQuery;
    i: Integer;
begin
   try
      Query := TSQLQuery.Create(Nil);
      Query.SQLConnection := Conexao;
      Query.SQL.Add(Comando);
      Query.Open;
      i := 0;
      CheckList.Items.Clear;
      while not Query.Eof do begin
         CheckList.Items.Add(Query.FieldByName(Resultado).AsString);
         CheckList.Checked[i] := Checados;
         Inc(i);
         Query.Next;
      end;
   finally
      FreeAndNil(Query);
   end;
end;

function GetAveCharSize(Canvas: TCanvas): TPoint;
var
  I: Integer;
  Buffer: array[0..51] of Char;
begin
  for I := 0 to 25 do Buffer[I] := Chr(I + Ord('A'));
  for I := 0 to 25 do Buffer[I + 26] := Chr(I + Ord('a'));
  GetTextExtentPoint(Canvas.Handle, Buffer, 52, TSize(Result));
  Result.X := Result.X div 52;
end;

function InputQueryPT(const ACaption, APrompt: string; var Value: string; Cap1 : string = '&OK'; Cap2 : string = '&Cancelar'): Boolean;
var
  Form: TForm;
  Prompt: TLabel;
  Edit: TEdit;
  DialogUnits: TPoint;
  ButtonTop, ButtonWidth, ButtonHeight: Integer;
begin
  Result := False;
  Form := TForm.Create(Application);
  with Form do
    try
      Font.Size   := 10;
      Font.Name   := 'Tahoma';
      Canvas.Font := Font;
      Canvas.Font.Size := 10;
      DialogUnits := GetAveCharSize(Canvas);
      BorderStyle := bsDialog;
      Caption := ACaption;
      ClientWidth := MulDiv(180, DialogUnits.X, 4);
      ClientHeight := MulDiv(63, DialogUnits.Y, 8);
      Position := poScreenCenter;
      Prompt := TLabel.Create(Form);
      with Prompt do
      begin
        Parent := Form;
        AutoSize := True;
        Left := MulDiv(8, DialogUnits.X, 4);
        Top := MulDiv(8, DialogUnits.Y, 8);
        Caption := APrompt;
      end;
      Edit := TEdit.Create(Form);
      with Edit do
      begin
        Parent := Form;
        Left := Prompt.Left;
        Top := MulDiv(19, DialogUnits.Y, 8);
        Width := MulDiv(164, DialogUnits.X, 4);
        MaxLength := 255;
        Text := Value;
        SelectAll;
      end;
      ButtonTop := MulDiv(41, DialogUnits.Y, 8);
      ButtonWidth := MulDiv(50, DialogUnits.X, 4);
      ButtonHeight := MulDiv(14, DialogUnits.Y, 8);
      with TButton.Create(Form) do
      begin
        Parent  := Form;
        Caption := cap1;
        Cursor  := crHandPoint;
        Height  := 27;
        Width   := 95;
        ModalResult := mrOk;  // Unit Controls
        Default := True;
        SetBounds(MulDiv(38, DialogUnits.X, 4), ButtonTop, ButtonWidth, ButtonHeight);
      end;
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := cap2;
        Cursor := crHandPoint;
        Height := 27;
        Width  := 95;
        ModalResult := mrCancel;
        Cancel := True;
        SetBounds(MulDiv(92, DialogUnits.X, 4), ButtonTop, ButtonWidth,ButtonHeight);
      end;
      if ShowModal = mrOk then
      begin
        Value := Edit.Text;
        Result := True;
      end;
    finally
      Form.Free;
    end;
end;

function Win64 : Boolean;
var Reg : TRegistry;
    sLeitura : String;
begin
   //HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CENTRALPROCESSOR\0\
   Reg := TRegistry.Create;
   Reg.RootKey:=HKEY_LOCAL_MACHINE;
   Reg.OpenKeyReadOnly('\HARDWARE\DESCRIPTION\SYSTEM\CENTRALPROCESSOR\0');
   if Reg.ValueExists('Identifier') then sLeitura := Reg.ReadString('Identifier');
   if Pos('X86',sLeitura) > 0 then Result := False else Result := True;
   FreeAndNil(Reg);
end;

function CustomStrToDate(ADate: String): TDate;
var
   vFormatSettings: TFormatSettings;
begin
   vFormatSettings.ShortDateFormat := 'YYYY-MM-DD';
   vFormatSettings.DateSeparator := '-';
   Result := StrToDate(ADate, vFormatSettings);
end;

function VersaoWin: string;
//identifica a versao do windows
var
   InfoVersao: TOSVersionInfo;
begin
   InfoVersao.dwOSVersionInfoSize:=SizeOf(InfoVersao);
   GetVersionEx(InfoVersao);
   Result:='';
   with InfoVersao do begin
      case dwPlatformId of
         1:
            case dwMinorVersion of
               0: Result:='95';
               10: Result:='98';
               90: Result:='Me';

            end;
         2:
            case dwMajorVersion of
               3: Result:='NT';
               4: Result:='NT';
               5:
                  case dwMinorVersion of
                     0: Result:='2000';
                     1: Result:='XP';
                  end;
               6: Result:='V7';
            end;
      end;
   end;
   if (Result='') then
      Result:='Não Identificado';
end;

function VersaoArquivo(const NomeArq: string) : String;
//fonte: http://www.devmedia.com.br/forum/carregar-versao-no-caption-do-form/397661
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
  Maior, Menor, Release, Build: Word;
begin
   VerInfoSize := GetFileVersionInfoSize( PChar(NomeArq), Dummy );
   GetMem( VerInfo, VerInfoSize );
   try
      GetFileVersionInfo( PChar(NomeArq), 0, VerInfoSize, VerInfo );
      VerQueryValue( VerInfo, '', Pointer(VerValue), VerValueSize );
      with VerValue^ do begin
         Maior := dwFileVersionMS shr 16;
         Menor := dwFileVersionMS and $FFFF;
         Release := dwFileVersionLS shr 16;
         Build := dwFileVersionLS and $FFFF;
      end;
   finally
      FreeMem( VerInfo, VerInfoSize );
   end;
   Result := IntToStr(Maior) + '.' + IntToStr(Menor) + '.' +
      IntToStr(Release) + '.' + IntToStr(Build);
end;

procedure VersaoArquivo(const NomeArq: string;
  var iVersao: array of integer);
//fonte: http://www.devmedia.com.br/forum/carregar-versao-no-caption-do-form/397661
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
  Maior, Menor, Release, Build: Word;
begin
//   System.SetLength(iVersao,4);
   iVersao[0] := -1; iVersao[1] := -1; iVersao[2] := -1; iVersao[3] := -1;
   VerInfoSize := GetFileVersionInfoSize( PChar(NomeArq), Dummy );
   GetMem( VerInfo, VerInfoSize );
   try
      GetFileVersionInfo( PChar(NomeArq), 0, VerInfoSize, VerInfo );
      VerQueryValue( VerInfo, '', Pointer(VerValue), VerValueSize );
      with VerValue^ do begin
         Maior := dwFileVersionMS shr 16;
         Menor := dwFileVersionMS and $FFFF;
         Release := dwFileVersionLS shr 16;
         Build := dwFileVersionLS and $FFFF;
      end;
   finally
      FreeMem( VerInfo, VerInfoSize );
   end;
   iVersao[0] := Maior; iVersao[1] := Menor; iVersao[2] := Release; iVersao[3] := Build;
end;

procedure VersaoArquivo(const NomeArq: string; var sVersao: String; bBuild : Boolean);
var iVersao : Array of integer;
begin
   SetLength(iVersao,4);
   VersaoArquivo(NomeArq,iVersao);
   sVersao := IntToStr(iVersao[0]) + '.' + IntToStr(iVersao[1]) + '.' +
      IntToStr(iVersao[2]);
   if bBuild then sVersao := sVersao+ '.' + IntToStr(iVersao[3]);
end;

procedure DeletaDiretorio(const Dir: string);
// Apaga diretorio com arquivos
var
   FileOp: TSHFileOpStruct;
begin
   FillChar(FileOp, SizeOf(FileOp), 0);
   FileOp.wFunc := FO_DELETE;
   FileOp.pFrom := PChar(Dir+#0);
   FileOp.fFlags := FOF_SILENT or FOF_NOERRORUI or FOF_NOCONFIRMATION;
   SHFileOperation(FileOp);
end;

procedure DeletaDiretorioRecursivo(const Diretorio : string);
//fonte : http://stackoverflow.com/questions/11798783/delete-all-files-and-folders-recursively-using-delphi
var F: TSearchRec;
begin
  if FindFirst(Diretorio + '\*', faAnyFile, F) = 0 then begin
    try
      repeat
        if (F.Attr and faDirectory <> 0) then begin
          if (F.Name <> '.') and (F.Name <> '..') then begin
            DeletaDiretorioRecursivo(Diretorio + '\' + F.Name);
          end;
        end else begin
          DeleteFile(Diretorio + '\' + F.Name);
        end;
      until FindNext(F) <> 0;
    finally
      FindClose(F);
    end;
    RemoveDir(Diretorio);
  end;
end;

function TruncTo(Valor: Double; CasasDecimais: Integer): Double;
begin
   Result := Trunc(Valor * Power(10,CasasDecimais)) / Power(10,CasasDecimais);
end;

function ArredondarParaCima(Valor : Real; CasasDec : Integer): Real;
var
   Resto : Real;
begin
   Resto := 0;
   Result := 0;
   Resto := Valor - RoundTo(Valor, CasasDec);
   Result := RoundTo(Valor, CasasDec);
   if Resto > 0 then begin
      Resto  := Resto;
      Result := Result + Power(10, -1);
   end;
end;

function MinParaHora(Minuto: integer): string;
var
   hr, min : Integer;
begin
   hr := 0;
   while minuto >= 60 do begin
      minuto := minuto - 60;
      hr := hr + 1;
   end;
   min := minuto;
   Result := FormatFloat('00:', hr) + FormatFloat('00', min);
end;

function FormataTipi(sValue:String) : String;
begin
   sValue := Sonumero(sValue);
   sValue := copy(sValue,1,4) + '.' + copy(sValue,5,2) + '.' + copy(sValue,7,2);
   Result := sValue;
end;

function ValiInteiro(Valor : String): Boolean;
var
   Aux : Integer;
begin
   Result := True;
   try
      Aux := StrToInt(Valor);
   except
      Result := False;
   end;
end;

function ValiFlutuante(Valor : Variant): Boolean;
var
   eTemp : Extended;
   iErro : Integer;
begin
   Val(Valor, eTemp, iErro);
   Result := (iErro = 0);
end;

function ConvertBitmapToGrayscale(const Bitmap: TBitmap): TBitmap;
var
   i, j: Integer;
   Grayshade, Red, Green, Blue: Byte;
   PixelColor: Longint;
begin
   with Bitmap do
      for i := 0 to Width - 1 do
         for j := 0 to Height - 1 do begin
            PixelColor := ColorToRGB(Canvas.Pixels[i, j]);
            Red        := PixelColor;
            Green      := PixelColor shr 8;
            Blue       := PixelColor shr 16;
            Grayshade  := Round(0.3 * Red + 0.6 * Green + 0.1 * Blue);
            Canvas.Pixels[i, j] := RGB(Grayshade, Grayshade, Grayshade);
         end;
   Result := Bitmap;
end;

procedure ClonarComboBoxEx(Origem, Destino : TComboBoxEx);
var
   i : Integer;
begin
   Destino.Clear;
   for i := 0 to Origem.ItemsEx.Count -1 do
       Destino.ItemsEx.AddItem(Origem.ItemsEx[i].Caption,Origem.ItemsEx[i].ImageIndex,-1,-1,-1,Nil);
   Destino.ItemIndex := Origem.ItemIndex;
end;

procedure AjustaVisualDBGrid(DataSource : TDataSource; Propriedades : String);
var
   i, x  : Integer;
   Conf, Prop : String;
   //Exibicao = Tamanho|Alinhamento|Máscara  § Item 2... § Item 3
   //Exemplo: 20|C|0.00§5|D|R$ #,##0.00§12,D
begin
   x := ContaSeq(Propriedades,'§');
   if x > 0 then begin
      for i := 0 to x-1 do begin
         Conf := PegaSeq(Propriedades, i+1, '§');
         Prop := PegaSeq(Conf,1,'|');
         if Prop <> '' then if ValiInteiro(Prop)
            then DataSource.DataSet.Fields[i].DisplayWidth := StrToInt(Prop);
         Prop := PegaSeq(Conf,2,'|');
         if  Pos(Prop,'CDE') > 0 then begin
            DataSource.DataSet.Fields[i].Alignment := Alinhamento(Prop);
         end;
         Prop := PegaSeq(Conf,3,'|');
         if Prop <> '' then TFloatField(DataSource.DataSet.Fields[i]).DisplayFormat := Prop;
      end;
   end;
end;

function Alinhamento(Sigla : String) : TAlignment;
begin
   if Sigla = 'C' then Result := taCenter else
   if Sigla = 'D' then Result := taRightJustify else
   if Sigla = 'E' then Result := taLeftJustify;
end;

function FormatMinToHour(Min: LongInt): String;
var
   Hrs : Word;
begin
   Hrs := Min div 60;
   Min := Hrs mod 60;
   Result := Format('%d:%d', [Hrs, Min]);
end;

//=======================================================//
// Formata horas decimais em formato de exibição de hora //
// Criada por Alexssandro de Souza Marcelino             //
// Exemplo: 50,95 = 50:57 horas                          //
//=======================================================//
function FormatDecHourToHour(Hours : Extended) : String;
var
   Hrs : Integer;
   Min, t : Double;
begin
   Hours := RoundTo(Hours,-2);//Incluído em 26.01.15 por Iago César. Problemas com tipos de dados.
   Hrs := Trunc(Hours);
   Min := Hours - Hrs;
   Min := Round(Min * 60);
   if Hours = 0 then Result := '-'
   else Result := IntToStr(Hrs) + ':' + StrZero(FloatToStr(Min),2);
end;

function DistinctSeq(Lista: String; Sep: Char = #44): String;
var
   sl : TStringList;
   i : Integer;
begin
   Result := '';
   sl := TStringList.Create;
   with sl do begin
      Sorted     := True;
      Duplicates := dupIgnore;
      for i := 0 to ContaSeq(Lista,Sep) -1 do sl.Add(PegaSeq(Lista,i+1,Sep));
   end;
   for i := 0 to sl.Count -1 do Result := Result + sl.Strings[i] + Sep;
   Result := ValidaSequencia(Result);
   sl.Free;
end;


function IfThenString(AValue: Boolean; const ATrue: string; const AFalse: string): string;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function FormataDataRelatorio(dtIni, dtFin : TDate): string;
var
   sAux : string;
   wAux : Word;
begin
   sAux := '';
   if dtIni = dtFin then
   begin
      //verificando se o filtro se trata de apenas um dia
      sAux := '0|';
      wAux := DayOf(dtIni);
      sAux := sAux + IntToStr(wAux);
      sAux := sAux + ' de ' + Trim(DizMes(MonthOf(dtIni))) + ' de ';
      wAux := Yearof(dtIni); sAux := sAux + IntToStr(wAux) + '|';
   end else
   if (MonthOf(dtIni) = MonthOf(dtFin)) and (YearOf(dtIni) = YearOf(dtFin))
       and ((DayOf(dtIni) = 1) and (DateToStr(dtFin) = DateToStr(FloatToDateTime(EndOfTheMonth(dtFin))))) then
   begin
      //verificando se o filtro se trata de apenas um mês (completo)
      sAux := '1|';
      sAux := sAux + Trim(DizMes(MonthOf(dtIni))) + '/';
      wAux := YearOf(dtIni); sAux := sAux + IntToStr(wAux) + '|';
   end else
   begin
      sAux := '2|';
      sAux := sAux + DateToStr(dtIni) + '|' + DateToStr(dtFin);
   end;
   result := sAux;
end;

function EnderecoMAC : string;
var
   Lib: Cardinal;
   Func: function(GUID: PGUID): Longint; stdcall;
   GUID1, GUID2: TGUID;
begin
   Result := '';
   Lib := LoadLibrary('rpcrt4.dll');
   if Lib <> 0 then
   begin
      @Func := GetProcAddress(Lib, 'UuidCreateSequential');
      if Assigned(Func) then
      begin
         if (Func(@GUID1) = 0) and
            (Func(@GUID2) = 0) and
            (GUID1.D4[2] = GUID2.D4[2]) and
            (GUID1.D4[3] = GUID2.D4[3]) and
            (GUID1.D4[4] = GUID2.D4[4]) and
            (GUID1.D4[5] = GUID2.D4[5]) and
            (GUID1.D4[6] = GUID2.D4[6]) and
            (GUID1.D4[7] = GUID2.D4[7]) then
         begin
            Result :=
            IntToHex(GUID1.D4[2], 2) + '-' +
            IntToHex(GUID1.D4[3], 2) + '-' +
            IntToHex(GUID1.D4[4], 2) + '-' +
            IntToHex(GUID1.D4[5], 2) + '-' +
            IntToHex(GUID1.D4[6], 2) + '-' +
            IntToHex(GUID1.D4[7], 2);
         end;
      end;
   end;
end;

{ TMeuBallonHint }

destructor TMeuBallonHint.HideBallon(Window: HWnd);
var Ballon : TEditBalloonTip;
begin
   SendMessageW(Window, Ballon.EM_HIDEBALLOONTIP, 0, 0);
end;

constructor TMeuBallonHint.ShowBallon(Window: HWnd; Texto, Titulo: PWideChar;
  Tipo: TTipoBallon);
var Ballon : TEditBalloonTip;
begin
   Ballon.cbStruct := SizeOf(TEditBalloonTip);
   Ballon.pszText := Texto;
   Ballon.pszTitle := Titulo;
   Ballon.ttiIcon := Integer(Tipo);
   SendMessageW(Window, Ballon.EM_SHOWBALLOONTIP, 0,Integer(@Ballon));
end;

procedure RetornaRosca(Percentual: Double; CorPerc, CorResto, CorFundo: TColor; Tamanho: Smallint; Destino: TImage; Espessura: TEspessuraRosca; Negrito : Boolean = False; CorNegat: TColor = clMaroon; ExibePerc: Boolean = True);
var
   Center: TPoint;
   Bitmap: TBitmap;
   BitRed: TBitmap;
   Radius: Integer;
   Esp : Real;
   PercTexto : Double;

   procedure DrawPieSlice(const Canvas: TCanvas; const Center: TPoint;
   const Radius: Integer; const StartDegrees, StopDegrees: Double);
   const
     Offset = 0;
   var
     X1, X2, X3, X4: Integer;
     Y1, Y2, Y3, Y4: Integer;
   begin
     X1 := Center.X - Radius;
     Y1 := Center.Y - Radius;
     X2 := Center.X + Radius;
     Y2 := Center.Y + Radius;
     X3 := Center.X + Round(Radius * Cos(DegToRad(Offset + StartDegrees)));
     Y3 := Center.y - Round(Radius * Sin(DegToRad(Offset + StartDegrees)));
     X4 := Center.X + Round(Radius * Cos(DegToRad(Offset + StopDegrees)));
     Y4 := Center.y - Round(Radius * Sin(DegToRad(Offset + StopDegrees)));
     Canvas.Pie(X1, Y1, X2, Y2, X3, Y3, X4, Y4);
   end;

begin
   PercTexto := Percentual;
   if Percentual < 0 then begin
      Percentual := Percentual * -1;
      CorPerc := CorNegat;
   end;
   if Percentual > 100 then Percentual := 100;
   Bitmap := TBitmap.Create;
   BitRed := TBitmap.Create;
   if Espessura = espMedia  then Esp := 2.5 else
   if Espessura = espFina   then Esp := 3.5 else
   if Espessura = espGrossa then Esp := 2.0;
   try
      Destino.Picture.Bitmap := Nil;
      Tamanho := Tamanho * 2;
      Bitmap.Width  := Tamanho;
      Bitmap.Height := Tamanho;
      Bitmap.PixelFormat := pf24bit;
      Bitmap.Canvas.Brush.Color := CorPerc;
      Bitmap.Canvas.Pen.Color := CorPerc;
      Center := Point(Bitmap.Width div 2, Bitmap.Height div 2);
      Radius := Bitmap.Width div 2;
      DrawPieSlice (Bitmap.Canvas, Center, Radius,  0, Trunc((Percentual * 360) / 100));
      Bitmap.Canvas.Brush.Color := CorResto;
      Bitmap.Canvas.Pen.Color := CorResto;
      if Percentual < 100 then DrawPieSlice (Bitmap.Canvas, Center, Radius, Trunc((Percentual * 360) / 100), 360);
      Bitmap.Canvas.Brush.Color := CorFundo;
      Bitmap.Canvas.Pen.Color := CorFundo;
      Bitmap.Canvas.Ellipse(
         Trunc((Tamanho / Esp) / Esp),
         Trunc((Tamanho / Esp) / Esp),
         Tamanho - Trunc((Tamanho / Esp) / Esp),
         Tamanho - Trunc((Tamanho / Esp) / Esp));
      Bitmap.TransparentMode := tmAuto;
      Bitmap.Transparent := True;
      Bitmap.TransparentColor := CorFundo;
      Bitmap.Canvas.Font.Size := Trunc(Tamanho / 8);
      SetTextAlign(Bitmap.Canvas.Handle,TA_CENTER);
      if Negrito then Bitmap.Canvas.Font.Style := [fsBold]
      else Bitmap.Canvas.Font.Style := [];
      if ExibePerc then begin
         Bitmap.Canvas.TextOut(
            Tamanho div 2,
            (Tamanho div 2) - (Bitmap.Canvas.TextHeight('100%') div 2),
            FormatFloat('0%',PercTexto));
      end;
      QualityResizeBitmap(Bitmap,BitRed,Tamanho div 2,Tamanho div 2);
      Destino.Picture.Bitmap := BitRed;
   finally
      FreeAndNil(Bitmap);
      FreeAndNil(BitRed);
   end;
end;


procedure QualityResizeBitmap(bmpOrig, bmpDest: TBitmap; newWidth,
  newHeight: Integer);
var
   xIni, xFin, yIni, yFin, xSalt, ySalt: Double;
   X, Y, pX, pY, tpX: Integer;
   R, G, B: Integer;
   pxColor: TColor;
begin
   bmpDest.Width  := newWidth;
   bmpDest.Height := newHeight;
   xSalt := bmpOrig.Width / newWidth;
   ySalt := bmpOrig.Height / newHeight;
   yFin := 0;
   for Y := 0 to pred(newHeight) do begin
      yIni := yFin;
      yFin := yIni + ySalt;
      if yFin >= bmpOrig.Height then yFin := pred(bmpOrig.Height);
      xFin := 0;
      for X := 0 to pred(newWidth) do begin
         xIni := xFin;
         xFin := xIni + xSalt;
         if xFin >= bmpOrig.Width then xFin := pred(bmpOrig.Width);
         R   := 0;
         G   := 0;
         B   := 0;
         tpX := 0;
         for pY := Round(yIni) to Round(yFin) do
            for pX := Round(xIni) to Round(xFin) do begin
               Inc(tpX);
               pxColor := ColorToRGB(bmpOrig.Canvas.Pixels[pX, pY]);
               R := R + GetRValue(pxColor);
               G := G + GetGValue(pxColor);
               B := B + GetBValue(pxColor);
            end;
         bmpDest.Canvas.Pixels[X,Y] := RGB(Round(R/tpX),Round(G/tpX),Round(B/tpX));
      end;
   end;
end;

function base64Encode(Texto : AnsiString):AnsiString;
var
   Encoder : TIdEncoderMime;
begin
   Encoder := TIdEncoderMime.Create(Nil);
   try
      Result := Encoder.EncodeString(Texto);
   finally
      FreeAndNil(Encoder);
   end;
end;

function base64Decode(Texto : AnsiString):AnsiString;
var
   Decoder : TIdDecoderMime;
begin
   Decoder := TIdDecoderMime.Create(nil);
   try
      Result := Decoder.DecodeString(Texto);
   finally
      FreeAndNil(Decoder)
   end
end;

function ConverteEncodingXML(var sPathArquivoXML : String; Salvar : Boolean; PaiTemp : TWinControl) : Boolean;
var
   WB      : TWebBrowser;
   WbDoc   : IHTMLDocument2 ;
   mXML    : TStrings;
   Visivel : Boolean;
   sAux    : String;
   I       : Integer;
begin
   try
      try
         result := True;
         Visivel := TWinControl(PaiTemp).Visible;
         WB := TWebBrowser.Create(Nil);
         WB.RegisterAsBrowser:= True;
         TWinControl(WB).Name   := 'wbLeitorXML';
         TWinControl(WB).Parent := PaiTemp;
         WB.Silent := true; WB.Visible:= False; //TWinControl(PaiTemp).Visible := False;
         WB.Top := 1; WB.Left := 1; WB.Height := 6; WB.Width := 8;

         mXML := TStringList.Create;
         WB.Visible := False;
         sAux := 'file://\'+sPathArquivoXML;
         //sAux := TIdURI.PathEncode(sAux);
         {$IFDEF DEBUG}
            Clipboard.AsText := sAux;
         {$ENDIF }
         WB.Navigate(sAux);
         while WB.ReadyState <> READYSTATE_COMPLETE do Application.ProcessMessages;

         WbDoc := WB.Document as IHTMLDocument2;
         while WbDoc.readyState <> 'complete' do Application.ProcessMessages;
         mXML.Text := WbDoc.body.innerText;
         mXML[0] := StringReplace(mXML[0],'" ?>','"?>',[rfReplaceAll]);
         mXML[0] := StringReplace(mXML[0],'  ','',[rfReplaceAll]);

         for I := 0 to mXML.Count-1 do begin
            if Copy(mXML[i],1,2) = '- ' then
               mXML[i] := Copy(mXML[i],2,Length(mXML[i]));
            mXML[i] := Trim(mXML[i]);
         end;

         if Salvar then mXML.SaveToFile(sPathArquivoXML,TEncoding.UTF8) else sPathArquivoXML := mXml.Text;
      except
         On e : Exception do begin
            Result := False;
            {$IFDEF DEBUG}
               Clipboard.AsText := e.Message;
            {$ENDIF }
         end;
      end;
   finally
      FreeAndNil(WB);
      FreeAndNil(mXML);
      TWinControl(PaiTemp).Visible := Visivel;
   end;
end;

function Hex2Dec(texto : string) : string;
var sHex : string; iDec : integer;
begin
   sHex := copy(texto,Pos('%',texto)+1,2);
//      if UpperCase(sHex) = 'NUL'   then iDec := 0  else if UpperCase(sHex) = 'SOH' then iDec := 1 else
//      if UpperCase(sHex) = 'STX'   then iDec := 2  else if UpperCase(sHex) = 'ETX' then iDec := 3 else
//      if UpperCase(sHex) = 'EOT'   then iDec := 4  else if UpperCase(sHex) = 'ENQ' then iDec := 5 else
//      if UpperCase(sHex) = 'ACK'   then iDec := 6  else if UpperCase(sHex) = 'BEL' then iDec := 7 else
//      if UpperCase(sHex) = 'BS'    then iDec := 8  else if UpperCase(sHex) = 'TAB' then iDec := 9 else
//      if UpperCase(sHex) = 'LF'    then iDec := 10 else if UpperCase(sHex) = 'VT'  then iDec := 11 else
//      if UpperCase(sHex) = 'FF'    then iDec := 12 else if UpperCase(sHex) = 'CR'  then iDec := 13 else
//      if UpperCase(sHex) = 'SO'    then iDec := 14 else if UpperCase(sHex) = 'SI'  then iDec := 15 else
//      if UpperCase(sHex) = 'DLE'   then iDec := 16 else if UpperCase(sHex) = 'DC1' then iDec := 17 else
//      if UpperCase(sHex) = 'DC2'   then iDec := 18 else if UpperCase(sHex) = 'DC3' then iDec := 19 else
//      if UpperCase(sHex) = 'DC4'   then iDec := 20 else if UpperCase(sHex) = 'NAK' then iDec := 21 else
//      if UpperCase(sHex) = 'SYN'   then iDec := 22 else if UpperCase(sHex) = 'ETB' then iDec := 23 else
//      if UpperCase(sHex) = 'CAN'   then iDec := 24 else if UpperCase(sHex) = 'EM'  then iDec := 25 else
//      if UpperCase(sHex) = 'SUB'   then iDec := 26 else if UpperCase(sHex) = 'ESC' then iDec := 27 else
//      if UpperCase(sHex) = 'FS'    then iDec := 28 else if UpperCase(sHex) = 'GS'  then iDec := 29 else
//      if UpperCase(sHex) = 'RS'    then iDec := 30 else if UpperCase(sHex) = 'US'  then iDec := 31 else
//      if UpperCase(sHex) = 'SPACE' then iDec := 32 else iDec := StrToInt('$'+sHex);
   try
      iDec := StrToInt('$'+sHex);
      texto := StringReplace(texto,'%'+sHex, Chr(iDec), [rfReplaceAll]);
   except
      texto := StringReplace(texto,'%'+sHex, '', [rfReplaceAll]);
   end;

   texto := StringReplace(texto,'%'+sHex, Chr(iDec), [rfReplaceAll]);
   if Pos('%',texto) > 0 then texto := Hex2Dec(texto);
   Result := texto;
end;

function GerarStringRandom(Size : Integer; Tipo : Integer = 1) : String;
//fonte: http://showdelphi.com.br/dica-funcao-para-gerar-uma-senha-aleatoria-delphi/
var I: Integer; Chave: String;
const
   str1 = '1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
   str2 = '1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ';
   str3 = '1234567890abcdefghijklmnopqrstuvwxyz';
   str4 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
   str5 = '1234567890';
   str6 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
   str7 = 'abcdefghijklmnopqrstuvwxyz';
begin
   Chave := '';
   for I := 1 to Size do begin
      case Tipo of
         1 : Chave := Chave + str1[Random(Length(str1)) + 1];
         2 : Chave := Chave + str2[Random(Length(str2)) + 1];
         3 : Chave := Chave + str3[Random(Length(str3)) + 1];
         4 : Chave := Chave + str4[Random(Length(str4)) + 1];
         5 : Chave := Chave + str5[Random(Length(str5)) + 1];
         6 : Chave := Chave + str6[Random(Length(str6)) + 1];
         7 : Chave := Chave + str7[Random(Length(str7)) + 1];
      end;
   end;
   Result := Chave;
end;

function PosStringInArray(Texto : String; Vetor : Array of String) : Integer;
var i : integer;
begin
   Result := -1;
   for I := 0 to Length(Vetor)-1 do begin
      if Vetor[i] = Texto then begin
         Result := i;
         exit;
      end;
   end;
end;

function RecortarImagem(Imagem : TImage; NewWidth, NewHeigth : Integer) : TImage;
var
   DstRect, SrcRect: TRect;
   H,W : Integer;
begin
   W := Imagem.Picture.Width;
   H := Imagem.Picture.Height;
   W := Trunc((W - NewWidth)/2);
   H := Trunc((H - NewHeigth)/2);

   DstRect := Rect(0, 0,     NewWidth,     NewHeigth);
   SrcRect := Rect(W, H, W + NewWidth, H + NewHeigth);

   Imagem.Canvas.CopyMode := cmSrcAnd;
   Result := TImage.Create(nil);
   Result.Width := NewWidth;  Result.Height := NewHeigth;
   Result.Canvas.CopyRect(DstRect, Imagem.Canvas, SrcRect);
end;

procedure NaoEscondeJvCaptionPanel(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var pFilho: TJvCaptionPanel; pPai  :   TWinControl;
begin
   pFilho := TJvCaptionPanel(Sender);
   pPai   := TWinControl(TWinControl(Sender).Parent);
   if Button = mbLeft then begin
      if pFilho.Left + pFilho.Width > pPai.Width  then pFilho.Left := pPai.Width - pFilho.Width;
      if pFilho.Left < 0  then pFilho.Left := 0;
      if pFilho.Top < 0 then pFilho.Top := 0;
      if pFilho.Top > pPai.Height - pFilho.Height  then pFilho.Top := pPai.Height - pFilho.Height;
   end;
end;

procedure EscondeSheets(pcGeral : TPageControl);
var i: Integer;
begin
   pcGeral.ActivePage.TabVisible := True;
   for i := 0 to pcGeral.PageCount - 1 do begin
      pcGeral.Pages[i].TabVisible := pcGeral.ActivePageIndex = i;
   end;
end;

procedure CentralizarJvCaptionPanel(Panel : TJvCaptionPanel; Referencia : TObject;
   mkpH : Integer = 4; mkpW : Integer = 2);
var W,H : Integer;
begin
   W := TWinControl(Referencia).Width;
   H := TWinControl(Referencia).Height;
   TJvCaptionPanel(Panel).Left := (W - TJvCaptionPanel(Panel).Width ) div mkpW;
   TJvCaptionPanel(Panel).Top  := (H - TJvCaptionPanel(Panel).Height) div mkpH;
end;

procedure WBLoadHTML(WebBrowser: TWebBrowser; HTMLCode: string);
var
   sl: TStringList;
   ms: TMemoryStream;
begin
   if (WebBrowser = nil) then Exit;
   WebBrowser.Navigate('about:blank');
   if HTMLCode <> '' then
   while WebBrowser.ReadyState < READYSTATE_INTERACTIVE do Application.ProcessMessages;
   if Assigned(WebBrowser.Document) then begin
      sl := TStringList.Create;
      try
         ms := TMemoryStream.Create;
         try
            sl.Text := HTMLCode;
            sl.SaveToStream(ms) ;
            ms.Seek(0,0) ;
            (WebBrowser.Document as IPersistStreamInit).Load(TStreamAdapter.Create(ms));
         finally
            ms.Free;
         end;
      finally
         sl.Free;
      end;
   end;
end;

procedure RetornaDifHorasMaior24(DataHoraA, DataHoraB : TDateTime; var Horas: integer; var Minutos: integer; var Segundos: integer; var Milissegundos: integer);
var Mil : Integer;
begin
   Mil            := MilliSecondsBetween(DataHoraA,DataHoraB);
   Horas          := Mil div 3600000;
   Mil            := Mil - Horas*3600000;
   Minutos        := Mil div 60000;
   Mil            := Mil - Minutos*60000;
   Segundos       := Mil div 1000;
   Mil            := Mil - Segundos*1000;
   Milissegundos  := Mil;
end;

procedure ExportaCsv(cdsExpor: TClientDataSet; sSeparador: String = ';';
   bAutoExecutar: Boolean = False; sNomeArq: string = '');
var
   i: integer;
   sl: TStringList;
   st: string;
begin
   if cdsExpor.IsEmpty then Abort;
   cdsExpor.DisableControls;
   cdsExpor.First;
   sl := TStringList.Create;
   try
      st := '';
      for i := 0 to cdsExpor.Fields.Count - 1 do
         st := st + cdsExpor.Fields[i].DisplayLabel + sSeparador;
      sl.Add(st);
      cdsExpor.First;
      while not cdsExpor.Eof do begin
         st := '';
         for i := 0 to cdsExpor.Fields.Count - 1 do
            st := st + cdsExpor.Fields[i].DisplayText + sSeparador;
         sl.Add(st);
         cdsExpor.Next;
      end;
      if sNomeArq <> '' then sl.SaveToFile(ExtractFilePath(Application.ExeName)+ sNomeArq)
      else sl.SaveToFile(ExtractFilePath(Application.ExeName)+'Export.csv');
   finally
      FreeAndNil(sl);
      cdsExpor.EnableControls;
   end;
end;

function FileTimeToDTime(FTime: TFileTime): TDateTime;
var
  LocalFTime: TFileTime;
  STime: TSystemTime;
begin
  FileTimeToLocalFileTime(FTime, LocalFTime);
  FileTimeToSystemTime(LocalFTime, STime);
  Result := SystemTimeToDateTime(STime);
end;

function RetornaDataArquivo(NomeArquivo, TipoData : String) : TDateTime;
var
  SR: TSearchRec;
  CreateDT, AccessDT, ModifyDT: TDateTime;
begin
   try
      TipoData := UpperCase(TipoData);
      if FindFirst(NomeArquivo, faAnyFile, SR) = 0 then begin
         CreateDT := FileTimeToDTime(SR.FindData.ftCreationTime);
         AccessDT := FileTimeToDTime(SR.FindData.ftLastAccessTime);
         ModifyDT := FileTimeToDTime(SR.FindData.ftLastWriteTime);
         if TipoData = 'C' then Result := CreateDT else
         if TipoData = 'A' then Result := AccessDT else
         if TipoData = 'M' then Result := ModifyDT else
         Result := Now;
      end else raise Exception.Create('Arquivo não encontrado.');
   finally
      FindClose(SR);
   end;
end;

function VerificaEXE(NomeEXE: String) : Boolean;
var
   Processo: TProcessEntry32;
   Hnd: THandle;
   Fnd: Boolean;
   List : TStrings;
begin
   List := TStringList.Create;
   List.Clear;
   Hnd := CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
   if Hnd <> -1 then begin
      Processo.dwSize := SizeOf(TProcessEntry32);
      Fnd := Process32First(Hnd, Processo);
      while Fnd do begin
         List.Add(Processo.szExeFile);
         Fnd := Process32Next(Hnd, Processo);
      end;
      CloseHandle(Hnd);
   end;
   if List.IndexOf(NomeEXE) > -1 then Result := True else Result := False;
   FreeAndNil(List);
end;

procedure LimparMemoriaResidual;
var MainHandle : THandle;
begin
   try
      MainHandle := OpenProcess(PROCESS_ALL_ACCESS, false, GetCurrentProcessID) ;
      SetProcessWorkingSetSize(MainHandle, $FFFFFFFF, $FFFFFFFF) ;
      CloseHandle(MainHandle) ;
   except end;
   Application.ProcessMessages;
end;

procedure PiscaTela(iHandle : Cardinal; iQuantidade : Cardinal = 10; iIntervalo : Cardinal = 500);
var pfwi : FLASHWINFO;
begin
   try
      pfwi.cbSize     := SizeOf(pfwi);
      pfwi.hwnd       := iHandle;
      pfwi.dwFlags    := FLASHW_ALL or FLASHW_TIMER; // or FLASHW_TIMER;//or FLASHW_TIMERNOFG;
      pfwi.uCount     := iQuantidade;
      pfwi.dwTimeout  := iIntervalo;
      FlashWindowEx(pfwi);
   except end;
end;

END.
