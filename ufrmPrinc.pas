unit ufrmPrinc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Vcl.Clipbrd, SynEdit, SynHighlighterSQL,
  SynEditHighlighter, SynHighlighterDWS, Vcl.ExtCtrls, SynEditCodeFolding,
  Vcl.Themes, System.IniFiles,

  uConfiguracoes;

type
  TfrmPrinc = class(TForm)
    SynDWSSyn1: TSynDWSSyn;
    SynSQLSyn1: TSynSQLSyn;
    ScrollBox1: TScrollBox;
    Panel1: TPanel;
    Label3: TLabel;
    mmSQL: TSynEdit;
    Splitter1: TSplitter;
    Panel2: TPanel;
    Label4: TLabel;
    mmPascal: TSynEdit;
    Panel3: TPanel;
    btnClipBoard1: TButton;
    Panel4: TPanel;
    btnClipBoard2: TButton;
    pTop: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    btnAdicionar: TButton;
    btnRemover: TButton;
    edtVariavel: TEdit;
    cmbClasse: TComboBox;
    cmbEstilos: TComboBox;
    Label5: TLabel;
    procedure btnAdicionarClick(Sender: TObject);
    procedure btnRemoverClick(Sender: TObject);
    procedure btnClipBoardClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure cmbEstilosSelect(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    xmlConfig : TConfiguracao;
    procedure IniciarEstilos;
    function VersaoSistema: String;
    function Padr(texto : String;Tam : Integer):String;
    function RetornaPrefixo: string;
    function RetornaSufixo: string;

  public
    { Public declarations }
  end;

var
  frmPrinc: TfrmPrinc;

const
  opTSQLQuery     = 0;
  opTFDQuery      = 1;
  opTStrings      = 2;
  opString        = 3;
  opStringSimples = 4;

implementation


{$R *.dfm}

function TfrmPrinc.RetornaPrefixo: string;
begin
  if cmbClasse.ItemIndex = opTStrings then
    Result := edtVariavel.Text + '.Add('+#39
  else if cmbClasse.ItemIndex = opString then
    Result := edtVariavel.Text + ' := '+edtVariavel.Text+' + '+ #39
  else if cmbClasse.ItemIndex = opStringSimples then
    Result := '+ ' + #39
  else
    result := edtVariavel.Text + '.SQL.Add('+#39;
end;

function TfrmPrinc.RetornaSufixo: string;
begin
  if cmbClasse.ItemIndex = opString then
    Result := #39 +' + #13#10;'
  else if cmbClasse.ItemIndex = opStringSimples then
    Result := ' '+#39
  else
    Result := #39+');';
end;

function TfrmPrinc.VersaoSistema: String;
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
  Maior, Menor, Release, Build: Word;
  NomeArq: String;
begin
  NomeArq := Application.ExeName;
  VerInfoSize := GetFileVersionInfoSize( PChar(NomeArq), Dummy );
  GetMem( VerInfo, VerInfoSize );
  try
    GetFileVersionInfo( PChar(NomeArq), 0, VerInfoSize, VerInfo );
    VerQueryValue( VerInfo, '', Pointer(VerValue), VerValueSize );
    with VerValue^ do begin
      Maior    := dwFileVersionMS shr 16;
      Menor    := dwFileVersionMS and $FFFF;
      Release  := dwFileVersionLS shr 16;
      Build    := dwFileVersionLS and $FFFF;
    end;
  finally
    FreeMem( VerInfo, VerInfoSize );
  end;
  Result := Maior.ToString
    +'.'+ Menor.ToString
    +'.'+ Release.ToString
    //+'.'+ Build.ToString
  ;
end;

procedure TfrmPrinc.btnAdicionarClick(Sender: TObject);
var
  r, i, qtd     : integer;
  s, idn, sTemp : string;
  slParam       : TStrings;

  procedure ExtraiParametro(sAux : string; slAux : TStrings);
  var c,i : integer; tmp : string;
  begin
    c := pos(':',sAux); i := pos(' ',sAux,c);
    tmp := copy (sAux,c+1,i-c-1);
    if tmp <> '' then begin
      if tmp[Length(tmp)] = ',' then Delete(tmp, Length(tmp), 1);
      if tmp[Length(tmp)] = ')' then Delete(tmp, Length(tmp), 1);
      if pos(tmp,slAux.Text) = 0 then slAux.Add(tmp);
    end;
    sAux := copy(sAux,pos(' ',sAux,c),Length(sAux));
    c := pos(':',sAux);
    if c > 0 then ExtraiParametro(sAux,slAux);
  end;
begin
   mmPascal.Clear;
   idn := '';

   if mmSQL.Lines.Count = 0  then
   begin
      Application.MessageBox(
        'Você deve informar o SQL original no memo à esquerda.',
        PWideChar(Application.Title),MB_ICONWARNING + MB_OK
      );
      Abort;
   end else begin
      slParam := TStringList.Create;
      try
        {$REGION 'Inicializando objetos'}
        if cmbClasse.ItemIndex = opTSQLQuery then begin
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := '+cmbClasse.Text+'.Create(Self);');
          mmPascal.Lines.Add(idn + edtVariavel.Text + '.SQLConnection := dm.SQLConnection;');
        end else
        if cmbClasse.ItemIndex = opTFDQuery then begin
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := '+cmbClasse.Text+'.Create(Self);');
          mmPascal.Lines.Add(idn + edtVariavel.Text + '.Connection := dm.FDConnection;');
        end else
        if cmbClasse.ItemIndex = opTStrings then begin
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := TStringList.Create;');
        end else
        if cmbClasse.ItemIndex = opString then begin
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := '+#39#39 + ';');
        end else
        if cmbClasse.ItemIndex = opStringSimples then begin
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := #13#10 ');
        end;
        {$ENDREGION}

        mmPascal.Lines.Add(' ');
        {$REGION 'Transcrevendo o SQL'}
        for r := 0 to mmSQL.Lines.Count-1 do begin
          sTemp := mmSQL.Lines.Strings[r];
          if Trim(sTemp) = '' then
            mmPascal.Lines.Add(' ')
          else begin
            sTemp := StringReplace(sTemp, #39, #39#39, [rfReplaceAll]);

            mmPascal.Lines.Add(
              idn + RetornaPrefixo + sTemp + RetornaSufixo
            );

            if pos(':',sTemp) > 0 then
              ExtraiParametro(sTemp+' ', slParam);
          end;
        end;
        {$ENDREGION}

        if cmbClasse.ItemIndex = opStringSimples then
          mmPascal.Lines.Add(idn + '; ');
        mmPascal.Lines.Add(' ');
        {$REGION 'Listando os parâmetros'}
        for r := 0 to slParam.Count-1 do begin
          if cmbClasse.ItemIndex = opTStrings then
            mmPascal.Lines.Add(edtVariavel.Text+'.Text := StringReplace('+edtVariavel.Text+'.Text,'+#39+':'+slParam.Strings[r]+#39+','+#39+'MinhaVariavel'+#39+', [rfReplaceAll]) ;')
          else if cmbClasse.ItemIndex in [opString, opStringSimples] then
            mmPascal.Lines.Add(edtVariavel.Text+' := StringReplace('+edtVariavel.Text+','+#39+':'+slParam.Strings[r]+#39+','+#39+'MinhaVariavel'+#39+', [rfReplaceAll]) ;')
          else
            mmPascal.Lines.Add(edtVariavel.Text+'.ParamByName('+#39+slParam.Strings[r]+#39+').AsString := '+#39+'MinhaVariavel'+#39+' ;');
        end;
        {$ENDREGION}

        if not (cmbClasse.ItemIndex in [opTStrings,opString,opStringSimples]) then begin
          mmPascal.Lines.Add(' ');
          mmPascal.Lines.Add(idn + edtVariavel.Text +  '.Open; ' );
        end;

      finally
        FreeAndNil(slParam);
        mmPascal.SetFocus;
      end;
   end;
end;

procedure TfrmPrinc.btnRemoverClick(Sender: TObject);
var
  r, i : integer;
  iIni, iFim: Integer;
  sTemp : string;
  sufixo,prefixo: string;

begin
  mmSQL.Clear;
  prefixo := RetornaPrefixo;
  sufixo  := RetornaSufixo;

  for r := 0 to mmPascal.Lines.Count do begin
    sTemp := mmPascal.Lines.Strings[r];

    iIni  := AnsiPos(prefixo, sTemp);
    if iIni = 0 then continue;
    inc(iIni,Length(prefixo));
    iFim  := LastDelimiter(sufixo, sTemp)+1-Length(sufixo);

    sTemp := Copy(sTemp, iIni, iFim-iIni);
    sTemp := StringReplace(sTemp, #39#39,#39,[rfReplaceAll]);

    if not ((mmSQL.Lines[mmSQL.Lines.Count-1] = '') and (sTemp = '')) then
      mmSQL.lines.add(sTemp);
  end;
  mmSQL.SetFocus;
end;

procedure TfrmPrinc.cmbEstilosSelect(Sender: TObject);
begin
  TStyleManager.SetStyle(cmbEstilos.Text);
end;

procedure TfrmPrinc.btnClipBoardClick(Sender: TObject);
var Lines: TStrings;
begin
  if TButton(Sender).Tag = 1 then
    Lines  := mmPascal.Lines
  else
    Lines := mmSQL.Lines;

   if Lines.Count > 0 then
   begin
      Clipboard.Astext := Lines.Text;
      Application.MessageBox(
        'Texto copiado para o clipboard!',
        PWideChar(Application.Title), MB_ICONINFORMATION + MB_OK
      );
   end;
end;

procedure TfrmPrinc.FormCreate(Sender: TObject);
begin
  Self.Caption  := Self.Caption + ' {vs: '+VersaoSistema+'}';
  xmlConfig := TConfiguracao.Create(Self,ChangeFileExt(Application.ExeName,'.xml'));
  IniciarEstilos;
end;

procedure TfrmPrinc.FormDestroy(Sender: TObject);
begin
  xmlConfig.EscreverValor('Core', 'Variavel', edtVariavel.Text);
  xmlConfig.EscreverValor('Core', 'Classe',   cmbClasse.Text);
  xmlConfig.EscreverValor('Core', 'SQL',      mmSql.Lines.Text);
  xmlConfig.EscreverValor('Core', 'Pascal',   mmPascal.Lines.Text);

  xmlConfig.EscreverValor('Estilo','Estilo',  cmbEstilos.Text);

  xmlConfig.Salvar;
  xmlConfig.DisposeOf;
end;

procedure TfrmPrinc.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case Key of
      VK_F8 : btnAdicionar.Click;
      VK_F9 : btnRemover.Click;
      VK_F4 : btnClipBoard1.Click;
      VK_F5 : btnClipBoard2.Click;
   end;
end;

procedure TfrmPrinc.FormShow(Sender: TObject);
begin
  edtVariavel.Text    := xmlConfig.ObterValor('Core', 'Variavel', edtVariavel.Text);
  mmSql.Lines.Text    := xmlConfig.ObterValor('Core', 'SQL',      '');
  mmPascal.Lines.Text := xmlConfig.ObterValor('Core', 'Pascal',   '');

  cmbClasse.ItemIndex := cmbClasse.Items.IndexOf( xmlConfig.ObterValor('Core', 'Classe', cmbClasse.Text));
end;

procedure TfrmPrinc.IniciarEstilos;
var
  v :String;
  vIndice :Integer;
begin
  cmbEstilos.Clear;
  for v in TStyleManager.StyleNames do
    cmbEstilos.Items.Add(v);
  cmbEstilos.Sorted := True;

  if xmlConfig.ObterValor('Estilo', 'Estilo', '') <> '' then begin
    try
      cmbEstilos.ItemIndex := cmbEstilos.Items.IndexOf(
        xmlConfig.ObterValor('Estilo', 'Estilo',TStyleManager.ActiveStyle.Name)
      );
      TStyleManager.SetStyle(cmbEstilos.Text);
    except
      TStyleManager.SetStyle(TStyleManager.ActiveStyle.Name);
      cmbEstilos.ItemIndex := cmbEstilos.Items.IndexOf(TStyleManager.ActiveStyle.Name);
    end;
  end else
    cmbEstilos.ItemIndex := cmbEstilos.Items.IndexOf(TStyleManager.ActiveStyle.Name);
end;

function TfrmPrinc.Padr(texto: String; Tam: Integer): String;
begin
   if length(texto) > Tam then Texto := copy (texto,1,tam);
   while length(texto) < tam do
      Texto := ' ' + texto;
   Result := Texto;
end;

end.
