unit ufrmPrinc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Vcl.Clipbrd, SynEdit, SynHighlighterSQL,
  SynEditHighlighter, SynHighlighterDWS, Vcl.ExtCtrls, SynEditCodeFolding;

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
    procedure btnAdicionarClick(Sender: TObject);
    procedure btnRemoverClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure btnClipBoard1Click(Sender: TObject);
    procedure btnClipBoard2Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    function Padr(texto : String;Tam : Integer):String;
    function RetornaPrefixo: string;
    function RetornaSufixo: string;

  public
    { Public declarations }
  end;

var
  frmPrinc: TfrmPrinc;

const
  opTSQLQuery = 0;
  opTFDQuery  = 1;
  opTStrings  = 2;
  opString    = 3;

implementation


{$R *.dfm}

function TfrmPrinc.RetornaPrefixo: string;
begin
  if cmbClasse.ItemIndex = opTStrings then
    Result := edtVariavel.Text + '.Add('+#39

  else if cmbClasse.ItemIndex = opString then
    Result := edtVariavel.Text + ' := '+edtVariavel.Text+' + '+ #39

  else
    result := edtVariavel.Text + '.SQL.Add('+#39;
end;

function TfrmPrinc.RetornaSufixo: string;
begin
  if cmbClasse.ItemIndex = opString then
    Result := #39 +';'
  else
    Result := #39+');';
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
    if tmp <> '' then
    begin
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
   idn := '';//identação pré definida

   if mmSQL.Lines.Count = 0  then
   begin
      Application.MessageBox(
        'Você deve informar o SQL original no memo à esquerda.',
        'Gerador de String',MB_ICONWARNING + MB_OK
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
          mmPascal.Lines.Add(idn + edtVariavel.Text + ' := '+#39#39 );
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

            //verificando se existe parâmetro na linha atual
            if pos(':',sTemp) > 0 then
              ExtraiParametro(sTemp+' ', slParam);
          end;
        end;
        {$ENDREGION}

        mmPascal.Lines.Add(' ');
        {$REGION 'Listando os parâmetros'}
        for r := 0 to slParam.Count-1 do begin
          if cmbClasse.ItemIndex = opTStrings then
            mmPascal.Lines.Add(edtVariavel.Text+'.Text := StringReplace('+edtVariavel.Text+'.Text,'+#39+':'+slParam.Strings[r]+#39+','+#39+'MinhaVariavel'+#39+', [rfReplaceAll]) ;')
          else if cmbClasse.ItemIndex = opString then
            mmPascal.Lines.Add(edtVariavel.Text+'.Text := StringReplace('+edtVariavel.Text+','+#39+':'+slParam.Strings[r]+#39+','+#39+'MinhaVariavel'+#39+', [rfReplaceAll]) ;')
          else
            mmPascal.Lines.Add(edtVariavel.Text+'.ParamByName('+#39+slParam.Strings[r]+#39+').AsString := '+#39+'MinhaVariavel'+#39+' ;');
        end;
        {$ENDREGION}

        if not (cmbClasse.ItemIndex in [opTStrings,opString]) then begin
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
sTemp : string;

begin
   mmSQL.Clear;
   for r := 0 to mmPascal.Lines.Count do
   begin
      sTemp := mmPascal.Lines.Strings[r];
      i     := AnsiPos('(', sTemp);
      sTemp := copy (sTemp,i + 2, Length(sTemp));
      sTemp := copy (sTemp,1, Length(sTemp) - 3);
      sTemp := StringReplace(sTemp, #39#39,#39,[rfReplaceAll]);

      mmSQL.lines.add (sTemp);
      //mmPascal.Lines.Add(edit1.Text + '.ADD("' + mmSQL.Lines.Strings[r] + '")' )
   end;
   mmSQL.SetFocus;
end;

procedure TfrmPrinc.btnClipBoard1Click(Sender: TObject);
begin
   if mmSQL.Lines.Count > 0 then
   begin
      Clipboard.Astext := mmSQL.Lines.Text;
      Application.MessageBox('Texto copiado para o clipboard!','Gerador de Strings',MB_ICONINFORMATION + MB_OK);
   end;
end;

procedure TfrmPrinc.btnClipBoard2Click(Sender: TObject);
begin
   if mmPascal.Lines.Count > 0 then
   begin
      Clipboard.Astext := mmPascal.Lines.Text;
      Application.MessageBox('Texto copiado para o clipboard!','Gerador de Strings',MB_ICONINFORMATION + MB_OK);
   end;
end;

procedure TfrmPrinc.Button3Click(Sender: TObject);
var
sTemp : String;
sCampo: String;
sTipo : String;
sVar  : String;
sMemo : String;
r, i, iCount : integer;

begin
   mmPascal.Clear;
   iCount :=0;
   for r := 0 to mmSQL.Lines.Count do
   begin
      sTemp  := StringReplace(Trim(mmSQL.Lines.Strings[r]), '"','',[rfReplaceAll, rfIgnoreCase]);
      sCampo := copy(sTemp,1, pos (' ',sTemp)-1);
      sTipo  := copy (sTemp,pos (' ',sTemp) + 1 , (pos ('(',sTemp) - pos (' ',sTemp)) - 1 );
      if sTipo = 'NUMBER' then sTipo := 'AsInteger'
      else if sTipo = 'VARCHAR2' then sTipo := 'AsString'
      else if sTipo = 'VARCHAR' then sTipo := 'AsString'
      else if sTipo = 'FLOAT' then sTipo := 'AsFloat'
      else sTipo:= 'Value';
      sTemp := edtVariavel.Text + '.ParamByName("' + sCampo + '").' + sTipo + '=' ;
      mmPascal.Lines.Add(sTemp);
      if iCount < Length(sTemp) then iCount := Length(sTemp);
   end;
   sMemo := mmSQL.Text;
   mmSQL.Text := mmPascal.Text;
   mmPascal.Clear;

   iCount:= iCount - 1;

   for r := 0 to mmSQL.Lines.Count do
   begin
      sTemp  := copy(Trim(mmSQL.Lines.Strings[r]),1, Length (mmSQL.Lines.Strings[r]) -1) ;
      sTemp  := padr(stemp,icount) + '=';
      mmPascal.Lines.Add(sTemp);
   end;
   mmSQL.Text := sMemo;
end;

procedure TfrmPrinc.Button4Click(Sender: TObject);
var
sCampos : array[0..255] of string;
sCampo  : string;
sTemp   : string;
sValues : string;
r, l : Integer;
begin
   l:=-1;
   for r := 0 to mmSQL.Lines.Count do
   begin
      sTemp  := StringReplace(Trim(mmSQL.Lines.Strings[r]), '"','',[rfReplaceAll, rfIgnoreCase]);
      sCampo := copy(sTemp,1, pos (' ',sTemp)-1);
      sCampos[r] := sCampo;
      if Length (sCampo) > 0 then inc(l);
      mmPascal.Lines.Add(sTemp);
   end;
   sTemp   := 'Insert Into ' +  edtVariavel.Text + ' (';
   sValues := ') Values (';
   for r := 0 to l do
   begin
      if sCampos[r] <> 'Z_GRUPO' then
      begin
         sTemp := sTemp + sCampos[r];
         sValues := sValues + ':' + sCampos[r];
         if r <> l then
         begin
            sTemp := sTemp + ',';
            sValues := sValues + ',';
         end;
      end;
   end;
   sValues := sValues + ')';
   mmPascal.Clear;
   mmPascal.lines.add (stemp + svalues);
end;

procedure TfrmPrinc.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case Key of
      VK_F8 : btnAdicionarClick(Sender);
      VK_F9 : btnRemoverClick(Sender);
      VK_F4 : btnClipBoard1Click(Sender);
      VK_F5 : btnClipBoard2Click(Sender);
   end;
end;

function TfrmPrinc.Padr(texto: String; Tam: Integer): String;
begin
   if length(texto) > Tam then Texto := copy (texto,1,tam);
   while length(texto) < tam do
      Texto := ' ' + texto;
   Result := Texto;
end;

end.
