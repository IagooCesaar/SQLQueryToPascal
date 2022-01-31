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
    bClipBoard1: TButton;
    Panel4: TPanel;
    bClipBoard2: TButton;
    pTop: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    bAdicionar: TButton;
    bRemover: TButton;
    edtPrefixo: TEdit;
    cmbClasse: TComboBox;
    procedure bAdicionarClick(Sender: TObject);
    procedure bRemoverClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure bClipBoard1Click(Sender: TObject);
    procedure bClipBoard2Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    function Padr(texto : String;Tam : Integer):String;
  public
    { Public declarations }
  end;

var
  frmPrinc: TfrmPrinc;

implementation


{$R *.dfm}

procedure TfrmPrinc.bAdicionarClick(Sender: TObject);
var
   r, i, qtd : integer;
   s, idn, sTemp  : string;
   slParam : TStrings;

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
      Application.MessageBox('Você deve informar o SQL original no memo à esquerda.','Gerador de String',MB_ICONWARNING + MB_OK);
      Abort;
   end else
   begin
      slParam := TStringList.Create;

      if cmbClasse.ItemIndex = 0 then
      begin
         mmPascal.Lines.Add(idn + edtPrefixo.Text + ' := '+cmbClasse.Text+'.Create(Self);');
         mmPascal.Lines.Add(idn + edtPrefixo.Text + '.SQLConnection := dm.conRemoto;');
      end else
      if cmbClasse.ItemIndex = 1 then begin
         mmPascal.Lines.Add(idn + edtPrefixo.Text + ' := '+cmbClasse.Text+'.Create(Self);');
         mmPascal.Lines.Add(idn + edtPrefixo.Text + '.Connection := dm.conFireDac;');
      end else
      if cmbClasse.ItemIndex = 2 then begin
         mmPascal.Lines.Add(idn + edtPrefixo.Text + ' := TStringList.Create;');
      end else
      if cmbClasse.ItemIndex = 3 then begin
        mmPascal.Lines.Add(idn + edtPrefixo.Text + ' := '' ');
      end;

      mmPascal.Lines.Add(' ');

      for r := 0 to mmSQL.Lines.Count-1 do
      begin
         sTemp := mmSQL.Lines.Strings[r];
         if Trim(sTemp) = '' then mmPascal.Lines.Add(' ') else
         begin
            sTemp := StringReplace(sTemp, #39, #39, [rfReplaceAll]);
            if cmbClasse.ItemIndex = 2 then begin
               mmPascal.Lines.Add(idn + edtPrefixo.Text + '.Add(''' + sTemp + ' '');' );

            end else if cmbClasse.ItemIndex = 3 then begin
               mmPascal.Lines.Add(idn + edtPrefixo.Text + ' := '+edtPrefixo.Text+' + '+ #39 + sTemp + #39 );

            end else begin
               mmPascal.Lines.Add(idn + edtPrefixo.Text + '.SQL.Add(''' + sTemp + ' '');' );
            end;

            //verificando se existe parâmetro na linha atual
            i := pos(':',sTemp);
            if i > 0 then ExtraiParametro(sTemp+' ', slParam);
         end;
      end;

      mmPascal.Lines.Add(' ');

      for r := 0 to slParam.Count-1 do
      begin
         if cmbClasse.ItemIndex = 2 then begin
            mmPascal.Lines.Add(edtPrefixo.Text+'.Text := StringReplace('+edtPrefixo.Text+'.Text,'+#39+':'+slParam.Strings[r]+#39+','+#39+'MinhaVariavel'+#39+', [rfReplaceAll]) ;');
         end else begin
            mmPascal.Lines.Add(edtPrefixo.Text+'.ParamByName('+#39+slParam.Strings[r]+#39+').AsString := '+#39+'MinhaVariavel'+#39+' ;');
         end;
      end;

      if cmbClasse.ItemIndex <> 2 then begin
         mmPascal.Lines.Add(' ');
         mmPascal.Lines.Add(idn + edtPrefixo.Text +  '.Open; ' );
      end;

      FreeAndNil(slParam);
      mmPascal.SetFocus;
   end;
end;

procedure TfrmPrinc.bRemoverClick(Sender: TObject);
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

procedure TfrmPrinc.bClipBoard1Click(Sender: TObject);
begin
   if mmSQL.Lines.Count > 0 then
   begin
      Clipboard.Astext := mmSQL.Lines.Text;
      Application.MessageBox('Texto copiado para o clipboard!','Gerador de Strings',MB_ICONINFORMATION + MB_OK);
   end;
end;

procedure TfrmPrinc.bClipBoard2Click(Sender: TObject);
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
      sTemp := edtPrefixo.Text + '.ParamByName("' + sCampo + '").' + sTipo + '=' ;
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
   sTemp   := 'Insert Into ' +  edtPrefixo.Text + ' (';
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
      VK_F8 : bAdicionarClick(Sender);
      VK_F9 : bRemoverClick(Sender);
      VK_F4 : bClipBoard1Click(Sender);
      VK_F5 : bClipBoard2Click(Sender);
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
