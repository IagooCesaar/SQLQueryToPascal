unit uConfiguracoes;

interface


uses Windows, Messages, SysUtils, Variants, Classes, Xml.XMLDoc, Xml.XMLIntf;


type TConfiguracao = class
  private
    FXmlPath: String;
    FConfig: TXMLDocument;


  public
    procedure Salvar;
    procedure EscreverValor(NoPai, Chave, Valor: string);
    function ObterValor(NoPai, Chave, ValorDefault: String): String;

    constructor Create(AOwner: TComponent; XmlPath: String);
    destructor Destroy;
end;

implementation

{ TConfiguracao }

destructor TConfiguracao.Destroy;
begin
  FConfig.DisposeOf;
end;

procedure TConfiguracao.EscreverValor(NoPai, Chave, Valor: string);
var nodeRoot, nodeGrupo, nodeRegistro: IXMLNode;
begin
  NoPai     := LowerCase(Nopai);
  Chave     := LowerCase(Chave);

  nodeRoot  := Self.FConfig.ChildNodes.FindNode('root');
  if nodeRoot = nil then
    nodeRoot := Self.FConfig.AddChild('root');

  nodeGrupo := nodeRoot.ChildNodes.FindNode(NoPai);
  if nodeGrupo = nil then
    nodeGrupo := nodeRoot.AddChild(NoPai);

  nodeRegistro := nodeGrupo.ChildNodes.FindNode(Chave);
  if nodeRegistro = nil then
    nodeRegistro  := nodeGrupo.AddChild(Chave);

  nodeRegistro.NodeValue  := Valor;
end;

function TConfiguracao.ObterValor(NoPai, Chave, ValorDefault: String): String;
var nodeRoot, nodeCore, nodeRegistro, nodeChave: IXMLNode;
begin
  Result   := ValorDefault;
  NoPai    := LowerCase(Nopai);
  Chave    := LowerCase(Chave);

  nodeRoot := Self.FConfig.DocumentElement;
  if nodeRoot = nil then Exit;

  nodeCore :=  nodeRoot.ChildNodes.FindNode(NoPai);
  if nodeCore = nil then Exit;

  nodeRegistro  := nodeCore.ChildNodes.FindNode(Chave);
  if nodeRegistro = nil then Exit;

  if nodeRegistro.NodeValue = null then Exit;  
  Result := nodeRegistro.NodeValue;
end;

procedure TConfiguracao.Salvar;
begin
  Self.FConfig.SaveToFile(Self.FXmlPath);
end;

constructor TConfiguracao.Create(AOwner: TComponent; XmlPath: String);
begin
   Self.FXmlPath  := XmlPath;
   Self.FConfig   := TXMLDocument.Create(AOwner);

   if FileExists(Self.FXmlPath) then
    Self.FConfig.LoadFromFile(Self.FXmlPath);

   Self.FConfig.Active  := False;
   Self.FConfig.Options := [doNodeAutoIndent];

   Self.FConfig.Active  := True;
end;

end.
