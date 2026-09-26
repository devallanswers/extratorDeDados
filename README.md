# Extrator de dados de fichas

Script Python que usa o Google Gemini para ler imagens de fichas preenchidas à mão, extrair os campos definidos no código e montar uma planilha Excel. Também tenta recortar uma foto localizada em uma região fixa da imagem.

O formato dos campos e a posição do recorte foram configurados para um modelo específico de ficha. Para usar outro formulário, revise `CAMPOS`, `PROMPT` e as constantes `FOTO_*` em `extrator_de_dados.py`.

## Requisitos

- Python 3.10 ou superior
- Chave de API do Google AI Studio
- Conexão com a internet
- Imagens `.jpg`, `.jpeg`, `.png`, `.gif` ou `.webp`

> O script usa `google-generativeai`, SDK que atualmente emite aviso de descontinuação e não recebe novas atualizações. Para uso continuado, planeje migrar para o SDK `google-genai` e valide a compatibilidade do modelo Gemini escolhido.

## Instalação

Crie e ative um ambiente virtual na pasta do projeto e instale as dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

No macOS/Linux, a ativação é `source .venv/bin/activate`.

## Configuração

Defina a chave em uma variável de ambiente. No PowerShell:

```powershell
$env:GEMINI_API_KEY = "SUA_CHAVE"
```

No macOS/Linux:

```sh
export GEMINI_API_KEY="SUA_CHAVE"
```

Não coloque a chave no código, no README, em arquivos versionados ou em capturas de tela.

## Uso

Coloque as imagens diretamente na pasta `fichas/` e execute:

```powershell
python extrator_de_dados.py
```

Também é possível passar uma imagem ou pasta como argumento:

```powershell
python extrator_de_dados.py "fichas\imagem_001.jpg"
python extrator_de_dados.py "C:\caminho\para\imagens"
```

O script cria `dados_extraidos.xlsx` na pasta atual. Imagens já registradas na planilha são ignoradas em execuções seguintes. O arquivo `erros_processamento.txt` só é criado quando há falhas. Confira manualmente os dados extraídos: OCR de escrita manual pode cometer erros.

## Privacidade e segurança

As imagens são enviadas ao serviço do Google Gemini para processamento. As fichas podem conter dados pessoais, inclusive informações de menores e dados sensíveis de saúde. Antes de usar o script, obtenha as autorizações necessárias, avalie se o envio ao serviço externo é permitido para esses dados e consulte os termos e configurações de privacidade atuais do Google.

Não adicione ao GitHub imagens reais, planilhas geradas, logs, chaves de API ou exemplos que contenham dados pessoais. O `.gitignore` exclui os arquivos de entrada e saída usuais, mas confira sempre `git status` antes de publicar. Armazene os arquivos localmente com acesso restrito e remova-os quando deixarem de ser necessários.