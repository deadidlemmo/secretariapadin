# Sistema da Secretaria Escolar - E.M Jose Padin Mouta

Aplicacao web interna em Flask para apoiar rotinas da secretaria escolar da E.M Jose Padin Mouta. O sistema trabalha principalmente com a Lista Piloto em Excel, fotos de alunos e modelos oficiais de planilhas/documentos para gerar declaracoes, carteirinhas e quadros de acompanhamento.

> Este projeto manipula dados pessoais de alunos. Trate arquivos em `uploads/`, `static/fotos/`, planilhas e documentos gerados como informacao sensivel.

## Funcionalidades

- Login simples por token de acesso.
- Upload da Lista Piloto do Ensino Fundamental e, opcionalmente, da Lista Piloto EJA.
- Geracao de declaracoes escolares:
  - Escolaridade.
  - Transferencia.
  - Conclusao.
  - Frequencia.
  - Declaracoes em lote para alunos do 5o ano.
  - Declaracoes personalizadas de conclusao, matricula cancelada e NCOM.
- Geracao de carteirinhas com foto, dados do aluno e controle de carteirinhas ja impressas.
- Upload individual, multiplo e inline de fotos de alunos.
- Geracao de quadros em Excel a partir da Lista Piloto:
  - Quadro Quantitativo de Inclusao.
  - Quadro de Atendimento Mensal.
  - Informativo Semanal / transferencias.
  - Quadro Quantitativo Mensal de Transferencias Expedidas.
- Conferencia de listas por blueprint em `/confere`, comparando dados da Lista Piloto com PDFs do SED.
- Alertas de prazo calculados a partir de feriados em `modelos/feriados.json`.

## Stack

- Python 3.
- Flask 3.
- Jinja2.
- pandas.
- openpyxl e xlrd para leitura e escrita de Excel.
- pdfplumber para leitura de PDF no modulo de conferencia.
- Bootstrap, Font Awesome e CSS nos templates.
- Gunicorn para deploy.

As dependencias estao em `requirements.txt`.

## Estrutura do Projeto

```text
.
|-- app.py                         # Aplicacao Flask principal
|-- config.py                      # Configuracoes por ambiente e seguranca basica
|-- confere.py                     # Blueprint de conferencia de listas
|-- requirements.txt               # Dependencias Python
|-- Procfile                       # Comando de deploy com Gunicorn
|-- gunicorn.conf.py               # Configuracao alternativa do Gunicorn
|-- services/
|   |-- carteirinhas_log.py        # Log SQLite das carteirinhas impressas
|   |-- fotos.py                   # Busca e salvamento de fotos por RM
|   |-- prazos.py                  # Regras de feriados e alertas de prazo
|   `-- upload_sessions.py         # Salvamento de Excel e persistencia na sessao
|-- utils/                         # Helpers reutilizaveis de Excel e uploads
|   |-- excel.py                   # Escrita segura em celulas mescladas
|   `-- uploads.py                 # Validacao e nomes seguros de uploads
|-- tests/                         # Testes unitarios/estaticos basicos
|-- templates/                     # Telas Jinja2
|-- static/
|   |-- logos/                     # Logos usadas nas telas/declaracoes
|   `-- fotos/                     # Fotos dos alunos, nomeadas por RM
|-- modelos/                       # Modelos oficiais em Excel, DOC, DOCX e PDF
|-- uploads/                       # Arquivos enviados e logs gerados em runtime
`-- README.md
```

## Preparando o Ambiente

No Windows PowerShell, a partir da raiz do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Se o ambiente virtual ja existir, basta ativar e instalar/atualizar as dependencias:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Rodando Localmente

Copie `.env.example` para `.env` e ajuste ao menos:

```text
FLASK_SECRET_KEY=uma-chave-longa-e-segura
ACCESS_TOKEN=token-usado-no-login
MAX_CONTENT_LENGTH_MB=50
SCHOOL_YEAR=2026
CONCLUSAO_5ANO_DATE_TEXT=Praia Grande, 22 de dezembro de 2026
```

```powershell
python app.py
```

Por padrao, o Flask sobe em:

```text
http://127.0.0.1:5000
```

Tambem e possivel usar:

```powershell
flask --app app run --debug
```

## Fluxo de Uso

1. Acesse a tela de login.
2. Informe o token de acesso configurado em `ACCESS_TOKEN` no `.env`.
3. Envie a Lista Piloto do Ensino Fundamental.
4. Envie a Lista Piloto da EJA somente se for usar fluxos que dependem dela.
5. Use o painel principal para acessar Declaracoes, Carteirinhas, Quadros ou Conferencia de Listas.

O app guarda os caminhos das listas carregadas na sessao Flask. Ao trocar de navegador, limpar sessao ou reiniciar com outro ambiente, pode ser necessario reenviar a Lista Piloto.

## Arquivos Esperados

### Lista Piloto Fundamental

Varios fluxos esperam uma planilha Excel com a aba `LISTA CORRIDA`. Dependendo da funcionalidade, sao usadas colunas como:

- `RM`
- `NOME`
- `DATA NASC.`
- `RA`
- `SAI SOZINHO?`
- `SERIE` / `SERIE`
- `HORARIO`
- `OBS`
- `LOCAL TE`
- `TIPO TE`

O Quadro de Atendimento Mensal tambem espera a aba `Total de Alunos`.

### Lista Piloto EJA

A EJA e opcional no fluxo geral. Algumas rotinas mantem suporte a EJA quando ha arquivo carregado e, em certos quadros, quando `ENABLE_EJA` esta ativo.

### Fotos dos Alunos

As fotos ficam em `static/fotos/` e sao associadas pelo RM do aluno. O sistema aceita extensoes como `.jpg`, `.jpeg`, `.png`, `.bmp` e `.gif`.

### Modelos

Os modelos ficam em `modelos/` e sao usados como base para os arquivos gerados:

- `Quadro Quantitativo Mensal - Modelo.xlsx`
- `Quadro Quantitativo de Inclusao - Modelo.xlsx`
- `Quadro Informativo - Modelo.xlsx`
- `Quadro de Atendimento Mensal - Modelo.xlsx`
- `Quadro de Alunos com Deficiencia - Modelo.xlsx`
- Modelos de matricula, ata, prontuario e estagio em subpastas.

Evite renomear ou mover esses arquivos sem atualizar os caminhos em `app.py`.

## Rotas Principais

| Rota | Descricao |
| --- | --- |
| `/login` | Login por token |
| `/logout` | Encerra a sessao |
| `/upload_listas` | Upload das listas piloto |
| `/` | Painel principal |
| `/declaracao/tipo` | Tela principal de declaracoes |
| `/declaracao/conclusao_5ano` | Declaracoes de conclusao em lote para 5o ano |
| `/declaracao/escolaridade_5ano` | Declaracoes de escolaridade em lote para 5o ano |
| `/carteirinhas` | Geracao de carteirinhas |
| `/carteirinhas/marcar_impressas` | Marca carteirinhas como impressas |
| `/upload_foto` | Upload individual de foto |
| `/upload_multiplas_fotos` | Upload em lote de fotos |
| `/upload_inline_foto` | Upload de foto durante o fluxo de carteirinha |
| `/quadros` | Menu de quadros |
| `/quantinclusao` | Quadro Quantitativo de Inclusao |
| `/quadros/atendimento_mensal` | Quadro de Atendimento Mensal |
| `/quadros/transferencias` | Informativo Semanal / transferencias |
| `/quadros/quantitativo_mensal` | Quadro Quantitativo Mensal de Transferencias |
| `/confere/` | Conferencia de listas |
| `/confere/upload_excel` | Upload da Lista Piloto para conferencia |
| `/escolas/search` | Busca de escolas para Select2 |

## Configuracoes Importantes

As principais configuracoes por variavel de ambiente sao:

- `FLASK_SECRET_KEY`: chave de sessao do Flask.
- `ACCESS_TOKEN`: token usado na tela de login.
- `MAX_CONTENT_LENGTH_MB`: limite global de upload, em MB.
- `SCHOOL_YEAR`: ano letivo usado em documentos gerados.
- `CONCLUSAO_5ANO_DATE_TEXT`: texto da data usado nas declaracoes em lote do 5o ano.
- `INFORMATIVO_WEEKDAY_DUE`: dia da semana do Informativo Semanal (`0` segunda, `4` sexta).
- `ENABLE_EJA`: ativa trechos opcionais de EJA em quadros suportados.
- `SESSION_COOKIE_SECURE`: marque `1` apenas quando servido exclusivamente via HTTPS.

Tambem existe suporte a:

```powershell
$env:ENABLE_EJA="1"
```

Quando `ENABLE_EJA=1`, os trechos de quadros que suportam EJA tentam considerar a lista EJA carregada.

## Deploy

O `Procfile` define o comando:

```text
web: gunicorn app:app --bind 0.0.0.0:$PORT --workers 3 --timeout 180 --graceful-timeout 180
```

Esse formato e compativel com plataformas como Render, que fornecem a variavel `$PORT`.

Antes de publicar, revise:

- Token e chave secreta.
- Persistencia da pasta `uploads/`.
- Persistencia da pasta `static/fotos/`.
- Disponibilidade dos modelos em `modelos/`.
- Limites de upload e tempo de resposta para planilhas grandes.

## Cuidados com Dados

Este repositorio contem arquivos que podem identificar alunos, como fotos, planilhas, PDFs e documentos gerados. Boas praticas:

- Nao publique dados reais em repositorios publicos.
- Nao envie `uploads/` e `static/fotos/` para ambientes desnecessarios.
- Remova arquivos temporarios e planilhas antigas quando nao forem mais necessarios.
- Evite colocar tokens, senhas ou chaves diretamente no codigo.

## Manutencao e Verificacao

Para uma checagem rapida de sintaxe:

```powershell
python -m py_compile app.py confere.py
```

Para rodar os testes estaticos de seguranca:

```powershell
python -m unittest discover -s tests
```

Ao alterar fluxos de planilha, teste manualmente com uma Lista Piloto representativa, porque boa parte das regras depende da estrutura das abas e colunas.

## Removendo Dados Sensíveis do Git

O `.gitignore` protege novos arquivos em `uploads/`, `static/fotos/`, ambientes virtuais e caches. Para deixar de versionar arquivos ja rastreados sem apagar do disco, use comandos como:

```powershell
git rm --cached -r uploads static/fotos __pycache__
```

Revise o resultado com `git status` antes de commitar.

## Pontos de Atencao

- `app.py` concentra quase toda a aplicacao. Mudancas pequenas devem ser bem localizadas.
- `confere.py` guarda o caminho do Excel na sessao para evitar mistura entre usuarios e workers.
- `templates/prazos_alertas.html` referencia `alerts_mark_sent`, mas nao ha rota correspondente encontrada no codigo atual.
- A rota `/quadros/inclusao` esta desativada e redireciona para o menu de quadros.
- As dependencias incluem pacotes de banco de dados/migracao, mas o codigo atual nao usa modelos SQLAlchemy.
