# Orientacoes para Agentes

Este arquivo orienta manutencoes automatizadas neste repositorio. O escopo vale para toda a arvore do projeto.

## Visao Geral

O projeto e uma aplicacao Flask de secretaria escolar. A maior parte da logica esta em `app.py`; o modulo `confere.py` registra um blueprint em `/confere` para comparar listas da escola com PDFs do SED. A aplicacao gera HTML imprimivel e arquivos Excel a partir de planilhas reais da Lista Piloto, fotos de alunos e modelos oficiais em `modelos/`.

Trate o repositorio como sensivel: ele contem fotos, planilhas, uploads e documentos com dados de alunos.

## Stack e Comandos

- Python / Flask.
- pandas, openpyxl, xlrd e pdfplumber.
- Templates Jinja2 em `templates/`.
- Arquivos estaticos em `static/`.
- Modelos oficiais em `modelos/`.

Comandos uteis:

```powershell
python -m py_compile app.py confere.py
python -m unittest discover -s tests
python app.py
flask --app app run --debug
pip install -r requirements.txt
```

O deploy usa:

```text
gunicorn app:app --bind 0.0.0.0:$PORT --workers 3 --timeout 180 --graceful-timeout 180
```

## Regras de Edicao

- Antes de editar, leia o trecho relevante em `app.py`, o template correspondente e, quando houver, o modelo em `modelos/`.
- Preserve o comportamento existente dos fluxos escolares. Muitas regras dependem do layout exato das planilhas.
- Mantenha textos de interface em portugues do Brasil.
- Evite refatoracoes amplas em `app.py` sem pedido explicito. O arquivo e grande e mistura varias rotinas historicas.
- Nao renomeie modelos, templates, fotos ou uploads sem atualizar todos os caminhos.
- Nao remova arquivos de `uploads/`, `static/fotos/` ou `modelos/` sem confirmacao do usuario.
- Nao exponha nem copie dados reais de alunos em respostas, logs ou novos arquivos de exemplo.
- Ao adicionar documentacao, prefira exemplos ficticios.

## Dados e Arquivos Sensiveis

Diretorios e arquivos que merecem cuidado extra:

- `uploads/`: listas enviadas, planilhas temporarias e logs.
- `static/fotos/`: fotos dos alunos, normalmente associadas por RM.
- `modelos/`: modelos oficiais que servem de base para geracao.
- Arquivos `.xlsx`, `.xlsm`, `.pdf`, `.doc` e `.docx` na raiz ou subpastas.

Se precisar criar arquivos de teste, use nomes claramente ficticios e remova-os ao final quando nao forem entregaveis.

## Convencoes Locais

- A sessao Flask guarda caminhos como `lista_fundamental`, `lista_eja`, `declaracao_excel` e filtros das carteirinhas.
- Segredos e configuracoes sensiveis devem vir de ambiente ou `.env`: `FLASK_SECRET_KEY`, `ACCESS_TOKEN`, `MAX_CONTENT_LENGTH_MB`, `SCHOOL_YEAR`, `CONCLUSAO_5ANO_DATE_TEXT`.
- A Lista Piloto Fundamental geralmente usa a aba `LISTA CORRIDA`.
- O Quadro de Atendimento Mensal tambem usa a aba `Total de Alunos`.
- O HTML/CSS final das declaracoes unitarias fica em `templates/declaracao_print.html`; mantenha textos/regras separados do layout quando possivel.
- Fotos de alunos ficam em `static/fotos/` e sao buscadas por RM.
- Modelos de quadro sao arquivos Excel em `modelos/`.
- O sistema calcula prazos usando `modelos/feriados.json` e timezone de Sao Paulo quando disponivel.
- `ENABLE_EJA=1` ativa partes opcionais de processamento EJA em alguns quadros.
- Configuracoes globais ficam em `config.py`.
- Regras de feriados e alertas de prazo ficam em `services/prazos.py`.
- Log transacional de carteirinhas impressas fica em `services/carteirinhas_log.py`.
- Busca e salvamento de fotos por RM ficam em `services/fotos.py`.
- Salvamento de uploads Excel com caminho em sessao fica em `services/upload_sessions.py`.
- Helpers do quadro de atendimento mensal ficam em `services/quadros_atendimento.py`.
- Helpers do quadro quantitativo de inclusao ficam em `services/quadros_inclusao.py`.
- Helpers de quadros de transferencia e quantitativo mensal ficam em `services/quadros_transferencias.py`.
- Helpers compartilhados de data, texto, planilha e upload ficam em `utils/dates.py`, `utils/text.py`, `utils/excel.py` e `utils/uploads.py`.

## Rotas Relevantes

- `/login`, `/logout`, `/upload_listas`, `/`.
- `/declaracao/tipo`, `/declaracao/conclusao_5ano`, `/declaracao/escolaridade_5ano`.
- `/carteirinhas`, `/carteirinhas/marcar_impressas`.
- `/upload_foto`, `/upload_multiplas_fotos`, `/upload_inline_foto`.
- `/quadros`, `/quantinclusao`, `/quadros/atendimento_mensal`, `/quadros/transferencias`, `/quadros/quantitativo_mensal`.
- `/confere/` e `/confere/upload_excel`.
- `/escolas/search`.

## Pontos de Atencao Tecnicos

- Nao reintroduza segredo hard-coded em `app.py`, `confere.py` ou templates.
- `confere.py` deve manter estado por sessao ou armazenamento persistente; nao reintroduza estado global para uploads.
- O log de carteirinhas impressas deve continuar usando `services.carteirinhas_log`, sem voltar para escrita JSON direta.
- `templates/prazos_alertas.html` referencia `alerts_mark_sent`, mas a rota nao existe no codigo atual.
- `/quadros/inclusao` existe, mas esta desativada e redireciona para `/quadros`.
- Ha dependencias de SQLAlchemy/Migrate no `requirements.txt`, mas o codigo atual nao define banco de dados.
- Varios templates carregam assets externos por CDN; ambientes sem internet podem renderizar com estilo incompleto.

## Como Testar Mudancas

Para alteracoes simples de Python:

```powershell
python -m py_compile app.py confere.py
```

Para alteracoes de rotas/templates:

```powershell
python app.py
```

Depois, acesse `http://127.0.0.1:5000`.

Para alteracoes em geracao de Excel, teste com uma planilha que contenha as abas e colunas esperadas. Confira se o arquivo baixado preserva formulas, celulas mescladas e preenchimentos do modelo.

Para alteracoes em carteirinhas ou declaracoes, confira a pagina renderizada no navegador e o comportamento de impressao.

## Estilo de Implementacao

- Prefira ajustes pequenos e verificaveis.
- Reaproveite helpers existentes como `utils.excel.set_merged_cell_value`, `utils.uploads.validate_excel_upload`, `_find_sheet_case_insensitive`, `_build_colmap`, `_pick_col` e os parsers de data.
- Ao lidar com planilhas, use pandas/openpyxl em vez de manipulacao manual de texto.
- Ao validar dados vindos de formulario, mantenha mensagens claras com `flash`.
- Ao retornar arquivos gerados, mantenha `send_file` com `BytesIO`, `download_name` e mimetype correto.
