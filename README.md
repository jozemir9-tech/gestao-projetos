# App Gestão de Projetos

Painel web (Streamlit) para gestão de projetos com cronograma (Gantt + curva S),
histogramas de mão de obra e equipamentos, recursos humanos e contratos.
A base de dados oficial é o Google Sheets; sem acesso, o app usa os CSVs locais
de exemplo em `data/exemplos/` com os mesmos cabeçalhos.

## Como rodar

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py   # (equivalente: streamlit run app.py)
```

Abre em http://localhost:8501. Não há etapa de build nem banco de dados.
O `streamlit_app.py` é só o entry point que chama o `main()` do `app.py` —
é o arquivo que o **Streamlit Community Cloud** procura por padrão na raiz do
repositório (o app de verdade continua sendo o `app.py`).

## Instalar no desktop (Windows)

1. **Servidor + navegador num clique** — crie um atalho de `iniciar_app.bat` na
   área de trabalho (clique direito no arquivo → *Enviar para* → *Área de
   trabalho (criar atalho)*). Ele sobe o Streamlit e abre
   http://localhost:8501 quando o app responder. Fechar a janela minimizada do
   servidor encerra o app.
2. **Ícone próprio, sem barra do navegador** — com o app aberto no navegador:
   - **Chrome:** ⋮ (três pontos) → *Salvar e compartilhar* → *Criar atalho…*
     (em versões novas: *Instalar site como aplicativo*) → nome
     "Gestão de Projetos" → marcar **Abrir como janela**.
   - **Edge:** ⋮ → *Aplicativos* → *Instalar site como aplicativo*.
   - Resultado: ícone na área de trabalho e no menu Iniciar que abre o app em
     janela isolada (como se fosse um programa instalado).
3. **Abrir junto com o PC** (opcional) — `Win+R` → digite `shell:startup` →
   Enter → arraste o atalho do `iniciar_app.bat` para essa pasta.

Requisitos: Python no PATH com `pip install -r requirements.txt` já executado.

## Estrutura

| Arquivo/pasta | Papel |
|---|---|
| `app.py` | Aplicação inteira (8 páginas na barra lateral, na ordem do fluxo: Contratos → Importar → Dashboard → Timeline → Histograma → RH → Relatório → Editar Dados) |
| `.streamlit/config.toml` | Tema escuro, headless, telemetria off e `enableStaticServing` (arquivos do PWA) |
| `conversor-cronograma.py` | Conversor CLI reutilizado pelo app para import/export (não editar sem ler os testes) |
| `leitor_mpp.py` | Leitor puro Python de `.mpp` (MPP14+) usado pela página Importar/Exportar |
| `static/` | Arquivos do **PWA**: `manifest.json`, `sw.js` (service worker) e ícones `icon-192/512.png` — servidos em `/app/static/` |
| `gerar_icones_pwa.py` | Regenera os ícones do PWA (só PIL): `python gerar_icones_pwa.py` |
| `data/exemplos/*.csv` | 6 CSVs de exemplo (separador `;`, UTF-8 com BOM) |
| `requirements.txt` | streamlit, pandas, plotly, openpyxl, gspread, google-auth, requests, olefile |

## Tema e interface

- **6 paletas de tema** — as opções ficam atrás do ícone **🎨 Tema** no topo da
  barra lateral (clique desoculta as paletas): **Escuro**, **Claro**,
  **Azul Profundo**, **Verde Engenharia**, **Sepia Documento** e **Alto
  Contraste**. A troca é em tempo real (sem reiniciar). Cada paleta
  define fundo, superfície, texto, acento, glass e blobs próprios no dicionário
  `TEMAS` do `app.py`.
- O `.streamlit/config.toml` só define o visual da **primeira carga** (fundo
  `#0B1220`); depois disso o CSS do app (`var(--fundo) !important` em
  `html, body, .stApp`, `stHeader` e `stToolbar`) é quem manda — sem isso, tema
  claro ficaria com texto escuro sobre fundo escuro.
- O visual custom (fonte Inter, cartões glass, KPIs, abas, alertas, gráficos,
  animação `surge` de entrada com respeito a `prefers-reduced-motion`) é o bloco
  `_CSS_CORPO` no `app.py`, injetado via `st.markdown`. Os seletores usam os
  `data-testid` do **Streamlit 1.66** — se atualizar o Streamlit, confira os
  seletores no navegador antes de mexer. Hovers usam `color-mix()`.
- Os gráficos Plotly usam `theme=None` + `config={"displaylogo": False}`; sem
  `theme=None`, o tema padrão do Streamlit sobrescreve o estilo custom.
- Todo gráfico define margens próprias: o Plotly **corta** (não expande) texto que
  não cabe na margem — o Gantt calcula a margem esquerda pelo maior nome de tarefa
  e os eixos X usam `b ≥ 34` para as datas não sumirem.
- **Assinatura “por Jozemir Leite”** aparece no rodapé de todo cabeçalho de página
  (classe `.pag-ass`).

## Fonte de dados

- **CSV local (exemplos)** — padrão. Nada a configurar.
- **Google Sheets (API)** — recomendado para apps externos. Publice a planilha
  como API via Google Apps Script (arquivo `apps_script_API.gs`) e informe na
  barra lateral a **URL do Web App** (terminando em `/exec`) e a **API Key**.
  Não precisa de conta de serviço nem `secrets.toml`. Se qualquer aba faltar,
  cai para o CSV local com aviso. Passo a passo no fim deste README.
- **Google Sheets** (conta de serviço) — informe o ID/URL da planilha.
  A planilha precisa das abas `DB_Cronograma_MSProject`, `DB_Histograma_MO`,
  `DB_Histograma_Equip`, `DB_Recursos_Humanos`, `DB_Contratos`, `Timeline_Visao`,
  e as credenciais de uma conta de serviço em `.streamlit/secrets.toml`:

  ```toml
  [gcp_service_account]
  type = "service_account"
  project_id = "seu-projeto"
  private_key_id = "..."
  private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
  client_email = "app@seu-projeto.iam.gserviceaccount.com"
  client_id = "..."
  auth_uri = "https://accounts.google.com/o/oauth2/auth"
  token_uri = "https://oauth2.googleapis.com/token"
  auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
  client_x509_cert_url = "..."
  ```

  Compartilhe a planilha com o e-mail da conta de serviço. Se qualquer aba ou
  credencial faltar, **aquela aba cai automaticamente para o CSV local** e um
  aviso aparece no topo do app (nunca falha em branco).

## Páginas

1. **Dashboard** — KPIs, curva S (planejado x real), avanço por contrato e
   alertas (tarefas atrasadas, críticas, alocação > 100%, contratos vencendo,
   `[FALTANTE]`). Quatro recursos:
   - **2 abas** no topo do painel:
     - **🌍 Visão Geral** — o projeto inteiro numa tela: KPIs, curva S,
       alertas + **indicadores novos**: *Marcos atrasados* (vencidos na
       `Timeline_Visao`), *Frentes paradas com prazo ≤ 45 dias* e
       *físico × financeiro por contrato* (avanço do cronograma vs % medido,
       com alerta de desvio > 10 p.p.).
     - **📅 Por Cronograma (avanço rápido)** — o Gantt com caminho crítico em
       vermelho + tabela editável **igual ao MS Project**: duplo clique na
       coluna *% de avanço*, digite o valor e **✅ Aplicar**; a ✕ na lateral
       marca tarefas para exclusão (com confirmação). Tudo recalcula na hora
       (KPIs, curva S, gráficos) — a gravação na fonte continua pelo *Editar
       Dados*.
   - **2 templates** escolhidos na barra lateral (*Template do Dashboard*):
     **Executivo** (6 KPIs grandes + gráficos amplos — reunião/diretoria) e
     **Analítico** (KPIs compactos + tabelas densas de tarefas e equipe —
     acompanhamento diário). Vale para a aba Visão Geral.
   - **Filtro por Contrato** no topo: recorta KPIs, gráficos, tabelas, RH,
     contratos **e as duas abas** (vinculação vinda da coluna `Contrato` do
     cronograma ou do mapa Frente→Projeto da `Timeline_Visao`); alertas
     continuam do projeto inteiro.
   - **Drill-down nos KPIs**: cada indicador tem um gatilho 🔍 que abre a lista
     exata por trás do número (todas as tarefas, atrasadas com dias de atraso,
     críticas, avanço por frente, equipe, contratos, marcos e frentes paradas).
2. **Cronograma Timeline** — Gantt Plotly (mesmo da aba Por Cronograma) com
   filtros por contrato/frente/responsável/status; caminho crítico em vermelho;
   linha "hoje".
3. **Histograma MO + Equip** — barras Previsto x Real mês a mês (HH, quantidade
   ou custo), com filtros e tabela de desvios.
4. **Recursos Humanos** — tabela com filtro ativo/inativo e alerta de alocação > 100%.
5. **Contratos** — cadastro, vínculo e exclusão:
   - **➕ Cadastrar novo contrato** (expander no topo): ID, cliente, objeto,
     valor e período — entra na hora nos seletores e no Dashboard. Um checkbox
     **💾 Salvar na fonte agora** grava o contrato direto no Sheets/CSV (senão,
     vale só até você usar *💾 Salvar tudo na fonte* na barra lateral).
   - **📎 Vincular cronograma/histograma**: anexa arquivos a um contrato
     (inclui no grupo, renumera IDs, etiqueta as linhas com o contrato).
   - **🗑️ Excluir contrato**: mostra o impacto (quantas tarefas/linhas/marcos
     perdem o vínculo), pede confirmação e remove o contrato + vínculos.
   - Após cadastrar, um botão **✏️ Ir para Editar Dados** leva direto à página
     de edição da base.
   - Card por contrato: valor total, medido, saldo, % avanço e marcos
     vinculados (vem da `Timeline_Visao`).
6. **Relatório (IA)** — relatório executivo em markdown gerado por IA a partir
   dos **números reais** do projeto (tarefas, atrasos, **marcos atrasados**,
   **frentes paradas**, contratos com **físico × financeiro**, efetivo), nos
   estilos **Executivo / Técnico / Ata de reunião** e com suas observações.
   Funciona com qualquer API compatível com OpenAI (seção `[ia]` no
   `secrets.toml` — OpenAI, Groq, DeepSeek, OpenRouter, Gemini) e tem **modo
   local sem IA** que gera o mesmo relatório determinístico. Resultado sai na
   tela e baixa em `.md`.
7. **Importar/Exportar** — upload CSV/XLSX/XML do MS Project **ou do próprio
   arquivo `.mpp`** (MS Project 2010+). A importação é um **assistente**:
   - **Nome do projeto** — etiqueta gravada na coluna `Projeto_Nome`;
   - **Qual contrato** — gravado na coluna `Contrato` (é o que o Dashboard usa
     no filtro por contrato) — com a opção **➕ Novo contrato...** que cadastra
     na hora;
   - **Frente** — escolhe ou cria (➕) a frente das tarefas: reescreve o 1º
     nível do WBS de todas as linhas importadas;
   - **Responsável** — escolhe ou cria (➕): preenche só as tarefas importadas
     que estão sem responsável;
   - **Vínculo com a Cronograma Timeline** — cria as linhas Frente→Projeto na
     `Timeline_Visao` ligadas ao contrato (é o que faz o filtro por contrato
     funcionar);
   - **Modo**: **➕ Acrescentar (incluir no grupo)** — junta o arquivo ao
     cronograma já carregado e renumera os IDs para não colidir (vários
     projetos convivem) — ou **🔄 Substituir** (comportamento antigo).
   - O histograma (MO/Equip) também pergunta **qual contrato** (com ➕ Novo
     contrato) antes de aplicar; o export pergunta o **nome do projeto** para
     nomear os arquivos baixados (MS Project CSV, P6 CSV e XLSX via
      `conversor-cronograma.py`).

## Alterar de qualquer aba e salvar na fonte

**Tudo está conectado**: uma alteração feita em **qualquer** página vale já em
**todas** as outras (Dashboard, Timeline, Histograma, RH, Contratos, Relatório).
Excluir um contrato na página Contratos, por exemplo, some ele da Timeline, do
Dashboard, do filtro por contrato e do relatório na mesma hora.

- **Onde alterar**: avanços/exclusões de tarefas na aba *📅 Por Cronograma* do
  Dashboard; cadastro, vínculo e exclusão de contratos na página *📑
  Contratos*; importações no *🔁 Importar/Exportar*.
- **💾 Salvar tudo na fonte** — quando há alteração pendente, aparece um bloco
  na **barra lateral** listando as abas alteradas e o botão **💾 Salvar agora**.
  Ele grava todas elas de uma vez na fonte ativa (Google Sheets ou CSV local),
  para valer no arquivo/planilha. Sem isso, as alterações valem só na sessão.
- **↩ Restaurar dados de origem** (também na lateral) descarta as alterações
  pendentes e volta ao que está gravado.

## Publicar na nuvem (Android, iOS e acesso de qualquer lugar)

A nuvem hospeda o app num endereço fixo — o PC pode ficar desligado. Sem
`secrets.toml` na nuvem o app roda com os CSVs locais; com os secrets, o Google
Sheets passa a ser lido de lá.

### 1. Subir o código para o GitHub

**Sem Git (só navegador):** crie o repositório em <https://github.com/new>
(*New repository*, público ou privado) e use *Add file → Upload files* para
enviar `app.py`, `conversor-cronograma.py`, `leitor_mpp.py`,
`gerar_icones_pwa.py`, `requirements.txt`, `README.md`, `.gitignore`,
`.streamlit/config.toml`, a pasta `data/exemplos/` e a pasta `static/`
(arquivos do PWA). Um ZIP pronto com exatamente esses arquivos (sem segredos)
fica na pasta do projeto: `gestao-projetos-publicacao.zip` — descompacte e
arraste os arquivos no GitHub.

**Com Git (opcional):**

```bash
git init
git add .
git commit -m "App Gestao de Projetos"
git branch -M main
git remote add origin https://github.com/SEU-USER/SEU-REPO.git
git push -u origin main
```

O `.gitignore` já exclui `.streamlit/secrets.toml`, credenciais e a documentação
interna. **Nunca** envie o `secrets.toml` na mão (pelo navegador ele fica de
fora do gitignore — cuidado ao arrastar a pasta `.streamlit`).

### 2. Publicar no Streamlit Community Cloud

1. Entre em <https://share.streamlit.io> com a conta do GitHub.
2. *Create app* → repositório escolhido, branch `main`, arquivo do app: `app.py`.
3. *Advanced settings* → *Secrets*: cole o bloco `[gcp_service_account]` do
   README (opcional — sem ele o app usa os CSVs de exemplo).
4. *Deploy* (cerca de 2 minutos). O app fica em `https://SEU-APP.streamlit.app`.

### 2.1 Acesso dos colegas: link direto ou convite por e-mail

- **App público (padrão):** qualquer pessoa com o link acessa. Use o botão
  **🔗 Compartilhar** da barra lateral para copiar o link e mande por
  e-mail/WhatsApp.
- **App privado com convite por e-mail:** no Community Cloud, abra o app → menu
  ⋮ → *Settings* → *Visibility* → **Private** → *Invite viewers* e informe os
  e-mails. Cada pessoa recebe um e-mail de convite do Streamlit e, após criar
  uma conta gratuita (e-mail, Google ou GitHub), passa a acessar pelo link.
  Convidados **veem apenas o app** — não editam código nem dados.
- Os **dados** seguem protegidos à parte: o app lê a planilha com a credencial
  dele — quem acessa o app vê o que o app mostra, mesmo sem acesso direto ao
  Google Sheets.

### 3. Instalar no celular (PWA)

O app já é um **PWA instalável**: `static/manifest.json` + `sw.js` + ícones,
com o `<head>` (manifest, meta do iOS e registro do service worker) injetado por
`injetar_pwa()` no `app.py`. Funciona em Android e iOS sem loja e sem conta paga.

- **Bloco 📲 Instalar + 🔗 Compartilhar** no fim da barra lateral (visível em
  qualquer página): *Instalar* abre o passo a passo do seu aparelho
  (Android/iOS/Windows) e, quando o navegador oferecer instalação, o botão
  **Instalar agora** dispara o install prompt de verdade; *Compartilhar* copia
  o link do app para a área de transferência (no celular abre o
  compartilhamento nativo) — é o link para mandar por e-mail.
- **Android (Chrome 108+):** abra a URL → menu ⋮ → **Instalar app** /
  *Adicionar à tela inicial*. O Chrome pode pedir conexão prévia com a URL.
- **iOS (Safari):** abra a URL → ícone de compartilhar → **Adicionar à Tela de
  Início** → *Adicionar* (o Safari ignora o service worker, mas o app abre em
  tela cheia).
- **Windows (Chrome/Edge):** menu ⋮ → **Instalar app** — o app entra no Menu
  Iniciar como aplicativo de verdade.
- O ícone abre o app em tela cheia, como se fosse nativo (o CSS do app já tem
  layout responsivo para celular).
- Depois de gerar o `.apk` (seção 4), preencha a constante `URL_APK` no topo do
  `injetar_pwa()` (`app.py`) com o link do arquivo: o painel 📲 ganha o link
  **⬇ Baixar .apk** ao lado do "Instalar agora".

### 4. Gerar o `.apk` (Android, opcional — se quiser arquivo instalável)

O PWA acima já instala no Android sem APK. Se precisar de um **`.apk`/`.aab` de
verdade** (para Play Store ou compartilhar o arquivo), use o **PWABuilder** —
mas ele **exige a URL pública** do passo 2 (não roda em `localhost`):

1. Publique no Community Cloud e confirme que
   `https://SEU-APP.streamlit.app/app/static/manifest.json` responde 200.
2. Abra <https://www.pwabuilder.com> → cole a URL do app → *Start*.
3. O PWABuilder valida o manifest e o service worker e oferece **Android →
   Generate** → baixa o pacote (`.aab` para Play Store e `.apk` para instalar
   direto).
4. Para a Play Store o `.aab` precisa de assinatura (conta de desenvolvedor
   Google, US$ 25 únicos) e passa pela revisão da Google. O `.apk` “package ID”
   vem do manifest — se quiser mudar nome/pacote, edite `static/manifest.json`.

**Alternativa local (Bubblewrap):** exige Node.js + Java + Android SDK nesta
máquina (nenhum dos três está instalado aqui). Se quiser rodar local, instale
`npm i -g @bubblewrap/cli` e `bubblewrap init --manifest
https://SEU-APP.streamlit.app/app/static/manifest.json` → `bubblewrap build`.

### Atenção na nuvem

- A URL é **pública** por padrão: quem tiver o link vê os dados exibidos. Para
  dados sensíveis, mantenha o repositório privado e avalie restrições de acesso
  antes de compartilhar o link.
- O repo precisa conter `data/exemplos/` (fallback local) e
  `conversor-cronograma.py` (exportadores).

## Convenções de dados

- Datas **DD/MM/AAAA**; moeda **R$**; percentual 0–100; CSV com separador **`;`**
  e codificação UTF-8 com BOM (o que o Excel/Sheets/MS Project esperam).
- **Frente = 1º nível do WBS** (`FUND.20` → frente `FUND`).
- Vínculo tarefa→contrato: o app usa a coluna `Contrato` se existir no
  cronograma; caso contrário, mapeia Frente→Projeto pela `Timeline_Visao`
  (a coluna `Projeto` dela recebe o `ID_Contrato`).
- Nada é inventado: campos vazios obrigatórios aparecem como `[FALTANTE]` nos alertas.

## Publicar a planilha como API (Google Apps Script Web App)

Transforma a planilha numa API HTTP com API Key, para o app (ou qualquer
aplicativo externo) consultar sem conta de serviço. Use o arquivo
`apps_script_API.gs`.

1. **Criar o script:** abra a planilha no navegador → menu **Extensões →
   Apps Script**. Apague o que estiver em `Código.gs` e cole o conteúdo de
   `apps_script_API.gs`. Salve (Ctrl+S).
2. **Gerar a chave:** no seletor de funções (acima do editor) escolha
   **`gerarEInstalarChave`** → **Executar**. Na 1ª vez autorize a planilha
   (clique *Revisar permissões* → sua conta → *Avançado* → *Ir para o projeto
   (inseguro)* → *Permitir*). A chave aparece no **Registro de execução**
   (painel embaixo). **Copie e guarde** — é a sua senha (API Key).
3. **Publicar:** clique em **Implantar → Nova implantação**. No ícone de
   engrenagem escolha **Tipo: Aplicativo da Web**. Preencha:
   - *Executar como:* **Eu** (o dono da planilha)
   - *Quem tem acesso:* **Qualquer pessoa**
   Clique em **Implantar** e **copie a URL** mostrada. Ela termina em
   `/exec` — é a URL da API.
4. **Testar:** abra a URL no navegador adicionando a chave, por exemplo:
   `SUA_URL/exec?action=ping&api_key=SUA_CHAVE`. Deve responder
   `{"success":true,"message":"API funcionando!"}`. Depois tente
   `?action=info&api_key=SUA_CHAVE` (mostra as abas e colunas) e
   `?action=list&aba=DB_Contratos&api_key=SUA_CHAVE` (lista uma aba).
5. **Usar no app:** na barra lateral escolha **Google Sheets (API)** e cole a
   URL da API e a API Key. O app baixa as 6 abas e mostra "✅ API conectada".

**Endpoints (GET):** `action=list` (parâmetro `aba`, opcional `busca`,
`coluna`+`valor`, `limite`), `action=get` (`aba` + `chave`), `action=info`,
`action=ping`. **Escrita (POST, corpo JSON):** `action=create|update|delete`
com `aba` e os campos da linha.

**Atenção:** se você editar o script, clique **Implantar → Gerenciar
implantações → ✏️ → Nova versão** e reimplante — a URL antiga não atualiza
sozinha. A chave fica salva nas propriedades do script; rode
`gerarEInstalarChave` de novo só se quiser trocá-la. A coluna-chave de cada
aba (`ID_Tarefa`, `ID_Contrato`, etc.) é definida em `CONFIG.ABAS` no topo do
script — ajuste se os nomes das colunas-chave forem diferentes.

## Atenção

- `Status` **não sobrevive ao round-trip** com MS Project/P6 (o formato não tem
  essa coluna): o app deriva `Status` de `%_Concluido` na importação. Na aba
  padrão do Sheets ele é preservado.
- **Importar `.mpp` (MS Project 2010+/MPP14+):** a página Importar/Exportar
  aceita o arquivo binário `.mpp` direto — o leitor é puro Python
  (`leitor_mpp.py`, usa só `olefile`), então roda também na nuvem, sem Java nem
  MS Project instalado. O que **não** vem do binário e é derivado/aviso:
  `Caminho_Critico_SIM_NAO` (marcado `NAO`), datas reais (derivadas do
  `%_Concluido`) e `Responsavel` (fica vazio — o vínculo exato recurso↔tarefa
  usa índices internos não decifrados; nada é inventado). O WBS é reconstruído
  pela hierarquia de datas (rolamento validado 43/43 no arquivo real).
  Versões antigas (MPP9 e anteriores) são recusadas com mensagem clara — salve
  como XML ou CSV nesses casos.
- `.xer` (P6) não é lido: exporte antes para CSV no software.
- Os CSVs de exemplo estão no período set–dez/2026; edite-os direto para testar
  (mantenha cabeçalhos, separador `;` e datas DD/MM/AAAA).
