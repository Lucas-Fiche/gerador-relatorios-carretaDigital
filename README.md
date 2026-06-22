# Gerador de Relatórios / Recibos — Carreta Digital

Sistema que substitui o painel antigo (em que era preciso selecionar **um**
colaborador e **um** mês por vez). Agora um único painel permite:

- Escolher o **mês de competência** (preenche `<<mes_competencia>>`);
- Marcar **"Todos os colaboradores"** ou selecionar **individualmente** quem terá relatório gerado;
- Gerar **1 PDF por colaborador** automaticamente, a partir do modelo, salvos em uma pasta do Drive.

Os dados vêm da aba **`Dados_Destino`** da planilha.

---

## Como funciona

1. O menu **📄 Gerador de Relatórios → Abrir painel** abre uma barra lateral.
2. O painel lê a aba `Dados_Destino` e lista os colaboradores.
3. Ao clicar em **Gerar relatórios (PDF)**, o script:
   - copia o Google Doc modelo,
   - substitui os campos `<<...>>` com os dados de cada colaborador,
   - exporta um PDF para a pasta de destino,
   - devolve os links dos PDFs gerados.

---

## Campos do modelo (placeholders)

O modelo (Google Doc) deve conter os marcadores abaixo. Espaços internos são
tolerados (ex.: `<< Cargo >>` funciona igual a `<<Cargo>>`).

| Placeholder no modelo            | Conteúdo gerado                                   |
| -------------------------------- | ------------------------------------------------- |
| `<<N_Recibo>>`                   | Nº do recibo no formato `AAAAMM-ID` (ex.: `202606-1`) |
| `<<mes_competencia>>`            | Mês escolhido no painel, formato `maio de 2026`   |
| `<<ID>>`                         | `Id`                                              |
| `<<Nome Completo>>`              | `Nome Completo`                                   |
| `<<Telefone>>`                   | `Telefone`                                        |
| `<<Cargo>>` (Cargo **e** Função) | `Cargo`                                           |
| `<<Email>>`                      | `Email`                                           |
| `<<Endereço>>`                   | `Endereço`                                        |
| `<<CEP>>`                        | `CEP`                                             |
| `<<Cidade>>`                     | `Cidade`                                          |
| `<<CPF>>`                        | `CNPJ / CPF / IG Favorecido`                      |
| `<<Descrição das Atividades>>`   | `DESCRIÇÃO ATIVIDADES`                            |
| `<<valor>>`                      | `Valor DL` formatado **sem** símbolo (`15.000,00`) |
| `<<valor_extenso>>`              | `Valor DL` por extenso (ex.: `quinze mil reais`)  |

> **`<<valor>>` não inclui o "R$".** No modelo o texto já é `R$ <<valor>>`,
> então o placeholder devolve apenas o número (`15.000,00`). Não coloque outro
> "R$" junto ao `<<valor>>`.
>
> **`<<N_Recibo>>`** é montado automaticamente como `ano + mês(2 dígitos) + "-" + Id`.
> Ex.: junho de 2026, bolsista de Id 1 → `202606-1`.

---

## Instalação (passo a passo)

### 1. Converter o modelo `.docx` para Google Docs
- Faça upload do `.docx` no Google Drive.
- Abra-o e use **Arquivo → Salvar como Documentos Google**.
- Ajuste os placeholders conforme a tabela acima (principalmente `<<valor>>` e
  `<<valor_extenso>>`).
- Copie o **ID do documento** da URL:
  `https://docs.google.com/document/d/`**`ESTE_É_O_ID`**`/edit`

### 2. Criar a pasta de destino dos PDFs
- Crie uma pasta no Drive e copie o **ID** da URL:
  `https://drive.google.com/drive/folders/`**`ESTE_É_O_ID`**

### 3. Instalar o script na planilha
- Abra a planilha → **Extensões → Apps Script**.
- Crie os arquivos com o mesmo nome dos que estão em `apps-script/`:
  - `Codigo.gs`  → cole o conteúdo de `apps-script/Codigo.gs`
  - `Painel.html` → cole o conteúdo de `apps-script/Painel.html`
  - (opcional) ajuste o manifesto `appsscript.json` com o de `apps-script/`.
- No topo do `Codigo.gs`, preencha:
  ```js
  TEMPLATE_DOC_ID: 'ID_DO_DOC_MODELO',
  OUTPUT_FOLDER_ID: 'ID_DA_PASTA_DESTINO'
  ```
- Salve.

### 4. Usar
- Recarregue a planilha (o menu **📄 Gerador de Relatórios** aparecerá).
- **Abrir painel** → escolher mês → marcar colaboradores → **Gerar relatórios (PDF)**.
- Na primeira execução o Google pedirá autorização das permissões.

---

## Estrutura do projeto

```
apps-script/
  Codigo.gs        # lógica: menu, API Web App, leitura da planilha, geração de PDFs, valor por extenso
  Painel.html      # interface da barra lateral (uso direto dentro da planilha)
  appsscript.json  # manifesto / escopos / configuração do Web App
docs/              # site estático servido pelo GitHub Pages (Opção B de uso)
  index.html       # mesma tela, porém standalone, consumindo a API Web App
  config.js        # URL do Web App + token (você preenche)
```

---

## Usar pela WEB (GitHub Pages + API Web App)

Em vez de abrir o painel dentro da planilha, é possível hospedar a tela no
GitHub Pages. A página é **estática**; quem lê a planilha e gera os PDFs continua
sendo o Apps Script, agora publicado como **Web App** (API). A página apenas
chama essa API.

```
[ GitHub Pages: docs/index.html ]  --fetch-->  [ Apps Script Web App ]  -->  Sheets / Docs / Drive
```

### 1. Definir o token
- Em `apps-script/Codigo.gs`, defina `CONFIG.API_TOKEN` com uma string longa e
  aleatória.
- Use exatamente o **mesmo** valor em `docs/config.js` (`API_TOKEN`).

### 2. Publicar o Apps Script como Web App
- No editor do Apps Script: **Implantar → Nova implantação → Tipo: App da Web**.
- Configure:
  - **Executar como:** *Eu* (você, dono da planilha) — assim ninguém precisa logar.
  - **Quem pode acessar:** *Qualquer pessoa*.
- Implante, autorize as permissões e copie a **URL** (termina em `/exec`).

### 3. Configurar o site
- Em `docs/config.js`, preencha `API_URL` com a URL `/exec` e `API_TOKEN` com o token.

### 4. Ativar o GitHub Pages
- No GitHub: **Settings → Pages**.
- **Source:** *Deploy from a branch* → branch `main` (ou a sua) → pasta **`/docs`**.
- Salve. A URL pública aparece em instantes (ex.: `https://usuario.github.io/repo/`).

### Segurança (importante)
- O GitHub Pages é **público**, então o `API_TOKEN` em `config.js` fica visível
  no código-fonte do site. Ele serve para **dificultar acesso casual**, não como
  segredo forte. Quem obtiver o token e a URL consegue chamar a API.
- Para proteção real, prefira o uso pelo painel dentro da planilha, ou a
  abordagem 100% no navegador com login Google (OAuth por usuário).
- Os links de PDF retornados abrem para quem tem acesso à pasta do Drive
  (normalmente você, dono). Compartilhe a pasta se outras pessoas precisarem abrir.

## Observações técnicas

- O nome do arquivo gerado é `Recibo <N_Recibo> - <Nome>.pdf` (ex.: `Recibo 202606-1 - João Silva.pdf`).
- O Google Doc temporário criado para cada colaborador é descartado (lixeira)
  após exportar o PDF — só os PDFs ficam na pasta.
- A conversão do `Valor DL` aceita número puro ou texto no formato brasileiro
  (`1.500,00`).
- Para **muitos** colaboradores de uma vez, lembre-se do limite de tempo de
  execução do Apps Script (~6 min). Se necessário, gere em lotes.
