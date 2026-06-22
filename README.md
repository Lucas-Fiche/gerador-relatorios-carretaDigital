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

| Placeholder no modelo            | Coluna na aba `Dados_Destino`     |
| -------------------------------- | --------------------------------- |
| `<<mes_competencia>>`            | (escolhido no painel)             |
| `<<ID>>`                         | `Id`                              |
| `<<Nome Completo>>`              | `Nome Completo`                   |
| `<<Telefone>>`                   | `Telefone`                        |
| `<<Cargo>>` (Cargo **e** Função) | `Cargo`                           |
| `<<Email>>`                      | `Email`                           |
| `<<Endereço>>`                   | `Endereço`                        |
| `<<CEP>>`                        | `CEP`                             |
| `<<Cidade>>`                     | `Cidade`                          |
| `<<CPF>>`                        | `CNPJ / CPF / IG Favorecido`      |
| `<<Descrição das Atividades>>`   | `DESCRIÇÃO ATIVIDADES`            |
| `<<valor>>`                      | `Valor DL` (formatado `R$ x,xx`)  |
| `<<valor_extenso>>`              | `Valor DL` (por extenso, gerado)  |

> **Atenção:** no modelo original o valor estava fixo (`R$ 15.000,00
> (quinze mil reais)`). Substitua esse trecho por `<<valor>> (<<valor_extenso>>)`
> para que cada recibo use o `Valor DL` do colaborador.
>
> Se quiser que o trecho do corpo "...relativa ao mês abril/2026..." também
> acompanhe o mês escolhido, troque `abril/2026` por `<<mes_competencia>>`.

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
  Codigo.gs        # lógica: menu, leitura da planilha, geração de PDFs, valor por extenso
  Painel.html      # interface (barra lateral) com seleção de mês e colaboradores
  appsscript.json  # manifesto / escopos de permissão
```

## Observações técnicas

- O nome do arquivo gerado é `Recibo - <Nome> - <mês>/<ano>.pdf`.
- O Google Doc temporário criado para cada colaborador é descartado (lixeira)
  após exportar o PDF — só os PDFs ficam na pasta.
- A conversão do `Valor DL` aceita número puro ou texto no formato brasileiro
  (`1.500,00`).
- Para **muitos** colaboradores de uma vez, lembre-se do limite de tempo de
  execução do Apps Script (~6 min). Se necessário, gere em lotes.
