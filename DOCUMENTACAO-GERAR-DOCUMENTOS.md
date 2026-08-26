# Documentação — Página de Geração de Documentos

> **Arquivo:** `gerar-documentos.html`  
> **URL de produção:** `https://murarogoncalves.com/gerar-documentos.html`  
> **Acesso:** restrito — somente quem tem o link (página oculta, sem link no menu, noindex)

---

## O que faz

Gera automaticamente os 5 documentos jurídicos que a Isabelle usa com cada cliente, a partir do relatório exportado pelo **Cálculo Jurídico (CJ)**. Tudo roda no navegador — zero backend, zero servidor, zero custo.

**Documentos gerados:**
1. Procuração Judicial
2. Declaração de Hipossuficiência
3. Termo de Renúncia
4. Termo de Representação e Autorização
5. Contrato de Honorários

---

## Como usar (passo a passo)

### 1. Exportar o relatório do cliente no CJ

1. Acesse o **Cálculo Jurídico** → menu **Contatos**
2. Localize o cliente pelo nome e clique sobre ele
3. Dentro do perfil, clique em **"Gerar Relatório"**
4. Selecione os campos: **Nome, CPF, Telefone, Nacionalidade, Estado civil, Profissão, Endereço**
5. Clique em **Gerar Relatório** → salva um PDF no computador

### 2. Abrir a página

Acesse: `https://murarogoncalves.com/gerar-documentos.html`

### 3. Importar o PDF

- Clique na área de upload ou arraste o PDF exportado do CJ
- Os campos são preenchidos automaticamente

### 4. Conferir os dados

- Verifique se nome, CPF, endereço, etc. estão corretos
- Ajuste manualmente se necessário
- Preencha a **data do atendimento**

### 5. Selecionar o tipo de benefício

- Marque os tipos de benefício que serão requeridos
- Essa informação vai para o **Termo de Representação e Autorização**

### 6. Gerar e baixar

- Clique em **"Gerar e baixar ZIP com os 5 documentos"**
- Um arquivo `.zip` com os 5 `.docx` é baixado automaticamente
- Abra, imprima e o cliente assina

---

## Como rodar localmente (para desenvolvimento)

### Pré-requisitos

- [Git](https://git-scm.com/)
- Qualquer servidor HTTP local (não abre direto no navegador por restrição do PDF.js com `file://`)

### Clonar o repositório

```bash
git clone https://github.com/d-dgan/mgadv.git
cd mgadv
```

### Rodar localmente

**Opção A — Python (mais simples):**
```bash
python3 -m http.server 8080
```
Acesse: `http://localhost:8080/gerar-documentos.html`

**Opção B — Node.js:**
```bash
npx serve .
```
Acesse: `http://localhost:3000/gerar-documentos.html`

**Opção C — VS Code:**
- Instale a extensão **Live Server**
- Clique com botão direito em `gerar-documentos.html` → **Open with Live Server**

> ⚠️ **Não abra o arquivo diretamente no navegador** (`file://`). O PDF.js não funciona com protocolo `file://` por restrições de segurança do navegador.

---

## Estrutura do arquivo `gerar-documentos.html`

```
gerar-documentos.html
│
├── HEAD
│   ├── Bootstrap + Bootstrap Icons (local)
│   ├── PDF.js (CDN) — lê o PDF do CJ no navegador
│   ├── PizZip + docxtemplater (local) — preenchem os modelos .docx
│   └── JSZip (CDN) — empacota os 5 arquivos em ZIP
│
├── HTML
│   ├── Header (igual ao site principal)
│   ├── Passo 1: Upload do PDF
│   ├── Passo 2: Formulário de dados do cliente
│   ├── Passo 3: Tipo de benefício (checkboxes + escolha única nas sub-opções)
│   ├── Botão gerar + status
│   └── Footer (igual ao site principal)
│
└── JavaScript (type="module")
    ├── processarPDF(file) — extrai texto do PDF via PDF.js
    ├── reconstruirLinhas(items) — reagrupa os fragmentos do PDF em linhas
    ├── extrairDadosCJ(texto) — parseia os campos do relatório do CJ
    ├── preencherFormulario(dados) — preenche os inputs
    ├── getBeneficios() — monta a frase da CLÁUSULA PRIMEIRA do contrato
    ├── marcasSubopcoes() — decide o X das sub-opções (urbana/rural etc.)
    ├── baixarTemplate(url) — busca o modelo .docx
    ├── preencherModelo(buffer, dados) — troca as tags pelos dados do cliente
    └── gerarDocumentos() — orquestra os 5 docs e empacota no ZIP
```

### Arquivos de apoio

```
assets/docs/modelos/     — os 5 .docx originais do escritório (fonte da verdade)
assets/docs/templates/   — os mesmos .docx com tags {nome}, {cpf}… (gerados)
assets/vendor/docx/      — PizZip e docxtemplater
tools/build-templates.py — regera assets/docs/templates/ a partir de modelos/
```

---

## Como funciona a extração do PDF

O relatório do CJ tem este formato:
```
Nome
José Bernardo da Silva
CPF
440.858.009-06
Nacionalidade
brasileiro
Estado civil
Viúvo(a)
...
```

Na verdade o PDF é uma **tabela de duas colunas**: o rótulo fica em `x≈49` e o valor em `x≈235`, os dois na *mesma* altura da página. O PDF.js não devolve linhas — devolve fragmentos soltos com coordenadas.

Por isso `reconstruirLinhas()` reagrupa os fragmentos pela coordenada Y (`item.transform[5]`, tolerância de 3pt), ordena cada grupo por X e insere espaço só onde existe vão horizontal real. O resultado é uma linha por linha visual:

```
Nome José Bernardo da Silva
CPF 440.858.009-06
Nacionalidade brasileiro
```

Aí `extrairDadosCJ()` casa o rótulo no começo da linha (ignorando maiúsculas e um `:` opcional) e pega o resto como valor.

---

## Como funciona a geração dos .docx

Os documentos são gerados **100% no navegador**, sem enviar dados para nenhum servidor.

Os `.docx` **não são montados do zero** — são os próprios modelos do escritório, preenchidos. É isso que garante que o documento gerado saia idêntico ao modelo: cabeçalho, rodapé, logo, fontes, margens e todas as cláusulas vêm do arquivo original e nunca são tocados.

1. `tools/build-templates.py` (rodado **em desenvolvimento**, não no navegador) copia cada modelo de `assets/docs/modelos/` para `assets/docs/templates/`, trocando só os valores variáveis por tags `{nome}`, `{cpf}`, `{endereco}`…
2. No navegador, `baixarTemplate()` busca o template e `preencherModelo()` usa o docxtemplater pra trocar as tags pelos dados do cliente
3. Os 5 arquivos `.docx` são empacotados num único `.zip` pelo JSZip
4. O ZIP é baixado direto pelo navegador

### O que é variável em cada documento

| Tag | Vem de |
|---|---|
| `{nome}` `{cpf}` `{nacionalidade}` `{estado_civil}` `{profissao}` `{celular}` `{endereco}` | Formulário (preenchido pelo PDF do CJ) |
| `{local}` `{data}` | Formulário |
| `{beneficios}` | CLÁUSULA PRIMEIRA do contrato — benefícios marcados no formulário |
| `{u_urbana}` `{u_rural}` `{u_temporaria}` `{u_permanente}` `{u_idoso}` `{u_pcd}` | `X` ou `_` nas sub-opções do Termo de Representação |

Duas decisões tomadas com o escritório:

- **A lista dos 12 benefícios do Termo de Representação sai sempre inteira marcada** (`X`), dando poderes amplos perante o INSS. Só as sub-opções (urbana/rural, temporária/permanente, idoso/PCD) é que seguem o formulário — e são de escolha única.
- **A data virou texto fixo.** O modelo usava um campo `TIME` do Word, que se atualizava sozinho toda vez que o arquivo era aberto — o documento assinado ficaria com a data errada.

### Como alterar um modelo

Edite o `.docx` em `assets/docs/modelos/`, rode `python tools/build-templates.py` e faça commit dos dois (modelo e template). O script falha se o layout esperado mudar, em vez de gerar um documento silenciosamente errado.

---

## Deploy

O deploy é automático via **GitHub Pages**:

- Qualquer `git push` na branch `main` dispara o workflow `.github/workflows/static.yml`
- O site atualiza em ~1-2 minutos
- Domínio customizado configurado no arquivo `CNAME`: `murarogoncalves.com`

---

## Pendências / próximos passos

- [ ] **Remover os logs de debug** (`console.log`) após validar a extração do PDF
- [ ] Validar os .docx gerados no Word e no LibreOffice
- [ ] Testar com outros PDFs do CJ (clientes diferentes)

---

## Tecnologias utilizadas

| Biblioteca | Versão | Uso |
|---|---|---|
| Bootstrap | 5.x | Layout e componentes visuais |
| PDF.js | 4.0.379 | Leitura do PDF no navegador |
| PizZip | 3.2.0 | Dependência do docxtemplater (local, `assets/vendor/docx/`) |
| docxtemplater | 3.69.3 | Preenchimento dos modelos .docx (local, `assets/vendor/docx/`) |
| JSZip | 3.10.1 | Empacotamento em ZIP |

> O PizZip e o docxtemplater ficam no repositório, não em CDN: as URLs do cdnjs usadas antes retornavam **404** e as bibliotecas nunca chegavam a carregar.
