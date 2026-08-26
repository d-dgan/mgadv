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
│   ├── Bootstrap + Bootstrap Icons (CDN)
│   ├── PDF.js (CDN) — lê o PDF do CJ no navegador
│   ├── PizZip + docxtemplater (CDN) — gera os .docx
│   └── JSZip (CDN) — empacota os 5 arquivos em ZIP
│
├── HTML
│   ├── Header (igual ao site principal)
│   ├── Passo 1: Upload do PDF
│   ├── Passo 2: Formulário de dados do cliente
│   ├── Passo 3: Tipo de benefício (checkboxes)
│   ├── Botão gerar + status
│   └── Footer (igual ao site principal)
│
└── JavaScript (type="module")
    ├── processarPDF(file) — extrai texto do PDF via PDF.js
    ├── extrairDadosCJ(texto) — parseia os campos do relatório do CJ
    ├── preencherFormulario(dados) — preenche os inputs
    ├── gerarDocumentos() — orquestra a geração dos 5 docs
    ├── criarDocx(titulo, secoes) — monta o XML do .docx e empacota
    └── gerarProcuracao/Declaracao/Renuncia/Representacao/Contrato(campos)
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

A função `extrairDadosCJ()` lê linha a linha: quando encontra um label conhecido (ex: `"Nome"`), pega a próxima linha como valor.

---

## Como funciona a geração dos .docx

Os documentos são gerados **100% no navegador**, sem enviar dados para nenhum servidor:

1. A função `criarDocx()` monta o XML interno de um arquivo `.docx` (formato Open XML)
2. O JSZip empacota esse XML nos arquivos necessários (document.xml, styles.xml, etc.)
3. Os 5 arquivos `.docx` são empacotados num único `.zip`
4. O ZIP é baixado direto pelo navegador

---

## Deploy

O deploy é automático via **GitHub Pages**:

- Qualquer `git push` na branch `main` dispara o workflow `.github/workflows/static.yml`
- O site atualiza em ~1-2 minutos
- Domínio customizado configurado no arquivo `CNAME`: `murarogoncalves.com`

---

## Pendências / próximos passos

- [ ] **Remover os logs de debug** (`console.log`) após validar a extração do PDF
- [ ] Validar geração dos .docx no Word e no LibreOffice
- [ ] Conferir se o layout das tabelas nos documentos está igual ao original
- [ ] Testar com outros PDFs do CJ (clientes diferentes)
- [ ] (Opcional) Melhorar o Contrato de Honorários com todas as cláusulas completas

---

## Tecnologias utilizadas

| Biblioteca | Versão | Uso |
|---|---|---|
| Bootstrap | 5.x | Layout e componentes visuais |
| PDF.js | 4.0.379 | Leitura do PDF no navegador |
| PizZip | 3.1.4 | Dependência do docxtemplater |
| docxtemplater | 3.47.5 | Geração de .docx |
| JSZip | 3.10.1 | Empacotamento em ZIP |

> Todas as bibliotecas são carregadas via CDN — sem instalação de dependências necessária.
