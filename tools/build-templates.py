#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera os templates do docxtemplater a partir dos .docx modelos do escritorio.

Os modelos em assets/docs/modelos/ sao os arquivos originais, com cabecalho,
rodape, logo, fontes e todas as clausulas. Este script copia cada um deles
trocando apenas os valores variaveis por tags {tag}, sem tocar em mais nada
do pacote — por isso o documento gerado sai visualmente identico ao modelo.

Uso:  python tools/build-templates.py
Saida: assets/docs/templates/*.docx
"""

import os
import re
import shutil
import sys
import zipfile

RAIZ    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ORIGEM  = os.path.join(RAIZ, 'assets', 'docs', 'modelos')
DESTINO = os.path.join(RAIZ, 'assets', 'docs', 'templates')

MODELOS = [
    ('1. Procuração judicial Muraro Gonçalves.docx',    '1-procuracao.docx'),
    ('2. Declaração de hipossuficiência.docx',          '2-declaracao.docx'),
    ('3. Termo de renúncia.docx',                       '3-renuncia.docx'),
    ('4. Termo de representação e autorização.docx',    '4-representacao.docx'),
    ('5. Contrato de honorários.docx',                  '5-contrato.docx'),
]

# Titulo do bloco de qualificacao do cliente — muda de documento pra documento.
# Os dados da advogada (OUTORGADO / REPRESENTANTE / CONTRATADO) ficam intactos.
BLOCOS_CLIENTE = ('OUTORGANTE', 'Declarante', 'REPRESENTADO(A)', 'CONTRATANTE')

# Ordem fixa dos 5 valores que vem logo depois do rotulo "Celular:"
VALORES_APOS_CELULAR = ['nacionalidade', 'estado_civil', 'profissao', 'cpf', 'celular']

RE_PARA  = re.compile(r'<w:p(?:\s[^>]*)?(?:/>|>.*?</w:p>)', re.S)
RE_TEXTO = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)
RE_RUN   = re.compile(r'<w:r(?:\s[^>]*)?>.*?</w:r>', re.S)
RE_RPR   = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
RE_ABRE  = re.compile(r'^(<w:p(?:\s[^>]*)?>)(<w:pPr>.*?</w:pPr>)?', re.S)


def esc(txt):
    return txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def texto_de(para):
    return ''.join(RE_TEXTO.findall(para))


def trocar_texto(para, novo):
    """Substitui todo o texto do paragrafo por um unico run com `novo`.

    Preserva o <w:pPr> (alinhamento, marcador de lista, recuo) e o <w:rPr> do
    primeiro run com texto. Como todos os paragrafos que este script mexe tem
    formatacao uniforme entre os runs, colapsar em um run so nao muda o visual —
    e e obrigatorio pro docxtemplater, que exige a tag inteira num unico <w:t>.
    """
    m = RE_ABRE.match(para)
    if not m:
        raise ValueError('paragrafo em formato inesperado: %r' % para[:120])
    abre, ppr = m.group(1), (m.group(2) or '')

    rpr = ''
    for run in RE_RUN.findall(para):
        if '<w:t' in run:
            achou = RE_RPR.search(run)
            if achou:
                rpr = achou.group(0)
            break

    return '%s%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (
        abre, ppr, rpr, esc(novo))


def indice_por_texto(textos, alvo, a_partir_de=0):
    for i in range(a_partir_de, len(textos)):
        if textos[i] == alvo:
            return i
    raise LookupError('paragrafo %r nao encontrado' % alvo)


def taguear_bloco_cliente(paras, textos):
    """Troca os 7 dados do cliente por tags. Devolve dict {indice: novo_texto}."""
    inicio = None
    for bloco in BLOCOS_CLIENTE:
        if bloco in textos:
            inicio = textos.index(bloco)
            break
    if inicio is None:
        raise LookupError('nenhum bloco de qualificacao do cliente encontrado')

    trocas = {}

    # "Nome:" e, no paragrafo seguinte, o nome.
    i_nome = indice_por_texto(textos, 'Nome:', inicio)
    trocas[i_nome + 1] = '{nome}'

    # Depois de "Celular:" vem, em sequencia, os 5 valores da linha de cima.
    i_celular = indice_por_texto(textos, 'Celular:', inicio)
    for n, tag in enumerate(VALORES_APOS_CELULAR, start=1):
        trocas[i_celular + n] = '{%s}' % tag

    # "Endereço:" e, no paragrafo seguinte, o endereco.
    i_endereco = indice_por_texto(textos, 'Endereço:', inicio)
    trocas[i_endereco + 1] = '{endereco}'

    # Conferencia: o que esta sendo trocado tem que ser um valor de verdade,
    # nunca um rotulo — se o layout do modelo mudar, o build falha aqui em vez
    # de gerar um documento silenciosamente errado.
    for i in trocas:
        if textos[i].endswith(':') or textos[i] in BLOCOS_CLIENTE:
            raise ValueError(
                'ia trocar o rotulo %r (paragrafo %d) por um valor' % (textos[i], i))

    return trocas


def taguear_data(paras, textos):
    """Troca o paragrafo da data por "{local}, {data}.".

    O modelo usa um campo TIME do Word, que se atualiza sozinho toda vez que o
    arquivo e aberto — o documento assinado ficaria com a data errada. Trocar o
    paragrafo inteiro por texto simples elimina o campo.
    """
    trocas = {}
    for i, p in enumerate(paras):
        if 'TIME \\@' in p:
            trocas[i] = '{local}, {data}.'
    if not trocas:
        raise LookupError('paragrafo da data (campo TIME) nao encontrado')
    return trocas


SUBOPCOES = [
    ('(_) urbana (_) rural',         '({u_urbana}) urbana ({u_rural}) rural'),
    ('(_) temporária (_) permanente', '({u_temporaria}) temporária ({u_permanente}) permanente'),
    ('(_) idoso (_) PCD',            '({u_idoso}) idoso ({u_pcd}) PCD'),
]


def taguear_subopcoes(paras, textos):
    trocas = {}
    for alvo, novo in SUBOPCOES:
        i = indice_por_texto(textos, alvo)
        trocas[i] = novo
    return trocas


def taguear_clausula_primeira(paras, textos):
    """Deixa o tipo de beneficio da CLAUSULA PRIMEIRA variavel."""
    trocas = {}
    for i, t in enumerate(textos):
        if t.startswith('CLÁUSULA PRIMEIRA:'):
            if 'benefício por incapacidade' not in t:
                raise ValueError('CLAUSULA PRIMEIRA sem o trecho esperado: %r' % t)
            trocas[i] = t.replace('benefício por incapacidade', '{objeto_contrato}')
    if not trocas:
        raise LookupError('CLAUSULA PRIMEIRA nao encontrada')
    return trocas


def inserir_honorarios_extra(xml):
    """Acrescenta uma alinea opcional na CLAUSULA QUARTA (honorarios).

    Pedido do escritorio: as duas alineas fixas cobrem incapacidade temporaria e
    o grupo aposentadoria/pensao/BPC/incapacidade permanente/auxilio-acidente.
    Quando o cliente contrata algo fora disso (salario-maternidade, auxilio-
    reclusao, retificacao de CTC), precisa de um item escrito na hora.

    Entra depois da alinea b) e antes do "Tais valores serao pagos...", herdando
    a formatacao de lista (letras a, b, c). Vem envolvido num bloco condicional:
    com o campo vazio, o docxtemplater apaga os tres paragrafos e o contrato sai
    identico ao modelo.
    """
    for p in RE_PARA.findall(xml):
        if texto_de(p).startswith('Em caso de aposentadoria'):
            alvo = p
            break
    else:
        raise LookupError('alinea b) da CLAUSULA QUARTA nao encontrada')

    novos = ''.join(trocar_texto(alvo, t)
                    for t in ('{#honorarios_extra}', '{texto}', '{/honorarios_extra}'))
    pos = xml.index(alvo) + len(alvo)
    return xml[:pos] + novos + xml[pos:]


def marcar_lista_beneficios(numbering_xml):
    """Troca o marcador da lista de beneficios de "o" (vazio) para "X" (marcado).

    Decisao do escritorio: o Termo de Representacao sai com os 12 beneficios ja
    marcados, dando poderes amplos perante o INSS. So as sub-opcoes (urbana /
    rural etc.) e que variam conforme o formulario.
    """
    novo, trocas = re.subn(
        r'(<w:abstractNum w:abstractNumId="0"[ >].*?<w:lvl w:ilvl="0"[ >].*?)'
        r'<w:lvlText w:val="o"/>',
        r'\1<w:lvlText w:val="X"/>',
        numbering_xml, count=1, flags=re.S)
    if not trocas:
        raise LookupError('marcador da lista de beneficios nao encontrado')
    return novo


def aplicar(xml, funcoes):
    paras = RE_PARA.findall(xml)
    textos = [texto_de(p) for p in paras]

    trocas = {}
    for f in funcoes:
        for i, novo in f(paras, textos).items():
            trocas[i] = novo

    # Substitui de tras pra frente pra nao invalidar as posicoes ainda nao usadas.
    for i in sorted(trocas, reverse=True):
        antigo = paras[i]
        pos = xml.index(antigo)
        xml = xml[:pos] + trocar_texto(antigo, trocas[i]) + xml[pos + len(antigo):]

    return xml, len(trocas)


def construir(entrada, saida):
    nome = os.path.basename(saida)
    funcoes = [taguear_bloco_cliente, taguear_data]
    if nome.startswith('4-'):
        funcoes.append(taguear_subopcoes)
    if nome.startswith('5-'):
        funcoes.append(taguear_clausula_primeira)

    origem = zipfile.ZipFile(entrada)
    doc, trocas = aplicar(origem.read('word/document.xml').decode('utf-8'), funcoes)
    if nome.startswith('5-'):
        doc = inserir_honorarios_extra(doc)

    with zipfile.ZipFile(saida, 'w', zipfile.ZIP_DEFLATED) as destino:
        for item in origem.infolist():
            if item.filename == 'word/document.xml':
                destino.writestr(item, doc.encode('utf-8'))
            elif item.filename == 'word/numbering.xml' and nome.startswith('4-'):
                numbering = origem.read(item.filename).decode('utf-8')
                destino.writestr(item, marcar_lista_beneficios(numbering).encode('utf-8'))
            else:
                destino.writestr(item, origem.read(item.filename))
    origem.close()

    tags = sorted(set(re.findall(r'\{(\w+)\}', doc)))
    print('  %-22s %2d paragrafos tagueados  ->  %s' % (nome, trocas, ', '.join(tags)))


def main():
    if not os.path.isdir(ORIGEM):
        sys.exit('pasta de modelos nao encontrada: %s' % ORIGEM)
    if os.path.isdir(DESTINO):
        shutil.rmtree(DESTINO)
    os.makedirs(DESTINO)

    print('Gerando templates em %s' % os.path.relpath(DESTINO, RAIZ))
    for modelo, saida in MODELOS:
        entrada = os.path.join(ORIGEM, modelo)
        if not os.path.isfile(entrada):
            sys.exit('modelo faltando: %s' % entrada)
        construir(entrada, os.path.join(DESTINO, saida))
    print('OK — %d templates gerados.' % len(MODELOS))


if __name__ == '__main__':
    main()
