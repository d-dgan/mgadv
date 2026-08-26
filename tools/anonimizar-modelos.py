#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Troca os dados de cliente real dos modelos por um cliente ficticio.

Os .docx que a Isabelle usa no dia a dia vem preenchidos com o ultimo cliente
atendido — nome, CPF, telefone e endereco de uma pessoa de verdade. Como
assets/ e publicado no GitHub Pages, esses arquivos nao podem ir pro repositorio
com dado real.

O build (build-templates.py) troca exatamente esses paragrafos por tags {tag},
entao o conteudo ficticio nunca aparece em documento gerado — ele so existe pra
o modelo continuar legivel e editavel no Word.

Uso:  python tools/anonimizar-modelos.py           (mostra o que seria trocado)
      python tools/anonimizar-modelos.py --aplicar (grava)
"""

import os
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from importlib import import_module

build = import_module('build-templates')

FICTICIO = {
    'nome':          'Fulano de Tal',
    'nacionalidade': 'Brasileiro(a)',
    'estado_civil':  'Casado(a)',
    'profissao':     'Aposentado',
    'cpf':           '000.000.000-00',
    'celular':       '(00) 00000-0000',
    'endereco':      'Rua Exemplo, nº 100, centro, Siqueira Campos/PR',
}


def valores_ficticios(paras, textos):
    """Mesmo mapeamento do build, mas escrevendo o valor ficticio no lugar da tag."""
    trocas = build.taguear_bloco_cliente(paras, textos)
    return {i: FICTICIO[tag.strip('{}')] for i, tag in trocas.items()}


def main():
    aplicar = '--aplicar' in sys.argv

    for modelo, _ in build.MODELOS:
        caminho = os.path.join(build.ORIGEM, modelo)
        origem = zipfile.ZipFile(caminho)
        xml = origem.read('word/document.xml').decode('utf-8')

        paras = build.RE_PARA.findall(xml)
        textos = [build.texto_de(p) for p in paras]
        novos = valores_ficticios(paras, textos)

        print(modelo)
        for i in sorted(novos):
            print('   %-52r ->  %r' % (textos[i], novos[i]))

        if aplicar:
            novo, _ = build.aplicar(xml, [valores_ficticios])
            temporario = caminho + '.tmp'
            with zipfile.ZipFile(temporario, 'w', zipfile.ZIP_DEFLATED) as destino:
                for item in origem.infolist():
                    if item.filename == 'word/document.xml':
                        destino.writestr(item, novo.encode('utf-8'))
                    else:
                        destino.writestr(item, origem.read(item.filename))
            origem.close()
            shutil.move(temporario, caminho)
        else:
            origem.close()

    print()
    print('gravado.' if aplicar else 'nada gravado — rode com --aplicar pra valer.')


if __name__ == '__main__':
    main()
