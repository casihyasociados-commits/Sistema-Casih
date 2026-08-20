# -*- coding: utf-8 -*-
"""
Normaliza el modelo de telegrama: deja TODO el formulario anclado a la pagina.

El modelo original tiene el formulario partido en tres sistemas de
coordenadas: una parte anclada a la pagina (filas 3 y 4, renglones internos,
pie) y otras dos ancladas a parrafos (los banners con las filas 1 y 2, y el
recuadro del cuerpo). Como la posicion de un parrafo depende de como cada
programa calcule el alto del texto anterior, esas partes caen distinto en
Word, Pages y Google Docs: el encabezado se corre y el recuadro se sube 2 cm
tapando la fila de Localidad/Provincia.

Al pasar todo a coordenadas de pagina, el formulario queda igual en cualquier
programa y los datos siempre caen sobre su renglon.

Los desplazamientos se calcularon para que las cuatro filas queden
equiespaciadas 1,02 cm (que es la separacion real entre las filas 3 y 4, las
unicas que ya estaban ancladas a la pagina) y para que el recuadro arranque
debajo de las etiquetas de Localidad/Provincia.
"""
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
EMU = 360000.0
MARGEN_IZQ_CM = 0.88          # 500 twips

# Desplazamiento vertical a aplicar segun el parrafo que contiene el ancla.
# par 2  -> banners + filas 1 y 2  (fila 1 pasa de 2.04 a 4.14)
# par 14 -> recuadro del cuerpo    (pasa de 0.07 a 8.60)
DESPLAZAMIENTO = {2: 2.10, 14: 8.53}


def normalizar(entrada, salida):
    zin = zipfile.ZipFile(entrada)
    doc = zin.read('word/document.xml').decode('utf-8')

    # Recorremos el XML por parrafo de nivel superior para saber a cual
    # pertenece cada ancla, y reescribimos solo las que no son de pagina.
    root = ET.fromstring(doc.encode('utf-8'))
    body = root.find(W + 'body')

    # Mapa: (posOffset original, eje) -> nuevo valor, por parrafo
    cambios = []
    for i, p in enumerate(list(body)):
        if p.tag != W + 'p':
            continue
        desp = DESPLAZAMIENTO.get(i)
        for a in p.iter(WP + 'anchor'):
            ph = a.find(WP + 'positionH')
            pv = a.find(WP + 'positionV')
            if ph is not None and (ph.get('relativeFrom') or '') != 'page':
                o = ph.find(WP + 'posOffset')
                viejo = int(o.text) if o is not None and o.text else 0
                cambios.append(('H', viejo, viejo + int(MARGEN_IZQ_CM * EMU)))
            if pv is not None and (pv.get('relativeFrom') or '') != 'page' and desp is not None:
                o = pv.find(WP + 'posOffset')
                viejo = int(o.text) if o is not None and o.text else 0
                cambios.append(('V', viejo, viejo + int(desp * EMU)))

    # Reescritura sobre el texto: reemplazamos cada bloque positionH/V que no
    # sea "page" por su equivalente en coordenadas de pagina. Se hace con una
    # pasada por parrafo para aplicar el desplazamiento correcto a cada uno.
    partes = re.split(r'(<w:p(?=[ >/]))', doc)
    # Reconstruimos separando por parrafos de nivel superior usando el conteo
    # de anidamiento (el formulario tiene parrafos dentro de cajas de texto).
    salida_xml = []
    idx_par = -1
    nivel = 0
    i = 0
    token = re.compile(r'<w:p(?=[ >/])|</w:p>')
    pos = 0
    trozos = []
    for m in token.finditer(doc):
        if m.group(0) == '</w:p>':
            nivel -= 1
            if nivel == 0:
                trozos.append((idx_par, pos, m.end()))
                pos = m.end()
        else:
            if nivel == 0:
                idx_par += 1
            nivel += 1
    # cola final
    if pos < len(doc):
        trozos.append((None, pos, len(doc)))

    def reescribir(fragmento, desp_v):
        def repl(m):
            eje = m.group('eje')
            rel = m.group('rel')
            val = int(m.group('val'))
            if rel == 'page':
                return m.group(0)
            if eje == 'H':
                nuevo = val + int(MARGEN_IZQ_CM * EMU)
            else:
                if desp_v is None:
                    return m.group(0)
                nuevo = val + int(desp_v * EMU)
            return ('<wp:position%s relativeFrom="page"><wp:posOffset>%d</wp:posOffset>'
                    '</wp:position%s>' % (eje, nuevo, eje))
        patron = re.compile(
            r'<wp:position(?P<eje>[HV]) relativeFrom="(?P<rel>[^"]+)">'
            r'<wp:posOffset>(?P<val>-?\d+)</wp:posOffset>'
            r'</wp:position(?P=eje)>')
        return patron.sub(repl, fragmento)

    for idx, ini, fin in trozos:
        salida_xml.append(reescribir(doc[ini:fin], DESPLAZAMIENTO.get(idx)))
    doc = ''.join(salida_xml)

    with zipfile.ZipFile(salida, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = doc.encode('utf-8')
            zout.writestr(item, data)
    print('normalizado ->', salida)
    print('anclas reubicadas:', len(cambios))


if __name__ == '__main__':
    normalizar(sys.argv[1], sys.argv[2])
