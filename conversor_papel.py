#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
conversor_papel.py  v4
Convierte PDF o DOCX de unidad online IC Grupo → DOCX papel editorial.
Sin IA. Reglas 100% deterministas.

Cambios v4:
  - URL de imagen como marcador [IMG] con estilo propio, nunca como párrafo normal
  - Opciones tipo test en párrafos separados con estilo EjerciciosOpciones (no duplicar letras)
  - Desplegables como viñetas con título negrita; sublistas como nivel 2
  - Instrucciones digitales eliminadas
  - Bloques especiales en estilos editoriales correctos
  - Importante → estilo Importante (no Nota/Ejemplos)
  - Pies e imagen con estilos Pie-de-imagen y Descripcion-imagen
  - Actividades y tareas en bloques diferenciados (EjerciciosPregunta / EjerciciosOpciones)
  - Solución y Feedback solo en versión docente (marcados pero no emitidos)
  - Fórmulas detectadas y con estilo Formula
  - Residuos de mapa conceptual eliminados / maquetados

Uso: python conversor_papel.py UNIDAD.pdf  EJEMPLO_MAQUETADO.docx [SALIDA.docx]
     python conversor_papel.py UNIDAD.docx EJEMPLO_MAQUETADO.docx [SALIDA.docx]

Dependencias: pip install pypdf python-docx
"""
import sys, re, zipfile
from pathlib import Path

# ── Expresiones regulares ─────────────────────────────────────────────────────
RE_SEC1     = re.compile(r'^(\d+)\.\s+(.+)')
RE_SEC2     = re.compile(r'^(\d+)\.(\d+)\s+(.+)')
RE_SEC3     = re.compile(r'^(\d+)\.(\d+)\.(\d+)\s+(.+)')
RE_ENCAB    = re.compile(r'^(MPCOM_|BFCOM_|MP$|BF$|UA$|COM_)')
RE_ENCAB2   = re.compile(r'^MPCOM_B_\d+\s+BFCOM_')
RE_TE       = re.compile(r'^TE(\d+)\s+Tarea de evaluación$')
RE_TAREA_A  = re.compile(r'^Tarea de evaluación \d+ asociada')
RE_COLAB    = re.compile(r'^Actividad colaborativa (\d+)')
# Texto con espaciado entre caracteres (desplegables del PDF)
RE_ESPAC    = re.compile(r'^[A-Za-záéíóúüñÁÉÍÓÚÜÑ¿¡]'
                         r'(\s{1,2}[A-Za-záéíóúüñÁÉÍÓÚÜÑ,;:]){3,}'
                         r'(\s{1,2}[A-Za-záéíóúüñÁÉÍÓÚÜÑ,;:])?$')
RE_URL      = re.compile(r'^https?://')
RE_URL_PAR  = re.compile(r'^\(https?://[^\)]+\)$')
RE_NUM_SOLO = re.compile(r'^\d+$')
RE_MAPA     = re.compile(r'^\d+\.\d+\s+.{5,}')
# Caracteres de área de uso privado (iconos de fuente custom del PDF; solo iconos, no texto)
RE_PUA_ONLY = re.compile(r'^[\ue000-\uf8ff\U000f0000-\U000fffff\s]+$')
RE_PUA      = re.compile(r'[\ue000-\uf8ff\U000f0000-\U000fffff]+')
# Descripción corta tipo "X, Y, Z, etc."
RE_DESC_ETC = re.compile(r'.+,\s+etc\.?$', re.IGNORECASE)
# Título corto sin puntuación final (candidato a cabecera de desplegable inline)
# No puede empezar con minúscula ni ser demasiado largo
RE_TITULO_CORTO = re.compile(r'^[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñA-ZÁÉÍÓÚÜÑ\s]{3,55}$')
# Palabras que al final de línea indican que el título continúa en la siguiente
PREPOSICIONES_CONTINUACION = frozenset([
    'en', 'de', 'y', 'e', 'la', 'el', 'los', 'las', 'a', 'al',
    'del', 'por', 'con', 'o', 'u', 'su', 'sus',
])
# Detección de fórmulas matemáticas / indicadores con operadores
RE_FORMULA  = re.compile(
    r'(?:'
    r'\d+\s*[×x\*/÷]\s*\d+'       # multiplicación o división
    r'|[A-Za-z]+\s*=\s*.+[/÷×]'   # asignación con operador
    r'|\b\w+\s*/\s*\w+\s*[×x]\s*100'  # ratio × 100
    r')'
)
# Opción tipo test: a) texto, b) texto …
RE_OPCION   = re.compile(r'^([a-h])\)\s+(.+)')

BLOQUES_ESP = {
    'Nota', 'Ejemplo', 'Sabías que...', 'Sabías que…',
    'Consejo', 'Definición', 'Hilo conductor',
    'Para saber más', 'Vídeo', 'Importante', 'Recuerda',
}

# Instrucciones digitales que deben eliminarse completamente
INSTRUCCIONES_DIGITALES = frozenset([
    'Instrucción:', 'Haz clic', 'Pulsa en ', 'Pulsa para ',
    'Avanza para ', 'Haz clic ', 'Pincha ', 'Cambio de pantalla',
    'Instrucción', 'Clic para',
])

PREFIJOS_ELIM = (
    'Para realizar las Actividades colaborativas',
    'Para realizar las Tareas de evaluación',
    'Las instrucciones para realizar la tarea',
    'Podrás compartir ',           # cualquier variante: "con el resto", "la solución", etc.
    'Podrás identificar las',
    'Es el momento de realizar la siguiente',
    'No obstante puedes seguir estudiando',
    'La duración aproximada de la misma',
    'Mapa conceptual o esquema de contenidos',
    'Para realizar las',
    'valoración será tenida en cuenta',
    'encontrarás la información necesaria',
    'otro momento que te sea más favorable',
    # Online navigation / digital instructions
    'Cambio de pantalla',
    'Pulsa en ',
    'Pulsa para ',
    'Avanza para ',
    'Haz clic ',
    'Pincha ',
    'Instrucción: ',
    'Haz clic para ',
    'Haz clic en ',
    # Residuos de actividades colaborativas online
    'Podrás debatir',
    'poder debatir y aportar',
    'En esta actividad colaborativa',
)

LISTA_NUM_PISTAS = (
    'Organizar la información',
    'Conservar los documentos',
    'Proteger la información',
    'Conservarse durante',
    'Destruirse mediante',
)

RESIDUOS_RESUMEN = {
    'Conservación o eliminación', 'Almacenamiento y custodia',
    'Uso y tramitación', 'Clasificación o registro', 'Creación o recepción',
    'Clasificación y registro',
}

# Archivos de estilo que se copian del ejemplo maquetado al DOCX de salida
ARCHIVOS_ESTILO = frozenset([
    'word/styles.xml',
    'word/stylesWithEffects.xml',
    'word/fontTable.xml',
    'word/settings.xml',
    'word/numbering.xml',
])

# Símbolos de viñeta por nivel (spec del prompt)
VINETA_SIM = {1: '\u25CF', 2: '\u25CB', 3: '\u25AA', 4: '\u2013'}
VINETA_EST = {1: 'Vietanvl11d', 2: 'Vietanvl21d',
              3: 'Vietanvl31d', 4: 'Vietanvl41d'}


# ── Helpers de texto ──────────────────────────────────────────────────────────

def limpiar_pua(t: str) -> str:
    """Elimina caracteres de área de uso privado (iconos de fuente del online)."""
    return RE_PUA.sub('', t).strip()

def es_solo_pua(t: str) -> bool:
    """True si la línea es solo iconos PUA (sin texto real)."""
    return bool(t) and bool(RE_PUA_ONLY.match(t))

def limpiar_titulo(t: str) -> str:
    t = re.sub(r'\s*\(CE\s+[a-z]\)\s*y\s*\(CE\s+[a-z]\)', '', t)
    t = re.sub(r'\s*\(CE\s+[a-z…]+\)', '', t)
    t = re.sub(r'\s*…+$', '', t)      # eliminar elipsis al final
    t = re.sub(r'\s*\.{2,}$', '', t)  # eliminar "..." al final
    t = t.rstrip('…. ').strip()
    return t

def es_titulo_seccion(texto: str) -> bool:
    for p in LISTA_NUM_PISTAS:
        if texto.startswith(p):
            return False
    if len(texto) > 100:
        return False
    return True

def es_espaciado(l: str) -> bool:
    return bool(RE_ESPAC.match(l))

def quitar_espaciado(l: str) -> str:
    """Reconstruye texto con espaciado inter-carácter del PDF."""
    palabras = re.split(r'\s{2,}', l)
    return ' '.join(re.sub(r'\s', '', w) for w in palabras).strip()

def infinitivo_a_imperativo(t: str) -> str:
    PARES = [
        ('buscar ', 'Busca '), ('identificar ', 'Identifica '),
        ('analizar ', 'Analiza '), ('elaborar ', 'Elabora '),
        ('diseñar ', 'Diseña '), ('crear ', 'Crea '),
        ('realizar ', 'Realiza '), ('comparar ', 'Compara '),
        ('describir ', 'Describe '), ('explicar ', 'Explica '),
        ('completar ', 'Completa '), ('redactar ', 'Redacta '),
        ('investigar ', 'Investiga '), ('seleccionar ', 'Selecciona '),
        ('clasificar ', 'Clasifica '), ('calcular ', 'Calcula '),
        ('consultar ', 'Consulta '), ('revisar ', 'Revisa '),
    ]
    for inf, imp in PARES:
        if t.lower().startswith(inf):
            return imp + t[len(inf):]
    return t[0].upper() + t[1:] if t else t

def debe_elim(l: str) -> bool:
    if l in RESIDUOS_RESUMEN:
        return True
    return any(l.startswith(p) for p in PREFIJOS_ELIM)

def esc(t: str) -> str:
    return (t.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;'))


# ── Constructores XML ─────────────────────────────────────────────────────────

def p(texto, estilo, negrita=False):
    """Párrafo simple con un único run de texto."""
    te = esc(str(texto))
    sp = ' xml:space="preserve"' if te and te != te.strip() else ''
    if not te:
        return f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr></w:p>'
    b = '<w:rPr><w:b/><w:bCs/></w:rPr>' if negrita else ''
    return (f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
            f'<w:r>{b}<w:t{sp}>{te}</w:t></w:r></w:p>')

def p_vineta(texto, nivel=1):
    """Viñeta con símbolo Unicode correcto según nivel (1-4)."""
    estilo  = VINETA_EST.get(nivel, 'Vietanvl11d')
    simbolo = esc(VINETA_SIM.get(nivel, '\u25CF'))
    te      = esc(texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
            f'<w:r><w:t>{simbolo}</w:t></w:r>'
            f'<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>')

def p_vineta_ejemplo(texto):
    """Viñeta dentro de bloque Ejemplos-."""
    simbolo = esc(VINETA_SIM[1])
    te      = esc(texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="Ejemplos-Vietanvl1"/></w:pPr>'
            f'<w:r><w:t>{simbolo}</w:t></w:r>'
            f'<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>')

def p_desp(titulo, desc, nivel=1):
    """Desplegable: viñeta con título en negrita + descripción en el mismo párrafo."""
    estilo  = VINETA_EST.get(nivel, 'Vietanvl11d')
    simbolo = esc(VINETA_SIM.get(nivel, '\u25CF'))
    t, d    = esc(titulo), esc(desc)
    return (f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
            f'<w:r><w:t>{simbolo}</w:t></w:r>'
            f'<w:r><w:tab/></w:r>'
            f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
            f'<w:t xml:space="preserve">{t}. </w:t></w:r>'
            f'<w:r><w:t>{d}</w:t></w:r></w:p>')

def p_vineta_bold(texto, nivel=1):
    """Viñeta con título en negrita (sin descripción adjunta)."""
    estilo  = VINETA_EST.get(nivel, 'Vietanvl11d')
    simbolo = esc(VINETA_SIM.get(nivel, '\u25CF'))
    te      = esc(texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
            f'<w:r><w:t>{simbolo}</w:t></w:r>'
            f'<w:r><w:tab/></w:r>'
            f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>{te}</w:t></w:r></w:p>')

def p_url_imagen(url):
    """URL de imagen: marcador de fuente con estilo propio, sin prefijo [IMG]."""
    te = esc(url)
    return (f'    <w:p><w:pPr><w:pStyle w:val="Marcadorimagen"/></w:pPr>'
            f'<w:r><w:t>{te}</w:t></w:r></w:p>')

def p_url_recurso(url):
    """URL de recurso/enlace externo (no imagen): estilo URL propio."""
    te = esc(url)
    return (f'    <w:p><w:pPr><w:pStyle w:val="URLrecurso"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
            f'<w:t>{te}</w:t></w:r></w:p>')

def p_pie_imagen(texto):
    """Pie de imagen con estilo específico — incluye etiqueta 'Pie de imagen:'."""
    # La etiqueta forma parte del contenido: el estilo la colorea
    label = 'Pie de imagen: '
    if texto.startswith('Pie de imagen:'):
        te = esc(texto)          # ya incluye la etiqueta
    else:
        te = esc(label + texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="Piedeimagen"/></w:pPr>'
            f'<w:r><w:t>{te}</w:t></w:r></w:p>')

def p_desc_imagen(texto):
    """Descripción de imagen con estilo específico — incluye etiqueta."""
    label = 'Descripción de la imagen: '
    if re.match(r'^Descripci[oó]n de (la )?imagen:', texto):
        te = esc(texto)
    else:
        te = esc(label + texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="Descripcionimagen"/></w:pPr>'
            f'<w:r><w:t>{te}</w:t></w:r></w:p>')

def p_formula(texto):
    """Fórmula o cálculo destacado."""
    te = esc(texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="Formula"/></w:pPr>'
            f'<w:r><w:t>{te}</w:t></w:r></w:p>')

def p_opcion_test(letra, texto):
    """Opción de pregunta tipo test en párrafo separado con estilo propio."""
    te = esc(texto)
    return (f'    <w:p><w:pPr><w:pStyle w:val="EjerciciosOpciones"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{esc(letra)}) </w:t></w:r>'
            f'<w:r><w:t>{te}</w:t></w:r></w:p>')

def _es_url_imagen(url: str) -> bool:
    """Detecta si una URL es fuente de imagen (no recurso navegable)."""
    DOMINIOS_IMAGEN = ('shutterstock.com', 'gettyimages.', 'istockphoto.com',
                       'unsplash.com', 'pexels.com', 'freepik.com',
                       'pixabay.com', 'depositphotos.com', 'adobe.com/stock',
                       'stock.adobe.com')
    url_lower = url.lower()
    if any(d in url_lower for d in DOMINIOS_IMAGEN):
        return True
    # Extensión de imagen
    if re.search(r'\.(jpg|jpeg|png|gif|webp|svg|bmp|tiff?)(\?|$)', url_lower):
        return True
    return False


# ── Extracción de texto del PDF ───────────────────────────────────────────────

def extraer_texto(pdf: Path) -> list[str]:
    from pypdf import PdfReader
    r = PdfReader(str(pdf))
    return [page.extract_text() or '' for page in r.pages]


def extraer_texto_docx(docx: Path) -> list[str]:
    """
    Extrae texto de un DOCX de unidad online como lista de 'páginas'
    (simuladas por los saltos de página explícitos del documento).
    Cada párrafo se convierte en una línea; los saltos de página
    (<w:br w:type="page"/>) delimitan páginas nuevas.
    """
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(str(docx))
    paginas = []
    pagina_actual = []

    for para in doc.paragraphs:
        # Detectar salto de página dentro del párrafo
        tiene_salto = any(
            br.get(qn('w:type')) == 'page'
            for br in para._p.iter(qn('w:br'))
        )

        texto = para.text.strip()
        if texto:
            pagina_actual.append(texto)

        if tiene_salto:
            paginas.append('\n'.join(pagina_actual))
            pagina_actual = []

    if pagina_actual:
        paginas.append('\n'.join(pagina_actual))

    return paginas if paginas else ['\n'.join(pagina_actual)]


def extraer_fuente(ruta: Path) -> list[str]:
    """Detecta PDF o DOCX y llama al extractor adecuado."""
    suf = ruta.suffix.lower()
    if suf == '.pdf':
        return extraer_texto(ruta)
    elif suf in ('.docx', '.doc'):
        return extraer_texto_docx(ruta)
    else:
        raise ValueError(f'Formato no soportado: {suf}  (usa .pdf o .docx)')


# ── Reconstrucción de párrafos ────────────────────────────────────────────────

def reconstruir_parrafos(lineas_raw: list[str]) -> list[str]:
    """
    Une líneas partidas por el ancho de columna del PDF.
    Mejoras v3:
    - No vuelca el buffer ante línea vacía si la línea previa no cierra oración.
    - Tolera acumulación de 1 línea vacía interna sin romper párrafo incompleto.
    """
    resultado = []
    buf = ''
    vacios_pendientes = 0

    for l in lineas_raw:
        l_clean = limpiar_pua(l)  # quitar iconos de fuente

        if not l_clean:
            if buf:
                ultima = buf[-1]
                # Si la frase no termina en puntuación, aguantar 1 línea vacía
                if ultima not in '.?!:;…' and vacios_pendientes == 0:
                    vacios_pendientes += 1
                    continue
                # Más de 1 vacío consecutivo → flush definitivo
                resultado.append(buf)
                buf = ''
            vacios_pendientes = 0
            resultado.append('')
            continue

        vacios_pendientes = 0

        if not buf:
            buf = l_clean
            continue

        ultima  = buf[-1] if buf else ''
        primera = l_clean[0] if l_clean else ''

        termina = ultima in '.?!:;…'

        nueva_seccion = (
            RE_SEC1.match(l_clean) or RE_SEC2.match(l_clean) or
            l_clean in BLOQUES_ESP or RE_TE.match(l_clean) or
            RE_COLAB.match(l_clean) or RE_ENCAB.match(l_clean) or
            RE_ENCAB2.match(l_clean) or RE_NUM_SOLO.match(l_clean) or
            l_clean in ('Resumen', 'Introducción')
        )

        if not termina and not nueva_seccion and (primera.islower() or ultima == ','):
            buf = buf + ' ' + l_clean
        else:
            resultado.append(buf)
            buf = l_clean

    if buf:
        resultado.append(buf)
    return resultado


# ── Extracción de títulos completos del mapa ──────────────────────────────────

def extraer_mapa_titulos(todas: list[str]) -> dict[str, str]:
    """
    Extrae títulos completos de sección nivel 1 del mapa de contenidos
    (líneas anteriores a 'Introducción').

    Gestiona títulos partidos en 2 líneas por el ancho de columna del PDF:
    detecta cuando la línea anterior termina con preposición/artículo.
    Devuelve {prefijo_25_chars: titulo_limpio}.
    """
    titulos = {}
    buf = ''

    IGNORAR = {
        'UA', 'MP', 'BF', 'Unidad de aprendizaje', 'Objetivos',
        'Objetivos específicos:', 'Módulo profesional', 'Bloque formativo',
        'Mapa conceptual o esquema de contenidos',
        '¿Qué aprenderás en esta unidad? Criterios de evaluación',
        'Introducción',
    }

    for l in todas:
        if l == 'Introducción':
            if buf:
                t = limpiar_titulo(buf)
                if len(t) > 15:
                    titulos[t[:25].lower()] = t
            break

        if not l:
            continue
        if RE_ENCAB.match(l) or RE_ENCAB2.match(l):
            continue
        if l in IGNORAR or l.startswith('CE') or l.startswith('Se han'):
            continue
        if RE_NUM_SOLO.match(l) or RE_MAPA.match(l):
            # Número de sección o línea de nivel 2 → flush
            if buf:
                t = limpiar_titulo(buf)
                if len(t) > 15:
                    titulos[t[:25].lower()] = t
                buf = ''
            continue

        # ¿Es candidato a título de sección nivel 1?
        if len(l) < 15 or l[0].isdigit():
            continue
        if l.startswith('Identificar') or l.startswith('Aplicar'):
            continue  # son los objetivos de evaluación, no títulos

        # ¿La línea anterior termina con preposición/artículo? → continuación
        if buf:
            ultima_palabra = buf.split()[-1].lower()
            if ultima_palabra in PREPOSICIONES_CONTINUACION:
                buf = buf + ' ' + l
                continue
            else:
                # Nueva entrada: flush del anterior
                t = limpiar_titulo(buf)
                if len(t) > 15:
                    titulos[t[:25].lower()] = t
                buf = l
        else:
            buf = l

    return titulos


def completar_titulo(titulo: str, titulos_mapa: dict) -> str:
    """
    Si el título está truncado (termina con … o en minúsculas abruptamente),
    intenta encontrar la versión completa en el mapa.
    """
    truncado = (titulo.endswith('…') or titulo.endswith('...') or
                (titulo and titulo[-1].islower() and len(titulo) > 10))
    if not truncado:
        return titulo

    base = re.sub(r'[…\.]+$', '', titulo).strip()
    clave = base[:20].lower()

    # Búsqueda por prefijo
    for k, v in titulos_mapa.items():
        if k.startswith(clave) or clave.startswith(k[:15]):
            return v

    # Si no encontramos, devolver el limpio sin elipsis
    return base


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear(paginas: list[str]) -> dict:
    raw = []
    for pag in paginas:
        for l in pag.split('\n'):
            raw.append(l.strip().replace('\x00', ''))

    todas = reconstruir_parrafos(raw)

    # ── Unidad y módulo ──────────────────────────────────────────
    ua_num = ''
    ua_titulo = ''
    for i, l in enumerate(todas[:30]):
        if l == 'UA' and i + 1 < len(todas) and todas[i + 1].isdigit():
            ua_num = todas[i + 1]
            if i + 2 < len(todas) and todas[i + 2] == 'Unidad de aprendizaje':
                ua_titulo = todas[i + 3] if i + 3 < len(todas) else ''
            break

    # ── Objetivos: entre "Mapa conceptual" y "CE" ───────────────
    objetivos = []
    en_zona   = False
    buf_obj   = []
    for l in todas:
        if l == 'Mapa conceptual o esquema de contenidos':
            en_zona = True
            continue
        if en_zona:
            if l == 'CE' or l.startswith('CE a'):
                break
            if l and not l.startswith('CE'):
                buf_obj.append(l)
    j = 0
    while j < len(buf_obj):
        linea = buf_obj[j]
        while (j + 1 < len(buf_obj) and buf_obj[j + 1] and
               buf_obj[j + 1][0].islower()):
            j += 1
            linea += ' ' + buf_obj[j]
        objetivos.append(linea.strip())
        j += 1

    # ── Mapa de títulos completos ────────────────────────────────
    titulos_mapa = extraer_mapa_titulos(todas)

    # ── Mapa de contenidos (para saltarlo en el body) ────────────
    mapa_secciones = set()
    for l in todas:
        if l == 'Introducción':
            break
        m = RE_MAPA.match(l)
        if m:
            mapa_secciones.add(l)

    # ── Localizar inicio del contenido real ──────────────────────
    inicio = -1
    for i, l in enumerate(todas):
        if l == 'Introducción' and i > 5:
            prev = todas[i - 1] if i > 0 else ''
            if not RE_SEC2.match(prev) and '(CE' not in prev:
                inicio = i
                break
    if inicio == -1:
        for i, l in enumerate(todas):
            m = RE_SEC1.match(l)
            if m and i > 10 and es_titulo_seccion(m.group(2)):
                inicio = i
                break
    if inicio == -1:
        inicio = 0

    lineas = todas[inicio:]

    # ── Estado del parser ────────────────────────────────────────
    secciones  = []
    sec = sub = sub2 = None
    bloques    = None
    num_mayor  = 1

    def ns1(num, tit):
        nonlocal sec, sub, sub2, bloques, num_mayor
        # Intentar completar título truncado
        tit_full = completar_titulo(tit, titulos_mapa)
        sec = {'num': num, 'titulo': tit_full,
               'bloques': [], 'subsecciones': []}
        secciones.append(sec)
        sub = sub2 = None
        bloques   = sec['bloques']
        num_mayor = num

    def ns2(num, tit):
        nonlocal sub, sub2, bloques
        sub = {'num': num, 'titulo': tit, 'bloques': [], 'subsecciones': []}
        if sec:
            sec['subsecciones'].append(sub)
        sub2    = None
        bloques = sub['bloques']

    def ns3(num, tit):
        nonlocal sub2, bloques
        sub2 = {'num': num, 'titulo': tit, 'bloques': [], 'subsecciones': []}
        if sub:
            sub['subsecciones'].append(sub2)
        elif sec:
            sec['subsecciones'].append(sub2)
        bloques = sub2['bloques']

    def add(b):
        if bloques is not None:
            bloques.append(b)

    def ultimo_bloque():
        return bloques[-1] if bloques else None

    ns1(1, 'Introducción')
    i = 1 if (lineas and lineas[0] == 'Introducción') else 0

    while i < len(lineas):
        l = lineas[i]

        # ── Filtros de ruido ─────────────────────────────────────
        if not l:
            i += 1
            continue

        # Líneas compuestas SOLO de iconos PUA → descartar (son marcadores visuales del PDF)
        if es_solo_pua(l):
            i += 1
            continue

        if RE_ENCAB.match(l) or RE_ENCAB2.match(l):
            i += 1
            continue

        # Palabras sueltas que son residuos de encabezados de página
        if l in ('impresión', 'impresión.', 'organización', 'organización.'):
            # Solo eliminar si el párrafo anterior no termina con vocal
            ult = ultimo_bloque()
            if ult and ult.get('tipo') == 'parrafo':
                txt = ult['texto']
                if txt and txt[-1] not in '.?!:;…':
                    ult['texto'] = txt + ' ' + l
                    i += 1
                    continue
            i += 1
            continue

        # ── Saltar mapa de contenidos embebido ───────────────────
        if RE_NUM_SOLO.match(l):
            j = i + 1
            while j < len(lineas) and (
                RE_MAPA.match(lineas[j]) or
                RE_NUM_SOLO.match(lineas[j]) or
                lineas[j] == ''
            ):
                j += 1
            i = j
            continue

        # ── Unir párrafo inconcluso con la línea actual ──────────
        # (resuelve cortes por salto de página con mapa intermedio)
        ult = ultimo_bloque()
        if (ult and ult.get('tipo') == 'parrafo' and
                ult['texto'] and ult['texto'][-1] not in '.?!:;…' and
                l and l[0].islower() and
                not RE_SEC1.match(l) and not RE_SEC2.match(l) and
                not es_espaciado(l) and l not in BLOQUES_ESP):
            ult['texto'] += ' ' + l
            i += 1
            continue

        # ── Sección nivel 1 ──────────────────────────────────────
        m = RE_SEC1.match(l)
        if m and es_titulo_seccion(m.group(2)):
            ns1(int(m.group(1)) + 1, limpiar_titulo(m.group(2)))
            i += 1
            continue

        if l == 'Resumen':
            ns1(num_mayor + 1, 'Resumen')
            i += 1
            continue

        # ── Sección nivel 2 ──────────────────────────────────────
        m = RE_SEC2.match(l)
        if m:
            ns2(f'{int(m.group(1)) + 1}.{m.group(2)}',
                limpiar_titulo(m.group(3)))
            i += 1
            continue

        # ── Sección nivel 3 ──────────────────────────────────────
        m = RE_SEC3.match(l)
        if m:
            ns3(f'{int(m.group(1)) + 1}.{m.group(2)}.{m.group(3)}',
                limpiar_titulo(m.group(4)))
            i += 1
            continue

        # ── Lista numerada dentro de bloque (no sección) ─────────
        m = RE_SEC1.match(l)
        if m and not es_titulo_seccion(m.group(2)):
            add({'tipo': 'parrafo', 'texto': l})
            i += 1
            continue

        # ── Tarea de evaluación ──────────────────────────────────
        m = RE_TE.match(l)
        if m:
            num_te = m.group(1)
            i += 1
            while i < len(lineas):
                nl = lineas[i]
                if RE_TAREA_A.match(nl) or nl in ('Objetivo:', 'Enunciado:', '') or debe_elim(nl):
                    i += 1
                    continue
                break
            cont = []
            while i < len(lineas):
                nl = lineas[i]
                m2 = RE_SEC1.match(nl)
                if m2 and es_titulo_seccion(m2.group(2)):
                    break
                if RE_SEC2.match(nl) or (RE_TE.match(nl) and cont):
                    break
                if RE_COLAB.match(nl):
                    break
                if RE_ENCAB.match(nl) or RE_ENCAB2.match(nl):
                    i += 1
                    continue
                if debe_elim(nl) or nl in ('Objetivo:', 'Enunciado:', 'impresión'):
                    i += 1
                    continue
                if RE_TAREA_A.match(nl):
                    i += 1
                    continue
                if nl in objetivos:
                    i += 1
                    continue
                if nl:
                    cont.append(nl)
                i += 1
            add({'tipo': 'tarea', 'etiqueta': f'Tarea {num_te}', 'lineas': cont})
            continue

        # ── Actividad colaborativa ───────────────────────────────
        m = RE_COLAB.match(l)
        if m:
            i += 1
            cont = []
            while i < len(lineas):
                nl = lineas[i]
                m2 = RE_SEC1.match(nl)
                if m2 and es_titulo_seccion(m2.group(2)):
                    break
                if RE_SEC2.match(nl) or RE_TE.match(nl) or RE_COLAB.match(nl):
                    break
                if RE_ENCAB.match(nl) or RE_ENCAB2.match(nl):
                    i += 1
                    continue
                if debe_elim(nl):
                    i += 1
                    continue
                if nl.startswith('En esta actividad deberás '):
                    cont.append(
                        infinitivo_a_imperativo(
                            nl[len('En esta actividad deberás '):]
                        )
                    )
                    i += 1
                    continue
                if nl:
                    cont.append(nl)
                i += 1
            add({'tipo': 'actividad_complementaria',
                 'etiqueta': 'Actividad complementaria', 'lineas': cont})
            continue

        # ── Bloque Vídeo ─────────────────────────────────────────
        if l == 'Vídeo':
            i += 1
            cont = []
            while i < len(lineas):
                nl = lineas[i]
                m2 = RE_SEC1.match(nl)
                if m2 and es_titulo_seccion(m2.group(2)):
                    break
                if RE_SEC2.match(nl):
                    break
                if nl in BLOQUES_ESP and cont:
                    break
                if RE_TE.match(nl) or RE_COLAB.match(nl):
                    break
                if RE_ENCAB.match(nl) or RE_ENCAB2.match(nl):
                    i += 1
                    continue
                if nl == '':
                    i += 1
                    continue
                if RE_URL_PAR.match(nl):
                    cont.append(nl[1:-1])
                    i += 1
                    continue
                cont.append(nl)
                i += 1
            add({'tipo': 'video', 'etiqueta': 'Vídeo', 'lineas': cont})
            continue

        # ── Bloques especiales (Nota, Ejemplo, Consejo…) ─────────
        if l in BLOQUES_ESP:
            etiq = l
            i += 1
            TIPO_MAP = {
                'Nota':           'nota',
                'Ejemplo':        'ejemplo',
                'Sabías que...':  'sabias_que',
                'Sabías que…':    'sabias_que',
                'Consejo':        'consejo',
                'Definición':     'definicion',
                'Hilo conductor': 'hilo_conductor',
                'Para saber más': 'para_saber_mas',
                'Vídeo':          'video',
                'Importante':     'importante',
                'Recuerda':       'recuerda',
            }
            cont   = []
            vacios = 0
            while i < len(lineas):
                nl = lineas[i]
                m2 = RE_SEC1.match(nl)
                if m2 and es_titulo_seccion(m2.group(2)):
                    break
                if RE_SEC2.match(nl):
                    break
                if nl in BLOQUES_ESP and cont:
                    break
                if RE_TE.match(nl) or RE_COLAB.match(nl):
                    break
                if RE_ENCAB.match(nl) or RE_ENCAB2.match(nl):
                    i += 1
                    continue
                if debe_elim(nl):
                    i += 1
                    continue
                if nl == '':
                    vacios += 1
                    if vacios > 1:
                        break
                    i += 1
                    continue
                vacios = 0
                cont.append(nl)
                i += 1
            tipo = TIPO_MAP.get(etiq, 'ejemplo')
            add({'tipo': tipo, 'etiqueta': etiq, 'lineas': cont})
            continue

        # ── Desplegable con espaciado inter-carácter ─────────────
        if es_espaciado(l):
            tit = quitar_espaciado(l)
            # El bloque anterior en el PDF es la descripción del desplegable
            # (el PDF extrae primero el cuerpo, luego el título del cuadro)
            desc = ''
            if bloques and bloques[-1].get('tipo') == 'parrafo':
                prev = bloques[-1]['texto']
                # Solo usar como descripción si es texto corto-medio (< 250 chars)
                if len(prev) < 250:
                    desc = prev
                    bloques.pop()
            add({'tipo': 'desplegable', 'titulo': tit, 'descripcion': desc})
            i += 1
            continue

        # ── Desplegable inline título+descripción ─────────────────
        # Patrón: título corto (sin punt. final) + descripción tipo "X, Y, Z, etc."
        # Ej: "Generación de documentos oficiales" + "Facturas, contratos, etc."
        sig = lineas[i + 1] if i + 1 < len(lineas) else ''
        if (sig and RE_DESC_ETC.match(sig) and RE_TITULO_CORTO.match(l) and
                not l.startswith(('La ', 'El ', 'Los ', 'Las ', 'Un ', 'Una ',
                                   'En ', 'A ', 'De ', 'Por ')) and
                not RE_SEC1.match(l) and not RE_SEC2.match(l) and
                l not in BLOQUES_ESP and not debe_elim(l)):
            add({'tipo': 'desplegable', 'titulo': l, 'descripcion': sig})
            i += 2
            continue

        if debe_elim(l):
            i += 1
            continue
        if l in RESIDUOS_RESUMEN:
            i += 1
            continue

        # ── Imagen ───────────────────────────────────────────────
        # Línea de pie + línea de descripción
        if (l.startswith(('La ', 'El ', 'Una ', 'Los ', 'Las ')) and
                len(l) < 150 and
                sig.startswith(('En la imagen', 'En el gráfico'))):
            add({'tipo': 'imagen', 'pie': l, 'descripcion': sig})
            i += 2
            continue

        if l.startswith(('En la imagen', 'En el gráfico')):
            add({'tipo': 'imagen', 'pie': '', 'descripcion': l})
            i += 1
            continue

        # Pie de imagen / descripción explícitos (líneas con prefijo)
        if l.startswith('Pie de imagen:'):
            add({'tipo': 'pie_imagen', 'texto': l[len('Pie de imagen:'):].strip()})
            i += 1
            continue
        if l.startswith('Descripción de la imagen:') or l.startswith('Descripción de imagen:'):
            add({'tipo': 'desc_imagen', 'texto': re.sub(r'^Descripci[oó]n de (la )?imagen:\s*', '', l)})
            i += 1
            continue

        # ── URLs ─────────────────────────────────────────────────
        # URL de imagen (Shutterstock, etc.) → marcador, no cuerpo
        if RE_URL.match(l):
            url = l
            if RE_URL_PAR.match(l):
                url = l[1:-1]
            # Detectar si es URL de imagen por dominio conocido o extensión
            if _es_url_imagen(url):
                add({'tipo': 'url_imagen', 'url': url})
            else:
                add({'tipo': 'url', 'url': url})
            i += 1
            continue
        if RE_URL_PAR.match(l):
            url = l[1:-1]
            if _es_url_imagen(url):
                add({'tipo': 'url_imagen', 'url': url})
            else:
                add({'tipo': 'url', 'url': url})
            i += 1
            continue

        # ── Párrafo de cuerpo ────────────────────────────────────
        add({'tipo': 'parrafo', 'texto': l})
        i += 1

    return {
        'titulo_unidad': 'Unidad de aprendizaje ' + ua_num,
        'titulo_modulo': ua_titulo,
        'objetivos':     objetivos,
        'secciones':     secciones,
    }




# ── Parser de interacciones (archivo auxiliar) ────────────────────────────────

def parsear_interacciones(docx: Path) -> dict:
    """
    Parsea el archivo de interacciones y devuelve:
    {
      N: {'tipo': 'desplegables', 'items': [{'titulo':..,'contenido':[..],'multi':bool}]},
      N: {'tipo': 'opciones',     'opciones': [...], 'solucion': str, 'feedback': str},
    }
    """
    from docx import Document
    doc = Document(str(docx))
    result = {}

    for tbl in doc.tables:
        # Row 0: header "Interacción N" or "Interacción N. Texto"
        header = tbl.rows[0].cells[0].text.strip()
        m = re.match(r'Interacci[oó]n\s+(\d+)', header)
        if not m:
            continue
        n = int(m.group(1))
        raw = tbl.rows[1].cells[0].text if len(tbl.rows) > 1 else ''

        if raw.strip().startswith('Opciones:'):
            # Ejercicio tipo test
            opciones = []
            solucion = ''
            feedback = ''
            zona = 'opciones'
            for line in raw.split('\n'):
                line = line.strip()
                if not line:
                    continue
                if line == 'Opciones:':
                    zona = 'opciones'
                elif re.match(r'Soluci[oó]n', line):
                    zona = 'solucion'
                    rest = line.split(':', 1)[1].strip() if ':' in line else ''
                    if rest:
                        solucion = rest
                elif line.startswith('Feedback:'):
                    zona = 'feedback'
                    rest = line.split(':', 1)[1].strip() if ':' in line else ''
                    if rest:
                        feedback = rest
                elif zona == 'opciones':
                    # Evitar incluir la letra si ya viene prefijada (a) b) c) d))
                    m_opt = re.match(r'^[a-h]\)\s*(.+)', line)
                    if m_opt:
                        opciones.append(m_opt.group(1).strip())
                    else:
                        opciones.append(line)
                elif zona == 'solucion':
                    solucion = (solucion + ' ' + line).strip()
                elif zona == 'feedback':
                    feedback = (feedback + ' ' + line).strip()
            result[n] = {
                'tipo': 'opciones',
                'opciones': opciones,
                'solucion': solucion,    # solo versión docente
                'feedback': feedback,    # solo versión docente
            }
        elif 'Desplegables:' in raw or 'Instrucción:' in raw:
            # Eliminar línea de instrucción
            if 'Desplegables:' in raw:
                raw_items = raw.split('Desplegables:', 1)[1].strip()
            else:
                raw_items = raw.strip()

            multi_mode = '\n\n' in raw_items
            items = []

            if multi_mode:
                blocks = [b.strip() for b in raw_items.split('\n\n') if b.strip()]
                i = 0
                while i < len(blocks):
                    titulo = blocks[i]
                    contenido = []
                    if i + 1 < len(blocks):
                        contenido = [l.strip() for l in blocks[i+1].split('\n') if l.strip()]
                        i += 2
                    else:
                        i += 1
                    items.append({'titulo': titulo, 'contenido': contenido, 'multi': True})
            else:
                # Título: línea corta (≤65 chars), sin puntuación final, no en minúscula
                def _es_titulo_desp(l):
                    if not l or len(l) > 65:
                        return False
                    if l[-1] in '.,;:':
                        return False
                    if l[0].islower():
                        return False
                    if re.match(r'^(Por ejemplo|Asimismo|Puede |También |Además |Se calcula|Número )', l):
                        return False
                    return True

                lines = [l.strip() for l in raw_items.split('\n') if l.strip()]
                i = 0
                while i < len(lines):
                    if not _es_titulo_desp(lines[i]):
                        i += 1
                        continue
                    titulo = lines[i]
                    i += 1
                    contenido_lines = []
                    while i < len(lines) and not _es_titulo_desp(lines[i]):
                        contenido_lines.append(lines[i])
                        i += 1
                    contenido = ' '.join(contenido_lines)
                    items.append({'titulo': titulo, 'contenido': [contenido], 'multi': False})

            result[n] = {'tipo': 'desplegables', 'items': items}

    return result


def parsear_interacciones_texto(path: Path) -> dict:
    """
    Parsea el archivo de interacciones en formato texto/markdown (tablas pipe).
    Misma salida que parsear_interacciones().
    """
    text = path.read_text(encoding='utf-8', errors='replace')
    result = {}

    # Cada interacción es un bloque de tabla separado por línea en blanco
    bloques = re.split(r'\n\n+', text)
    n_actual = None
    contenido_actual = ''

    for bloque in bloques:
        lines = [l.strip() for l in bloque.strip().split('\n') if l.strip()]
        if not lines:
            continue

        # Buscar cabecera "Interacción N"
        for line in lines:
            # Limpiar markdown de negrita: **texto**
            clean = re.sub(r'\*+', '', line).strip('| ').strip()
            m = re.match(r'Interacci[oó]n\s+(\d+)', clean)
            if m:
                n_actual = int(m.group(1))
                contenido_actual = ''
                break
        else:
            # No hay cabecera → es contenido de la interacción anterior
            if n_actual is not None:
                for line in lines:
                    if line == '| --- |' or line == '---':
                        continue
                    # Es una línea de contenido (celda de tabla)
                    cell = line.strip('| ').strip()
                    # Limpiar links markdown [texto](url) → url
                    cell = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\2', cell)
                    if cell and not re.match(r'^-+$', cell):
                        contenido_actual = (contenido_actual + ' ' + cell).strip()
            continue

        # Buscar el contenido en el mismo bloque (fila 3 de la tabla)
        content_lines = []
        for line in lines:
            if '| --- |' in line or line == '---':
                continue
            clean_line = line.strip('| ').strip()
            clean_line = re.sub(r'\*+', '', clean_line)
            clean_line = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\2', clean_line)
            m2 = re.match(r'Interacci[oó]n\s+\d+', clean_line)
            if m2:
                continue
            if clean_line:
                content_lines.append(clean_line)

        if content_lines:
            contenido_actual = ' '.join(content_lines)

        if n_actual is not None and contenido_actual:
            _procesar_contenido_interaccion(n_actual, contenido_actual, result)

    return result


def _procesar_contenido_interaccion(n: int, raw: str, result: dict):
    """Procesa el contenido de una celda de interacción y lo añade a result."""
    raw = raw.strip()

    # Eliminar prefijo de instrucción digital
    raw = re.sub(r'^Instrucci[oó]n:[^.]+\.\s*', '', raw)
    raw = re.sub(r'^Haz clic para[^.]+\.\s*', '', raw)

    if re.match(r'Opciones:', raw):
        # Tipo test
        opciones = []
        solucion = ''
        feedback = ''
        # Dividir por palabras clave con regex
        partes = re.split(r'\b(Soluci[oó]n:|Feedback:)', raw)
        zona_actual = 'opciones'
        for parte in partes:
            if re.match(r'Soluci[oó]n:', parte):
                zona_actual = 'solucion'
                continue
            if parte == 'Feedback:':
                zona_actual = 'feedback'
                continue
            if zona_actual == 'opciones':
                texto_opts = re.sub(r'^Opciones:\s*', '', parte).strip()
                if not texto_opts:
                    continue
                # Intentar separar por letra explícita: a) b) c) d)
                opts_con_letra = re.split(r'\s+(?=[a-h]\)\s)', texto_opts)
                if len(opts_con_letra) > 1:
                    for opt in opts_con_letra:
                        opt = opt.strip()
                        if not opt:
                            continue
                        m_opt = re.match(r'^([a-h])\)\s*(.+)', opt)
                        if m_opt:
                            opciones.append(m_opt.group(2).strip())
                        else:
                            opciones.append(opt)
                else:
                    # Sin letras → separar por ". " seguido de mayúscula
                    # Heurística: cada opción es una oración independiente
                    frases = re.split(r'(?<=\.)\s+(?=[A-ZÁÉÍÓÚÜÑ])', texto_opts)
                    opciones = [f.strip() for f in frases if f.strip()]
            elif zona_actual == 'solucion':
                solucion = (solucion + ' ' + parte).strip()
            elif zona_actual == 'feedback':
                feedback = (feedback + ' ' + parte).strip()
        result[n] = {
            'tipo': 'opciones',
            'opciones': opciones,
            'solucion': solucion,    # solo versión docente
            'feedback': feedback,    # solo versión docente
        }

    elif 'Desplegables:' in raw:
        raw_items = raw.split('Desplegables:', 1)[1].strip()
        items = _parsear_desplegables_texto(raw_items)
        result[n] = {'tipo': 'desplegables', 'items': items}

    elif raw:
        # Sin Opciones ni Desplegables → tratar como bloque de texto
        pass


def _parsear_desplegables_texto(raw: str) -> list:
    """
    Extrae items de desplegable de texto plano.
    Heurística: línea corta sin puntuación final = título; resto = contenido.
    """
    items = []

    def _es_titulo(t: str) -> bool:
        if not t or len(t) > 80:
            return False
        if t[-1] in '.,;:':
            return False
        if t[0].islower():
            return False
        if re.match(r'^(Por ejemplo|Asimismo|Puede |También |Además |Se calcula|'
                    r'Número |Estándar|Urgente|Programada|Entrega|Recogida|'
                    r'Encuestas|Tiempo de|Tasa de|Retenci)', t):
            return False
        return True

    # Intentar separar por títulos conocidos
    # Estrategia: dividir el texto buscando palabras con mayúscula al inicio
    # que no continúan un párrafo
    tokens = raw.split()
    current_title = []
    current_body = []
    in_body = False

    # Reconstruir "palabras" agrupadas por mayúsculas al inicio
    sentences = re.split(r'(?<=[.!?])\s+', raw)
    # Si hay pocas frases → título simple
    if len(sentences) <= 2:
        # Todo es contenido de un solo desplegable
        words = raw.split()
        # Primera "oración" corta sin punt. = título
        for i, word in enumerate(words):
            if i > 0 and words[i-1][-1] in '.!?' and word[0].isupper():
                titulo = ' '.join(words[:i])
                resto  = ' '.join(words[i:])
                if _es_titulo(titulo):
                    items.append({'titulo': titulo, 'contenido': [resto], 'multi': False})
                    return items
        items.append({'titulo': raw[:50], 'contenido': [raw[50:].strip()], 'multi': False})
        return items

    # Múltiples frases → buscar patrones título+cuerpo
    # Patrón más robusto: dividir en "chunks" donde cada chunk empieza con mayúscula
    # y la frase anterior terminó en punto
    chunks = re.split(r'(?<=\.) (?=[A-ZÁÉÍÓÚÜÑ])', raw)

    titulo_actual = ''
    body_lines = []

    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        # ¿Es un título potencial? (corto, sin punt. final, mayúscula)
        if _es_titulo(chunk) and len(chunk.split()) <= 8:
            # Guardar el item anterior
            if titulo_actual:
                items.append({
                    'titulo': titulo_actual,
                    'contenido': [' '.join(body_lines)] if body_lines else [],
                    'multi': False,
                })
            titulo_actual = chunk
            body_lines = []
        else:
            if titulo_actual:
                body_lines.append(chunk)
            else:
                # Sin título aún: primera frase larga se parte
                words = chunk.split()
                # Tomar las primeras palabras como título si son pocas
                for cut in range(3, min(8, len(words))):
                    cand = ' '.join(words[:cut])
                    if _es_titulo(cand):
                        titulo_actual = cand
                        body_lines = [' '.join(words[cut:])]
                        break
                else:
                    titulo_actual = chunk[:40].rstrip()
                    body_lines = [chunk[40:].strip()] if len(chunk) > 40 else []

    if titulo_actual:
        items.append({
            'titulo': titulo_actual,
            'contenido': [' '.join(body_lines)] if body_lines else [],
            'multi': False,
        })

    return items if items else [{'titulo': raw[:60], 'contenido': [raw[60:]], 'multi': False}]


def expandir_interaccion(n: int, interacciones: dict) -> list:
    """
    Convierte una interacción en lista de bloques para bloques_xml.
    - Desplegables simples → desplegable_simple (título + contenido en misma viñeta)
    - Desplegables multi   → desplegable_multi  (título bold + sub-viñetas nivel 2)
    - Opciones             → lista de bloques vacía (el contexto se maneja externamente)
    - Instrucciones digitales (Instrucción: Haz clic...) → eliminadas
    """
    inter = interacciones.get(n, {})
    if not inter:
        return []

    if inter['tipo'] == 'opciones':
        return []  # gestionado externamente

    bloques = []
    for item in inter.get('items', []):
        titulo = item['titulo']
        if any(titulo.startswith(p) for p in PREFIJOS_ELIM):
            continue

        if item['multi']:
            bloques.append({
                'tipo': 'desplegable_multi',
                'titulo': titulo,
                'items': item['contenido'],
            })
        else:
            contenido = item['contenido'][0] if item['contenido'] else ''
            if contenido:
                bloques.append({
                    'tipo': 'desplegable_simple',
                    'titulo': titulo,
                    'contenido': contenido,
                })
            else:
                bloques.append({
                    'tipo': 'p_vineta_bold',
                    'texto': titulo,
                    'nivel': 1,
                })
    return bloques
    """
    Convierte una interacción en lista de bloques para bloques_xml.
    - Desplegables simples → desplegable_simple (título + contenido en misma viñeta)
    - Desplegables multi   → desplegable_multi  (título bold + sub-viñetas nivel 2)
    - Opciones             → lista de bloques vacía (el contexto se maneja en parsear_docx_fuente)
    - Instrucciones digitales (Instrucción: Haz clic...) → eliminadas
    """
    inter = interacciones.get(n, {})
    if not inter:
        return []

    if inter['tipo'] == 'opciones':
        return []  # gestionado externamente

    bloques = []
    for item in inter.get('items', []):
        # Filtrar instrucciones digitales del título
        titulo = item['titulo']
        if any(titulo.startswith(p) for p in PREFIJOS_ELIM):
            continue

        if item['multi']:
            # Sublistas → nivel 2
            sub_items = item['contenido']
            bloques.append({
                'tipo': 'desplegable_multi',
                'titulo': titulo,
                'items': sub_items,
            })
        else:
            contenido = item['contenido'][0] if item['contenido'] else ''
            if contenido:
                bloques.append({
                    'tipo': 'desplegable_simple',
                    'titulo': titulo,
                    'contenido': contenido,
                })
            else:
                bloques.append({
                    'tipo': 'p_vineta_bold',
                    'texto': titulo,
                    'nivel': 1,
                })
    return bloques


# ── Parser de DOCX fuente (basado en estilos de párrafo) ─────────────────────

def parsear_docx_fuente(docx: Path, interacciones: dict) -> dict:
    """
    Parsea el DOCX fuente aprovechando los estilos de párrafo.
    Devuelve la misma estructura que parsear() para ser compatible con generar_docx().
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph as DParagraph
    from docx.table import Table as DTable

    doc = Document(str(docx))

    ua_num    = ''
    ua_titulo = ''
    objetivos = []
    secciones = []

    current_sec  = None
    current_sub  = None
    current_sub2 = None
    en_objetivos = False
    sec_count    = 0
    blk          = None  # bloque multi-párrafo en construcción

    # Mapeo: estilo fuente → tipo bloque
    MULTI_STYLES = {
        'Hilo conductor':       'hilo_conductor',
        'Ejemplo':              'ejemplo',
        'Sabiasque':            None,   # dinámico según primer texto
        'Vídeo':                'video',
        'Actividad colaborativa': 'actividad_complementaria',
        'Importante':           'importante',
        'Aplicación práctica':  'tarea',
    }
    # Qué label usar como etiqueta de bloque
    LABEL_MAP = {
        'Sabías que…': 'sabias_que',
        'Definición':  'definicion',
        'Definición ': 'definicion',
        'Importante':  'importante',
    }

    def bloques_activos():
        if current_sub2:
            return current_sub2['bloques']
        if current_sub:
            return current_sub['bloques']
        if current_sec:
            return current_sec['bloques']
        return []

    def flush():
        nonlocal blk
        if blk and blk.get('lineas'):
            bloques_activos().append({k: v for k, v in blk.items() if k != '_estilo'})
        blk = None

    def add(b):
        flush()
        bloques_activos().append(b)

    SKIP_TEXT = {
        'Cambio de pantalla', 'Específicos', 'Objetivos',
        'Criterios de evaluación: ', '',
    }
    SKIP_BOLD = {'Cambio de pantalla'}
    SKIP_PREFIX = ('CE ', 'a) Se han', 'b) Se han', 'c) Se han', 'd) Se han',
                   'e) Se han', 'Duración:', 'Objetivo:', 'Enunciado:')

    # Itera elementos del body en orden (párrafos Y tablas)
    for child in doc.element.body:
        is_para = child.tag == qn('w:p')
        is_tbl  = child.tag == qn('w:tbl')

        if is_tbl:
            tbl = DTable(child, doc)
            flush()
            for row in tbl.rows:
                cell_txt = row.cells[0].text.strip() if row.cells else ''
                if cell_txt:
                    bloques_activos().append({'tipo': 'p_vineta', 'texto': cell_txt, 'nivel': 1})
            continue

        if not is_para:
            continue

        para  = DParagraph(child, doc)
        style = para.style.name
        txt   = para.text.strip()
        bold  = any(r.bold for r in para.runs if r.text.strip())

        # ── Título del documento y headings ANTES de filtros ───
        if style == 'Title':
            m = re.match(r'Unidad de aprendizaje\s+(\d+)[.\-–\s]\s*(.+)', txt)
            if m:
                ua_num    = m.group(1)
                ua_titulo = m.group(2).strip()
            continue

        # ── Heading 1 (secciones y zonas especiales) ─────────
        if style == 'Heading 1':
            flush()
            en_objetivos = False
            current_sub  = None
            current_sub2 = None

            if txt == 'Objetivos':
                en_objetivos = True
                continue

            if txt == 'Introducción':
                sec_count  += 1
                current_sec = {'num': str(sec_count), 'titulo': 'Introducción',
                                'bloques': [], 'subsecciones': []}
                secciones.append(current_sec)
                continue

            # "1. El servicio postventa (Ce a)" → num+1, sin etiqueta CE
            m = re.match(r'(\d+)\.\s+(.+?)(?:\s*\(Ce[^)]*\))?$', txt, re.IGNORECASE)
            if m:
                sec_count  += 1
                titulo      = re.sub(r'\s*\(Ce[^)]*\)', '', m.group(2), flags=re.IGNORECASE).strip()
                current_sec = {'num': str(sec_count), 'titulo': titulo,
                                'bloques': [], 'subsecciones': []}
                secciones.append(current_sec)
                continue

            # "Resumen", "Conclusiones" u otros títulos sin número
            if txt:
                sec_count  += 1
                current_sec = {'num': str(sec_count), 'titulo': txt,
                                'bloques': [], 'subsecciones': []}
                secciones.append(current_sec)
            continue

        # ── Heading 2 / Heading 3 ─────────────────────────────
        if style == 'Heading 2':
            flush()
            m = re.match(r'(\d+\.\d+)\s+(.+)', txt)
            if m and current_sec:
                current_sub  = {'num': m.group(1), 'titulo': limpiar_titulo(m.group(2)),
                                 'bloques': [], 'subsecciones': []}
                current_sub2 = None
                current_sec['subsecciones'].append(current_sub)
            continue

        if style == 'Heading 3':
            flush()
            m = re.match(r'(\d+\.\d+\.\d+)\s+(.+)', txt)
            if m and current_sub:
                current_sub2 = {'num': m.group(1), 'titulo': limpiar_titulo(m.group(2)),
                                  'bloques': []}
                current_sub['subsecciones'].append(current_sub2)
            continue

        # ── Filtros globales (tras headings) ──────────────────
        if (bold and txt in SKIP_BOLD) or txt.startswith(SKIP_PREFIX):
            flush()
            continue
        if txt in SKIP_TEXT and style not in MULTI_STYLES:
            flush()
            continue

        # ── Objetivos ─────────────────────────────────────────
        if en_objetivos:
            if style == 'List Paragraph' and txt:
                objetivos.append(txt)
            continue

        # No hay sección activa todavía
        if not current_sec:
            continue

        # ── Bloques multi-párrafo por estilo ──────────────────
        if style in MULTI_STYLES:
            estilo_orig = style

            if style == 'Sabiasque':
                if blk is None or blk.get('_estilo') != 'Sabiasque':
                    flush()
                    tipo_blk = LABEL_MAP.get(txt, 'sabias_que')
                    etq      = txt
                    blk = {'tipo': tipo_blk, 'etiqueta': etq, 'lineas': [], '_estilo': 'Sabiasque'}
                elif txt:
                    if txt.startswith('- '):
                        blk['lineas'].append(txt[2:])
                    else:
                        blk['lineas'].append(txt)
                continue

            if style == 'Hilo conductor':
                if txt == 'Hilo conductor':
                    flush()
                    blk = {'tipo': 'hilo_conductor', 'etiqueta': 'Hilo conductor',
                            'lineas': [], '_estilo': 'Hilo conductor'}
                elif blk and blk.get('_estilo') == 'Hilo conductor' and txt:
                    blk['lineas'].append(txt)
                continue

            if style == 'Ejemplo':
                if txt == 'Ejemplo':
                    flush()
                    blk = {'tipo': 'ejemplo', 'etiqueta': 'Ejemplo', 'lineas': [], '_estilo': 'Ejemplo'}
                elif blk and blk.get('_estilo') == 'Ejemplo' and txt:
                    blk['lineas'].append(txt)
                continue

            if style == 'Vídeo':
                if txt == 'Vídeo':
                    flush()
                    blk = {'tipo': 'video', 'etiqueta': 'Vídeo', 'lineas': [], '_estilo': 'Vídeo'}
                elif blk and blk.get('_estilo') == 'Vídeo' and txt:
                    blk['lineas'].append(txt)
                continue

            if style == 'Actividad colaborativa':
                if blk is None or blk.get('_estilo') != 'Actividad colaborativa':
                    flush()
                    blk = {'tipo': 'actividad_complementaria',
                            'etiqueta': 'Actividad complementaria',
                            'lineas': [], '_estilo': 'Actividad colaborativa'}
                else:
                    if txt:
                        m2 = re.match(r'\d+\.\s+En esta actividad deberás (.+)', txt)
                        texto_act = infinitivo_a_imperativo(m2.group(1)) if m2 else txt
                        blk['lineas'].append(texto_act)
                continue

            if style == 'Importante':
                if blk is None or blk.get('_estilo') != 'Importante':
                    flush()
                    etq = txt if txt else 'Importante'
                    blk = {'tipo': 'importante', 'etiqueta': etq, 'lineas': [], '_estilo': 'Importante'}
                elif txt:
                    blk['lineas'].append(txt)
                continue

            if style == 'Aplicación práctica':
                if blk is None or blk.get('_estilo') != 'Aplicación práctica':
                    flush()
                    mt = re.match(r'Tarea de evaluaci[oó]n\s+(\d+)', txt, re.IGNORECASE)
                    num_t = mt.group(1) if mt else '1'
                    blk = {'tipo': 'tarea', 'etiqueta': f'Tarea {num_t}',
                            'lineas': [], '_estilo': 'Aplicación práctica'}
                else:
                    # Omitir líneas de metadatos en negrita (Duración, Objetivo, Enunciado)
                    if txt and not bold:
                        blk['lineas'].append(txt)
                    elif txt and bold and not re.match(r'(Duración|Objetivo|Enunciado)', txt):
                        blk['lineas'].append(txt)
                continue

        # Salir de bloque multi-párrafo si el estilo cambió
        if blk and blk.get('_estilo') and style not in MULTI_STYLES:
            flush()

        # ── Marcador de interacción ───────────────────────────
        if bold and re.match(r'Interacci[oó]n\s+\d+', txt):
            m = re.match(r'Interacci[oó]n\s+(\d+)(\.?\s+(.+))?', txt)
            if not m:
                continue
            n       = int(m.group(1))
            label   = (m.group(3) or '').strip()
            inter   = interacciones.get(n, {})

            if inter.get('tipo') == 'opciones':
                # Actividad de evaluación: recoger enunciado del contexto DOCX
                mt2 = re.match(r'Actividad de evaluaci[oó]n\s+(\d+)', label, re.IGNORECASE)
                num_act = mt2.group(1) if mt2 else '1'
                blk = {'tipo': '_ejercicio_pendiente',
                        '_estilo': '_ejercicio',
                        'etiqueta': f'Actividad {num_act}',
                        'lineas': [],
                        'inter': inter}
            else:
                for vb in expandir_interaccion(n, interacciones):
                    bloques_activos().append(vb)
            continue

        # ── Recoger enunciado de ejercicio pendiente ──────────
        if blk and blk.get('tipo') == '_ejercicio_pendiente':
            if txt and not bold:
                txt_clean = re.sub(r'^Enunciado:\s*', '', txt)
                blk['lineas'].append(txt_clean)
            elif bold or not txt:
                # El ejercicio ha terminado → resolver
                inter    = blk['inter']
                lineas   = list(blk['lineas'])
                letras   = 'abcdefgh'
                opciones = []
                for j, opt in enumerate(inter.get('opciones', [])):
                    # No duplicar letra si ya viene prefijada
                    m_opt = re.match(r'^[a-h]\)\s*(.+)', opt)
                    if m_opt:
                        opciones.append({'letra': letras[j], 'texto': m_opt.group(1).strip()})
                    else:
                        opciones.append({'letra': letras[j], 'texto': opt})
                bloques_activos().append({
                    'tipo':      'tarea',
                    'etiqueta':  blk['etiqueta'],
                    'lineas':    lineas,
                    'opciones':  opciones,   # lista de {letra, texto} – sin duplicar letras
                    # Solución y feedback: solo versión docente, no se emiten en cuerpo
                    '_solucion': inter.get('solucion', ''),
                    '_feedback': inter.get('feedback', ''),
                })
                blk = None
                # Re-procesar la línea actual si era un título de sección
                if bold and txt:
                    bloques_activos().append({'tipo': 'parrafo', 'texto': txt})
            continue

        # ── Párrafo de imagen (URL / pie / descripción) ───────
        if txt.startswith('https://') or txt.startswith('http://'):
            url = txt
            if _es_url_imagen(url):
                add({'tipo': 'url_imagen', 'url': url})
            else:
                add({'tipo': 'url', 'url': url})
            continue

        if txt.startswith('Pie de imagen:'):
            add({'tipo': 'pie_imagen', 'texto': txt[len('Pie de imagen:'):].strip()})
            continue

        if re.match(r'^Descripci[oó]n de (la )?imagen:', txt):
            add({'tipo': 'desc_imagen', 'texto': re.sub(r'^Descripci[oó]n de (la )?imagen:\s*', '', txt)})
            continue

        # ── Párrafo de cuerpo ────────────────────────────────
        if txt and style in ('Normal', 'Normal (Web)', 'List Paragraph', 'Cuerpo parrafo'):
            if style == 'List Paragraph':
                bloques_activos().append({'tipo': 'p_vineta', 'texto': txt, 'nivel': 1})
            else:
                bloques_activos().append({'tipo': 'parrafo', 'texto': txt})
            continue

    # Flush bloque final
    flush()
    # Resolver ejercicio pendiente al final
    if blk and blk.get('tipo') == '_ejercicio_pendiente':
        inter  = blk['inter']
        lineas = list(blk['lineas'])
        letras = 'abcdefgh'
        opciones = []
        for j, opt in enumerate(inter.get('opciones', [])):
            m_opt = re.match(r'^[a-h]\)\s*(.+)', opt)
            if m_opt:
                opciones.append({'letra': letras[j], 'texto': m_opt.group(1).strip()})
            else:
                opciones.append({'letra': letras[j], 'texto': opt})
        bloques_activos().append({
            'tipo':      'tarea',
            'etiqueta':  blk['etiqueta'],
            'lineas':    lineas,
            'opciones':  opciones,
            '_solucion': inter.get('solucion', ''),
            '_feedback': inter.get('feedback', ''),
        })

    return {
        'titulo_unidad': f'Unidad de aprendizaje {ua_num}',
        'titulo_modulo': ua_titulo,
        'objetivos':     objetivos,
        'secciones':     secciones,
    }

# ── Generación de XML ─────────────────────────────────────────────────────────

def bloques_xml(bloques: list) -> list[str]:
    out = []
    for b in bloques:
        t = b.get('tipo', '')

        if t == 'parrafo':
            texto = b['texto']
            # Detectar fórmula dentro de párrafo de cuerpo
            if RE_FORMULA.search(texto) and len(texto) < 200:
                out.append(p_formula(texto))
            else:
                out.append(p(texto, 'Cuerpoparrafo'))

        elif t == 'desplegable':
            if b.get('descripcion'):
                out.append(p_desp(b['titulo'], b['descripcion'], nivel=1))
            else:
                out.append(p_vineta_bold(b['titulo'], nivel=1))

        elif t in ('nota', 'ejemplo', 'sabias_que', 'consejo', 'definicion',
                   'hilo_conductor', 'para_saber_mas', 'video'):
            out.append(p(b['etiqueta'], 'Ejemplos-01lneainicio'))
            for l in b.get('lineas', []):
                if not l.strip():
                    continue
                # Detectar viñetas dentro del bloque
                if l.startswith('- ') or l.startswith('* '):
                    out.append(p_vineta_ejemplo(l[2:].strip()))
                elif l.startswith('● ') or l.startswith('●\t'):
                    out.append(p_vineta_ejemplo(l[1:].strip()))
                else:
                    out.append(p(l, 'Ejemplos-Cuerpoparrafo'))
            out.append(p('', 'Ejemplos-02lneafin'))

        elif t == 'importante':
            # Importante usa los mismos estilos de caja editorial que Nota/Ejemplo
            out.append(p(b['etiqueta'], 'Ejemplos-01lneainicio'))
            for l in b.get('lineas', []):
                if not l.strip():
                    continue
                if l.startswith('- ') or l.startswith('* '):
                    out.append(p_vineta_ejemplo(l[2:].strip()))
                elif l.startswith('● ') or l.startswith('●\t'):
                    out.append(p_vineta_ejemplo(l[1:].strip()))
                else:
                    out.append(p(l, 'Ejemplos-Cuerpoparrafo'))
            out.append(p('', 'Ejemplos-02lneafin'))

        elif t == 'recuerda':
            out.append(p(b['etiqueta'], 'Recuerda-00lneainicio'))
            for l in b.get('lineas', []):
                if l.strip():
                    out.append(p(l, 'Recuerda-cuerpoparrafo'))
            out.append(p('', 'Recuerda-01lneafin'))

        elif t == 'tarea':
            out.append(p(b['etiqueta'], 'Ejercicios-01lneainicio'))
            for l in b.get('lineas', []):
                if l.strip():
                    out.append(p(l, 'EjerciciosPregunta'))
            # Opciones tipo test en párrafos separados con estilo propio
            for opt in b.get('opciones', []):
                out.append(p_opcion_test(opt['letra'], opt['texto']))
            # Solución y feedback: solo versión docente → NO se emiten aquí
            out.append(p('', 'Ejercicios-02lneafin'))

        elif t == 'actividad_complementaria':
            out.append(p('Actividad complementaria', 'Ejercicios-01lneainicio'))
            for l in b.get('lineas', []):
                if not l.strip():
                    continue
                # Líneas con número (1. texto) → viñeta numerada de ejercicio
                m_num = re.match(r'^(\d+)\.\s+(.+)', l)
                if m_num:
                    out.append(p(f'{m_num.group(1)}. {m_num.group(2)}', 'EjerciciosPregunta'))
                elif l.startswith('- ') or l.startswith('* '):
                    out.append(p(l[2:].strip(), 'EjerciciosPregunta'))
                else:
                    out.append(p(l, 'EjerciciosPregunta'))
            out.append(p('', 'Ejercicios-02lneafin'))

        elif t == 'imagen':
            # Marcador de imagen (placeholder)
            out.append(p('[IMAGEN]', 'Marcadorimagen'))
            if b.get('pie'):
                out.append(p_pie_imagen(b['pie']))
            if b.get('descripcion'):
                out.append(p_desc_imagen(b['descripcion']))

        elif t == 'url_imagen':
            # URL de imagen → marcador con estilo propio, nunca texto normal
            out.append(p_url_imagen(b['url']))

        elif t == 'url':
            # URL de recurso/enlace externo con estilo URL
            out.append(p_url_recurso(b['url']))

        elif t == 'pie_imagen':
            out.append(p_pie_imagen(b['texto']))

        elif t == 'desc_imagen':
            out.append(p_desc_imagen(b['texto']))

        elif t == 'p_vineta':
            out.append(p_vineta(b['texto'], nivel=b.get('nivel', 1)))

        elif t == 'p_vineta_bold':
            out.append(p_vineta_bold(b['texto'], nivel=b.get('nivel', 1)))

        elif t == 'p_vineta_ejemplo':
            out.append(p_vineta_ejemplo(b['texto']))

        elif t == 'desplegable_simple':
            out.append(p_desp(b['titulo'], b['contenido'], nivel=1))

        elif t == 'opcion_test_suelta':
            out.append(p_opcion_test(b['letra'], b['texto']))

        elif t == 'parrafo_formula':
            out.append(p_formula(b['texto']))

        elif t == 'desplegable_multi':
            # Título en negrita como viñeta nivel 1; sublista en nivel 2
            out.append(p_vineta_bold(b['titulo'] + ':', nivel=1))
            for item in b.get('items', []):
                if item.strip():
                    out.append(p_vineta(item, nivel=2))

    return out


def generar_docx(est: dict, ejemplo: Path | None, plantilla: Path, salida: Path):
    NS = (
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'mc:Ignorable="w14"'
    )

    pars = []

    # Portada
    pars.append(p(est['titulo_unidad'], 'TITULOUNIDAD1'))
    pars.append(p(est['titulo_modulo'], 'TITULOUNIDAD2', negrita=True))

    if est['objetivos']:
        pars.append(p('Los objetivos específicos de esta Unidad de Aprendizaje son:',
                      'Cuerpoparrafo'))
        for obj in est['objetivos']:
            pars.append(p_vineta(obj, nivel=1))

    # Secciones
    for sec in est['secciones']:
        pars.append(p(f"{sec['num']}. {sec['titulo']}", '1Ttulonvl1'))
        pars += bloques_xml(sec['bloques'])
        for sub in sec.get('subsecciones', []):
            pars.append(p(f"{sub['num']} {sub['titulo']}", '2Ttulonvl2'))
            pars += bloques_xml(sub['bloques'])
            for sub2 in sub.get('subsecciones', []):
                pars.append(p(f"{sub2['num']} {sub2['titulo']}", '3Ttulonvl3'))
                pars += bloques_xml(sub2['bloques'])

    pars.append(p('', 'Cuerpoparrafo'))

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}><w:body>\n'
        + '\n'.join(pars)
        + '\n<w:sectPr/></w:body></w:document>'
    ).encode('utf-8')

    # ── Construir el archivo ZIP de salida ────────────────────────
    def _es_zip(path: Path) -> bool:
        try:
            with open(path, 'rb') as f:
                return f.read(4) == b'PK\x03\x04'
        except Exception:
            return False

    # 1. Base: plantilla (si es ZIP válido)
    if plantilla.exists() and _es_zip(plantilla):
        with zipfile.ZipFile(str(plantilla), 'r') as zin:
            archivos = {n: zin.read(n) for n in zin.namelist()}
    else:
        # Sin plantilla válida: construir estructura DOCX mínima
        print('  ⚠ Plantilla no es un DOCX binario: generando estructura mínima')
        archivos = _crear_docx_minimo()

    # 2. Sobrescribir estilos con los del ejemplo maquetado (si es ZIP válido)
    if ejemplo and ejemplo.exists() and _es_zip(ejemplo):
        with zipfile.ZipFile(str(ejemplo), 'r') as zej:
            for nombre in zej.namelist():
                if nombre in ARCHIVOS_ESTILO:
                    archivos[nombre] = zej.read(nombre)
                    print(f'  ✓ Estilo copiado: {nombre}')
                elif nombre.startswith('word/theme/'):
                    archivos[nombre] = zej.read(nombre)
                    print(f'  ✓ Tema copiado:   {nombre}')
    else:
        print('  ⚠ Sin ejemplo maquetado binario: usando estilos de la plantilla')

    # 3. Inyectar el documento generado
    archivos['word/document.xml'] = doc_xml

    with zipfile.ZipFile(str(salida), 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, d in archivos.items():
            zout.writestr(n, d)


def _crear_docx_minimo() -> dict:
    """Crea la estructura ZIP mínima de un DOCX vacío válido."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    ).encode('utf-8')

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    ).encode('utf-8')

    word_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    ).encode('utf-8')

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        '<w:rPr><w:sz w:val="24"/></w:rPr>'
        '</w:style>'
        '</w:styles>'
    ).encode('utf-8')

    return {
        '[Content_Types].xml': content_types,
        '_rels/.rels': rels,
        'word/_rels/document.xml.rels': word_rels,
        'word/styles.xml': styles,
    }


# ── Punto de entrada ──────────────────────────────────────────────────────────

def parsear_docx_fuente_texto(path: Path, interacciones: dict) -> dict:
    """
    Parsea un archivo de unidad en formato texto/markdown.
    Detecta estructuras por patrones de línea en lugar de estilos de párrafo.
    Devuelve la misma estructura que parsear_docx_fuente().
    """
    text = path.read_text(encoding='utf-8', errors='replace')

    # Limpiar markdown
    def clean_md(s: str) -> str:
        # Links [text](url) → url o texto según si es imagen
        s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', lambda m:
                   m.group(2) if m.group(2).startswith('http') else m.group(1), s)
        # Negrita/cursiva **texto** / *texto* / _texto_
        s = re.sub(r'\*{1,3}([^*]+?)\*{1,3}', r'\1', s)
        s = re.sub(r'_{1,2}([^_]+?)_{1,2}', r'\1', s)
        # Tablas pipe: eliminar celdas vacías
        if s.startswith('|'):
            s = s.strip('| ').strip()
        return s.strip()

    lines_raw = text.split('\n')
    lines = [clean_md(l) for l in lines_raw]

    ua_num    = ''
    ua_titulo = ''
    objetivos = []
    secciones = []

    current_sec  = None
    current_sub  = None
    current_sub2 = None
    sec_count    = 0
    en_objetivos = False
    blk          = None

    def bloques_activos():
        if current_sub2:
            return current_sub2['bloques']
        if current_sub:
            return current_sub['bloques']
        if current_sec:
            return current_sec['bloques']
        return []

    def flush():
        nonlocal blk
        if blk is None:
            return
        if blk.get('tipo') == '_ejercicio_pendiente':
            # Resolver ejercicio antes de hacer flush normal
            inter    = blk['inter']
            lineas   = list(blk['lineas'])
            letras   = 'abcdefgh'
            opciones = []
            for j, opt in enumerate(inter.get('opciones', [])):
                m_opt = re.match(r'^[a-h]\)\s*(.+)', opt)
                if m_opt:
                    opciones.append({'letra': letras[j], 'texto': m_opt.group(1).strip()})
                else:
                    opciones.append({'letra': letras[j], 'texto': opt})
            bloques_activos().append({
                'tipo':      'tarea',
                'etiqueta':  blk['etiqueta'],
                'lineas':    lineas,
                'opciones':  opciones,
                '_solucion': inter.get('solucion', ''),
                '_feedback': inter.get('feedback', ''),
            })
        elif blk.get('lineas'):
            bloques_activos().append({k: v for k, v in blk.items() if not k.startswith('_')})
        blk = None

    def add(b):
        flush()
        bloques_activos().append(b)

    i = 0
    while i < len(lines):
        l = lines[i]
        orig = lines_raw[i].strip() if i < len(lines_raw) else l

        # Saltar separadores de tabla y líneas vacías
        if not l or re.match(r'^-+$', l) or l == '---':
            i += 1
            continue

        # Saltar líneas solo PUA
        if es_solo_pua(l):
            i += 1
            continue

        # Saltar instrucciones digitales
        if debe_elim(l):
            i += 1
            continue

        # Título principal
        if not ua_num:
            m = re.match(r'Unidad de aprendizaje\s+(\d+)[.\-–\s]\s*(.+)', l)
            if m:
                ua_num    = m.group(1)
                ua_titulo = m.group(2).strip()
                i += 1
                continue
            # También sin número
            m2 = re.match(r'Unidad de aprendizaje\s+(\d+)', l)
            if m2:
                ua_num = m2.group(1)
                i += 1
                continue

        # Título de bloque de unidad (primera línea bold corta después del UA)
        if not ua_titulo and ua_num and i > 0:
            if re.match(r'\*\*', lines_raw[i]) and len(l) < 80:
                ua_titulo = l
                i += 1
                continue

        # Objetivos
        if l == 'Objetivos' or l == 'Específicos':
            en_objetivos = True
            i += 1
            continue

        if en_objetivos:
            # Turn-off conditions
            if (re.match(r'^\d+\.', l) or l in BLOQUES_ESP or RE_SEC1.match(l)
                    or l in ('Introducción', 'Resumen')
                    or l.startswith('Criterios de evaluación')
                    or (l and l[0].isdigit() and '.' in l)):
                en_objetivos = False
                # Fall through to process this line normally
            elif l and not l.startswith('CE') and not re.match(r'^[a-e]\) Se han', l):
                # Limpieza de viñeta markdown
                obj_txt = re.sub(r'^[-*]\s*', '', l).strip()
                if obj_txt and len(obj_txt) > 10:
                    objetivos.append(obj_txt)
                i += 1
                continue
            else:
                # CE criteria → skip
                i += 1
                continue

        # Heading 1: "1. Título"
        m = RE_SEC1.match(l)
        if m and es_titulo_seccion(m.group(2)):
            flush()
            sec_count += 1
            titulo = re.sub(r'\s*\(Ce[^)]*\)', '', limpiar_titulo(m.group(2)),
                            flags=re.IGNORECASE).strip()
            current_sec  = {'num': str(sec_count), 'titulo': titulo,
                             'bloques': [], 'subsecciones': []}
            current_sub  = None
            current_sub2 = None
            secciones.append(current_sec)
            i += 1
            continue

        if l in ('Introducción', 'Resumen'):
            flush()
            sec_count += 1
            current_sec  = {'num': str(sec_count), 'titulo': l,
                             'bloques': [], 'subsecciones': []}
            current_sub  = None
            current_sub2 = None
            secciones.append(current_sec)
            i += 1
            continue

        if not current_sec:
            i += 1
            continue

        # Heading 2
        m = RE_SEC2.match(l)
        if m:
            flush()
            current_sub  = {'num': m.group(1), 'titulo': limpiar_titulo(m.group(3)),
                             'bloques': [], 'subsecciones': []}
            current_sub2 = None
            current_sec['subsecciones'].append(current_sub)
            i += 1
            continue

        # Heading 3
        m = RE_SEC3.match(l)
        if m:
            flush()
            current_sub2 = {'num': m.group(1), 'titulo': limpiar_titulo(m.group(4)),
                             'bloques': []}
            if current_sub:
                current_sub['subsecciones'].append(current_sub2)
            else:
                current_sec['subsecciones'].append(current_sub2)
            i += 1
            continue

        # Bloques especiales
        if l in BLOQUES_ESP:
            flush()
            TIPO_MAP_T = {
                'Nota': 'nota', 'Ejemplo': 'ejemplo',
                'Sabías que...': 'sabias_que', 'Sabías que…': 'sabias_que',
                'Consejo': 'consejo', 'Definición': 'definicion',
                'Hilo conductor': 'hilo_conductor', 'Para saber más': 'para_saber_mas',
                'Vídeo': 'video', 'Importante': 'importante', 'Recuerda': 'recuerda',
            }
            tipo = TIPO_MAP_T.get(l, 'ejemplo')
            blk = {'tipo': tipo, 'etiqueta': l, 'lineas': [], '_estilo': l}
            i += 1
            continue

        # Continuar bloque especial
        if blk and blk.get('_estilo') in BLOQUES_ESP:
            # Terminadores del bloque especial
            es_fin_bloque = (
                l in BLOQUES_ESP
                or RE_SEC1.match(l) and es_titulo_seccion(RE_SEC1.match(l).group(2))
                or RE_SEC2.match(l)
                or re.match(r'Interacci[oó]n\s+\d+', l)
                or RE_URL.match(l)
                or l.startswith('Pie de imagen:')
                or re.match(r'^Descripci[oó]n de (la )?imagen:', l)
                or l in ('Introducción', 'Resumen')
            )
            if es_fin_bloque:
                flush()
                # No incrementar i → releer esta línea en el contexto normal
            else:
                if l:
                    blk['lineas'].append(l)
                i += 1
                continue

        # Continuar bloque Actividad colaborativa
        if blk and blk.get('_estilo') == 'Actividad colaborativa':
            es_fin_act = (
                RE_SEC1.match(l) and es_titulo_seccion(RE_SEC1.match(l).group(2))
                or RE_SEC2.match(l)
                or l in BLOQUES_ESP
                or re.match(r'Interacci[oó]n\s+\d+', l)
                or l in ('Introducción', 'Resumen')
                or re.match(r'^Actividad (colaborativa|complementaria)\b', l, re.IGNORECASE)
                # Terminar en párrafo de cuerpo largo que no es una tarea
                or (len(l) > 120 and not re.match(r'^\d+\.', l) and blk['lineas'])
            )
            if es_fin_act:
                flush()
            else:
                txt_act = re.sub(r'^[-*]\s*', '', l).strip()
                txt_act = re.sub(r'^\d+\.\s+En esta actividad deberás\s+', '', txt_act)
                txt_act = infinitivo_a_imperativo(txt_act)
                if txt_act:
                    blk['lineas'].append(txt_act)
                i += 1
                continue

        # URL de imagen
        if RE_URL.match(l):
            flush()
            url = l
            if _es_url_imagen(url):
                add({'tipo': 'url_imagen', 'url': url})
            else:
                add({'tipo': 'url', 'url': url})
            i += 1
            continue

        # Pie de imagen / descripción
        if l.startswith('Pie de imagen:'):
            flush()
            add({'tipo': 'pie_imagen', 'texto': l[len('Pie de imagen:'):].strip()})
            i += 1
            continue
        if re.match(r'^Descripci[oó]n de (la )?imagen:', l):
            flush()
            add({'tipo': 'desc_imagen',
                 'texto': re.sub(r'^Descripci[oó]n de (la )?imagen:\s*', '', l)})
            i += 1
            continue

        # Actividad colaborativa (markdown: "**Actividad colaborativa**" o variantes)
        if re.match(r'^Actividad (colaborativa|complementaria)\b', l, re.IGNORECASE):
            flush()
            blk = {'tipo': 'actividad_complementaria',
                   'etiqueta': 'Actividad complementaria',
                   'lineas': [], '_estilo': 'Actividad colaborativa'}
            i += 1
            continue

        # Marcador de interacción
        m_inter = re.match(r'Interacci[oó]n\s+(\d+)(\.?\s+(.+))?', l)
        if m_inter:
            flush()
            n_int = int(m_inter.group(1))
            label = (m_inter.group(3) or '').strip()
            inter = interacciones.get(n_int, {})
            if inter.get('tipo') == 'opciones':
                mt2 = re.match(r'Actividad de evaluaci[oó]n\s+(\d+)', label, re.IGNORECASE)
                num_act = mt2.group(1) if mt2 else '1'
                blk = {'tipo': '_ejercicio_pendiente', '_estilo': '_ejercicio',
                       'etiqueta': f'Actividad {num_act}', 'lineas': [], 'inter': inter}
            else:
                for vb in expandir_interaccion(n_int, interacciones):
                    bloques_activos().append(vb)
            i += 1
            continue

        # Ejercicio pendiente
        if blk and blk.get('tipo') == '_ejercicio_pendiente':
            # Terminar en sección nueva, bloque especial, o línea en negrita sin enunciado
            es_seccion = (RE_SEC1.match(l) and es_titulo_seccion(RE_SEC1.match(l).group(2))
                          or RE_SEC2.match(l) or l in ('Resumen',))
            es_metadato = re.match(r'^(Duración|Objetivo|Enunciado[^:]|Criterios|CE |Solución|Feedback)', l)
            es_fin = (es_seccion or
                      (re.match(r'\*\*', lines_raw[i]) and l not in ('',)) or
                      l in BLOQUES_ESP)

            if es_fin:
                # Resolver y cerrar ejercicio
                inter    = blk['inter']
                lineas   = list(blk['lineas'])
                letras   = 'abcdefgh'
                opciones = []
                for j, opt in enumerate(inter.get('opciones', [])):
                    m_opt = re.match(r'^[a-h]\)\s*(.+)', opt)
                    if m_opt:
                        opciones.append({'letra': letras[j], 'texto': m_opt.group(1).strip()})
                    else:
                        opciones.append({'letra': letras[j], 'texto': opt})
                bloques_activos().append({
                    'tipo':      'tarea',
                    'etiqueta':  blk['etiqueta'],
                    'lineas':    lineas,
                    'opciones':  opciones,
                    '_solucion': inter.get('solucion', ''),
                    '_feedback': inter.get('feedback', ''),
                })
                blk = None
                # Re-procesar esta misma línea si es un nuevo contexto
                continue  # NO incrementar i — releer la línea
            elif es_metadato:
                # Saltar metadatos (Duración, Objetivo, etc.) pero mantener Enunciado:
                if l.startswith('Enunciado:'):
                    txt_clean = l[len('Enunciado:'):].strip()
                    if txt_clean:
                        blk['lineas'].append(txt_clean)
                i += 1
                continue
            else:
                txt_clean = re.sub(r'^Enunciado:\s*', '', l)
                blk['lineas'].append(txt_clean)
                i += 1
                continue

        # Opción tipo test suelta (a) ... b) ...)
        m_opt = RE_OPCION.match(l)
        if m_opt:
            add({'tipo': 'opcion_test_suelta',
                 'letra': m_opt.group(1), 'texto': m_opt.group(2)})
            i += 1
            continue

        # Viñeta markdown
        if orig.startswith('**●**') or orig.startswith('- ') or orig.startswith('* '):
            texto = re.sub(r'^(\*\*●\*\*|-|\*)\s*', '', orig).strip()
            texto = clean_md(texto)
            # Detectar si es desplegable (negrita al inicio)
            m_desp = re.match(r'\*\*(.+?)\.\*\*\s*(.+)', orig)
            if m_desp:
                add({'tipo': 'desplegable_simple',
                     'titulo': clean_md(m_desp.group(1)),
                     'contenido': clean_md(m_desp.group(2))})
            else:
                add({'tipo': 'p_vineta', 'texto': texto, 'nivel': 1})
            i += 1
            continue

        # Fórmula
        if RE_FORMULA.search(l) and len(l) < 200:
            add({'tipo': 'parrafo_formula', 'texto': l})
            i += 1
            continue

        # Párrafo normal
        add({'tipo': 'parrafo', 'texto': l})
        i += 1

    flush()

    # Resolver ejercicio pendiente al final
    if blk and blk.get('tipo') == '_ejercicio_pendiente':
        inter  = blk['inter']
        lineas = list(blk['lineas'])
        letras = 'abcdefgh'
        opciones = []
        for j, opt in enumerate(inter.get('opciones', [])):
            m_opt = re.match(r'^[a-h]\)\s*(.+)', opt)
            if m_opt:
                opciones.append({'letra': letras[j], 'texto': m_opt.group(1).strip()})
            else:
                opciones.append({'letra': letras[j], 'texto': opt})
        bloques_activos().append({
            'tipo':      'tarea',
            'etiqueta':  blk['etiqueta'],
            'lineas':    lineas,
            'opciones':  opciones,
            '_solucion': inter.get('solucion', ''),
            '_feedback': inter.get('feedback', ''),
        })

    return {
        'titulo_unidad': f'Unidad de aprendizaje {ua_num}',
        'titulo_modulo': ua_titulo,
        'objetivos':     objetivos,
        'secciones':     secciones,
    }


def _es_docx_binario(path: Path) -> bool:
    """Devuelve True si el archivo es un DOCX binario (ZIP), False si es texto."""
    try:
        with open(path, 'rb') as f:
            magic = f.read(4)
        return magic == b'PK\x03\x04'
    except Exception:
        return False


def main():
    args = sys.argv[1:]

    # Modos:
    # 2: UNIDAD EJEMPLO
    # 3: UNIDAD EJEMPLO SALIDA
    # 4: UNIDAD EJEMPLO INTERACCIONES SALIDA
    # 5: UNIDAD EJEMPLO PLANTILLA INTERACCIONES SALIDA (compatibilidad)
    interac_path = None
    if len(args) == 2:
        pdf, ejemplo = (Path(a) for a in args)
        plantilla = ejemplo
        salida = pdf.with_name(pdf.stem + '_resultado.docx')
    elif len(args) == 3:
        pdf, ejemplo, salida = (Path(a) for a in args)
        plantilla = ejemplo
    elif len(args) == 4:
        pdf, ejemplo, interac_path_str, salida = args
        pdf, ejemplo, salida = Path(pdf), Path(ejemplo), Path(salida)
        plantilla = ejemplo
        interac_path = Path(interac_path_str)
    elif len(args) == 5:
        pdf, ejemplo, plantilla, interac_path_str, salida = args
        pdf, ejemplo, plantilla, salida = Path(pdf), Path(ejemplo), Path(plantilla), Path(salida)
        interac_path = Path(interac_path_str)
    else:
        print('Uso: python conversor_papel.py UNIDAD EJEMPLO [INTERACCIONES] [SALIDA]')
        sys.exit(1)

    pdf, ejemplo, plantilla, salida = Path(pdf), Path(ejemplo), Path(plantilla), Path(salida)

    for f, n in [(pdf, 'Unidad'), (ejemplo, 'Ejemplo maquetado')]:
        if not f.exists():
            print(f'ERROR: No existe {n}: {f}')
            sys.exit(1)

    # ── Detectar / cargar interacciones ──────────────────────────────────────
    interacciones = {}
    if pdf.suffix.lower() in ('.docx', '.doc'):
        if interac_path is None:
            candidatos = sorted(pdf.parent.glob('interacciones_*.docx'))
            if candidatos:
                interac_path = candidatos[0]
                print(f'  → Interacciones detectadas: {interac_path.name}')
        if interac_path and interac_path.exists():
            print(f'→ Parseando interacciones de {interac_path.name}...')
            if _es_docx_binario(interac_path):
                interacciones = parsear_interacciones(interac_path)
            else:
                interacciones = parsear_interacciones_texto(interac_path)
            print(f'  {len(interacciones)} interacciones cargadas')

    print(f'→ Parseando {pdf.name}...')
    if pdf.suffix.lower() in ('.docx', '.doc'):
        if _es_docx_binario(pdf):
            est = parsear_docx_fuente(pdf, interacciones)
        else:
            est = parsear_docx_fuente_texto(pdf, interacciones)
    else:
        paginas = extraer_texto(pdf)
        print(f'  {len(paginas)} páginas')
        est = parsear(paginas)
    print(f'  {est["titulo_unidad"]} — {est["titulo_modulo"]}')
    print(f'  {len(est["objetivos"])} objetivos')

    total_bloques = 0
    for s in est['secciones']:
        nb  = len(s['bloques'])
        nss = len(s.get('subsecciones', []))
        nsb = sum(len(ss['bloques']) for ss in s.get('subsecciones', []))
        total_bloques += nb + nsb
        print(f'    {s["num"]}. {s["titulo"]}  '
              f'[{nb} bloques, {nss} subsecs, {nsb} bloques subsec]')
    print(f'  Total bloques: {total_bloques}')

    print(f'→ Generando {salida.name}...')
    generar_docx(est, ejemplo, plantilla, salida)
    print(f'✓ Listo: {salida}')


if __name__ == '__main__':
    main()