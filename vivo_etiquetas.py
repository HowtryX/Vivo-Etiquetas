from __future__ import annotations

import json
import os
import sys
import subprocess
import tempfile
import math
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    DecodedStreamObject,
    DictionaryObject,
    NameObject,
)

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None


# ============================================================
# ABRIR ARQUIVO
# ============================================================

def abrir_arquivo(caminho):
    caminho = str(caminho)
    if sys.platform.startswith("win"):
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

APP_DIR = Path.home() / ".vivo_etiquetas"
CONFIG_FILE = APP_DIR / "config.json"

CONFIG_PADRAO = {
    "plano_vivo": "Vivo Controle 12GB",
    "valor_plano_vivo": 57.00,
    "arquivo_excel": "",
    "arquivo_pdf_modelo": "",
    "tema": "sistema",          # "sistema", "claro" ou "escuro"
}

# Excel
ABA_EXCEL = "SMARTPHONES"
COL_NOME = "E"
COL_PRE_PAGO = "F"
COL_ENTRADA = "AX"

# No bloco "CONTROLE ENTRADA" da tabela, cada coluna de parcelamento traz o
# VALOR TOTAL a pagar naquele parcelamento (a parcela = total / nº de parcelas).
COL_TOTAL_12X = "BJ"
COL_TOTAL_21X = "BS"

# PDF
# O número de etiquetas por página é descoberto no próprio PDF modelo
# (função definir_capacidade). 10 é só o valor inicial.
ETIQUETAS_POR_PAGINA = 10
CAMPOS_POR_ETIQUETA = 13

# Textos fixos escritos em toda etiqueta (ficam iguais ao modelo se já forem
# iguais; só são reescritos quando diferentes)
SEGURO_NOME = "Vivo Seguro Celular"
SEGURO_MESES = "10x"

# Nome do aparelho em MAIÚSCULAS? (False = como está no Excel)
NOME_EM_MAIUSCULAS = False

# Tamanho mínimo da letra quando o texto precisa "encolher" para caber
TAMANHO_MINIMO = 5.0


# ============================================================
# CONFIGURAÇÃO
# ============================================================

def salvar_config(config):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as arquivo:
        json.dump(config, arquivo, ensure_ascii=False, indent=4)


def carregar_config():
    APP_DIR.mkdir(parents=True, exist_ok=True)

    if not CONFIG_FILE.exists():
        salvar_config(CONFIG_PADRAO.copy())
        return CONFIG_PADRAO.copy()

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        config = CONFIG_PADRAO.copy()
        config.update(dados)
        config.pop("taxa_cartao_mes", None)   # campo antigo (não é mais usado)
        return config
    except (OSError, json.JSONDecodeError):
        return CONFIG_PADRAO.copy()


# ============================================================
# MOEDA / DECIMAL
# ============================================================

def moeda(valor):
    return Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def texto_formatado(valor):
    valor = moeda(valor)
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return texto


def formatar_reais(valor):
    return f"R$ {texto_formatado(valor)}"


def decimal_brasileiro(valor):
    if valor is None:
        return Decimal("0.00")

    texto = str(valor).strip()
    if not texto:
        return Decimal("0.00")

    texto = texto.replace("R$", "").strip()

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    return moeda(texto)


# ============================================================
# SEGURO
# ============================================================

SEGURO_FAIXAS = [
    {"limite": Decimal("2000.00"), "valor": Decimal("16.90")},
    {"limite": Decimal("3000.00"), "valor": Decimal("34.90")},
    {"limite": Decimal("5000.00"), "valor": Decimal("69.90")},
    {"limite": Decimal("10000.00"), "valor": Decimal("129.90")},
    {"limite": Decimal("15000.00"), "valor": Decimal("199.90")},
]


def calcular_seguro(valor_aparelho):
    valor_aparelho = moeda(valor_aparelho)

    for faixa in SEGURO_FAIXAS:
        if valor_aparelho <= faixa["limite"]:
            return moeda(faixa["valor"])

    raise ValueError(
        "O valor do aparelho ultrapassa a maior faixa de seguro."
    )


# ============================================================
# PIX
# ============================================================

def calcular_pix(controle_entrada):
    return moeda(moeda(controle_entrada) * Decimal("0.90"))


# ============================================================
# PARCELAMENTO (valores vêm da tabela de preços)
# ============================================================

def calcular_parcela_tabela(total, quantidade_parcelas, nome_aparelho=""):
    """
    A tabela informa o TOTAL a pagar no parcelamento (já com os juros,
    quando houver). O valor de cada parcela é esse total dividido pelo
    número de parcelas.
    """
    total = moeda(total)

    if total <= 0:
        raise ValueError(
            f"A tabela não tem o valor do parcelamento em "
            f"{quantidade_parcelas}x para este aparelho"
            + (f" ({nome_aparelho})." if nome_aparelho else ".")
        )

    return moeda(total / quantidade_parcelas)


# ============================================================
# APARELHO
# ============================================================

class Aparelho:

    def __init__(self, modelo, pre_pago, controle_entrada,
                 total_12x=0, total_21x=0):
        self.modelo = modelo
        self.pre_pago = moeda(pre_pago)
        self.controle_entrada = moeda(controle_entrada)
        self.total_12x = moeda(total_12x)
        self.total_21x = moeda(total_21x)


# ============================================================
# CÁLCULO COMPLETO DA ETIQUETA
# ============================================================

def calcular_etiqueta(aparelho, plano_vivo, valor_plano):
    return {
        "nome_dispositivo": aparelho.modelo,
        "valor_pix": calcular_pix(aparelho.controle_entrada),
        "valor_pre_pago": aparelho.pre_pago,
        "valor_seguro": calcular_seguro(aparelho.controle_entrada),
        "controle_entrada": aparelho.controle_entrada,
        "total_12x": aparelho.total_12x,
        "total_21x": aparelho.total_21x,
        "parcela_12x": calcular_parcela_tabela(
            aparelho.total_12x, 12, aparelho.modelo
        ),
        "parcela_21x": calcular_parcela_tabela(
            aparelho.total_21x, 21, aparelho.modelo
        ),
        "plano_vivo": plano_vivo,
        "valor_plano_vivo": moeda(valor_plano),
    }


# ============================================================
# REPOSITÓRIO EXCEL
# ============================================================

class ExcelRepository:

    def __init__(self):
        self.linhas = []

    def carregar(self, caminho):
        if load_workbook is None:
            raise RuntimeError(
                "A biblioteca openpyxl não está instalada.\n\n"
                "Execute:\npip install openpyxl"
            )

        arquivo = Path(caminho)
        if not arquivo.exists():
            raise FileNotFoundError("Arquivo Excel não encontrado.")

        workbook = load_workbook(arquivo, data_only=True)

        if ABA_EXCEL not in workbook.sheetnames:
            raise ValueError(
                f"A aba '{ABA_EXCEL}' não foi encontrada.\n\n"
                "Abas disponíveis:\n" + "\n".join(workbook.sheetnames)
            )

        planilha = workbook[ABA_EXCEL]
        self.linhas = []

        for numero_linha in range(4, planilha.max_row + 1):
            nome = planilha[f"{COL_NOME}{numero_linha}"].value
            if nome is None:
                continue

            nome = str(nome).strip()
            if not nome:
                continue

            self.linhas.append({
                "Nome Comercial": nome,
                "Pré-pago": planilha[f"{COL_PRE_PAGO}{numero_linha}"].value,
                "Controle Entrada": planilha[f"{COL_ENTRADA}{numero_linha}"].value,
                "Total 12x": planilha[f"{COL_TOTAL_12X}{numero_linha}"].value,
                "Total 21x": planilha[f"{COL_TOTAL_21X}{numero_linha}"].value,
                "linha_excel": numero_linha,
            })

        return len(self.linhas)

    def pesquisar_modelo(self, texto):
        texto = str(texto).lower().strip()
        if not texto:
            return self.linhas.copy()

        return [
            linha for linha in self.linhas
            if texto in str(linha.get("Nome Comercial", "")).lower()
        ]


# ============================================================
# LEITURA DA ESTRUTURA DO PDF MODELO
# ============================================================
#
# Cada etiqueta tem 13 campos. Em vez de depender da ORDEM dos campos
# no arquivo (que muda), descobrimos qual campo é qual pela POSIÇÃO
# dele na página:
#
#   - o campo "TL ..." (texto pequeno de rodapé) é a âncora da etiqueta
#     e NUNCA é alterado;
#   - os demais campos são ligados à etiqueta mais próxima;
#   - o papel de cada campo vem do nome ou da posição dentro da etiqueta.

def _nome_campo(anot):
    nome = anot.get("/T")
    if nome is None and anot.get("/Parent") is not None:
        nome = anot["/Parent"].get_object().get("/T")
    return str(nome) if nome is not None else ""


def _valor_campo(anot):
    valor = anot.get("/V")
    if valor is None and anot.get("/Parent") is not None:
        valor = anot["/Parent"].get_object().get("/V")
    return str(valor) if valor is not None else ""


def _widgets_da_pagina(pagina):
    widgets = []

    for ref in pagina.get("/Annots", []) or []:
        anot = ref.get_object()

        if anot.get("/Subtype") != "/Widget":
            continue

        r = [float(v) for v in anot["/Rect"]]
        x0, x1 = sorted((r[0], r[2]))
        y0, y1 = sorted((r[1], r[3]))

        widgets.append({
            "anot": anot,
            "nome": _nome_campo(anot),
            "x0": x0, "y0": y0, "x1": x1, "y1": y1,
        })

    return widgets


def mapear_etiquetas(pagina):
    """
    Retorna uma lista com 8 dicionários (esquerda→direita, cima→baixo),
    cada um com os 13 campos da etiqueta separados por papel.
    """
    widgets = _widgets_da_pagina(pagina)

    ancoras = [w for w in widgets if w["nome"].upper().startswith("TL")]

    if not ancoras:
        raise ValueError(
            "O PDF modelo não tem a estrutura esperada: "
            "nenhuma etiqueta foi encontrada na página 1."
        )

    ancoras.sort(key=lambda a: (-round(a["y0"]), a["x0"]))

    grupos = [[] for _ in ancoras]

    for w in widgets:
        colunas = [a for a in ancoras if a["x0"] <= w["x0"] + 5]
        if not colunas:
            continue

        x_coluna = max(a["x0"] for a in colunas)
        candidatas = [
            (i, a) for i, a in enumerate(ancoras)
            if abs(a["x0"] - x_coluna) < 1 and a["y0"] <= w["y0"] + 1
        ]
        if not candidatas:
            continue

        indice, _ = max(candidatas, key=lambda ia: ia[1]["y0"])
        grupos[indice].append(w)

    etiquetas = []

    for ancora, grupo in zip(ancoras, grupos):

        if len(grupo) != CAMPOS_POR_ETIQUETA:
            raise ValueError(
                "O PDF modelo não possui a quantidade esperada "
                "de campos por etiqueta.\n\n"
                f"Encontrados: {len(grupo)}\n"
                f"Esperados: {CAMPOS_POR_ETIQUETA}"
            )

        papeis = {}
        precos = []

        for w in grupo:
            n = w["nome"].upper()

            if n.startswith("TL"):
                papeis["tl"] = w
            elif n.startswith("ESCOLHA DO PLANO"):
                papeis["plano"] = w
            elif n.startswith("COMBO BOX"):
                papeis["seguro_nome"] = w
            elif n.startswith("TEXT FIELD"):
                papeis["pix"] = w
            elif n.startswith("12X") or n.startswith("PARCELA SEGURO"):
                papeis["seguro_meses"] = w
            elif n.startswith("NOME DO APARELHO"):
                papeis["nome"] = w
            elif n.startswith("PRE"):
                precos.append(w)

        # Se o modelo não tem um campo "Text Field" (PIX), o PIX é o campo
        # numérico de maior altura (fonte grande).
        if "pix" not in papeis and precos:
            pix = max(precos, key=lambda w: w["y1"] - w["y0"])
            precos.remove(pix)
            papeis["pix"] = pix

        esquerda = sorted(
            [w for w in precos if w["x0"] - ancora["x0"] < 50],
            key=lambda w: w["y0"])
        meio = sorted(
            [w for w in precos if 50 <= w["x0"] - ancora["x0"] < 120],
            key=lambda w: w["y0"])
        direita = sorted(
            [w for w in precos if w["x0"] - ancora["x0"] >= 120],
            key=lambda w: w["y0"])

        if (len(esquerda), len(meio), len(direita)) != (2, 2, 3) or len(papeis) != 6:
            raise ValueError(
                "Não consegui identificar os campos de uma das etiquetas "
                "do PDF modelo."
            )

        papeis["seguro_valor"], papeis["pre_pago"] = esquerda
        # (ordenado de baixo p/ cima) linha de baixo = total em 21x,
        # linha de cima = à vista
        papeis["total_21x"], papeis["a_vista"] = meio
        papeis["plano_valor"], papeis["parcela_21"], papeis["parcela_12"] = direita

        etiquetas.append(papeis)

    return etiquetas


# ============================================================
# TEXTO DENTRO DO CAMPO (aparência)
# ============================================================
#
# Em vez de pedir para o visualizador "regerar" o campo (o que faz o
# Windows aumentar a letra), trocamos apenas o TEXTO dentro da aparência
# que o próprio PDF já tem. Posição, fonte, cor e fundo ficam idênticos.

_RE_TF = re.compile(r"/(\S+)\s+([\d.]+)\s+Tf")

ESPACAMENTO_LINHAS = 1.242   # igual ao do modelo (9,936 pt para letra 8)


def _num(valor):
    if hasattr(valor, "get_object"):
        valor = valor.get_object()
    return float(valor)


def _codificar(texto):
    # As fontes Vivo do modelo usam codificação MacRoman
    return str(texto).encode("mac_roman", errors="replace")


def _literal_pdf(dados):
    dados = dados.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")
    return dados.decode("latin-1")


def _fontes_do_formulario(writer):
    try:
        dr = writer._root_object["/AcroForm"]["/DR"].get_object()
        return dr["/Font"].get_object()
    except Exception:
        return {}


def _largura_texto(fonte, dados, tamanho):
    primeiro = int(fonte.get("/FirstChar", 0))
    larguras = fonte.get("/Widths")
    total = 0.0

    for byte in dados:
        i = byte - primeiro
        if larguras is not None and 0 <= i < len(larguras):
            total += _num(larguras[i])
        else:
            total += 500.0

    return total * tamanho / 1000.0


def _ascent_descent(fonte):
    """Ascent/descent da fonte em fração do tamanho (ex.: 1.005 / -0.287)."""
    try:
        fd = fonte["/FontDescriptor"].get_object()
        return _num(fd["/Ascent"]) / 1000.0, _num(fd["/Descent"]) / 1000.0
    except Exception:
        return 0.95, -0.25


def _quebrar_linhas(texto, fonte, tamanho, largura):
    linhas = []
    atual = ""

    for palavra in str(texto).split():
        teste = f"{atual} {palavra}".strip()

        if _largura_texto(fonte, _codificar(teste), tamanho) <= largura or not atual:
            atual = teste
        else:
            linhas.append(atual)
            atual = palavra

    if atual:
        linhas.append(atual)

    return linhas


def _da_do_campo(anot):
    da = anot.get("/DA")
    if da is None and anot.get("/Parent") is not None:
        da = anot["/Parent"].get_object().get("/DA")
    return str(da) if da is not None else ""


def definir_texto_campo(writer, widget, texto, multilinha=False):
    """
    Escreve o texto de um campo (vazio ou preenchido) montando a aparência
    do zero, com as mesmas regras de posição do modelo. Fundo (ex.: cinza
    dos campos de plano) e demais elementos da aparência são mantidos.
    """
    anot = widget["anot"]
    ap = anot["/AP"]["/N"].get_object()

    dados = ap.get_data().decode("latin-1")

    # Tudo que vem ANTES do texto (ex.: retângulo cinza de fundo) é mantido
    corte = dados.find("/Tx BMC")
    prefixo = dados[:corte] if corte >= 0 else dados

    # Fonte e tamanho vêm do /DA do campo; se for "automático" (0),
    # usa o tamanho que já está na aparência atual.
    m = _RE_TF.search(_da_do_campo(anot)) or _RE_TF.search(dados)
    if not m:
        return

    fonte_nome = m.group(1)
    tamanho = float(m.group(2))

    bbox = [float(v) for v in ap["/BBox"]]
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]

    if tamanho <= 0:
        m_ap = _RE_TF.search(dados)
        tamanho = float(m_ap.group(2)) if m_ap and float(m_ap.group(2)) > 0 else round(h * 0.6177, 3)

    fontes = _fontes_do_formulario(writer)
    fonte = fontes.get("/" + fonte_nome)
    fonte = fonte.get_object() if fonte is not None else None

    largura = w - 4
    texto = str(texto)

    asc, desc = _ascent_descent(fonte) if fonte is not None else (0.95, -0.25)

    # ----- escolhe tamanho e quebra de linhas -----
    if multilinha:
        linhas = _quebrar_linhas(texto, fonte, tamanho, largura) if fonte else [texto]

        def cabe(tam, lns):
            return (
                len(lns) * tam * ESPACAMENTO_LINHAS <= h - 4
                and all(_largura_texto(fonte, _codificar(l), tam) <= largura for l in lns)
            )

        while fonte and tamanho > TAMANHO_MINIMO and not cabe(tamanho, linhas):
            tamanho -= 0.25
            linhas = _quebrar_linhas(texto, fonte, tamanho, largura)

        lead = tamanho * ESPACAMENTO_LINHAS
        y_inicial = h - 2 - lead
    else:
        linhas = [texto]

        while (
            fonte
            and tamanho > TAMANHO_MINIMO
            and _largura_texto(fonte, _codificar(texto), tamanho) > largura
        ):
            tamanho -= 0.25

        lead = 0
        # centralizado na vertical (mesma conta do modelo)
        y_inicial = (h - (asc - desc) * tamanho) / 2 - desc * tamanho

    # ----- monta a aparência -----
    corpo = (
        "/Tx BMC\nq\n"
        f"1 1 {w - 2:.4f} {h - 2:.4f} re\nW\nn\n"
        f"BT\n/{fonte_nome} {tamanho:g} Tf\n0 g\n"
    )

    for i, linha in enumerate(linhas or [""]):
        literal = _literal_pdf(_codificar(linha))
        if i == 0:
            corpo += f"2 {y_inicial:.4f} Td\n({literal}) Tj\n"
        else:
            corpo += f"0 -{lead:.4f} Td\n({literal}) Tj\n"

    corpo += "ET\nQ\nEMC\n"

    novo = DecodedStreamObject()
    for chave, valor in ap.items():
        if chave not in ("/Length", "/Filter", "/DecodeParms", "/Resources"):
            novo[NameObject(chave)] = valor

    novo[NameObject("/Resources")] = DictionaryObject({
        NameObject("/Font"): DictionaryObject({k: v for k, v in fontes.items()})
    })
    novo.set_data((prefixo + corpo).encode("latin-1"))

    anot["/AP"][NameObject("/N")] = writer._add_object(novo)


# ============================================================
# ACHATAR A PÁGINA (campos viram conteúdo fixo)
# ============================================================

def achatar_pagina(writer, pagina):
    """
    Desenha a aparência de cada campo diretamente na página e remove os
    campos. Nenhum visualizador recalcula nada depois disso, então o
    tamanho das letras fica exatamente como no PDF modelo.
    """
    fontes = _fontes_do_formulario(writer)

    resources = pagina.get("/Resources")
    if resources is None:
        resources = DictionaryObject()
        pagina[NameObject("/Resources")] = resources
    resources = resources.get_object()

    xobjects = resources.get("/XObject")
    if xobjects is None:
        xobjects = DictionaryObject()
        resources[NameObject("/XObject")] = xobjects
    xobjects = xobjects.get_object()

    desenho = []
    restantes = ArrayObject()

    for i, ref in enumerate(pagina.get("/Annots", []) or []):
        anot = ref.get_object()

        if anot.get("/Subtype") != "/Widget":
            restantes.append(ref)
            continue

        # campo oculto → não desenha
        if int(anot.get("/F", 0) or 0) & 2:
            continue

        ap = anot.get("/AP")
        if ap is None or "/N" not in ap:
            continue

        n_ref = ap.raw_get("/N")
        n = n_ref.get_object()

        if not hasattr(n, "get_data"):
            continue

        # garante que o stream é um Form XObject com recursos próprios
        n[NameObject("/Type")] = NameObject("/XObject")
        n[NameObject("/Subtype")] = NameObject("/Form")
        if "/Resources" not in n:
            n[NameObject("/Resources")] = DictionaryObject({
                NameObject("/Font"): DictionaryObject(
                    {k: v for k, v in fontes.items()})
            })

        # matriz que leva o BBox da aparência até o retângulo do campo
        r = [float(v) for v in anot["/Rect"]]
        rx0, rx1 = sorted((r[0], r[2]))
        ry0, ry1 = sorted((r[1], r[3]))

        bbox = [float(v) for v in n["/BBox"]]
        m = [float(v) for v in n.get("/Matrix", [1, 0, 0, 1, 0, 0])]

        cantos = [
            (bbox[0], bbox[1]), (bbox[2], bbox[1]),
            (bbox[0], bbox[3]), (bbox[2], bbox[3]),
        ]
        xs = [m[0] * x + m[2] * y + m[4] for x, y in cantos]
        ys = [m[1] * x + m[3] * y + m[5] for x, y in cantos]
        tx0, tx1, ty0, ty1 = min(xs), max(xs), min(ys), max(ys)

        if tx1 - tx0 <= 0 or ty1 - ty0 <= 0:
            continue

        sx = (rx1 - rx0) / (tx1 - tx0)
        sy = (ry1 - ry0) / (ty1 - ty0)
        ex = rx0 - tx0 * sx
        fy = ry0 - ty0 * sy

        nome = f"/FlatW{i}"
        xobjects[NameObject(nome)] = n_ref
        desenho.append(f"q {sx:.6f} 0 0 {sy:.6f} {ex:.6f} {fy:.6f} cm {nome} Do Q")

    # junta o conteúdo original com os campos desenhados
    original = pagina.get_contents()
    base = original.get_data() if original is not None else b""

    novo = DecodedStreamObject()
    novo.set_data(
        b"q\n" + base + b"\nQ\n" + "\n".join(desenho).encode("latin-1") + b"\n"
    )
    pagina[NameObject("/Contents")] = writer._add_object(novo)

    if restantes:
        pagina[NameObject("/Annots")] = restantes
    elif "/Annots" in pagina:
        del pagina["/Annots"]


# ============================================================
# TEXTOS DE UMA ETIQUETA
# ============================================================

def textos_da_etiqueta(etiqueta):
    """Dicionário papel → texto. Etiqueta None = posição vazia."""

    papeis = [
        "nome", "pix", "pre_pago", "seguro_valor",
        "a_vista", "total_21x", "parcela_12", "parcela_21",
        "plano", "plano_valor", "seguro_nome", "seguro_meses",
    ]

    if etiqueta is None:
        return {papel: "" for papel in papeis}

    entrada = texto_formatado(etiqueta["controle_entrada"])

    return {
        "nome": (
            str(etiqueta["nome_dispositivo"]).strip().upper()
            if NOME_EM_MAIUSCULAS
            else str(etiqueta["nome_dispositivo"]).strip()
        ),
        "pix": texto_formatado(etiqueta["valor_pix"]),
        "pre_pago": texto_formatado(etiqueta["valor_pre_pago"]),
        "seguro_valor": texto_formatado(etiqueta["valor_seguro"]),
        # Linha de cima: À vista (ao lado do "em 12x de")
        "a_vista": entrada,
        # Linha de baixo: total do parcelamento em 21x (ao lado do "em até 21x de")
        "total_21x": texto_formatado(etiqueta["total_21x"]),
        "parcela_12": texto_formatado(etiqueta["parcela_12x"]),
        "parcela_21": texto_formatado(etiqueta["parcela_21x"]),
        "plano": etiqueta["plano_vivo"],
        "plano_valor": texto_formatado(etiqueta["valor_plano_vivo"]),
        "seguro_nome": SEGURO_NOME,
        "seguro_meses": SEGURO_MESES,
    }


def definir_capacidade(arquivo_modelo):
    """Descobre quantas etiquetas cabem por página no PDF modelo."""
    global ETIQUETAS_POR_PAGINA

    reader = PdfReader(arquivo_modelo)
    ancoras = [
        w for w in _widgets_da_pagina(reader.pages[0])
        if w["nome"].upper().startswith("TL")
    ]

    if not ancoras:
        raise ValueError(
            "Este PDF não parece ser o modelo de etiquetas: "
            "nenhuma etiqueta foi encontrada."
        )

    ETIQUETAS_POR_PAGINA = len(ancoras)
    return ETIQUETAS_POR_PAGINA


# ============================================================
# GERA UMA PÁGINA PREENCHIDA (já achatada)
# ============================================================

def gerar_pagina_preenchida(arquivo_modelo, etiquetas_da_pagina):
    reader = PdfReader(arquivo_modelo)

    if not reader.get_fields():
        raise ValueError("O PDF modelo não possui campos preenchíveis.")

    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    pagina = writer.pages[0]

    mapa = mapear_etiquetas(pagina)

    for numero, papeis in enumerate(mapa):
        etiqueta = (
            etiquetas_da_pagina[numero]
            if numero < len(etiquetas_da_pagina)
            else None
        )

        textos = textos_da_etiqueta(etiqueta)

        for papel, texto in textos.items():
            widget = papeis[papel]

            # Campos fixos (plano, seguro, "12"): só mexe se o valor mudou
            if papel in ("plano", "seguro_nome", "seguro_meses"):
                if texto == _valor_campo(widget["anot"]):
                    continue

            definir_texto_campo(
                writer, widget, texto, multilinha=(papel == "nome")
            )

        # O campo "TL" (texto pequeno) não é tocado.

    achatar_pagina(writer, pagina)

    return writer, pagina


# ============================================================
# GERA PDF COMPLETO
# ============================================================

def gerar_pdf_completo(arquivo_modelo, arquivo_saida, etiquetas):
    if not etiquetas:
        raise ValueError("Nenhuma etiqueta foi adicionada.")

    definir_capacidade(arquivo_modelo)

    quantidade = len(etiquetas)
    quantidade_paginas = math.ceil(quantidade / ETIQUETAS_POR_PAGINA)

    writer_final = PdfWriter()

    for numero_pagina in range(quantidade_paginas):
        inicio = numero_pagina * ETIQUETAS_POR_PAGINA
        fim = min(inicio + ETIQUETAS_POR_PAGINA, quantidade)

        etiquetas_da_pagina = list(etiquetas[inicio:fim])
        while len(etiquetas_da_pagina) < ETIQUETAS_POR_PAGINA:
            etiquetas_da_pagina.append(None)

        _writer, pagina = gerar_pagina_preenchida(
            arquivo_modelo, etiquetas_da_pagina
        )

        writer_final.add_page(pagina)

    with open(arquivo_saida, "wb") as arquivo:
        writer_final.write(arquivo)


# ============================================================
# APLICAÇÃO
# ============================================================

# ============================================================
# TEMAS (claro / escuro / padrão do dispositivo)
# ============================================================

TEMAS = {
    "claro": {
        "bg": "#f3f3f3", "fg": "#1a1a1a", "campo": "#ffffff",
        "botao": "#e4e4e4", "botao_hover": "#d6d6d6", "borda": "#c4c4c4",
        "selecao": "#0a64c8", "selecao_fg": "#ffffff",
    },
    "escuro": {
        "bg": "#1f1f1f", "fg": "#e6e6e6", "campo": "#2b2b2b",
        "botao": "#363636", "botao_hover": "#444444", "borda": "#4a4a4a",
        "selecao": "#2f6fb5", "selecao_fg": "#ffffff",
    },
}

OPCOES_TEMA = {
    "Padrão do dispositivo": "sistema",
    "Claro": "claro",
    "Escuro": "escuro",
}


def tema_do_sistema():
    """Descobre se o dispositivo está no modo claro ou escuro."""
    try:
        if sys.platform.startswith("win"):
            import winreg
            chave = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            claro, _ = winreg.QueryValueEx(chave, "AppsUseLightTheme")
            return "claro" if claro else "escuro"

        if sys.platform == "darwin":
            r = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True, text=True,
            )
            return "escuro" if "Dark" in r.stdout else "claro"
    except Exception:
        pass

    return "claro"


def barra_de_titulo(janela, escuro):
    """Windows 10/11: deixa a barra de título da janela escura ou clara."""
    if not sys.platform.startswith("win"):
        return

    try:
        import ctypes

        janela.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
        valor = ctypes.c_int(1 if escuro else 0)

        for atributo in (20, 19):   # 20 = Win10 2004+/Win11, 19 = anteriores
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor)
            )
    except Exception:
        pass


class Aplicacao:

    def __init__(self, root):
        self.root = root
        self.root.title("Vivo - Gerador de Etiquetas")
        self.root.geometry("1100x800")
        self.root.minsize(850, 600)

        self.config = carregar_config()
        self.excel = ExcelRepository()
        self.etiquetas = []
        self.resultados = []
        self.indice_etiqueta_selecionada = 0

        self.cores_tema = TEMAS["claro"]
        self.tema_atual = "claro"

        self.criar_interface()
        self.carregar_caminhos_salvos()

        self.aplicar_tema()
        self.root.after(3000, self.vigiar_tema_do_sistema)

    # --------------------------------------------------------
    # INTERFACE
    # --------------------------------------------------------

    def criar_interface(self):
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(container, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            container, orient="vertical", command=self.canvas.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.frame_conteudo = ttk.Frame(self.canvas)
        self.window_canvas = self.canvas.create_window(
            (0, 0), window=self.frame_conteudo, anchor="nw"
        )

        self.frame_conteudo.bind("<Configure>", self.atualizar_scroll)
        self.canvas.bind("<Configure>", self.ajustar_largura_canvas)
        self.canvas.bind_all("<MouseWheel>", self.rolar_mouse)
        self.canvas.bind_all("<Button-4>", self.rolar_mouse_linux)
        self.canvas.bind_all("<Button-5>", self.rolar_mouse_linux)

        self.montar_conteudo()

    def montar_conteudo(self):
        frame = self.frame_conteudo

        # ---------------- TEMA ----------------
        barra_tema = ttk.Frame(frame)
        barra_tema.pack(fill="x", padx=15, pady=(10, 0))

        salvo = self.config.get("tema", "sistema")
        self.var_tema = tk.StringVar(value=next(
            (nome for nome, valor in OPCOES_TEMA.items() if valor == salvo),
            "Padrão do dispositivo",
        ))

        self.combo_tema = ttk.Combobox(
            barra_tema, textvariable=self.var_tema,
            values=list(OPCOES_TEMA), state="readonly", width=22,
        )
        self.combo_tema.pack(side="right")
        self.combo_tema.bind("<<ComboboxSelected>>", self.mudar_tema)
        ttk.Label(barra_tema, text="Tema:").pack(side="right", padx=(0, 6))

        ttk.Label(
            frame, text="Gerador de Etiquetas Vivo",
            font=("TkDefaultFont", 18, "bold"),
        ).pack(pady=(15, 5))

        ttk.Label(
            frame, text="Excel → cálculo automático → etiquetas → PDF"
        ).pack(pady=(0, 15))

        # ---------------- ARQUIVOS ----------------
        frame_arquivos = ttk.LabelFrame(frame, text="Arquivos")
        frame_arquivos.pack(fill="x", padx=15, pady=5)
        frame_arquivos.columnconfigure(1, weight=1)

        ttk.Label(frame_arquivos, text="Excel:").grid(
            row=0, column=0, padx=8, pady=8)
        self.var_excel = tk.StringVar()
        ttk.Entry(frame_arquivos, textvariable=self.var_excel).grid(
            row=0, column=1, sticky="ew", padx=5)
        ttk.Button(frame_arquivos, text="Selecionar",
                   command=self.selecionar_excel).grid(row=0, column=2, padx=5)
        ttk.Button(frame_arquivos, text="Carregar Excel",
                   command=self.carregar_excel).grid(row=0, column=3, padx=5)

        ttk.Label(frame_arquivos, text="PDF modelo:").grid(
            row=1, column=0, padx=8, pady=8)
        self.var_pdf = tk.StringVar()
        ttk.Entry(frame_arquivos, textvariable=self.var_pdf).grid(
            row=1, column=1, sticky="ew", padx=5)
        ttk.Button(frame_arquivos, text="Selecionar",
                   command=self.selecionar_pdf).grid(row=1, column=2, padx=5)

        # ---------------- CONFIGURAÇÕES ----------------
        frame_config = ttk.LabelFrame(frame, text="Configurações")
        frame_config.pack(fill="x", padx=15, pady=5)

        ttk.Label(frame_config, text="Plano:").grid(
            row=0, column=0, padx=8, pady=8)
        self.var_plano = tk.StringVar(value=self.config["plano_vivo"])
        ttk.Entry(frame_config, textvariable=self.var_plano, width=25).grid(
            row=0, column=1, padx=5)

        ttk.Label(frame_config, text="Valor plano:").grid(
            row=0, column=2, padx=8)
        self.var_valor_plano = tk.StringVar(
            value=str(self.config["valor_plano_vivo"]))
        ttk.Entry(frame_config, textvariable=self.var_valor_plano,
                  width=10).grid(row=0, column=3, padx=5)

        # ---------------- PESQUISA ----------------
        frame_pesquisa = ttk.LabelFrame(frame, text="Pesquisar aparelho")
        frame_pesquisa.pack(fill="x", padx=15, pady=5)

        self.var_busca = tk.StringVar()
        entrada_busca = ttk.Entry(frame_pesquisa, textvariable=self.var_busca)
        entrada_busca.pack(side="left", fill="x", expand=True, padx=8, pady=8)
        entrada_busca.bind("<KeyRelease>", lambda event: self.pesquisar())

        ttk.Button(frame_pesquisa, text="Pesquisar",
                   command=self.pesquisar).pack(side="left", padx=5)

        # ---------------- RESULTADOS ----------------
        frame_resultados = ttk.Frame(frame)
        frame_resultados.pack(fill="both", expand=True, padx=15, pady=5)
        frame_resultados.columnconfigure(0, weight=1)
        frame_resultados.columnconfigure(1, weight=1)

        frame_lista = ttk.LabelFrame(
            frame_resultados, text="Aparelhos encontrados")
        frame_lista.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.lista_modelos = tk.Listbox(frame_lista, height=10)
        self.lista_modelos.pack(fill="both", expand=True, padx=8, pady=8)
        self.lista_modelos.bind("<<ListboxSelect>>", self.selecionar_modelo)

        frame_dados = ttk.LabelFrame(frame_resultados, text="Dados do aparelho")
        frame_dados.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        frame_dados.columnconfigure(1, weight=1)

        self.var_nome = tk.StringVar()
        self.var_pre_pago = tk.StringVar()
        self.var_entrada = tk.StringVar()
        self.var_total_12x = tk.StringVar()
        self.var_total_21x = tk.StringVar()

        campos = [
            ("Nome:", self.var_nome),
            ("Pré-pago:", self.var_pre_pago),
            ("Controle Entrada:", self.var_entrada),
            ("Total em 12x:", self.var_total_12x),
            ("Total em 21x:", self.var_total_21x),
        ]

        for linha, (nome, variavel) in enumerate(campos):
            ttk.Label(frame_dados, text=nome).grid(
                row=linha, column=0, sticky="w", padx=8, pady=8)
            ttk.Entry(frame_dados, textvariable=variavel,
                      state="readonly").grid(
                row=linha, column=1, sticky="ew", padx=8, pady=8)

        ttk.Button(frame_dados, text="Calcular aparelho",
                   command=self.calcular_atual).grid(
            row=len(campos), column=0, columnspan=2, pady=12)

        # ---------------- ETIQUETAS ----------------
        frame_etiquetas = ttk.LabelFrame(frame, text="Etiquetas adicionadas")
        frame_etiquetas.pack(fill="x", padx=15, pady=5)

        self.lista_etiquetas = tk.Listbox(frame_etiquetas, height=10)
        self.lista_etiquetas.pack(
            side="left", fill="both", expand=True, padx=8, pady=8)
        self.lista_etiquetas.bind(
            "<<ListboxSelect>>", self.selecionar_etiqueta)

        frame_botoes = ttk.Frame(frame_etiquetas)
        frame_botoes.pack(side="right", padx=8)

        ttk.Button(frame_botoes, text="➕ Adicionar etiqueta",
                   command=self.adicionar_etiqueta).pack(fill="x", pady=3)
        ttk.Button(frame_botoes, text="🗑 Remover selecionada",
                   command=self.remover_etiqueta).pack(fill="x", pady=3)

        self.var_status = tk.StringVar(value="Nenhuma etiqueta adicionada.")
        ttk.Label(frame, textvariable=self.var_status,
                  font=("TkDefaultFont", 10, "bold")).pack(pady=5)

        # ---------------- BOTÕES PDF ----------------
        frame_pdf = ttk.Frame(frame)
        frame_pdf.pack(pady=(5, 20))

        ttk.Button(frame_pdf, text="👁 Pré-visualizar PDF",
                   command=self.previsualizar_pdf).pack(
            side="left", padx=8, ipadx=12, ipady=7)
        ttk.Button(frame_pdf, text="💾 Gerar PDF",
                   command=self.salvar_pdf).pack(
            side="left", padx=8, ipadx=12, ipady=7)

    # --------------------------------------------------------
    # TEMA
    # --------------------------------------------------------

    def aplicar_tema(self):
        escolha = self.config.get("tema", "sistema")
        nome = tema_do_sistema() if escolha == "sistema" else escolha
        if nome not in TEMAS:
            nome = "claro"

        c = TEMAS[nome]
        self.cores_tema = c
        self.tema_atual = nome

        estilo = ttk.Style(self.root)
        estilo.theme_use("clam")

        self.root.configure(bg=c["bg"])

        estilo.configure(
            ".", background=c["bg"], foreground=c["fg"],
            fieldbackground=c["campo"], bordercolor=c["borda"],
            lightcolor=c["bg"], darkcolor=c["bg"], troughcolor=c["bg"],
            insertcolor=c["fg"], focuscolor=c["bg"],
        )
        estilo.configure("TLabelframe", bordercolor=c["borda"])
        estilo.configure("TLabelframe.Label",
                         background=c["bg"], foreground=c["fg"])

        estilo.configure("TButton", background=c["botao"],
                         lightcolor=c["botao"], darkcolor=c["botao"],
                         bordercolor=c["borda"], focuscolor=c["botao"])
        estilo.map("TButton", background=[
            ("active", c["botao_hover"]), ("pressed", c["botao_hover"])])

        estilo.configure("TEntry", fieldbackground=c["campo"],
                         foreground=c["fg"], insertcolor=c["fg"])
        estilo.map("TEntry",
                   fieldbackground=[("readonly", c["botao"])],
                   foreground=[("readonly", c["fg"])])

        estilo.configure("TCombobox", fieldbackground=c["campo"],
                         background=c["botao"], foreground=c["fg"],
                         arrowcolor=c["fg"])
        estilo.map("TCombobox",
                   fieldbackground=[("readonly", c["campo"])],
                   foreground=[("readonly", c["fg"])],
                   selectbackground=[("readonly", c["campo"])],
                   selectforeground=[("readonly", c["fg"])])

        estilo.configure("Vertical.TScrollbar", background=c["botao"],
                         lightcolor=c["botao"], darkcolor=c["botao"],
                         troughcolor=c["bg"], bordercolor=c["borda"],
                         arrowcolor=c["fg"])

        # Widgets "clássicos" do Tk (não seguem o estilo ttk)
        self.canvas.configure(bg=c["bg"])

        for lista in (self.lista_modelos, self.lista_etiquetas):
            lista.configure(
                bg=c["campo"], fg=c["fg"],
                selectbackground=c["selecao"], selectforeground=c["selecao_fg"],
                highlightbackground=c["borda"], highlightcolor=c["borda"],
            )

        # Lista que se abre no seletor de tema
        try:
            popdown = self.root.tk.eval(
                f"ttk::combobox::PopdownWindow {self.combo_tema}")
            self.root.tk.call(
                f"{popdown}.f.l", "configure",
                "-background", c["campo"], "-foreground", c["fg"],
                "-selectbackground", c["selecao"],
                "-selectforeground", c["selecao_fg"],
            )
        except Exception:
            pass

        barra_de_titulo(self.root, nome == "escuro")

    def mudar_tema(self, event=None):
        self.config["tema"] = OPCOES_TEMA.get(self.var_tema.get(), "sistema")
        salvar_config(self.config)
        self.aplicar_tema()
        self.combo_tema.selection_clear()

    def vigiar_tema_do_sistema(self):
        """No modo 'Padrão do dispositivo', acompanha mudanças do Windows."""
        if (self.config.get("tema", "sistema") == "sistema"
                and tema_do_sistema() != self.tema_atual):
            self.aplicar_tema()

        self.root.after(3000, self.vigiar_tema_do_sistema)

    # --------------------------------------------------------
    # SCROLL
    # --------------------------------------------------------

    def atualizar_scroll(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def ajustar_largura_canvas(self, event):
        self.canvas.itemconfigure(self.window_canvas, width=event.width)

    def rolar_mouse(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def rolar_mouse_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(3, "units")

    # --------------------------------------------------------
    # CONFIGURAÇÕES SALVAS
    # --------------------------------------------------------

    def carregar_caminhos_salvos(self):
        self.var_excel.set(self.config.get("arquivo_excel", ""))
        self.var_pdf.set(self.config.get("arquivo_pdf_modelo", ""))

        try:
            caminho = self.config.get("arquivo_pdf_modelo", "")
            if caminho and Path(caminho).exists():
                definir_capacidade(caminho)
        except Exception:
            pass

    def salvar_config_atual(self):
        try:
            valor_plano = float(self.var_valor_plano.get().replace(",", "."))
        except ValueError:
            raise ValueError("Valor do plano inválido.")

        self.config["plano_vivo"] = self.var_plano.get()
        self.config["valor_plano_vivo"] = valor_plano
        self.config["arquivo_excel"] = self.var_excel.get()
        self.config["arquivo_pdf_modelo"] = self.var_pdf.get()

        salvar_config(self.config)

    # --------------------------------------------------------
    # SELECIONAR ARQUIVOS
    # --------------------------------------------------------

    def selecionar_excel(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"),
                       ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.var_excel.set(caminho)
            self.salvar_config_atual()

    def selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar PDF modelo",
            filetypes=[("PDF", "*.pdf")],
        )
        if caminho:
            self.var_pdf.set(caminho)
            self.salvar_config_atual()

            try:
                definir_capacidade(caminho)
                self.atualizar_lista_etiquetas()
            except Exception as erro:
                messagebox.showerror("PDF modelo", str(erro))

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    def carregar_excel(self):
        caminho = self.var_excel.get().strip()

        if not caminho:
            messagebox.showwarning(
                "Excel", "Selecione um arquivo Excel primeiro.")
            return

        try:
            quantidade = self.excel.carregar(caminho)
            self.resultados = self.excel.pesquisar_modelo("")
            self.atualizar_lista_modelos()
            self.salvar_config_atual()

            messagebox.showinfo(
                "Excel carregado",
                f"{quantidade} aparelhos carregados com sucesso.",
            )
        except Exception as erro:
            messagebox.showerror("Erro ao carregar Excel", str(erro))

    def pesquisar(self):
        self.resultados = self.excel.pesquisar_modelo(self.var_busca.get())
        self.atualizar_lista_modelos()

    def atualizar_lista_modelos(self):
        self.lista_modelos.delete(0, tk.END)
        for linha in self.resultados:
            self.lista_modelos.insert(tk.END, linha["Nome Comercial"])

    def selecionar_modelo(self, event=None):
        selecao = self.lista_modelos.curselection()
        if not selecao:
            return
        self.preencher_dados(self.resultados[selecao[0]])

    def preencher_dados(self, linha):
        self.var_nome.set(linha.get("Nome Comercial", ""))
        self.var_pre_pago.set(texto_formatado(
            decimal_brasileiro(linha.get("Pré-pago", 0))))
        self.var_entrada.set(texto_formatado(
            decimal_brasileiro(linha.get("Controle Entrada", 0))))
        self.var_total_12x.set(texto_formatado(
            decimal_brasileiro(linha.get("Total 12x", 0))))
        self.var_total_21x.set(texto_formatado(
            decimal_brasileiro(linha.get("Total 21x", 0))))

    # --------------------------------------------------------
    # CÁLCULO
    # --------------------------------------------------------

    def obter_aparelho_atual(self):
        return Aparelho(
            self.var_nome.get(),
            decimal_brasileiro(self.var_pre_pago.get()),
            decimal_brasileiro(self.var_entrada.get()),
            decimal_brasileiro(self.var_total_12x.get()),
            decimal_brasileiro(self.var_total_21x.get()),
        )

    def calcular_atual(self):
        if not self.var_nome.get():
            messagebox.showwarning(
                "Aparelho", "Selecione um aparelho primeiro.")
            return None

        try:
            aparelho = self.obter_aparelho_atual()
            valor_plano = decimal_brasileiro(self.var_valor_plano.get())

            resultado = calcular_etiqueta(
                aparelho, self.var_plano.get(), valor_plano
            )

            self.mostrar_resultado(resultado)
            return resultado

        except Exception as erro:
            messagebox.showerror("Erro no cálculo", str(erro))
            return None

    def mostrar_resultado(self, resultado):
        janela = tk.Toplevel(self.root)
        janela.title("Resultado do cálculo")
        janela.geometry("450x560")

        c = self.cores_tema
        janela.configure(bg=c["bg"])
        barra_de_titulo(janela, self.tema_atual == "escuro")

        texto = tk.Text(
            janela, wrap="word", bg=c["campo"], fg=c["fg"],
            insertbackground=c["fg"], relief="flat", highlightthickness=1,
            highlightbackground=c["borda"], highlightcolor=c["borda"],
        )
        texto.pack(fill="both", expand=True, padx=10, pady=10)

        texto.insert(
            "1.0",
            f"APARELHO\n{resultado['nome_dispositivo']}\n\n"
            f"PRÉ-PAGO\n{formatar_reais(resultado['valor_pre_pago'])}\n\n"
            f"CONTROLE ENTRADA\n{formatar_reais(resultado['controle_entrada'])}\n\n"
            f"SEGURO\n{formatar_reais(resultado['valor_seguro'])}\n\n"
            f"PIX\n{formatar_reais(resultado['valor_pix'])}\n\n"
            f"12X DE\n{formatar_reais(resultado['parcela_12x'])}"
            f"  (total {formatar_reais(resultado['total_12x'])})\n\n"
            f"21X DE\n{formatar_reais(resultado['parcela_21x'])}"
            f"  (total {formatar_reais(resultado['total_21x'])})\n\n"
            f"PLANO\n{resultado['plano_vivo']}\n"
            f"{formatar_reais(resultado['valor_plano_vivo'])}",
        )
        texto.config(state="disabled")

        ttk.Button(janela, text="Fechar", command=janela.destroy).pack(pady=10)

    # --------------------------------------------------------
    # ETIQUETAS
    # --------------------------------------------------------

    def selecionar_etiqueta(self, event=None):
        selecao = self.lista_etiquetas.curselection()
        if selecao:
            self.indice_etiqueta_selecionada = selecao[0]

    def adicionar_etiqueta(self):
        resultado = self.calcular_atual()
        if resultado is None:
            return

        self.etiquetas.append(resultado)
        indice = len(self.etiquetas) - 1

        self.atualizar_lista_etiquetas()

        self.lista_etiquetas.selection_clear(0, tk.END)
        self.lista_etiquetas.selection_set(indice)
        self.lista_etiquetas.see(indice)
        self.lista_etiquetas.activate(indice)
        self.indice_etiqueta_selecionada = indice

        self.var_status.set(
            f"Etiqueta {indice + 1} adicionada: "
            f"{resultado['nome_dispositivo']}."
        )

    def remover_etiqueta(self):
        if not self.etiquetas:
            return

        selecao = self.lista_etiquetas.curselection()
        indice = selecao[0] if selecao else len(self.etiquetas) - 1

        etiqueta = self.etiquetas[indice]
        del self.etiquetas[indice]

        self.atualizar_lista_etiquetas()

        if self.etiquetas:
            novo = min(indice, len(self.etiquetas) - 1)
            self.lista_etiquetas.selection_set(novo)
            self.lista_etiquetas.see(novo)
            self.indice_etiqueta_selecionada = novo
        else:
            self.indice_etiqueta_selecionada = 0

        self.var_status.set(
            f"Etiqueta removida: {etiqueta['nome_dispositivo']}.")

    def atualizar_lista_etiquetas(self):
        self.lista_etiquetas.delete(0, tk.END)
        quantidade = len(self.etiquetas)

        if quantidade == 0:
            self.lista_etiquetas.insert(tk.END, "Nenhuma etiqueta adicionada.")
        else:
            for indice, etiqueta in enumerate(self.etiquetas):
                pagina = indice // ETIQUETAS_POR_PAGINA + 1
                posicao = indice % ETIQUETAS_POR_PAGINA + 1
                self.lista_etiquetas.insert(
                    tk.END,
                    f"{indice + 1:03d}  |  Página {pagina} • "
                    f"posição {posicao}  |  {etiqueta['nome_dispositivo']}",
                )

        paginas = math.ceil(quantidade / ETIQUETAS_POR_PAGINA) if quantidade else 0

        self.var_status.set(
            f"{quantidade} etiqueta(s) | {paginas} página(s) no PDF | "
            f"{ETIQUETAS_POR_PAGINA} por página"
        )

    # --------------------------------------------------------
    # VALIDAÇÃO / PDF
    # --------------------------------------------------------

    def validar_para_pdf(self):
        arquivo_modelo = self.var_pdf.get().strip()

        if not arquivo_modelo:
            raise ValueError("Selecione o PDF modelo primeiro.")

        if not Path(arquivo_modelo).exists():
            raise FileNotFoundError(
                "O arquivo PDF modelo não foi encontrado.")

        definir_capacidade(arquivo_modelo)

        if not self.etiquetas:
            raise ValueError("Nenhuma etiqueta foi adicionada.")

        return arquivo_modelo, len(self.etiquetas)

    def previsualizar_pdf(self):
        try:
            arquivo_modelo, quantidade = self.validar_para_pdf()

            tmp = tempfile.NamedTemporaryFile(
                prefix="vivo_preview_", suffix=".pdf", delete=False)
            caminho_preview = Path(tmp.name)
            tmp.close()

            gerar_pdf_completo(
                arquivo_modelo, caminho_preview, self.etiquetas.copy())

            abrir_arquivo(caminho_preview)

            paginas = math.ceil(quantidade / ETIQUETAS_POR_PAGINA)
            self.var_status.set(
                f"Pré-visualização aberta: {quantidade} etiqueta(s) "
                f"em {paginas} página(s)."
            )
        except Exception as erro:
            messagebox.showerror("Erro na pré-visualização", str(erro))

    def salvar_pdf(self):
        try:
            arquivo_modelo, quantidade = self.validar_para_pdf()
            self.salvar_config_atual()

            paginas = math.ceil(quantidade / ETIQUETAS_POR_PAGINA)

            arquivo_saida = filedialog.asksaveasfilename(
                title="Salvar PDF preenchido",
                defaultextension=".pdf",
                initialfile="etiquetas_vivo_preenchidas.pdf",
                filetypes=[("PDF", "*.pdf")],
            )

            if not arquivo_saida:
                return

            gerar_pdf_completo(
                arquivo_modelo, arquivo_saida, self.etiquetas.copy())

            if messagebox.askyesno(
                "PDF gerado",
                "PDF gerado com sucesso!\n\n"
                f"Etiquetas: {quantidade}\n"
                f"Páginas: {paginas}\n\n"
                "Deseja abrir o PDF agora?",
            ):
                abrir_arquivo(arquivo_saida)

            self.var_status.set(
                f"PDF gerado: {quantidade} etiqueta(s), {paginas} página(s).")

        except Exception as erro:
            messagebox.showerror("Erro ao gerar PDF", str(erro))


# ============================================================
# MAIN
# ============================================================

def main():
    root = tk.Tk()
    Aplicacao(root)
    root.mainloop()


if __name__ == "__main__":
    main()
