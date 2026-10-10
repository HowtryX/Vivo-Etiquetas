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
from pypdf.generic import NumberObject, NameObject, TextStringObject

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

try:
    import fitz  # PyMuPDF  (pip install pymupdf)
except ImportError:
    fitz = None

try:
    from reportlab.pdfbase.pdfmetrics import stringWidth  # pip install reportlab
except ImportError:
    stringWidth = None


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
    "taxa_cartao_mes": 1.00,
    "plano_vivo": "Controle 12GB",
    "valor_plano_vivo": 57.00,
    "arquivo_excel": "",
    "arquivo_pdf_modelo": "",
}

# Excel
ABA_EXCEL = "SMARTPHONES"
COL_NOME = "E"
COL_PRE_PAGO = "F"
COL_ENTRADA = "AX"

# PDF
ETIQUETAS_POR_PAGINA = 10
CAMPOS_POR_ETIQUETA = 13
TOTAL_CAMPOS_POR_PAGINA = ETIQUETAS_POR_PAGINA * CAMPOS_POR_ETIQUETA

PDF_FLAG_MULTILINE = 4096  # bit 13

# ------------------------------------------------------------
# AJUSTES DO CAMPO "NOME DO APARELHO" (campo 6)
# ------------------------------------------------------------
FONTE_NOME_METRICA = "Helvetica"   # fonte usada só para MEDIR a largura do texto
TAMANHO_NOME = 8                   # tamanho da letra do nome (ajuste se precisar)
MARGEM_CAMPO = 4                   # margem interna (pontos) descontada da largura


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
# PARCELAMENTO PRICE
# ============================================================

def calcular_parcela_price(valor, quantidade_parcelas, taxa_mensal):
    valor = moeda(valor)
    taxa = Decimal(str(taxa_mensal)) / Decimal("100")

    if quantidade_parcelas <= 0:
        raise ValueError("Quantidade de parcelas inválida.")

    if taxa == 0:
        return moeda(valor / quantidade_parcelas)

    fator = (Decimal("1") + taxa) ** quantidade_parcelas
    parcela = valor * (taxa * fator) / (fator - Decimal("1"))
    return moeda(parcela)


# ============================================================
# APARELHO
# ============================================================

class Aparelho:

    def __init__(self, modelo, pre_pago, controle_entrada):
        self.modelo = modelo
        self.pre_pago = moeda(pre_pago)
        self.controle_entrada = moeda(controle_entrada)


# ============================================================
# CÁLCULO COMPLETO DA ETIQUETA
# ============================================================

def calcular_etiqueta(aparelho, taxa_cartao, plano_vivo, valor_plano):
    return {
        "nome_dispositivo": aparelho.modelo,
        "valor_pix": calcular_pix(aparelho.controle_entrada),
        "valor_pre_pago": aparelho.pre_pago,
        "valor_seguro": calcular_seguro(aparelho.controle_entrada),
        "controle_entrada": aparelho.controle_entrada,
        "parcela_12x": calcular_parcela_price(
            aparelho.controle_entrada, 12, taxa_cartao
        ),
        "parcela_21x": calcular_parcela_price(
            aparelho.controle_entrada, 21, taxa_cartao
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
# MAPA DOS CAMPOS DA ETIQUETA
# ============================================================

def preencher_campos_etiqueta(grupo, valores):
    # O campo 5 (grupo[4]) NÃO é alterado.
    return {
        grupo[0]: valores["pre_pago"],
        grupo[1]: "Vivo Seguro Celular",
        grupo[2]: valores["seguro"],
        grupo[3]: "10x",
        grupo[5]: valores["nome"],
        grupo[6]: valores["plano"],
        grupo[7]: valores["valor_plano"],
        grupo[8]: valores["entrada"],
        grupo[9]: valores["entrada"],
        grupo[10]: valores["parcela_12"],
        grupo[11]: valores["parcela_21"],
        grupo[12]: valores["pix"],
    }


# ============================================================
# UTILITÁRIOS DE WIDGET / CAMPO
# ============================================================

def _nome_widget(anotacao):
    nome = anotacao.get("/T")
    if nome is None and anotacao.get("/Parent") is not None:
        nome = anotacao["/Parent"].get_object().get("/T")
    return nome


def _widgets_por_nome(pagina, nome_campo):
    """Gera (widget, pai_ou_None) de todos os widgets com esse nome."""
    for ref in pagina.get("/Annots", []) or []:
        try:
            anot = ref.get_object()
            if anot.get("/Subtype") != "/Widget":
                continue
            if _nome_widget(anot) != nome_campo:
                continue
            pai = anot.get("/Parent")
            yield anot, (pai.get_object() if pai is not None else None)
        except Exception:
            continue


def configurar_multiline(pagina, writer, nome_campo):
    """Liga o flag Multiline somente no campo informado."""
    for anot, pai in _widgets_por_nome(pagina, nome_campo):
        alvo = anot if anot.get("/Ff") is not None or pai is None else pai
        flags = int(alvo.get("/Ff", 0) or 0) | PDF_FLAG_MULTILINE
        alvo[NameObject("/Ff")] = NumberObject(flags)


def obter_largura_campo(pagina, nome_campo):
    for anot, _ in _widgets_por_nome(pagina, nome_campo):
        x0, y0, x1, y1 = [float(v) for v in anot["/Rect"]]
        return abs(x1 - x0) - MARGEM_CAMPO
    return None


def fixar_fonte_campo(pagina, nome_campo, tamanho=TAMANHO_NOME):
    """
    Troca SOMENTE o tamanho da fonte no /DA, mantendo o nome da fonte
    original do campo. Evita que o visualizador use tamanho 'auto'.
    """
    for anot, pai in _widgets_por_nome(pagina, nome_campo):
        for obj in (anot, pai):
            if obj is None:
                continue

            da = obj.get("/DA")
            da = str(da) if da is not None else ""

            if re.search(r"/\S+\s+[\d.]+\s+Tf", da):
                novo = re.sub(
                    r"(/\S+)\s+[\d.]+\s+Tf",
                    rf"\1 {tamanho} Tf",
                    da,
                    count=1,
                )
            else:
                novo = f"/Helv {tamanho} Tf 0 g"

            obj[NameObject("/DA")] = TextStringObject(novo)


def quebrar_texto(texto, largura, tamanho=TAMANHO_NOME):
    """Quebra o texto em linhas que caibam na largura do campo."""
    texto = str(texto).strip()

    if stringWidth is None or not largura:
        return texto

    linhas, atual = [], ""

    for palavra in texto.split():
        teste = f"{atual} {palavra}".strip()
        if stringWidth(teste, FONTE_NOME_METRICA, tamanho) <= largura:
            atual = teste
        else:
            if atual:
                linhas.append(atual)
            atual = palavra

    if atual:
        linhas.append(atual)

    return "\n".join(linhas)


# ============================================================
# ACHATAR (FLATTEN) PDF
# ============================================================

def achatar_pdf(origem, destino):
    """
    Transforma os campos preenchíveis em conteúdo fixo, para o Windows
    (Edge/Adobe/etc.) não recalcular a aparência e mudar o tamanho da letra.
    Retorna True se achatou, False se não foi possível.
    """
    if fitz is None:
        return False

    try:
        doc = fitz.open(str(origem))
        doc.bake(annots=True, widgets=True)
        doc.save(str(destino))
        doc.close()
        return True
    except Exception:
        return False


# ============================================================
# GERA UMA PÁGINA PREENCHIDA
# ============================================================

def gerar_pagina_preenchida(arquivo_modelo, etiquetas_da_pagina):
    reader = PdfReader(arquivo_modelo)
    campos_dict = reader.get_fields()

    if not campos_dict:
        raise ValueError("O PDF modelo não possui campos preenchíveis.")

    campos = list(campos_dict.keys())

    if len(campos) != TOTAL_CAMPOS_POR_PAGINA:
        raise ValueError(
            "O PDF modelo não possui a quantidade esperada de campos.\n\n"
            f"Encontrados: {len(campos)}\n"
            f"Esperados: {TOTAL_CAMPOS_POR_PAGINA}"
        )

    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    pagina = writer.pages[0]

    todos_os_dados = {}

    for numero_etiqueta in range(ETIQUETAS_POR_PAGINA):
        etiqueta = etiquetas_da_pagina[numero_etiqueta]

        if etiqueta is None:
            continue

        inicio = numero_etiqueta * CAMPOS_POR_ETIQUETA
        grupo = campos[inicio:inicio + CAMPOS_POR_ETIQUETA]

        valores = {
            "nome": etiqueta["nome_dispositivo"],
            "pre_pago": texto_formatado(etiqueta["valor_pre_pago"]),
            "seguro": texto_formatado(etiqueta["valor_seguro"]),
            "entrada": texto_formatado(etiqueta["controle_entrada"]),
            "parcela_12": texto_formatado(etiqueta["parcela_12x"]),
            "parcela_21": texto_formatado(etiqueta["parcela_21x"]),
            "pix": texto_formatado(etiqueta["valor_pix"]),
            "plano": etiqueta["plano_vivo"],
            "valor_plano": texto_formatado(etiqueta["valor_plano_vivo"]),
        }

        dados = preencher_campos_etiqueta(grupo, valores)

        # --- Campo 6: nome do aparelho -----------------------------
        campo_nome = grupo[5]

        configurar_multiline(pagina, writer, campo_nome)
        fixar_fonte_campo(pagina, campo_nome)

        largura = obter_largura_campo(pagina, campo_nome)
        dados[campo_nome] = quebrar_texto(valores["nome"], largura)

        todos_os_dados.update(dados)

    writer.update_page_form_field_values(
        pagina,
        {k: ("" if v is None else str(v)) for k, v in todos_os_dados.items()},
        auto_regenerate=False,
    )

    try:
        writer.set_need_appearances_writer(False)
    except Exception:
        pass

    tmp = tempfile.NamedTemporaryFile(
        prefix="vivo_pagina_", suffix=".pdf", delete=False
    )
    caminho = Path(tmp.name)
    tmp.close()

    with open(caminho, "wb") as arquivo:
        writer.write(arquivo)

    return caminho


# ============================================================
# GERA PDF COMPLETO
# ============================================================

def gerar_pdf_completo(arquivo_modelo, arquivo_saida, etiquetas):
    if not etiquetas:
        raise ValueError("Nenhuma etiqueta foi adicionada.")

    quantidade = len(etiquetas)
    quantidade_paginas = math.ceil(quantidade / ETIQUETAS_POR_PAGINA)

    temporarios = []
    paginas_finais = []

    try:
        for numero_pagina in range(quantidade_paginas):
            inicio = numero_pagina * ETIQUETAS_POR_PAGINA
            fim = min(inicio + ETIQUETAS_POR_PAGINA, quantidade)

            etiquetas_da_pagina = list(etiquetas[inicio:fim])
            while len(etiquetas_da_pagina) < ETIQUETAS_POR_PAGINA:
                etiquetas_da_pagina.append(None)

            caminho_pagina = gerar_pagina_preenchida(
                arquivo_modelo, etiquetas_da_pagina
            )
            temporarios.append(caminho_pagina)

            # Achata cada página ANTES de juntar (o AcroForm ainda existe aqui)
            caminho_achatado = caminho_pagina.with_name(
                caminho_pagina.stem + "_flat.pdf"
            )
            if achatar_pdf(caminho_pagina, caminho_achatado):
                temporarios.append(caminho_achatado)
                paginas_finais.append(caminho_achatado)
            else:
                paginas_finais.append(caminho_pagina)

        # ---- Junta as páginas ----
        if fitz is not None:
            saida = fitz.open()
            for caminho in paginas_finais:
                with fitz.open(str(caminho)) as parte:
                    saida.insert_pdf(parte)
            saida.save(str(arquivo_saida))
            saida.close()
        else:
            writer_final = PdfWriter()
            for caminho in paginas_finais:
                writer_final.add_page(PdfReader(caminho).pages[0])
            with open(arquivo_saida, "wb") as arquivo:
                writer_final.write(arquivo)

    finally:
        for caminho in temporarios:
            try:
                Path(caminho).unlink(missing_ok=True)
            except OSError:
                pass


# ============================================================
# APLICAÇÃO
# ============================================================

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

        self.criar_interface()
        self.carregar_caminhos_salvos()

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

        ttk.Label(frame_config, text="Taxa cartão (% a.m.):").grid(
            row=0, column=0, padx=8, pady=8)
        self.var_taxa = tk.StringVar(
            value=str(self.config["taxa_cartao_mes"]))
        ttk.Entry(frame_config, textvariable=self.var_taxa, width=10).grid(
            row=0, column=1, padx=5)

        ttk.Label(frame_config, text="Plano:").grid(row=0, column=2, padx=8)
        self.var_plano = tk.StringVar(value=self.config["plano_vivo"])
        ttk.Entry(frame_config, textvariable=self.var_plano, width=25).grid(
            row=0, column=3, padx=5)

        ttk.Label(frame_config, text="Valor plano:").grid(
            row=0, column=4, padx=8)
        self.var_valor_plano = tk.StringVar(
            value=str(self.config["valor_plano_vivo"]))
        ttk.Entry(frame_config, textvariable=self.var_valor_plano,
                  width=10).grid(row=0, column=5, padx=5)

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

        campos = [
            ("Nome:", self.var_nome),
            ("Pré-pago:", self.var_pre_pago),
            ("Controle Entrada:", self.var_entrada),
        ]

        for linha, (nome, variavel) in enumerate(campos):
            ttk.Label(frame_dados, text=nome).grid(
                row=linha, column=0, sticky="w", padx=8, pady=8)
            ttk.Entry(frame_dados, textvariable=variavel,
                      state="readonly").grid(
                row=linha, column=1, sticky="ew", padx=8, pady=8)

        ttk.Button(frame_dados, text="Calcular aparelho",
                   command=self.calcular_atual).grid(
            row=3, column=0, columnspan=2, pady=12)

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

    def salvar_config_atual(self):
        try:
            taxa = float(self.var_taxa.get().replace(",", "."))
            valor_plano = float(self.var_valor_plano.get().replace(",", "."))
        except ValueError:
            raise ValueError("Taxa do cartão ou valor do plano inválido.")

        self.config["taxa_cartao_mes"] = taxa
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

    # --------------------------------------------------------
    # CÁLCULO
    # --------------------------------------------------------

    def obter_aparelho_atual(self):
        return Aparelho(
            self.var_nome.get(),
            decimal_brasileiro(self.var_pre_pago.get()),
            decimal_brasileiro(self.var_entrada.get()),
        )

    def calcular_atual(self):
        if not self.var_nome.get():
            messagebox.showwarning(
                "Aparelho", "Selecione um aparelho primeiro.")
            return None

        try:
            aparelho = self.obter_aparelho_atual()
            taxa = Decimal(self.var_taxa.get().replace(",", "."))
            valor_plano = decimal_brasileiro(self.var_valor_plano.get())

            resultado = calcular_etiqueta(
                aparelho, taxa, self.var_plano.get(), valor_plano
            )

            self.mostrar_resultado(resultado)
            return resultado

        except Exception as erro:
            messagebox.showerror("Erro no cálculo", str(erro))
            return None

    def mostrar_resultado(self, resultado):
        janela = tk.Toplevel(self.root)
        janela.title("Resultado do cálculo")
        janela.geometry("450x500")

        texto = tk.Text(janela, wrap="word")
        texto.pack(fill="both", expand=True, padx=10, pady=10)

        texto.insert(
            "1.0",
            f"APARELHO\n{resultado['nome_dispositivo']}\n\n"
            f"PRÉ-PAGO\n{formatar_reais(resultado['valor_pre_pago'])}\n\n"
            f"CONTROLE ENTRADA\n{formatar_reais(resultado['controle_entrada'])}\n\n"
            f"SEGURO\n{formatar_reais(resultado['valor_seguro'])}\n\n"
            f"PIX\n{formatar_reais(resultado['valor_pix'])}\n\n"
            f"12X\n{formatar_reais(resultado['parcela_12x'])}\n\n"
            f"21X\n{formatar_reais(resultado['parcela_21x'])}\n\n"
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
