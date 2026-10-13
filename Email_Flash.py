import smtplib
import imaplib
import os
import json
import threading
import tkinter as tk
from tkinter import messagebox, ttk, filedialog, scrolledtext
from email.message import EmailMessage
from email.utils import make_msgid
from datetime import datetime, timezone
import re
import binascii

try:
    import keyring
    HAS_KEYRING = True
except ImportError:
    HAS_KEYRING = False


# ===== ARQUIVO DE CONFIGURAÇÕES =====
PASTA_APP_DADOS = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "EmailFlash")
os.makedirs(PASTA_APP_DADOS, exist_ok=True)
CONFIG_FILE = os.path.join(PASTA_APP_DADOS, "config.json")
SERVICE_NAME = "EmailFlash_App"

CONFIG_PADRAO = {
    "titulo_app": "Automação de E-mail",
    "email_remetente": "",
    "senha_email": "",
    "smtp_servidor": "",
    "smtp_porta": "465",
    "smtp_tipo_conexao": "SSL",
    "imap_servidor": "",
    "imap_porta": "993",
    "email_cc": "",
    "email_cco": "",
    "pasta_nfs": "",
    "pasta_boletos": "",
    "arquivo_assinatura": "",
    "assunto_template": "NF {NFS} {BOLETO_SUBJECT} - {CLIENTE}",
    "corpo_template": (
        "Olá,\n\nSegue anexa a {NF_MSG} {NFS} {BOLETO_MSG}.\n\n"
        "Qualquer dúvida estou à disposição!\n\n"
        "Atenciosamente,"
    )
}

def encode_mutf7(text: str) -> bytes:
    """Implementa o Modified UTF-7 exigido pelo protocolo IMAP."""
    res = bytearray()
    b = []
    for c in text:
        if c == '&':
            if b:
                b64 = binascii.b2a_base64("".join(b).encode('utf-16-be')).replace(b'/', b',').rstrip(b'\n=')
                res.extend(b'&' + b64 + b'-')
                b = []
            res.extend(b'&-')
        elif 0x20 <= ord(c) <= 0x7E:
            if b:
                b64 = binascii.b2a_base64("".join(b).encode('utf-16-be')).replace(b'/', b',').rstrip(b'\n=')
                res.extend(b'&' + b64 + b'-')
                b = []
            res.append(ord(c))
        else:
            b.append(c)
    if b:
        b64 = binascii.b2a_base64("".join(b).encode('utf-16-be')).replace(b'/', b',').rstrip(b'\n=')
        res.extend(b'&' + b64 + b'-')
    return bytes(res)


def _get_senha(config: dict) -> str:
    if HAS_KEYRING and config.get("email_remetente"):
        try:
            senha = keyring.get_password(SERVICE_NAME, config["email_remetente"])
            if senha:
                return senha
        except Exception:
            pass
    return config.get("senha_email", "")


def _set_senha(config: dict, senha: str):
    if HAS_KEYRING and config.get("email_remetente"):
        try:
            keyring.set_password(SERVICE_NAME, config["email_remetente"], senha)
            config["senha_email"] = ""
            return
        except Exception:
            pass
    config["senha_email"] = senha


def carregar_config():
    dados = CONFIG_PADRAO.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                dados_salvos = json.load(f)
                dados.update(dados_salvos)
        except Exception:
            pass
            
    for k, v in CONFIG_PADRAO.items():
        if k not in dados:
            dados[k] = v
            
    # Migra senha para o keyring se estiver solta no JSON
    if HAS_KEYRING and dados.get("senha_email") and dados.get("email_remetente"):
        _set_senha(dados, dados["senha_email"])
        salvar_config(dados)

    return dados


def salvar_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


# === CLASSE PRINCIPAL =========================================================
class FaturamentoApp:

    def __init__(self, master):
        self.master = master
        self.config = carregar_config()
        master.title(self._get_config_str("titulo_app", "Automação de E-mail"))
        master.geometry("780x560")
        master.resizable(True, True)
        master.minsize(600, 460)

        self.nf_var      = tk.StringVar()
        self.nf_sem_boleto_var = tk.StringVar()
        self.nome_var    = tk.StringVar()
        self.email_var   = tk.StringVar()

        self.v_email_rem    = tk.StringVar(value=self._get_config_str("email_remetente"))
        self.v_senha        = tk.StringVar(value=_get_senha(self.config))
        self.v_smtp_serv    = tk.StringVar(value=self._get_config_str("smtp_servidor"))
        self.v_smtp_porta   = tk.StringVar(value=self._get_config_str("smtp_porta"))
        self.v_smtp_tipo    = tk.StringVar(value=self._get_config_str("smtp_tipo_conexao", "SSL"))
        self.v_imap_serv    = tk.StringVar(value=self._get_config_str("imap_servidor"))
        self.v_imap_porta   = tk.StringVar(value=self._get_config_str("imap_porta"))
        self.v_cc           = tk.StringVar(value=self._get_config_str("email_cc"))
        self.v_cco          = tk.StringVar(value=self._get_config_str("email_cco"))

        self.v_pasta_nfs       = tk.StringVar(value=self._get_config_str("pasta_nfs"))
        self.v_pasta_boletos   = tk.StringVar(value=self._get_config_str("pasta_boletos"))
        self.v_assinatura      = tk.StringVar(value=self._get_config_str("arquivo_assinatura"))

        self._configurar_estilos()

        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        self.aba1 = ttk.Frame(self.notebook, padding=15)
        self.aba2 = ttk.Frame(self.notebook, padding=15)
        self.aba3 = ttk.Frame(self.notebook, padding=15)
        self.aba4 = ttk.Frame(self.notebook, padding=15)
        self.aba5 = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.aba1, text="  📧 Envio  ")
        self.notebook.add(self.aba2, text="  ✏️ Mensagem  ")
        self.notebook.add(self.aba3, text="  📊 Envio em Lote  ")
        self.notebook.add(self.aba4, text="  📁 Config. Gerais  ")
        self.notebook.add(self.aba5, text="  ⚙️ Config. E-mail  ")

        self._criar_aba1()
        self._criar_aba2()
        self._criar_aba3()
        self._criar_aba4()
        self._criar_aba5()

        ttk.Label(master, text="© Todos os direitos reservados", style="Copyright.TLabel").pack(side=tk.BOTTOM, pady=4)

    # -------------------------------------------------------------------------
    def _configurar_estilos(self):
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10, "bold"))
        style.configure("Head.TLabel", font=("Helvetica", 10, "bold"), foreground="#0056b3")
        style.configure("Green.TLabel", font=("Helvetica", 10, "bold"), foreground="green")
        style.configure("Copyright.TLabel", font=("Helvetica", 8), foreground="gray")
        style.configure("Status.TLabel", font=("Helvetica", 10))

    def _criar_campo(self, parent, row, label_text, var, show="", width=45, is_readonly=False):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky='w', padx=5, pady=(5, 2))
        state = 'readonly' if is_readonly else 'normal'
        entry = ttk.Entry(parent, textvariable=var, show=show, width=width, state=state)
        entry.grid(row=row, column=1, columnspan=2, sticky='ew', pady=(2, 5), padx=(0, 10))
        return entry

    def _get_config_str(self, chave: str, default: str = "") -> str:
        return str(self.config.get(chave, default))

    def _get_config_int(self, chave: str, default: int) -> int:
        try:
            return int(self.config.get(chave, default))
        except ValueError:
            return default

    @staticmethod
    def _validar_email_simples(email: str) -> bool:
        email = email.strip()
        if not email: return False
        return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email))

    @staticmethod
    def _validar_lista_emails(texto: str) -> bool:
        texto = texto.strip()
        if not texto: return True
        return all(FaturamentoApp._validar_email_simples(e) for e in texto.split(',') if e.strip())

    # =========================================================================
    # ABA 1 – ENVIO MANUAL
    # =========================================================================
    def _criar_aba1(self):
        f = self.aba1
        f.grid_columnconfigure(1, weight=1)

        ttk.Label(f, text="Orientações de entrada:", style="Head.TLabel").grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 4))
        for i, (pre, verde, pos) in enumerate([
            ("1. Digite o número da NF e a série (Ex.:", "99999-1", ")"),
            ("2. Para mais de uma NF, separe por vírgula (Ex.:", "99999-1, 9999-2", ")"),
            ("3. Use o campo abaixo para NFs sem boleto (Ex.: bonificação, à vista)", "", "")
        ], start=1):
            fr = ttk.Frame(f)
            fr.grid(row=i, column=0, columnspan=3, sticky='w', padx=2, pady=1)
            ttk.Label(fr, text=pre).grid(row=0, column=0, sticky='w')
            if verde:
                ttk.Label(fr, text=verde, style="Green.TLabel").grid(row=0, column=1, sticky='w')
                ttk.Label(fr, text=pos).grid(row=0, column=2, sticky='w')

        self.nf_entry            = self._criar_campo(f, 4, "Número(s) da NF:", self.nf_var)
        self.nf_sem_boleto_entry = self._criar_campo(f, 5, "NF's sem boleto:", self.nf_sem_boleto_var)
        self.nome_entry          = self._criar_campo(f, 6, "Nome do Cliente:", self.nome_var)
        self.email_entry         = self._criar_campo(f, 7, "E-mail do Cliente:", self.email_var)

        self.btn_enviar = ttk.Button(f, text="Enviar E-mail", command=self.processar_envio)
        self.btn_enviar.grid(row=8, column=0, columnspan=3, pady=14, sticky='ew', padx=5)

        self.status_label = ttk.Label(f, text="Pronto para o envio.", foreground="blue", style="Status.TLabel")
        self.status_label.grid(row=9, column=0, columnspan=3, sticky='ew', pady=4)
        self.nf_entry.focus()

    # =========================================================================
    # ABA 2 – CONFIGURAÇÃO DA MENSAGEM
    # =========================================================================
    def _criar_aba2(self):
        f = self.aba2
        f.grid_columnconfigure(1, weight=1)
        f.grid_rowconfigure(3, weight=1)

        ttk.Label(f, text="Configure o template da mensagem de e-mail:", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 6))
        info = "Variáveis:\n{NFS} = NF(s) | {CLIENTE} = cliente | {NF_MSG} = Singular/Plural | {BOLETO_MSG} = Complemento boleto | {BOLETO_SUBJECT} = + BOLETOS"
        ttk.Label(f, text=info, foreground="#555555", font=("Helvetica", 9), justify='left').grid(row=1, column=0, columnspan=2, sticky='w', pady=(0, 10))

        ttk.Label(f, text="Assunto:").grid(row=2, column=0, sticky='w', padx=5, pady=(0, 4))
        self.assunto_entry = ttk.Entry(f, width=70)
        self.assunto_entry.insert(0, self._get_config_str("assunto_template", CONFIG_PADRAO["assunto_template"]))
        self.assunto_entry.grid(row=2, column=1, sticky='ew', pady=(0, 4), padx=(0, 5))

        ttk.Label(f, text="Corpo:").grid(row=3, column=0, sticky='nw', padx=5, pady=(4, 0))
        self.corpo_text = scrolledtext.ScrolledText(f, width=60, height=12, font=("Helvetica", 10), wrap=tk.WORD)
        self.corpo_text.grid(row=3, column=1, sticky='nsew', padx=(0, 5), pady=(4, 6))
        self.corpo_text.insert("1.0", self._get_config_str("corpo_template", CONFIG_PADRAO["corpo_template"]))

        ttk.Button(f, text="💾 Salvar Mensagem", command=self._salvar_mensagem).grid(row=4, column=0, columnspan=2, pady=8, sticky='ew', padx=5)

    def _salvar_mensagem(self):
        self.config["assunto_template"] = self.assunto_entry.get().strip()
        self.config["corpo_template"]   = self.corpo_text.get("1.0", tk.END).strip()
        salvar_config(self.config)
        messagebox.showinfo("Salvo", "Template salvo!")

    # =========================================================================
    # ABA 3 – ENVIO EM LOTE (EXCEL)
    # =========================================================================
    def _criar_aba3(self):
        f = self.aba3
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(3, weight=1)

        ttk.Label(f, text="Envio em Lote via Planilha Excel:", style="Head.TLabel").grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 4))
        ttk.Label(f, text="Colunas Linha 1: SERIE NF | NUMERO NF | CLIENTE | EMAIL", foreground="#555555", font=("Helvetica", 9)).grid(row=1, column=0, columnspan=3, sticky='w', pady=(0, 8))

        fr_arq = ttk.Frame(f)
        fr_arq.grid(row=2, column=0, columnspan=3, sticky='ew', pady=(0, 8))
        fr_arq.grid_columnconfigure(1, weight=1)

        self.v_excel_path = tk.StringVar()
        ttk.Label(fr_arq, text="Arquivo Excel:").grid(row=0, column=0, sticky='w', padx=(0, 5))
        ttk.Entry(fr_arq, textvariable=self.v_excel_path, state='readonly', width=50).grid(row=0, column=1, sticky='ew', padx=(0, 6))
        ttk.Button(fr_arq, text="📂 Selecionar", command=self._selecionar_excel).grid(row=0, column=2)

        cols = ("SERIE NF", "NUMERO NF", "CLIENTE", "EMAIL")
        self.tree = ttk.Treeview(f, columns=cols, show='headings', height=10)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80 if "NF" in col else 180, anchor='w')
        self.tree.grid(row=3, column=0, columnspan=3, sticky='nsew', pady=(0, 8))

        scroll_y = ttk.Scrollbar(f, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=scroll_y.set)
        scroll_y.grid(row=3, column=3, sticky='ns')

        self.sem_boleto_lote_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Todas as NFs da planilha são SEM BOLETO", variable=self.sem_boleto_lote_var).grid(row=4, column=0, columnspan=3, sticky='w', pady=(0, 6))
        self.btn_enviar_lote = ttk.Button(f, text="🚀 Enviar Todos", command=self._enviar_lote)
        self.btn_enviar_lote.grid(row=5, column=0, columnspan=3, pady=6, sticky='ew')
        self.status_lote = ttk.Label(f, text="", foreground="blue", style="Status.TLabel")
        self.status_lote.grid(row=6, column=0, columnspan=3, sticky='ew', pady=4)

    def _selecionar_excel(self):
        caminho = filedialog.askopenfilename(title="Selecionar Planilha", filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
        if caminho:
            self.v_excel_path.set(caminho)
            self._carregar_preview_excel(caminho)

    def _carregar_preview_excel(self, caminho):
        try:
            if caminho.lower().endswith(".xlsx"):
                import openpyxl
                linhas = list(openpyxl.load_workbook(caminho, data_only=True).active.iter_rows(values_only=True))
            else:
                import xlrd
                wb = xlrd.open_workbook(caminho)
                ws = wb.sheet_by_index(0)
                linhas = [ws.row_values(r) for r in range(ws.nrows)]

            for row in self.tree.get_children(): self.tree.delete(row)

            for linha in linhas[1:]:
                if len(linha) >= 4:
                    vals = [str(v).strip().removesuffix(".0") if v else "" for v in linha[:4]]
                    if any(vals): self.tree.insert("", tk.END, values=vals)
        except Exception as e:
            messagebox.showerror("Erro Planilha", f"Dependência ou erro:\n{e}")

    def _enviar_lote(self):
        itens = self.tree.get_children()
        if not itens: return messagebox.showwarning("Aviso", "Carregue os dados.")
        
        dados = [self.tree.item(i, "values") for i in itens]
        self.btn_enviar_lote.config(state='disabled')
        self.status_lote.config(text="Iniciando...", foreground="gray")
        
        threading.Thread(target=self._enviar_lote_thread, args=(dados, self.sem_boleto_lote_var.get()), daemon=True).start()

    def _enviar_lote_thread(self, dados, sem_boleto):
        total, sucessos, erros = len(dados), 0, []
        for idx, (serie, num, cliente, email) in enumerate(dados, 1):
            if not all([serie, num, cliente, email]):
                erros.append(f"Linha {idx}: dados incompletos.")
                continue

            self.master.after(0, lambda i=idx, c=cliente: self.status_lote.config(text=f"Enviando {i}/{total}: {c}...", foreground="gray"))
            try:
                arq_anexos, missing, _ = self._localizar_arquivos_da_nf(f"{num}-{serie}", sem_boleto)
                if missing:
                    erros.append(f"Linha {idx} ({cliente}): Arquivo não encontrado: {missing}")
                    continue
                    
                self.enviar_email(cliente, email, num, arq_anexos, _silencioso=True)
                sucessos += 1
            except Exception as e:
                erros.append(f"Linha {idx} ({cliente}): {str(e)}")

        self.master.after(0, lambda: self._finalizar_lote_ui(sucessos, total, erros))

    def _finalizar_lote_ui(self, sucessos, total, erros):
        resumo = f"✅ Lote concluído: {sucessos}/{total} enviados."
        if erros:
            resumo += f" | ❌ {len(erros)} erro(s)."
            messagebox.showwarning("Aviso Lote", "\n".join(erros[:15]) + ("\n..." if len(erros) > 15 else ""))
        else:
            messagebox.showinfo("Lote concluído", f"{sucessos} e-mail(s) enviados!")
        self.status_lote.config(text=resumo, foreground="green" if not erros else "orange")
        self.btn_enviar_lote.config(state='normal')

    # =========================================================================
    # ABA 4 – CONFIGURAÇÕES GERAIS
    # =========================================================================
    def _criar_aba4(self):
        f = self.aba4
        f.grid_columnconfigure(1, weight=1)
        ttk.Label(f, text="Configurações de Caminhos e Arquivos:", style="Head.TLabel").grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 12))

        comps = [
            ("Arquivo de Assinatura (.jpg/.png):", self.v_assinatura, lambda: self._browse_file("arquivo_assinatura", self.v_assinatura, [("Imagens", "*.jpg *.jpeg *.png"), ("Todos", "*.*")])),
            ("Pasta das NFs:", self.v_pasta_nfs, lambda: self._browse_folder("pasta_nfs", self.v_pasta_nfs)),
            ("Pasta dos Boletos:", self.v_pasta_boletos, lambda: self._browse_folder("pasta_boletos", self.v_pasta_boletos)),
        ]
        for i, (lbl, var, cmd) in enumerate(comps, 1):
            ttk.Label(f, text=lbl).grid(row=i*2-1, column=0, columnspan=3, sticky='w', padx=5, pady=(10, 2))
            fr = ttk.Frame(f)
            fr.grid(row=i*2, column=0, columnspan=3, sticky='ew', padx=5, pady=(0, 4))
            fr.grid_columnconfigure(0, weight=1)
            ttk.Entry(fr, textvariable=var, state='readonly', width=55).grid(row=0, column=0, sticky='ew', padx=(0, 6))
            ttk.Button(fr, text="📂 Selecionar", command=cmd).grid(row=0, column=1)

        ttk.Button(f, text="💾 Salvar Caminhos", command=self._salvar_caminhos).grid(row=10, column=0, columnspan=3, pady=20, sticky='ew', padx=5)

    def _browse_file(self, chave, var, tipos):
        caminho = filedialog.askopenfilename(title="Selecionar arquivo", filetypes=tipos)
        if caminho:
            var.set(caminho)
            self.config[chave] = caminho

    def _browse_folder(self, chave, var):
        caminho = filedialog.askdirectory(title="Selecionar pasta")
        if caminho:
            var.set(caminho)
            self.config[chave] = caminho

    def _salvar_caminhos(self):
        self.config["arquivo_assinatura"] = self.v_assinatura.get()
        self.config["pasta_nfs"]          = self.v_pasta_nfs.get()
        self.config["pasta_boletos"]      = self.v_pasta_boletos.get()
        salvar_config(self.config)
        messagebox.showinfo("Salvo", "Caminhos salvos!")

    # =========================================================================
    # ABA 5 – CONFIGURAÇÕES DE E-MAIL
    # =========================================================================
    def _criar_aba5(self):
        f = self.aba5
        f.grid_columnconfigure(1, weight=1)
        ttk.Label(f, text="Configurações de E-mail:", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        self._criar_campo(f, 1, "E-mail Remetente:", self.v_email_rem)
        self._criar_campo(f, 2, "Senha do E-mail:", self.v_senha, show="*")
        self._criar_campo(f, 3, "Servidor SMTP:", self.v_smtp_serv)
        self._criar_campo(f, 4, "Porta SMTP:", self.v_smtp_porta)
        
        ttk.Label(f, text="Conexão SMTP:").grid(row=5, column=0, sticky='w', padx=5, pady=(5, 2))
        cb_tipo = ttk.Combobox(f, textvariable=self.v_smtp_tipo, values=["SSL", "STARTTLS"], state="readonly", width=42)
        cb_tipo.grid(row=5, column=1, sticky='ew', pady=(2, 5), padx=(0, 10))

        self._criar_campo(f, 6, "Servidor IMAP:", self.v_imap_serv)
        self._criar_campo(f, 7, "Porta IMAP:", self.v_imap_porta)
        self._criar_campo(f, 8, "CC:", self.v_cc)
        self._criar_campo(f, 9, "CCO:", self.v_cco)

        ttk.Button(f, text="💾 Salvar Configurações", command=self._salvar_email_config).grid(row=10, column=0, columnspan=2, pady=18, sticky='ew', padx=5)

    def _salvar_email_config(self):
        if not self._validar_lista_emails(self.v_cc.get()):
            return messagebox.showerror("Inválido", "O campo CC contém e-mail inválido.")
        if not self._validar_lista_emails(self.v_cco.get()):
            return messagebox.showerror("Inválido", "O campo CCO contém e-mail inválido.")

        self.config["email_remetente"] = self.v_email_rem.get().strip()
        self.config["smtp_servidor"]   = self.v_smtp_serv.get().strip()
        self.config["smtp_porta"]      = self.v_smtp_porta.get().strip()
        self.config["smtp_tipo_conexao"] = self.v_smtp_tipo.get().strip()
        self.config["imap_servidor"]   = self.v_imap_serv.get().strip()
        self.config["imap_porta"]      = self.v_imap_porta.get().strip()
        self.config["email_cc"]        = self.v_cc.get().strip()
        self.config["email_cco"]       = self.v_cco.get().strip()
        
        _set_senha(self.config, self.v_senha.get())
        salvar_config(self.config)
        messagebox.showinfo("Salvo", "Configurações salvas!")

    # =========================================================================
    # LÓGICA DE UNIFICAÇÃO (Processamento de Arquivos)
    # =========================================================================
    def _localizar_arquivos_da_nf(self, nf_item_str, sem_boleto_global=False):
        """Retorna (lista_de_anexos, missing_str, nf_numero)."""
        partes = nf_item_str.strip().upper().split('-')
        if len(partes) < 2:
            raise ValueError(f"Formato inválido: '{nf_item_str}'")

        nf_numero, nf_serie = partes[0], partes[1]
        sem_boleto = sem_boleto_global or (len(partes) == 3 and partes[2] == 'SBOLETO')

        NF_BUSCA = f"danfe_{nf_serie}_{nf_numero}"
        BOLETO_BUSCA = None if sem_boleto else f"boleto nf {int(nf_numero):08d}"

        nf_path = os.path.join(self._get_config_str("pasta_nfs"), f"{NF_BUSCA}.pdf")
        arquivos = []
        
        if not os.path.exists(nf_path):
            return [], f"{NF_BUSCA}.pdf", nf_numero

        arquivos.append({'path': nf_path, 'name': f"NF_S{nf_serie}_{nf_numero}.pdf"})

        if BOLETO_BUSCA:
            boleto_path = os.path.join(self._get_config_str("pasta_boletos"), f"{BOLETO_BUSCA}.pdf")
            if os.path.exists(boleto_path):
                arquivos.append({'path': boleto_path, 'name': f"Boleto_NF_{nf_numero}.pdf"})
                
        return arquivos, None, nf_numero

    def atualizar_status(self, mensagem, cor="blue"):
        self.master.after(0, lambda: self.status_label.config(text=mensagem, foreground=cor))

    def processar_envio(self):
        self.atualizar_status("Processando...", "gray")
        if not self._get_config_str("pasta_nfs") or not self._get_config_str("pasta_boletos"):
            return self._erro_e_retorno("Configure as pastas na Aba de Configurações.")

        nf_input, nf_sem_boleto_input, cliente_nome, cliente_email = self.nf_var.get().strip(), self.nf_sem_boleto_var.get().strip(), self.nome_var.get().strip(), self.email_var.get().strip()

        if not (nf_input or nf_sem_boleto_input) or not all([cliente_nome, cliente_email]):
            return self._erro_e_retorno("Preencha ao menos um campo de NF, além de Nome e E-mail.")
        if not self._validar_email_simples(cliente_email):
            return self._erro_e_retorno("O E-mail do cliente possui formato inválido.")

        arquivos_para_anexar, faltantes, nums = [], [], []

        for item in [n.strip() for n in nf_input.split(',') if n.strip()]:
            try:
                anexos, missing, num = self._localizar_arquivos_da_nf(item, False)
                if missing: faltantes.append(missing)
                else: 
                    arquivos_para_anexar.extend(anexos)
                    nums.append(num)
            except Exception as e:
                return self._erro_e_retorno(str(e))

        for item in [n.strip() for n in nf_sem_boleto_input.split(',') if n.strip()]:
            try:
                anexos, missing, num = self._localizar_arquivos_da_nf(item, True)
                if missing: faltantes.append(missing)
                else: 
                    arquivos_para_anexar.extend(anexos)
                    nums.append(num)
            except Exception as e:
                return self._erro_e_retorno(str(e))

        if faltantes: return self._erro_e_retorno("NF(s) não encontrada(s):\n" + "\n".join(faltantes))
        if not arquivos_para_anexar: return self._erro_e_retorno("Nenhuma NF localizada.")

        self.btn_enviar.config(state='disabled')
        threading.Thread(target=self._enviar_email_thread, args=(cliente_nome, cliente_email, ", ".join(nums), arquivos_para_anexar), daemon=True).start()

    def _enviar_email_thread(self, cliente_nome, cliente_email, nfs_string, arquivos_para_anexar):
        try:
            self.enviar_email(cliente_nome, cliente_email, nfs_string, arquivos_para_anexar)
        except Exception as e:
            self._erro_e_retorno(f"Erro ao enviar: {str(e)}")

    # =========================================================================
    # ENVIO DE E-MAIL SMTP / IMAP
    # =========================================================================
    def enviar_email(self, cliente_nome, cliente_email, nfs_string, arquivos, _silencioso=False):
        msg = EmailMessage()
        msg["From"] = self._get_config_str("email_remetente")
        msg["To"]   = cliente_email
        if self._get_config_str("email_cc"): msg["Cc"] = self._get_config_str("email_cc")
        if self._get_config_str("email_cco"): msg["Bcc"] = self._get_config_str("email_cco")

        is_multi = len(nfs_string.split(',')) > 1
        tem_bol = any("Boleto" in a['name'] for a in arquivos)

        bol_sub = "+ BOLETOS" if is_multi else ("+ BOLETO" if tem_bol else "")
        msg_nf = "Notas Fiscais nº" if is_multi else "Nota Fiscal nº"
        msg_bol = "e respectivos boletos" if is_multi and tem_bol else ("e o respectivo boleto" if tem_bol else "")

        def sub(txt): return txt.replace("{NFS}", nfs_string).replace("{CLIENTE}", cliente_nome).replace("{BOLETO_SUBJECT}", bol_sub).replace("{NF_MSG}", msg_nf).replace("{BOLETO_MSG}", msg_bol)

        msg["Subject"] = sub(self._get_config_str("assunto_template", CONFIG_PADRAO["assunto_template"]))
        corpo_base = sub(self._get_config_str("corpo_template", CONFIG_PADRAO["corpo_template"]))
        
        cid = make_msgid()[1:-1]
        msg.set_content(corpo_base)
        msg.add_alternative(f"<html><body>{''.join(f'<p>{l}</p>' if l.strip() else '<br>' for l in corpo_base.splitlines())}<img src='cid:{cid}' style='margin-top:10px;max-width:300px;height:auto;'></body></html>", subtype='html')

        html_part = next((p for p in msg.walk() if p.get_content_type() == 'text/html'), None)
        ass_path = self._get_config_str("arquivo_assinatura")
        
        if html_part and ass_path and os.path.exists(ass_path):
            try:
                ext = ass_path.split('.')[-1].lower()
                with open(ass_path, 'rb') as f:
                    html_part.add_related(f.read(), 'image', 'jpeg' if ext in ('jpg', 'jpeg') else ext, cid=f"<{cid}>")
            except Exception as e:
                self.atualizar_status(f"⚠️ Erro na assinatura: {e}", "orange")

        for anexo in arquivos:
            try:
                with open(anexo['path'], "rb") as f:
                    msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=anexo['name'])
            except FileNotFoundError:
                raise FileNotFoundError(f"Arquivo não localizado no disco: {anexo['name']}")
            except PermissionError:
                raise PermissionError(f"Sem permissão para ler o arquivo: {anexo['name']}")

        senha = _get_senha(self.config)
        porta = self._get_config_int("smtp_porta", 465)
        srv = self._get_config_str("smtp_servidor")

        try:
            if self._get_config_str("smtp_tipo_conexao") == "STARTTLS":
                with smtplib.SMTP(srv, porta) as smtp:
                    smtp.starttls()
                    smtp.login(self._get_config_str("email_remetente"), senha)
                    smtp.send_message(msg)
            else:
                with smtplib.SMTP_SSL(srv, porta) as smtp:
                    smtp.login(self._get_config_str("email_remetente"), senha)
                    smtp.send_message(msg)

            imap_salvo = self._salvar_copia_imap(msg, senha)
            if not _silencioso: self._limpar_e_finalizar(cliente_nome, nfs_string, imap_salvo)

        except smtplib.SMTPAuthenticationError:
            err = "Falha de Autenticação SMTP. Verifique e-mail/senha."
            if _silencioso: raise PermissionError(err)
            self._erro_e_retorno(err)
        except Exception as e:
            if _silencioso: raise
            self._erro_e_retorno(f"Erro SMTP: {e}")

    def _salvar_copia_imap(self, msg, senha):
        pastas = ['Sent', 'INBOX.Sent', 'Enviados', 'INBOX.Enviados', '[Gmail]/Sent', 'sent-mail', 'Sent Items', 'OUTBOX']
        try:
            with imaplib.IMAP4_SSL(self._get_config_str("imap_servidor"), self._get_config_int("imap_porta", 993)) as imap:
                imap.login(self._get_config_str("email_remetente"), senha)
                dt_str = imaplib.Time2Internaldate(datetime.now(timezone.utc))
                b_msg = msg.as_bytes()

                for pasta in pastas:
                    try:
                        # Utiliza a codificação UTF-7 modificada customizada (sem o codec legado)
                        if imap.append(encode_mutf7(pasta), '\\Seen', dt_str, b_msg)[0] == 'OK': return True
                    except imaplib.IMAP4.error:
                        continue
                raise Exception("Pasta não localizada no servidor IMAP.")
        except imaplib.IMAP4.error as e:
            self.atualizar_status(f"❌ Erro de protocolo IMAP: {e}", "red")
            return False
        except Exception as e:
            self.atualizar_status(f"❌ Erro salvando IMAP: {e}", "red")
            return False

    def _limpar_e_finalizar(self, cliente_nome, nfs_string, imap_salvo):
        st = f"{'✅ Enviado e salvo' if imap_salvo else '⚠️ Enviado sem cópia IMAP'}: {cliente_nome} - NF {nfs_string}"
        self.atualizar_status(st, "green" if imap_salvo else "orange")
        def _f():
            if imap_salvo: messagebox.showinfo("Sucesso", "E-mail enviado e cópia IMAP salva!")
            else: messagebox.showwarning("Aviso IMAP", "Enviado, mas falha ao salvar nos Enviados.")
            self.nf_var.set(""); self.nf_sem_boleto_var.set(""); self.nome_var.set(""); self.email_var.set("")
            self.btn_enviar.config(state='normal'); self.nf_entry.focus()
        self.master.after(0, _f)

    def _erro_e_retorno(self, msg):
        self.atualizar_status("❌ Falha no Envio.", "red")
        self.master.after(0, lambda: messagebox.showerror("Erro", msg))
        self.master.after(0, lambda: self.btn_enviar.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = FaturamentoApp(root)
    root.mainloop()