from datetime import datetime, timedelta
import json
import customtkinter as ctk
import tkinter as tk
import os
import requests

import requests

# ===== Aparência e tema =====
ctk.set_appearance_mode("Dark")            # força dark mode
ctk.set_default_color_theme("dark-blue")   # base: escuro com azul

# Paleta customizada
REFRESH_URL = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authentication/actions/refreshToken"
TOKEN_DIR = "tokens"
TOKEN_FILE = os.path.join(TOKEN_DIR, "tokens.json")
ACCENT = "#7C83FF"         # azul arroxeado
ACCENT_DARK = "#626BFF"
ERROR = "#FF6464"
SUCCESS = "#66D19E"
FG_MUTED = "#A9ABB6"
CARD_BG = ("#1D1F27", "#1B1D24")           # gradiente leve simulado
BG_TOP = "#0E1015"                          # fundo mais escuro
BG_BOTTOM = "#141722"

# ===== Tooltip simples =====
class ToolTip:
    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self._id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)

    def _schedule(self, _=None):
        self._id = self.widget.after(self.delay, self._show)

    def _show(self):
        if self.tipwindow or not self.text:
            return
        x, y, _, h = self.widget.bbox("insert") if self.widget == self.widget.focus_get() else (0, 0, 0, 0)
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.configure(bg="#000000", padx=1, pady=1)
        frame = tk.Frame(tw, bg="#111216")
        frame.pack()
        label = tk.Label(
            frame, text=self.text, bg="#111216", fg="#D7DAE0",
            font=("Segoe UI", 9), padx=8, pady=5, justify="left"
        )
        label.pack()

        # posiciona
        tw.wm_geometry(f"+{x}+{y}")

    def _hide(self, _=None):
        if self._id:
            self.widget.after_cancel(self._id)
            self._id = None
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


class TokenApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Token Refresher")
        self.geometry("560x360")
        self.minsize(520, 340)
        self.resizable(True, True)
        self._hidden = True
        self.configure(fg_color=BG_TOP)

        # gradiente de fundo “simulado” com duas frames
        self._bg_top = ctk.CTkFrame(self, fg_color=BG_TOP, corner_radius=0)
        self._bg_top.pack(fill="both", expand=True)
        self._bg_bottom = ctk.CTkFrame(self._bg_top, fg_color=BG_BOTTOM, corner_radius=0)
        self._bg_bottom.place(relx=0, rely=1, anchor="sw", relwidth=1, relheight=0.55)

        self._build_ui()
        self._center_window()

        # atalhos globais
        self.bind_all("<Return>", self._on_enter)
        self.bind_all("<Escape>", lambda e: self.destroy())
        self.bind_all("<Control-v>", self._paste_clipboard)
        self.bind_all("<Control-c>", self._copy_token)
        self.bind_all("<Control-k>", self._clear_entry)

    # ===== UI =====
    def _build_ui(self):
        # Header (titulo + subtitulo)
        header = ctk.CTkFrame(self._bg_top, fg_color="transparent")
        header.pack(fill="x", padx=22, pady=(18, 6))

        title_font = ctk.CTkFont(family="Segoe UI Semibold", size=18, weight="bold")
        subtitle_font = ctk.CTkFont(family="Segoe UI", size=12)

        title = ctk.CTkLabel(
            header,
            text="🔐 Token Refresher",
            font=title_font,
            text_color="#E8EAF0"
        )
        title.pack(anchor="w")

        subtitle = ctk.CTkLabel(
            header,
            text="Cole seu token com segurança. Pressione ⏎ para confirmar.",
            font=subtitle_font,
            text_color=FG_MUTED
        )
        subtitle.pack(anchor="w", pady=(2, 0))

        # Card principal (efeito glass leve)
        card = ctk.CTkFrame(
            self._bg_top,
            corner_radius=14,
            fg_color=CARD_BG[0]
        )
        card.pack(fill="x", padx=22, pady=(8, 14))
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)

        label_font = ctk.CTkFont(family="Segoe UI", size=13)
        self.label = ctk.CTkLabel(
            card, text="Informe o token", font=label_font, text_color="#DADDE6"
        )
        self.label.grid(row=0, column=0, sticky="w", padx=14, pady=(14, 6))

        # Linha de entrada + botões
        entry_frame = ctk.CTkFrame(card, fg_color=("transparent"))
        entry_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        entry_frame.grid_columnconfigure(0, weight=1)
        entry_frame.grid_columnconfigure(1, weight=0)
        entry_frame.grid_columnconfigure(2, weight=0)
        entry_frame.grid_columnconfigure(3, weight=0)

        self.entry = ctk.CTkEntry(
            entry_frame,
            placeholder_text="Cole ou digite seu token aqui",
            height=36,
            show="•",
            border_width=1,
            corner_radius=10,
        )
        self.entry.grid(row=0, column=0, sticky="ew", padx=(4, 8), pady=8)
        self.entry.bind("<KeyRelease>", self._on_entry_change)
        self.entry.focus_set()

        # Botão mostrar/ocultar
        self.toggle_btn = ctk.CTkButton(
            entry_frame,
            text="👁 Mostrar",
            width=86,
            fg_color=ACCENT,
            hover_color=ACCENT_DARK,
            corner_radius=10,
            command=self._toggle_visibility
        )
        self.toggle_btn.grid(row=0, column=1, padx=(0, 6), pady=8)
        ToolTip(self.toggle_btn, "Mostrar/ocultar token")

        # Botão copiar
        self.copy_btn = ctk.CTkButton(
            entry_frame,
            text="📋 Copiar",
            width=86,
            fg_color="#2A2E3A",
            hover_color="#34394A",
            command=self._copy_token
        )
        self.copy_btn.grid(row=0, column=2, padx=(0, 6), pady=8)
        ToolTip(self.copy_btn, "Copiar token para a área de transferência (Ctrl+C)")

        # Botão limpar
        self.clear_btn = ctk.CTkButton(
            entry_frame,
            text="🧹 Limpar",
            width=86,
            fg_color="#2A2E3A",
            hover_color="#34394A",
            command=self._clear_entry
        )
        self.clear_btn.grid(row=0, column=3, padx=(0, 4), pady=8)
        ToolTip(self.clear_btn, "Limpar campo (Ctrl+K)")

        # Barra de qualidade (indicativa pelo tamanho)
        bar_row = ctk.CTkFrame(card, fg_color="transparent")
        bar_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 6))
        bar_row.grid_columnconfigure(0, weight=1)
        self.strength_bar = ctk.CTkProgressBar(bar_row, height=10, corner_radius=6, progress_color=ACCENT)
        self.strength_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))
        self.strength_bar.set(0)

        self.strength_label = ctk.CTkLabel(
            bar_row, text="", text_color=FG_MUTED, font=ctk.CTkFont(size=11)
        )
        self.strength_label.grid(row=1, column=0, sticky="w", padx=6, pady=(2, 8))

        # Linha final: Confirmar
        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=0)

        self.button = ctk.CTkButton(
            actions,
            text="⏎ Confirmar",
            height=38,
            width=120,
            fg_color=ACCENT,
            hover_color=ACCENT_DARK,
            corner_radius=12,
            command=self.on_confirm,
            state="disabled"
        )
        self.button.grid(row=0, column=1, sticky="e", padx=(6, 6), pady=(6, 6))
        ToolTip(self.button, "Confirmar (Enter)")

        # Mensagens
        self.message = ctk.CTkLabel(
            self._bg_top, text="", text_color=FG_MUTED, justify="left"
        )
        self.message.pack(fill="x", padx=24, pady=(0, 10))

        # Rodapé
        footer = ctk.CTkLabel(
            self._bg_top,
            text="Dica: tokens sensíveis não devem ser logados em produção.",
            text_color=FG_MUTED,
            font=ctk.CTkFont(size=10)
        )
        footer.pack(side="bottom", pady=8)

    # ===== Handlers =====
    def _on_entry_change(self, event=None):
        value = self.entry.get().strip()
        self.button.configure(state=("normal" if value else "disabled"))
        self._update_strength(value)
        if self.message.cget("text"):
            self._set_message("", FG_MUTED)

    def _on_enter(self, event=None):
        if self.button.cget("state") == "normal":
            self.on_confirm()

    def _paste_clipboard(self, event=None):
        try:
            clip = self.clipboard_get()
            # substitui o conteúdo atual pelo clipboard
            self.entry.delete(0, ctk.END)
            self.entry.insert(ctk.END, clip)
            self._on_entry_change()
        except Exception:
            pass

    def _copy_token(self, event=None):
        try:
            token = self.entry.get()
            if not token:
                self._set_message("⚠️ Nada para copiar.", ERROR)
                return
            self.clipboard_clear()
            self.clipboard_append(token)
            self._set_message("✅ Token copiado para a área de transferência.", SUCCESS)
        except Exception:
            self._set_message("⚠️ Não foi possível copiar.", ERROR)

    def _clear_entry(self, event=None):
        self.entry.delete(0, ctk.END)
        self._on_entry_change()
        self._set_message("Campo limpo.", FG_MUTED)

    def _toggle_visibility(self):
        self._hidden = not self._hidden
        self.entry.configure(show=("•" if self._hidden else ""))
        self.toggle_btn.configure(text=("👁 Mostrar" if self._hidden else "🙈 Ocultar"))

    def _update_strength(self, value: str):
        # Indicador simples por tamanho (ex.: <8 fraco, 8–24 médio, >24 forte)
        n = len(value)
        if n == 0:
            self.strength_bar.set(0)
            self.strength_label.configure(text="")
            return
        if n < 8:
            self.strength_bar.set(0.25)
            self.strength_bar.configure(progress_color=ERROR)
            self.strength_label.configure(text=f"Tamanho: fraco ({n} chars)", text_color=ERROR)
        elif n < 24:
            self.strength_bar.set(0.6)
            self.strength_bar.configure(progress_color="#F2C94C")  # amarelo
            self.strength_label.configure(text=f"Tamanho: médio ({n} chars)", text_color="#F2C94C")
        else:
            self.strength_bar.set(1.0)
            self.strength_bar.configure(progress_color=SUCCESS)
            self.strength_label.configure(text=f"Tamanho: forte ({n} chars)", text_color=SUCCESS)

    def on_confirm(self):
        value = self.entry.get().strip()
        if not value:
            self._set_message("⚠️ Entrada vazia — insira um valor.", ERROR)
            return

        if len(value) < 8:
            self._set_message("⚠️ Token muito curto (mínimo de 8 caracteres).", ERROR)
            return

        # Feedback de sucesso
        self._set_message(f"✅ Valor recebido ({len(value)} caracteres).", SUCCESS)
        # Em produção, evite logar o token completo:
        print("Valor do input (uso interno):", value)

        # Se desejar limpar:
        self._clear_entry()

        # Ou processe aqui:
        self._process_token(value)

    def renovar_token(self,refresh_token):
        headers = {
            "Authorization": f"Bearer {refresh_token}",
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json"
        }
        payload = {"refreshToken": refresh_token}
        try:
            response = requests.post(REFRESH_URL, headers=headers, json=payload)
            if response.status_code != 200:
                print(f"❌ Falha ao renovar token: {response.status_code} {response.text}")
                return None
            return json.loads(response.json()["jsonToken"])
        except Exception as e:
            print(f"❌ Erro na requisição de refresh: {e}")
            return None


    def _process_token(self, token: str):
        novo_token = self.renovar_token(token)
        if novo_token:
            expires_seconds = novo_token["expires_in"]
            expiration_date = datetime.now() + timedelta(seconds=expires_seconds)
            resultado = {
                "access_token": novo_token["access_token"],
                "refresh_token": novo_token["refresh_token"],
                "expiration_date": expiration_date.isoformat(),
                "expires_in_seconds": expires_seconds
            }
            print("✅ Token renovado com sucesso:", resultado)
            os.makedirs(TOKEN_DIR, exist_ok=True)
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                json.dump(resultado, f, indent=4, ensure_ascii=False)
            self._set_message("✔ Token renovado e salvo com sucesso!", SUCCESS)
        else:
            self._set_message("❌ Falha ao renovar o token.", ERROR)
        

    def _set_message(self, text: str, color: str = FG_MUTED):
        self.message.configure(text=text, text_color=color)

    def _center_window(self):
        self.update_idletasks()
        w, h = 560, 360
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 3
        self.geometry(f"{w}x{h}+{x}+{y}")


if __name__ == "__main__":
    app = TokenApp()
    app.mainloop()
