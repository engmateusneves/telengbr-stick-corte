# -*- coding: utf-8 -*-
# tela principal - tkinter
# v0.11 beta - ainda to arrumando o layout

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .parser import read_input, validate_input, infer_weight_per_meter, find_best_sheet
from openpyxl import load_workbook
from .models import Parameters
from .optimizer import build_plan
from .output import generate_output

# perfis que sempre aparecem, deixa aqui pra facilitar
DEFAULT_PROFILES = [
    "G90-1.25",
    "89S-41-0.8",
    "140S-1.2",
    "nodus-rfy-metric-067",
    "nodus-rfy-metric-116",
]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Tel.eng.br - Stick Corte v0.11 beta")
        self.root.geometry("820x620")
        # tenta colocar icone, se não achar ignora
        try:
            if os.path.exists("icon_profiles.ico"):
                self.root.iconbitmap("icon_profiles.ico")
        except:
            pass

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.kerf = tk.StringVar(value="2")
        self.scrap = tk.StringVar(value="100")
        self.mode = tk.StringVar(value="Global")
        self.price = tk.StringVar(value="25,55")
        self.exact_limit = tk.StringVar(value="22")
        self.profile_vars = {}
        self._build()

    def _build(self):
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Tel.eng.br - Stick Corte v0.11 beta", font=("Arial", 13, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frame, text="TEL.ENG.BR | Mateus Neves (47) 9 8836-1017", font=("Arial", 9)).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,8))

        ttk.Label(frame, text="Excel do Nodus (.xlsx):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.input_path, width=70).grid(row=3, column=0, sticky="ew")
        ttk.Button(frame, text="Selecionar...", command=self.select_input).grid(row=3, column=1, padx=6)

        ttk.Label(frame, text="Saída:").grid(row=4, column=0, sticky="w", pady=(10,0))
        ttk.Entry(frame, textvariable=self.output_path, width=70).grid(row=5, column=0, sticky="ew")
        ttk.Button(frame, text="Salvar como...", command=self.select_output).grid(row=5, column=1, padx=6)

        params = ttk.LabelFrame(frame, text="Parâmetros", padding=8)
        params.grid(row=6, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Label(params, text="Kerf:").grid(row=0, column=0, sticky="w")
        ttk.Entry(params, textvariable=self.kerf, width=8).grid(row=0, column=1, sticky="w")
        ttk.Label(params, text="Sobra min:").grid(row=0, column=2, sticky="w", padx=(12,0))
        ttk.Entry(params, textvariable=self.scrap, width=8).grid(row=0, column=3, sticky="w")
        ttk.Label(params, text="R$/kg:").grid(row=0, column=4, sticky="w", padx=(12,0))
        ttk.Entry(params, textvariable=self.price, width=8).grid(row=0, column=5, sticky="w")
        ttk.Label(params, text="Modo:").grid(row=0, column=6, sticky="w", padx=(12,0))
        ttk.Combobox(params, textvariable=self.mode, values=["Por painel", "Global"], state="readonly", width=12).grid(row=0, column=7, sticky="w")

        self.profile_frame = ttk.LabelFrame(frame, text="Perfis (detectados automaticamente)", padding=6)
        self.profile_frame.grid(row=7, column=0, columnspan=2, sticky="ew")
        ttk.Label(self.profile_frame, text="Selecione o arquivo e clique em Carregar").pack(anchor="w")

        btns = ttk.Frame(frame)
        btns.grid(row=8, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Button(btns, text="Carregar perfis", command=self.load_profiles).pack(side="left")
        ttk.Button(btns, text="GERAR", command=self.generate).pack(side="right")

        self.status = tk.StringVar(value="Aguardando arquivo - v0.11")
        ttk.Label(frame, textvariable=self.status, wraplength=750, font=("Arial", 8)).grid(row=9, column=0, columnspan=2, sticky="w")

    def select_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                self.output_path.set(os.path.splitext(path)[0] + "_StickCorte_v011.xlsx")

    def select_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def load_profiles(self):
        if not self.input_path.get():
            messagebox.showwarning("Atenção", "Selecione o Excel primeiro.")
            return
        try:
            wb = load_workbook(self.input_path.get(), data_only=True, read_only=True)
            sheet = find_best_sheet(wb)
            pieces, _, _ = read_input(self.input_path.get(), sheet_name=sheet)
            profiles = sorted(set(p.profile for p in pieces) | set(DEFAULT_PROFILES))
            for child in self.profile_frame.winfo_children():
                child.destroy()
            for p in profiles:
                var = tk.BooleanVar(value=True)
                self.profile_vars[p] = var
                ttk.Checkbutton(self.profile_frame, text=p, variable=var).pack(anchor="w")
            self.status.set(f"{len(pieces)} peças da aba '{sheet}' | {len(profiles)} perfis")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _params(self, pieces):
        inferred = infer_weight_per_meter(pieces)
        profiles = sorted(set(p.profile for p in pieces) | set(DEFAULT_PROFILES))
        bar_lengths = {p: [3000, 6000] for p in profiles}
        optimize = {p: self.profile_vars.get(p, tk.BooleanVar(value=True)).get() for p in profiles}
        weights = {p: inferred.get(p) for p in profiles}
        return Parameters(
            bar_lengths=bar_lengths,
            kerf_mm=int(float(self.kerf.get().replace(",", "."))),
            useful_scrap_min_mm=int(float(self.scrap.get().replace(",", "."))),
            optimization_mode=self.mode.get(),
            steel_price_per_kg=float(self.price.get().replace(".", "").replace(",", ".")),
            weight_per_meter=weights,
            optimize_profile=optimize,
            exact_piece_limit=int(self.exact_limit.get()),
        )

    def generate(self):
        if not self.input_path.get():
            messagebox.showwarning("Atenção", "Selecione o Excel.")
            return
        output = self.output_path.get() or os.path.splitext(self.input_path.get())[0] + "_StickCorte_v011.xlsx"
        try:
            self.status.set("Processando... pode demorar 1 min se tiver muita peça")
            self.root.update_idletasks()
            wb = load_workbook(self.input_path.get(), data_only=True, read_only=True)
            sheet = find_best_sheet(wb)
            pieces, subtotals, line_alerts = read_input(self.input_path.get(), sheet_name=sheet)
            alerts = validate_input(pieces, subtotals, line_alerts)
            params = self._params(pieces)
            plan = build_plan(pieces, params, alerts, subtotals)
            generate_output(plan, params, output)
            self.status.set(f"Pronto: {len(plan.baseline_bars)} original vs {len(plan.bars)} otimizado | {output}")
            messagebox.showinfo("Pronto", f"ORIGINAL: {len(plan.baseline_bars)} barras\nOTIMIZADO: {len(plan.bars)} barras\n\nArquivo:\n{output}")
        except Exception as exc:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro", str(exc))

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
