#!/usr/bin/env python3
"""Tela desktop do Robo Sallus.

Abre uma interface simples com:
- botao Etapa 1: baixar fila Salus
- botao Etapa 2: processar/lancar evolucao
- cards com contagens da fila e da planilha clinica
- paciente/senha atual
- log de execucao
"""

from __future__ import annotations

import datetime as dt
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from etapa2_lancar_evolucao_salus import (
    SUCCESS_STATUSES,
    PatientResult,
    QueuePatient,
    process_patients,
    read_clinical,
    read_queue,
    read_successful_passwords,
    write_report,
)


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "exports"


def newest(pattern: str, fallback: str) -> Path:
    matches = sorted(EXPORTS.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else EXPORTS / fallback


def today_report() -> Path:
    return EXPORTS / f"relatorio_lancamentos_{dt.date.today().isoformat()}.xlsx"


class RoboSallusApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Robo Sallus")
        self.geometry("1180x760")
        self.minsize(980, 640)

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = False

        self.fila_var = tk.StringVar(value=str(newest("pacientes_sirio_libanes_*.xlsx", "pacientes_sirio_libanes_DATA.xlsx")))
        self.clinica_var = tk.StringVar(value=str(newest("preenchimento_evolucao_clinica_*_colorido.xlsx", "preenchimento_evolucao_clinica_DATA.xlsx")))
        self.relatorio_var = tk.StringVar(value=str(today_report()))
        self.real_run_var = tk.BooleanVar(value=False)
        self.only_password_var = tk.StringVar(value="")

        self.total_salus_var = tk.StringVar(value="-")
        self.total_excel_var = tk.StringVar(value="-")
        self.total_encontrados_var = tk.StringVar(value="-")
        self.total_faltam_var = tk.StringVar(value="-")
        self.processados_var = tk.StringVar(value="0")
        self.current_patient_var = tk.StringVar(value="Nenhum paciente em execucao")
        self.current_password_var = tk.StringVar(value="-")
        self.status_var = tk.StringVar(value="Pronto")

        self._build_style()
        self._build_ui()
        self.after(150, self._poll_events)
        self.refresh_counts()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.configure("TFrame", background="#f4f6f8")
        style.configure("Root.TFrame", background="#f4f6f8")
        style.configure("Header.TLabel", background="#f4f6f8", foreground="#1d2733", font=("Helvetica", 22, "bold"))
        style.configure("Sub.TLabel", background="#f4f6f8", foreground="#5c6670", font=("Helvetica", 12))
        style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("CardTitle.TLabel", background="#ffffff", foreground="#6b7280", font=("Helvetica", 11))
        style.configure("CardValue.TLabel", background="#ffffff", foreground="#1f4e78", font=("Helvetica", 24, "bold"))
        style.configure("Current.TFrame", background="#eaf2f8", relief="solid", borderwidth=1)
        style.configure("CurrentTitle.TLabel", background="#eaf2f8", foreground="#506070", font=("Helvetica", 12))
        style.configure("CurrentValue.TLabel", background="#eaf2f8", foreground="#17324d", font=("Helvetica", 18, "bold"))
        style.configure("Primary.TButton", font=("Helvetica", 12, "bold"), padding=(14, 10))
        style.configure("Secondary.TButton", font=("Helvetica", 11), padding=(10, 8))
        style.configure("Danger.TCheckbutton", background="#f4f6f8", foreground="#9b1c1c", font=("Helvetica", 11, "bold"))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=18)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root, style="Root.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="Robo Sallus", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Etapa 1 baixa a fila. Etapa 2 confere senha por senha e prepara/lanca a evolucao.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        files = ttk.Frame(root, style="Root.TFrame")
        files.pack(fill="x", pady=(0, 12))
        self._file_row(files, 0, "Fila Salus", self.fila_var, [("Excel", "*.xlsx")])
        self._file_row(files, 1, "Planilha Clinica", self.clinica_var, [("Excel", "*.xlsx")])
        self._file_row(files, 2, "Relatorio", self.relatorio_var, [("Excel", "*.xlsx")], save=True)

        controls = ttk.Frame(root, style="Root.TFrame")
        controls.pack(fill="x", pady=(0, 14))
        self.btn_etapa1 = ttk.Button(controls, text="Etapa 1: Baixar Fila Salus", style="Primary.TButton", command=self.run_etapa1)
        self.btn_etapa1.pack(side="left", padx=(0, 10))
        self.btn_etapa2 = ttk.Button(controls, text="Etapa 2: Lançar Evolução", style="Primary.TButton", command=self.run_etapa2)
        self.btn_etapa2.pack(side="left", padx=(0, 10))
        ttk.Button(controls, text="Atualizar Cards", style="Secondary.TButton", command=self.refresh_counts).pack(side="left", padx=(0, 10))
        ttk.Label(controls, text="Somente senha:", background="#f4f6f8").pack(side="left", padx=(16, 5))
        ttk.Entry(controls, textvariable=self.only_password_var, width=14).pack(side="left")
        ttk.Checkbutton(
            controls,
            text="Executar no Salus real",
            variable=self.real_run_var,
            style="Danger.TCheckbutton",
        ).pack(side="left", padx=(18, 0))

        cards = ttk.Frame(root, style="Root.TFrame")
        cards.pack(fill="x", pady=(0, 14))
        for idx, (title, var) in enumerate(
            [
                ("Pacientes no Salus", self.total_salus_var),
                ("Pacientes no Excel", self.total_excel_var),
                ("Senhas encontradas", self.total_encontrados_var),
                ("Faltam lançar", self.total_faltam_var),
                ("Processados agora", self.processados_var),
            ]
        ):
            card = self._card(cards, title, var)
            card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 10, 0))
            cards.columnconfigure(idx, weight=1)

        current = ttk.Frame(root, style="Current.TFrame", padding=14)
        current.pack(fill="x", pady=(0, 14))
        left = ttk.Frame(current, style="Current.TFrame")
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text="Paciente atual", style="CurrentTitle.TLabel").pack(anchor="w")
        ttk.Label(left, textvariable=self.current_patient_var, style="CurrentValue.TLabel").pack(anchor="w", pady=(4, 0))
        right = ttk.Frame(current, style="Current.TFrame")
        right.pack(side="right")
        ttk.Label(right, text="Senha", style="CurrentTitle.TLabel").pack(anchor="e")
        ttk.Label(right, textvariable=self.current_password_var, style="CurrentValue.TLabel").pack(anchor="e", pady=(4, 0))

        log_frame = ttk.Frame(root, style="Root.TFrame")
        log_frame.pack(fill="both", expand=True)
        ttk.Label(log_frame, textvariable=self.status_var, background="#f4f6f8", foreground="#4b5563").pack(anchor="w", pady=(0, 6))
        self.log_text = tk.Text(log_frame, height=16, wrap="word", font=("Menlo", 12), bg="#111827", fg="#e5e7eb", insertbackground="#e5e7eb")
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _file_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar, filetypes, save: bool = False) -> None:
        ttk.Label(parent, text=label, background="#f4f6f8", width=18).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4, padx=(0, 8))
        command = lambda: self._choose_file(var, filetypes, save)
        ttk.Button(parent, text="Escolher", command=command).grid(row=row, column=2, sticky="e", pady=4)
        parent.columnconfigure(1, weight=1)

    def _card(self, parent: ttk.Frame, title: str, value: tk.StringVar) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=14)
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(frame, textvariable=value, style="CardValue.TLabel").pack(anchor="w", pady=(8, 0))
        return frame

    def _choose_file(self, var: tk.StringVar, filetypes, save: bool) -> None:
        initial = Path(var.get()).parent if var.get() else EXPORTS
        if save:
            chosen = filedialog.asksaveasfilename(initialdir=initial, filetypes=filetypes, defaultextension=".xlsx")
        else:
            chosen = filedialog.askopenfilename(initialdir=initial, filetypes=filetypes)
        if chosen:
            var.set(chosen)
            self.refresh_counts()

    def log(self, message: str) -> None:
        timestamp = dt.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

    def set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        self.btn_etapa1.configure(state=state)
        self.btn_etapa2.configure(state=state)

    def refresh_counts(self) -> None:
        try:
            fila = Path(self.fila_var.get())
            clinica = Path(self.clinica_var.get())
            relatorio = Path(self.relatorio_var.get())

            queue_patients = read_queue(fila) if fila.exists() else []
            clinical_by_password, _, _ = read_clinical(clinica) if clinica.exists() else ({}, {}, [])
            successful = read_successful_passwords(relatorio)

            clinical_unique = {senha for senha, rows in clinical_by_password.items() if len(rows) == 1}
            queue_passwords = [p.senha for p in queue_patients]
            found = sum(1 for senha in queue_passwords if senha in clinical_unique)
            missing_to_launch = sum(1 for senha in queue_passwords if senha not in successful)

            self.total_salus_var.set(str(len(queue_patients)))
            self.total_excel_var.set(str(sum(len(rows) for rows in clinical_by_password.values())))
            self.total_encontrados_var.set(str(found))
            self.total_faltam_var.set(str(missing_to_launch))
            self.status_var.set("Cards atualizados")
        except Exception as exc:
            self.status_var.set("Erro ao atualizar cards")
            self.log(f"Erro ao atualizar cards: {exc}")

    def run_etapa1(self) -> None:
        if self.running:
            return
        self.set_running(True)
        self.status_var.set("Executando Etapa 1...")
        self.log("Etapa 1 iniciada: baixar fila Salus.")
        thread = threading.Thread(target=self._run_etapa1_worker, daemon=True)
        thread.start()

    def _run_etapa1_worker(self) -> None:
        try:
            output = EXPORTS / f"pacientes_sirio_libanes_{dt.date.today().isoformat()}.xlsx"
            cmd = [sys.executable, str(ROOT / "scripts" / "gerar_lista_pacientes.py"), "--saida", str(output)]
            completed = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, check=False)
            self.events.put(("log", completed.stdout.strip() or "Etapa 1 concluida."))
            if completed.stderr.strip():
                self.events.put(("log", completed.stderr.strip()))
            if completed.returncode != 0:
                self.events.put(("error", f"Etapa 1 falhou com codigo {completed.returncode}."))
            else:
                self.events.put(("fila", str(output)))
                self.events.put(("done", "Etapa 1 concluida."))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def run_etapa2(self) -> None:
        if self.running:
            return
        if self.real_run_var.get():
            ok = messagebox.askyesno(
                "Confirmar execução real",
                "Voce marcou execucao real no Salus. Confirma que deseja continuar?",
            )
            if not ok:
                return
        self.set_running(True)
        self.processados_var.set("0")
        self.status_var.set("Executando Etapa 2...")
        mode = "REAL" if self.real_run_var.get() else "DRY-RUN"
        self.log(f"Etapa 2 iniciada em modo {mode}.")
        thread = threading.Thread(target=self._run_etapa2_worker, daemon=True)
        thread.start()

    def _run_etapa2_worker(self) -> None:
        try:
            fila = Path(self.fila_var.get())
            clinica = Path(self.clinica_var.get())
            relatorio = Path(self.relatorio_var.get())
            only_password = self.only_password_var.get().strip() or None

            queue_patients = read_queue(fila)
            clinical_by_password, field_meta, field_headers = read_clinical(clinica)
            successful_passwords = read_successful_passwords(relatorio)
            processed_count = 0

            def progress(event: str, patient: QueuePatient, result: PatientResult | None) -> None:
                nonlocal processed_count
                if event == "inicio":
                    self.events.put(("current", (patient.nome or "-", patient.senha)))
                elif event == "fim" and result:
                    processed_count += 1
                    self.events.put(("processed", processed_count))
                    self.events.put(("log", f"{result.senha} - {result.nome} - {result.status}: {result.mensagem}"))

            results = process_patients(
                queue_patients=queue_patients,
                clinical_by_password=clinical_by_password,
                field_meta=field_meta,
                field_headers=field_headers,
                successful_passwords=successful_passwords,
                dry_run=not self.real_run_var.get(),
                only_password=only_password,
                progress_callback=progress,
            )
            write_report(results, relatorio)
            self.events.put(("done", f"Etapa 2 concluida. Relatorio: {relatorio}"))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "log":
                    self.log(str(payload))
                elif event == "error":
                    self.log(f"ERRO: {payload}")
                    self.status_var.set("Erro")
                    self.set_running(False)
                    self.refresh_counts()
                elif event == "done":
                    self.log(str(payload))
                    self.status_var.set(str(payload))
                    self.set_running(False)
                    self.refresh_counts()
                elif event == "fila":
                    self.fila_var.set(str(payload))
                elif event == "current":
                    nome, senha = payload  # type: ignore[misc]
                    self.current_patient_var.set(str(nome))
                    self.current_password_var.set(str(senha))
                elif event == "processed":
                    self.processados_var.set(str(payload))
        except queue.Empty:
            pass
        self.after(150, self._poll_events)


def main() -> int:
    app = RoboSallusApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
