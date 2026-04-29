from __future__ import annotations

from pathlib import Path
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback

from .document import create_template, generate_ebook
from .spreadsheet import read_companies


class EbookFiscalApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Gerador de E-book Fiscal")
        self.geometry("760x430")
        self.minsize(680, 390)

        self.planilha_var = tk.StringVar()
        self.saida_var = tk.StringVar(value=str(Path.cwd() / "ebooks_gerados"))
        self.aba_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Selecione uma planilha Excel para começar.")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(container, text="Gerador de E-book Fiscal", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 14))

        self._add_file_row(
            container,
            row=1,
            label="Planilha Excel",
            variable=self.planilha_var,
            command=self._choose_spreadsheet,
        )
        self._add_file_row(
            container,
            row=2,
            label="Pasta de saída",
            variable=self.saida_var,
            command=self._choose_output_dir,
        )
        self._add_file_row(
            container,
            row=3,
            label="Template Word",
            variable=self.template_var,
            command=self._choose_template,
            optional=True,
        )

        ttk.Label(container, text="Aba da planilha").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Entry(container, textvariable=self.aba_var).grid(row=4, column=1, sticky="ew", pady=8)
        ttk.Label(container, text="Opcional. Em branco usa a primeira aba.").grid(
            row=4, column=2, sticky="w", padx=(8, 0), pady=8
        )

        buttons = ttk.Frame(container)
        buttons.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(16, 8))
        buttons.columnconfigure(0, weight=1)

        ttk.Button(buttons, text="Criar planilha/template de exemplo", command=self._create_examples).grid(
            row=0, column=0, sticky="w"
        )
        self.generate_button = ttk.Button(buttons, text="Gerar e-books", command=self._generate)
        self.generate_button.grid(row=0, column=1, sticky="e")

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 8))

        status = ttk.Label(container, textvariable=self.status_var, wraplength=700, foreground="#444444")
        status.grid(row=7, column=0, columnspan=3, sticky="w")

        help_text = (
            "Colunas esperadas: Número da empresa, Razão social, CNPJ, UF, Regime tributário, "
            "Tipo Comercial, Seguimento/atividade, Responsável, E-mail, Telefone, Cargo, "
            "Obrigações fiscais, Processos, Particulariedades e Pontos de atenção."
        )
        ttk.Label(container, text=help_text, wraplength=700).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(18, 0)
        )

    def _add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
        optional: bool = False,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=8)
        button_text = "Selecionar"
        ttk.Button(parent, text=button_text, command=command).grid(row=row, column=2, sticky="e", padx=(8, 0))
        if optional:
            variable.set("")

    def _choose_spreadsheet(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.planilha_var.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Selecione a pasta de saída")
        if path:
            self.saida_var.set(path)

    def _choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecione o template Word",
            filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.template_var.set(path)

    def _create_examples(self) -> None:
        output_dir = Path(self.saida_var.get() or Path.cwd())
        output_dir.mkdir(parents=True, exist_ok=True)
        template_path = output_dir / "Ebook_Fiscal_Template.docx"
        spreadsheet_path = output_dir / "empresas_fiscais_exemplo.xlsx"

        create_template(template_path)
        from .cli import example_command
        import argparse

        example_command(argparse.Namespace(saida=str(spreadsheet_path)))
        self.status_var.set(f"Exemplos criados em: {output_dir}")
        messagebox.showinfo("Exemplos criados", f"Arquivos criados em:\n{output_dir}")

    def _generate(self) -> None:
        spreadsheet = self.planilha_var.get().strip()
        output_dir = self.saida_var.get().strip()
        if not spreadsheet:
            messagebox.showwarning("Planilha obrigatória", "Selecione a planilha Excel.")
            return
        if not output_dir:
            messagebox.showwarning("Saída obrigatória", "Selecione a pasta de saída.")
            return

        self._set_running(True)
        Thread(target=self._generate_worker, daemon=True).start()

    def _generate_worker(self) -> None:
        try:
            spreadsheet = self.planilha_var.get().strip()
            output_dir = Path(self.saida_var.get().strip())
            sheet_name = self.aba_var.get().strip() or None
            template_path = self.template_var.get().strip() or None

            companies = read_companies(spreadsheet, sheet_name=sheet_name)
            output_dir.mkdir(parents=True, exist_ok=True)

            generated_paths: list[Path] = []
            for company in companies:
                output_path = output_dir / company.output_filename()
                generate_ebook(company, output_path, template_path=template_path)
                generated_paths.append(output_path)

            self.after(0, self._show_success, generated_paths, output_dir)
        except Exception as exc:  # noqa: BLE001 - desktop app must surface unexpected errors.
            details = traceback.format_exc()
            self.after(0, self._show_error, exc, details)

    def _set_running(self, running: bool) -> None:
        if running:
            self.status_var.set("Gerando e-books, aguarde...")
            self.generate_button.state(["disabled"])
            self.progress.start(10)
        else:
            self.generate_button.state(["!disabled"])
            self.progress.stop()

    def _show_success(self, generated_paths: list[Path], output_dir: Path) -> None:
        self._set_running(False)
        self.status_var.set(f"Concluído. {len(generated_paths)} e-book(s) gerado(s) em {output_dir}.")
        messagebox.showinfo(
            "Geração concluída",
            f"{len(generated_paths)} e-book(s) gerado(s).\n\nPasta:\n{output_dir}",
        )

    def _show_error(self, exc: Exception, details: str) -> None:
        self._set_running(False)
        self.status_var.set(f"Erro: {exc}")
        messagebox.showerror("Erro ao gerar e-books", f"{exc}\n\nDetalhes técnicos:\n{details}")


def main() -> None:
    app = EbookFiscalApp()
    app.mainloop()


if __name__ == "__main__":
    main()
