from __future__ import annotations

import re
from pathlib import Path
from tkinter import StringVar, Tk, filedialog, messagebox
from tkinter import ttk

PDF_BACKEND = "pypdf"
try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None


UPPER_NAME_PATTERN = re.compile(
    r"\b([A-Z]{2,}(?:\s+(?:[A-Z]{2,}|DA|DE|DI|DO|DU|DAS|DOS|E)){1,8})\b"
)
LINE_BREAK_SPACES = re.compile(r"\s+")
TRAILING_FIELDS_PATTERN = re.compile(
    r"\b(CARGO|CTPS|LOCALIZACAO|CATEGORIA|HORARIOS)\s*:",
    flags=re.IGNORECASE,
)
NOISE_WORDS = {
    "COMERCIALIZACAO",
    "BRASIL",
    "MENSALISTA",
    "CARGO",
    "CTPS",
    "LOCALIZACAO",
    "CATEGORIA",
    "TRABALHO",
    "FALTAS",
    "ATRASOS",
}
LOWERCASE_PARTICLES = {"da", "de", "di", "do", "du", "das", "dos", "e"}


def normalize_spaces(value: str) -> str:
    return LINE_BREAK_SPACES.sub(" ", value).strip()


def normalize_name_case(name: str) -> str:
    words = normalize_spaces(name).split(" ")
    normalized: list[str] = []
    for i, word in enumerate(words):
        lower = word.lower()
        if i > 0 and lower in LOWERCASE_PARTICLES:
            normalized.append(lower)
        else:
            normalized.append(lower.capitalize())
    return " ".join(normalized)


def looks_like_name(name: str) -> bool:
    if not name:
        return False
    parts = [p for p in normalize_spaces(name).split(" ") if p]
    if len(parts) < 2:
        return False
    if any(any(ch.isdigit() for ch in part) for part in parts):
        return False
    return True


def extract_text_from_pdf(pdf_path: Path) -> str:
    if PdfReader is None:
        raise RuntimeError("PDF library not found. Run: pip install pypdf")

    reader = PdfReader(str(pdf_path))
    chunks: list[str] = []
    for page in reader.pages:
        text = page.extract_text(extraction_mode="layout") or ""
        if not text:
            text = page.extract_text() or ""
        if text:
            chunks.append(text)
    return "\n".join(chunks)


def _pick_upper_name_candidate(text: str) -> str | None:
    candidates = UPPER_NAME_PATTERN.findall(text.upper())
    for candidate in candidates:
        clean = normalize_spaces(candidate)
        tokens = clean.split(" ")
        if len(tokens) < 2:
            continue
        if len(tokens) > 8:
            continue
        if any(token in NOISE_WORDS for token in tokens):
            continue
        return normalize_name_case(clean)
    return None


def extract_name_from_employee_field(text: str) -> str | None:
    lines = [normalize_spaces(line) for line in text.splitlines() if line.strip()]

    for idx, line in enumerate(lines):
        lower_line = line.lower()
        if "empregado:" not in lower_line or "empregador:" in lower_line:
            continue

        after_label = line.split(":", 1)[1].strip() if ":" in line else line
        after_label = re.sub(r"^\d+\s*", "", after_label)
        after_label = TRAILING_FIELDS_PATTERN.split(after_label)[0].strip()

        if looks_like_name(after_label):
            return normalize_name_case(after_label)

        upper_guess = _pick_upper_name_candidate(after_label)
        if upper_guess:
            return upper_guess

        if idx + 1 < len(lines):
            next_line = lines[idx + 1]
            next_line_clean = TRAILING_FIELDS_PATTERN.split(next_line)[0].strip()
            if looks_like_name(next_line_clean):
                return normalize_name_case(next_line_clean)
            upper_guess = _pick_upper_name_candidate(next_line_clean)
            if upper_guess:
                return upper_guess

    linear = normalize_spaces(text.upper())
    block_match = re.search(
        r"EMPREGADO:\s*\d*\s*(.+?)\s*(?:CARGO:|CTPS:|LOCALIZACAO:|CATEGORIA:|HORARIOS:|DT\s*SEM|MARCACOES|$)",
        linear,
    )
    if block_match:
        return _pick_upper_name_candidate(block_match.group(1))

    return None


def extract_name_from_filename(pdf_path: Path) -> str | None:
    stem = pdf_path.stem.replace("_", " ").replace("-", " ").strip()
    if looks_like_name(stem):
        return normalize_name_case(stem)
    return None


class PdfNameReaderApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Leitor de Nomes em PDFs")
        self.root.geometry("1120x560")

        self.folder_var = StringVar()
        self.status_var = StringVar(value="Selecione uma pasta com PDFs.")
        self.results: list[tuple[str, str, str]] = []

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        top = ttk.Frame(main)
        top.pack(fill="x")

        ttk.Label(top, text="Pasta de PDFs:").pack(side="left")

        folder_entry = ttk.Entry(top, textvariable=self.folder_var)
        folder_entry.pack(side="left", fill="x", expand=True, padx=8)

        ttk.Button(top, text="Selecionar", command=self.choose_folder).pack(side="left")
        ttk.Button(top, text="Ler PDFs", command=self.scan_pdfs).pack(side="left", padx=8)
        ttk.Button(top, text="Copiar Grid", command=self.copy_grid_to_clipboard).pack(side="left")

        columns = ("arquivo", "nome", "detalhe")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("arquivo", text="Arquivo PDF")
        self.tree.heading("nome", text="Nome encontrado")
        self.tree.heading("detalhe", text="Detalhe")
        self.tree.column("arquivo", width=300, anchor="w")
        self.tree.column("nome", width=320, anchor="w")
        self.tree.column("detalhe", width=450, anchor="w")

        scrollbar = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True, pady=10)
        scrollbar.pack(side="left", fill="y", pady=10)

        status = ttk.Label(main, textvariable=self.status_var, anchor="w")
        status.pack(fill="x")

        self.root.bind("<Control-c>", self.copy_grid_to_clipboard_event)

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if folder:
            self.folder_var.set(folder)
            self.status_var.set("Pasta selecionada. Clique em 'Ler PDFs'.")

    def scan_pdfs(self) -> None:
        if PdfReader is None:
            messagebox.showerror(
                "Dependencia ausente",
                "Nao encontrei biblioteca para ler PDF.\n\n"
                "No terminal, rode:\n"
                "pip install pypdf",
            )
            return

        folder_text = self.folder_var.get().strip()
        if not folder_text:
            messagebox.showwarning("Aviso", "Selecione uma pasta primeiro.")
            return

        folder = Path(folder_text)
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("Erro", "Pasta invalida.")
            return

        pdf_files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])
        if not pdf_files:
            messagebox.showinfo("Resultado", "Nenhum arquivo PDF encontrado na pasta.")
            return

        self.results.clear()
        self.tree.delete(*self.tree.get_children())

        errors = 0
        for pdf_path in pdf_files:
            try:
                text = extract_text_from_pdf(pdf_path)
                name = extract_name_from_employee_field(text)
                detail = "extraido do campo Empregado"

                if not name:
                    fallback = extract_name_from_filename(pdf_path)
                    if fallback:
                        name = fallback
                        detail = "fallback: nome do arquivo"

                if not name:
                    self.results.append((pdf_path.name, "(nenhum nome identificado)", ""))
                else:
                    self.results.append((pdf_path.name, name, detail))
            except Exception as exc:
                errors += 1
                detail = f"{type(exc).__name__}: {exc}".strip()
                self.results.append((pdf_path.name, "(erro ao ler arquivo)", detail[:420]))

        for row in self.results:
            self.tree.insert("", "end", values=row)

        backend_text = PDF_BACKEND or "desconhecido"
        self.status_var.set(
            f"Concluido: {len(pdf_files)} PDF(s), {len(self.results)} linha(s), "
            f"{errors} erro(s). Motor: {backend_text}"
        )

    def copy_grid_to_clipboard_event(self, _event) -> None:
        self.copy_grid_to_clipboard()

    def copy_grid_to_clipboard(self) -> None:
        rows = self.tree.selection()
        if rows:
            selected = [self.tree.item(iid, "values") for iid in rows]
        else:
            selected = [self.tree.item(iid, "values") for iid in self.tree.get_children()]

        if not selected:
            messagebox.showwarning("Aviso", "Nao ha dados para copiar.")
            return

        header = ["Arquivo PDF", "Nome encontrado", "Detalhe"]
        lines = ["\t".join(header)]
        lines.extend("\t".join(str(value) for value in row) for row in selected)
        tsv = "\n".join(lines)

        self.root.clipboard_clear()
        self.root.clipboard_append(tsv)
        self.status_var.set(
            f"Grid copiado para area de transferencia ({len(selected)} linha(s)). "
            f"Cole no Excel com Ctrl+V."
        )


def main() -> None:
    root = Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    PdfNameReaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
