from __future__ import annotations

import calendar
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from database import Database
from document_editor import export_docx, export_pdf
from printing import calculate_physical_sheets, detect_page_count, send_to_printer


MONTHS = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]


class CopyManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Fotocopias - Albergue Universitario")
        self.geometry("1200x760")
        self.minsize(980, 650)
        self.db = Database()
        self.selected_student_id: int | None = None
        self.student_rows: dict[str, int] = {}

        self._configure_style()
        self._build_ui()
        self.refresh_students()
        self.refresh_history()

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Metric.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Muted.TLabel", foreground="#555555")

    def _build_ui(self):
        header = ttk.Frame(self, padding=(16, 12))
        header.pack(fill="x")
        ttk.Label(header, text="Sistema de Fotocopias", style="Title.TLabel").pack(side="left")

        now = datetime.now()
        period_frame = ttk.Frame(header)
        period_frame.pack(side="right")
        ttk.Label(period_frame, text="Periodo:").pack(side="left", padx=(0, 6))
        self.month_var = tk.StringVar(value=MONTHS[now.month - 1])
        self.year_var = tk.IntVar(value=now.year)
        month_box = ttk.Combobox(
            period_frame, textvariable=self.month_var, values=MONTHS,
            width=13, state="readonly"
        )
        month_box.pack(side="left", padx=(0, 5))
        year_box = ttk.Spinbox(
            period_frame, from_=2020, to=2100, textvariable=self.year_var, width=7
        )
        year_box.pack(side="left")
        month_box.bind("<<ComboboxSelected>>", lambda _: self.on_period_changed())
        year_box.bind("<FocusOut>", lambda _: self.on_period_changed())
        year_box.bind("<Return>", lambda _: self.on_period_changed())

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.management_tab = ttk.Frame(self.notebook, padding=10)
        self.history_tab = ttk.Frame(self.notebook, padding=10)
        self.editor_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.management_tab, text="Becados y saldo")
        self.notebook.add(self.history_tab, text="Historial")
        self.notebook.add(self.editor_tab, text="Redactar / preparar documento")

        self._build_management_tab()
        self._build_history_tab()
        self._build_editor_tab()

    def _build_management_tab(self):
        paned = ttk.Panedwindow(self.management_tab, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, padding=(0, 0, 10, 0))
        right = ttk.Frame(paned, padding=(10, 0, 0, 0))
        paned.add(left, weight=2)
        paned.add(right, weight=3)

        search_frame = ttk.Frame(left)
        search_frame.pack(fill="x", pady=(0, 8))
        self.search_var = tk.StringVar()
        search = ttk.Entry(search_frame, textvariable=self.search_var)
        search.pack(side="left", fill="x", expand=True)
        search.bind("<KeyRelease>", lambda _: self.refresh_students())
        ttk.Button(search_frame, text="Nuevo becado", command=self.add_student).pack(side="left", padx=(8, 0))

        columns = ("apellido", "nombre", "identificador")
        self.students_tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse")
        self.students_tree.heading("apellido", text="Apellido")
        self.students_tree.heading("nombre", text="Nombre")
        self.students_tree.heading("identificador", text="DNI / Legajo")
        self.students_tree.column("apellido", width=155)
        self.students_tree.column("nombre", width=155)
        self.students_tree.column("identificador", width=105)
        self.students_tree.pack(fill="both", expand=True)
        self.students_tree.bind("<<TreeviewSelect>>", self.on_student_selected)

        actions = ttk.Frame(left)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Desactivar becado", command=self.deactivate_student).pack(side="left")
        ttk.Button(actions, text="Actualizar", command=self.refresh_students).pack(side="right")

        self.student_name_label = ttk.Label(right, text="Seleccione un becado", style="Title.TLabel")
        self.student_name_label.pack(anchor="w", pady=(0, 10))

        metrics = ttk.Frame(right)
        metrics.pack(fill="x")
        self.metric_vars = {
            "base": tk.StringVar(value="-"),
            "printed": tk.StringVar(value="-"),
            "received": tk.StringVar(value="-"),
            "donated": tk.StringVar(value="-"),
            "available": tk.StringVar(value="-"),
        }
        labels = [
            ("Cupo base", "base"),
            ("Impresas", "printed"),
            ("Recibidas", "received"),
            ("Donadas", "donated"),
            ("Disponible", "available"),
        ]
        for i, (label, key) in enumerate(labels):
            box = ttk.LabelFrame(metrics, text=label, padding=10)
            box.grid(row=0, column=i, padx=(0, 7), sticky="nsew")
            ttk.Label(box, textvariable=self.metric_vars[key], style="Metric.TLabel").pack()
            metrics.columnconfigure(i, weight=1)

        self.last_print_var = tk.StringVar(value="Ultimo retiro: -")
        ttk.Label(right, textvariable=self.last_print_var, style="Muted.TLabel").pack(anchor="w", pady=(10, 18))

        primary = ttk.LabelFrame(right, text="Operaciones", padding=14)
        primary.pack(fill="x")
        ttk.Button(primary, text="Imprimir archivo", command=self.print_file).pack(fill="x", pady=(0, 8))
        ttk.Button(primary, text="Donar hojas", command=self.donate_sheets).pack(fill="x", pady=(0, 8))
        ttk.Button(primary, text="Corregir ultima impresion", command=self.reverse_last_print).pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=18)
        ttk.Label(
            right,
            text=(
                "El saldo se calcula para el periodo seleccionado. Las donaciones y correcciones "
                "quedan registradas permanentemente en el historial."
            ),
            wraplength=520,
            style="Muted.TLabel",
        ).pack(anchor="w")

    def _build_history_tab(self):
        top = ttk.Frame(self.history_tab)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Movimientos del periodo seleccionado", style="Title.TLabel").pack(side="left")
        ttk.Button(top, text="Actualizar", command=self.refresh_history).pack(side="right")

        cols = ("fecha", "tipo", "becado", "relacionado", "cantidad", "archivo", "nota")
        self.history_tree = ttk.Treeview(self.history_tab, columns=cols, show="headings")
        headings = {
            "fecha": "Fecha", "tipo": "Tipo", "becado": "Becado",
            "relacionado": "Relacionado", "cantidad": "Hojas",
            "archivo": "Archivo", "nota": "Nota",
        }
        widths = {"fecha": 155, "tipo": 130, "becado": 180, "relacionado": 180, "cantidad": 65, "archivo": 180, "nota": 260}
        for col in cols:
            self.history_tree.heading(col, text=headings[col])
            self.history_tree.column(col, width=widths[col], anchor="w")
        self.history_tree.pack(fill="both", expand=True)

    def _build_editor_tab(self):
        toolbar = ttk.Frame(self.editor_tab)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Editor basico", style="Title.TLabel").pack(side="left")
        ttk.Button(toolbar, text="Guardar como Word", command=self.save_editor_docx).pack(side="right", padx=(6, 0))
        ttk.Button(toolbar, text="Guardar como PDF", command=self.save_editor_pdf).pack(side="right", padx=(6, 0))
        ttk.Button(toolbar, text="Limpiar", command=lambda: self.editor.delete("1.0", "end")).pack(side="right")

        ttk.Label(
            self.editor_tab,
            text="Redacte aqui un documento simple. Luego guardelo como PDF o Word y use 'Imprimir archivo' para asociarlo a un becado.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(0, 8))
        self.editor = tk.Text(self.editor_tab, wrap="word", undo=True, font=("Segoe UI", 11), padx=12, pady=12)
        self.editor.pack(fill="both", expand=True)

    def current_period(self) -> str:
        month = MONTHS.index(self.month_var.get()) + 1
        return self.db.normalize_period(int(self.year_var.get()), month)

    def on_period_changed(self):
        self.refresh_selected_student()
        self.refresh_history()

    def refresh_students(self):
        selected_id = self.selected_student_id
        for item in self.students_tree.get_children():
            self.students_tree.delete(item)
        self.student_rows.clear()
        for row in self.db.list_students(self.search_var.get()):
            item = self.students_tree.insert(
                "", "end", values=(row["last_name"], row["first_name"], row["identifier"] or "")
            )
            self.student_rows[item] = int(row["id"])
            if selected_id == row["id"]:
                self.students_tree.selection_set(item)
                self.students_tree.focus(item)
        if selected_id:
            self.refresh_selected_student()

    def on_student_selected(self, _event=None):
        selection = self.students_tree.selection()
        if not selection:
            return
        self.selected_student_id = self.student_rows.get(selection[0])
        self.refresh_selected_student()

    def refresh_selected_student(self):
        if not self.selected_student_id:
            return
        student = self.db.get_student(self.selected_student_id)
        if not student:
            return
        balance = self.db.get_balance(self.selected_student_id, self.current_period())
        self.student_name_label.config(text=f"{student['last_name']}, {student['first_name']}")
        self.metric_vars["base"].set(str(balance.base_quota))
        self.metric_vars["printed"].set(str(balance.printed))
        self.metric_vars["received"].set(f"+{balance.received}")
        self.metric_vars["donated"].set(f"-{balance.donated}")
        self.metric_vars["available"].set(str(balance.available))
        self.last_print_var.set(f"Ultimo retiro: {self._format_timestamp(balance.last_print_at)}")

    def add_student(self):
        dialog = StudentDialog(self)
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            self.db.add_student(*dialog.result)
            self.refresh_students()
        except Exception as exc:
            messagebox.showerror("No se pudo crear", str(exc), parent=self)

    def deactivate_student(self):
        if not self.selected_student_id:
            messagebox.showinfo("Becado", "Seleccione un becado primero.", parent=self)
            return
        if not messagebox.askyesno(
            "Desactivar becado",
            "El becado dejara de aparecer en el listado activo. Su historial no se elimina.\n\n¿Continuar?",
            parent=self,
        ):
            return
        self.db.deactivate_student(self.selected_student_id)
        self.selected_student_id = None
        self.student_name_label.config(text="Seleccione un becado")
        for value in self.metric_vars.values():
            value.set("-")
        self.last_print_var.set("Ultimo retiro: -")
        self.refresh_students()

    def donate_sheets(self):
        if not self._require_student():
            return
        donor = self.db.get_student(self.selected_student_id)
        students = [s for s in self.db.list_students() if s["id"] != self.selected_student_id]
        if not students:
            messagebox.showinfo("Donacion", "No hay otro becado activo para recibir hojas.", parent=self)
            return
        dialog = DonationDialog(self, donor, students)
        self.wait_window(dialog)
        if not dialog.result:
            return
        recipient_id, quantity, note = dialog.result
        try:
            self.db.donate(self.selected_student_id, recipient_id, quantity, self.current_period(), note)
            self.refresh_selected_student()
            self.refresh_history()
            messagebox.showinfo("Donacion registrada", f"Se transfirieron {quantity} hojas correctamente.", parent=self)
        except Exception as exc:
            messagebox.showerror("No se pudo donar", str(exc), parent=self)

    def print_file(self):
        if not self._require_student():
            return
        path = filedialog.askopenfilename(
            parent=self,
            title="Seleccionar archivo para imprimir",
            filetypes=[
                ("Documentos", "*.pdf *.doc *.docx *.txt *.rtf"),
                ("PDF", "*.pdf"), ("Word", "*.doc *.docx"), ("Todos", "*.*"),
            ],
        )
        if not path:
            return
        try:
            detected = detect_page_count(path)
        except Exception as exc:
            messagebox.showerror("Archivo", f"No se pudo leer el documento:\n{exc}", parent=self)
            return

        dialog = PrintDialog(self, Path(path).name, detected)
        self.wait_window(dialog)
        if not dialog.result:
            return
        pages, copies, duplex, note = dialog.result
        sheets = calculate_physical_sheets(pages, copies, duplex)
        balance = self.db.get_balance(self.selected_student_id, self.current_period())
        if sheets > balance.available:
            messagebox.showerror(
                "Saldo insuficiente",
                f"La impresion requiere {sheets} hojas y el becado dispone de {balance.available}.",
                parent=self,
            )
            return

        student = self.db.get_student(self.selected_student_id)
        if not messagebox.askyesno(
            "Confirmar impresion",
            f"Becado: {student['last_name']}, {student['first_name']}\n"
            f"Archivo: {Path(path).name}\nPaginas: {pages}\nCopias: {copies}\n"
            f"Modo: {'Doble faz' if duplex else 'Simple faz'}\n\n"
            f"Se descontaran {sheets} hojas.\n¿Enviar a imprimir?",
            parent=self,
        ):
            return

        try:
            send_to_printer(path)
            self.db.register_print(
                self.selected_student_id,
                sheets,
                self.current_period(),
                file_name=Path(path).name,
                note=note,
            )
            self.refresh_selected_student()
            self.refresh_history()
            messagebox.showinfo(
                "Impresion registrada",
                f"Windows acepto la orden. Se descontaron {sheets} hojas.\n\n"
                "Si la impresion salio mal, use 'Corregir ultima impresion'.",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror(
                "No se registro la impresion",
                f"No se desconto saldo porque la orden no pudo completarse:\n\n{exc}",
                parent=self,
            )

    def reverse_last_print(self):
        if not self._require_student():
            return
        movement = self.db.get_last_reversible_print(self.selected_student_id, self.current_period())
        if not movement:
            messagebox.showinfo("Correccion", "No hay una impresion para corregir en este periodo.", parent=self)
            return
        reason = simpledialog.askstring(
            "Corregir impresion",
            f"Se reintegraran {movement['quantity']} hojas de la impresion:\n"
            f"{movement['file_name'] or 'Sin archivo'}\n\nMotivo de la correccion:",
            parent=self,
        )
        if not reason:
            return
        try:
            self.db.reverse_last_print(self.selected_student_id, self.current_period(), reason)
            self.refresh_selected_student()
            self.refresh_history()
            messagebox.showinfo("Correccion registrada", "Las hojas fueron reintegradas y la correccion quedo en el historial.", parent=self)
        except Exception as exc:
            messagebox.showerror("Correccion", str(exc), parent=self)

    def refresh_history(self):
        if not hasattr(self, "history_tree"):
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        type_names = {"PRINT": "Impresion", "DONATION": "Donacion", "PRINT_REVERSAL": "Correccion"}
        for row in self.db.list_movements(self.current_period()):
            self.history_tree.insert(
                "", "end",
                values=(
                    self._format_timestamp(row["created_at"]),
                    type_names.get(row["movement_type"], row["movement_type"]),
                    row["student_name"] or "",
                    row["related_student_name"] or "",
                    row["quantity"],
                    row["file_name"] or "",
                    row["note"] or "",
                ),
            )

    def save_editor_docx(self):
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".docx", filetypes=[("Word", "*.docx")]
        )
        if path:
            try:
                export_docx(self.editor.get("1.0", "end-1c"), path)
                messagebox.showinfo("Documento guardado", f"Se guardo:\n{path}", parent=self)
            except Exception as exc:
                messagebox.showerror("Guardar", str(exc), parent=self)

    def save_editor_pdf(self):
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
        )
        if path:
            try:
                export_pdf(self.editor.get("1.0", "end-1c"), path)
                messagebox.showinfo("Documento guardado", f"Se guardo:\n{path}", parent=self)
            except Exception as exc:
                messagebox.showerror("Guardar", str(exc), parent=self)

    def _require_student(self) -> bool:
        if not self.selected_student_id:
            messagebox.showinfo("Becado", "Seleccione un becado primero.", parent=self)
            return False
        return True

    @staticmethod
    def _format_timestamp(value: str | None) -> str:
        if not value:
            return "-"
        try:
            dt = datetime.fromisoformat(value)
            return dt.strftime("%d/%m/%Y %H:%M")
        except ValueError:
            return value


class StudentDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Nuevo becado")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        frame = ttk.Frame(self, padding=16)
        frame.pack()
        self.last = self._row(frame, "Apellido", 0)
        self.first = self._row(frame, "Nombre", 1)
        self.identifier = self._row(frame, "DNI / Legajo (opcional)", 2)
        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="left")
        ttk.Button(buttons, text="Guardar", command=self.submit).pack(side="left", padx=(8, 0))
        self.last.focus_set()

    @staticmethod
    def _row(frame, label, row):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 8))
        entry = ttk.Entry(frame, width=32)
        entry.grid(row=row, column=1, pady=4)
        return entry

    def submit(self):
        if not self.last.get().strip() or not self.first.get().strip():
            messagebox.showwarning("Datos", "Apellido y nombre son obligatorios.", parent=self)
            return
        self.result = (self.last.get(), self.first.get(), self.identifier.get())
        self.destroy()


class DonationDialog(tk.Toplevel):
    def __init__(self, parent, donor, students):
        super().__init__(parent)
        self.title("Donar hojas")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.students = students
        frame = ttk.Frame(self, padding=16)
        frame.pack()
        ttk.Label(frame, text=f"Donante: {donor['last_name']}, {donor['first_name']}").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Receptor").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        names = [f"{s['last_name']}, {s['first_name']}" for s in students]
        self.recipient = ttk.Combobox(frame, values=names, width=34, state="readonly")
        self.recipient.grid(row=1, column=1, pady=4)
        self.recipient.current(0)
        ttk.Label(frame, text="Cantidad").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.quantity = ttk.Spinbox(frame, from_=1, to=9999, width=12)
        self.quantity.set(1)
        self.quantity.grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(frame, text="Nota (opcional)").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.note = ttk.Entry(frame, width=37)
        self.note.grid(row=3, column=1, pady=4)
        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="left")
        ttk.Button(buttons, text="Confirmar donacion", command=self.submit).pack(side="left", padx=(8, 0))

    def submit(self):
        try:
            quantity = int(self.quantity.get())
            if quantity <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Cantidad", "Ingrese una cantidad valida.", parent=self)
            return
        index = self.recipient.current()
        self.result = (int(self.students[index]["id"]), quantity, self.note.get())
        self.destroy()


class PrintDialog(tk.Toplevel):
    def __init__(self, parent, file_name: str, detected_pages: int | None):
        super().__init__(parent)
        self.title("Preparar impresion")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        frame = ttk.Frame(self, padding=16)
        frame.pack()
        ttk.Label(frame, text=file_name, font=("Segoe UI", 10, "bold"), wraplength=420).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Paginas del documento").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.pages = ttk.Spinbox(frame, from_=1, to=100000, width=12)
        self.pages.set(detected_pages or 1)
        self.pages.grid(row=1, column=1, sticky="w", pady=4)
        if detected_pages is None:
            ttk.Label(frame, text="No se pudo detectar automaticamente: verifique este valor.", style="Muted.TLabel").grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Label(frame, text="Copias").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.copies = ttk.Spinbox(frame, from_=1, to=1000, width=12)
        self.copies.set(1)
        self.copies.grid(row=3, column=1, sticky="w", pady=4)
        self.duplex = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Doble faz", variable=self.duplex).grid(row=4, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Label(frame, text="Nota (opcional)").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=4)
        self.note = ttk.Entry(frame, width=36)
        self.note.grid(row=5, column=1, pady=4)
        buttons = ttk.Frame(frame)
        buttons.grid(row=6, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="left")
        ttk.Button(buttons, text="Continuar", command=self.submit).pack(side="left", padx=(8, 0))

    def submit(self):
        try:
            pages = int(self.pages.get())
            copies = int(self.copies.get())
            if pages <= 0 or copies <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Impresion", "Paginas y copias deben ser mayores que cero.", parent=self)
            return
        self.result = (pages, copies, bool(self.duplex.get()), self.note.get())
        self.destroy()


if __name__ == "__main__":
    app = CopyManagementApp()
    app.mainloop()
