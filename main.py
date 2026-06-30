import customtkinter as ctk
import win32api
import win32print
import win32service
import win32serviceutil
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import re
import shutil
import subprocess

NFE_NUMBER_PATTERN = re.compile(r"(?<!\d)\d{20}55\d{3}(\d{9})\d{10}(?!\d)")

def restart():
    service_name = "Spooler"
    try:
        win32serviceutil.RestartService(service_name)
    except Exception:
        try:
            win32serviceutil.StopService(service_name)
            win32serviceutil.WaitForServiceStatus(
                service_name,
                win32service.SERVICE_STOPPED,
                30
            )
        except Exception:
            pass
        win32serviceutil.StartService(service_name)


def _printer_info_value(info, key, index):
    if isinstance(info, dict):
        return info.get(key, "N/A")
    try:
        return info[index]
    except Exception:
        return "N/A"


def get_printer_info():
    try:
        default_printer = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(default_printer)
        info = win32print.GetPrinter(handle, 2)
        win32print.ClosePrinter(handle)

        printer_name = _printer_info_value(info, "pPrinterName", 1)
        driver_name = _printer_info_value(info, "pDriverName", 4)
        port_name = _printer_info_value(info, "pPortName", 3)

        return (
            f"Impressora predeterminada: {printer_name}\n"
            f"Controlador: {driver_name}\n"
            f"Porta: {port_name}"
        )
    except Exception as error:
        return f"Erro ao obter informação da impressora: {error}"


def list_printers():
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        printers = win32print.EnumPrinters(flags, None, 1)
        printer_names = [printer[2] for printer in printers] if printers else ["Nenhuma impressora encontrada"]
        return "Impressoras instaladas:\n" + "\n".join(printer_names)
    except Exception as error:
        return f"Erro ao listar impressoras: {error}"


def refresh_info():
    info_label.configure(text=get_printer_info())


def show_printers():
    info_label.configure(text=list_printers())


def get_print_queue():
    try:
        printer_name = win32print.GetDefaultPrinter()
        handle = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
        win32print.ClosePrinter(handle)

        if not jobs:
            return f"Fila de impressão vazia para: {printer_name}"

        lines = [f"Fila de impressão ({printer_name}):"]
        for job in jobs:
            job_id = _printer_info_value(job, "JobId", 0)
            document = _printer_info_value(job, "pDocument", 4)
            status = _printer_info_value(job, "Status", 6)
            lines.append(f"ID {job_id} - {document} - Status {status}")
        return "\n".join(lines)
    except Exception as error:
        return f"Erro ao obter fila de impressão: {error}"


def show_print_queue():
    try:
        printer_name = win32print.GetDefaultPrinter()
        subprocess.Popen([
            "rundll32.exe",
            "printui.dll,PrintUIEntry",
            "/o",
            "/n",
            printer_name
        ])
        info_label.configure(text=f"Fila de impressão aberta para: {printer_name}")
    except Exception as error:
        try:
            subprocess.Popen(["control", "printers"])
            info_label.configure(text="Painel de impressoras aberto.")
        except Exception:
            info_label.configure(text=f"Erro ao abrir fila de impressão: {error}")


def cancel_print_queue():
    try:
        printer_name = win32print.GetDefaultPrinter()
    except Exception as error:
        info_label.configure(text=f"Erro ao obter impressora padrão: {error}")
        return

    confirmed = messagebox.askyesno(
        "Cancelar fila de impressão",
        f"Deseja cancelar todos os trabalhos pendentes de:\n{printer_name}?"
    )
    if not confirmed:
        return

    handle = None
    try:
        handle = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)

        if not jobs:
            info_label.configure(text=f"Nenhum trabalho pendente em: {printer_name}")
            return

        cancelled_count = 0
        for job in jobs:
            job_id = _printer_info_value(job, "JobId", 0)
            win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_CANCEL)
            cancelled_count += 1

        info_label.configure(
            text=f"Cancelado(s) {cancelled_count} trabalho(s) de: {printer_name}"
        )
    except Exception as error:
        info_label.configure(text=f"Erro ao cancelar fila de impressão: {error}")
    finally:
        if handle is not None:
            win32print.ClosePrinter(handle)


def open_control_printers():
    try:
        subprocess.Popen(["control", "printers"])
        info_label.configure(text="Painel de impressão aberto.")
    except Exception as error:
        info_label.configure(text=f"Erro ao abrir Control Printers: {error}")


def extract_nfe_number(filename):
    match = NFE_NUMBER_PATTERN.search(filename)
    if not match:
        return None
    return int(match.group(1))


def show_nfe_print_selection(folder, pdf_files):
    nfe_files = []
    unidentified_pdf_files = []
    for pdf_file in pdf_files:
        nfe_number = extract_nfe_number(os.path.basename(pdf_file))
        if nfe_number is not None:
            nfe_files.append((nfe_number, pdf_file))
        else:
            unidentified_pdf_files.append(pdf_file)

    nfe_files.sort(key=lambda item: item[0])
    min_nfe = nfe_files[0][0] if nfe_files else None
    max_nfe = nfe_files[-1][0] if nfe_files else None
    selected_files = None

    dialog = ctk.CTkToplevel(app)
    dialog.title("Seleção de NF-es para impressão")
    dialog.geometry("440x430")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()

    mode_var = ctk.StringVar(value="nfe_all" if nfe_files else "pdf_all")
    start_var = ctk.StringVar(value=str(min_nfe) if min_nfe is not None else "")
    end_var = ctk.StringVar(value=str(max_nfe) if max_nfe is not None else "")

    container = ctk.CTkFrame(dialog)
    container.pack(padx=20, pady=20, fill="both", expand=True)

    title_label = ctk.CTkLabel(
        container,
        text="Informações da pasta",
        font=ctk.CTkFont(size=16, weight="bold")
    )
    title_label.pack(anchor="w", pady=(0, 10))

    folder_info_label = ctk.CTkLabel(
        container,
        text=(
            f"PDFs encontrados: {len(pdf_files)}\n"
            f"NF-es detectadas: {len(nfe_files)}\n"
            f"PDFs sem NF-e detectada: {len(unidentified_pdf_files)}\n\n"
            f"Menor NF-e encontrada: {min_nfe if min_nfe is not None else 'N/A'}\n"
            f"Maior NF-e encontrada: {max_nfe if max_nfe is not None else 'N/A'}"
        ),
        justify="left"
    )
    folder_info_label.pack(anchor="w", fill="x", pady=(0, 14))

    all_radio = ctk.CTkRadioButton(
        container,
        text="Imprimir todas as NF-es",
        variable=mode_var,
        value="nfe_all",
        state="normal" if nfe_files else "disabled"
    )
    all_radio.pack(anchor="w", pady=(0, 8))

    range_radio = ctk.CTkRadioButton(
        container,
        text="Imprimir intervalo específico",
        variable=mode_var,
        value="range",
        state="normal" if nfe_files else "disabled"
    )
    range_radio.pack(anchor="w", pady=(0, 12))

    all_pdfs_radio = ctk.CTkRadioButton(
        container,
        text="Imprimir todos os PDFs mesmo sem NF-e",
        variable=mode_var,
        value="pdf_all"
    )
    all_pdfs_radio.pack(anchor="w", pady=(0, 12))

    range_frame = ctk.CTkFrame(container, fg_color="transparent")
    range_frame.pack(fill="x", pady=(0, 16))
    range_frame.grid_columnconfigure(1, weight=1)

    start_label = ctk.CTkLabel(range_frame, text="NF-e inicial:")
    start_label.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    start_entry = ctk.CTkEntry(range_frame, textvariable=start_var)
    start_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))

    end_label = ctk.CTkLabel(range_frame, text="NF-e final:")
    end_label.grid(row=1, column=0, sticky="w", padx=(0, 10))
    end_entry = ctk.CTkEntry(range_frame, textvariable=end_var)
    end_entry.grid(row=1, column=1, sticky="ew")

    def update_range_state(*_):
        state = "normal" if mode_var.get() == "range" and nfe_files else "disabled"
        start_entry.configure(state=state)
        end_entry.configure(state=state)

    def confirm_print():
        nonlocal selected_files
        if mode_var.get() == "pdf_all":
            selected_files = pdf_files
            dialog.destroy()
            return

        if mode_var.get() == "nfe_all":
            selected_files = [pdf_file for _, pdf_file in nfe_files]
            dialog.destroy()
            return

        try:
            start_nfe = int(start_var.get().strip())
            end_nfe = int(end_var.get().strip())
        except ValueError:
            messagebox.showerror(
                "Intervalo inválido",
                "Informe apenas números nos campos de NF-e.",
                parent=dialog
            )
            return

        if start_nfe > end_nfe:
            messagebox.showerror(
                "Intervalo inválido",
                "A NF-e inicial deve ser menor ou igual a NF-e final.",
                parent=dialog
            )
            return

        selected_files = [
            pdf_file
            for nfe_number, pdf_file in nfe_files
            if start_nfe <= nfe_number <= end_nfe
        ]
        if not selected_files:
            messagebox.showerror(
                "Nenhuma NF-e encontrada",
                "Não existem PDFs no intervalo informado.",
                parent=dialog
            )
            return

        dialog.destroy()

    def cancel_print():
        dialog.destroy()

    mode_var.trace_add("write", update_range_state)
    update_range_state()

    buttons_frame = ctk.CTkFrame(container, fg_color="transparent")
    buttons_frame.pack(fill="x")
    buttons_frame.grid_columnconfigure((0, 1), weight=1)

    print_button = ctk.CTkButton(buttons_frame, text="Imprimir", command=confirm_print)
    print_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

    cancel_button = ctk.CTkButton(buttons_frame, text="Cancelar", command=cancel_print)
    cancel_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))

    dialog.protocol("WM_DELETE_WINDOW", cancel_print)
    dialog.focus()
    app.wait_window(dialog)

    return selected_files


def print_folder_files():
    global selected_folder
    folder = selected_folder or filedialog.askdirectory()
    if not folder:
        return
    selected_folder = folder
    folder_label.configure(text=f"Pasta selecionada: {folder}")
    update_folder_preview(folder)

    pdf_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    ]
    if not pdf_files:
        info_label.configure(text=f"Nenhum PDF encontrado em:\n{folder}")
        return

    files_to_print = show_nfe_print_selection(folder, pdf_files)
    if not files_to_print:
        return

    try:
        for pdf_file in files_to_print:
            win32api.ShellExecute(0, "print", os.path.abspath(pdf_file), "", ".", 0)
        info_label.configure(
            text=f"Enviados para impressão {len(files_to_print)} PDF(s) de:\n{folder}"
        )
    except Exception as error:
        info_label.configure(text=f"Erro ao imprimir arquivos da pasta: {error}")


def clear_folder_preview():
    for widget in preview_frame.winfo_children():
        widget.destroy()


def format_file_size(size_bytes):
    units = ["B", "KB", "MB", "GB"]
    size = float(size_bytes)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} {unit}"
        size /= 1024


def update_folder_preview(folder):
    clear_folder_preview()

    try:
        files = sorted(os.listdir(folder), key=str.lower)
    except Exception as error:
        preview_summary_label.configure(text="Não foi possível ler a pasta")
        preview_status_label.configure(text="Status: erro ao carregar a prévia")
        ctk.CTkLabel(
            preview_frame,
            text=f"Erro: {error}",
            justify="left",
            anchor="w",
            wraplength=360
        ).pack(anchor="w", fill="x", padx=12, pady=10)
        return

    pdf_count = sum(1 for filename in files if filename.lower().endswith(".pdf"))
    nfe_count = 0
    pdf_total_size = 0

    for filename in files:
        if not filename.lower().endswith(".pdf"):
            continue

        file_path = os.path.join(folder, filename)
        if os.path.isfile(file_path):
            pdf_total_size += os.path.getsize(file_path)
        if extract_nfe_number(filename) is not None:
            nfe_count += 1

    preview_summary_label.configure(
        text=f"{len(files)} arquivo(s) | {pdf_count} PDF(s)"
    )
    preview_status_label.configure(
        text=(
            f"NF-es detectadas: {nfe_count} | "
            f"Tamanho dos PDFs: {format_file_size(pdf_total_size)} | "
            f"Atualizado: {datetime.now().strftime('%H:%M')}"
        )
    )

    if not files:
        ctk.CTkLabel(
            preview_frame,
            text="A pasta selecionada está vazia.",
            justify="left",
            anchor="w"
        ).pack(anchor="w", fill="x", padx=12, pady=10)
        return

    for filename in files:
        is_pdf = filename.lower().endswith(".pdf")
        nfe_number = extract_nfe_number(filename) if is_pdf else None
        details = []
        details.append("PDF" if is_pdf else "Arquivo")
        if nfe_number is not None:
            details.append(f"NF-e {nfe_number}")

        file_row = ctk.CTkFrame(
            preview_frame,
            fg_color="#202020" if is_pdf else "#191919",
            corner_radius=6
        )
        file_row.pack(fill="x", padx=8, pady=(0, 8))

        ctk.CTkLabel(
            file_row,
            text=filename,
            justify="left",
            anchor="w",
            wraplength=360
        ).pack(anchor="w", fill="x", padx=10, pady=(8, 2))

        ctk.CTkLabel(
            file_row,
            text=" | ".join(details),
            text_color="#9ca3af",
            justify="left",
            anchor="w"
        ).pack(anchor="w", fill="x", padx=10, pady=(0, 8))


app = ctk.CTk()
app.geometry("880x520")
app.title("Printer Master")
app.grid_columnconfigure(0, weight=0)
app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

left_panel = ctk.CTkFrame(app, fg_color="transparent")
left_panel.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=20)

right_panel = ctk.CTkFrame(app, fg_color="#111111", corner_radius=8)
right_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=20)
right_panel.grid_rowconfigure(2, weight=1)
right_panel.grid_columnconfigure(0, weight=1)

preview_title_label = ctk.CTkLabel(
    right_panel,
    text="Prévia da pasta",
    font=ctk.CTkFont(size=16, weight="bold"),
    anchor="w"
)
preview_title_label.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 2))

preview_summary_label = ctk.CTkLabel(
    right_panel,
    text="Selecione uma pasta para visualizar os arquivos.",
    text_color="#9ca3af",
    anchor="w"
)
preview_summary_label.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))

preview_frame = ctk.CTkScrollableFrame(
    right_panel,
    fg_color="#151515",
    scrollbar_button_color="#2b2b2b",
    scrollbar_button_hover_color="#3a3a3a"
)
preview_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 8))

preview_status_label = ctk.CTkLabel(
    right_panel,
    text="Status: aguardando seleção de pasta",
    text_color="#8b949e",
    fg_color="#0d0d0d",
    corner_radius=6,
    anchor="w",
    height=26
)
preview_status_label.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 14))

btn_restart = ctk.CTkButton(
    left_panel,
    text="Reiniciar Spooler",
    command=restart
)

btn_refresh = ctk.CTkButton(
    left_panel,
    text="Atualizar informações da impressora",
    command=refresh_info
)

btn_list = ctk.CTkButton(
    left_panel,
    text="Listar impressoras",
    command=show_printers
)

btn_print_queue = ctk.CTkButton(
    left_panel,
    text="Fila de impressão",
    command=show_print_queue
)

btn_cancel_print_queue = ctk.CTkButton(
    left_panel,
    text="Cancelar fila de impressão",
    command=cancel_print_queue
)

btn_control_printers = ctk.CTkButton(
    left_panel,
    text="Abrir painel de impressoras",
    command=open_control_printers
)

def select_folder():
    global selected_folder
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        folder_label.configure(text=f"Pasta selecionada: {folder}")
        update_folder_preview(folder)

selected_folder = None

btn_select_folder = ctk.CTkButton(
    left_panel,
    text="Selecionar pasta",
    command=select_folder
)

btn_print_folder = ctk.CTkButton(
    left_panel,
    text="Imprimir arquivos da pasta",
    command=print_folder_files
)

info_label = ctk.CTkLabel(
    left_panel,
    text="Pressione um botão para obter informações sobre a impressora ou listar as impressoras instaladas.",
    justify="left",
    wraplength=330
)

btn_restart.pack(padx=20, pady=(20, 10), fill="x")
btn_refresh.pack(padx=20, pady=10, fill="x")
btn_list.pack(padx=20, pady=10, fill="x")
btn_select_folder.pack(padx=20, pady=10, fill="x")
btn_print_folder.pack(padx=20, pady=10, fill="x")
btn_print_queue.pack(padx=20, pady=10, fill="x")
btn_cancel_print_queue.pack(padx=20, pady=10, fill="x")
btn_control_printers.pack(padx=20, pady=10, fill="x")
info_label.pack(padx=20, pady=(10, 20), fill="x")

folder_label = ctk.CTkLabel(
    left_panel,
    text="Pasta selecionada: Nenhuma",
    justify="left",
    wraplength=330
)
folder_label.pack(padx=20, pady=(0, 20), fill="x")

app.mainloop()
