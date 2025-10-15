
import tkinter as tk
from tkinter import ttk, messagebox
import sv_ttk
import datetime
import threading
import time # For potential delays or sleep in automation
import sys
import os

# Ensure project root is in path for web_actions and config imports
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from web_actions import setup_driver, login, open_shipment_explorer, process_single_shipment
from config import LOGIN_URL # Only LOGIN_URL is needed from config now
from utils import create_or_load_excel # Import create_or_load_excel

class Toast(tk.Toplevel):
    """A temporary, toast-like notification window."""
    def __init__(self, parent, message, success=True):
        super().__init__(parent)
        self.overrideredirect(True)
        style = "Success.TLabel" if success else "Error.TLabel"
        self.label = ttk.Label(self, text=message, padding=15, style=style, anchor="center")
        self.label.pack()
        s = ttk.Style()
        s.configure("Success.TLabel", background="#2a9d8f", foreground="#ffffff", font=("", 10, "bold"))
        s.configure("Error.TLabel", background="#e76f51", foreground="#ffffff", font=("", 10, "bold"))
        parent.update_idletasks()
        parent_x, parent_y, parent_width = parent.winfo_x(), parent.winfo_y(), parent.winfo_width()
        self.update_idletasks()
        toast_width = self.winfo_width()
        x = parent_x + (parent_width // 2) - (toast_width // 2)
        y = parent_y + 50
        self.geometry(f"+{x}+{y}")
        self.after(3000, self.destroy)

class AutomationController:
    def __init__(self, app_instance, username, password, guides):
        self.app = app_instance
        self.username = username
        self.password = password
        self.guides = guides
        self.driver = None
        self.stop_event = threading.Event()

    def _update_progress_ui(self, processed_count, total_count):
        percentage = int((processed_count / total_count) * 100)
        self.app.after(0, lambda: self.app.status_bar.set_progress(percentage))
        self.app.after(0, lambda: self.app.status_bar.set_status(f"Procesando guía {processed_count}/{total_count} ({percentage}%)"))

    def run_automation(self):
        total_shipments = len(self.guides)
        processed_count = 0
        current_shipment_index = 0
        
        wb, ws, file_name = create_or_load_excel()

        try:
            self.app.after(0, lambda: self.app.status_bar.set_status("Iniciando navegador..."))
            self.driver = setup_driver()
            
            self.app.after(0, lambda: self.app.status_bar.set_status("Iniciando sesión..."))
            login(self.driver, self.username, self.password)
            
            # Open shipment explorer ONCE before the loop
            self.app.after(0, lambda: self.app.status_bar.set_status("Abriendo explorador de envíos..."))
            open_shipment_explorer(self.driver)

            while current_shipment_index < total_shipments:
                shipment_to_process = self.guides[current_shipment_index]

                self.app.after(0, lambda: self.app.status_bar.set_status(f"Procesando guía {current_shipment_index + 1}/{total_shipments} ({shipment_to_process})..."))
                
                success_one, needs_reopen = process_single_shipment(
                    self.driver, 
                    shipment_to_process, 
                    lambda: self._update_progress_ui(processed_count + 1, total_shipments), # Callback for one item
                    wb, ws, file_name
                )

                if needs_reopen:
                    self.app.after(0, lambda: Toast(self.app, f"⚠️ Redirección detectada. Reabriendo explorador para {shipment_to_process}...", success=False))
                    self.app.after(0, lambda: self.app.status_bar.set_status(f"Reabriendo explorador para {shipment_to_process}..."))
                    open_shipment_explorer(self.driver)
                    continue # This will re-process the same shipment_to_process

                # ONLY advance to the next shipment if no re-open was needed
                if not needs_reopen:
                    if success_one:
                        processed_count += 1
                        self.app.after(0, lambda: Toast(self.app, f"✅ Guía {shipment_to_process} procesada.", success=True))
                    else:
                        self.app.after(0, lambda: Toast(self.app, f"❌ Guía {shipment_to_process} falló.", success=False))
                    
                    current_shipment_index += 1 # Moved inside the 'if not needs_reopen' block

                # Update overall progress after each shipment
                self._update_progress_ui(processed_count, total_shipments)

            # After loop finishes (all shipments attempted)
            self.app.after(0, lambda: Toast(self.app, "✅ Proceso de todas las guías completado.", success=True))
            self.app.after(0, self.app.guides_frame.clear_entries)
            self.app.after(0, lambda: self.app.guides_frame.on_key_release(None)) # Update counter and button state

        except Exception as e:
            print(f"Automation error: {e}")
            import traceback
            traceback.print_exc()
            error_message = str(e)
            self.app.after(0, lambda msg=error_message: Toast(self.app, f"❌ Error crítico: {msg}", success=False))
        finally:
            if self.driver:
                self.driver.quit()
            self.app.after(0, self._reset_ui_state)

    def _reset_ui_state(self):
        self.app.config(cursor="")
        self.app.status_bar.start_button.config(state="normal")
        self.app.credentials_frame.enable_fields()
        self.app.guides_frame.enable()
        self.app.status_bar.set_status("Listo para iniciar." if self.app.guides_frame.get_guides() else "Ingrese guías para comenzar.")
        self.app.status_bar.set_progress(0)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Interrapidisimo Bot v1.0")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        app_width = int(screen_width * 0.7)
        app_height = int(screen_height * 0.8)
        center_x = int((screen_width / 2) - (app_width / 2))
        center_y = int((screen_height / 2) - (app_height / 2))
        self.geometry(f"{app_width}x{app_height}+{center_x}+{center_y}")
        self.minsize(800, 600)

        sv_ttk.set_theme("dark")
        style = ttk.Style()
        style.map("TEntry", fieldbackground=[("disabled", "#3c3c3c")])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.create_header()
        self.create_body()
        self.create_footer()
        self.automation_thread = None

    def create_header(self):
        header = ttk.Frame(self, padding=(20, 10, 20, 10))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        title = ttk.Label(header, text="Interrapidisimo Bot", font=("", 24, "bold"))
        title.grid(row=0, column=0, sticky="w")
        theme_toggle = ttk.Checkbutton(header, style="Switch.TCheckbutton", command=self.toggle_theme)
        theme_toggle.grid(row=0, column=1, sticky="e")

    def toggle_theme(self):
        sv_ttk.set_theme("light" if sv_ttk.get_theme() == "dark" else "dark")

    def create_body(self):
        body = ttk.Frame(self)
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        self.status_bar = StatusBar(body, self.start_bot_process)
        self.guides_frame = GuidesFrame(body, status_bar=self.status_bar)
        self.credentials_frame = CredentialsFrame(body)

        self.credentials_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.guides_frame.grid(row=1, column=0, sticky="nsew")
        self.status_bar.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    def create_footer(self):
        footer = ttk.Frame(self, padding=10)
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        current_year = datetime.date.today().year
        copyright_label = ttk.Label(footer, text=f"© {current_year} Geiler Orlando Hipia Mejia. Todos los derechos reservados.", font=("", 8))
        copyright_label.grid(sticky="e")

    def start_bot_process(self):
        guides = self.guides_frame.get_guides()
        if not guides:
            Toast(self, "❌ Error: No hay guías para procesar.", success=False)
            return
        
        username = self.credentials_frame.user_entry.get()
        password = self.credentials_frame.pass_entry.get()

        self.config(cursor="watch")
        self.status_bar.start_button.config(state="disabled")
        self.credentials_frame.disable_fields()
        self.guides_frame.disable()
        self.status_bar.set_progress(0)
        self.status_bar.set_status("Iniciando proceso de automatización...")

        self.automation_controller = AutomationController(self, username, password, guides)
        self.automation_thread = threading.Thread(target=self.automation_controller.run_automation, daemon=True)
        self.automation_thread.start()

class CredentialsFrame(ttk.LabelFrame):
    def __init__(self, parent):
        super().__init__(parent, text="🔑 Credenciales", padding=15)
        self.columnconfigure(1, weight=1)
        self.user_entry = self._create_row("Username:", 0, "age7179.dagua")
        self.pass_entry = self._create_row("Password:", 1, "1111111111")

    def _create_row(self, label_text, row, default_value):
        label = ttk.Label(self, text=label_text)
        label.grid(row=row, column=0, sticky="w", padx=5, pady=10)
        entry = ttk.Entry(self, width=30)
        entry.insert(0, default_value)
        entry.grid(row=row, column=1, sticky="ew", padx=5)
        return entry

    def disable_fields(self):
        self.user_entry.config(state="disabled")
        self.pass_entry.config(state="disabled")

    def enable_fields(self):
        self.user_entry.config(state="normal")
        self.pass_entry.config(state="normal")

class GuidesFrame(ttk.LabelFrame):
    def __init__(self, parent, status_bar):
        super().__init__(parent, text="📦 Números de Guía", padding=15)
        self.status_bar = status_bar
        self.num_rows = 10
        self.entries = []
        self.string_vars = [] # To hold StringVar for each entry

        controls_frame = ttk.Frame(self)
        controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        controls_frame.columnconfigure(0, weight=1)
        self.counter_label = ttk.Label(controls_frame, text="Guías Ingresadas: 0")
        self.counter_label.grid(row=0, column=0, sticky="w")
        clear_button = ttk.Button(controls_frame, text="🗑️ Limpiar Todo", command=self.confirm_clear)
        clear_button.grid(row=0, column=1, sticky="e")

        self.canvas = tk.Canvas(self, relief="flat", highlightthickness=0)
        self.h_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=1, column=1, sticky="ns")
        self.h_scrollbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        for j in range(6):
            self.add_column(j)

    def add_column(self, col_index):
        col_entries = []
        col_string_vars = []
        for i in range(self.num_rows):
            string_var = tk.StringVar()
            entry = ttk.Entry(self.scrollable_frame, width=15, textvariable=string_var)
            string_var.trace_add("write", lambda name, index, mode, sv=string_var, entry_widget=entry: self._validate_length(sv, entry_widget))
            entry.grid(row=i, column=col_index, padx=3, pady=3, ipady=2)
            entry.bind("<KeyRelease>", self.on_key_release)
            entry.bind("<Return>", self.on_enter_pressed)
            entry.bind("<<Paste>>", self.on_paste)
            col_entries.append(entry)
            col_string_vars.append(string_var)
        self.entries.append(col_entries)
        self.string_vars.append(col_string_vars)

    def _validate_length(self, string_var, entry_widget):
        value = string_var.get()
        if len(value) > 12:
            string_var.set(value[:12])
        # Always call on_key_release to update counter and button state
        # and potentially trigger auto-advance if length is 12
        self.on_key_release(event=None) # Pass None as event, as we don't need it here

    def confirm_clear(self):
        if messagebox.askyesno("Confirmar Limpieza", "¿Estás seguro de que quieres borrar todas las guías ingresadas?"):
            self.clear_entries()
            self.on_key_release(None)

    def on_key_release(self, event):
        guides = self.get_guides()
        self.counter_label.config(text=f"Guías Ingresadas: {len(guides)}")
        self.status_bar.toggle_start_button(not guides)

        # Auto-advance logic only if a real key event occurred
        if event and event.widget:
            current_widget = event.widget
            if len(current_widget.get()) == 12:
                self.on_enter_pressed(event)

    def on_enter_pressed(self, event):
        current_widget = event.widget
        grid_info = current_widget.grid_info()
        current_row, current_col = grid_info["row"], grid_info["column"]
        next_row, next_col = current_row + 1, current_col
        if next_row >= self.num_rows:
            next_row = 0
            next_col += 1
        if next_col >= len(self.entries):
            self.add_column(next_col)
            self.canvas.update_idletasks()
            self.canvas.xview_moveto(1.0)
        self.entries[next_col][next_row].focus()
        return "break"

    def on_paste(self, event):
        clipboard_content = self.clipboard_get()
        pasted_guides = [g.strip() for g in clipboard_content.split('\n') if g.strip()]

        if not pasted_guides:
            return "break"

        current_widget = event.widget
        grid_info = current_widget.grid_info()
        current_row, current_col = grid_info["row"], grid_info["column"]

        guide_index = 0
        col_idx = current_col
        row_idx = current_row

        while guide_index < len(pasted_guides):
            if col_idx >= len(self.entries):
                self.add_column(col_idx)
                self.canvas.update_idletasks()
                self.canvas.xview_moveto(1.0)

            if row_idx >= self.num_rows:
                row_idx = 0
                col_idx += 1
                if col_idx >= len(self.entries):
                    self.add_column(col_idx)
                    self.canvas.update_idletasks()
                    self.canvas.xview_moveto(1.0)

            entry_to_fill = self.entries[col_idx][row_idx]
            entry_to_fill.delete(0, "end")
            entry_to_fill.insert(0, pasted_guides[guide_index][:12])

            guide_index += 1
            row_idx += 1
        
        next_row, next_col = row_idx, col_idx
        if next_row >= self.num_rows:
            next_row = 0
            next_col += 1
        if next_col >= len(self.entries):
            self.add_column(next_col)
            self.canvas.update_idletasks()
            self.canvas.xview_moveto(1.0)
        self.entries[next_col][next_row].focus()

        self.on_key_release(event=None)

        return "break"

    def get_guides(self):
        guides = []
        for col in self.entries:
            for entry in col:
                guide = entry.get().strip()
                if guide:
                    guides.append(guide)
        return guides

    def disable(self):
        for col in self.entries:
            for entry in col:
                entry.config(state="disabled")

    def enable(self):
        for col in self.entries:
            for entry in col:
                entry.config(state="normal")

    def clear_entries(self):
        for col_string_vars in self.string_vars:
            for string_var in col_string_vars:
                string_var.set("")

class StatusBar(ttk.Frame):
    def __init__(self, parent, start_command):
        super().__init__(parent)
        self.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(self, text="Ingrese guías para comenzar.")
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode='determinate')
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5, ipady=8)
        self.start_button = ttk.Button(self, text="Iniciar Bot", style="Accent.TButton", command=start_command, state="disabled")
        self.start_button.grid(row=0, column=1, sticky="e")

    def toggle_start_button(self, is_empty):
        self.start_button.config(state="disabled" if is_empty else "normal")
        self.set_status("Listo para iniciar." if not is_empty else "Ingrese guías para comenzar.")

    def set_progress(self, value):
        self.progress_bar["value"] = value

    def set_status(self, text):
        self.status_label.config(text=text)

if __name__ == "__main__":
    app = App()
    app.mainloop()
