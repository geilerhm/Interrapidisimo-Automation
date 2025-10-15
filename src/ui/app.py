
import tkinter as tk
from tkinter import ttk, messagebox
import sv_ttk
import datetime
import threading
import time # For potential delays or sleep in automation
import sys
import os

from src.automation.web_actions import setup_driver, login, open_shipment_explorer, process_single_shipment, AuthenticationError
from src.config import LOGIN_URL # Only LOGIN_URL is needed from config now
from src.utils import create_or_load_excel # Import create_or_load_excel

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

class SettingsFrame(ttk.LabelFrame):
    def __init__(self, parent):
        super().__init__(parent, text="⚙️ Configuración", padding=15)
        self.columnconfigure(0, weight=1)

        # Dark/Light Mode Toggle
        theme_label = ttk.Label(self, text="Modo Oscuro/Claro:")
        theme_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.theme_toggle = ttk.Checkbutton(self, style="Switch.TCheckbutton", command=self._toggle_theme)
        self.theme_toggle.grid(row=0, column=1, sticky="e", padx=5, pady=5)

        # Show Browser Toggle
        browser_label = ttk.Label(self, text="Mostrar Navegador:")
        browser_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.show_browser_var = tk.BooleanVar(value=False) # Default to NOT showing browser
        self.show_browser_toggle = ttk.Checkbutton(self, style="Switch.TCheckbutton", variable=self.show_browser_var)
        self.show_browser_toggle.grid(row=1, column=1, sticky="e", padx=5, pady=5)

    def _toggle_theme(self):
        sv_ttk.set_theme("light" if sv_ttk.get_theme() == "dark" else "dark")

    def get_show_browser_setting(self):
        return self.show_browser_var.get()

    def disable_fields(self):
        self.theme_toggle.config(state="disabled")
        self.show_browser_toggle.config(state="disabled")

    def enable_fields(self):
        self.theme_toggle.config(state="normal")
        self.show_browser_toggle.config(state="normal")


class ProgressModal(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Proceso de Automatización")
        self.transient(parent)  # Make it a transient window for the parent
        self.grab_set()  # Make it modal
        self.resizable(False, False)

        # UI elements
        self.status_label = ttk.Label(self, text="Iniciando...", font=("", 12, "bold"), wraplength=300)
        self.status_label.pack(pady=20, padx=20)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=250)
        self.progress_bar.pack(pady=10, padx=20)

        self.current_item_label = ttk.Label(self, text="", wraplength=300)
        self.current_item_label.pack(pady=(0, 20), padx=20)

        # Center the modal after widgets are packed
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        # Prevent closing with window manager buttons
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        # Do nothing to prevent closing
        pass

    def update_status(self, message):
        self.status_label.config(text=message)
        self.update_idletasks()

    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.update_idletasks()

    def update_current_item(self, item_text):
        self.current_item_label.config(text=item_text)
        self.update_idletasks()


class AutomationController:
    def __init__(self, app_instance, username, password, guides, show_browser): # Add show_browser
        self.app = app_instance
        self.username = username
        self.password = password
        self.guides = guides
        self.show_browser = show_browser # Store show_browser
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
            self.driver = setup_driver(show_browser=self.show_browser)
            
            self.app.after(0, lambda: self.app.status_bar.set_status("Iniciando sesión..."))
            login(self.driver, self.username, self.password)
            
            self.app.after(0, lambda: self.app.status_bar.set_status("Abriendo explorador de envíos..."))
            open_shipment_explorer(self.driver)

            while current_shipment_index < total_shipments:
                shipment_to_process = self.guides[current_shipment_index]

                self.app.after(0, lambda: self.app.status_bar.set_status(f"Procesando guía {current_shipment_index + 1}/{total_shipments} ({shipment_to_process})..."))
                
                success_one, needs_reopen = process_single_shipment(
                    self.driver, 
                    shipment_to_process, 
                    lambda: self._update_progress_ui(processed_count + 1, total_shipments),
                    wb, ws, file_name
                )

                if needs_reopen:
                    self.app.after(0, lambda: Toast(self.app, f"⚠️ Redirección detectada. Reabriendo explorador...", success=False))
                    self.app.after(0, lambda: self.app.status_bar.set_status(f"Reabriendo explorador..."))
                    open_shipment_explorer(self.driver)
                    continue

                if not needs_reopen:
                    if success_one:
                        processed_count += 1
                        self.app.after(0, lambda: Toast(self.app, f"✅ Guía {shipment_to_process} procesada.", success=True))
                    else:
                        self.app.after(0, lambda: Toast(self.app, f"❌ Guía {shipment_to_process} falló.", success=False))
                    current_shipment_index += 1

                self._update_progress_ui(processed_count, total_shipments)

            # --- SUCCESS PATH ---
            self.app.after(0, lambda: Toast(self.app, "✅ Proceso de todas las guías completado.", success=True))
            self.app.after(1000, self.app.guides_frame.clear_entries) # Clear entries after 1s
            self.app.after(1000, lambda: self.app.guides_frame.on_key_release(None))
            self.app.after(1500, self._reset_ui_state) # Reset UI after 1.5s

        except AuthenticationError as e:
            # --- AUTHENTICATION FAILURE PATH ---
            error_message = str(e)
            self.app.after(0, lambda msg=error_message: Toast(self.app, f"❌ {msg}", success=False))
            self.app.after(3500, self._reset_ui_state) # Reset UI after 3.5s to ensure toast is read
        
        except Exception as e:
            # --- GENERIC FAILURE PATH ---
            print(f"Automation error: {e}")
            import traceback
            traceback.print_exc()
            error_message = str(e)
            self.app.after(0, lambda msg=error_message: Toast(self.app, f"❌ Error crítico: {msg}", success=False))
            self.app.after(3500, self._reset_ui_state) # Reset UI after 3.5s

        finally:
            # --- CLEANUP ---
            # The finally block is now only responsible for closing the browser driver.
            if self.driver:
                self.driver.quit()

    def _reset_ui_state(self):
        self.app.config(cursor="")
        self.app.status_bar.start_button.config(state="normal")
        self.app.credentials_frame.enable_fields()
        self.app.settings_frame.enable_fields() # Enable settings fields
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

        sv_ttk.set_theme("light")
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

    def create_body(self):
        body = ttk.Frame(self)
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1) # Added this line
        body.rowconfigure(1, weight=1)

        self.status_bar = StatusBar(body, self.start_bot_process)
        self.guides_frame = GuidesFrame(body, status_bar=self.status_bar)
        self.credentials_frame = CredentialsFrame(body)
        self.settings_frame = SettingsFrame(body) # Added this line

        self.credentials_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10), padx=(0, 5))
        self.settings_frame.grid(row=0, column=1, sticky="ew", pady=(0, 10), padx=(5, 0))
        self.guides_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.status_bar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

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
        show_browser = self.settings_frame.get_show_browser_setting() # Get setting

        self.config(cursor="watch")
        self.status_bar.start_button.config(state="disabled")
        self.credentials_frame.disable_fields()
        self.settings_frame.disable_fields() # Disable settings fields
        self.guides_frame.disable()
        self.status_bar.set_progress(0)
        self.status_bar.set_status("Iniciando proceso de automatización...")

        self.automation_controller = AutomationController(self, username, password, guides, show_browser) # Pass show_browser
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
        self.max_rows_per_column = 9
        self.entries = []  # List of lists, entries[col_idx][row_idx]
        self.string_vars = []  # List of lists, string_vars[col_idx][row_idx]
        self.last_active_entry = None # To keep track of the last entry that had focus

        controls_frame = ttk.Frame(self)
        controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        controls_frame.columnconfigure(0, weight=1)
        self.counter_label = ttk.Label(controls_frame, text="Guías Ingresadas: 0")
        self.counter_label.grid(row=0, column=0, sticky="w")
        clear_button = ttk.Button(controls_frame, text="🗑️ Limpiar Todo", command=self.confirm_clear)
        clear_button.grid(row=0, column=1, sticky="e")

        self.canvas = tk.Canvas(self, relief="flat", highlightthickness=0)
        self.h_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas.grid(row=1, column=0, sticky="nsew", columnspan=2)
        
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        # Initialize with one column and two entries
        self._add_column() # This creates the first column (col_idx 0)
        self._add_entry(0, 0)
        self._add_entry(0, 1)
        self._update_scrollbars() # Call to update scrollbars initially

    def _update_scrollbars(self):
        if len(self.entries) > 6:
            self.h_scrollbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        else:
            self.h_scrollbar.grid_forget()

    def _add_column(self):
        col_idx = len(self.entries)
        self.entries.append([])
        self.string_vars.append([])
        # Configure column weight for the scrollable frame
        self.scrollable_frame.columnconfigure(col_idx, weight=1)
        self._update_scrollbars() # Update scrollbars when a new column is added
        return col_idx

    def _add_entry(self, col_idx, row_idx, value=""):
        string_var = tk.StringVar(value=value)
        entry = ttk.Entry(self.scrollable_frame, width=15, textvariable=string_var)
        string_var.trace_add("write", lambda name, index, mode, sv=string_var, entry_widget=entry: self._validate_and_manage_entries(sv, entry_widget))
        entry.grid(row=row_idx, column=col_idx, padx=3, pady=3, ipady=2)
        entry.bind("<KeyRelease>", self.on_key_release)
        entry.bind("<Return>", self.on_enter_pressed)
        entry.bind("<<Paste>>", self.on_paste)
        self.entries[col_idx].append(entry)
        self.string_vars[col_idx].append(string_var)
        return entry

    def _validate_and_manage_entries(self, string_var, entry_widget):
        value = string_var.get()
        if len(value) > 12:
            string_var.set(value[:12])
            value = value[:12] # Update value after truncation

        # Update counter and button state
        self.on_key_release(event=None) 

        # Dynamic entry creation logic
        col_idx = entry_widget.grid_info()["column"]
        row_idx = entry_widget.grid_info()["row"]

        # If the current entry is filled to 12 chars and it's the last entry in its column
        if len(value) == 12 and row_idx == len(self.entries[col_idx]) - 1:
            if row_idx < self.max_rows_per_column - 1: # If not the last row in column
                # Add a new entry in the same column
                self._add_entry(col_idx, row_idx + 1)
            elif col_idx == len(self.entries) - 1: # If it's the last row in the last column
                # Add a new column and then add an entry to it
                new_col_idx = self._add_column()
                self._add_entry(new_col_idx, 0)
                self.scrollable_frame.update_idletasks() # Force update of layout
                self.canvas.configure(scrollregion=self.canvas.bbox("all")) # Explicitly update scrollregion
                self.canvas.after(0, lambda: self.canvas.xview_moveto(1.0)) # Defer scroll to new column

    def confirm_clear(self):
        if messagebox.askyesno("Confirmar Limpieza", "¿Estás seguro de que quieres borrar todas las guías ingresadas?"):
            self.clear_entries()
            self.on_key_release(None)

    def on_key_release(self, event):
        # Update the counter and button state
        guides = self.get_guides()
        self.counter_label.config(text=f"Guías Ingresadas: {len(guides)}")
        self.status_bar.toggle_start_button(not guides)

        # If a real key event occurred, store the widget
        if event and event.widget:
            self.last_active_entry = event.widget

    def on_enter_pressed(self, event):
        current_widget = event.widget
        grid_info = current_widget.grid_info()
        current_row, current_col = grid_info["row"], grid_info["column"]

        # Try to move to the next entry in the current column
        if current_row < len(self.entries[current_col]) - 1:
            self.entries[current_col][current_row + 1].focus()
        elif current_col < len(self.entries) - 1: # Move to the first entry of the next column
            self.entries[current_col + 1][0].focus()
        else: # If it's the last entry in the last column, and a new one was created by _validate_and_manage_entries
            # Check if a new entry was created at (current_col, current_row + 1) or (current_col + 1, 0)
            if current_row < self.max_rows_per_column - 1 and len(self.entries[current_col]) > current_row + 1:
                self.entries[current_col][current_row + 1].focus()
            elif len(self.entries) > current_col + 1 and len(self.entries[current_col + 1]) > 0:
                self.entries[current_col + 1][0].focus()
            else:
                # If no new entry was created (e.g., current entry is not full), just stay focused
                pass

        return "break" # Prevent default Tkinter behavior

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
            # Ensure column exists
            if col_idx >= len(self.entries):
                self._add_column()
                self.scrollable_frame.update_idletasks() # Force update of layout
                self.canvas.configure(scrollregion=self.canvas.bbox("all")) # Explicitly update scrollregion
                self.canvas.after(0, lambda: self.canvas.xview_moveto(1.0)) # Defer scroll to new column

            # Ensure row exists in current column, or add new entry if needed
            if row_idx >= len(self.entries[col_idx]):
                self._add_entry(col_idx, row_idx)
            
            entry_to_fill = self.entries[col_idx][row_idx]
            entry_to_fill.delete(0, "end")
            entry_to_fill.insert(0, pasted_guides[guide_index][:12])

            guide_index += 1
            row_idx += 1

            if row_idx >= self.max_rows_per_column: # Move to next column if current is full
                row_idx = 0
                col_idx += 1
        
        # After pasting, focus on the next logical entry
        next_row, next_col = row_idx, col_idx
        if next_col >= len(self.entries): # If we ended up needing a new column
            next_col = self._add_column()
            next_row = 0
            self._add_entry(next_col, next_row)
            self.scrollable_frame.update_idletasks() # Force update of layout
            self.canvas.configure(scrollregion=self.canvas.bbox("all")) # Explicitly update scrollregion
            self.canvas.after(0, lambda: self.canvas.xview_moveto(1.0)) # Defer scroll to new column
        elif next_row >= len(self.entries[next_col]): # If we ended up needing a new entry in existing column
            self._add_entry(next_col, next_row)
        
        self.entries[next_col][next_row].focus()
        self.on_key_release(event=None) # Update counter and button state

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
        # Clear all existing entries and reset to initial state (2 entries in 1st column)
        for col_string_vars in self.string_vars:
            for string_var in col_string_vars:
                string_var.set("")
        
        # Remove all but the first column and its first two entries
        for col_idx in range(len(self.entries) -1, 0, -1):
            for entry in self.entries[col_idx]:
                entry.destroy()
            del self.entries[col_idx]
            del self.string_vars[col_idx]
        
        # Ensure only two entries in the first column are visible and clear
        for row_idx in range(len(self.entries[0]) -1, 1, -1):
            self.entries[0][row_idx].destroy()
            del self.entries[0][row_idx]
            del self.string_vars[0][row_idx]
        
        # If for some reason there are less than 2 entries, add them back
        while len(self.entries[0]) < 2:
            self._add_entry(0, len(self.entries[0]))

        self.entries[0][0].focus() # Focus on the first entry
        self.canvas.xview_moveto(0.0) # Scroll to the beginning
        self.on_key_release(None) # Update counter and button state
        self._update_scrollbars() # Update scrollbars after clearing

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
