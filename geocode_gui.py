"""CustomTkinter GUI for geocoding Excel address tables."""
from __future__ import annotations

import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from geocode_addresses import (
    AddressColumnError,
    GeocodingCancelledError,
    DEFAULT_USER_AGENT,
    geocode_excel_file,
    get_excel_columns,
    infer_address_column,
    reveal_in_file_manager,
    export_to_shapefile,
)

from map import generate_map


class MultiSelectDropdown(ctk.CTkFrame):
    def __init__(self, master, placeholder: str = "Select items", command=None) -> None:
        super().__init__(master)
        self._placeholder = placeholder
        self._command = command
        self._values: list[str] = []
        self._variables: dict[str, ctk.BooleanVar] = {}
        self._menu: ctk.CTkToplevel | None = None

        self._button = ctk.CTkButton(self, text=self._placeholder, anchor="w", command=self._toggle_menu)
        self._button.pack(fill="x")
        self._button.configure(state="disabled")

    def configure_values(self, values: list[str], selected: list[str] | None = None) -> None:
        if self._menu is not None and self._menu.winfo_exists():
            self._menu.destroy()
            self._menu = None

        self._values = list(values)
        self._variables = {value: ctk.BooleanVar(value=False) for value in self._values}
        if selected:
            for value in selected:
                if value in self._variables:
                    self._variables[value].set(True)

        self._button.configure(state="normal" if self._values else "disabled")
        self._update_button_text()

    def clear(self) -> None:
        self.configure_values([])

    def get_selected(self) -> list[str]:
        return [name for name, var in self._variables.items() if var.get()]

    def destroy(self) -> None:  # type: ignore[override]
        if self._menu is not None and self._menu.winfo_exists():
            self._menu.destroy()
            self._menu = None
        super().destroy()

    def _toggle_menu(self) -> None:
        if not self._values:
            return
        if self._menu is not None and self._menu.winfo_exists():
            self._close_menu()
        else:
            self._open_menu()

    def _open_menu(self) -> None:
        if not self._values:
            return
        top = ctk.CTkToplevel(self)
        top.title("Select columns")
        top.resizable(False, False)
        top.protocol("WM_DELETE_WINDOW", self._close_menu)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        top.geometry(f"+{x}+{y}")

        scroll_height = min(220, max(140, len(self._values) * 32))
        frame = ctk.CTkScrollableFrame(top, width=260, height=scroll_height)
        frame.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        for value, variable in self._variables.items():
            checkbox = ctk.CTkCheckBox(frame, text=value, variable=variable, command=self._on_check)
            checkbox.pack(anchor="w", padx=6, pady=2)

        done_button = ctk.CTkButton(top, text="Done", command=self._close_menu)
        done_button.pack(fill="x", padx=10, pady=10)

        top.grab_set()
        top.focus_force()
        self._menu = top

    def _close_menu(self) -> None:
        if self._menu is None:
            return
        if self._menu.winfo_exists():
            try:
                self._menu.grab_release()
            except Exception:
                pass
            menu = self._menu

            def _destroy() -> None:
                if menu.winfo_exists():
                    menu.destroy()

            menu.after(20, _destroy)
        self._menu = None

    def _on_check(self) -> None:
        self._update_button_text()
        if self._command is not None:
            self._command(self.get_selected())

    def _update_button_text(self) -> None:
        selected = self.get_selected()
        if not selected:
            display = self._placeholder
        else:
            if len(selected) <= 3:
                display = ", ".join(selected)
            else:
                display = ", ".join(selected[:3]) + f" +{len(selected) - 3}"
        self._button.configure(text=display)


class GeocodeApp(ctk.CTk):
    """Desktop UI wrapper around the Excel geocoding workflow."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Address Geocoder")
        self.geometry("560x420")
        self.resizable(False, False)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self._file_path_var = ctk.StringVar(value="")
        self._status_var = ctk.StringVar(value="Select an Excel file to begin.")
        self._address_column_var = ctk.StringVar(value="")
        self._color_column_var = ctk.StringVar(value="")
        self._popup_selector: MultiSelectDropdown | None = None
        self._stop_requested = False

        self._build_layout()

    def _build_layout(self) -> None:
        padding = {"padx": 20, "pady": 10}

        file_frame = ctk.CTkFrame(self)
        file_frame.pack(fill="x", **padding)

        file_label = ctk.CTkLabel(file_frame, text="Excel file:")
        file_label.pack(anchor="w")

        file_path_entry = ctk.CTkEntry(file_frame, textvariable=self._file_path_var, state="readonly")
        file_path_entry.pack(side="left", fill="x", expand=True, pady=(0, 8))

        browse_button = ctk.CTkButton(file_frame, text="Browse", command=self._on_browse)
        browse_button.pack(side="left", padx=(10, 0))

        column_frame = ctk.CTkFrame(self)
        column_frame.pack(fill="x", **padding)

        column_label = ctk.CTkLabel(column_frame, text="Address column:")
        column_label.pack(anchor="w")
        self._address_selector = ctk.CTkOptionMenu(column_frame, variable=self._address_column_var, values=["Select column"], state="disabled")
        self._address_selector.pack(fill="x")

        popup_label = ctk.CTkLabel(column_frame, text="Map popup columns:")
        popup_label.pack(anchor="w", pady=(10, 0))
        self._popup_selector = MultiSelectDropdown(column_frame, placeholder="Select columns to display")
        self._popup_selector.pack(fill="x")

        color_label = ctk.CTkLabel(column_frame, text="Color markers by column (optional):")
        color_label.pack(anchor="w", pady=(10, 0))
        self._color_selector = ctk.CTkOptionMenu(column_frame, variable=self._color_column_var, values=["No coloring"], state="disabled")
        self._color_selector.pack(fill="x")


        action_frame = ctk.CTkFrame(self)
        action_frame.pack(fill="x", **padding)

        self._start_button = ctk.CTkButton(action_frame, text="Start geocoding", command=self._on_start)
        self._start_button.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self._stop_button = ctk.CTkButton(action_frame, text="Stop", command=self._on_stop, state="disabled", fg_color="darkred", hover_color="red")
        self._stop_button.pack(side="left", fill="x", expand=True, padx=(5, 0))

        progress_frame = ctk.CTkFrame(self)
        progress_frame.pack(fill="x", **padding)

        self._progress_bar = ctk.CTkProgressBar(progress_frame)
        self._progress_bar.pack(fill="x")
        self._progress_bar.set(0)

        status_label = ctk.CTkLabel(self, textvariable=self._status_var, wraplength=500, justify="left")
        status_label.pack(fill="x", padx=20, pady=(0, 20))

    def _on_browse(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if not selected:
            return

        self._file_path_var.set(selected)
        self._status_var.set("Loading column names...")
        self._load_columns(Path(selected))

    def _load_columns(self, path: Path) -> None:
        try:
            columns = get_excel_columns(path)
        except Exception as exc:
            messagebox.showerror("Unable to read file", f"Could not read the Excel file.\n{exc}")
            self._address_selector.configure(values=["Select column"], state="disabled")
            self._address_column_var.set("")
            self._color_selector.configure(values=["No coloring"], state="disabled")
            self._color_column_var.set("")
            if self._popup_selector is not None:
                self._popup_selector.clear()
            self._status_var.set("Select a valid Excel file to continue.")
            return

        if not columns:
            messagebox.showerror("No columns found", "The Excel file does not contain any columns.")
            self._address_selector.configure(values=["Select column"], state="disabled")
            self._address_column_var.set("")
            self._color_selector.configure(values=["No coloring"], state="disabled")
            self._color_column_var.set("")
            if self._popup_selector is not None:
                self._popup_selector.clear()
            self._status_var.set("Select a valid Excel file to continue.")
            return

        inferred = infer_address_column(columns)
        self._address_selector.configure(values=columns, state="normal")
        self._address_column_var.set(inferred or columns[0])

        # Configure color column selector
        self._color_selector.configure(values=["No coloring"] + columns, state="normal")
        self._color_column_var.set("No coloring")

        if self._popup_selector is not None:
            preselected = [name for name in ("geocoded_address", "address") if name in columns]
            self._popup_selector.configure_values(columns, selected=preselected)

        self._status_var.set("Ready. Choose the address column and start geocoding.")
        self._progress_bar.set(0)

    def _on_start(self) -> None:
        file_path = self._file_path_var.get().strip()
        if not file_path:
            messagebox.showinfo("Select a file", "Please choose an Excel file before starting.")
            return

        column_name = self._address_column_var.get().strip()
        if not column_name or column_name == "Select column":
            messagebox.showinfo("Select column", "Please pick the column that contains address data.")
            return

        user_agent = DEFAULT_USER_AGENT
        popup_columns: list[str] | None = None
        if self._popup_selector is not None:
            popup_columns = self._popup_selector.get_selected()
            if not popup_columns:
                popup_columns = None

        color_column: str | None = self._color_column_var.get().strip()
        if color_column == "No coloring":
            color_column = None

        self._stop_requested = False
        self._start_button.configure(state="disabled")
        self._stop_button.configure(state="normal")
        self._status_var.set("Geocoding in progress... This may take a while for large files.")
        self._progress_bar.set(0)

        thread = threading.Thread(
            target=self._geocode_worker,
            args=(Path(file_path), column_name, user_agent, popup_columns, color_column),
            daemon=True,
        )
        thread.start()

    def _on_stop(self) -> None:
        self._stop_requested = True
        self._stop_button.configure(state="disabled")
        self._status_var.set("Stopping... Please wait for current address to finish.")

    def _check_stop(self) -> bool:
        """Return True if stop was requested."""
        return self._stop_requested

    def _geocode_worker(
        self,
        path: Path,
        column_name: str,
        user_agent: str,
        popup_columns: list[str] | None,
        color_column: str | None,
    ) -> None:
        try:
            output_path = geocode_excel_file(
                excel_path=path,
                address_column=column_name,
                user_agent=user_agent,
                progress_callback=self._on_progress_update,
                stop_check=self._check_stop,
            )
        except GeocodingCancelledError:
            self._on_cancelled()
            return
        except AddressColumnError as exc:
            self._notify_error("Column error", str(exc))
            return
        except FileNotFoundError:
            self._notify_error("File missing", "The selected Excel file could not be found.")
            return
        except Exception as exc:  # pragma: no cover - surface unexpected issues
            self._notify_error("Processing error", str(exc))
            return

        map_path = None
        map_error = None
        self.after(0, lambda: self._status_var.set("Geocoding complete. Building map..."))
        try:
            map_path = generate_map(output_path, popup_columns=popup_columns, color_column=color_column)
        except Exception as exc:
            map_error = str(exc)
        self._on_success(output_path, map_path, map_error)

    def _on_cancelled(self) -> None:
        def notify() -> None:
            self._start_button.configure(state="normal")
            self._stop_button.configure(state="disabled")
            self._status_var.set("Geocoding stopped by user.")

        self.after(0, notify)

    def _on_progress_update(self, current: int, total: int, address: str) -> None:
        def update() -> None:
            fraction = current / total if total else 0
            self._progress_bar.set(fraction)
            snippet = (address[:50] + "...") if len(address) > 50 else address
            self._status_var.set(f"Processing {current}/{total}: {snippet}")

        self.after(0, update)

    def _notify_error(self, title: str, message: str) -> None:
        def notify() -> None:
            self._start_button.configure(state="normal")
            self._stop_button.configure(state="disabled")
            self._status_var.set("Geocoding failed. Please review the message and try again.")
            messagebox.showerror(title, message)

        self.after(0, notify)

    def _on_success(self, output_path: Path, map_path: Path | None, map_error: str | None = None) -> None:
        def notify() -> None:
            self._start_button.configure(state="normal")
            self._stop_button.configure(state="disabled")
            self._progress_bar.set(1)
            status_message = f"Success! Saved to {output_path.name}."
            if map_path:
                status_message += f" Map saved to {map_path.name}."
            elif map_error:
                status_message += " Map generation failed."
            self._status_var.set(status_message)
            
            details = [f"Geocoded data saved to:\n{output_path}"]
            if map_path:
                details.append(f"Map saved to:\n{map_path}")
            elif map_error:
                trimmed = map_error if len(map_error) <= 200 else map_error[:197] + "..."
                details.append(f"Map generation failed:\n{trimmed}")
            
            # Generate shapefile
            try:
                shp_path = export_to_shapefile(output_path)
                details.append(f"Shapefile saved to:\n{shp_path}")
            except Exception as exc:
                details.append(f"Shapefile generation failed:\n{str(exc)}")
            
            messagebox.showinfo("Done", "\n\n".join(details))
            reveal_in_file_manager(output_path)

        self.after(0, notify)

def main() -> None:
    app = GeocodeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
