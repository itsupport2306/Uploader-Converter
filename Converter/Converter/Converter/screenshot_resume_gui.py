from __future__ import annotations

import queue
import threading
import traceback
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk

from screenshot_to_word import (
    IMAGE_EXTS,
    configure_tesseract,
    convert,
)


class ScreenshotResumeApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("Screenshot to Resume Converter")
        self.root.geometry("760x520")
        self.root.minsize(680, 460)

        self.input_dir = StringVar()
        self.output_dir = StringVar()
        self.tesseract_path = StringVar()
        self.force = BooleanVar(value=False)
        self.running = False
        self.log_queue: queue.Queue[str] = queue.Queue()

        self._build_ui()
        self._poll_log()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(6, weight=1)

        title = ttk.Label(
            outer,
            text="Screenshot to Resume Converter",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(outer, text="Input screenshot folder").grid(
            row=1, column=0, sticky="w", pady=6
        )
        ttk.Entry(outer, textvariable=self.input_dir).grid(
            row=1, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(outer, text="Browse", command=self._browse_input).grid(
            row=1, column=2, pady=6
        )

        ttk.Label(outer, text="Output resume folder").grid(
            row=2, column=0, sticky="w", pady=6
        )
        ttk.Entry(outer, textvariable=self.output_dir).grid(
            row=2, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(outer, text="Browse", command=self._browse_output).grid(
            row=2, column=2, pady=6
        )

        ttk.Label(outer, text="Tesseract.exe (optional)").grid(
            row=3, column=0, sticky="w", pady=6
        )
        ttk.Entry(outer, textvariable=self.tesseract_path).grid(
            row=3, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(outer, text="Browse", command=self._browse_tesseract).grid(
            row=3, column=2, pady=6
        )

        ttk.Checkbutton(
            outer,
            text="Overwrite existing .docx files",
            variable=self.force,
        ).grid(row=4, column=1, sticky="w", padx=8, pady=(6, 10))

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 10))

        log_frame = ttk.LabelFrame(outer, text="Log")
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log = ttk.Treeview(log_frame, columns=("message",), show="headings")
        self.log.heading("message", text="Status")
        self.log.column("message", anchor="w", stretch=True)
        self.log.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

        buttons = ttk.Frame(outer)
        buttons.grid(row=7, column=0, columnspan=3, sticky="e", pady=(14, 0))
        self.run_button = ttk.Button(buttons, text="Convert", command=self._start)
        self.run_button.pack(side="right")

    def _browse_input(self) -> None:
        path = filedialog.askdirectory(title="Select screenshot input folder")
        if path:
            self.input_dir.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Select output resume folder")
        if path:
            self.output_dir.set(path)

    def _browse_tesseract(self) -> None:
        path = filedialog.askopenfilename(
            title="Select tesseract.exe",
            filetypes=[("Tesseract executable", "tesseract.exe"), ("Executables", "*.exe")],
        )
        if path:
            self.tesseract_path.set(path)

    def _start(self) -> None:
        if self.running:
            return

        input_text = self.input_dir.get().strip()
        output_text = self.output_dir.get().strip()
        input_dir = Path(input_text)
        output_dir = Path(output_text)
        tesseract = self.tesseract_path.get().strip() or None

        if not input_text or not input_dir.is_dir():
            messagebox.showerror("Missing input", "Please choose a valid input folder.")
            return
        if not output_text:
            messagebox.showerror("Missing output", "Please choose an output folder.")
            return

        images = [p for p in sorted(input_dir.iterdir()) if p.suffix.lower() in IMAGE_EXTS]
        if not images:
            messagebox.showerror("No screenshots", "No supported image files were found.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.progress.configure(value=0, maximum=len(images))
        self.log.delete(*self.log.get_children())
        self.running = True
        self.run_button.configure(state="disabled")

        worker = threading.Thread(
            target=self._run_conversion,
            args=(images, output_dir, tesseract),
            daemon=True,
        )
        worker.start()

    def _run_conversion(
        self,
        images: list[Path],
        output_dir: Path,
        tesseract: str | None,
    ) -> None:
        done = skipped = failed = 0
        try:
            configure_tesseract(tesseract)
            for image_path in images:
                out_path = output_dir / f"{image_path.stem}.docx"
                if out_path.exists() and not self.force.get():
                    skipped += 1
                    self._log(f"Skip  {out_path.name} (already exists)")
                    self._advance()
                    continue

                try:
                    convert(image_path, out_path)
                    done += 1
                    self._log(f"Wrote {out_path}")
                except Exception as exc:
                    failed += 1
                    self._log(f"FAIL  {image_path.name}: {exc}")
                finally:
                    self._advance()
        except SystemExit as exc:
            failed += 1
            self._log(str(exc))
        except Exception:
            failed += 1
            self._log(traceback.format_exc())
        finally:
            self._log(
                f"Done. {done} converted, {skipped} skipped, {failed} failed "
                f"({len(images)} total)."
            )
            self.log_queue.put("__DONE__")

    def _log(self, message: str) -> None:
        self.log_queue.put(message)

    def _advance(self) -> None:
        self.log_queue.put("__ADVANCE__")

    def _poll_log(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                if message == "__ADVANCE__":
                    self.progress.step(1)
                elif message == "__DONE__":
                    self.running = False
                    self.run_button.configure(state="normal")
                    messagebox.showinfo("Finished", "Conversion finished.")
                else:
                    self.log.insert("", "end", values=(message,))
                    children = self.log.get_children()
                    if children:
                        self.log.see(children[-1])
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log)


def main() -> None:
    root = Tk()
    try:
        root.call("tk", "scaling", 1.25)
    except Exception:
        pass
    ScreenshotResumeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
