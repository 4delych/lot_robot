from ui import ProcurementApp
import tkinter as tk
from tkinter import messagebox
import logging
logger = logging.getLogger(__name__)
def main():
    """Main application entry point."""
    try:
        root = tk.Tk()

        try:
            root.state("zoomed")
        except:
            try:
                root.attributes("-zoomed", True)
            except:
                pass

        app = ProcurementApp(root)

        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (900 // 2)
        y = (root.winfo_screenheight() // 2) - (700 // 2)
        root.geometry(f"900x700+{x}+{y}")

        root.mainloop()

    except Exception as e:
        logger.error(f"Application failed to start: {e}")
        messagebox.showerror(
            "Критическая ошибка", f"Не удалось запустить приложение:\n{e}"
        )


if __name__ == "__main__":
    main()