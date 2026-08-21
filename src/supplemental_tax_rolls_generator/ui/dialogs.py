import customtkinter as ctk


class AnimatedDialog(ctk.CTkToplevel):
    """Custom animated modal pop-up window that handles success and error feedback."""

    def __init__(self, parent, title, message, mode="success"):
        super().__init__(parent)
        self.title("")
        self.geometry("400x220")
        self.resizable(False, False)

        self.transient(parent)
        self.grab_set()

        self.icon_color = "#2ecc71" if mode == "success" else "#e74c3c"
        self.symbol = "✓" if mode == "success" else "❌"

        self.frame = ctk.CTkFrame(self, fg_color="transparent")
        self.frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.icon_lbl = ctk.CTkLabel(
            self.frame, text=self.symbol, font=ctk.CTkFont(size=1, weight="bold"),
            text_color=self.icon_color
        )
        self.icon_lbl.pack(pady=(15, 5))

        self.title_lbl = ctk.CTkLabel(
            self.frame, text=title, font=ctk.CTkFont(size=16, weight="bold")
        )
        self.title_lbl.pack(pady=5)

        self.msg_lbl = ctk.CTkLabel(
            self.frame, text=message, font=ctk.CTkFont(size=13), wraplength=350
        )
        self.msg_lbl.pack(pady=5)

        self.btn = ctk.CTkButton(
            self.frame, text="Dismiss", width=120, fg_color="#34495e",
            hover_color="#2c3e50", command=self.destroy
        )
        self.btn.pack(pady=(10, 15))

        self.current_size = 1
        self.animate_pop()

    def animate_pop(self):
        if self.current_size < 48:
            self.current_size += 4
            self.icon_lbl.configure(font=ctk.CTkFont(size=self.current_size, weight="bold"))
            self.after(10, self.animate_pop)
