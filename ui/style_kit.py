import customtkinter as ctk
from tkinter import ttk, font as tkfont


def apply_global_style(root):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    _apply_tk_fonts(root)
    _apply_ttk_styles(root)


def _apply_tk_fonts(root):
    # Возвращаемся к системному шрифту mac-like вида.
    # Меняем только размер и убираем жирность.
    for font_name, size in [
        ("TkDefaultFont", 11),
        ("TkTextFont", 11),
        ("TkMenuFont", 11),
        ("TkHeadingFont", 11),
        ("TkCaptionFont", 11),
    ]:
        try:
            f = tkfont.nametofont(font_name)
            f.configure(size=size, weight="normal")
        except Exception:
            pass


def _apply_ttk_styles(root):
    try:
        style = ttk.Style(root)
        default_family = tkfont.nametofont("TkDefaultFont").cget("family")
        style.configure("Treeview", font=(default_family, 10), rowheight=24)
        style.configure("Treeview.Heading", font=(default_family, 10, "normal"))
    except Exception:
        pass


def page_shell(parent):
    shell = ctk.CTkFrame(parent, fg_color="transparent")
    shell.pack(fill="both", expand=True, padx=8, pady=8)
    return shell


def section_frame(parent):
    frame = ctk.CTkFrame(parent, corner_radius=14)
    return frame


def title_font(size=22):
    return ctk.CTkFont(size=size, weight="normal")


def text_font(size=11):
    return ctk.CTkFont(size=size, weight="normal")


def button(parent, text, command=None, width=150):
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        font=text_font(11),
        corner_radius=10,
        height=34,
    )


def home_button(parent, command):
    return button(parent, "🏠 На главную", command=command, width=140)


def label(parent, text, size=12):
    return ctk.CTkLabel(parent, text=text, font=text_font(size))


def title_label(parent, text, size=22):
    return ctk.CTkLabel(parent, text=text, font=title_font(size))
