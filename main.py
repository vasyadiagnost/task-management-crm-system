from database import init_db
from ui.main_window import MainWindow


def main():
    init_db()

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()