from gradebook.gradebook import Gradebook
import ttkbootstrap as ttk


def main():
    app = ttk.Window("Marks Sender", "superhero", resizable=(True, True))
    Gradebook(app)
    app.mainloop()


if __name__ == "__main__":
    main()
