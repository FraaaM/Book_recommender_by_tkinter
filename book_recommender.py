import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import json

class BookRecommender:
    def __init__(self, root):
        self.root = root
        self.books = self.load_books()
        self.genres = sorted({book["genre"] for book in self.books})
        self.authors = sorted({author for book in self.books for author in book["author"]})
        self.selected_authors = set()
        self.tree = None

        self.genre_vars = {genre: tk.BooleanVar() for genre in self.genres}
        self.sort_option = tk.StringVar(value="alphabet")
        self.sort_order = tk.StringVar(value="asc")
        self.only_selected_genres_var = tk.BooleanVar(value=False)

        self.selected_authors_text = tk.StringVar(value="")
        self.setup_ui()

    @staticmethod
    def load_books(filename="books.json"):
        with open(filename, "r", encoding="utf-8") as file:
            return json.load(file)

    def setup_ui(self):
        self.root.title("Рекомендательная система книг")
        self.root.geometry("1200x650")
        self.root.resizable(False, False)

        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=10, pady=5, fill="both", expand=True)
        main_frame.grid_columnconfigure(1, weight=1)

        # Жанры
        self.create_genres_frame(main_frame)

        # Года
        self.year_from_entry, self.year_to_entry = self.create_years_frame(main_frame)

        # Авторы
        self.create_authors_frame(main_frame)

        # Ключевые слова
        self.keywords_entry = self.create_keywords_frame(main_frame)

        # Рекомендации
        self.results_frame = self.create_results_frame(main_frame)

        # Сортировка
        self.create_sort_frame(main_frame)

        # Кнопки
        self.create_actions_frame(main_frame)

    def create_genres_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Жанры", padx=10, pady=10)
        frame.grid(row=0, column=0, sticky="nsew")

        num_columns = 4
        for i, genre in enumerate(self.genres):
            tk.Checkbutton(frame, text=genre, variable=self.genre_vars[genre]).grid(
                row=i // num_columns, column=i % num_columns, sticky="w", padx=5, pady=2
            )

        tk.Checkbutton(
            frame, text="Рекомендовать только указанные жанры", variable=self.only_selected_genres_var
        ).grid(row=len(self.genres) // num_columns + 1, column=0, columnspan=num_columns, sticky="w")

    def create_years_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Года", padx=5, pady=5)
        frame.grid(row=1, column=0, sticky="nsew")

        tk.Label(frame, text="Начиная с:").grid(row=0, column=0, padx=(10, 3), pady=(0, 4), sticky="w")
        year_from_entry = tk.Entry(frame, width=10)
        year_from_entry.grid(row=0, column=1, padx=0, pady=(0, 4))

        tk.Label(frame, text="До:").grid(row=0, column=2, padx=(10, 3), pady=(0, 4), sticky="w")
        year_to_entry = tk.Entry(frame, width=10)
        year_to_entry.grid(row=0, column=3, padx=0, pady=(0, 4))

        return year_from_entry, year_to_entry

    def create_authors_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Авторы", padx=0, pady=0)
        frame.grid(row=2, column=0, sticky="nsew")

        tk.Label(frame, text="Выбранные авторы:").pack(anchor="w", padx=(10, 0))
        tk.Label(frame, textvariable=self.selected_authors_text, wraplength=400, anchor="w", justify="left").pack(
            fill="x", padx=10, pady=5
        )

        search_frame = tk.Frame(frame)
        search_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(search_frame, text="Введите имя автора:").pack(side="left", padx=(0, 5))
        author_search_entry = tk.Entry(search_frame)
        author_search_entry.pack(fill="x", expand=True)

        suggestions_frame = self.create_suggestions_frame(frame)
        author_search_entry.bind(
            "<KeyRelease>", lambda e: self.update_author_suggestions(author_search_entry, suggestions_frame)
        )

    def create_suggestions_frame(self, parent):
        canvas = tk.Canvas(parent, height=150)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        suggestions_frame = tk.Frame(canvas)

        canvas.create_window((0, 0), window=suggestions_frame, anchor="nw")
        canvas.config(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        suggestions_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        return suggestions_frame

    def update_author_suggestions(self, entry, suggestions_frame):
        query = entry.get().strip().lower()
        for widget in suggestions_frame.winfo_children():
            widget.destroy()

        if not query:
            return

        matching_authors = [author for author in self.authors if query in author.lower()]
        for author in matching_authors[:20]:
            tk.Button(
                suggestions_frame,
                text=author,
                anchor="w",
                relief="flat",
                bg="#f0f0f0",
                command=lambda a=author: self.select_author(a),
            ).pack(fill="x", padx=5, pady=2)

    def select_author(self, author):
        if author in self.selected_authors:
            if not messagebox.askyesno("Подтверждение", f"Удалить автора '{author}'?"):
                return
            self.selected_authors.remove(author)
        else:
            self.selected_authors.add(author)
        self.update_selected_authors()

    def update_selected_authors(self):
        self.selected_authors_text.set(", ".join(self.selected_authors))

    def create_keywords_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Ключевые слова", padx=10, pady=10)
        frame.grid(row=3, column=0, columnspan=2, sticky="nsew")
        keywords_entry = tk.Entry(frame)
        keywords_entry.pack(fill="x")
        return keywords_entry

    def create_results_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Рекомендации", padx=0, pady=0)
        frame.grid(row=0, column=1, rowspan=4, sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)

        # Treeview to show recommendations
        self.tree = ttk.Treeview(frame, columns=("Title", "Author", "Genre", "Year"), show="headings")
        self.tree.heading("Title", text="Название")
        self.tree.heading("Author", text="Автор")
        self.tree.heading("Genre", text="Жанр")
        self.tree.heading("Year", text="Год")
        self.tree.pack(fill="both", expand=True)

        return frame

    def create_sort_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Выбор сортировки", padx=10, pady=10)
        frame.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(10, 0))

        tk.Radiobutton(frame, text="По алфавиту", variable=self.sort_option, value="alphabet").grid(
            row=0, column=0, sticky="w", padx=5
        )
        tk.Radiobutton(frame, text="По году публикации", variable=self.sort_option, value="year").grid(
            row=0, column=1, sticky="w", padx=5
        )
        tk.Radiobutton(frame, text="Возрастание", variable=self.sort_order, value="asc").grid(
            row=1, column=0, sticky="w", padx=5
        )
        tk.Radiobutton(frame, text="Убывание", variable=self.sort_order, value="desc").grid(
            row=1, column=1, sticky="w", padx=5
        )

    def create_actions_frame(self, parent):
        frame = tk.Frame(parent)
        frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=(0, 1))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        tk.Button(frame, text="Сохранить в xlsx", command=self.save_to_read_list).grid(
            row=0, column=0, columnspan=2, sticky="nsew"
        )
        tk.Button(
            frame,
            text="Получить рекомендации",
            command=self.get_recommendations,
        ).grid(row=1, column=0, columnspan=2, sticky="nsew")

    def save_to_read_list(self):
        # Implement saving logic to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Recommendations"
        ws.append(["Title", "Author", "Genre", "Year"])

        for row in self.tree.get_children():
            book = self.tree.item(row)["values"]
            ws.append(book)

        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            wb.save(filename)

    def get_recommendations(self):
        # Получаем выбранные параметры фильтрации
        selected_genres = [
            genre for genre, var in self.genre_vars.items() if var.get() or (self.only_selected_genres_var.get() is False)
        ]
        selected_year_from = self.year_from_entry.get().strip()
        selected_year_to = self.year_to_entry.get().strip()

        # Фильтруем книги по жанрам, авторам и годам
        filtered_books = [book for book in self.books if
                        (not selected_genres or book["genre"] in selected_genres)
                        and (not selected_year_from or book["first_publish_year"] >= int(selected_year_from))
                        and (not selected_year_to or book["first_publish_year"] <= int(selected_year_to))
                        and (not self.selected_authors or any(author in self.selected_authors for author in book["author"]))
        ]

        # Получаем выбранную сортировку
        sort_by = self.sort_option.get()
        sort_order = self.sort_order.get()

        # Сортировка по выбранному критерию
        if sort_by == "alphabet":
            filtered_books.sort(key=lambda x: x["title"].lower(), reverse=(sort_order == "desc"))
        elif sort_by == "year":
            filtered_books.sort(key=lambda x: x["first_publish_year"], reverse=(sort_order == "desc"))

        # Очищаем старые рекомендации из таблицы
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Добавляем отсортированные книги в таблицу
        for book in filtered_books:
            self.tree.insert("", "end", values=(book["title"], ", ".join(book["author"]), book["genre"], book["first_publish_year"]))



if __name__ == "__main__":
    root = tk.Tk()
    app = BookRecommender(root)
    root.mainloop()
