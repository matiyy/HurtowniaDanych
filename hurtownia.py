import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Scrollbar, ttk
import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns


class DataWarehouseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Projekt Hurtownie Danych - Mateusz Florian")
        self.root.geometry("1400x900")
        self.df = None
        self.original_df = None  # Kopia oryginalnych danych
        self.dark_mode = False

        # Kolory dla motywów
        self.themes = {
            'light': {
                'bg': '#ffffff',
                'fg': '#000000',
                'select_bg': '#0078d4',
                'select_fg': '#ffffff',
                'button_bg': '#f0f0f0',
                'entry_bg': '#ffffff',
                'text_bg': '#ffffff'
            },
            'dark': {
                'bg': '#2d2d2d',
                'fg': '#ffffff',
                'select_bg': '#404040',
                'select_fg': '#ffffff',
                'button_bg': '#404040',
                'entry_bg': '#404040',
                'text_bg': '#1e1e1e'
            }
        }

        self.create_widgets()
        self.apply_theme()

    def create_widgets(self):
        # Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Menu Widok
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Widok", menu=view_menu)
        view_menu.add_command(label="Przełącz motyw", command=self.toggle_theme)

        # Główny frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        # Frame na przyciski
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)

        # Pierwsza linia przycisków
        tk.Button(button_frame, text="📂 Wybierz plik CSV", width=20, command=self.load_csv).grid(row=0, column=0,
                                                                                                 padx=3, pady=3)
        tk.Button(button_frame, text="📊 Pokaż statystyki", width=20, command=self.show_statistics).grid(row=0, column=1,
                                                                                                        padx=3, pady=3)
        tk.Button(button_frame, text="🔗 Oblicz korelację", width=20, command=self.calculate_correlation).grid(row=0,
                                                                                                              column=2,
                                                                                                              padx=3,
                                                                                                              pady=3)
        tk.Button(button_frame, text="📈 Wykres kolumny", width=20, command=self.plot_column).grid(row=0, column=3,
                                                                                                  padx=3, pady=3)

        # Druga linia przycisków
        tk.Button(button_frame, text="🗂️ Wyodrębnij podtabelę", width=20, command=self.extract_subtable).grid(row=1,
                                                                                                              column=0,
                                                                                                              padx=3,
                                                                                                              pady=3)
        tk.Button(button_frame, text="🔄 Zamień wartości", width=20, command=self.replace_values).grid(row=1, column=1,
                                                                                                      padx=3, pady=3)
        tk.Button(button_frame, text="💾 Zapisz do CSV", width=20, command=self.save_to_csv).grid(row=1, column=2,
                                                                                                 padx=3, pady=3)
        tk.Button(button_frame, text="↩️ Wróć do widoku danych", width=20, command=self.show_data).grid(row=1, column=3,
                                                                                                        padx=3, pady=3)

        # Trzecia linia przycisków - nowe funkcje
        tk.Button(button_frame, text="🔍 Filtruj dane", width=20, command=self.filter_data).grid(row=2, column=0, padx=3,
                                                                                                pady=3)
        tk.Button(button_frame, text="⚠️ Wykryj outliery", width=20, command=self.detect_outliers).grid(row=2, column=1,
                                                                                                        padx=3, pady=3)
        tk.Button(button_frame, text="❓ Analiza braków", width=20, command=self.analyze_missing_data).grid(row=2,
                                                                                                           column=2,
                                                                                                           padx=3,
                                                                                                           pady=3)
        tk.Button(button_frame, text="🔄 Resetuj filtr", width=20, command=self.reset_filter).grid(row=2, column=3,
                                                                                                  padx=3, pady=3)

        # Frame na dane z suwakami
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Suwak pionowy
        self.scroll_y = Scrollbar(text_frame, orient="vertical")
        self.scroll_y.pack(side="right", fill="y")

        # Suwak poziomy
        self.scroll_x = Scrollbar(text_frame, orient="horizontal")
        self.scroll_x.pack(side="bottom", fill="x")

        # Tekst z wyłączonym zawijaniem
        self.text = tk.Text(text_frame, wrap="none",
                            yscrollcommand=self.scroll_y.set,
                            xscrollcommand=self.scroll_x.set,
                            font=("Consolas", 10))
        self.text.pack(side="left", fill="both", expand=True)

        # Konfiguracja suwaków
        self.scroll_y.configure(command=self.text.yview)
        self.scroll_x.configure(command=self.text.xview)

        # Wiązanie dwukliku dla edycji komórek
        self.text.bind("<Double-Button-1>", self.on_double_click)

        # Zmienne do śledzenia struktury tabeli
        self.table_start_line = None
        self.column_positions = []
        self.row_positions = []

        # Status bar
        self.status_frame = tk.Frame(main_frame)
        self.status_frame.pack(side="bottom", fill="x", pady=(5, 0))

        self.status_label = tk.Label(self.status_frame, text="Gotowy do pracy",
                                     relief="sunken", anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)

        self.info_label = tk.Label(self.status_frame, text="",
                                   relief="sunken", anchor="e", width=30)
        self.info_label.pack(side="right")

        # Przechowywanie referencji do elementów interfejsu
        self.ui_elements = [
            button_frame, text_frame, self.text, self.scroll_y, self.scroll_x,
            self.status_frame, self.status_label, self.info_label
        ]

    def apply_theme(self):
        """Aplikuje wybrany motyw kolorystyczny"""
        theme = self.themes['dark' if self.dark_mode else 'light']

        # Główne okno
        self.root.configure(bg=theme['bg'])

        # Konfiguracja wszystkich frame'ów
        for widget in self.root.winfo_children():
            self.configure_widget_theme(widget, theme)

    def configure_widget_theme(self, widget, theme):
        """Rekurencyjnie konfiguruje motyw dla widget'a i jego dzieci"""
        try:
            widget_class = widget.winfo_class()

            if widget_class == 'Frame':
                widget.configure(bg=theme['bg'])
            elif widget_class == 'Button':
                widget.configure(bg=theme['button_bg'], fg=theme['fg'],
                                 activebackground=theme['select_bg'],
                                 activeforeground=theme['select_fg'])
            elif widget_class == 'Label':
                widget.configure(bg=theme['bg'], fg=theme['fg'])
            elif widget_class == 'Text':
                widget.configure(bg=theme['text_bg'], fg=theme['fg'],
                                 insertbackground=theme['fg'],
                                 selectbackground=theme['select_bg'],
                                 selectforeground=theme['select_fg'])
            elif widget_class == 'Scrollbar':
                widget.configure(bg=theme['button_bg'],
                                 troughcolor=theme['bg'],
                                 activebackground=theme['select_bg'])
        except tk.TclError:
            pass

        # Rekurencyjnie dla dzieci
        for child in widget.winfo_children():
            self.configure_widget_theme(child, theme)

    def toggle_theme(self):
        """Przełącza między motywem jasnym a ciemnym"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        self.update_status(f"Przełączono na motyw {'ciemny' if self.dark_mode else 'jasny'}")

    def update_status(self, message, info=""):
        """Aktualizuje status bar"""
        self.status_label.config(text=message)
        if info:
            self.info_label.config(text=info)

    def update_data_info(self):
        """Aktualizuje informacje o danych w status bar"""
        if self.df is not None:
            rows, cols = self.df.shape
            info = f"Wiersze: {rows} | Kolumny: {cols}"
            if self.original_df is not None and len(self.df) != len(self.original_df):
                info += f" | Filtrowane z {len(self.original_df)}"
            self.info_label.config(text=info)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                self.df = pd.read_csv(file_path, sep=';')
                self.original_df = self.df.copy()  # Kopia oryginału
                self.show_data()
                self.update_status("Plik CSV został pomyślnie wczytany")
                self.update_data_info()
                messagebox.showinfo("Sukces", "Plik CSV został pomyślnie wczytany.")
            except Exception as e:
                self.update_status("Błąd wczytywania pliku")
                messagebox.showerror("Błąd", f"Nie udało się wczytać pliku CSV: {e}")

    def show_data(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Brak wczytanego pliku CSV.")
            return

        self.text.delete('1.0', tk.END)
        table_content = f"Podgląd danych:\n{self.df.to_string(index=False)}\n"
        self.text.insert(tk.END, table_content)

        # Zapisz pozycje dla edycji komórek
        self.parse_table_structure()
        self.update_status("Wyświetlono dane")

    def show_statistics(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        desc = self.df.describe(include='all')
        self.text.delete('1.0', tk.END)
        self.text.insert(tk.END, f"Statystyki tabeli:\n{desc.to_string()}\n")
        self.update_status("Wyświetlono statystyki tabeli")
        messagebox.showinfo("Info", "Wyświetlono statystyki tabeli.")

    def calculate_correlation(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        numeric_df = self.df.select_dtypes(include=[np.number])
        if numeric_df.empty:
            messagebox.showerror("Błąd", "Brak danych numerycznych w zbiorze.")
            return

        correlation = numeric_df.corr(method='pearson')
        self.text.delete('1.0', tk.END)
        self.text.insert(tk.END, f"Macierz korelacji:\n{correlation.to_string()}\n")

        # Wykres korelacji
        plt.figure(figsize=(10, 8))
        sns.heatmap(correlation, annot=True, cmap='coolwarm')
        plt.title('Macierz Korelacji')
        plt.show()

        self.update_status("Wyświetlono macierz korelacji")
        messagebox.showinfo("Info", "Wyświetlono macierz korelacji i wykres.")

    def plot_column(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        column = simpledialog.askstring("Wykres kolumny", "Podaj nazwę kolumny do wykresu:")

        if column not in self.df.columns:
            messagebox.showerror("Błąd", "Podano nieprawidłową kolumnę.")
            return

        try:
            plt.figure(figsize=(10, 6))
            self.df[column].value_counts().plot(kind='bar', color='skyblue')
            plt.title(f"Wykres słupkowy dla kolumny: {column}")
            plt.xlabel(column)
            plt.ylabel("Liczba wystąpień")
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.show()

            self.update_status(f"Wyświetlono wykres dla kolumny: {column}")
            messagebox.showinfo("Sukces", f"Wyświetlono wykres dla kolumny: {column}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się utworzyć wykresu: {e}")

    def filter_data(self):
        """Filtrowanie danych według wartości w kolumnie"""
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        # Okno dialogowe do filtrowania
        filter_window = tk.Toplevel(self.root)
        filter_window.title("Filtrowanie danych")
        filter_window.geometry("400x300")

        # Wybór kolumny
        tk.Label(filter_window, text="Wybierz kolumnę:").pack(pady=5)
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(filter_window, textvariable=column_var,
                                    values=list(self.df.columns), state="readonly")
        column_combo.pack(pady=5)

        # Typ filtra
        tk.Label(filter_window, text="Typ filtra:").pack(pady=5)
        filter_type = tk.StringVar(value="równa się")
        filter_combo = ttk.Combobox(filter_window, textvariable=filter_type,
                                    values=["równa się", "zawiera", "większe niż", "mniejsze niż", "nie równa się"],
                                    state="readonly")
        filter_combo.pack(pady=5)

        # Wartość do filtrowania
        tk.Label(filter_window, text="Wartość:").pack(pady=5)
        value_var = tk.StringVar()
        value_entry = tk.Entry(filter_window, textvariable=value_var, width=30)
        value_entry.pack(pady=5)

        def apply_filter():
            column = column_var.get()
            filter_op = filter_type.get()
            value = value_var.get()

            if not column or not value:
                messagebox.showerror("Błąd", "Wypełnij wszystkie pola.")
                return

            try:
                if filter_op == "równa się":
                    filtered_df = self.original_df[self.original_df[column].astype(str) == value]
                elif filter_op == "zawiera":
                    filtered_df = self.original_df[self.original_df[column].astype(str).str.contains(value, na=False)]
                elif filter_op == "nie równa się":
                    filtered_df = self.original_df[self.original_df[column].astype(str) != value]
                elif filter_op == "większe niż":
                    filtered_df = self.original_df[
                        pd.to_numeric(self.original_df[column], errors='coerce') > float(value)]
                elif filter_op == "mniejsze niż":
                    filtered_df = self.original_df[
                        pd.to_numeric(self.original_df[column], errors='coerce') < float(value)]

                self.df = filtered_df
                self.show_data()
                self.update_data_info()
                self.update_status(f"Zastosowano filtr: {column} {filter_op} {value}")
                filter_window.destroy()

            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się zastosować filtra: {e}")

        tk.Button(filter_window, text="Zastosuj filtr", command=apply_filter).pack(pady=10)
        tk.Button(filter_window, text="Anuluj", command=filter_window.destroy).pack(pady=5)

    def reset_filter(self):
        """Resetuje filtr i przywraca oryginalne dane"""
        if self.original_df is not None:
            self.df = self.original_df.copy()
            self.show_data()
            self.update_data_info()
            self.update_status("Zresetowano filtr - przywrócono wszystkie dane")
        else:
            messagebox.showinfo("Info", "Brak danych do przywrócenia.")

    def detect_outliers(self):
        """Wykrywa outliery w danych numerycznych"""
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        numeric_columns = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_columns) == 0:
            messagebox.showerror("Błąd", "Brak kolumn numerycznych do analizy outlierów.")
            return

        outliers_info = []

        for col in numeric_columns:
            # Metoda IQR (Interquartile Range)
            Q1 = self.df[col].quantile(0.25)
            Q3 = self.df[col].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR

            outliers = self.df[(self.df[col] < lower_bound) | (self.df[col] > upper_bound)]

            if not outliers.empty:
                outliers_info.append(f"\n--- Kolumna: {col} ---")
                outliers_info.append(f"Granice: {lower_bound:.2f} - {upper_bound:.2f}")
                outliers_info.append(f"Liczba outlierów: {len(outliers)}")
                outliers_info.append(f"Wartości outlierów: {outliers[col].tolist()}")

                # Z-score method jako dodatkowa informacja
                z_scores = np.abs(stats.zscore(self.df[col].dropna()))
                z_outliers = len(z_scores[z_scores > 3])
                outliers_info.append(f"Outliery (Z-score > 3): {z_outliers}")

        if outliers_info:
            result = "ANALIZA OUTLIERÓW:\n" + "\n".join(outliers_info)
            self.text.delete('1.0', tk.END)
            self.text.insert(tk.END, result)
            self.update_status("Przeprowadzono analizę outlierów")
        else:
            self.text.delete('1.0', tk.END)
            self.text.insert(tk.END, "ANALIZA OUTLIERÓW:\n\nNie znaleziono outlierów w danych numerycznych.")
            self.update_status("Nie znaleziono outlierów")

    def analyze_missing_data(self):
        """Analizuje brakujące dane"""
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        missing_analysis = []
        missing_analysis.append("ANALIZA BRAKUJĄCYCH DANYCH:\n")

        # Ogólne statystyki
        total_cells = self.df.size
        missing_cells = self.df.isnull().sum().sum()
        missing_percentage = (missing_cells / total_cells) * 100

        missing_analysis.append(f"Całkowita liczba komórek: {total_cells}")
        missing_analysis.append(f"Brakujące komórki: {missing_cells}")
        missing_analysis.append(f"Procent brakujących danych: {missing_percentage:.2f}%\n")

        # Analiza per kolumna
        missing_analysis.append("BRAKUJĄCE DANE PER KOLUMNA:")
        missing_analysis.append("-" * 50)

        for col in self.df.columns:
            missing_count = self.df[col].isnull().sum()
            missing_pct = (missing_count / len(self.df)) * 100
            data_type = str(self.df[col].dtype)

            missing_analysis.append(f"{col}:")
            missing_analysis.append(f"  - Typ danych: {data_type}")
            missing_analysis.append(f"  - Brakujące wartości: {missing_count}")
            missing_analysis.append(f"  - Procent brakujących: {missing_pct:.2f}%")

            if missing_count > 0:
                # Indeksy wierszy z brakującymi danymi
                missing_indices = self.df[self.df[col].isnull()].index.tolist()
                missing_analysis.append(
                    f"  - Indeksy z brakami: {missing_indices[:10]}{'...' if len(missing_indices) > 10 else ''}")

            missing_analysis.append("")

        # Wiersze z największą liczbą braków
        missing_per_row = self.df.isnull().sum(axis=1)
        worst_rows = missing_per_row.nlargest(5)

        if worst_rows.max() > 0:
            missing_analysis.append("WIERSZE Z NAJWIĘKSZĄ LICZBĄ BRAKÓW:")
            missing_analysis.append("-" * 40)
            for idx, missing_count in worst_rows.items():
                if missing_count > 0:
                    missing_analysis.append(f"Wiersz {idx}: {missing_count} brakujących wartości")

        # Wzorce brakujących danych
        if missing_cells > 0:
            missing_analysis.append("\nWZORCE BRAKUJĄCYCH DANYCH:")
            missing_analysis.append("-" * 30)
            missing_combinations = self.df.isnull().value_counts().head(5)
            for pattern, count in missing_combinations.items():
                missing_cols = [col for col, is_missing in zip(self.df.columns, pattern) if is_missing]
                if missing_cols:
                    missing_analysis.append(f"Braki w kolumnach {missing_cols}: {count} wierszy")

        result = "\n".join(missing_analysis)
        self.text.delete('1.0', tk.END)
        self.text.insert(tk.END, result)
        self.update_status("Przeprowadzono analizę brakujących danych")

    def extract_subtable(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        choice = simpledialog.askstring("Podtabela",
                                        "Podaj numery wierszy (np. 0,2,4) lub nazwy kolumn (np. Kolumna1,Kolumna2):")
        if not choice:
            return

        try:
            if all(item.strip().isdigit() for item in choice.split(',')):
                rows = [int(i.strip()) for i in choice.split(',')]
                subtable = self.df.iloc[rows]
            else:
                cols = [col.strip() for col in choice.split(',')]
                for col in cols:
                    if col not in self.df.columns:
                        messagebox.showerror("Błąd", f"Kolumna '{col}' nie istnieje.")
                        return
                subtable = self.df[cols]

            self.text.delete('1.0', tk.END)
            self.text.insert(tk.END, f"Wyodrębniona podtabela:\n{subtable.to_string(index=False)}\n")
            self.update_status("Wyodrębniono podtabelę")
            messagebox.showinfo("Info", "Podtabela została wyodrębniona.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Coś poszło nie tak: {e}")

    def replace_values(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        column = simpledialog.askstring("Zamiana wartości", "Podaj nazwę kolumny:")
        if column not in self.df.columns:
            messagebox.showerror("Błąd", "Podano nieprawidłową kolumnę.")
            return

        old_value = simpledialog.askstring("Zamiana wartości", "Podaj wartość do zamiany:")
        new_value = simpledialog.askstring("Zamiana wartości", "Podaj nową wartość:")

        if old_value is None or new_value is None:
            return

        try:
            # Próba konwersji do liczby
            if pd.api.types.is_numeric_dtype(self.df[column]):
                old_value = pd.to_numeric(old_value, errors='coerce')
                new_value = pd.to_numeric(new_value, errors='coerce')

            # Zamiana wartości
            self.df[column] = self.df[column].replace(old_value, new_value)

            self.text.delete('1.0', tk.END)
            self.text.insert(tk.END, f"Zamieniono '{old_value}' na '{new_value}' w kolumnie '{column}'.\n")
            self.show_data()
            self.update_status(f"Zamieniono wartości w kolumnie {column}")
            messagebox.showinfo("Sukces", "Wartości zostały pomyślnie zamienione.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zamienić wartości: {e}")

    def save_to_csv(self):
        if self.df is None:
            messagebox.showerror("Błąd", "Najpierw wczytaj plik CSV.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if save_path:
            try:
                self.df.to_csv(save_path, index=False, sep=';')
                self.update_status(f"Zapisano plik: {save_path}")
                messagebox.showinfo("Sukces", f"Plik został zapisany jako: {save_path}")
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się zapisać pliku: {e}")

    def parse_table_structure(self):
        """Analizuje strukturę tabeli w widgecie Text dla edycji komórek"""
        content = self.text.get('1.0', tk.END)
        lines = content.split('\n')

        # Znajdź linię z nagłówkami (pierwsza linia po "Podgląd danych:")
        header_line_idx = None
        for i, line in enumerate(lines):
            if line.strip() and not line.startswith("Podgląd danych:") and line.strip() != "":
                header_line_idx = i
                break

        if header_line_idx is None:
            return

        self.table_start_line = header_line_idx + 1  # +1 bo tkinter numeruje od 1

        # Analizuj pozycje kolumn na podstawie nagłówków
        header_line = lines[header_line_idx]
        self.column_positions = []

        # Znajdź kolumny DataFrame
        if self.df is not None:
            # Używamy to_string() żeby uzyskać identyczną strukturę
            table_str = self.df.to_string(index=False)
            table_lines = table_str.split('\n')

            if len(table_lines) > 0:
                header = table_lines[0]

                # Znajdź pozycje kolumn
                current_pos = 0
                for col_name in self.df.columns:
                    col_pos = header.find(col_name, current_pos)
                    if col_pos != -1:
                        col_end = col_pos + len(str(col_name))
                        # Znajdź koniec danych kolumny (następna kolumna lub koniec linii)
                        next_col_pos = len(header)
                        for next_col in self.df.columns[list(self.df.columns).index(col_name) + 1:]:
                            next_pos = header.find(next_col, col_end)
                            if next_pos != -1:
                                next_col_pos = next_pos
                                break

                        self.column_positions.append({
                            'name': col_name,
                            'start': col_pos,
                            'end': next_col_pos,
                            'header_end': col_end
                        })
                        current_pos = col_end

    def on_double_click(self, event):
        """Obsługuje dwuklik na komórce dla edycji"""
        if self.df is None or not self.column_positions:
            return

        # Pobierz pozycję kursora
        cursor_pos = self.text.index(tk.INSERT)
        line_num = int(cursor_pos.split('.')[0])
        char_num = int(cursor_pos.split('.')[1])

        # Sprawdź czy klik był w obszarze tabeli
        if self.table_start_line is None or line_num <= self.table_start_line:
            return

        # Oblicz wiersz danych (pomijając nagłówek)
        data_row = line_num - self.table_start_line - 1

        if data_row < 0 or data_row >= len(self.df):
            return

        # Znajdź kolumnę na podstawie pozycji znaku
        selected_column = None
        col_index = None

        for i, col_info in enumerate(self.column_positions):
            if col_info['start'] <= char_num < col_info['end']:
                selected_column = col_info['name']
                col_index = i
                break

        if selected_column is None:
            return

        # Pobierz aktualną wartość
        current_value = self.df.iloc[data_row, col_index]

        # Utwórz okno edycji
        self.create_cell_editor(data_row, col_index, selected_column, current_value,
                                line_num, col_info['start'], col_info['end'])

    def create_cell_editor(self, row_idx, col_idx, col_name, current_value,
                           text_line, col_start, col_end):
        """Tworzy okno edycji dla pojedynczej komórki"""
        editor = tk.Toplevel(self.root)
        editor.title(f"Edycja komórki [{row_idx}, {col_name}]")
        editor.geometry("400x200")
        editor.transient(self.root)
        editor.grab_set()

        # Informacje o komórce
        info_frame = tk.Frame(editor)
        info_frame.pack(pady=10, padx=10, fill="x")

        tk.Label(info_frame, text=f"Wiersz: {row_idx}", font=("Arial", 9)).pack(anchor="w")
        tk.Label(info_frame, text=f"Kolumna: {col_name}", font=("Arial", 9)).pack(anchor="w")
        tk.Label(info_frame, text=f"Typ danych: {str(self.df[col_name].dtype)}", font=("Arial", 9)).pack(anchor="w")

        # Pole edycji
        tk.Label(editor, text="Nowa wartość:", font=("Arial", 10, "bold")).pack(pady=(10, 5))

        entry_var = tk.StringVar(value=str(current_value) if pd.notna(current_value) else "")
        entry = tk.Entry(editor, textvariable=entry_var, font=("Arial", 11), width=40)
        entry.pack(pady=5, padx=10)
        entry.focus_set()
        entry.select_range(0, tk.END)

        # Przyciski
        button_frame = tk.Frame(editor)
        button_frame.pack(pady=20)

        def save_changes():
            new_value = entry_var.get()

            try:
                # Próba konwersji do odpowiedniego typu
                if pd.api.types.is_numeric_dtype(self.df[col_name]):
                    if new_value.strip() == "":
                        converted_value = np.nan
                    else:
                        converted_value = pd.to_numeric(new_value)
                elif pd.api.types.is_datetime64_any_dtype(self.df[col_name]):
                    if new_value.strip() == "":
                        converted_value = pd.NaT
                    else:
                        converted_value = pd.to_datetime(new_value)
                else:
                    converted_value = new_value if new_value.strip() != "" else np.nan

                # Zapisz zmianę
                old_value = self.df.iloc[row_idx, col_idx]
                self.df.iloc[row_idx, col_idx] = converted_value

                # Odśwież widok
                self.show_data()
                self.update_data_info()

                # Aktualizuj status
                self.update_status(f"Zmieniono komórkę [{row_idx}, {col_name}]: '{old_value}' → '{converted_value}'")

                editor.destroy()

            except Exception as e:
                messagebox.showerror("Błąd konwersji",
                                     f"Nie można przekonwertować wartości '{new_value}' "
                                     f"na typ {self.df[col_name].dtype}:\n{str(e)}")

        def cancel_changes():
            editor.destroy()

        # Wiązanie Enter dla zapisania
        entry.bind('<Return>', lambda e: save_changes())
        entry.bind('<KP_Enter>', lambda e: save_changes())
        editor.bind('<Escape>', lambda e: cancel_changes())

        tk.Button(button_frame, text="💾 Zapisz", command=save_changes,
                  width=12, bg="#4CAF50", fg="white").pack(side="left", padx=5)
        tk.Button(button_frame, text="❌ Anuluj", command=cancel_changes,
                  width=12, bg="#f44336", fg="white").pack(side="left", padx=5)

        # Dodatkowe informacje
        help_frame = tk.Frame(editor)
        help_frame.pack(pady=5, padx=10, fill="x")

        help_text = "💡 Naciśnij Enter aby zapisać, Escape aby anulować"
        tk.Label(help_frame, text=help_text, font=("Arial", 8), fg="gray").pack()

        # Wyśrodkuj okno
        editor.update_idletasks()
        x = (editor.winfo_screenwidth() // 2) - (editor.winfo_width() // 2)
        y = (editor.winfo_screenheight() // 2) - (editor.winfo_height() // 2)
        editor.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataWarehouseApp(root)
    root.mainloop()