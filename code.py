import tkinter as tk
from tkinter import messagebox, ttk
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd  # Добавляем библиотеку pandas для работы с Excel

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка заказов")
        self.current_user = None
        self.cart = []  # Корзина пользователя
        self.balance_var = tk.DoubleVar()  # Переменная для баланса
        self.create_database()  # Создаем базу данных
        self.create_login_window()

        # Настройка стилей для кнопок
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 12), padding=10, width=20)
        self.style.configure("Admin.TButton", background="lightblue", font=("Arial", 12), padding=10, width=20)
        self.style.configure("User.TButton", background="lightgreen", font=("Arial", 12), padding=10, width=20)

    def create_database(self):
        """Создание базы данных и таблиц."""
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()

        # Создание таблицы users
        c.execute('''CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE NOT NULL,
                        password TEXT NOT NULL,
                        balance REAL DEFAULT 0
                    )''')

        # Создание таблицы products
        c.execute('''CREATE TABLE IF NOT EXISTS products (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        price REAL NOT NULL,
                        quantity INTEGER NOT NULL
                    )''')

        # Создание таблицы orders
        c.execute('''CREATE TABLE IF NOT EXISTS orders (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER NOT NULL,
                        product_id INTEGER NOT NULL,
                        quantity INTEGER NOT NULL,
                        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (user_id) REFERENCES users(id),
                        FOREIGN KEY (product_id) REFERENCES products(id)
                    )''')

        conn.commit()
        conn.close()

    def create_login_window(self):
        self.login_frame = tk.Frame(self.root)
        self.login_frame.pack()

        # Приветствие жирным шрифтом
        welcome_label = tk.Label(self.login_frame, text="Добро пожаловать", font=("Arial", 14, "bold"))
        welcome_label.grid(row=0, column=0, columnspan=2, pady=10)

        tk.Label(self.login_frame, text="Логин:").grid(row=1, column=0)
        self.login_entry = tk.Entry(self.login_frame, font=("Arial", 12), width=20)  # Увеличиваем размер поля
        self.login_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(self.login_frame, text="Пароль:").grid(row=2, column=0)
        self.password_entry = tk.Entry(self.login_frame, show="*", font=("Arial", 12), width=20)  # Увеличиваем размер поля
        self.password_entry.grid(row=2, column=1, padx=10, pady=5)

        ttk.Button(self.login_frame, text="Войти", command=self.login, style="TButton").grid(row=3, column=0, pady=10)
        ttk.Button(self.login_frame, text="Регистрация", command=self.create_register_window, style="TButton").grid(row=3, column=1, pady=10)

    def create_register_window(self):
        self.login_frame.pack_forget()
        self.register_frame = tk.Frame(self.root)
        self.register_frame.pack()

        tk.Label(self.register_frame, text="Логин:").grid(row=0, column=0)
        self.reg_login_entry = tk.Entry(self.register_frame)
        self.reg_login_entry.grid(row=0, column=1)

        tk.Label(self.register_frame, text="Пароль:").grid(row=1, column=0)
        self.reg_password_entry = tk.Entry(self.register_frame, show="*")
        self.reg_password_entry.grid(row=1, column=1)

        ttk.Button(self.register_frame, text="Зарегистрироваться", command=self.register, style="TButton").grid(row=2, column=0)
        ttk.Button(self.register_frame, text="Назад", command=self.back_to_login, style="TButton").grid(row=2, column=1)

    def back_to_login(self):
        # Проверяем, существует ли register_frame, и если да, то скрываем его
        if hasattr(self, 'register_frame'):
            self.register_frame.pack_forget()
        # Проверяем, существует ли admin_frame, и если да, то скрываем его
        if hasattr(self, 'admin_frame'):
            self.admin_frame.pack_forget()
        # Проверяем, существует ли user_frame, и если да, то скрываем его
        if hasattr(self, 'user_frame'):
            self.user_frame.pack_forget()
        # Возвращаемся на экран входа
        self.login_frame.pack()
        self.current_user = None  # Сбрасываем текущего пользователя

    def login(self):
        username = self.login_entry.get()
        password = self.password_entry.get()
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username=? AND password=?", (username, password))
        user = c.fetchone()
        conn.close()
        if user:
            self.current_user = user
            self.balance_var.set(user[3])  # Устанавливаем начальный баланс
            if username == "admin":
                self.create_admin_window()
            else:
                self.create_user_window()
        else:
            messagebox.showerror("Ошибка", "Неверный логин или пароль")

    def register(self):
        username = self.reg_login_entry.get()
        password = self.reg_password_entry.get()
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        try:
            c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, password))
            conn.commit()
            conn.close()
            messagebox.showinfo("Успех", "Регистрация прошла успешно")
            self.back_to_login()
        except sqlite3.IntegrityError:
            messagebox.showerror("Ошибка", "Пользователь с таким логином уже существует")

    def create_admin_window(self):
        self.login_frame.pack_forget()
        self.admin_frame = tk.Frame(self.root)
        self.admin_frame.pack()

        # Приветствие с логином
        welcome_label = tk.Label(self.admin_frame, text=f"Добро пожаловать, {self.current_user[1]}", font=("Arial", 14, "bold"))
        welcome_label.grid(row=0, column=0, columnspan=2, pady=10)

        tk.Label(self.admin_frame, text="Админ-панель").grid(row=1, column=0, columnspan=2)
        ttk.Button(self.admin_frame, text="Добавить товар", command=self.add_product, style="Admin.TButton").grid(row=2, column=0, pady=10)
        ttk.Button(self.admin_frame, text="Список пользователей", command=self.view_users, style="Admin.TButton").grid(row=2, column=1, pady=10)
        ttk.Button(self.admin_frame, text="Список заказов", command=self.view_orders, style="Admin.TButton").grid(row=3, column=0, pady=10)
        ttk.Button(self.admin_frame, text="Удалить товар", command=self.delete_product, style="Admin.TButton").grid(row=3, column=1, pady=10)
        ttk.Button(self.admin_frame, text="Удалить пользователя", command=self.delete_user, style="Admin.TButton").grid(row=4, column=0, pady=10)
        ttk.Button(self.admin_frame, text="Список товаров", command=self.view_products, style="Admin.TButton").grid(row=4, column=1, pady=10)
        ttk.Button(self.admin_frame, text="Статистика", command=self.view_statistics, style="Admin.TButton").grid(row=5, column=0, pady=10)
        ttk.Button(self.admin_frame, text="Сменить пользователя", command=self.back_to_login, style="Admin.TButton").grid(row=5, column=1, pady=10)

        # Добавляем кнопку для экспорта в Excel
        ttk.Button(self.admin_frame, text="Экспорт заказов в Excel", command=self.export_to_excel, style="Admin.TButton").grid(row=6, column=0, pady=10)

    def add_product(self):
        self.add_product_window = tk.Toplevel(self.root)
        tk.Label(self.add_product_window, text="Название:").grid(row=0, column=0)
        name_entry = tk.Entry(self.add_product_window)
        name_entry.grid(row=0, column=1)

        tk.Label(self.add_product_window, text="Цена:").grid(row=1, column=0)
        price_entry = tk.Entry(self.add_product_window)
        price_entry.grid(row=1, column=1)

        tk.Label(self.add_product_window, text="Количество:").grid(row=2, column=0)
        quantity_entry = tk.Entry(self.add_product_window)
        quantity_entry.grid(row=2, column=1)

        ttk.Button(self.add_product_window, text="Добавить", command=lambda: self.save_product(name_entry.get(), price_entry.get(), quantity_entry.get()), style="TButton").grid(row=3, column=0)

    def save_product(self, name, price, quantity):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("INSERT INTO products (name, price, quantity) VALUES (?, ?, ?)", (name, price, quantity))
        conn.commit()
        conn.close()
        self.add_product_window.destroy()

    def view_users(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM users")
        users = c.fetchall()
        conn.close()

        self.view_users_window = tk.Toplevel(self.root)
        for i, user in enumerate(users):
            tk.Label(self.view_users_window, text=f"ID: {user[0]}, Логин: {user[1]}, Баланс: {user[3]}").grid(row=i, column=0)

    def view_orders(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT orders.id, users.username, products.name, orders.quantity, orders.created_at FROM orders JOIN users ON orders.user_id = users.id JOIN products ON orders.product_id = products.id")
        orders = c.fetchall()
        conn.close()

        self.view_orders_window = tk.Toplevel(self.root)
        for i, order in enumerate(orders):
            tk.Label(self.view_orders_window, text=f"ID заказа: {order[0]}, Пользователь: {order[1]}, Товар: {order[2]}, Количество: {order[3]}, Дата: {order[4]}").grid(row=i, column=0)

    def delete_product(self):
        self.delete_product_window = tk.Toplevel(self.root)
        tk.Label(self.delete_product_window, text="ID товара:").grid(row=0, column=0)
        product_id_entry = tk.Entry(self.delete_product_window)
        product_id_entry.grid(row=0, column=1)

        ttk.Button(self.delete_product_window, text="Удалить", command=lambda: self.remove_product(product_id_entry.get()), style="TButton").grid(row=1, column=0)

    def remove_product(self, product_id):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("DELETE FROM products WHERE id=?", (product_id,))
        conn.commit()
        conn.close()
        self.delete_product_window.destroy()

    def delete_user(self):
        self.delete_user_window = tk.Toplevel(self.root)
        tk.Label(self.delete_user_window, text="ID пользователя:").grid(row=0, column=0)
        user_id_entry = tk.Entry(self.delete_user_window)
        user_id_entry.grid(row=0, column=1)

        ttk.Button(self.delete_user_window, text="Удалить", command=lambda: self.remove_user(user_id_entry.get()), style="TButton").grid(row=1, column=0)

    def remove_user(self, user_id):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("DELETE FROM users WHERE id=?", (user_id,))
        conn.commit()
        conn.close()
        self.delete_user_window.destroy()

    def view_products(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM products")
        products = c.fetchall()
        conn.close()

        self.view_products_window = tk.Toplevel(self.root)
        self.view_products_window.minsize(600, 400)  # Устанавливаем минимальный размер окна

        # Создаем Listbox с прокруткой
        self.product_listbox = tk.Listbox(self.view_products_window, font=("Arial", 12), width=80, height=20)
        self.product_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Добавляем вертикальную прокрутку
        scrollbar = tk.Scrollbar(self.view_products_window, command=self.product_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.product_listbox.config(yscrollcommand=scrollbar.set)

        for product in products:
            self.product_listbox.insert(tk.END, f"ID: {product[0]}, Название: {product[1]}, Цена: {product[2]}, Количество: {product[3]}")

    def create_user_window(self):
        self.login_frame.pack_forget()
        self.user_frame = tk.Frame(self.root)
        self.user_frame.pack()

        # Приветствие с логином
        welcome_label = tk.Label(self.user_frame, text=f"Добро пожаловать, {self.current_user[1]}", font=("Arial", 14, "bold"))
        welcome_label.grid(row=0, column=0, columnspan=2, pady=10)

        tk.Label(self.user_frame, text="Список товаров").grid(row=1, column=0)
        self.product_listbox = tk.Listbox(self.user_frame, font=("Arial", 12), width=80, height=20)
        self.product_listbox.grid(row=2, column=0)
        self.load_products()

        ttk.Button(self.user_frame, text="Добавить в корзину", command=self.add_to_cart, style="User.TButton").grid(row=3, column=0, pady=10)
        ttk.Button(self.user_frame, text="Корзина", command=self.view_cart, style="User.TButton").grid(row=4, column=0, pady=10)
        ttk.Button(self.user_frame, text="Пополнить баланс", command=self.top_up_balance, style="User.TButton").grid(row=5, column=0, pady=10)
        ttk.Button(self.user_frame, text="История заказов", command=self.view_order_history, style="User.TButton").grid(row=6, column=0, pady=10)
        ttk.Button(self.user_frame, text="Поиск товара", command=self.search_products, style="User.TButton").grid(row=7, column=0, pady=10)
        ttk.Button(self.user_frame, text="Сменить пользователя", command=self.back_to_login, style="User.TButton").grid(row=8, column=0, pady=10)

        # Показ баланса пользователя
        tk.Label(self.user_frame, text="Баланс:").grid(row=9, column=0, sticky="w")  # Выравниваем по левому краю
        tk.Label(self.user_frame, textvariable=self.balance_var).grid(row=9, column=1, sticky="e")  # Выравниваем по правому краю

    def load_products(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM products")
        products = c.fetchall()
        conn.close()

        self.product_listbox.delete(0, tk.END)
        for product in products:
            self.product_listbox.insert(tk.END, f"ID: {product[0]}, Название: {product[1]}, Цена: {product[2]}, Количество: {product[3]}")

    def add_to_cart(self):
        self.add_to_cart_window = tk.Toplevel(self.root)
        tk.Label(self.add_to_cart_window, text="ID товара:").grid(row=0, column=0)
        product_id_entry = tk.Entry(self.add_to_cart_window)
        product_id_entry.grid(row=0, column=1)

        tk.Label(self.add_to_cart_window, text="Количество:").grid(row=1, column=0)
        quantity_entry = tk.Entry(self.add_to_cart_window)
        quantity_entry.grid(row=1, column=1)

        ttk.Button(self.add_to_cart_window, text="Добавить", command=lambda: self.add_product_to_cart(product_id_entry.get(), quantity_entry.get()), style="TButton").grid(row=2, column=0)

    def add_product_to_cart(self, product_id, quantity):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM products WHERE id=?", (product_id,))
        product = c.fetchone()
        conn.close()

        if product:
            if int(quantity) > product[3]:
                messagebox.showerror("Ошибка", "Недостаточно товара на складе")
            else:
                for _ in range(int(quantity)):
                    self.cart.append(product)
                messagebox.showinfo("Успех", f"{quantity} шт. товара добавлено в корзину")
                self.add_to_cart_window.destroy()
        else:
            messagebox.showerror("Ошибка", "Товар не найден")

    def view_cart(self):
        self.cart_window = tk.Toplevel(self.root)
        total_cost = sum(item[2] for item in self.cart)
        tk.Label(self.cart_window, text=f"Общая сумма: {total_cost}").grid(row=0, column=0, columnspan=2)

        # Используем Treeview для отображения корзины
        self.cart_tree = ttk.Treeview(self.cart_window, columns=("ID", "Название", "Цена", "Количество"), show="headings")
        self.cart_tree.heading("ID", text="ID")
        self.cart_tree.heading("Название", text="Название")
        self.cart_tree.heading("Цена", text="Цена")
        self.cart_tree.heading("Количество", text="Количество")
        self.cart_tree.grid(row=1, column=0, columnspan=2)

        for item in self.cart:
            self.cart_tree.insert("", tk.END, values=(item[0], item[1], item[2], item[3]))

        ttk.Button(self.cart_window, text="Оплатить", command=self.checkout, style="TButton").grid(row=2, column=0)
        ttk.Button(self.cart_window, text="Удалить товар", command=self.remove_from_cart, style="TButton").grid(row=2, column=1)
        ttk.Button(self.cart_window, text="Очистить корзину", command=self.clear_cart, style="TButton").grid(row=3, column=0)

    def remove_from_cart(self):
        self.remove_from_cart_window = tk.Toplevel(self.root)
        tk.Label(self.remove_from_cart_window, text="ID товара:").grid(row=0, column=0)
        product_id_entry = tk.Entry(self.remove_from_cart_window)
        product_id_entry.grid(row=0, column=1)

        tk.Label(self.remove_from_cart_window, text="Количество:").grid(row=1, column=0)
        quantity_entry = tk.Entry(self.remove_from_cart_window)
        quantity_entry.grid(row=1, column=1)

        ttk.Button(self.remove_from_cart_window, text="Удалить", command=lambda: self.remove_product_from_cart(product_id_entry.get(), quantity_entry.get()), style="TButton").grid(row=2, column=0)

    def remove_product_from_cart(self, product_id, quantity):
        try:
            quantity = int(quantity)
            product_id = int(product_id)
            count = 0
            for item in self.cart:
                if item[0] == product_id:
                    if count < quantity:
                        self.cart.remove(item)
                        count += 1
            messagebox.showinfo("Успех", f"{quantity} шт. товара удалено из корзины")
            self.remove_from_cart_window.destroy()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверное количество")

    def clear_cart(self):
        self.cart = []
        messagebox.showinfo("Успех", "Корзина очищена")

    def checkout(self):
        total_cost = sum(item[2] for item in self.cart)
        if self.current_user[3] >= total_cost:
            conn = sqlite3.connect('orders.db')
            c = conn.cursor()
            for item in self.cart:
                c.execute("UPDATE products SET quantity=quantity-1 WHERE id=?", (item[0],))
                if item[3] == 1:
                    c.execute("DELETE FROM products WHERE id=?", (item[0],))
                c.execute("INSERT INTO orders (user_id, product_id, quantity, created_at) VALUES (?, ?, ?, ?)", (self.current_user[0], item[0], 1, datetime.now()))
            c.execute("UPDATE users SET balance=balance-? WHERE id=?", (total_cost, self.current_user[0]))
            conn.commit()
            conn.close()
            self.cart = []
            messagebox.showinfo("Успех", "Оплата прошла успешно")
            self.cart_window.destroy()
            self.update_user_balance()  # Обновляем баланс пользователя
        else:
            messagebox.showerror("Ошибка", "Недостаточно средств на балансе")

    def top_up_balance(self):
        self.top_up_window = tk.Toplevel(self.root)
        tk.Label(self.top_up_window, text="Сумма пополнения:").grid(row=0, column=0)
        amount_entry = tk.Entry(self.top_up_window)
        amount_entry.grid(row=0, column=1)

        ttk.Button(self.top_up_window, text="Пополнить", command=lambda: self.update_balance(amount_entry.get()), style="TButton").grid(row=1, column=0)

    def update_balance(self, amount):
        try:
            amount = float(amount)
            if amount <= 0:
                messagebox.showerror("Ошибка", "Сумма пополнения должна быть положительной")
                return

            conn = sqlite3.connect('orders.db')
            c = conn.cursor()
            c.execute("UPDATE users SET balance=balance+? WHERE id=?", (amount, self.current_user[0]))
            conn.commit()
            conn.close()
            self.top_up_window.destroy()
            messagebox.showinfo("Успех", f"Баланс пополнен на {amount}")
            self.update_user_balance()  # Обновляем баланс пользователя
        except ValueError:
            messagebox.showerror("Ошибка", "Неверная сумма. Введите число.")

    def update_user_balance(self):
        """Обновляет текущий баланс пользователя и переменную balance_var."""
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT balance FROM users WHERE id=?", (self.current_user[0],))
        new_balance = c.fetchone()[0]
        conn.close()
        self.current_user = (self.current_user[0], self.current_user[1], self.current_user[2], new_balance)
        self.balance_var.set(new_balance)

    def get_user_data(self, user_id):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE id=?", (user_id,))
        user = c.fetchone()
        conn.close()
        return user

    def view_order_history(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT orders.id, products.name, orders.quantity, orders.created_at FROM orders JOIN products ON orders.product_id = products.id WHERE user_id=?", (self.current_user[0],))
        orders = c.fetchall()
        conn.close()

        self.order_history_window = tk.Toplevel(self.root)
        for i, order in enumerate(orders):
            tk.Label(self.order_history_window, text=f"Заказ №{order[0]}, Товар: {order[1]}, Количество: {order[2]}, Дата: {order[3]}").grid(row=i, column=0)

    def search_products(self):
        self.search_window = tk.Toplevel(self.root)
        tk.Label(self.search_window, text="Поиск товара:").grid(row=0, column=0)
        search_entry = tk.Entry(self.search_window)
        search_entry.grid(row=0, column=1)
        ttk.Button(self.search_window, text="Найти", command=lambda: self.show_search_results(search_entry.get()), style="TButton").grid(row=1, column=0)

    def show_search_results(self, query):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("SELECT * FROM products WHERE name LIKE ?", (f"%{query}%",))
        results = c.fetchall()
        conn.close()

        self.search_results_window = tk.Toplevel(self.root)
        for i, product in enumerate(results):
            tk.Label(self.search_results_window, text=f"ID: {product[0]}, Название: {product[1]}, Цена: {product[2]}, Количество: {product[3]}").grid(row=i, column=0)

    def view_statistics(self):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        c.execute("""
            SELECT products.name, SUM(orders.quantity) as total_sales
            FROM orders
            JOIN products ON orders.product_id = products.id
            GROUP BY products.name
            ORDER BY total_sales DESC
            LIMIT 5
        """)
        top_products = c.fetchall()
        conn.close()

        self.statistics_window = tk.Toplevel(self.root)
        self.statistics_window.minsize(800, 600)  # Увеличиваем размер окна
        self.statistics_window.resizable(True, True)  # Разрешаем изменение размера окна

        # Создание графика
        fig, ax = plt.subplots(figsize=(8, 6))
        products, sales = zip(*top_products)
        ax.barh(products, sales, color='skyblue')
        ax.set_xlabel('Количество продаж')
        ax.set_title('Топ-5 товаров по продажам')

        # Встраивание графика в окно
        canvas = FigureCanvasTkAgg(fig, master=self.statistics_window)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Добавляем текст с данными
        text_frame = tk.Frame(self.statistics_window)
        text_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        tk.Label(text_frame, text="Топ-5 товаров по продажам:", font=("Arial", 14, "bold")).pack()
        for product, sales in top_products:
            tk.Label(text_frame, text=f"{product}: {sales} продаж").pack()

    def filter_products(self, sort_by="price"):
        conn = sqlite3.connect('orders.db')
        c = conn.cursor()
        if sort_by == "price":
            c.execute("SELECT * FROM products ORDER BY price ASC")
        elif sort_by == "quantity":
            c.execute("SELECT * FROM products ORDER BY quantity DESC")
        products = c.fetchall()
        conn.close()

        self.product_listbox.delete(0, tk.END)
        for product in products:
            self.product_listbox.insert(tk.END, f"ID: {product[0]}, Название: {product[1]}, Цена: {product[2]}, Количество: {product[3]}")

    def export_to_excel(self):
        """Экспорт данных из таблицы orders в Excel."""
        conn = sqlite3.connect('orders.db')
        query = "SELECT * FROM orders"
        df = pd.read_sql_query(query, conn)
        conn.close()

        # Сохранение в Excel
        file_path = "orders_export.xlsx"
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Успех", f"Данные успешно экспортированы в файл {file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()