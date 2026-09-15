import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkinter.filedialog import asksaveasfilename
from tkinter import ttk
from datetime import date
import tempfile
import os

from tkcalendar import Calendar
from tkcalendar import DateEntry
#database
from real.dblogin import Registerdata
from real.dbWarehouse import warehouse
from real.dbTotal_TempReal import DatabaseTotal
from real.dbSalesReportReal import DatabaseSales
import random
import customtkinter as ctk
from tkinter import ttk, messagebox
import sqlite3
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime



# ============================================================
# CUSTOMTKINTER SETTINGS
# ============================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# ============================================================
# FIRST PAGE - LOGIN
# ============================================================

class FirstPage(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#f3f5f9")

        self.controller = controller

        # ----------------------------------------------------
        # Background
        # ----------------------------------------------------

        self.bg_frame = ctk.CTkFrame(
            self,
            fg_color="#f3f5f9"
        )
        self.bg_frame.pack(fill="both", expand=True)

        # ----------------------------------------------------
        # Main PanedWindow
        # ----------------------------------------------------

        self.paned = tk.PanedWindow(
            self.bg_frame,
            orient=tk.HORIZONTAL,
            sashwidth=0,
            bd=0,
            relief="flat",
            bg="#f3f5f9"
        )

        self.paned.pack(
            fill="both",
            expand=True,
            padx=40,
            pady=40
        )

        # ----------------------------------------------------
        # LEFT PANEL
        # ----------------------------------------------------

        self.left_panel = ctk.CTkFrame(
            self.paned,
            fg_color="#1e293b",
            corner_radius=25
        )

        self.paned.add(
            self.left_panel,
            minsize=400
        )

        # Logo / title area

        self.logo_title = ctk.CTkLabel(
            self.left_panel,
            text="ABK",
            font=ctk.CTkFont(
                family="Arial",
                size=70,
                weight="bold"
            ),
            text_color="#ffffff"
        )

        self.logo_title.pack(
            pady=(120, 10)
        )

        self.logo_line = ctk.CTkFrame(
            self.left_panel,
            width=100,
            height=5,
            fg_color="#fbbf24",
            corner_radius=5
        )

        self.logo_line.pack(
            pady=10
        )

        self.system_title = ctk.CTkLabel(
            self.left_panel,
            text="POS APPLICATION",
            font=ctk.CTkFont(
                family="Arial",
                size=25,
                weight="bold"
            ),
            text_color="#ffffff"
        )

        self.system_title.pack(
            pady=(20, 5)
        )

        self.system_subtitle = ctk.CTkLabel(
            self.left_panel,
            text="Point of Sale Management System",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            ),
            text_color="#cbd5e1"
        )

        self.system_subtitle.pack(
            pady=5
        )

        self.info_text = ctk.CTkLabel(
            self.left_panel,
            text="Secure • Fast • Simple",
            font=ctk.CTkFont(
                family="Arial",
                size=14
            ),
            text_color="#fbbf24"
        )

        self.info_text.pack(
            pady=30
        )

        # ----------------------------------------------------
        # RIGHT PANEL
        # ----------------------------------------------------

        self.right_panel = ctk.CTkFrame(
            self.paned,
            fg_color="#ffffff",
            corner_radius=25
        )

        self.paned.add(
            self.right_panel,
            minsize=650
        )

        # Login title

        self.login_title = ctk.CTkLabel(
            self.right_panel,
            text="Welcome Back",
            font=ctk.CTkFont(
                family="Arial",
                size=38,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.login_title.pack(
            pady=(90, 5)
        )

        self.login_subtitle = ctk.CTkLabel(
            self.right_panel,
            text="Please login to continue",
            font=ctk.CTkFont(
                family="Arial",
                size=16
            ),
            text_color="#64748b"
        )

        self.login_subtitle.pack(
            pady=(0, 35)
        )

        # ----------------------------------------------------
        # LOGIN FORM
        # ----------------------------------------------------

        self.form_frame = ctk.CTkFrame(
            self.right_panel,
            fg_color="transparent"
        )

        self.form_frame.pack(
            padx=80,
            fill="x"
        )

        # Username

        self.username_label = ctk.CTkLabel(
            self.form_frame,
            text="Username",
            anchor="w",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155"
        )

        self.username_label.pack(
            fill="x",
            pady=(0, 8)
        )

        text1 = tk.StringVar()
        text2 = tk.StringVar()

        self.T1 = ctk.CTkEntry(
            self.form_frame,
            height=50,
            corner_radius=10,
            border_width=1,
            border_color="#cbd5e1",
            fg_color="#f8fafc",
            text_color="#111827",
            font=ctk.CTkFont(
                family="Arial",
                size=16
            ),
            textvariable=text1,
            placeholder_text="Enter username"
        )

        self.T1.pack(
            fill="x",
            pady=(0, 20)
        )

        # Password

        self.password_label = ctk.CTkLabel(
            self.form_frame,
            text="Password",
            anchor="w",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155"
        )

        self.password_label.pack(
            fill="x",
            pady=(0, 8)
        )

        self.T2 = ctk.CTkEntry(
            self.form_frame,
            height=50,
            corner_radius=10,
            border_width=1,
            border_color="#cbd5e1",
            fg_color="#f8fafc",
            text_color="#111827",
            font=ctk.CTkFont(
                family="Arial",
                size=16
            ),
            textvariable=text2,
            placeholder_text="Enter password",
            show="*"
        )

        self.T2.pack(
            fill="x",
            pady=(0, 30)
        )

        # ----------------------------------------------------
        # DATABASE
        # ----------------------------------------------------

        dbRegister = Registerdata(
            "real/Registerdatabase.db"
        )

        register = []
        register = dbRegister.fetch()

        registerUser = dbRegister.fetchUser()

        # ----------------------------------------------------
        # ADMIN VERIFY
        # ----------------------------------------------------

        def verify():

            if self.T1.get() == "admin" and self.T2.get() == "12345@":
                controller.current_username = self.T1.get()
                text1.set('')
                text2.set('')

                controller.show_frame(SecondPage)

            else:

                try:

                    i = 0

                    for a in range(len(register)):

                        if (
                            register[a][1] == self.T1.get()
                            and
                            register[a][2] == self.T2.get()
                        ):
                            controller.current_username = register[a][1]
                            controller.show_frame(SecondPage)

                            i = 1
                            break

                    if i == 0:

                        messagebox.showinfo(
                            "Error",
                            "Please provide correct username and password!!"
                        )

                except:

                    messagebox.showinfo(
                        "Error",
                        "Please provide correct username and password!!"
                    )

        # ----------------------------------------------------
        # USER VERIFY
        # ----------------------------------------------------

        def Userverify():

            if self.T1.get() == "admin" and self.T2.get() == "12345@":
                controller.current_username = self.T1.get()
                text1.set('')
                text2.set('')

                controller.show_frame(UserBillPage)

            else:

                try:

                    i = 0

                    for a in range(len(register)):

                        if (
                            registerUser[a][1] == self.T1.get()
                            and
                            registerUser[a][2] == self.T2.get()
                        ):
                            controller.current_username = registerUser[a][1]
                            controller.show_frame(UserBillPage)

                            i = 1
                            break

                    if i == 0:

                        messagebox.showinfo(
                            "Error",
                            "Please provide correct username and password!!"
                        )

                except:

                    messagebox.showinfo(
                        "Error",
                        "Please provide correct username and password!!"
                    )

        # ----------------------------------------------------
        # BUTTONS
        # ----------------------------------------------------

        self.button_frame = ctk.CTkFrame(
            self.form_frame,
            fg_color="transparent"
        )

        self.button_frame.pack(
            fill="x"
        )

        self.admin_button = ctk.CTkButton(
            self.button_frame,
            text="Admin",
            height=48,
            corner_radius=10,
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=verify
        )

        self.admin_button.pack(
            side="left",
            expand=True,
            fill="x",
            padx=(0, 7)
        )

        self.user_button = ctk.CTkButton(
            self.button_frame,
            text="User",
            height=48,
            corner_radius=10,
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            fg_color="#f59e0b",
            hover_color="#d97706",
            command=Userverify
        )

        self.user_button.pack(
            side="left",
            expand=True,
            fill="x",
            padx=7
        )

        self.exit_button = ctk.CTkButton(
            self.button_frame,
            text="Exit",
            height=48,
            corner_radius=10,
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            fg_color="#10b981",
            hover_color="#059669",
            command=self.controller.destroy
        )

        self.exit_button.pack(
            side="left",
            expand=True,
            fill="x",
            padx=(7, 0)
        )

        # ----------------------------------------------------
        # FOOTER
        # ----------------------------------------------------

        self.footer = ctk.CTkLabel(
            self.right_panel,
            text="ABK POS System created by Magway Software House © 2026, Contact - 09777275950",
            font=ctk.CTkFont(
                family="Arial",
                size=12
            ),
            text_color="#94a3b8"
        )

        self.footer.pack(
            side="bottom",
            pady=25
        )



# ============================================================
# SECOND PAGE - POS DASHBOARD
# ============================================================

class SecondPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#d2daea"
        )

        self.controller = controller

        # ============================================================
        # DATABASE
        # ============================================================

        self.dbSales = DatabaseSales(
            "real/SalesReportdatabase.db"
        )

        self.dbWarehouse = warehouse(
            "real/Warehousedatabase.db"
        )

        # ============================================================
        # MAIN LAYOUT
        # ============================================================

        self.grid_columnconfigure(
            0,
            weight=0
        )

        self.grid_columnconfigure(
            1,
            weight=1
        )

        self.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # LEFT SIDEBAR
        # ============================================================

        self.sidebar = ctk.CTkFrame(
            self,
            width=220,
            corner_radius=0,
            fg_color="#111827"
        )

        self.sidebar.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        self.sidebar.grid_propagate(False)

        # ============================================================
        # LOGO / TITLE
        # ============================================================

        logo_frame = ctk.CTkFrame(
            self.sidebar,
            fg_color="transparent"
        )

        logo_frame.pack(
            fill="x",
            padx=15,
            pady=(25, 20)
        )

        ctk.CTkLabel(
            logo_frame,
            text="ABK",
            font=ctk.CTkFont(
                family="Arial",
                size=30,
                weight="bold"
            ),
            text_color="#ffffff"
        ).pack()

        ctk.CTkLabel(
            logo_frame,
            text="MiniMart POS",
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#9ca3af"
        ).pack(
            pady=(0, 5)
        )

        # ============================================================
        # NAVIGATION
        # ============================================================

        nav_frame = ctk.CTkFrame(
            self.sidebar,
            fg_color="transparent"
        )

        nav_frame.pack(
            fill="both",
            expand=True,
            padx=12
        )

        # Dashboard
        self.dashboard_button = self.create_sidebar_button(
            nav_frame,
            "Dashboard",
            lambda: controller.show_frame(SecondPage)
        )

        # Bill
        self.create_sidebar_button(
            nav_frame,
            "Bill",
            lambda: controller.show_frame(BillPage)
        )

        # Add Item
        self.create_sidebar_button(
            nav_frame,
            "Add Item",
            lambda: controller.show_frame(FivePage)
        )

        ## Barcode
        #self.create_sidebar_button(
        #    nav_frame,
        #    "Barcode",
        #    lambda: controller.show_frame(BarcodePage)
        #)
#
        ## Warehouse
        #self.create_sidebar_button(
        #    nav_frame,
        #    "Warehouse",
        #    lambda: controller.show_frame(WarehousePage)
        #)

        # Report
        self.create_sidebar_button(
            nav_frame,
            "Report",
            lambda: controller.show_frame(ReportPage)
        )

        # Setting
        self.create_sidebar_button(
            nav_frame,
            "Setting",
            lambda: controller.show_frame(SettingPage)
        )
        # Logout
        self.create_sidebar_button(
            nav_frame,
            "Logout",
            lambda: controller.show_frame(FirstPage)
        )

        # ============================================================
        # SIDEBAR BOTTOM
        # ============================================================

        bottom_frame = ctk.CTkFrame(
            self.sidebar,
            fg_color="transparent"
        )

        bottom_frame.pack(
            side="bottom",
            fill="x",
            padx=12,
            pady=15
        )

        ctk.CTkLabel(
            bottom_frame,
            text="POS System",
            font=ctk.CTkFont(
                family="Arial",
                size=11
            ),
            text_color="#6b7280"
        ).pack()

        # ============================================================
        # RIGHT DASHBOARD AREA
        # ============================================================

        self.dashboard_area = ctk.CTkFrame(
            self,
            fg_color="#99acd1"
        )

        self.dashboard_area.grid(
            row=0,
            column=1,
            sticky="nsew"
        )

        self.dashboard_area.grid_columnconfigure(
            0,
            weight=1
        )

        self.dashboard_area.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # SCROLLABLE DASHBOARD
        # ============================================================

        self.main_scroll = ctk.CTkScrollableFrame(
            self.dashboard_area,
            fg_color="#f3f5f9"
        )

        self.main_scroll.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        # ============================================================
        # HEADER
        # ============================================================

        header = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent"
        )

        header.pack(
            fill="x",
            padx=15,
            pady=(15, 15)
        )

        header_left = ctk.CTkFrame(
            header,
            fg_color="transparent"
        )

        header_left.pack(
            side="left"
        )

        ctk.CTkLabel(
            header_left,
            text="POS Dashboard",
            font=ctk.CTkFont(
                family="Arial",
                size=30,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w"
        )

        ctk.CTkLabel(
            header_left,
            text="Sales overview and inventory status",
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#6b7280"
        ).pack(
            anchor="w"
        )

        # Date
        self.date_label = ctk.CTkLabel(
            header,
            text=datetime.now().strftime(
                "%d %B %Y"
            ),
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#374151"
        )

        self.date_label.pack(
            side="right",
            pady=10
        )

        # ============================================================
        # SUMMARY CARDS
        # ============================================================

        summary_frame = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent"
        )

        summary_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 15)
        )

        for i in range(3):

            summary_frame.grid_columnconfigure(
                i,
                weight=1
            )

        self.sales_card = self.create_summary_card(
            summary_frame,
            0,
            "Today's Sales",
            "0",
            "💰",
            "#dcfce7"
        )

        self.orders_card = self.create_summary_card(
            summary_frame,
            1,
            "Today's Orders",
            "0",
            "🧾",
            "#dbeafe"
        )

        self.items_card = self.create_summary_card(
            summary_frame,
            2,
            "Total Items",
            "0",
            "📦",
            "#fef3c7"
        )

        # ============================================================
        # QUICK ACTIONS
        # ============================================================

        quick_card = ctk.CTkFrame(
            self.main_scroll,
            fg_color="#ffffff",
            corner_radius=14
        )

        quick_card.pack(
            fill="x",
            padx=15,
            pady=(0, 15)
        )

        ctk.CTkLabel(
            quick_card,
            text="Quick Actions",
            font=ctk.CTkFont(
                family="Arial",
                size=20,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=20,
            pady=(18, 12)
        )

        quick_buttons = ctk.CTkFrame(
            quick_card,
            fg_color="transparent"
        )

        quick_buttons.pack(
            fill="x",
            padx=20,
            pady=(0, 20)
        )

        for i in range(3):

            quick_buttons.grid_columnconfigure(
                i,
                weight=1
            )

        ctk.CTkButton(
            quick_buttons,
            text="New Bill",
            height=48,
            command=lambda: controller.show_frame(BillPage)
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 7)
        )

        ctk.CTkButton(
            quick_buttons,
            text="Add Item",
            height=48,
            fg_color="#16a34a",
            hover_color="#15803d",
            command=lambda: controller.show_frame(FivePage)
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=7
        )

        ctk.CTkButton(
            quick_buttons,
            text="Warehouse",
            height=48,
            fg_color="#9333ea",
            hover_color="#7e22ce",
            command=lambda: controller.show_frame(WarehousePage)
        ).grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(7, 0)
        )

        # ============================================================
        # CHART AREA
        # ============================================================

        chart_container = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent"
        )

        chart_container.pack(
            fill="x",
            padx=15,
            pady=(0, 15)
        )

        chart_container.grid_columnconfigure(
            0,
            weight=3
        )

        chart_container.grid_columnconfigure(
            1,
            weight=1
        )

        # Top Sales
        top_sales_card = ctk.CTkFrame(
            chart_container,
            fg_color="#ffffff",
            corner_radius=14
        )

        top_sales_card.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 7)
        )

        ctk.CTkLabel(
            top_sales_card,
            text="Top 10 Most Sales Items",
            font=ctk.CTkFont(
                family="Arial",
                size=20,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=20,
            pady=(18, 2)
        )

        ctk.CTkLabel(
            top_sales_card,
            text="Frequently sold items",
            font=ctk.CTkFont(
                family="Arial",
                size=12
            ),
            text_color="#6b7280"
        ).pack(
            anchor="w",
            padx=20,
            pady=(0, 10)
        )

        self.top_sales_canvas = tk.Canvas(
            top_sales_card,
            height=300,
            bg="#ffffff",
            highlightthickness=0
        )

        self.top_sales_canvas.pack(
            fill="x",
            padx=15,
            pady=(0, 15)
        )

        # Low Stock
        low_stock_card = ctk.CTkFrame(
            chart_container,
            fg_color="#ffffff",
            corner_radius=14
        )

        low_stock_card.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(7, 0)
        )

        ctk.CTkLabel(
            low_stock_card,
            text="Low Stock",
            font=ctk.CTkFont(
                family="Arial",
                size=20,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(18, 2)
        )

        ctk.CTkLabel(
            low_stock_card,
            text="Items with less than 2",
            font=ctk.CTkFont(
                family="Arial",
                size=12
            ),
            text_color="#6b7280"
        ).pack(
            anchor="w",
            padx=18,
            pady=(0, 10)
        )

        self.low_stock_canvas = tk.Canvas(
            low_stock_card,
            height=300,
            bg="#ffffff",
            highlightthickness=0
        )

        self.low_stock_canvas.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 15)
        )

        # ============================================================
        # FIRST LOAD
        # ============================================================

        self.refresh_dashboard()

    # =================================================================
    # SIDEBAR BUTTON
    # =================================================================

    def create_sidebar_button(
        self,
        parent,
        text,
        command
    ):

        button = ctk.CTkButton(
            parent,
            text=text,
            height=45,
            corner_radius=8,
            fg_color="transparent",
            hover_color="#1f2937",
            text_color="#d1d5db",
            anchor="w",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            command=command
        )

        button.pack(
            fill="x",
            pady=4
        )

        return button

    # =================================================================
    # SUMMARY CARD
    # =================================================================

    def create_summary_card(
        self,
        parent,
        column,
        title,
        value,
        icon,
        icon_bg
    ):

        card = ctk.CTkFrame(
            parent,
            fg_color="#ffffff",
            corner_radius=14,
            height=130
        )

        card.grid(
            row=0,
            column=column,
            sticky="ew",
            padx=7
        )

        card.grid_propagate(False)

        icon_frame = ctk.CTkFrame(
            card,
            width=55,
            height=55,
            corner_radius=12,
            fg_color=icon_bg
        )

        icon_frame.place(
            x=18,
            y=25
        )

        icon_frame.pack_propagate(False)

        ctk.CTkLabel(
            icon_frame,
            text=icon,
            font=ctk.CTkFont(
                size=22
            )
        ).pack(
            expand=True
        )

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#6b7280"
        ).place(
            x=90,
            y=25
        )

        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(
                family="Arial",
                size=24,
                weight="bold"
            ),
            text_color="#111827"
        )

        value_label.place(
            x=90,
            y=52
        )

        return value_label

    # =================================================================
    # REFRESH DASHBOARD
    # =================================================================

    def refresh_dashboard(self):

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        # Today's Sales
        try:

            result = self.dbSales.Total(
                today
            )

            value = 0

            if result:
                value = result[0] or 0

            self.sales_card.configure(
                text=f"{value:,.0f}"
            )

        except Exception as error:

            print(
                "Today's Sales Error:",
                error
            )

            self.sales_card.configure(
                text="0"
            )

        # Today's Orders
        try:

            result = self.dbSales.CountQuan(
                today
            )

            value = 0

            if result:
                value = result[0] or 0

            self.orders_card.configure(
                text=str(value)
            )

        except Exception as error:

            print(
                "Today's Orders Error:",
                error
            )

            self.orders_card.configure(
                text="0"
            )

        # Total Items
        try:

            result = self.dbWarehouse.TotalQuantity()

            value = 0

            if result:
                value = result[0] or 0

            self.items_card.configure(
                text=str(value)
            )

        except Exception as error:

            print(
                "Total Items Error:",
                error
            )

            self.items_card.configure(
                text="0"
            )

        self.load_top_sales_chart()
        self.load_low_stock_chart()

    # =================================================================
    # TOP 10 SALES
    # =================================================================

    def load_top_sales_chart(self):

        canvas = self.top_sales_canvas

        canvas.delete("all")

        try:
            # change itemcode shows to name top 10 type
            self.dbSales.cur.execute(
                """
                SELECT
                    name,
                    SUM(quantity) AS total_quantity
                FROM SalesReport
                WHERE itemcode IS NOT NULL
                AND itemcode != ''
                GROUP BY itemcode
                ORDER BY total_quantity DESC
                LIMIT 10
                """
            )

            rows = self.dbSales.cur.fetchall()

        except Exception as error:

            print(
                "Top Sales Error:",
                error
            )

            rows = []

        if not rows:

            canvas.create_text(
                300,
                140,
                text="No sales data available",
                font=("Arial", 14),
                fill="#9ca3af"
            )

            return

        max_value = max(
            row[1] or 0
            for row in rows
        )

        chart_width = 700
        left = 120
        right = 40
        top = 10
        bar_height = 20
        gap = 7

        for index, row in enumerate(rows):

            item_code = str(
                row[0]
            )

            quantity = row[1] or 0

            y = top + (
                index *
                (bar_height + gap)
            )

            canvas.create_text(
                left - 10,
                y + bar_height / 2,
                text=item_code,
                anchor="e",
                font=("Arial", 10, "bold"),
                fill="#374151"
            )

            canvas.create_rectangle(
                left,
                y,
                chart_width - right,
                y + bar_height,
                fill="#eef2f7",
                outline=""
            )

            if max_value > 0:

                bar_width = (
                    quantity / max_value
                ) * (
                    chart_width -
                    left -
                    right
                )

            else:

                bar_width = 0

            canvas.create_rectangle(
                left,
                y,
                left + bar_width,
                y + bar_height,
                fill="#2563eb",
                outline=""
            )

            canvas.create_text(
                left + bar_width + 8,
                y + bar_height / 2,
                text=str(quantity),
                anchor="w",
                font=("Arial", 10, "bold"),
                fill="#111827"
            )

    # =================================================================
    # LOW STOCK
    # =================================================================

    def load_low_stock_chart(self):

        canvas = self.low_stock_canvas

        canvas.delete("all")

        try:

            self.dbWarehouse.cur.execute(
                """
                SELECT Item, Name, Quantity
                FROM Order2
                WHERE Quantity < 2
                ORDER BY Quantity ASC
                """
            )

            rows = self.dbWarehouse.cur.fetchall()

        except Exception as error:

            print(
                "Low Stock Error:",
                error
            )

            rows = []

        if not rows:

            canvas.create_text(
                150,
                140,
                text="No low-stock items",
                font=("Arial", 13),
                fill="#9ca3af"
            )

            return

        y = 10

        for row in rows:

            item_code = str(
                row[0]
            )

            name = str(
                row[1] or ""
            )

            quantity = row[2] or 0

            canvas.create_rectangle(
                8,
                y,
                330,
                y + 48,
                fill="#fff7ed",
                outline=""
            )

            canvas.create_text(
                20,
                y + 16,
                text=item_code,
                anchor="w",
                font=("Arial", 11, "bold"),
                fill="#111827"
            )

            canvas.create_text(
                20,
                y + 34,
                text=name[:25],
                anchor="w",
                font=("Arial", 9),
                fill="#6b7280"
            )

            canvas.create_text(
                310,
                y + 24,
                text=str(quantity),
                anchor="e",
                font=("Arial", 12, "bold"),
                fill="#dc2626"
            )

            y += 55

            if y > 270:
                break

    # =================================================================
    # REFRESH WHEN DASHBOARD IS OPENED
    # =================================================================

    def tkraise(self, aboveThis=None):

        result = super().tkraise(
            aboveThis
        )

        self.refresh_dashboard()

        return result


# ============================================================
# EMPTY PAGE FUNCTION
# ============================================================

class EmptyPOSPage(ctk.CTkFrame):

    def __init__(
        self,
        parent,
        controller,
        title
    ):

        super().__init__(
            parent,
            fg_color="#f3f5f9"
        )

        self.controller = controller

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        self.header = ctk.CTkFrame(
            self,
            height=80,
            fg_color="#ffffff",
            corner_radius=0
        )

        self.header.pack(
            fill="x"
        )

        self.header.pack_propagate(False)

        ctk.CTkLabel(
            self.header,
            text=title,
            font=ctk.CTkFont(
                family="Arial",
                size=27,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left",
            padx=30
        )

        # ----------------------------------------------------
        # Content
        # ----------------------------------------------------

        self.content = ctk.CTkFrame(
            self,
            fg_color="#ffffff",
            corner_radius=20
        )

        self.content.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=30
        )

        ctk.CTkLabel(
            self.content,
            text=title,
            font=ctk.CTkFont(
                family="Arial",
                size=35,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            pady=(150, 10)
        )

        ctk.CTkLabel(
            self.content,
            text="This page is ready for your POS features.",
            font=ctk.CTkFont(
                family="Arial",
                size=16
            ),
            text_color="#64748b"
        ).pack(
            pady=5
        )

        ctk.CTkButton(
            self.content,
            text="Back to Dashboard",
            width=200,
            height=50,
            corner_radius=10,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            command=lambda: controller.show_frame(SecondPage)
        ).pack(
            pady=30
        )


# ============================================================
# 7 POS PAGES
# ============================================================
class BillPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#f3f5f9"
        )

        self.controller = controller

        # ============================================================
        # DATABASE
        # ============================================================

        self.dbWarehouse = warehouse(
            "real/Warehousedatabase.db"
        )

        self.dbTotal = DatabaseTotal(
            "real/Totaldatabase.db"
        )

        self.dbSales = DatabaseSales(
            "real/SalesReportdatabase.db"
        )

        # ============================================================
        # BILL VARIABLES
        # ============================================================

        self.bill_number = str(
            random.randint(100000, 999999)
        )

        self.cart_items = []

        self.total_amount = 0
        self.user_pay = 0
        self.cashback = 0

        # Prevent saving the same bill more than once
        self.bill_saved = False

        # ============================================================
        # SELLER
        # ============================================================

        self.seller_name = ""

        # ============================================================
        # NUMBER VALIDATION
        # ============================================================

        def only_numbers(value):

            if value == "":
                return True

            return value.isdigit()

        number_command = self.register(
            only_numbers
        )

        # ============================================================
        # MAIN SIDEBAR
        # ============================================================

        self.sidebar = ctk.CTkFrame(
            self,
            width=190,
            fg_color="#1e293b",
            corner_radius=0
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(False)

        # ============================================================
        # LOGO
        # ============================================================

        ctk.CTkLabel(
            self.sidebar,
            text="ABK",
            font=ctk.CTkFont(
                family="Arial",
                size=34,
                weight="bold"
            ),
            text_color="white"
        ).pack(
            pady=(25, 0)
        )

        ctk.CTkLabel(
            self.sidebar,
            text="POS SYSTEM",
            font=ctk.CTkFont(
                family="Arial",
                size=12,
                weight="bold"
            ),
            text_color="#fbbf24"
        ).pack(
            pady=(0, 20)
        )

        # ============================================================
        # SIDEBAR BUTTON
        # ============================================================

        def sidebar_button(
            text,
            page,
            active=False
        ):

            button = ctk.CTkButton(
                self.sidebar,
                text=text,
                height=42,
                corner_radius=8,
                anchor="w",
                font=ctk.CTkFont(
                    family="Arial",
                    size=14,
                    weight="bold"
                ),
                fg_color="#2563eb" if active else "#1e293b",
                hover_color="#334155",
                text_color="white",
                command=lambda: controller.show_frame(page)
            )

            button.pack(
                fill="x",
                padx=15,
                pady=3
            )

        # ============================================================
        # SIDEBAR MENU
        # ============================================================
        sidebar_button(
            "   Dashboard",
            SecondPage
        )

        sidebar_button(
            "   Bill",
            BillPage,
            True
        )

        sidebar_button(
            "   Add Item",
            FivePage
        )

        #sidebar_button(
        #    "   Barcode",
        #    BarcodePage
        #)
#
        #sidebar_button(
        #    "   Warehouse",
        #    WarehousePage
        #)

        sidebar_button(
            "   Report",
            ReportPage
        )

        sidebar_button(
            "   Setting",
            SettingPage
        )

        sidebar_button(
            "   Logout",
            FirstPage
        )

        # ============================================================
        # MAIN AREA
        # ============================================================

        self.main_area = ctk.CTkFrame(
            self,
            fg_color="#d6d9e5",
            corner_radius=0
        )

        self.main_area.pack(
            side="left",
            fill="both",
            expand=True
        )

        # ============================================================
        # TOP BAR
        # ============================================================

        topbar = ctk.CTkFrame(
            self.main_area,
            height=60,
            fg_color="white",
            corner_radius=0
        )

        topbar.pack(
            fill="x"
        )

        topbar.pack_propagate(False)

        ctk.CTkLabel(
            topbar,
            text="Bill / Sales",
            font=ctk.CTkFont(
                family="Arial",
                size=23,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left",
            padx=20
        )

        ctk.CTkLabel(
            topbar,
            text="Magway Software House",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            side="right",
            padx=20
        )

        # ============================================================
        # CONTENT background
        # ============================================================

        content = ctk.CTkFrame(
            self.main_area,
            fg_color="#d6d9e5",
            corner_radius=0
        )

        content.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=15
        )

        content.grid_columnconfigure(
            0,
            weight=1
        )

        content.grid_columnconfigure(
            1,
            weight=1
        )

        content.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # LEFT AREA
        # ============================================================

        left_area = ctk.CTkFrame(
            content,
            fg_color="transparent"
        )

        left_area.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 7)
        )

        # ============================================================
        # TODAY SALES CARD
        # ============================================================

        sales_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        sales_card.pack(
            fill="x",
            pady=(0, 20)
        )

        sales_title_row = ctk.CTkFrame(
            sales_card,
            fg_color="transparent"
        )

        sales_title_row.pack(
            fill="x",
            padx=18,
            pady=(10, 5)
        )

        ctk.CTkLabel(
            sales_title_row,
            text="Today's Sales",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left"
        )

        self.sales_date_label = ctk.CTkLabel(
            sales_title_row,
            text=datetime.now().strftime(
                "%Y-%m-%d"
            ),
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        )

        self.sales_date_label.pack(
            side="right"
        )

        sales_value_frame = ctk.CTkFrame(
            sales_card,
            fg_color="#c0d8f7",
            corner_radius=10
        )

        sales_value_frame.pack(
            fill="x",
            padx=18,
            pady=(0, 12)
        )

        self.today_sales_label = ctk.CTkLabel(
            sales_value_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=25,
                weight="bold"
            ),
            text_color="#2563eb"
        )

        self.today_sales_label.pack(
            pady=8
        )

        # ============================================================
        # GET TODAY SALES
        # ============================================================

        def update_today_sales():

            today = datetime.now().strftime(
                "%Y-%m-%d"
            )

            try:

                result = self.dbSales.Total(
                    today
                )

                if result and result[0] is not None:

                    total_sales = result[0]

                else:

                    total_sales = 0

                self.today_sales_label.configure(
                    text=f"{int(total_sales):,}"
                )

            except Exception:

                self.today_sales_label.configure(
                    text="0"
                )

        # ============================================================
        # BILL INFORMATION CARD
        # ============================================================

        bill_info = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        bill_info.pack(
            fill="x",
            pady=(50, 20)
        )

        ctk.CTkLabel(
            bill_info,
            text="Bill Information",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        bill_info_row = ctk.CTkFrame(
            bill_info,
            fg_color="transparent"
        )

        bill_info_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        bill_info_row.grid_columnconfigure(
            0,
            weight=1
        )

        bill_info_row.grid_columnconfigure(
            1,
            weight=1
        )

        # ============================================================
        # BILL NUMBER
        # ============================================================

        bill_number_frame = ctk.CTkFrame(
            bill_info_row,
            fg_color="#a4c6f2",
            corner_radius=9
        )

        bill_number_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            bill_number_frame,
            text="Bill Number",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=12,
            pady=(7, 0)
        )

        self.bill_number_label = ctk.CTkLabel(
            bill_number_frame,
            text="#" + self.bill_number,
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#2563eb"
        )

        self.bill_number_label.pack(
            anchor="w",
            padx=12,
            pady=(0, 7)
        )

        # ============================================================
        # SELLER
        # ============================================================

        seller_frame = ctk.CTkFrame(
            bill_info_row,
            fg_color="#ade9ee",
            corner_radius=9
        )

        seller_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(5, 0)
        )

        ctk.CTkLabel(
            seller_frame,
            text="Seller",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=12,
            pady=(7, 0)
        )

        self.seller_entry = ctk.CTkEntry(
            seller_frame,
            height=30,
            border_width=0,
            fg_color="transparent",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.seller_entry.pack(
            fill="x",
            padx=7,
            pady=(0, 5)
        )

        self.seller_entry.insert(
            0,
            self.seller_name
        )

        self.seller_entry.configure(
            state="readonly"
        )

        # ============================================================
        # ADD ITEM CARD
        # ============================================================

        add_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        add_card.pack(
            fill="x",
            pady=(0, 20)
        )

        ctk.CTkLabel(
            add_card,
            text="Add Item",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        item_row = ctk.CTkFrame(
            add_card,
            fg_color="transparent"
        )

        item_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        item_row.grid_columnconfigure(
            0,
            weight=2
        )

        item_row.grid_columnconfigure(
            1,
            weight=1
        )

        item_row.grid_columnconfigure(
            2,
            weight=1
        )

        # ============================================================
        # ITEM CODE
        # ============================================================

        item_frame = ctk.CTkFrame(
            item_row,
            fg_color="#b4cde6",
            corner_radius=8
        )

        item_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            item_frame,
            text="Item Code",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#5e7088"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.item_entry = ctk.CTkEntry(
            item_frame,
            height=40,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Scan / Enter",
            font=ctk.CTkFont(
                family="Arial",
                size=18
            )
        )

        self.item_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # ============================================================
        # QUANTITY
        # ============================================================

        qty_frame = ctk.CTkFrame(
            item_row,
            fg_color="#b4cde6",
            corner_radius=8
        )

        qty_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5
        )

        ctk.CTkLabel(
            qty_frame,
            text="Quantity",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#5e7088"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.qty_entry = ctk.CTkEntry(
            qty_frame,
            height=40,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Qty",
            font=ctk.CTkFont(
                family="Arial",
                size=18
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.qty_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # DEFAULT QUANTITY = 1
        self.qty_entry.insert(
            0,
            "1"
        )

        # ============================================================
        # ADD BUTTON
        # ============================================================

        self.add_item_button = ctk.CTkButton(
            item_row,
            text="+  Add Item",
            height=50,
            corner_radius=8,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.add_item_button.grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(5, 0)
        )

        # ============================================================
        # PAYMENT CARD
        # ============================================================

        payment_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        payment_card.pack(
            fill="x"
        )

        ctk.CTkLabel(
            payment_card,
            text="Payment",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        payment_row = ctk.CTkFrame(
            payment_card,
            fg_color="transparent"
        )

        payment_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        payment_row.grid_columnconfigure(
            0,
            weight=1
        )

        payment_row.grid_columnconfigure(
            1,
            weight=1
        )

        payment_row.grid_columnconfigure(
            2,
            weight=1
        )

        # ============================================================
        # TOTAL
        # ============================================================

        total_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#b4f5cc",
            corner_radius=8
        )

        total_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            total_frame,
            text="TOTAL",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.total_label = ctk.CTkLabel(
            total_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=17,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.total_label.pack(
            anchor="w",
            padx=10,
            pady=(0, 6)
        )

        # ============================================================
        # USER PAY
        # ============================================================

        pay_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#b6d0f2",
            corner_radius=8
        )

        pay_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5
        )

        ctk.CTkLabel(
            pay_frame,
            text="USER PAY",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.pay_entry = ctk.CTkEntry(
            pay_frame,
            height=30,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Payment",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.pay_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # ============================================================
        # CASHBACK
        # ============================================================

        cashback_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#e6ebab",
            corner_radius=8
        )

        cashback_frame.grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(5, 0)
        )

        ctk.CTkLabel(
            cashback_frame,
            text="CASHBACK",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.cashback_label = ctk.CTkLabel(
            cashback_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=17,
                weight="bold"
            ),
            text_color="#16a34a"
        )

        self.cashback_label.pack(
            anchor="w",
            padx=10,
            pady=(0, 6)
        )

        # ============================================================
        # RIGHT AREA
        # ============================================================

        right_area = ctk.CTkFrame(
            content,
            fg_color="transparent"
        )

        right_area.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(7, 0)
        )

        # ============================================================
        # RECEIPT CARD
        # ============================================================

        receipt_card = ctk.CTkFrame(
            right_area,
            fg_color="#cbd5e1",
            corner_radius=14
        )

        receipt_card.pack(
            fill="both",
            expand=True
        )

        ctk.CTkLabel(
            receipt_card,
            text="Receipt Preview • 80mm",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155"
        ).pack(
            pady=(8, 5)
        )

        # ============================================================
        # RECEIPT PAPER
        # ============================================================

        receipt_frame = ctk.CTkFrame(
            receipt_card,
            fg_color="white",
            corner_radius=3
        )

        receipt_frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 8)
        )

        # ============================================================
        # PREVIEW FONT
        # Larger for normal screen viewing
        # ============================================================

        preview_font = ctk.CTkFont(
            family="Courier New",
            size=11
        )

        preview_bold_font = ctk.CTkFont(
            family="Courier New",
            size=11,
            weight="bold"
        )

        # ============================================================
        # RECEIPT HEADER
        # ============================================================

        ctk.CTkLabel(
            receipt_frame,
            text="အောင်ဘုန်းခန့် MiniMart",
            font=ctk.CTkFont(
                family="Pyidaungsu",
                size=20,
                weight="bold"
            ),
            text_color="black"
        ).pack(
            pady=(8, 0)
        )

        ctk.CTkLabel(
            receipt_frame,
            text="Ph:09777775706",
            font=ctk.CTkFont(
                family="Arial",
                size=11
            ),
            text_color="black"
        ).pack(
            pady=(0, 2)
        )

        self.receipt_bill_label = ctk.CTkLabel(
            receipt_frame,
            text="Bill Number: " +
            self.bill_number,
            font=preview_bold_font,
            text_color="black"
        )

        self.receipt_bill_label.pack()

        self.receipt_date_label = ctk.CTkLabel(
            receipt_frame,
            text="Date: " +
            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),
            font=preview_font,
            text_color="black"
        )

        self.receipt_date_label.pack(
            pady=(0, 3)
        )

        ctk.CTkLabel(
            receipt_frame,
            text="-" * 48,
            font=ctk.CTkFont(
                family="Courier New",
                size=10
            ),
            text_color="black"
        ).pack()

        # ============================================================
        # TABLE HEADER
        # ============================================================

        header = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        header.pack(
            fill="x",
            padx=8
        )

        ctk.CTkLabel(
            header,
            text="No",
            width=30,
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Name",
            width=105,
            anchor="w",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Qty",
            width=35,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Price",
            width=55,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Amount",
            width=65,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        # ============================================================
        # RECEIPT ITEMS
        # ============================================================

        self.receipt_items = ctk.CTkScrollableFrame(
            receipt_frame,
            fg_color="white",
            corner_radius=0
        )

        self.receipt_items.pack(
            fill="both",
            expand=True,
            padx=4,
            pady=2
        )

        # ============================================================
        # RECEIPT BOTTOM
        # ============================================================

        ctk.CTkLabel(
            receipt_frame,
            text="-" * 48,
            font=ctk.CTkFont(
                family="Courier New",
                size=10
            ),
            text_color="black"
        ).pack()

        total_receipt_row = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        total_receipt_row.pack(
            fill="x",
            padx=12
        )

        ctk.CTkLabel(
            total_receipt_row,
            text="TOTAL",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        self.receipt_total_label = ctk.CTkLabel(
            total_receipt_row,
            text="0",
            font=preview_bold_font,
            text_color="black"
        )

        self.receipt_total_label.pack(
            side="right"
        )

        cashback_receipt_row = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        cashback_receipt_row.pack(
            fill="x",
            padx=12
        )

        ctk.CTkLabel(
            cashback_receipt_row,
            text="Cashback",
            font=preview_font,
            text_color="black"
        ).pack(
            side="left"
        )

        self.receipt_cashback_label = ctk.CTkLabel(
            cashback_receipt_row,
            text="0",
            font=preview_font,
            text_color="black"
        )

        self.receipt_cashback_label.pack(
            side="right"
        )

        ctk.CTkLabel(
            receipt_frame,
            text="Thank You!",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="black"
        ).pack(
            pady=(3, 5)
        )

        # ============================================================
        # BOTTOM BUTTONS
        # ============================================================

        button_row = ctk.CTkFrame(
            right_area,
            fg_color="transparent"
        )

        button_row.pack(
            fill="x",
            pady=(8, 0)
        )

        button_row.grid_columnconfigure(
            0,
            weight=1
        )

        button_row.grid_columnconfigure(
            1,
            weight=1
        )

        # ============================================================
        # CLEAR BUTTON
        # ============================================================

        self.clear_button = ctk.CTkButton(
            button_row,
            text="Clear Bill",
            height=44,
            corner_radius=9,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.clear_button.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        # ============================================================
        # PRINT BUTTON
        # ============================================================

        self.print_button = ctk.CTkButton(
            button_row,
            text="Print",
            height=44,
            corner_radius=9,
            fg_color="#16a34a",
            hover_color="#15803d",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.print_button.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(5, 0)
        )

        # ============================================================
        # PAYMENT CALCULATION
        # ============================================================

        def update_payment(event=None):

            try:

                if self.pay_entry.get() == "":

                    self.user_pay = 0

                else:

                    self.user_pay = int(
                        self.pay_entry.get()
                    )

                self.cashback = (
                    self.user_pay -
                    self.total_amount
                )

                self.cashback_label.configure(
                    text=f"{self.cashback:,}"
                )

                self.receipt_cashback_label.configure(
                    text=f"{self.cashback:,}"
                )

            except:

                self.user_pay = 0
                self.cashback = 0

                self.cashback_label.configure(
                    text="0"
                )

                self.receipt_cashback_label.configure(
                    text="0"
                )

        self.pay_entry.bind(
            "<KeyRelease>",
            update_payment
        )

        # ============================================================
        # REMOVE ITEM
        # ============================================================

        def remove_item(item):

            if item in self.cart_items:

                self.cart_items.remove(
                    item
                )

            self.calculate_total()

            refresh_receipt()

            self.bill_saved = False

            self.item_entry.focus()

        # ============================================================
        # REFRESH RECEIPT
        # ============================================================

        def refresh_receipt():

            for widget in (
                self.receipt_items.winfo_children()
            ):

                widget.destroy()

            for index, item in enumerate(
                self.cart_items,
                start=1
            ):

                row = ctk.CTkFrame(
                    self.receipt_items,
                    fg_color="white"
                )

                row.pack(
                    fill="x",
                    padx=2,
                    pady=2
                )

                ctk.CTkLabel(
                    row,
                    text=str(index),
                    width=30,
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                name = str(
                    item["name"]
                )

                if len(name) > 16:

                    name = name[:16]

                ctk.CTkLabel(
                    row,
                    text=name,
                    width=105,
                    anchor="w",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["quantity"]
                    ),
                    width=35,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["price"]
                    ),
                    width=55,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["amount"]
                    ),
                    width=65,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                remove_button = ctk.CTkButton(
                    row,
                    text="X",
                    width=30,
                    height=25,
                    corner_radius=5,
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    font=ctk.CTkFont(
                        family="Arial",
                        size=10,
                        weight="bold"
                    ),
                    command=lambda
                    current_item=item:
                    remove_item(
                        current_item
                    )
                )

                remove_button.pack(
                    side="left",
                    padx=(5, 0)
                )

            self.total_label.configure(
                text=f"{self.total_amount:,}"
            )

            self.receipt_total_label.configure(
                text=f"{self.total_amount:,}"
            )

            update_payment()

        # ============================================================
        # CALCULATE TOTAL
        # ============================================================

        def calculate_total():

            self.total_amount = 0

            for item in self.cart_items:

                self.total_amount += (
                    item["amount"]
                )

        self.calculate_total = calculate_total

        # ============================================================
        # ADD ITEM
        # ============================================================

        def add_item(event=None):

            ItemCode = (
                self.item_entry
                .get()
                .strip()
            )

            QtyText = (
                self.qty_entry
                .get()
                .strip()
            )

            if ItemCode == "":

                messagebox.showwarning(
                    "Warning",
                    "Please enter Item Code."
                )

                self.item_entry.focus()

                return

            if QtyText == "":

                QtyText = "1"

            Qty = int(
                QtyText
            )

            if Qty <= 0:

                messagebox.showwarning(
                    "Warning",
                    "Quantity must be greater than 0."
                )

                self.qty_entry.focus()

                return

            rows = self.dbWarehouse.Search(
                ItemCode
            )

            if not rows:

                messagebox.showerror(
                    "Item Not Found",
                    "Item Code does not exist.ဤပစွည်းမရှိသေးပါရှင့်။"
                )

                self.item_entry.focus()

                return

            item = rows[0]

            DBQuantity = int(
                item[4]
            )

            Name = item[2]

            Price = int(
                item[6]
            )

            already_in_cart = 0

            for cart_item in self.cart_items:

                if (
                    cart_item["itemcode"]
                    == ItemCode
                ):

                    already_in_cart = (
                        cart_item["quantity"]
                    )

                    break

            total_requested = (
                already_in_cart +
                Qty
            )

            if total_requested > DBQuantity:

                messagebox.showwarning(
                    "Not Enough Stock, ပစွည်းကုန်သွားပါပြီရှင့်",
                    "Available quantity, ဝယ်ယူနိုင်သောအရေအတွက် = "
                    + str(DBQuantity) + " ခု"
                )

                return

            found = False

            for cart_item in self.cart_items:

                if (
                    cart_item["itemcode"]
                    == ItemCode
                ):

                    cart_item["quantity"] += Qty

                    cart_item["amount"] = (
                        cart_item["quantity"]
                        * Price
                    )

                    found = True

                    break

            if not found:

                self.cart_items.append(
                    {
                        "itemcode": ItemCode,
                        "name": Name,
                        "quantity": Qty,
                        "price": Price,
                        "amount": Qty * Price
                    }
                )

            calculate_total()

            refresh_receipt()

            # New item starts with quantity 1
            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.item_entry.focus()

            self.bill_saved = False

        self.add_item_button.configure(
            command=add_item
        )

        # Enter on Item Code
        self.item_entry.bind(
            "<Return>",
            add_item
        )

        # Enter on Quantity
        self.qty_entry.bind(
            "<Return>",
            add_item
        )

        # ============================================================
        # SAVE BILL TO DATABASE
        # ============================================================

        def save_bill():

            if len(self.cart_items) == 0:

                messagebox.showwarning(
                    "Warning",
                    "Please add items before saving."
                )

                return False

            if self.bill_saved:

                messagebox.showinfo(
                    "Already Saved",
                    "This bill has already been saved."
                )

                return True

            seller = self.seller_name

            if seller == "":

                seller = getattr(
                    controller,
                    "current_username",
                    ""
                )

            if seller == "":

                messagebox.showwarning(
                    "Seller",
                    "Seller username is empty."
                )

                return False

            # --------------------------------------------------------
            # USER PAY
            # --------------------------------------------------------

            if self.pay_entry.get() == "":

                messagebox.showwarning(
                    "Payment",
                    "Please enter User Pay."
                )

                self.pay_entry.focus()

                return False

            try:

                user_pay = int(
                    self.pay_entry.get()
                )

            except:

                messagebox.showwarning(
                    "Payment",
                    "Please enter a valid amount."
                )

                self.pay_entry.focus()

                return False

            if user_pay < self.total_amount:

                messagebox.showwarning(
                    "Payment Error",
                    "User Pay is less than Total."
                )

                self.pay_entry.focus()

                return False

            self.user_pay = user_pay

            self.cashback = (
                user_pay -
                self.total_amount
            )

            update_payment()

            # --------------------------------------------------------
            # DATE
            # --------------------------------------------------------

            now = datetime.now()

            bill_date = now.strftime(
                "%Y-%m-%d"
            )

            try:

                # ====================================================
                # SAVE EACH ITEM
                # ====================================================

                for item in self.cart_items:

                    ItemCode = item["itemcode"]

                    Name = item["name"]

                    Qty = item["quantity"]

                    Price = item["price"]

                    Amount = item["amount"]

                    # ----------------------------------------------
                    # AddTotal
                    # ----------------------------------------------

                    self.dbTotal.insert(
                        ItemCode,
                        Name,
                        str(Amount)
                    )

                    # ----------------------------------------------
                    # AddTotalBillNum
                    # ----------------------------------------------

                    self.dbTotal.insertBill(
                        self.bill_number,
                        ItemCode,
                        Name,
                        Qty,
                        Price,
                        str(Amount),
                        seller
                    )

                    # ----------------------------------------------
                    # SalesReport
                    # ----------------------------------------------

                    self.dbSales.insert(
                        bill_date,
                        Amount,
                        self.bill_number,
                        seller,
                        Qty,
                        ItemCode,
                        Name,
                    )

                    # ----------------------------------------------
                    # Warehouse quantity decrease
                    # ----------------------------------------------

                    current = (
                        self.dbWarehouse
                        .getQuntity(
                            ItemCode
                        )
                    )

                    if current:

                        current_qty = int(
                            current[0]
                        )

                        new_quantity = (
                            current_qty -
                            Qty
                        )

                        if new_quantity < 0:

                            new_quantity = 0

                        self.dbWarehouse.updateQuantity(
                            ItemCode,
                            new_quantity
                        )

                # ====================================================
                # SAVED
                # ====================================================

                self.bill_saved = True

                update_today_sales()

                messagebox.showinfo(
                    "Saved",
                    "Bill saved successfully.\n\n"
                    "You can now press Print."
                )

                return True

            except Exception as e:

                messagebox.showerror(
                    "Database Error",
                    str(e)
                )

                return False

        # ============================================================
        # USER PAY ENTER = SAVE ONLY
        # ============================================================

        self.pay_entry.bind(
            "<Return>",
            lambda event: save_bill()
        )

        # ============================================================
        # CLEAR BILL
        # ============================================================

        def clear_bill():

            if (
                len(self.cart_items) == 0
                and not self.bill_saved
            ):

                return

            answer = messagebox.askyesno(
                "Clear Bill",
                "Are you sure you want to "
                "clear this bill?"
            )

            if not answer:

                return

            self.cart_items.clear()

            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0

            self.bill_saved = False

            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.pay_entry.delete(
                0,
                "end"
            )

            self.total_label.configure(
                text="0"
            )

            self.cashback_label.configure(
                text="0"
            )

            self.receipt_total_label.configure(
                text="0"
            )

            self.receipt_cashback_label.configure(
                text="0"
            )

            refresh_receipt()

            self.item_entry.focus()

        self.clear_button.configure(
            command=clear_bill
        )

        # ============================================================
        # PRINT BILL
        # ============================================================

        def print_bill():

            if len(self.cart_items) == 0:

                messagebox.showwarning(
                    "Warning",
                    "There are no items to print."
                )

                return

            # --------------------------------------------------------
            # If not saved, ask user to save first
            # --------------------------------------------------------

            if not self.bill_saved:

                answer = messagebox.askyesno(
                    "Bill Not Saved",
                    "This bill has not been saved.\n\n"
                    "Save it before printing?"
                )

                if not answer:

                    return

                if not save_bill():

                    return

            seller = self.seller_name

            if seller == "":

                seller = getattr(
                    controller,
                    "current_username",
                    ""
                )

            now = datetime.now()

            bill_datetime = now.strftime(
                "%Y-%m-%d %H:%M"
            )

            # ========================================================
            # 80MM PRINTER TEXT
            #
            # This remains smaller / fixed-width for the printer.
            # The preview above uses a larger font.
            # ========================================================

            receipt_lines = []

            receipt_lines.append(
                "အောင်ဘုန်းခန့် MiniMart"
            )

            receipt_lines.append(
                "Ph:09777775706"
            )

            receipt_lines.append(
                "Bill Number: "
                + self.bill_number
            )

            receipt_lines.append(
                "Date: "
                + bill_datetime
            )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                "No  Name              "
                "Qty Price  Amount"
            )

            receipt_lines.append(
                "-" * 48
            )

            for index, item in enumerate(
                self.cart_items,
                start=1
            ):

                name = str(
                    item["name"]
                )

                if len(name) > 17:

                    name = name[:17]

                line = (
                    f"{index:<3}"
                    f"{name:<18}"
                    f"{item['quantity']:>4}"
                    f"{item['price']:>6}"
                    f"{item['amount']:>9}"
                )

                receipt_lines.append(
                    line
                )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                f"{'TOTAL':<30}"
                f"{self.total_amount:>18}"
            )

            receipt_lines.append(
                f"{'PAY':<30}"
                f"{self.user_pay:>18}"
            )

            receipt_lines.append(
                f"{'CASHBACK':<30}"
                f"{self.cashback:>18}"
            )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                "Seller: "
                + seller
            )

            receipt_lines.append(
                "Thank You!"
            )

            receipt_text = "\n".join(
                receipt_lines
            )

            # ========================================================
            # TEMP FILE
            # ========================================================

            receipt_file = os.path.join(
                tempfile.gettempdir(),
                "ABK_Bill_"
                + self.bill_number
                + ".txt"
            )

            try:

                with open(
                    receipt_file,
                    "w",
                    encoding="utf-8"
                ) as file:

                    file.write(
                        receipt_text
                    )

            except Exception as e:

                messagebox.showerror(
                    "File Error",
                    str(e)
                )

                return

            # ========================================================
            # WINDOWS PRINT
            # ========================================================

            try:

                os.startfile(
                    receipt_file,
                    "print"
                )

            except Exception as e:

                messagebox.showerror(
                    "Printer Error",
                    "Windows could not send "
                    "the receipt to printer.\n\n"
                    + str(e)
                )

                return

            # ========================================================
            # PRINT SUCCESS
            # ========================================================

            messagebox.showinfo(
                "Print",
                "Bill sent to printer."
            )

            # ========================================================
            # START NEW BILL
            # ========================================================

            self.cart_items.clear()

            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0

            self.bill_saved = False

            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.pay_entry.delete(
                0,
                "end"
            )

            # ========================================================
            # NEW BILL NUMBER
            # ========================================================

            self.bill_number = str(
                random.randint(
                    100000,
                    999999
                )
            )

            self.bill_number_label.configure(
                text="#"
                + self.bill_number
            )

            self.receipt_bill_label.configure(
                text="Bill Number: "
                + self.bill_number
            )

            self.receipt_date_label.configure(
                text="Date: "
                + datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )

            self.total_label.configure(
                text="0"
            )

            self.cashback_label.configure(
                text="0"
            )

            self.receipt_total_label.configure(
                text="0"
            )

            self.receipt_cashback_label.configure(
                text="0"
            )

            refresh_receipt()

            update_today_sales()

            self.item_entry.focus()

        self.print_button.configure(
            command=print_bill
        )

        # ============================================================
        # INITIAL
        # ============================================================

        refresh_receipt()

        update_today_sales()

        self.item_entry.focus()

    def refresh_seller(self):

        username = getattr(
            self.controller,
            "current_username",
            ""
        )

        self.seller_name = username

        self.seller_entry.configure(
            state="normal"
        )

        self.seller_entry.delete(
            0,
            "end"
        )

        if username:
            self.seller_entry.insert(
                0,
                username
            )

        self.seller_entry.configure(
            state="readonly"


        )


    def tkraise(self, aboveThis=None):

        result = super().tkraise(aboveThis)

        # Update seller every time BillPage is opened
        if hasattr(self, "seller_entry"):
            self.refresh_seller()

        return result        


# ======================= User BillPage ======================
class UserBillPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#f3f5f9"
        )

        self.controller = controller

        # ============================================================
        # DATABASE
        # ============================================================

        self.dbWarehouse = warehouse(
            "real/Warehousedatabase.db"
        )

        self.dbTotal = DatabaseTotal(
            "real/Totaldatabase.db"
        )

        self.dbSales = DatabaseSales(
            "real/SalesReportdatabase.db"
        )

        # ============================================================
        # BILL VARIABLES
        # ============================================================

        self.bill_number = str(
            random.randint(100000, 999999)
        )

        self.cart_items = []

        self.total_amount = 0
        self.user_pay = 0
        self.cashback = 0

        # Prevent saving the same bill more than once
        self.bill_saved = False

        # ============================================================
        # SELLER
        # ============================================================

        self.seller_name = ""

        # ============================================================
        # NUMBER VALIDATION
        # ============================================================

        def only_numbers(value):

            if value == "":
                return True

            return value.isdigit()

        number_command = self.register(
            only_numbers
        )

        # ============================================================
        # MAIN SIDEBAR
        # ============================================================

        self.sidebar = ctk.CTkFrame(
            self,
            width=190,
            fg_color="#1e293b",
            corner_radius=0
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(False)

        # ============================================================
        # LOGO
        # ============================================================

        ctk.CTkLabel(
            self.sidebar,
            text="ABK",
            font=ctk.CTkFont(
                family="Arial",
                size=34,
                weight="bold"
            ),
            text_color="white"
        ).pack(
            pady=(25, 0)
        )

        ctk.CTkLabel(
            self.sidebar,
            text="POS SYSTEM",
            font=ctk.CTkFont(
                family="Arial",
                size=12,
                weight="bold"
            ),
            text_color="#fbbf24"
        ).pack(
            pady=(0, 20)
        )

        # ============================================================
        # SIDEBAR BUTTON
        # ============================================================

        def sidebar_button(
            text,
            page,
            active=False
        ):

            button = ctk.CTkButton(
                self.sidebar,
                text=text,
                height=42,
                corner_radius=8,
                anchor="w",
                font=ctk.CTkFont(
                    family="Arial",
                    size=14,
                    weight="bold"
                ),
                fg_color="#2563eb" if active else "#1e293b",
                hover_color="#334155",
                text_color="white",
                command=lambda: controller.show_frame(page)
            )

            button.pack(
                fill="x",
                padx=15,
                pady=3
            )

        # ============================================================
        # SIDEBAR MENU
        # ============================================================
        

        sidebar_button(
            "   Bill",
            BillPage,
            True
        )

        #sidebar_button(
        #    "   Add Item",
        #    FivePage
        #)
#
        #sidebar_button(
        #    "   Barcode",
        #    BarcodePage
        #)
#
        #sidebar_button(
        #    "   Warehouse",
        #    WarehousePage
        #)
#
        #sidebar_button(
        #    "   Report",
        #    ReportPage
        #)
#
        #sidebar_button(
        #    "   Setting",
        #    SettingPage
        #)

        sidebar_button(
            "   Logout",
            FirstPage,
        )

        # ============================================================
        # MAIN AREA
        # ============================================================

        self.main_area = ctk.CTkFrame(
            self,
            fg_color="#d6d9e5",
            corner_radius=0
        )

        self.main_area.pack(
            side="left",
            fill="both",
            expand=True
        )

        # ============================================================
        # TOP BAR
        # ============================================================

        topbar = ctk.CTkFrame(
            self.main_area,
            height=60,
            fg_color="white",
            corner_radius=0
        )

        topbar.pack(
            fill="x"
        )

        topbar.pack_propagate(False)

        ctk.CTkLabel(
            topbar,
            text="Bill / Sales",
            font=ctk.CTkFont(
                family="Arial",
                size=23,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left",
            padx=20
        )

        ctk.CTkLabel(
            topbar,
            text="Magway Software House",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            side="right",
            padx=20
        )

        # ============================================================
        # CONTENT background
        # ============================================================

        content = ctk.CTkFrame(
            self.main_area,
            fg_color="#d6d9e5",
            corner_radius=0
        )

        content.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=15
        )

        content.grid_columnconfigure(
            0,
            weight=1
        )

        content.grid_columnconfigure(
            1,
            weight=1
        )

        content.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # LEFT AREA
        # ============================================================

        left_area = ctk.CTkFrame(
            content,
            fg_color="transparent"
        )

        left_area.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 7)
        )

        # ============================================================
        # TODAY SALES CARD
        # ============================================================

        sales_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        sales_card.pack(
            fill="x",
            pady=(0, 20)
        )

        sales_title_row = ctk.CTkFrame(
            sales_card,
            fg_color="transparent"
        )

        sales_title_row.pack(
            fill="x",
            padx=18,
            pady=(10, 5)
        )

        ctk.CTkLabel(
            sales_title_row,
            text="Today's Sales",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left"
        )

        self.sales_date_label = ctk.CTkLabel(
            sales_title_row,
            text=datetime.now().strftime(
                "%Y-%m-%d"
            ),
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        )

        self.sales_date_label.pack(
            side="right"
        )

        sales_value_frame = ctk.CTkFrame(
            sales_card,
            fg_color="#c0d8f7",
            corner_radius=10
        )

        sales_value_frame.pack(
            fill="x",
            padx=18,
            pady=(0, 12)
        )

        self.today_sales_label = ctk.CTkLabel(
            sales_value_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=25,
                weight="bold"
            ),
            text_color="#2563eb"
        )

        self.today_sales_label.pack(
            pady=8
        )

        # ============================================================
        # GET TODAY SALES
        # ============================================================

        def update_today_sales():

            today = datetime.now().strftime(
                "%Y-%m-%d"
            )

            try:

                result = self.dbSales.Total(
                    today
                )

                if result and result[0] is not None:

                    total_sales = result[0]

                else:

                    total_sales = 0

                self.today_sales_label.configure(
                    text=f"{int(total_sales):,}"
                )

            except Exception:

                self.today_sales_label.configure(
                    text="0"
                )

        # ============================================================
        # BILL INFORMATION CARD
        # ============================================================

        bill_info = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        bill_info.pack(
            fill="x",
            pady=(50, 20)
        )

        ctk.CTkLabel(
            bill_info,
            text="Bill Information",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        bill_info_row = ctk.CTkFrame(
            bill_info,
            fg_color="transparent"
        )

        bill_info_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        bill_info_row.grid_columnconfigure(
            0,
            weight=1
        )

        bill_info_row.grid_columnconfigure(
            1,
            weight=1
        )

        # ============================================================
        # BILL NUMBER
        # ============================================================

        bill_number_frame = ctk.CTkFrame(
            bill_info_row,
            fg_color="#a4c6f2",
            corner_radius=9
        )

        bill_number_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            bill_number_frame,
            text="Bill Number",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=12,
            pady=(7, 0)
        )

        self.bill_number_label = ctk.CTkLabel(
            bill_number_frame,
            text="#" + self.bill_number,
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#2563eb"
        )

        self.bill_number_label.pack(
            anchor="w",
            padx=12,
            pady=(0, 7)
        )

        # ============================================================
        # SELLER
        # ============================================================

        seller_frame = ctk.CTkFrame(
            bill_info_row,
            fg_color="#ade9ee",
            corner_radius=9
        )

        seller_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(5, 0)
        )

        ctk.CTkLabel(
            seller_frame,
            text="Seller",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=12,
            pady=(7, 0)
        )

        self.seller_entry = ctk.CTkEntry(
            seller_frame,
            height=30,
            border_width=0,
            fg_color="transparent",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.seller_entry.pack(
            fill="x",
            padx=7,
            pady=(0, 5)
        )

        self.seller_entry.insert(
            0,
            self.seller_name
        )

        self.seller_entry.configure(
            state="readonly"
        )

        # ============================================================
        # ADD ITEM CARD
        # ============================================================

        add_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        add_card.pack(
            fill="x",
            pady=(0, 20)
        )

        ctk.CTkLabel(
            add_card,
            text="Add Item",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        item_row = ctk.CTkFrame(
            add_card,
            fg_color="transparent"
        )

        item_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        item_row.grid_columnconfigure(
            0,
            weight=2
        )

        item_row.grid_columnconfigure(
            1,
            weight=1
        )

        item_row.grid_columnconfigure(
            2,
            weight=1
        )

        # ============================================================
        # ITEM CODE
        # ============================================================

        item_frame = ctk.CTkFrame(
            item_row,
            fg_color="#b4cde6",
            corner_radius=8
        )

        item_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            item_frame,
            text="Item Code",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#5e7088"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.item_entry = ctk.CTkEntry(
            item_frame,
            height=40,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Scan / Enter",
            font=ctk.CTkFont(
                family="Arial",
                size=18
            )
        )

        self.item_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # ============================================================
        # QUANTITY
        # ============================================================

        qty_frame = ctk.CTkFrame(
            item_row,
            fg_color="#b4cde6",
            corner_radius=8
        )

        qty_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5
        )

        ctk.CTkLabel(
            qty_frame,
            text="Quantity",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#5e7088"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.qty_entry = ctk.CTkEntry(
            qty_frame,
            height=40,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Qty",
            font=ctk.CTkFont(
                family="Arial",
                size=18
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.qty_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # DEFAULT QUANTITY = 1
        self.qty_entry.insert(
            0,
            "1"
        )

        # ============================================================
        # ADD BUTTON
        # ============================================================

        self.add_item_button = ctk.CTkButton(
            item_row,
            text="+  Add Item",
            height=50,
            corner_radius=8,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.add_item_button.grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(5, 0)
        )

        # ============================================================
        # PAYMENT CARD
        # ============================================================

        payment_card = ctk.CTkFrame(
            left_area,
            fg_color="white",
            corner_radius=14
        )

        payment_card.pack(
            fill="x"
        )

        ctk.CTkLabel(
            payment_card,
            text="Payment",
            font=ctk.CTkFont(
                family="Arial",
                size=18,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=18,
            pady=(12, 8)
        )

        payment_row = ctk.CTkFrame(
            payment_card,
            fg_color="transparent"
        )

        payment_row.pack(
            fill="x",
            padx=18,
            pady=(0, 14)
        )

        payment_row.grid_columnconfigure(
            0,
            weight=1
        )

        payment_row.grid_columnconfigure(
            1,
            weight=1
        )

        payment_row.grid_columnconfigure(
            2,
            weight=1
        )

        # ============================================================
        # TOTAL
        # ============================================================

        total_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#b4f5cc",
            corner_radius=8
        )

        total_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        ctk.CTkLabel(
            total_frame,
            text="TOTAL",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.total_label = ctk.CTkLabel(
            total_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=17,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.total_label.pack(
            anchor="w",
            padx=10,
            pady=(0, 6)
        )

        # ============================================================
        # USER PAY
        # ============================================================

        pay_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#b6d0f2",
            corner_radius=8
        )

        pay_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5
        )

        ctk.CTkLabel(
            pay_frame,
            text="USER PAY",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.pay_entry = ctk.CTkEntry(
            pay_frame,
            height=30,
            border_width=0,
            fg_color="transparent",
            placeholder_text="Payment",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.pay_entry.pack(
            fill="x",
            padx=5,
            pady=(0, 4)
        )

        # ============================================================
        # CASHBACK
        # ============================================================

        cashback_frame = ctk.CTkFrame(
            payment_row,
            fg_color="#e6ebab",
            corner_radius=8
        )

        cashback_frame.grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(5, 0)
        )

        ctk.CTkLabel(
            cashback_frame,
            text="CASHBACK",
            font=ctk.CTkFont(
                family="Arial",
                size=10,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            anchor="w",
            padx=10,
            pady=(5, 0)
        )

        self.cashback_label = ctk.CTkLabel(
            cashback_frame,
            text="0",
            font=ctk.CTkFont(
                family="Arial",
                size=17,
                weight="bold"
            ),
            text_color="#16a34a"
        )

        self.cashback_label.pack(
            anchor="w",
            padx=10,
            pady=(0, 6)
        )

        # ============================================================
        # RIGHT AREA
        # ============================================================

        right_area = ctk.CTkFrame(
            content,
            fg_color="transparent"
        )

        right_area.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(7, 0)
        )

        # ============================================================
        # RECEIPT CARD
        # ============================================================

        receipt_card = ctk.CTkFrame(
            right_area,
            fg_color="#cbd5e1",
            corner_radius=14
        )

        receipt_card.pack(
            fill="both",
            expand=True
        )

        ctk.CTkLabel(
            receipt_card,
            text="Receipt Preview • 80mm",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155"
        ).pack(
            pady=(8, 5)
        )

        # ============================================================
        # RECEIPT PAPER
        # ============================================================

        receipt_frame = ctk.CTkFrame(
            receipt_card,
            fg_color="white",
            corner_radius=3
        )

        receipt_frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 8)
        )

        # ============================================================
        # PREVIEW FONT
        # Larger for normal screen viewing
        # ============================================================

        preview_font = ctk.CTkFont(
            family="Courier New",
            size=11
        )

        preview_bold_font = ctk.CTkFont(
            family="Courier New",
            size=11,
            weight="bold"
        )

        # ============================================================
        # RECEIPT HEADER
        # ============================================================

        ctk.CTkLabel(
            receipt_frame,
            text="အောင်ဘုန်းခန့် MiniMart",
            font=ctk.CTkFont(
                family="Pyidaungsu",
                size=20,
                weight="bold"
            ),
            text_color="black"
        ).pack(
            pady=(8, 0)
        )

        ctk.CTkLabel(
            receipt_frame,
            text="Ph:09777775706",
            font=ctk.CTkFont(
                family="Arial",
                size=11
            ),
            text_color="black"
        ).pack(
            pady=(0, 2)
        )

        self.receipt_bill_label = ctk.CTkLabel(
            receipt_frame,
            text="Bill Number: " +
            self.bill_number,
            font=preview_bold_font,
            text_color="black"
        )

        self.receipt_bill_label.pack()

        self.receipt_date_label = ctk.CTkLabel(
            receipt_frame,
            text="Date: " +
            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),
            font=preview_font,
            text_color="black"
        )

        self.receipt_date_label.pack(
            pady=(0, 3)
        )

        ctk.CTkLabel(
            receipt_frame,
            text="-" * 48,
            font=ctk.CTkFont(
                family="Courier New",
                size=10
            ),
            text_color="black"
        ).pack()

        # ============================================================
        # TABLE HEADER
        # ============================================================

        header = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        header.pack(
            fill="x",
            padx=8
        )

        ctk.CTkLabel(
            header,
            text="No",
            width=30,
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Name",
            width=105,
            anchor="w",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Qty",
            width=35,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Price",
            width=55,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            header,
            text="Amount",
            width=65,
            anchor="e",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        # ============================================================
        # RECEIPT ITEMS
        # ============================================================

        self.receipt_items = ctk.CTkScrollableFrame(
            receipt_frame,
            fg_color="white",
            corner_radius=0
        )

        self.receipt_items.pack(
            fill="both",
            expand=True,
            padx=4,
            pady=2
        )

        # ============================================================
        # RECEIPT BOTTOM
        # ============================================================

        ctk.CTkLabel(
            receipt_frame,
            text="-" * 48,
            font=ctk.CTkFont(
                family="Courier New",
                size=10
            ),
            text_color="black"
        ).pack()

        total_receipt_row = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        total_receipt_row.pack(
            fill="x",
            padx=12
        )

        ctk.CTkLabel(
            total_receipt_row,
            text="TOTAL",
            font=preview_bold_font,
            text_color="black"
        ).pack(
            side="left"
        )

        self.receipt_total_label = ctk.CTkLabel(
            total_receipt_row,
            text="0",
            font=preview_bold_font,
            text_color="black"
        )

        self.receipt_total_label.pack(
            side="right"
        )

        cashback_receipt_row = ctk.CTkFrame(
            receipt_frame,
            fg_color="white"
        )

        cashback_receipt_row.pack(
            fill="x",
            padx=12
        )

        ctk.CTkLabel(
            cashback_receipt_row,
            text="Cashback",
            font=preview_font,
            text_color="black"
        ).pack(
            side="left"
        )

        self.receipt_cashback_label = ctk.CTkLabel(
            cashback_receipt_row,
            text="0",
            font=preview_font,
            text_color="black"
        )

        self.receipt_cashback_label.pack(
            side="right"
        )

        ctk.CTkLabel(
            receipt_frame,
            text="Thank You!",
            font=ctk.CTkFont(
                family="Arial",
                size=11,
                weight="bold"
            ),
            text_color="black"
        ).pack(
            pady=(3, 5)
        )

        # ============================================================
        # BOTTOM BUTTONS
        # ============================================================

        button_row = ctk.CTkFrame(
            right_area,
            fg_color="transparent"
        )

        button_row.pack(
            fill="x",
            pady=(8, 0)
        )

        button_row.grid_columnconfigure(
            0,
            weight=1
        )

        button_row.grid_columnconfigure(
            1,
            weight=1
        )

        # ============================================================
        # CLEAR BUTTON
        # ============================================================

        self.clear_button = ctk.CTkButton(
            button_row,
            text="Clear Bill",
            height=44,
            corner_radius=9,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.clear_button.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 5)
        )

        # ============================================================
        # PRINT BUTTON
        # ============================================================

        self.print_button = ctk.CTkButton(
            button_row,
            text="Print",
            height=44,
            corner_radius=9,
            fg_color="#16a34a",
            hover_color="#15803d",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            )
        )

        self.print_button.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(5, 0)
        )

        # ============================================================
        # PAYMENT CALCULATION
        # ============================================================

        def update_payment(event=None):

            try:

                if self.pay_entry.get() == "":

                    self.user_pay = 0

                else:

                    self.user_pay = int(
                        self.pay_entry.get()
                    )

                self.cashback = (
                    self.user_pay -
                    self.total_amount
                )

                self.cashback_label.configure(
                    text=f"{self.cashback:,}"
                )

                self.receipt_cashback_label.configure(
                    text=f"{self.cashback:,}"
                )

            except:

                self.user_pay = 0
                self.cashback = 0

                self.cashback_label.configure(
                    text="0"
                )

                self.receipt_cashback_label.configure(
                    text="0"
                )

        self.pay_entry.bind(
            "<KeyRelease>",
            update_payment
        )

        # ============================================================
        # REMOVE ITEM
        # ============================================================

        def remove_item(item):

            if item in self.cart_items:

                self.cart_items.remove(
                    item
                )

            self.calculate_total()

            refresh_receipt()

            self.bill_saved = False

            self.item_entry.focus()

        # ============================================================
        # REFRESH RECEIPT
        # ============================================================

        def refresh_receipt():

            for widget in (
                self.receipt_items.winfo_children()
            ):

                widget.destroy()

            for index, item in enumerate(
                self.cart_items,
                start=1
            ):

                row = ctk.CTkFrame(
                    self.receipt_items,
                    fg_color="white"
                )

                row.pack(
                    fill="x",
                    padx=2,
                    pady=2
                )

                ctk.CTkLabel(
                    row,
                    text=str(index),
                    width=30,
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                name = str(
                    item["name"]
                )

                if len(name) > 16:

                    name = name[:16]

                ctk.CTkLabel(
                    row,
                    text=name,
                    width=105,
                    anchor="w",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["quantity"]
                    ),
                    width=35,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["price"]
                    ),
                    width=55,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                ctk.CTkLabel(
                    row,
                    text=str(
                        item["amount"]
                    ),
                    width=65,
                    anchor="e",
                    font=preview_font,
                    text_color="black"
                ).pack(
                    side="left"
                )

                remove_button = ctk.CTkButton(
                    row,
                    text="X",
                    width=30,
                    height=25,
                    corner_radius=5,
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    font=ctk.CTkFont(
                        family="Arial",
                        size=10,
                        weight="bold"
                    ),
                    command=lambda
                    current_item=item:
                    remove_item(
                        current_item
                    )
                )

                remove_button.pack(
                    side="left",
                    padx=(5, 0)
                )

            self.total_label.configure(
                text=f"{self.total_amount:,}"
            )

            self.receipt_total_label.configure(
                text=f"{self.total_amount:,}"
            )

            update_payment()

        # ============================================================
        # CALCULATE TOTAL
        # ============================================================

        def calculate_total():

            self.total_amount = 0

            for item in self.cart_items:

                self.total_amount += (
                    item["amount"]
                )

        self.calculate_total = calculate_total

        # ============================================================
        # ADD ITEM
        # ============================================================

        def add_item(event=None):

            ItemCode = (
                self.item_entry
                .get()
                .strip()
            )

            QtyText = (
                self.qty_entry
                .get()
                .strip()
            )

            if ItemCode == "":

                messagebox.showwarning(
                    "Warning",
                    "Please enter Item Code."
                )

                self.item_entry.focus()

                return

            if QtyText == "":

                QtyText = "1"

            Qty = int(
                QtyText
            )

            if Qty <= 0:

                messagebox.showwarning(
                    "Warning",
                    "Quantity must be greater than 0."
                )

                self.qty_entry.focus()

                return

            rows = self.dbWarehouse.Search(
                ItemCode
            )

            if not rows:

                messagebox.showerror(
                    "Item Not Found",
                    "Item Code does not exist.ဤပစွည်းမရှိသေးပါရှင့်။"
                )

                self.item_entry.focus()

                return

            item = rows[0]

            DBQuantity = int(
                item[4]
            )

            Name = item[2]

            Price = int(
                item[6]
            )

            already_in_cart = 0

            for cart_item in self.cart_items:

                if (
                    cart_item["itemcode"]
                    == ItemCode
                ):

                    already_in_cart = (
                        cart_item["quantity"]
                    )

                    break

            total_requested = (
                already_in_cart +
                Qty
            )

            if total_requested > DBQuantity:

                messagebox.showwarning(
                    "Not Enough Stock, ပစွည်းကုန်သွားပါပြီရှင့်",
                    "Available quantity, ဝယ်ယူနိုင်သောအရေအတွက် = "
                    + str(DBQuantity+"")
                )

                return

            found = False

            for cart_item in self.cart_items:

                if (
                    cart_item["itemcode"]
                    == ItemCode
                ):

                    cart_item["quantity"] += Qty

                    cart_item["amount"] = (
                        cart_item["quantity"]
                        * Price
                    )

                    found = True

                    break

            if not found:

                self.cart_items.append(
                    {
                        "itemcode": ItemCode,
                        "name": Name,
                        "quantity": Qty,
                        "price": Price,
                        "amount": Qty * Price
                    }
                )

            calculate_total()

            refresh_receipt()

            # New item starts with quantity 1
            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.item_entry.focus()

            self.bill_saved = False

        self.add_item_button.configure(
            command=add_item
        )

        # Enter on Item Code
        self.item_entry.bind(
            "<Return>",
            add_item
        )

        # Enter on Quantity
        self.qty_entry.bind(
            "<Return>",
            add_item
        )

        # ============================================================
        # SAVE BILL TO DATABASE
        # ============================================================

        def save_bill():

            if len(self.cart_items) == 0:

                messagebox.showwarning(
                    "Warning",
                    "Please add items before saving."
                )

                return False

            if self.bill_saved:

                messagebox.showinfo(
                    "Already Saved",
                    "This bill has already been saved."
                )

                return True

            seller = self.seller_name

            if seller == "":

                seller = getattr(
                    controller,
                    "current_username",
                    ""
                )

            if seller == "":

                messagebox.showwarning(
                    "Seller",
                    "Seller username is empty."
                )

                return False

            # --------------------------------------------------------
            # USER PAY
            # --------------------------------------------------------

            if self.pay_entry.get() == "":

                messagebox.showwarning(
                    "Payment",
                    "Please enter User Pay."
                )

                self.pay_entry.focus()

                return False

            try:

                user_pay = int(
                    self.pay_entry.get()
                )

            except:

                messagebox.showwarning(
                    "Payment",
                    "Please enter a valid amount."
                )

                self.pay_entry.focus()

                return False

            if user_pay < self.total_amount:

                messagebox.showwarning(
                    "Payment Error",
                    "User Pay is less than Total."
                )

                self.pay_entry.focus()

                return False

            self.user_pay = user_pay

            self.cashback = (
                user_pay -
                self.total_amount
            )

            update_payment()

            # --------------------------------------------------------
            # DATE
            # --------------------------------------------------------

            now = datetime.now()

            bill_date = now.strftime(
                "%Y-%m-%d"
            )

            try:

                # ====================================================
                # SAVE EACH ITEM
                # ====================================================

                for item in self.cart_items:

                    ItemCode = item["itemcode"]

                    Name = item["name"]

                    Qty = item["quantity"]

                    Price = item["price"]

                    Amount = item["amount"]

                    # ----------------------------------------------
                    # AddTotal
                    # ----------------------------------------------

                    self.dbTotal.insert(
                        ItemCode,
                        Name,
                        str(Amount)
                    )

                    # ----------------------------------------------
                    # AddTotalBillNum
                    # ----------------------------------------------

                    self.dbTotal.insertBill(
                        self.bill_number,
                        ItemCode,
                        Name,
                        Qty,
                        Price,
                        str(Amount),
                        seller
                    )

                    # ----------------------------------------------
                    # SalesReport
                    # ----------------------------------------------

                    self.dbSales.insert(
                        bill_date,
                        Amount,
                        self.bill_number,
                        seller,
                        Qty,
                        ItemCode,
                        Name
                    )

                    # ----------------------------------------------
                    # Warehouse quantity decrease
                    # ----------------------------------------------

                    current = (
                        self.dbWarehouse
                        .getQuntity(
                            ItemCode
                        )
                    )

                    if current:

                        current_qty = int(
                            current[0]
                        )

                        new_quantity = (
                            current_qty -
                            Qty
                        )

                        if new_quantity < 0:

                            new_quantity = 0

                        self.dbWarehouse.updateQuantity(
                            ItemCode,
                            new_quantity
                        )

                # ====================================================
                # SAVED
                # ====================================================

                self.bill_saved = True

                update_today_sales()

                messagebox.showinfo(
                    "Saved",
                    "Bill saved successfully.\n\n"
                    "You can now press Print."
                )

                return True

            except Exception as e:

                messagebox.showerror(
                    "Database Error",
                    str(e)
                )

                return False

        # ============================================================
        # USER PAY ENTER = SAVE ONLY
        # ============================================================

        self.pay_entry.bind(
            "<Return>",
            lambda event: save_bill()
        )

        # ============================================================
        # CLEAR BILL
        # ============================================================

        def clear_bill():

            if (
                len(self.cart_items) == 0
                and not self.bill_saved
            ):

                return

            answer = messagebox.askyesno(
                "Clear Bill",
                "Are you sure you want to "
                "clear this bill?"
            )

            if not answer:

                return

            self.cart_items.clear()

            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0

            self.bill_saved = False

            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.pay_entry.delete(
                0,
                "end"
            )

            self.total_label.configure(
                text="0"
            )

            self.cashback_label.configure(
                text="0"
            )

            self.receipt_total_label.configure(
                text="0"
            )

            self.receipt_cashback_label.configure(
                text="0"
            )

            refresh_receipt()

            self.item_entry.focus()

        self.clear_button.configure(
            command=clear_bill
        )

        # ============================================================
        # PRINT BILL
        # ============================================================

        def print_bill():

            if len(self.cart_items) == 0:

                messagebox.showwarning(
                    "Warning",
                    "There are no items to print."
                )

                return

            # --------------------------------------------------------
            # If not saved, ask user to save first
            # --------------------------------------------------------

            if not self.bill_saved:

                answer = messagebox.askyesno(
                    "Bill Not Saved",
                    "This bill has not been saved.\n\n"
                    "Save it before printing?"
                )

                if not answer:

                    return

                if not save_bill():

                    return

            seller = self.seller_name

            if seller == "":

                seller = getattr(
                    controller,
                    "current_username",
                    ""
                )

            now = datetime.now()

            bill_datetime = now.strftime(
                "%Y-%m-%d %H:%M"
            )

            # ========================================================
            # 80MM PRINTER TEXT
            #
            # This remains smaller / fixed-width for the printer.
            # The preview above uses a larger font.
            # ========================================================

            receipt_lines = []

            receipt_lines.append(
                "အောင်ဘုန်းခန့် MiniMart"
            )

            receipt_lines.append(
                "Ph:09777775706"
            )

            receipt_lines.append(
                "Bill Number: "
                + self.bill_number
            )

            receipt_lines.append(
                "Date: "
                + bill_datetime
            )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                "No  Name              "
                "Qty Price  Amount"
            )

            receipt_lines.append(
                "-" * 48
            )

            for index, item in enumerate(
                self.cart_items,
                start=1
            ):

                name = str(
                    item["name"]
                )

                if len(name) > 17:

                    name = name[:17]

                line = (
                    f"{index:<3}"
                    f"{name:<18}"
                    f"{item['quantity']:>4}"
                    f"{item['price']:>6}"
                    f"{item['amount']:>9}"
                )

                receipt_lines.append(
                    line
                )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                f"{'TOTAL':<30}"
                f"{self.total_amount:>18}"
            )

            receipt_lines.append(
                f"{'PAY':<30}"
                f"{self.user_pay:>18}"
            )

            receipt_lines.append(
                f"{'CASHBACK':<30}"
                f"{self.cashback:>18}"
            )

            receipt_lines.append(
                "-" * 48
            )

            receipt_lines.append(
                "Seller: "
                + seller
            )

            receipt_lines.append(
                "Thank You!"
            )

            receipt_text = "\n".join(
                receipt_lines
            )

            # ========================================================
            # TEMP FILE
            # ========================================================

            receipt_file = os.path.join(
                tempfile.gettempdir(),
                "ABK_Bill_"
                + self.bill_number
                + ".txt"
            )

            try:

                with open(
                    receipt_file,
                    "w",
                    encoding="utf-8"
                ) as file:

                    file.write(
                        receipt_text
                    )

            except Exception as e:

                messagebox.showerror(
                    "File Error",
                    str(e)
                )

                return

            # ========================================================
            # WINDOWS PRINT
            # ========================================================

            try:

                os.startfile(
                    receipt_file,
                    "print"
                )

            except Exception as e:

                messagebox.showerror(
                    "Printer Error",
                    "Windows could not send "
                    "the receipt to printer.\n\n"
                    + str(e)
                )

                return

            # ========================================================
            # PRINT SUCCESS
            # ========================================================

            messagebox.showinfo(
                "Print",
                "Bill sent to printer."
            )

            # ========================================================
            # START NEW BILL
            # ========================================================

            self.cart_items.clear()

            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0

            self.bill_saved = False

            self.item_entry.delete(
                0,
                "end"
            )

            self.qty_entry.delete(
                0,
                "end"
            )

            self.qty_entry.insert(
                0,
                "1"
            )

            self.pay_entry.delete(
                0,
                "end"
            )

            # ========================================================
            # NEW BILL NUMBER
            # ========================================================

            self.bill_number = str(
                random.randint(
                    100000,
                    999999
                )
            )

            self.bill_number_label.configure(
                text="#"
                + self.bill_number
            )

            self.receipt_bill_label.configure(
                text="Bill Number: "
                + self.bill_number
            )

            self.receipt_date_label.configure(
                text="Date: "
                + datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )
            )

            self.total_label.configure(
                text="0"
            )

            self.cashback_label.configure(
                text="0"
            )

            self.receipt_total_label.configure(
                text="0"
            )

            self.receipt_cashback_label.configure(
                text="0"
            )

            refresh_receipt()

            update_today_sales()

            self.item_entry.focus()

        self.print_button.configure(
            command=print_bill
        )

        # ============================================================
        # INITIAL
        # ============================================================

        refresh_receipt()

        update_today_sales()

        self.item_entry.focus()

    def refresh_seller(self):

        username = getattr(
            self.controller,
            "current_username",
            ""
        )

        self.seller_name = username

        self.seller_entry.configure(
            state="normal"
        )

        self.seller_entry.delete(
            0,
            "end"
        )

        if username:
            self.seller_entry.insert(
                0,
                username
            )

        self.seller_entry.configure(
            state="readonly"


        )


    def tkraise(self, aboveThis=None):

        result = super().tkraise(aboveThis)

        # Update seller every time BillPage is opened
        if hasattr(self, "seller_entry"):
            self.refresh_seller()

        return result        

# ============================================================
# FIVE PAGE - ADD ITEM / ITEM MANAGEMENT
# ============================================================

class FivePage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#f3f5f9"
        )

        self.controller = controller
        self.selected_id = None

        # ====================================================
        # DATABASE
        # ====================================================

        dbWarehouse = warehouse(
            "real/Warehousedatabase.db"
        )

        # ====================================================
        # NUMBER VALIDATION
        # Quantity / Cost / Price = numbers only
        # ====================================================

        def only_numbers(value):

            if value == "":
                return True

            return value.isdigit()

        number_command = self.register(
            only_numbers
        )

        # ====================================================
        # LEFT SIDEBAR
        # ====================================================

        self.sidebar = ctk.CTkFrame(
            self,
            width=230,
            fg_color="#1e293b",
            corner_radius=0
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(False)

        # ====================================================
        # LOGO
        # ====================================================

        ctk.CTkLabel(
            self.sidebar,
            text="ABK",
            font=ctk.CTkFont(
                family="Arial",
                size=38,
                weight="bold"
            ),
            text_color="#ffffff"
        ).pack(
            pady=(35, 5)
        )

        ctk.CTkLabel(
            self.sidebar,
            text="POS SYSTEM",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#fbbf24"
        ).pack(
            pady=(0, 30)
        )

        # ====================================================
        # SIDEBAR BUTTON FUNCTION
        # ====================================================

        def sidebar_button(
            text,
            page,
            fg_color="#1e293b"
        ):

            button = ctk.CTkButton(
                self.sidebar,
                text=text,
                height=50,
                corner_radius=10,
                anchor="w",
                font=ctk.CTkFont(
                    family="Arial",
                    size=15,
                    weight="bold"
                ),
                fg_color=fg_color,
                hover_color="#334155",
                text_color="#ffffff",
                command=lambda: controller.show_frame(page)
            )

            button.pack(
                fill="x",
                padx=20,
                pady=5
            )

            return button

        # ====================================================
        # SIDEBAR MENU
        # ====================================================
        sidebar_button(
            "   Dashboard",
            SecondPage
        )
        
        sidebar_button(
            "   Bill",
            BillPage
        )

        sidebar_button(
            "   Add Item",
            FivePage,
            fg_color="#2563eb"
        )

        #sidebar_button(
        #    "   Barcode",
        #    BarcodePage
        #)
#
        #sidebar_button(
        #    "   Warehouse",
        #    WarehousePage
        #)

        sidebar_button(
            "   Report",
            ReportPage
        )

        sidebar_button(
            "   Setting",
            SettingPage
        )

        sidebar_button(
            "   Logout",
            FirstPage,
            fg_color="#dc2626"
        )

        # ====================================================
        # MAIN AREA
        # ====================================================

        self.main_area = ctk.CTkFrame(
            self,
            fg_color="#f3f5f9",
            corner_radius=0
        )

        self.main_area.pack(
            side="left",
            fill="both",
            expand=True
        )

        # ====================================================
        # TOP BAR
        # ====================================================

        self.topbar = ctk.CTkFrame(
            self.main_area,
            height=80,
            fg_color="#ffffff",
            corner_radius=0
        )

        self.topbar.pack(
            fill="x"
        )

        self.topbar.pack_propagate(False)

        ctk.CTkLabel(
            self.topbar,
            text="Add Item (Item Management System)",
            font=ctk.CTkFont(
                family="Arial",
                size=27,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left",
            padx=30
        )

        ctk.CTkLabel(
            self.topbar,
            text="Admin",
            font=ctk.CTkFont(
                family="Arial",
                size=14,
                weight="bold"
            ),
            text_color="#64748b"
        ).pack(
            side="right",
            padx=30
        )

        # ====================================================
        # FULL ITEM MANAGEMENT SCROLL AREA
        # ====================================================

        self.content = ctk.CTkScrollableFrame(
            self.main_area,
            fg_color="#f3f5f9",
            corner_radius=0
        )

        self.content.pack(
            fill="both",
            expand=True
        )

        # ====================================================
        # TITLE
        # ====================================================
        #ctk.CTkLabel(
        #    self.content,
        #    text="Item Management",
        #    font=ctk.CTkFont(
        #        family="Arial",
        #        size=30,
        #        weight="bold"
        #    ),
        #    text_color="#111827"
        #).pack(
        #    anchor="w",
        #    padx=30,
        #    pady=(25, 0)
        #)

        #ctk.CTkLabel(
        #    self.content,
        #    text="Add, search, update and delete warehouse items.",
        #    font=ctk.CTkFont(
        #        family="Arial",
        #        size=3
        #    ),
        #    text_color="#64748b"
        #).pack(
        #    anchor="w",
        #    padx=30,
        #    pady=(1, 2)
        #)

        # ====================================================
        # FORM CARD
        # ====================================================

        self.form_card = ctk.CTkFrame(
            self.content,
            fg_color="#B4B1B1",
            corner_radius=18
        )

        self.form_card.pack(
            fill="x",
            padx=30
        )

        # ====================================================
        # FORM SCROLL
        # ====================================================

        self.form_scroll = ctk.CTkFrame(
            self.form_card,
            
            fg_color="#ffffff",
            corner_radius=18,
            height=500
        )

        self.form_scroll.pack(
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

        # ====================================================
        # FORM
        # ====================================================

        self.form = ctk.CTkFrame(
            self.form_scroll,
            fg_color="transparent"
        )

        self.form.pack(
            fill="x",
            padx=30,
            pady=20
        )

        self.form.grid_columnconfigure(
            0,
            weight=1
        )

        self.form.grid_columnconfigure(
            1,
            weight=1
        )

        # ====================================================
        # ITEM CODE
        # ====================================================

        ctk.CTkLabel(
            self.form,
            text="Item Code",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=(0, 8)
        )

        ctk.CTkLabel(
            self.form,
            text="Item Name",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=0,
            column=1,
            sticky="w",
            padx=(15, 0),
            pady=(0, 8)
        )

        self.item_code = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter item code",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            )
        )

        self.item_code.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=(0, 15),
            pady=(0, 20)
        )

        self.item_name = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter item name",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            )
        )

        self.item_name.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(15, 0),
            pady=(0, 20)
        )

        # ====================================================
        # CATEGORY / QUANTITY
        # ====================================================

        ctk.CTkLabel(
            self.form,
            text="Category",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=(0, 8)
        )

        ctk.CTkLabel(
            self.form,
            text="Quantity",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(15, 0),
            pady=(0, 8)
        )

        self.category = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter category",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            )
        )

        self.category.grid(
            row=3,
            column=0,
            sticky="ew",
            padx=(0, 15),
            pady=(0, 20)
        )

        self.quantity = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter quantity",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.quantity.grid(
            row=3,
            column=1,
            sticky="ew",
            padx=(15, 0),
            pady=(0, 20)
        )

        # ====================================================
        # COST / PRICE
        # ====================================================

        ctk.CTkLabel(
            self.form,
            text="Cost",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=4,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=(0, 8)
        )

        ctk.CTkLabel(
            self.form,
            text="Price",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=4,
            column=1,
            sticky="w",
            padx=(15, 0),
            pady=(0, 8)
        )

        self.cost = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter cost",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.cost.grid(
            row=5,
            column=0,
            sticky="ew",
            padx=(0, 15),
            pady=(0, 20)
        )

        self.price = ctk.CTkEntry(
            self.form,
            height=48,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Enter selling price",
            font=ctk.CTkFont(
                family="Arial",
                size=15
            ),
            validate="key",
            validatecommand=(
                number_command,
                "%P"
            )
        )

        self.price.grid(
            row=5,
            column=1,
            sticky="ew",
            padx=(15, 0),
            pady=(0, 20)
        )

        # ====================================================
        # EXPIRED DATE / ALERT DATE
        # ====================================================

        ctk.CTkLabel(
            self.form,
            text="Expired Date",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=6,
            column=0,
            sticky="w",
            padx=(0, 15),
            pady=(0, 8)
        )

        ctk.CTkLabel(
            self.form,
            text="Alert Date",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#334155",
            anchor="w"
        ).grid(
            row=6,
            column=1,
            sticky="w",
            padx=(15, 0),
            pady=(0, 8)
        )

        self.exp_date = DateEntry(
            self.form,
            width=20,
            background="#2563eb",
            foreground="white",
            borderwidth=0,
            date_pattern="yyyy-mm-dd",
            font=(
                "Arial",
                14
            )
        )

        self.exp_date.grid(
            row=7,
            column=0,
            sticky="ew",
            padx=(0, 15),
            pady=(0, 20),
            ipady=10
        )

        self.alert_date = DateEntry(
            self.form,
            width=20,
            background="#f59e0b",
            foreground="white",
            borderwidth=0,
            date_pattern="yyyy-mm-dd",
            font=(
                "Arial",
                14
            )
        )

        self.alert_date.grid(
            row=7,
            column=1,
            sticky="ew",
            padx=(15, 0),
            pady=(0, 20),
            ipady=10
        )

        # ====================================================
        # CLEAR FUNCTION
        # ====================================================

        def clear_form():

            self.item_code.delete(0, "end")
            self.item_name.delete(0, "end")
            self.category.delete(0, "end")
            self.quantity.delete(0, "end")
            self.cost.delete(0, "end")
            self.price.delete(0, "end")
        
            self.exp_date.set_date(date.today())
            self.alert_date.set_date(date.today())
        
            self.selected_id = None
        
            for selected in self.item_table.selection():
                self.item_table.selection_remove(selected)

        # ====================================================
        # SAVE FUNCTION
        # ====================================================

        def save_item(event=None):

            Item = self.item_code.get().strip()
            Name = self.item_name.get().strip()
            Category = self.category.get().strip()
            Quantity = self.quantity.get().strip()
            Cost = self.cost.get().strip()
            Price = self.price.get().strip()

            Exp_Date = self.exp_date.get()
            Alert_Date = self.alert_date.get()

            # ------------------------------------------------
            # VALIDATION
            # ------------------------------------------------

            if Item == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter item code."
                )

                self.item_code.focus()

                return

            if Name == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter item name."
                )

                self.item_name.focus()

                return

            if Category == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter category."
                )

                self.category.focus()

                return

            if Quantity == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter quantity."
                )

                self.quantity.focus()

                return

            if Cost == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter cost."
                )

                self.cost.focus()

                return

            if Price == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter price."
                )

                self.price.focus()

                return

            # ------------------------------------------------
            # DATABASE INSERT
            # ------------------------------------------------

            try:

                dbWarehouse.insert(
                    Item,
                    Name,
                    Category,
                    int(Quantity),
                    Cost,
                    int(Price),
                    Exp_Date,
                    Alert_Date
                )

                messagebox.showinfo(
                    "Success",
                    "Item added successfully."
                )

                clear_form()

                load_items()

            except Exception as e:

                messagebox.showerror(
                    "Database Error",
                    str(e)
                )

        # ====================================================
        # ENTER KEY
        #
        # Alert Date + ENTER
        #       ↓
        # save_item()
        # ====================================================

        self.alert_date.bind(
            "<Return>",
            save_item
        )

        # ====================================================
        # FORM BUTTON AREA
        # ====================================================

        self.button_frame = ctk.CTkFrame(
            self.form,
            fg_color="transparent"
        )

        self.button_frame.grid(
            row=8,
            column=0,
            columnspan=2,
            sticky="e",
            pady=(0, 10)
        )

        # ====================================================
        # CLEAR BUTTON
        # ====================================================

        ctk.CTkButton(
            self.button_frame,
            text="Clear",
            width=120,
            height=45,
            corner_radius=10,
            fg_color="#64748b",
            hover_color="#475569",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            command=clear_form
        ).pack(
            side="left",
            padx=5
        )

        # ====================================================
        # SAVE BUTTON
        # ====================================================

        self.save_button = ctk.CTkButton(
            self.button_frame,
            text="Save Item",
            width=150,
            height=45,
            corner_radius=10,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            command=save_item
        )

        self.save_button.pack(
            side="left",
            padx=5
        )

        # ====================================================
        # ITEMS DISPLAY CARD
        # ====================================================

        self.table_card = ctk.CTkFrame(
            self.content,
            fg_color="#9f9d9d",
            corner_radius=18
        )

        self.table_card.pack(
            fill="x",
            padx=30,
            pady=(20, 30)
        )

        # ====================================================
        # TABLE HEADER
        # ====================================================

        self.table_header = ctk.CTkFrame(
            self.table_card,
            fg_color="transparent"
        )

        self.table_header.pack(
            fill="x",
            padx=25,
            pady=(20, 10)
        )

        # ====================================================
        # ADDED ITEMS TITLE
        # ====================================================

        ctk.CTkLabel(
            self.table_header,
            text="Added Items",
            font=ctk.CTkFont(
                family="Arial",
                size=21,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left"
        )

        # ====================================================
        # RIGHT CONTROL AREA
        # Search / Update / Delete
        # ====================================================

        self.table_controls = ctk.CTkFrame(
            self.table_header,
            fg_color="transparent"
        )

        self.table_controls.pack(
            side="right"
        )

        # ====================================================
        # SEARCH ENTRY
        # ====================================================

        self.search_entry = ctk.CTkEntry(
            self.table_controls,
            width=210,
            height=40,
            corner_radius=10,
            border_width=1,
            border_color="#6788b1",
            fg_color="#f8fafc",
            placeholder_text="Search Item Code...",
            font=ctk.CTkFont(
                family="Arial",
                size=13
            )
        )

        self.search_entry.pack(
            side="left",
            padx=4
        )

        # ====================================================
        # SEARCH FUNCTION
        # ====================================================

        def search_item():

            search_value = self.search_entry.get().strip()

            # Clear table

            for row in self.item_table.get_children():

                self.item_table.delete(
                    row
                )

            # Search database

            if search_value == "":

                rows = dbWarehouse.fetch()

            else:

                rows = dbWarehouse.Search(
                    search_value
                )

            # Display result

            for row in rows:

                self.item_table.insert(
                    "",
                    "end",
                    values=row
                )

        # ====================================================
        # SEARCH BUTTON
        # ====================================================

        ctk.CTkButton(
            self.table_controls,
            text="Search",
            width=90,
            height=40,
            corner_radius=10,
            fg_color="#8b5cf6",
            hover_color="#7c3aed",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=search_item
        ).pack(
            side="left",
            padx=4
        )

        # ====================================================
        # UPDATE FUNCTION
        # ====================================================

        def update_item():

            if self.selected_id is None:

                messagebox.showwarning(
                    "Warning",
                    "Please select an item from the table first."
                )

                return

            Item = self.item_code.get().strip()
            Name = self.item_name.get().strip()
            Category = self.category.get().strip()
            Quantity = self.quantity.get().strip()
            Cost = self.cost.get().strip()
            Price = self.price.get().strip()

            Exp_Date = self.exp_date.get()
            Alert_Date = self.alert_date.get()

            # ------------------------------------------------
            # VALIDATION
            # ------------------------------------------------

            if Item == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter item code."
                )
                return

            if Name == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter item name."
                )
                return

            if Category == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter category."
                )
                return

            if Quantity == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter quantity."
                )
                return

            if Cost == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter cost."
                )
                return

            if Price == "":
                messagebox.showwarning(
                    "Warning",
                    "Please enter price."
                )
                return

            # ------------------------------------------------
            # UPDATE DATABASE
            # ------------------------------------------------

            try:

                dbWarehouse.update(
                    self.selected_id,
                    Item,
                    Name,
                    Category,
                    int(Quantity),
                    Cost,
                    int(Price),
                    Exp_Date,
                    Alert_Date
                )

                messagebox.showinfo(
                    "Success",
                    "Item updated successfully."
                )

                clear_form()

                load_items()

            except Exception as e:

                messagebox.showerror(
                    "Database Error",
                    str(e)
                )

        # ====================================================
        # UPDATE BUTTON
        # ====================================================

        ctk.CTkButton(
            self.table_controls,
            text="Update",
            width=90,
            height=40,
            corner_radius=10,
            fg_color="#f59e0b",
            hover_color="#d97706",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=update_item
        ).pack(
            side="left",
            padx=4
        )

        # ====================================================
        # DELETE FUNCTION
        # ====================================================

        def delete_item():

            if self.selected_id is None:

                messagebox.showwarning(
                    "Warning",
                    "Please select an item from the table first."
                )

                return

            confirm = messagebox.askyesno(
                "Delete Item",
                "Are you sure you want to delete this item?"
            )

            if not confirm:

                return

            try:

                dbWarehouse.remove(
                    self.selected_id
                )

                messagebox.showinfo(
                    "Success",
                    "Item deleted successfully."
                )

                clear_form()

                load_items()

            except Exception as e:

                messagebox.showerror(
                    "Database Error",
                    str(e)
                )

        # ====================================================
        # DELETE BUTTON
        # ====================================================

        ctk.CTkButton(
            self.table_controls,
            text="Delete",
            width=90,
            height=40,
            corner_radius=10,
            fg_color="#ef4444",
            hover_color="#dc2626",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=delete_item
        ).pack(
            side="left",
            padx=4
        )

        # ====================================================
        # TABLE FRAME
        # ====================================================

        self.table_frame = ctk.CTkFrame(
            self.table_card,
            fg_color="transparent"
        )

        self.table_frame.pack(
            fill="both",
            expand=True,
            padx=25,
            pady=(5, 25)
        )

        # ====================================================
        # TREEVIEW STYLE
        # ====================================================

        style = ttk.Style()

        try:

            style.theme_use(
                "clam"
            )

        except:

            pass

        style.configure(
            "Treeview",
            rowheight=35,
            font=(
                "Arial",
                10
            ),
            background="#ffffff",
            fieldbackground="#ffffff"
        )

        style.configure(
            "Treeview.Heading",
            font=(
                "Arial",
                10,
                "bold"
            ),
            padding=8
        )

        style.map(
            "Treeview",
            background=[
                (
                    "selected",
                    "#2563eb"
                )
            ],
            foreground=[
                (
                    "selected",
                    "#ffffff"
                )
            ]
        )

        # ====================================================
        # TABLE COLUMNS
        # ====================================================

        columns = (
            "ID",
            "Item",
            "Name",
            "Category",
            "Quantity",
            "Cost",
            "Price",
            "Exp_Date",
            "Alert_Date"
        )

        self.item_table = ttk.Treeview(
            self.table_frame,
            columns=columns,
            show="headings",
            height=8
        )

        # ====================================================
        # HEADINGS
        # ====================================================

        self.item_table.heading(
            "ID",
            text="ID"
        )

        self.item_table.heading(
            "Item",
            text="Item Code"
        )

        self.item_table.heading(
            "Name",
            text="Name"
        )

        self.item_table.heading(
            "Category",
            text="Category"
        )

        self.item_table.heading(
            "Quantity",
            text="Quantity"
        )

        self.item_table.heading(
            "Cost",
            text="Cost"
        )

        self.item_table.heading(
            "Price",
            text="Price"
        )

        self.item_table.heading(
            "Exp_Date",
            text="Expired Date"
        )

        self.item_table.heading(
            "Alert_Date",
            text="Alert Date"
        )

        # ====================================================
        # COLUMN WIDTH
        # ====================================================

        self.item_table.column(
            "ID",
            width=50,
            anchor="center"
        )

        self.item_table.column(
            "Item",
            width=110,
            anchor="center"
        )

        self.item_table.column(
            "Name",
            width=150,
            anchor="w"
        )

        self.item_table.column(
            "Category",
            width=120,
            anchor="w"
        )

        self.item_table.column(
            "Quantity",
            width=90,
            anchor="center"
        )

        self.item_table.column(
            "Cost",
            width=100,
            anchor="center"
        )

        self.item_table.column(
            "Price",
            width=100,
            anchor="center"
        )

        self.item_table.column(
            "Exp_Date",
            width=120,
            anchor="center"
        )

        self.item_table.column(
            "Alert_Date",
            width=120,
            anchor="center"
        )

        # ====================================================
        # TABLE SCROLLBAR
        # ====================================================

        table_scrollbar = ttk.Scrollbar(
            self.table_frame,
            orient="vertical",
            command=self.item_table.yview
        )

        self.item_table.configure(
            yscrollcommand=table_scrollbar.set
        )

        self.item_table.pack(
            side="left",
            fill="both",
            expand=True
        )

        table_scrollbar.pack(
            side="right",
            fill="y"
        )

        # ====================================================
        # LOAD ITEMS
        # ====================================================

        def load_items():

            for row in self.item_table.get_children():

                self.item_table.delete(
                    row
                )

            rows = dbWarehouse.fetch()

            for row in rows:

                self.item_table.insert(
                    "",
                    "end",
                    values=row
                )

        # ====================================================
        # SELECT TABLE ROW
        # ====================================================

        def select_item(event=None):

            selected = self.item_table.selection()

            if not selected:

                return

            row = self.item_table.item(
                selected[0]
            )

            values = row["values"]

            if not values:

                return

            # ------------------------------------------------
            # Save selected ID
            # ------------------------------------------------

            self.selected_id = values[0]

            # ------------------------------------------------
            # Item Code
            # ------------------------------------------------

            self.item_code.delete(
                0,
                "end"
            )

            self.item_code.insert(
                0,
                values[1]
            )

            # ------------------------------------------------
            # Item Name
            # ------------------------------------------------

            self.item_name.delete(
                0,
                "end"
            )

            self.item_name.insert(
                0,
                values[2]
            )

            # ------------------------------------------------
            # Category
            # ------------------------------------------------

            self.category.delete(
                0,
                "end"
            )

            self.category.insert(
                0,
                values[3]
            )

            # ------------------------------------------------
            # Quantity
            # ------------------------------------------------

            self.quantity.delete(
                0,
                "end"
            )

            self.quantity.insert(
                0,
                values[4]
            )

            # ------------------------------------------------
            # Cost
            # ------------------------------------------------

            self.cost.delete(
                0,
                "end"
            )

            self.cost.insert(
                0,
                values[5]
            )

            # ------------------------------------------------
            # Price
            # ------------------------------------------------

            self.price.delete(
                0,
                "end"
            )

            self.price.insert(
                0,
                values[6]
            )

            # ------------------------------------------------
            # Expired Date
            # ------------------------------------------------

            try:

                self.exp_date.set_date(
                    values[7]
                )

            except:

                pass

            # ------------------------------------------------
            # Alert Date
            # ------------------------------------------------

            try:

                self.alert_date.set_date(
                    values[8]
                )

            except:

                pass

            # ------------------------------------------------
            # Scroll form to top
            # ------------------------------------------------

            try:

                self.form_scroll._parent_canvas.yview_moveto(
                    0
                )

            except:

                pass

        self.item_table.bind(
            "<<TreeviewSelect>>",
            select_item
        )

        # ====================================================
        # SEARCH ENTER KEY
        # ====================================================

        self.search_entry.bind(
            "<Return>",
            lambda event: search_item()
        )

        # ====================================================
        # INITIAL DATABASE LOAD
        # ====================================================

        load_items()



# ============================================================
# Extra Page
# ============================================================



class BarcodePage(EmptyPOSPage):

    def __init__(self, parent, controller):
        super().__init__(
            parent,
            controller,
            "Barcode"
        )


class WarehousePage(EmptyPOSPage):

    def __init__(self, parent, controller):
        super().__init__(
            parent,
            controller,
            "Warehouse"
        )

# ======================== Report Page ====================

class ReportPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#f3f5f9"
        )

        self.controller = controller

        # ============================================================
        # DATABASE
        # ============================================================

        self.dbSales = DatabaseSales(
            "real/SalesReportdatabase.db"
        )

        self.dbWarehouse = warehouse(
            "real/Warehousedatabase.db"
        )

        # ============================================================
        # MAIN GRID
        # ============================================================

        self.grid_columnconfigure(
            0,
            weight=1
        )

        self.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # MAIN SCROLL
        # ============================================================

        self.main_scroll = ctk.CTkScrollableFrame(
            self,
            fg_color="#f3f5f9"
        )

        self.main_scroll.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        # ============================================================
        # HEADER
        # ============================================================

        header = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent"
        )

        header.pack(
            fill="x",
            padx=15,
            pady=(15, 10)
        )

        header_left = ctk.CTkFrame(
            header,
            fg_color="transparent"
        )

        header_left.pack(
            side="left"
        )

        ctk.CTkLabel(
            header_left,
            text="Sales Report",
            font=ctk.CTkFont(
                family="Arial",
                size=30,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w"
        )

        ctk.CTkLabel(
            header_left,
            text="Sales, orders and inventory report",
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#6b7280"
        ).pack(
            anchor="w"
        )

        ctk.CTkButton(
            header,
            text="Dashboard",
            width=120,
            height=40,
            corner_radius=8,
            fg_color="#111827",
            hover_color="#1f2937",
            command=lambda: controller.show_frame(SecondPage)
        ).pack(
            side="right",
            pady=5
        )

        # ============================================================
        # TOP 3 CARDS
        # ============================================================

        cards_frame = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent"
        )

        cards_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 15)
        )

        for i in range(3):

            cards_frame.grid_columnconfigure(
                i,
                weight=1
            )

        # Today's Sales
        self.sales_card = self.create_summary_card(
            cards_frame,
            0,
            "Today's Sales",
            "0",
            "💰",
            "#dcfce7"
        )

        # Today's Orders
        self.orders_card = self.create_summary_card(
            cards_frame,
            1,
            "Today's Orders",
            "0",
            "🧾",
            "#dbeafe"
        )

        # Total Items
        self.items_card = self.create_summary_card(
            cards_frame,
            2,
            "Total Items",
            "0",
            "📦",
            "#fef3c7"
        )

        # ============================================================
        # SALES REPORT CARD
        # ============================================================

        report_card = ctk.CTkFrame(
            self.main_scroll,
            fg_color="#ffffff",
            corner_radius=14
        )

        report_card.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 15)
        )

        # ============================================================
        # REPORT TITLE
        # ============================================================

        title_frame = ctk.CTkFrame(
            report_card,
            fg_color="transparent"
        )

        title_frame.pack(
            fill="x",
            padx=20,
            pady=(18, 10)
        )

        ctk.CTkLabel(
            title_frame,
            text="Sales Table",
            font=ctk.CTkFont(
                family="Arial",
                size=21,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            side="left"
        )

        # ============================================================
        # FILTER FRAME
        # ============================================================

        filter_frame = ctk.CTkFrame(
            report_card,
            fg_color="#f8fafc",
            corner_radius=10
        )

        filter_frame.pack(
            fill="x",
            padx=20,
            pady=(0, 15)
        )

        # ---------------- DATE ----------------

        ctk.CTkLabel(
            filter_frame,
            text="Date",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#374151"
        ).grid(
            row=0,
            column=0,
            padx=(15, 5),
            pady=15
        )

        self.date_entry = ctk.CTkEntry(
            filter_frame,
            width=150,
            height=38,
            placeholder_text="YYYY-MM-DD"
        )

        self.date_entry.grid(
            row=0,
            column=1,
            padx=5,
            pady=15
        )

        # Enter = search date
        self.date_entry.bind(
            "<Return>",
            lambda event: self.search_by_date()
        )

        # ---------------- SEARCH DATE ----------------

        ctk.CTkButton(
            filter_frame,
            text="Search Date",
            width=110,
            height=38,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self.search_by_date
        ).grid(
            row=0,
            column=2,
            padx=5,
            pady=15
        )

        # ---------------- SEARCH TOTAL ----------------

        ctk.CTkButton(
            filter_frame,
            text="Search Total",
            width=110,
            height=38,
            fg_color="#16a34a",
            hover_color="#15803d",
            command=self.search_total
        ).grid(
            row=0,
            column=3,
            padx=5,
            pady=15
        )

        # ---------------- DAILY ----------------

        ctk.CTkButton(
            filter_frame,
            text="Daily",
            width=90,
            height=38,
            fg_color="#7c3aed",
            hover_color="#6d28d9",
            command=self.show_daily
        ).grid(
            row=0,
            column=4,
            padx=5,
            pady=15
        )

        # ---------------- MONTHLY ----------------

        ctk.CTkButton(
            filter_frame,
            text="Monthly",
            width=90,
            height=38,
            fg_color="#ea580c",
            hover_color="#c2410c",
            command=self.show_monthly
        ).grid(
            row=0,
            column=5,
            padx=5,
            pady=15
        )

        # ---------------- YEARLY ----------------

        ctk.CTkButton(
            filter_frame,
            text="Yearly",
            width=90,
            height=38,
            fg_color="#0891b2",
            hover_color="#0e7490",
            command=self.show_yearly
        ).grid(
            row=0,
            column=6,
            padx=5,
            pady=15
        )

        # ---------------- ALL ----------------

        ctk.CTkButton(
            filter_frame,
            text="All",
            width=80,
            height=38,
            fg_color="#374151",
            hover_color="#1f2937",
            command=self.show_all
        ).grid(
            row=0,
            column=7,
            padx=(5, 15),
            pady=15
        )
        # ---------------- EXPORT EXCEL ----------------

        ctk.CTkButton(
            filter_frame,
            text="Export Excel",
            width=120,
            height=38,
            fg_color="#059669",
            hover_color="#047857",
            command=self.export_excel
        ).grid(
            row=0,
            column=8,
            padx=(5, 15),
            pady=15
        )
        # ============================================================
        # TOTAL RESULT
        # ============================================================

        total_frame = ctk.CTkFrame(
            report_card,
            fg_color="transparent"
        )

        total_frame.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        self.report_title = ctk.CTkLabel(
            total_frame,
            text="All Sales",
            font=ctk.CTkFont(
                family="Arial",
                size=15,
                weight="bold"
            ),
            text_color="#374151"
        )

        self.report_title.pack(
            side="left"
        )

        self.total_label = ctk.CTkLabel(
            total_frame,
            text="Total: 0",
            font=ctk.CTkFont(
                family="Arial",
                size=17,
                weight="bold"
            ),
            text_color="#16a34a"
        )

        self.total_label.pack(
            side="right"
        )

        # ============================================================
        # TABLE FRAME
        # ============================================================

        table_frame = ctk.CTkFrame(
            report_card,
            fg_color="#ffffff",
            corner_radius=8
        )

        table_frame.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=(0, 20)
        )

        # ============================================================
        # TREEVIEW STYLE
        # ============================================================

        style = ttk.Style()

        try:
            style.theme_use(
                "clam"
            )
        except:
            pass

        style.configure(
            "Treeview",
            background="#ffffff",
            foreground="#111827",
            rowheight=38,
            fieldbackground="#ffffff",
            font=("Arial", 10)
        )

        style.configure(
            "Treeview.Heading",
            background="#f1f5f9",
            foreground="#111827",
            font=(
                "Arial",
                10,
                "bold"
            ),
            padding=8
        )

        style.map(
            "Treeview",
            background=[
                ("selected", "#dbeafe")
            ],
            foreground=[
                ("selected", "#111827")
            ]
        )

        # ============================================================
        # TABLE
        # ============================================================

        columns = (
            "No",
            "Date",
            "Bill Number",
            "Seller",
            "Quantity",
            "Item Name",
            "Sales"
        )

        self.sales_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )

        self.sales_table.heading(
            "No",
            text="No"
        )

        self.sales_table.heading(
            "Date",
            text="Date"
        )

        self.sales_table.heading(
            "Bill Number",
            text="Bill Number"
        )

        self.sales_table.heading(
            "Seller",
            text="Seller"
        )

        self.sales_table.heading(
            "Quantity",
            text="Quantity"
        )

        self.sales_table.heading(
            "Item Name",
            text="Item Name"
        )

        self.sales_table.heading(
            "Sales",
            text="Sales"
        )

        self.sales_table.column(
            "No",
            width=55,
            anchor="center"
        )

        self.sales_table.column(
            "Date",
            width=110,
            anchor="center"
        )

        self.sales_table.column(
            "Bill Number",
            width=120,
            anchor="center"
        )

        self.sales_table.column(
            "Seller",
            width=130,
            anchor="center"
        )

        self.sales_table.column(
            "Quantity",
            width=90,
            anchor="center"
        )

        self.sales_table.column(
            "Item Name",
            width=130,
            anchor="center"
        )

        self.sales_table.column(
            "Sales",
            width=120,
            anchor="e"
        )

        # ============================================================
        # SCROLLBAR
        # ============================================================

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.sales_table.yview
        )

        self.sales_table.configure(
            yscrollcommand=scrollbar.set
        )

        self.sales_table.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        # ============================================================
        # LOAD DATA
        # ============================================================

        self.refresh_report()

    # ================================================================
    # SUMMARY CARD
    # ================================================================

    def create_summary_card(
        self,
        parent,
        column,
        title,
        value,
        icon,
        icon_bg
    ):

        card = ctk.CTkFrame(
            parent,
            fg_color="#ffffff",
            corner_radius=14,
            height=130
        )

        card.grid(
            row=0,
            column=column,
            sticky="ew",
            padx=7
        )

        card.grid_propagate(False)

        icon_frame = ctk.CTkFrame(
            card,
            width=55,
            height=55,
            corner_radius=12,
            fg_color=icon_bg
        )

        icon_frame.place(
            x=18,
            y=25
        )

        icon_frame.pack_propagate(False)

        ctk.CTkLabel(
            icon_frame,
            text=icon,
            font=ctk.CTkFont(
                size=22
            )
        ).pack(
            expand=True
        )

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#6b7280"
        ).place(
            x=90,
            y=25
        )

        value_label = ctk.CTkLabel(
            card,
            text=value,
            font=ctk.CTkFont(
                family="Arial",
                size=24,
                weight="bold"
            ),
            text_color="#111827"
        )

        value_label.place(
            x=90,
            y=52
        )

        return value_label

    # ================================================================
    # REFRESH REPORT
    # ================================================================

    def refresh_report(self):

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        # Today's Sales
        try:

            result = self.dbSales.Total(
                today
            )

            value = 0

            if result:
                value = result[0] or 0

            self.sales_card.configure(
                text=f"{value:,.0f}"
            )

        except Exception as error:

            print(
                "Today's Sales Error:",
                error
            )

            self.sales_card.configure(
                text="0"
            )

        # Today's Orders
        try:

            result = self.dbSales.CountQuan(
                today
            )

            value = 0

            if result:
                value = result[0] or 0

            self.orders_card.configure(
                text=str(value)
            )

        except Exception as error:

            print(
                "Today's Orders Error:",
                error
            )

            self.orders_card.configure(
                text="0"
            )

        # Total Items
        try:

            result = self.dbWarehouse.TotalQuantity()

            value = 0

            if result:
                value = result[0] or 0

            self.items_card.configure(
                text=str(value)
            )

        except Exception as error:

            print(
                "Total Items Error:",
                error
            )

            self.items_card.configure(
                text="0"
            )

        self.show_all()

    # ================================================================
    # CLEAR TABLE
    # ================================================================

    def clear_table(self):

        for item in self.sales_table.get_children():

            self.sales_table.delete(
                item
            )

    # ================================================================
    # INSERT TABLE ROW
    # ================================================================

    def insert_rows(
        self,
        rows
    ):

        self.clear_table()

        total = 0

        for index, row in enumerate(
            rows,
            start=1
        ):

            # Database structure:
            #
            # 0 = id
            # 1 = date
            # 2 = sales
            # 3 = billNumber
            # 4 = seller
            # 5 = quantity
            # 6 = itemcode

            db_id = row[0]
            date = row[1]
            sales = row[2] or 0
            bill_number = row[3]
            seller = row[4]
            quantity = row[5] or 0
            item_code = row[7]

            total += sales

            self.sales_table.insert(
                "",
                "end",
                values=(
                    index,
                    date,
                    bill_number,
                    seller,
                    quantity,
                    item_code,
                    f"{sales:,.0f}"
                )
            )

        self.total_label.configure(
            text=f"Total: {total:,.0f}"
        )

    # ================================================================
    # GET ALL SALES
    # ================================================================

    def show_all(self):

        try:

            self.dbSales.cur.execute(
                """
                SELECT *
                FROM SalesReport
                ORDER BY date DESC, id DESC
                """
            )

            rows = self.dbSales.cur.fetchall()

            self.report_title.configure(
                text="All Sales"
            )

            self.insert_rows(
                rows
            )

        except Exception as error:

            print(
                "All Sales Error:",
                error
            )

    def export_excel(self):

        # Get currently displayed table rows
        table_rows = self.sales_table.get_children()

        if not table_rows:

            messagebox.showwarning(
                "Export Excel",
                "There is no sales data to export."
            )

            return

        # Ask user where to save
        file_path = filedialog.asksaveasfilename(
            title="Export Sales Report",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ],
            initialfile="Sales_Report.xlsx"
        )

        if not file_path:
            return

        try:

            # ============================================================
            # CREATE EXCEL WORKBOOK
            # ============================================================

            workbook = Workbook()

            worksheet = workbook.active

            worksheet.title = "Sales Report"

            # ============================================================
            # TITLE
            # ============================================================

            worksheet["A1"] = "ABK MiniMart - Sales Report"

            worksheet["A1"].font = Font(
                size=16,
                bold=True
            )

            worksheet["A1"].alignment = Alignment(
                horizontal="center"
            )

            worksheet.merge_cells(
                "A1:G1"
            )

            # ============================================================
            # REPORT TYPE
            # ============================================================

            worksheet["A2"] = self.report_title.cget(
                "text"
            )

            worksheet["A2"].font = Font(
                size=12,
                bold=True
            )

            worksheet.merge_cells(
                "A2:G2"
            )

            # ============================================================
            # TABLE HEADERS
            # ============================================================

            headers = [
                "No",
                "Date",
                "Bill Number",
                "Seller",
                "Quantity",
                "Item Name",
                "Sales"
            ]

            for column, header in enumerate(
                headers,
                start=1
            ):

                cell = worksheet.cell(
                    row=4,
                    column=column,
                    value=header
                )

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

            # ============================================================
            # TABLE DATA
            # ============================================================

            excel_row = 5

            for item_id in table_rows:

                values = self.sales_table.item(
                    item_id,
                    "values"
                )

                for column, value in enumerate(
                    values,
                    start=1
                ):

                    cell = worksheet.cell(
                        row=excel_row,
                        column=column,
                        value=value
                    )

                    cell.alignment = Alignment(
                        horizontal="center"
                    )

                excel_row += 1

            # ============================================================
            # TOTAL
            # ============================================================

            total_text = self.total_label.cget(
                "text"
            )

            worksheet.cell(
                row=excel_row + 1,
                column=6,
                value="TOTAL"
            ).font = Font(
                bold=True
            )

            worksheet.cell(
                row=excel_row + 1,
                column=7,
                value=total_text.replace(
                    "Total:",
                    ""
                ).strip()
            ).font = Font(
                bold=True
            )

            # ============================================================
            # COLUMN WIDTH
            # ============================================================

            worksheet.column_dimensions["A"].width = 8
            worksheet.column_dimensions["B"].width = 15
            worksheet.column_dimensions["C"].width = 18
            worksheet.column_dimensions["D"].width = 18
            worksheet.column_dimensions["E"].width = 12
            worksheet.column_dimensions["F"].width = 18
            worksheet.column_dimensions["G"].width = 18

            # ============================================================
            # SAVE
            # ============================================================

            workbook.save(
                file_path
            )

            messagebox.showinfo(
                "Export Excel",
                "Sales report exported successfully."
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Could not export Excel file.\n\n{error}"
            )
    # ================================================================
    # SEARCH BY DATE
    # ================================================================

    def search_by_date(self):

        date = self.date_entry.get().strip()

        if date == "":

            messagebox.showwarning(
                "Date",
                "Please enter date.\nExample: 2026-08-13"
            )

            return

        try:

            self.dbSales.cur.execute(
                """
                SELECT *
                FROM SalesReport
                WHERE date=?
                ORDER BY id DESC
                """,
                (date,)
            )

            rows = self.dbSales.cur.fetchall()

            self.report_title.configure(
                text=f"Sales for {date}"
            )

            self.insert_rows(
                rows
            )

        except Exception as error:

            messagebox.showerror(
                "Error",
                str(error)
            )

    # ================================================================
    # SEARCH TOTAL
    # ================================================================

    def search_total(self):

        date = self.date_entry.get().strip()

        if date == "":

            messagebox.showwarning(
                "Date",
                "Please enter date first."
            )

            return

        try:

            result = self.dbSales.Total(
                date
            )

            total = 0

            if result:
                total = result[0] or 0

            self.total_label.configure(
                text=f"Total: {total:,.0f}"
            )

            self.report_title.configure(
                text=f"Total Sales - {date}"
            )

        except Exception as error:

            messagebox.showerror(
                "Error",
                str(error)
            )

    # ================================================================
    # DAILY
    # ================================================================

    def show_daily(self):

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        self.date_entry.delete(
            0,
            "end"
        )

        self.date_entry.insert(
            0,
            today
        )

        self.search_by_date()

    # ================================================================
    # MONTHLY
    # ================================================================

    def show_monthly(self):

        month = datetime.now().strftime(
            "%m"
        )

        try:

            self.dbSales.cur.execute(
                """
                SELECT *
                FROM SalesReport
                WHERE strftime('%m', date)=?
                ORDER BY date DESC, id DESC
                """,
                (month,)
            )

            rows = self.dbSales.cur.fetchall()

            self.report_title.configure(
                text="Current Month Sales"
            )

            self.insert_rows(
                rows
            )

        except Exception as error:

            messagebox.showerror(
                "Error",
                str(error)
            )

    # ================================================================
    # YEARLY
    # ================================================================

    def show_yearly(self):

        year = datetime.now().strftime(
            "%Y"
        )

        try:

            self.dbSales.cur.execute(
                """
                SELECT *
                FROM SalesReport
                WHERE strftime('%Y', date)=?
                ORDER BY date DESC, id DESC
                """,
                (year,)
            )

            rows = self.dbSales.cur.fetchall()

            self.report_title.configure(
                text=f"{year} Sales"
            )

            self.insert_rows(
                rows
            )

        except Exception as error:

            messagebox.showerror(
                "Error",
                str(error)
            )

    # ================================================================
    # REFRESH WHEN PAGE OPENS
    # ================================================================

    def tkraise(
        self,
        aboveThis=None
    ):

        result = super().tkraise(
            aboveThis
        )

        self.refresh_report()

        return result
    
##### Setting Page #############




class SettingPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#1e293b"
        )

        self.controller = controller

        # ============================================================
        # DATABASE
        # ============================================================

        self.db = Registerdata(
            "real/Registerdatabase.db"
        )

        # ============================================================
        # VARIABLES
        # ============================================================

        self.selected_id = None

        self.account_type = ctk.StringVar(
            value="Admin"
        )

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.confirm_var = ctk.StringVar()
        self.search_var = ctk.StringVar()

        # ============================================================
        # MAIN TITLE
        # ============================================================

        title_frame = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )
        title_frame.pack(
            fill="x",
            padx=25,
            pady=(20, 10)
        )

        ctk.CTkLabel(
            title_frame,
            text="Account Settings",
            font=ctk.CTkFont(
                family="Arial",
                size=28,
                weight="bold"
            ),
            text_color="gold"
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            title_frame,
            text="Create, update and delete accounts",
            font=ctk.CTkFont(
                family="Arial",
                size=13
            ),
            text_color="#dddfe2"
        ).pack(
            side="left",
            padx=15,
            pady=(8, 0)
        )

        # ============================================================
        # MAIN CONTENT
        # ============================================================

        content = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )
        content.pack(
            fill="both",
            expand=True,
            padx=25,
            pady=10
        )

        content.grid_columnconfigure(
            0,
            weight=0,
            minsize=350
        )

        content.grid_columnconfigure(
            1,
            weight=1
        )

        content.grid_rowconfigure(
            0,
            weight=1
        )

        # ============================================================
        # LEFT - ACCOUNT CREATION CARD
        # ============================================================

        form_card = ctk.CTkFrame(
            content,
            fg_color="#ffffff",
            corner_radius=14
        )

        form_card.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 10)
        )

        ctk.CTkLabel(
            form_card,
            text="Account Creation",
            font=ctk.CTkFont(
                family="Arial",
                size=21,
                weight="bold"
            ),
            text_color="#111827"
        ).pack(
            anchor="w",
            padx=20,
            pady=(20, 5)
        )

        ctk.CTkLabel(
            form_card,
            text="Manage system login accounts",
            font=ctk.CTkFont(
                family="Arial",
                size=12
            ),
            text_color="#6b7280"
        ).pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

        # ============================================================
        # ACCOUNT TYPE
        # ============================================================

        ctk.CTkLabel(
            form_card,
            text="Account Type",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#374151"
        ).pack(
            anchor="w",
            padx=20,
            pady=(5, 5)
        )

        self.type_menu = ctk.CTkOptionMenu(
            form_card,
            values=[
                "Admin",
                "User"
            ],
            variable=self.account_type,
            height=38,
            corner_radius=8,
            fg_color="#2563eb",
            button_color="#1d4ed8",
            button_hover_color="#1e40af",
            command=self.account_type_changed
        )

        self.type_menu.pack(
            fill="x",
            padx=20,
            pady=(0, 12)
        )

        # ============================================================
        # USERNAME
        # ============================================================

        ctk.CTkLabel(
            form_card,
            text="Username",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#374151"
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 5)
        )

        self.username_entry = ctk.CTkEntry(
            form_card,
            height=38,
            placeholder_text="Enter username",
            corner_radius=8,
            border_width=1,
            border_color="#d1d5db",
            fg_color="#f9fafb",
            textvariable=self.username_var
        )

        self.username_entry.pack(
            fill="x",
            padx=20,
            pady=(0, 12)
        )

        # ============================================================
        # PASSWORD
        # ============================================================

        ctk.CTkLabel(
            form_card,
            text="Password",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#374151"
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 5)
        )

        self.password_entry = ctk.CTkEntry(
            form_card,
            height=38,
            placeholder_text="Enter password",
            show="●",
            corner_radius=8,
            border_width=1,
            border_color="#d1d5db",
            fg_color="#f9fafb",
            textvariable=self.password_var
        )

        self.password_entry.pack(
            fill="x",
            padx=20,
            pady=(0, 12)
        )

        # ============================================================
        # CONFIRM PASSWORD
        # ============================================================

        ctk.CTkLabel(
            form_card,
            text="Confirm Password",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            text_color="#374151"
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 5)
        )

        self.confirm_entry = ctk.CTkEntry(
            form_card,
            height=38,
            placeholder_text="Confirm password",
            show="●",
            corner_radius=8,
            border_width=1,
            border_color="#d1d5db",
            fg_color="#f9fafb",
            textvariable=self.confirm_var
        )

        self.confirm_entry.pack(
            fill="x",
            padx=20,
            pady=(0, 18)
        )

        # ============================================================
        # BUTTONS
        # ============================================================

        button_frame = ctk.CTkFrame(
            form_card,
            fg_color="transparent"
        )

        button_frame.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        self.create_button = ctk.CTkButton(
            button_frame,
            text="Create Account",
            height=38,
            corner_radius=8,
            fg_color="#16a34a",
            hover_color="#15803d",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=self.create_account
        )

        self.create_button.pack(
            fill="x",
            pady=4
        )

        # ============================================================
        # BOTTOM ACTION BUTTONS
        # ============================================================
        
        button_frame = ctk.CTkFrame(
            form_card,
            fg_color="transparent"
        )
        
        button_frame.pack(
            fill="x",
            padx=20,
            pady=(5, 15)
        )
        
        button_frame.grid_columnconfigure(
            0,
            weight=1
        )
        
        button_frame.grid_columnconfigure(
            1,
            weight=1
        )
        
        button_frame.grid_columnconfigure(
            2,
            weight=1
        )
        
        button_frame.grid_columnconfigure(
            3,
            weight=1
        )
        
        # Update
        self.update_button = ctk.CTkButton(
            button_frame,
            text="Update",
            height=38,
            corner_radius=8,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=self.update_account
        )
        
        self.update_button.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 4)
        )
        
        # Delete
        self.delete_button = ctk.CTkButton(
            button_frame,
            text="Delete",
            height=38,
            corner_radius=8,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=self.delete_account
        )
        
        self.delete_button.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=4
        )
        
        # Clear
        self.clear_button = ctk.CTkButton(
            button_frame,
            text="Clear",
            height=38,
            corner_radius=8,
            fg_color="#6b7280",
            hover_color="#4b5563",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=self.clear_form
        )
        
        self.clear_button.grid(
            row=0,
            column=2,
            sticky="ew",
            padx=4
        )
        
        # Dashboard
        self.dashboard_button = ctk.CTkButton(
            button_frame,
            text="Dashboard",
            height=38,
            corner_radius=8,
            fg_color="#B7C23B",
            hover_color="#88930E",
            font=ctk.CTkFont(
                family="Arial",
                size=13,
                weight="bold"
            ),
            command=lambda: self.controller.show_frame(SecondPage)
        )
        
        self.dashboard_button.grid(
            row=0,
            column=3,
            sticky="ew",
            padx=(4, 0)
        )

        # ============================================================
        # RIGHT - ACCOUNT TABLE CARD
        # ============================================================

        table_card = ctk.CTkFrame(
            content,
            fg_color="#ffffff",
            corner_radius=14
        )

        table_card.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(10, 0)
        )

        table_card.grid_rowconfigure(
            2,
            weight=1
        )

        table_card.grid_columnconfigure(
            0,
            weight=1
        )

        # ============================================================
        # TABLE TITLE
        # ============================================================

        ctk.CTkLabel(
            table_card,
            text="Account List",
            font=ctk.CTkFont(
                family="Arial",
                size=21,
                weight="bold"
            ),
            text_color="#111827"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=20,
            pady=(20, 5)
        )

        # ============================================================
        # SEARCH BAR
        # ============================================================

        search_frame = ctk.CTkFrame(
            table_card,
            fg_color="transparent"
        )

        search_frame.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=20,
            pady=(10, 15)
        )

        search_frame.grid_columnconfigure(
            0,
            weight=1
        )

        self.search_entry = ctk.CTkEntry(
            search_frame,
            height=38,
            placeholder_text="Search username...",
            corner_radius=8,
            border_width=1,
            border_color="#d1d5db",
            fg_color="#f9fafb",
            textvariable=self.search_var
        )

        self.search_entry.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8)
        )

        ctk.CTkButton(
            search_frame,
            text="Search",
            width=90,
            height=38,
            corner_radius=8,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self.search_account
        ).grid(
            row=0,
            column=1
        )

        ctk.CTkButton(
            search_frame,
            text="Show All",
            width=90,
            height=38,
            corner_radius=8,
            fg_color="#6b7280",
            hover_color="#4b5563",
            command=self.load_accounts
        ).grid(
            row=0,
            column=2,
            padx=(8, 0)
        )

        # ============================================================
        # TREEVIEW
        # ============================================================

        tree_frame = ctk.CTkFrame(
            table_card,
            fg_color="#ffffff"
        )

        tree_frame.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=20,
            pady=(0, 20)
        )

        tree_frame.grid_rowconfigure(
            0,
            weight=1
        )

        tree_frame.grid_columnconfigure(
            0,
            weight=1
        )

        style = ttk.Style()

        style.configure(
            "Account.Treeview",
            rowheight=38,
            font=("Arial", 11),
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#111827"
        )

        style.configure(
            "Account.Treeview.Heading",
            font=("Arial", 11, "bold"),
            background="#e5e7eb",
            foreground="#111827"
        )

        style.map(
            "Account.Treeview",
            background=[
                ("selected", "#dbeafe")
            ],
            foreground=[
                ("selected", "#111827")
            ]
        )

        self.account_table = ttk.Treeview(
            tree_frame,
            columns=(
                "ID",
                "Username",
                "Password",
                "Confirm"
            ),
            show="headings",
            style="Account.Treeview"
        )

        self.account_table.heading(
            "ID",
            text="ID"
        )

        self.account_table.heading(
            "Username",
            text="Username"
        )

        self.account_table.heading(
            "Password",
            text="Password"
        )

        self.account_table.heading(
            "Confirm",
            text="Confirm"
        )

        self.account_table.column(
            "ID",
            width=60,
            anchor="center"
        )

        self.account_table.column(
            "Username",
            width=180,
            anchor="center"
        )

        self.account_table.column(
            "Password",
            width=180,
            anchor="center"
        )

        self.account_table.column(
            "Confirm",
            width=180,
            anchor="center"
        )

        self.account_table.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.account_table.yview
        )

        scrollbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        self.account_table.configure(
            yscrollcommand=scrollbar.set
        )

        self.account_table.bind(
            "<<TreeviewSelect>>",
            self.select_account
        )

        # Load data
        self.load_accounts()

    # ================================================================
    # ACCOUNT TYPE CHANGED
    # ================================================================

    def account_type_changed(self, value):

        self.clear_form()

        self.load_accounts()

    # ================================================================
    # GET TABLE NAME
    # ================================================================

    def get_table_name(self):

        if self.account_type.get() == "Admin":
            return "Admindata"

        return "Userdata"

    # ================================================================
    # LOAD ACCOUNTS
    # ================================================================

    def load_accounts(self):

        for row in self.account_table.get_children():
            self.account_table.delete(row)

        table_name = self.get_table_name()

        try:

            rows = self.db.cur.execute(
                f"""
                SELECT id, Username, Password, Confirm
                FROM {table_name}
                ORDER BY id DESC
                """
            ).fetchall()

            for row in rows:

                self.account_table.insert(
                    "",
                    "end",
                    values=row
                )

        except sqlite3.Error as error:

            messagebox.showerror(
                "Database Error",
                str(error)
            )

    # ================================================================
    # SEARCH
    # ================================================================

    def search_account(self):

        username = self.search_var.get().strip()

        if username == "":
            self.load_accounts()
            return

        for row in self.account_table.get_children():
            self.account_table.delete(row)

        table_name = self.get_table_name()

        try:

            rows = self.db.cur.execute(
                f"""
                SELECT id, Username, Password, Confirm
                FROM {table_name}
                WHERE Username LIKE ?
                ORDER BY id DESC
                """,
                (f"%{username}%",)
            ).fetchall()

            for row in rows:

                self.account_table.insert(
                    "",
                    "end",
                    values=row
                )

        except sqlite3.Error as error:

            messagebox.showerror(
                "Database Error",
                str(error)
            )

    # ================================================================
    # SELECT ACCOUNT
    # ================================================================

    def select_account(self, event=None):

        selected = self.account_table.selection()

        if not selected:
            return

        values = self.account_table.item(
            selected[0],
            "values"
        )

        if not values:
            return

        self.selected_id = values[0]

        self.username_var.set(values[1])
        self.password_var.set(values[2])
        self.confirm_var.set(values[3])

    # ================================================================
    # CREATE ACCOUNT
    # ================================================================

    def create_account(self):

        username = self.username_var.get().strip()
        password = self.password_var.get()
        confirm = self.confirm_var.get()

        if username == "":
            messagebox.showwarning(
                "Missing Username",
                "Please enter username."
            )
            self.username_entry.focus()
            return

        if password == "":
            messagebox.showwarning(
                "Missing Password",
                "Please enter password."
            )
            self.password_entry.focus()
            return

        if confirm == "":
            messagebox.showwarning(
                "Missing Confirmation",
                "Please confirm password."
            )
            self.confirm_entry.focus()
            return

        if password != confirm:
            messagebox.showwarning(
                "Password Error",
                "Password and Confirm Password do not match."
            )
            return

        table_name = self.get_table_name()

        try:

            exists = self.db.cur.execute(
                f"""
                SELECT id
                FROM {table_name}
                WHERE Username=?
                """,
                (username,)
            ).fetchone()

            if exists:

                messagebox.showwarning(
                    "Account Exists",
                    "This username already exists."
                )

                return

            self.db.cur.execute(
                f"""
                INSERT INTO {table_name}
                (
                    Username,
                    Password,
                    Confirm
                )
                VALUES (?, ?, ?)
                """,
                (
                    username,
                    password,
                    confirm
                )
            )

            self.db.con.commit()

            messagebox.showinfo(
                "Success",
                f"{self.account_type.get()} account created successfully."
            )

            self.clear_form()
            self.load_accounts()

        except sqlite3.Error as error:

            messagebox.showerror(
                "Database Error",
                str(error)
            )

    # ================================================================
    # UPDATE ACCOUNT
    # ================================================================

    def update_account(self):

        if self.selected_id is None:

            messagebox.showwarning(
                "No Account Selected",
                "Please select an account from the table."
            )

            return

        username = self.username_var.get().strip()
        password = self.password_var.get()
        confirm = self.confirm_var.get()

        if username == "":
            messagebox.showwarning(
                "Missing Username",
                "Please enter username."
            )
            return

        if password == "":
            messagebox.showwarning(
                "Missing Password",
                "Please enter password."
            )
            return

        if password != confirm:
            messagebox.showwarning(
                "Password Error",
                "Password and Confirm Password do not match."
            )
            return

        table_name = self.get_table_name()

        try:

            exists = self.db.cur.execute(
                f"""
                SELECT id
                FROM {table_name}
                WHERE Username=?
                AND id != ?
                """,
                (
                    username,
                    self.selected_id
                )
            ).fetchone()

            if exists:

                messagebox.showwarning(
                    "Username Exists",
                    "Another account already uses this username."
                )

                return

            self.db.cur.execute(
                f"""
                UPDATE {table_name}
                SET Username=?,
                    Password=?,
                    Confirm=?
                WHERE id=?
                """,
                (
                    username,
                    password,
                    confirm,
                    self.selected_id
                )
            )

            self.db.con.commit()

            messagebox.showinfo(
                "Success",
                "Account updated successfully."
            )

            self.clear_form()
            self.load_accounts()

        except sqlite3.Error as error:

            messagebox.showerror(
                "Database Error",
                str(error)
            )

    # ================================================================
    # DELETE ACCOUNT
    # ================================================================

    def delete_account(self):

        if self.selected_id is None:

            messagebox.showwarning(
                "No Account Selected",
                "Please select an account from the table."
            )

            return

        username = self.username_var.get()

        answer = messagebox.askyesno(
            "Delete Account",
            f"Are you sure you want to delete '{username}'?"
        )

        if not answer:
            return

        table_name = self.get_table_name()

        try:

            self.db.cur.execute(
                f"""
                DELETE FROM {table_name}
                WHERE id=?
                """,
                (self.selected_id,)
            )

            self.db.con.commit()

            messagebox.showinfo(
                "Deleted",
                "Account deleted successfully."
            )

            self.clear_form()
            self.load_accounts()

        except sqlite3.Error as error:

            messagebox.showerror(
                "Database Error",
                str(error)
            )

    # ================================================================
    # CLEAR FORM
    # ================================================================

    def clear_form(self):

        self.selected_id = None

        self.username_var.set("")
        self.password_var.set("")
        self.confirm_var.set("")

        self.account_table.selection_remove(
            self.account_table.selection()
        )

        self.username_entry.focus()

    # ================================================================
    # CLOSE DATABASE
    # ================================================================

    def close_database(self):

        try:
            self.db.con.close()
        except:
            pass

# ============================================================
# THIRD PAGE
# ============================================================
class ThirdPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(
            parent,
            fg_color="#f1f5f9"
        )

        self.controller = controller

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        self.header = ctk.CTkFrame(
            self,
            height=80,
            fg_color="#1e293b",
            corner_radius=0
        )

        self.header.pack(
            fill="x"
        )

        self.header_title = ctk.CTkLabel(
            self.header,
            text="ABK POS APPLICATION",
            font=ctk.CTkFont(
                family="Arial",
                size=25,
                weight="bold"
            ),
            text_color="#ffffff"
        )

        self.header_title.pack(
            side="left",
            padx=40,
            pady=20
        )

        # ----------------------------------------------------
        # Main Content
        # ----------------------------------------------------

        self.content = ctk.CTkFrame(
            self,
            fg_color="#ffffff",
            corner_radius=20
        )

        self.content.pack(
            fill="both",
            expand=True,
            padx=50,
            pady=50
        )

        self.content_title = ctk.CTkLabel(
            self.content,
            text="Welcome to ABK POS",
            font=ctk.CTkFont(
                family="Arial",
                size=35,
                weight="bold"
            ),
            text_color="#111827"
        )

        self.content_title.pack(
            pady=(100, 10)
        )

        self.content_text = ctk.CTkLabel(
            self.content,
            text=(
                "Store some content related to your project\n"
                "or what your application is made for.\n\n"
                "All the best!"
            ),
            font=ctk.CTkFont(
                family="Arial",
                size=18
            ),
            text_color="#64748b",
            justify="center"
        )

        self.content_text.pack(
            pady=20
        )

        self.home_button = ctk.CTkButton(
            self.content,
            text="Home",
            width=180,
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(
                family="Arial",
                size=16,
                weight="bold"
            ),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=lambda: controller.show_frame(FirstPage)
        )

        self.home_button.pack(
            pady=40
        )


# ============================================================
# APPLICATION
# ============================================================

class Application(ctk.CTk):

    def __init__(self, *args, **kwargs):

        super().__init__(*args, **kwargs)

        # ----------------------------------------------------
        # Window
        # ----------------------------------------------------

        self.title("ABK POS Application")

        # Get desktop resolution

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # ----------------------------------------------------
        # Desktop size
        #
        # Keep some space for Windows taskbar
        # ----------------------------------------------------

        window_width = int(screen_width * 0.90)
        window_height = int(screen_height * 0.88)

        # Minimum size

        window_width = max(window_width, 1100)
        window_height = max(window_height, 650)

        # Prevent window from becoming larger
        # than desktop

        window_width = min(
            window_width,
            screen_width
        )

        window_height = min(
            window_height,
            screen_height - 50
        )

        # Center window

        x = int(
            (screen_width - window_width) / 2
        )

        y = int(
            (screen_height - window_height) / 2
        )

        self.geometry(
            f"{window_width}x{window_height}+{x}+{y}"
        )

        self.minsize(
            1000,
            600
        )

        # ----------------------------------------------------
        # Icon
        # ----------------------------------------------------

        try:

            self.iconbitmap(
                "logoicon.ico"
            )

        except:

            pass

        # ----------------------------------------------------
        # Main Container
        # ----------------------------------------------------

        window = ctk.CTkFrame(
            self,
            fg_color="#f3f5f9",
            corner_radius=0
        )

        window.pack(
            fill="both",
            expand=True
        )

        window.grid_rowconfigure(
            0,
            weight=1
        )

        window.grid_columnconfigure(
            0,
            weight=1
        )

        # ----------------------------------------------------
        # Frames
        # ----------------------------------------------------

        self.frames = {}

        for F in (
            FirstPage,
            SecondPage,
            ThirdPage,
            BillPage,
            UserBillPage,
            FivePage,
            BarcodePage,
            WarehousePage,
            ReportPage,
            SettingPage
        ):

            frame = F(
                window,
                self
            )

            self.frames[F] = frame

            frame.grid(
                row=0,
                column=0,
                sticky="nsew"
            )

        # ----------------------------------------------------
        # Start Page
        # ----------------------------------------------------

        self.show_frame(
            FirstPage
        )

    # --------------------------------------------------------
    # SHOW FRAME
    # --------------------------------------------------------

    def show_frame(self, page):

        frame = self.frames[page]
        frame.tkraise()

        if page == BillPage:
            frame.refresh_seller()

# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app = Application()

    app.mainloop()