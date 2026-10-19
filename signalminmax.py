import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import math
import re
import ast
import operator
from pathlib import Path


# ============================================================
# CẤU HÌNH
# ============================================================

DEFAULT_INPUT_PATH = (
    r"C:\Users\VOM1YH\Documents\dev\ACTOUTPUT"
    r"\outputchannelmaxmin.xlsx"
)

DEFAULT_OUTPUT_PATH = (
    r"C:\Users\VOM1YH\Documents\dev\ACTOUTPUT"
    r"\channel_calculation_results.xlsx"
)


# ============================================================
# XỬ LÝ CÔNG THỨC
# ============================================================

def normalize_formula(formula):
    """
    Chuyển công thức user nhập thành cú pháp Python.

    Ví dụ:
        √(A^2+C^2)    -> sqrt(A**2+C**2)
        A+B+C          -> A+B+C
        A*2+C          -> A*2+C
    """

    formula = formula.strip()

    # Chuyển ^ thành phép lũy thừa Python
    formula = formula.replace("^", "**")

    # Chuyển ký hiệu căn bậc hai
    # Hỗ trợ trường hợp đơn giản như √(A**2+C**2)
    formula = re.sub(
        r"√\s*\(([^()]*)\)",
        r"sqrt(\1)",
        formula
    )

    # Chuyển math.sqrt thành sqrt để xử lý an toàn
    formula = formula.replace("math.sqrt", "sqrt")
    formula = formula.replace("np.sqrt", "sqrt")

    return formula


class SafeFormulaEvaluator:
    """
    Tính công thức bằng AST thay vì eval trực tiếp.

    Chỉ cho phép:
        +, -, *, /, //, %, **
        sqrt, abs, min, max, round
        sin, cos, tan, log, exp
        A, B, C
    """

    binary_operators = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.FloorDiv: operator.floordiv,
        ast.Mod: operator.mod,
        ast.Pow: operator.pow
    }

    unary_operators = {
        ast.UAdd: operator.pos,
        ast.USub: operator.neg
    }

    allowed_functions = {
        "sqrt": math.sqrt,
        "abs": abs,
        "min": min,
        "max": max,
        "round": round,
        "sin": math.sin,
        "cos": math.cos,
        "tan": math.tan,
        "log": math.log,
        "log10": math.log10,
        "exp": math.exp
    }

    def __init__(self, variables):
        self.variables = variables

    def evaluate(self, formula):
        parsed_formula = ast.parse(formula, mode="eval")
        return self._evaluate_node(parsed_formula.body)

    def _evaluate_node(self, node):

        if isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)):
                return node.value

            raise ValueError("Chỉ được sử dụng giá trị số.")

        if isinstance(node, ast.Name):
            if node.id in self.variables:
                return self.variables[node.id]

            raise ValueError(
                f"Biến '{node.id}' không được hỗ trợ. "
                f"Chỉ sử dụng A, B và C."
            )

        if isinstance(node, ast.BinOp):
            operator_type = type(node.op)

            if operator_type not in self.binary_operators:
                raise ValueError("Phép toán không được hỗ trợ.")

            left_value = self._evaluate_node(node.left)
            right_value = self._evaluate_node(node.right)

            return self.binary_operatorsleft_value,
                right_value
            

        if isinstance(node, ast.UnaryOp):
            operator_type = type(node.op)

            if operator_type not in self.unary_operators:
                raise ValueError("Phép toán một ngôi không được hỗ trợ.")

            value = self._evaluate_node(node.operand)

            return self.unary_operatorsvalue

        if isinstance(node, ast.Call):

            if not isinstance(node.func, ast.Name):
                raise ValueError("Cách gọi hàm không hợp lệ.")

            function_name = node.func.id

            if function_name not in self.allowed_functions:
                raise ValueError(
                    f"Hàm '{function_name}' không được hỗ trợ."
                )

            arguments = [
                self._evaluate_node(argument)
                for argument in node.args
            ]

            return self.allowed_functions*arguments

        raise ValueError(
            f"Thành phần công thức không được hỗ trợ: "
            f"{type(node).__name__}"
        )


def calculate_formula(formula, A, B, C):
    normalized_formula = normalize_formula(formula)

    evaluator = SafeFormulaEvaluator(
        variables={
            "A": A,
            "B": B,
            "C": C
        }
    )

    result = evaluator.evaluate(normalized_formula)

    return result, normalized_formula


# ============================================================
# DÒNG CHANNEL TRONG GUI
# ============================================================

class ChannelRow:
    def __init__(
        self,
        parent,
        row_number,
        channel,
        max_value,
        min_value
    ):
        self.channel = channel
        self.max_value = max_value
        self.min_value = min_value

        self.status_var = tk.StringVar(value="OFF")
        self.c_var = tk.StringVar()
        self.a_var = tk.StringVar()
        self.b_var = tk.StringVar()
        self.formula_var = tk.StringVar(value="sqrt(A^2+C^2)")
        self.result_var = tk.StringVar()

        self.create_widgets(parent, row_number)

        # OFF mặc định nên C = MAX
        self.update_c_value()

    def create_widgets(self, parent, row_number):

        padding = 4

        ttk.Label(
            parent,
            text=str(self.channel),
            width=18
        ).grid(
            row=row_number,
            column=0,
            padx=padding,
            pady=padding,
            sticky="w"
        )

        ttk.Label(
            parent,
            text=self.format_number(self.max_value),
            width=12
        ).grid(
            row=row_number,
            column=1,
            padx=padding,
            pady=padding
        )

        ttk.Label(
            parent,
            text=self.format_number(self.min_value),
            width=12
        ).grid(
            row=row_number,
            column=2,
            padx=padding,
            pady=padding
        )

        status_combo = ttk.Combobox(
            parent,
            textvariable=self.status_var,
            values=["OFF", "ON"],
            state="readonly",
            width=7
        )

        status_combo.grid(
            row=row_number,
            column=3,
            padx=padding,
            pady=padding
        )

        status_combo.bind(
            "<<ComboboxSelected>>",
            self.update_c_value
        )

        self.c_entry = ttk.Entry(
            parent,
            textvariable=self.c_var,
            width=12,
            state="readonly"
        )

        self.c_entry.grid(
            row=row_number,
            column=4,
            padx=padding,
            pady=padding
        )

        ttk.Entry(
            parent,
            textvariable=self.a_var,
            width=10
        ).grid(
            row=row_number,
            column=5,
            padx=padding,
            pady=padding
        )

        ttk.Entry(
            parent,
            textvariable=self.b_var,
            width=10
        ).grid(
            row=row_number,
            column=6,
            padx=padding,
            pady=padding
        )

        ttk.Entry(
            parent,
            textvariable=self.formula_var,
            width=28
        ).grid(
            row=row_number,
            column=7,
            padx=padding,
            pady=padding
        )

        ttk.Button(
            parent,
            text="Calculate",
            command=self.calculate
        ).grid(
            row=row_number,
            column=8,
            padx=padding,
            pady=padding
        )

        self.result_entry = ttk.Entry(
            parent,
            textvariable=self.result_var,
            width=15,
            state="readonly"
        )

        self.result_entry.grid(
            row=row_number,
            column=9,
            padx=padding,
            pady=padding
        )

    @staticmethod
    def format_number(value):

        if pd.isna(value):
            return ""

        try:
            number = float(value)

            if number.is_integer():
                return str(int(number))

            return str(number)

        except (TypeError, ValueError):
            return str(value)

    def update_c_value(self, event=None):

        status = self.status_var.get().strip().upper()

        if status == "OFF":
            selected_value = self.max_value
        else:
            selected_value = self.min_value

        if pd.isna(selected_value):
            self.c_var.set("")
        else:
            self.c_var.set(
                self.format_number(selected_value)
            )

        # Khi thay đổi ON/OFF, kết quả cũ không còn hợp lệ
        self.result_var.set("")

    def calculate(self, show_error=True):

        try:
            if not self.a_var.get().strip():
                raise ValueError("Chưa nhập giá trị A.")

            if not self.b_var.get().strip():
                raise ValueError("Chưa nhập giá trị B.")

            if not self.c_var.get().strip():
                raise ValueError("Không có giá trị C hợp lệ.")

            if not self.formula_var.get().strip():
                raise ValueError("Chưa nhập công thức.")

            A = float(self.a_var.get())
            B = float(self.b_var.get())
            C = float(self.c_var.get())

            formula = self.formula_var.get()

            result, normalized_formula = calculate_formula(
                formula=formula,
                A=A,
                B=B,
                C=C
            )

            if isinstance(result, (int, float, np.number)):
                result = float(result)

                if math.isfinite(result):
                    formatted_result = f"{result:.10g}"
                else:
                    formatted_result = str(result)
            else:
                formatted_result = str(result)

            self.result_var.set(formatted_result)

            return {
                "CHANNEL": self.channel,
                "MAX_OF_MAX": self.max_value,
                "MIN_OF_MIN": self.min_value,
                "ON_OFF_STATUS": self.status_var.get(),
                "A": A,
                "B": B,
                "C": C,
                "INPUT_FORMULA": formula,
                "NORMALIZED_FORMULA": normalized_formula,
                "CALCULATION_RESULT": result,
                "STATUS": "OK",
                "ERROR": ""
            }

        except Exception as error:

            self.result_var.set("ERROR")

            if show_error:
                messagebox.showerror(
                    "Calculation Error",
                    f"Channel: {self.channel}\n\n"
                    f"{error}"
                )

            return {
                "CHANNEL": self.channel,
                "MAX_OF_MAX": self.max_value,
                "MIN_OF_MIN": self.min_value,
                "ON_OFF_STATUS": self.status_var.get(),
                "A": self.a_var.get(),
                "B": self.b_var.get(),
                "C": self.c_var.get(),
                "INPUT_FORMULA": self.formula_var.get(),
                "NORMALIZED_FORMULA": "",
                "CALCULATION_RESULT": "",
                "STATUS": "ERROR",
                "ERROR": str(error)
            }

    def get_export_data(self):
        return self.calculate(show_error=False)


# ============================================================
# GUI CHÍNH
# ============================================================

class ChannelCalculationApp:

    def __init__(self, root):

        self.root = root
        self.root.title("ACT Channel Calculation")
        self.root.geometry("1450x750")
        self.root.minsize(1100, 600)

        self.input_path_var = tk.StringVar(
            value=DEFAULT_INPUT_PATH
        )

        self.status_message_var = tk.StringVar(
            value="Select an Excel file and press Load Data."
        )

        self.result_df = pd.DataFrame()
        self.channel_rows = []

        self.configure_style()
        self.create_gui()

    def configure_style(self):

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 16, "bold")
        )

        style.configure(
            "Header.TLabel",
            font=("Segoe UI", 10, "bold"),
            background="#D9EAF7",
            foreground="#1F1F1F",
            padding=6
        )

        style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(12, 7)
        )

    def create_gui(self):

        # ----------------------------------------------------
        # Title
        # ----------------------------------------------------

        title_frame = ttk.Frame(
            self.root,
            padding=(15, 12)
        )

        title_frame.pack(
            fill="x"
        )

        ttk.Label(
            title_frame,
            text="ACT Channel Calculation",
            style="Title.TLabel"
        ).pack(
            side="left"
        )

        # ----------------------------------------------------
        # Input file
        # ----------------------------------------------------

        file_frame = ttk.LabelFrame(
            self.root,
            text="Input Excel",
            padding=10
        )

        file_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 10)
        )

        file_frame.columnconfigure(
            0,
            weight=1
        )

        ttk.Entry(
            file_frame,
            textvariable=self.input_path_var
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8)
        )

        ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_input_file
        ).grid(
            row=0,
            column=1,
            padx=(0, 8)
        )

        ttk.Button(
            file_frame,
            text="Load Data",
            style="Action.TButton",
            command=self.load_excel_data
        ).grid(
            row=0,
            column=2
        )

        # ----------------------------------------------------
        # Formula information
        # ----------------------------------------------------

        information_frame = ttk.LabelFrame(
            self.root,
            text="Formula examples",
            padding=8
        )

        information_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 10)
        )

        ttk.Label(
            information_frame,
            text=(
                "Supported examples: √(A^2+C^2), "
                "sqrt(A^2+C^2), A+B+C, A*2+C, "
                "abs(A-C)+B, max(A,B,C)"
            )
        ).pack(
            anchor="w"
        )

        ttk.Label(
            information_frame,
            text=(
                "Rule: OFF → C = MAX_OF_MAX   |   "
                "ON → C = MIN_OF_MIN"
            )
        ).pack(
            anchor="w",
            pady=(3, 0)
        )

        # ----------------------------------------------------
        # Table container
        # ----------------------------------------------------

        self.table_container = ttk.Frame(
            self.root
        )

        self.table_container.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 10)
        )

        self.canvas = tk.Canvas(
            self.table_container,
            highlightthickness=0
        )

        vertical_scrollbar = ttk.Scrollbar(
            self.table_container,
            orient="vertical",
            command=self.canvas.yview
        )

        horizontal_scrollbar = ttk.Scrollbar(
            self.table_container,
            orient="horizontal",
            command=self.canvas.xview
        )

        self.canvas.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set
        )

        vertical_scrollbar.pack(
            side="right",
            fill="y"
        )

        horizontal_scrollbar.pack(
            side="bottom",
            fill="x"
        )

        self.canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.rows_frame = ttk.Frame(
            self.canvas
        )

        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.rows_frame,
            anchor="nw"
        )

        self.rows_frame.bind(
            "<Configure>",
            self.update_scroll_region
        )

        self.canvas.bind(
            "<Configure>",
            self.update_canvas_width
        )

        self.canvas.bind_all(
            "<MouseWheel>",
            self.on_mousewheel
        )

        self.create_table_header()

        # ----------------------------------------------------
        # Bottom buttons
        # ----------------------------------------------------

        bottom_frame = ttk.Frame(
            self.root,
            padding=(15, 5, 15, 12)
        )

        bottom_frame.pack(
            fill="x"
        )

        ttk.Label(
            bottom_frame,
            textvariable=self.status_message_var
        ).pack(
            side="left"
        )

        ttk.Button(
            bottom_frame,
            text="Calculate All",
            style="Action.TButton",
            command=self.calculate_all
        ).pack(
            side="right",
            padx=(8, 0)
        )

        ttk.Button(
            bottom_frame,
            text="Save Results to Excel",
            style="Action.TButton",
            command=self.save_results
        ).pack(
            side="right"
        )

    def create_table_header(self):

        headers = [
            ("CHANNEL", 18),
            ("MAX", 12),
            ("MIN", 12),
            ("ON/OFF", 8),
            ("C", 12),
            ("A", 10),
            ("B", 10),
            ("FORMULA", 28),
            ("ACTION", 12),
            ("RESULT", 15)
        ]

        for column_number, (title, width) in enumerate(headers):

            ttk.Label(
                self.rows_frame,
                text=title,
                width=width,
                style="Header.TLabel",
                anchor="center"
            ).grid(
                row=0,
                column=column_number,
                padx=2,
                pady=2,
                sticky="nsew"
            )

    def update_scroll_region(self, event=None):

        self.canvas.configure(
            scrollregion=self.canvas.bbox("all")
        )

    def update_canvas_width(self, event):

        required_width = self.rows_frame.winfo_reqwidth()

        self.canvas.itemconfigure(
            self.canvas_window,
            width=max(event.width, required_width)
        )

    def on_mousewheel(self, event):

        self.canvas.yview_scroll(
            int(-1 * (event.delta / 120)),
            "units"
        )

    def browse_input_file(self):

        current_path = Path(
            self.input_path_var.get()
        )

        if current_path.parent.exists():
            initial_directory = str(current_path.parent)
        else:
            initial_directory = str(Path.home())

        selected_file = filedialog.askopenfilename(
            title="Select input Excel file",
            initialdir=initial_directory,
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*")
            ]
        )

        if selected_file:
            self.input_path_var.set(selected_file)

    def clear_channel_rows(self):

        for widget in self.rows_frame.winfo_children():
            widget.destroy()

        self.channel_rows.clear()
        self.create_table_header()

    def load_excel_data(self):

        input_path = self.input_path_var.get().strip()

        if not input_path:
            messagebox.showwarning(
                "Missing file",
                "Please select an input Excel file."
            )
            return

        if not Path(input_path).exists():
            messagebox.showerror(
                "File not found",
                f"Input file does not exist:\n\n{input_path}"
            )
            return

        try:
            df = pd.read_excel(
                input_path,
                engine="openpyxl"
            )

            required_columns = [
                "CHANNEL",
                "MAX",
                "MIN"
            ]

            missing_columns = [
                column
                for column in required_columns
                if column not in df.columns
            ]

            if missing_columns:
                raise ValueError(
                    "Missing required columns: "
                    + ", ".join(missing_columns)
                )

            df["MAX"] = pd.to_numeric(
                df["MAX"],
                errors="coerce"
            )

            df["MIN"] = pd.to_numeric(
                df["MIN"],
                errors="coerce"
            )

            df = df.dropna(
                subset=["CHANNEL"]
            )

            self.result_df = (
                df.groupby(
                    "CHANNEL",
                    as_index=False
                )
                .agg(
                    MAX_OF_MAX=("MAX", "max"),
                    MIN_OF_MIN=("MIN", "min")
                )
            )

            self.clear_channel_rows()

            for row_number, (_, row) in enumerate(
                self.result_df.iterrows(),
                start=1
            ):
                channel_row = ChannelRow(
                    parent=self.rows_frame,
                    row_number=row_number,
                    channel=row["CHANNEL"],
                    max_value=row["MAX_OF_MAX"],
                    min_value=row["MIN_OF_MIN"]
                )

                self.channel_rows.append(
                    channel_row
                )

            self.status_message_var.set(
                f"Loaded {len(self.channel_rows)} channels."
            )

            self.update_scroll_region()

        except Exception as error:
            messagebox.showerror(
                "Load Error",
                f"Could not read the Excel file.\n\n{error}"
            )

    def calculate_all(self):

        if not self.channel_rows:
            messagebox.showwarning(
                "No data",
                "Please load Excel data first."
            )
            return

        success_count = 0
        error_channels = []

        for channel_row in self.channel_rows:

            calculation_data = channel_row.calculate(
                show_error=False
            )

            if calculation_data["STATUS"] == "OK":
                success_count += 1
            else:
                error_channels.append(
                    str(channel_row.channel)
                )

        if error_channels:
            self.status_message_var.set(
                f"Calculated: {success_count}, "
                f"Errors: {len(error_channels)}"
            )

            messagebox.showwarning(
                "Calculation completed",
                f"Successfully calculated: {success_count}\n"
                f"Error channels: {', '.join(error_channels)}"
            )

        else:
            self.status_message_var.set(
                f"All {success_count} channels calculated successfully."
            )

            messagebox.showinfo(
                "Calculation completed",
                f"All {success_count} channels were calculated successfully."
            )

    def save_results(self):

        if not self.channel_rows:
            messagebox.showwarning(
                "No data",
                "Please load Excel data first."
            )
            return

        export_records = [
            channel_row.get_export_data()
            for channel_row in self.channel_rows
        ]

        calculation_df = pd.DataFrame(
            export_records
        )

        default_output = Path(
            DEFAULT_OUTPUT_PATH
        )

        if default_output.parent.exists():
            initial_directory = str(default_output.parent)
        else:
            initial_directory = str(Path.home())

        output_path = filedialog.asksaveasfilename(
            title="Save calculation results",
            initialdir=initial_directory,
            initialfile=default_output.name,
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Workbook", "*.xlsx")
            ]
        )

        if not output_path:
            return

        try:
            with pd.ExcelWriter(
                output_path,
                engine="openpyxl"
            ) as writer:

                self.result_df.to_excel(
                    writer,
                    sheet_name="Channel Summary",
                    index=False
                )

                calculation_df.to_excel(
                    writer,
                    sheet_name="Calculation Results",
                    index=False
                )

            error_count = (
                calculation_df["STATUS"] == "ERROR"
            ).sum()

            self.status_message_var.set(
                f"Results saved: {output_path}"
            )

            if error_count > 0:
                messagebox.showwarning(
                    "Saved with errors",
                    f"Results were saved successfully.\n\n"
                    f"{error_count} channel(s) contain errors.\n\n"
                    f"File:\n{output_path}"
                )
            else:
                messagebox.showinfo(
                    "Save completed",
                    f"Results were saved successfully.\n\n"
                    f"{output_path}"
                )

        except PermissionError:
            messagebox.showerror(
                "Save Error",
                "Could not save the file.\n\n"
                "The output Excel file may currently be open. "
                "Please close it and try again."
            )

        except Exception as error:
            messagebox.showerror(
                "Save Error",
                f"Could not save the result file.\n\n{error}"
            )


# ============================================================
# CHẠY CHƯƠNG TRÌNH
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = ChannelCalculationApp(root)

    root.mainloop()