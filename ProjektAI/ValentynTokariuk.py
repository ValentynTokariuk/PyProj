import sys
import re
import copy
from PyQt5.QtWidgets import (QApplication, QColorDialog, QMainWindow, QTableWidget, QTableWidgetItem, 
                             QFileDialog, QToolButton, QAction, QFontDialog, QInputDialog, QLineEdit, QVBoxLayout, QWidget)
from PyQt5 import QtCore
from PyQt5.QtGui import QKeySequence
import pandas as pd
import matplotlib.pyplot as plt


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(250, 150, 800, 600)
        self.file_path = None
        self.undo_stack = []  # Stores the previous states for undo functionality
        self.redo_stack = []  # Stores the undone states for redo functionality
        self.formulas = {}  # Store formulas for cells
        self.current_cell = None  # Track the currently selected cell (row, col)

        # Main layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)

        # Formula bar
        self.formula_bar = QLineEdit(self)
        self.formula_bar.returnPressed.connect(self.apply_formula)
        self.layout.addWidget(self.formula_bar)

        # Create a table widget
        self.table = QTableWidget(self)
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.setStyleSheet("QHeaderView::section { background-color: grey; }")
        self.table.cellClicked.connect(self.update_formula_bar)  # Update formula bar on cell click
        self.table.cellClicked.connect(self.track_selected_cell)  # Track the selected cell
        self.layout.addWidget(self.table)

        # Toolbar buttons for file operations
        self.add_toolbar_action("Open", self.open_file, QKeySequence.Open)
        self.add_toolbar_action("Save", self.save_file)
        self.add_toolbar_action("Save As", self.save_as_file, QKeySequence.SaveAs)

        # Toolbar buttons for table modifications
        self.add_toolbar_button("Add Row", self.add_row)
        self.add_toolbar_button("Add Column", self.add_column)

        # Context menu actions for editing
        self.table.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.add_context_action("Copy", self.copy, QKeySequence.Copy)
        self.add_context_action("Cut", self.cut, QKeySequence.Cut)
        self.add_context_action("Paste", self.paste, QKeySequence.Paste)

        # Font and style actions
        self.add_context_action("Bold", self.toggle_bold, QKeySequence.Bold)
        self.add_context_action("Italic", self.toggle_italic, QKeySequence.Italic)
        self.add_context_action("Change Font", self.change_font)
        self.add_context_action("Change Font Color", self.change_font_color)

        # Toolbar button for diagram creation
        self.add_toolbar_button("Create Diagram", self.create_diagram)

        # Undo/Redo actions
        self.add_toolbar_action("Undo", self.undo, QKeySequence.Undo)
        self.add_toolbar_action("Redo", self.redo, QKeySequence.Redo)

        # Sorting button
        self.add_toolbar_button("Sort Column", self.sort_column)

        # Alignment actions
        self.add_context_action("Align Left", self.align_left)
        self.add_context_action("Align Center", self.align_center)
        self.add_context_action("Align Right", self.align_right)

        self.update_column_headers()  # Initialize column headers with letters

    def add_toolbar_action(self, name, handler, shortcut=None):
        """Utility to add actions to the toolbar."""
        action = self.addToolBar(name).addAction(name, handler)
        if shortcut:
            action.setShortcuts(shortcut)
        return action

    def add_toolbar_button(self, name, handler):
        """Utility to add tool buttons to the toolbar."""
        button = QToolButton(self)
        button.setText(name)
        button.clicked.connect(handler)
        self.addToolBar(name).addWidget(button)
        return button

    def add_context_action(self, name, handler, shortcut=None):
        """Utility to add actions to the context menu."""
        action = QAction(name, self)
        action.triggered.connect(handler)
        if shortcut:
            action.setShortcuts(shortcut)
        self.table.addAction(action)

    def open_file(self):
        """Open a CSV file and load it into the table."""
        options = QFileDialog.Options()
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if self.file_path:
            try:
                df = pd.read_csv(self.file_path)
                self.table.setRowCount(df.shape[0])
                self.table.setColumnCount(df.shape[1])
                for i, row in df.iterrows():
                    for j, value in enumerate(row):
                        self.table.setItem(i, j, QTableWidgetItem(str(value)))
                self.update_column_headers()  # Update column headers after loading
                self.auto_adjust_columns()
            except Exception as e:
                print(f"Error loading file: {e}")

    def save_file(self):
        """Save the table content back to the original CSV file."""
        if self.file_path:
            self.save_table_to_file(self.file_path)

    def save_as_file(self):
        """Save the table content to a new CSV file."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Save As", "", "CSV (*.csv)")
        if file_path:
            self.save_table_to_file(file_path)

    def save_table_to_file(self, file_path):
        """Helper function to save table data to a CSV file."""
        try:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            df = pd.DataFrame(data)
            df.to_csv(file_path, index=False, header=False)
        except Exception as e:
            print(f"Error saving file: {e}")

    def add_row(self):
        """Add a new empty row to the table."""
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)

    def add_column(self):
        """Add a new empty column to the table."""
        column_count = self.table.columnCount()
        self.table.insertColumn(column_count)
        self.update_column_headers()  # Update column headers after adding a column

    def update_column_headers(self):
        """Update column headers to use letters like A, B, C..."""
        column_count = self.table.columnCount()
        headers = [self.column_index_to_letter(i) for i in range(column_count)]
        self.table.setHorizontalHeaderLabels(headers)

    def column_index_to_letter(self, index):
        """Convert a column index (0, 1, 2, ...) to letter (A, B, C, ...)."""
        result = ""
        while index >= 0:
            result = chr(index % 26 + ord('A')) + result
            index = index // 26 - 1
        return result

    def copy(self):
        """Copy selected table content to the clipboard."""
        clipboard = QApplication.clipboard()
        selected_ranges = self.table.selectedRanges()
        if selected_ranges:
            data = ""
            for row in range(selected_ranges[0].topRow(), selected_ranges[0].bottomRow() + 1):
                row_data = []
                for col in range(selected_ranges[0].leftColumn(), selected_ranges[0].rightColumn() + 1):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                data += '\t'.join(row_data) + '\n'
            clipboard.setText(data.strip())

    def paste(self):
        """Paste content from the clipboard into the table."""
        clipboard = QApplication.clipboard().text()
        rows = clipboard.split("\n")
        for row_idx, row_data in enumerate(rows):
            columns = row_data.split("\t")
            for col_idx, col_data in enumerate(columns):
                item = self.table.item(row_idx, col_idx)
                if item:
                    item.setText(col_data)

    def cut(self):
        """Cut selected table content."""
        self.copy()
        for item in self.table.selectedItems():
            item.setText("")

    def track_selected_cell(self, row, col):
        """Track the selected cell when a user clicks on it."""
        self.current_cell = (row, col)

    def apply_formula(self):
        """Apply the formula typed in the formula bar to the selected cell."""
        if self.current_cell:
            row, col = self.current_cell
            formula = self.formula_bar.text()

            # Save the formula to the dictionary for later editing
            cell_key = f"{row},{col}"
            self.formulas[cell_key] = formula

            # Apply the formula if it starts with '='
            if formula.startswith("="):
                try:
                    result = self.parse_formula(formula[1:])
                    self.table.setItem(row, col, QTableWidgetItem(str(result)))
                except Exception as e:
                    self.table.setItem(row, col, QTableWidgetItem("ERROR"))
            else:
                # It's plain text, so just set the text in the cell
                self.table.setItem(row, col, QTableWidgetItem(formula))

    def update_formula_bar(self, row, col):
        """Update the formula bar when a cell is clicked."""
        cell_key = f"{row},{col}"
        formula = self.formulas.get(cell_key, "")
        self.formula_bar.setText(formula)

    def parse_formula(self, formula):
        """Parse and evaluate the formula."""
        formula = formula.strip()

        if "SUM" in formula:
            return self.evaluate_sum(formula)
        elif "AVERAGE" in formula:
            return self.evaluate_average(formula)
        else:
            return self.evaluate_arithmetic(formula)

    def evaluate_sum(self, formula):
        """Evaluate a SUM formula like 'SUM(A1:B2)'."""
        cell_ranges = self.extract_cell_ranges(formula)
        total = 0
        for cell_range in cell_ranges:
            total += self.get_range_sum(cell_range)
        return total

    def evaluate_average(self, formula):
        """Evaluate an AVERAGE formula like 'AVERAGE(A1:B2)'."""
        cell_ranges = self.extract_cell_ranges(formula)
        total, count = 0, 0
        for cell_range in cell_ranges:
            range_sum, range_count = self.get_range_sum_and_count(cell_range)
            total += range_sum
            count += range_count
        return total / count if count != 0 else 0

    def extract_cell_ranges(self, formula):
        """Extract cell ranges from a formula like 'A1:B2'."""
        return re.findall(r'([A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+)', formula)

    def get_range_sum(self, cell_range):
        """Sum the values in a cell range like 'A1:B2'."""
        start, end = cell_range.split(":")
        start_row, start_col = self.cell_to_indices(start)
        end_row, end_col = self.cell_to_indices(end)
        total = 0
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                total += int(self.table.item(row, col).text() or 0)
        return total

    def get_range_sum_and_count(self, cell_range):
        """Sum the values and count the cells in a cell range like 'A1:B2'."""
        start, end = cell_range.split(":")
        start_row, start_col = self.cell_to_indices(start)
        end_row, end_col = self.cell_to_indices(end)
        total, count = 0, 0
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                value = self.table.item(row, col).text() or "0"
                total += int(value)
                count += 1
        return total, count

    def evaluate_arithmetic(self, formula):
        """Evaluate simple arithmetic formulas like 'A1+B2'."""
        formula = self.replace_cell_references(formula)
        return eval(formula)

    def replace_cell_references(self, formula):
        """Replace cell references (e.g., 'A1') with their numeric values."""
        cells = re.findall(r'[A-Za-z]+[0-9]+', formula)
        for cell in cells:
            row, col = self.cell_to_indices(cell)
            value = self.table.item(row, col).text() or "0"
            formula = formula.replace(cell, value)
        return formula

    def cell_to_indices(self, cell):
        """Convert a cell reference like 'A1' to (row, col) indices."""
        match = re.match(r'([A-Za-z]+)([0-9]+)', cell)
        col = self.letter_to_index(match.group(1))
        row = int(match.group(2)) - 1
        return row, col

    def letter_to_index(self, letter):
        """Convert a letter like 'A' or 'AA' to a zero-based column index."""
        index = 0
        for char in letter:
            index = index * 26 + (ord(char.upper()) - ord('A')) + 1
        return index - 1

    def auto_adjust_columns(self):
        """Auto-adjust column widths based on content."""
        for col in range(self.table.columnCount()):
            self.table.resizeColumnToContents(col)

    def toggle_bold(self):
        """Toggle bold style for selected table items."""
        self.toggle_font_weight(True)

    def toggle_italic(self):
        """Toggle italic style for selected table items."""
        self.toggle_font_weight(False)

    def toggle_font_weight(self, bold):
        """Helper function to apply bold/italic styles to selected table items."""
        selected_items = self.table.selectedItems()
        if selected_items:
            for item in selected_items:
                font = item.font()
                font.setBold(bold) if bold else font.setItalic(True)
                item.setFont(font)

    def change_font(self):
        """Open font dialog to change the font of selected items."""
        font, ok = QFontDialog.getFont()
        if ok:
            for item in self.table.selectedItems():
                item.setFont(font)

    def change_font_color(self):
        """Open color dialog to change the font color of selected items."""
        color = QColorDialog.getColor()
        if color.isValid():
            for item in self.table.selectedItems():
                item.setForeground(color)

    def align_left(self):
        """Align text to the left."""
        self.align_cells(QtCore.Qt.AlignLeft)

    def align_center(self):
        """Align text to the center."""
        self.align_cells(QtCore.Qt.AlignCenter)

    def align_right(self):
        """Align text to the right."""
        self.align_cells(QtCore.Qt.AlignRight)

    def align_cells(self, alignment):
        """Helper function to align selected cells."""
        selected_items = self.table.selectedItems()
        for item in selected_items:
            item.setTextAlignment(alignment)

    def sort_column(self):
        """Sort the data in a specific column."""
        col, ok = QInputDialog.getInt(self, "Sort Column", "Enter the column number to sort:", min=1, max=self.table.columnCount())
        if ok:
            self.table.sortItems(col - 1)

    def create_diagram(self):
        """Create a bar chart from selected columns."""
        col1, ok1 = QInputDialog.getInt(self, "Choose first column", "Enter the first column number (x-axis):", min=1, max=self.table.columnCount())
        if ok1:
            col2, ok2 = QInputDialog.getInt(self, "Choose second column", "Enter the second column number (y-axis):", min=1, max=self.table.columnCount())
            if ok2:
                try:
                    # Prepare data for the x-axis (col1) and y-axis (col2)
                    x_data = []
                    y_data = []
                    
                    for row in range(self.table.rowCount()):
                        item1 = self.table.item(row, col1 - 1)
                        item2 = self.table.item(row, col2 - 1)
                        
                        # Get x-data from col1 (can be string or int)
                        if item1 is not None:
                            x_data.append(item1.text())
                        else:
                            x_data.append("")  # Default to empty string if cell is empty

                        # Get y-data from col2 (numeric data)
                        if item2 is not None:
                            try:
                                y_data.append(float(item2.text()))
                            except ValueError:
                                y_data.append(0)  # Default to 0 if the value cannot be converted to float

                    if len(x_data) > 0 and len(y_data) > 0:
                        plt.figure(figsize=(8, 6))
                        plt.bar(x_data, y_data, color='skyblue')
                        plt.xlabel(f"Column {col1}")
                        plt.ylabel(f"Column {col2}")
                        plt.title(f"Bar Chart for Column {col1} (x) vs Column {col2} (y)")
                        plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for better readability
                        plt.tight_layout()  # Adjust layout to prevent clipping
                        plt.show()
                    else:
                        print("Error: No data to plot.")

                except Exception as e:
                    print(f"Error creating diagram: {e}")


    def undo(self):
        """Undo the last action by restoring the previous state."""
        if self.undo_stack:
            current_state = self.save_current_state()  # Save current state for redo
            self.redo_stack.append(copy.deepcopy(current_state))  # Store it for redo
            previous_state = self.undo_stack.pop()  # Get the previous state
            self.load_state(previous_state)  # Load the previous state into the table

    def redo(self):
        """Redo the last undone action."""
        if self.redo_stack:
            current_state = self.save_current_state()  # Save current state for undo
            self.undo_stack.append(copy.deepcopy(current_state))  # Store it for undo
            next_state = self.redo_stack.pop()  # Get the next state
            self.load_state(next_state)  # Load the next state into the table

    def save_current_state(self):
        """Save the current state of the table."""
        state = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            state.append(row_data)
        return state

    def load_state(self, state):
        """Load a saved state into the table."""
        self.table.setRowCount(len(state))
        self.table.setColumnCount(len(state[0]) if state else 0)
        for row, row_data in enumerate(state):
            for col, text in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(text))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
