import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QFileDialog, QToolButton, QColorDialog, QAction, QFontDialog, QInputDialog
from PyQt5 import Qt, QtCore
from PyQt5.QtGui import QKeySequence

import numpy as np
import matplotlib.pyplot as plt



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(250, 150, self.width()*2, self.height()*2) 
        self.file_path = None
        
        # Create a table widget
        self.table = QTableWidget(self)
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.setStyleSheet("QHeaderView::section { background-color: grey; }")

        self.setCentralWidget(self.table)
        
        # Create a "Open" button to open a CSV file
        self.open_button = self.addToolBar("Open").addAction("Open", self.open_file)
        self.open_button.setShortcuts(QKeySequence.Open)

        #Save button
        self.save_button = self.addToolBar("Save").addAction("Save", self.save_file)
        
        #Save As Button
        self.save_as_button = self.addToolBar("Save As").addAction("Save As", self.save_as_file)
        self.save_as_button.setShortcuts(QKeySequence.SaveAs)

        #Change color button
        self.change_color_button = self.addToolBar("Change color").addAction("Change color", self.change_color)

        # Create an "Add Row" button
        self.add_row_button = QToolButton(self)
        self.add_row_button.setText("Add Row")
        self.add_row_button.clicked.connect(self.add_row)
        self.addToolBar("Add Row").addWidget(self.add_row_button)

        #"Add Column" button
        self.add_column_button = QToolButton(self)
        self.add_column_button.setText("Add Column")
        self.add_column_button.clicked.connect(self.add_column)
        self.addToolBar("Add Column").addWidget(self.add_column_button)

        #Context menu buttons
        self.table.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.change_color_action = QAction("Change color", self)
        self.change_color_action.triggered.connect(self.change_color)
        self.table.addAction(self.change_color_action)

        self.copy_action = QAction("Copy", self)
        self.copy_action.setShortcuts(QKeySequence.Copy)
        self.copy_action.triggered.connect(self.copy)
        self.table.addAction(self.copy_action)

        self.cut_action = QAction("Cut", self)
        self.cut_action.setShortcuts(QKeySequence.Cut)
        self.cut_action.triggered.connect(self.cut)
        self.table.addAction(self.cut_action)

        self.paste_action = QAction("Paste", self)
        self.paste_action.setShortcuts(QKeySequence.Paste)
        self.paste_action.triggered.connect(self.paste)
        self.table.addAction(self.paste_action)

        #Create diagram button
        self.create_diagram_button = QToolButton(self)
        self.create_diagram_button.setText("Create Diagram")
        self.create_diagram_button.clicked.connect(self.create_diagram)
        self.addToolBar("Create Diagram").addWidget(self.create_diagram_button)

        # Create actions for changing font
        self.bold_action = QAction("Bold", self)
        self.bold_action.triggered.connect(self.make_bold)
        self.cursive_action = QAction("Cursive", self)
        self.cursive_action.triggered.connect(self.make_cursive)
        self.change_font_action = QAction("Change font", self)
        self.change_font_action.triggered.connect(self.change_font)

        # Add actions to context menu policy
        self.bold_action.setShortcuts(QKeySequence.Bold)
        self.table.addAction(self.bold_action)
        self.cursive_action.setShortcuts(QKeySequence.Italic)
        self.table.addAction(self.cursive_action)
        self.table.addAction(self.change_font_action)

        self.change_font_color_action = QAction("Change font color", self)
        self.change_font_color_action.triggered.connect(self.change_font_color)
        self.table.addAction(self.change_font_color_action)

        

        
       

    def open_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if self.file_path:
            # Load data from CSV file
            data = np.genfromtxt(self.file_path, delimiter=',', dtype=str)
            
            # Set the table dimensions
            self.table.setRowCount(data.shape[0])
            self.table.setColumnCount(data.shape[1])
            
            # Populate the table with data
            for row in range(data.shape[0]):
                for col in range(data.shape[1]):
                    item = QTableWidgetItem(data[row, col])
                    self.table.setItem(row, col, item)
                    
            # Set row and column headers
            self.table.setVerticalHeaderLabels(["{}".format(i+1) for i in range(data.shape[0])])
            self.table.setHorizontalHeaderLabels(["{}".format(i+1) for i in range(data.shape[1])])
          
    def save_file(self):
        if self.file_path:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text())
                data.append(row_data)
            np.savetxt(self.file_path, np.array(data), delimiter=",", fmt='%s')        

    def save_as_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save As", "", "CSV (*.csv)")
        if file_path:
            file = open(file_path, 'w')
            for row_index in range(self.table.rowCount()):
                row = [self.table.item(row_index, column_index).text() for column_index in range(self.table.columnCount())]
                file.write(",".join(row) + "\n")
            file.close()

    def add_row(self):
        # Get the current number of rows
        rows = self.table.rowCount()
        
        # Insert a new row at the bottom of the table
        self.table.insertRow(rows)
        
        # Add empty items to the new row
        for col in range(self.table.columnCount()):
            self.table.setItem(rows, col, QTableWidgetItem(""))
    def add_column(self):
        row = self.table.rowCount()
        column = self.table.columnCount()
        self.table.insertColumn(column)
        for i in range(row):
            self.table.setItem(i, column, QTableWidgetItem(""))

    def change_color(self):
            color = QColorDialog.getColor()
            if color.isValid():
                self.setStyleSheet("QMainWindow { background-color: %s; color: %s; }" % (color.name(), color.name()))
    def change_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            for item in self.table.selectedItems():
                item.setBackground(color)
 
    def open_context_menu(self, position):
        self.context_menu.exec_(self.table.viewport().mapToGlobal(position))

    def copy(self):
        selected_items = self.table.selectedItems()
        if selected_items:
            data = ""
            for item in selected_items:
                data += item.text() + '\t'
            clipboard = QApplication.clipboard()
            clipboard.setText(data)

    def paste(self):
        clipboard = QApplication.clipboard()
        data = clipboard.text().split("\t")
        for item in self.table.selectedItems():
            item.setText(data.pop(0))
    def cut(self):
        selected_ranges = self.table.selectedRanges()
        selected_text = ""
        for i in range(len(selected_ranges)):
            for row in range(selected_ranges[i].topRow(), selected_ranges[i].bottomRow()+1):
                for col in range(selected_ranges[i].leftColumn(), selected_ranges[i].rightColumn()+1):
                    selected_text += self.table.item(row, col).text() + "\t"
                selected_text = selected_text[:-1] + "\n"
        selected_text = selected_text[:-1]
        QtWidgets.QApplication.clipboard().setText(selected_text)

        for i in range(len(selected_ranges)):
            for row in range(selected_ranges[i].topRow(), selected_ranges[i].bottomRow()+1):
                for col in range(selected_ranges[i].leftColumn(), selected_ranges[i].rightColumn()+1):
                    self.table.setItem(row, col, QTableWidgetItem(""))

    
    def create_diagram(self):
        col1, ok1 = QInputDialog.getInt(self, "Choose first column", "Enter the first column number:", min=1, max=self.table.columnCount())
        if ok1:
            col2, ok2 = QInputDialog.getInt(self, "Choose second column", "Enter the second column number:", min=1, max=self.table.columnCount())
            if ok2:
                data1 = []
                data2 = []
                for row in range(self.table.rowCount()):
                    item1 = self.table.item(row, col1-1)
                    if item1:
                        data1.append(item1.text())
                    item2 = self.table.item(row, col2-1)
                    if item2:
                        data2.append(item2.text())
                data1 = [int(i) for i in data1]
                data2 = [int(i) for i in data2]

                fig = plt.figure()
                ax = fig.add_axes([0,0,1,1])
                ax.bar(data1, data2)
                plt.show()
    def make_bold(self):
        selected_items = self.table.selectedItems()
        font = selected_items[0].font()
        font.setBold(True)
        for item in selected_items:
            item.setFont(font)
        
    def make_cursive(self):
        selected_items = self.table.selectedItems()
        font = selected_items[0].font()
        font.setItalic(True)
        for item in selected_items:
            item.setFont(font)
        
    def change_font(self):
        selected_items = self.table.selectedItems()
        font, ok = QFontDialog.getFont()
        if ok:
            for item in selected_items:
                item.setFont(font)

    def change_font_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            for item in self.table.selectedItems():
                item.setForeground(color)


    



if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
