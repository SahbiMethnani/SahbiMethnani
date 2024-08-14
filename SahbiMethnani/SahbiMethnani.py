import sys
import sqlite3
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import win32com.client

class Employee(QMainWindow):
    def __init__(self):
        super().__init__()
        self.conn = sqlite3.connect('/home/sahbimethnani/Bureau/Python_Projects/gesemployé/employee_management.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                age INTEGER NOT NULL,
                position TEXT NOT NULL
            )
        """)
        self.conn.commit()
        self.setWindowTitle("Gestion des employés")
        self.setGeometry(100, 100, 800, 600)  # Augmenter la taille pour inclure les nouveaux boutons
        
        # Main widget setup
        layout = QVBoxLayout()
        widget = QWidget()
        self.setCentralWidget(widget)
        widget.setLayout(layout)
        widget.setStyleSheet("Background-color: #ACFFFC")
        
        # Image label setup
        photo_label = QLabel()
        photo_pixmap = QPixmap("image.png")  # Assurez-vous que l'image est au bon chemin
        scaled_pixmap = photo_pixmap.scaledToHeight(200)  # Ajuster la largeur de l'image à la largeur de la fenêtre
        photo_label.setPixmap(scaled_pixmap)
        photo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(photo_label)

        # Form layout
        form_layout = QFormLayout()
        self.name_input, self.age_input, self.position_input = QLineEdit(), QLineEdit(), QLineEdit()
        form_layout.addRow("Nom:", self.name_input)
        form_layout.addRow("Âge:", self.age_input)
        form_layout.addRow("Poste:", self.position_input)
        layout.addLayout(form_layout)

        # Buttons layout
        button_layout = QHBoxLayout()
        buttons = [("Ajouter", self.add), ("Mettre à jour", self.update), 
                   ("Supprimer", self.delete), ("Réinitialiser", self.reset_form),
                   ("Importer Excel", self.import_excel), ("Exporter Excel", self.export_excel),
                   ("Quitter", self.close)]
        for text, slot in buttons:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            button_layout.addWidget(btn)
        layout.addLayout(button_layout)

        # Table setup
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "Nom", "Âge", "Poste"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)
        self.table.clicked.connect(self.select_employee)
    
        self.load_employees()
 
    def load_employees(self):
        self.table.setRowCount(0)
        self.cursor.execute("SELECT * FROM employees")
        for row_number, row_data in enumerate(self.cursor.fetchall()):
            self.table.insertRow(row_number)
            for col_number, data in enumerate(row_data):
                self.table.setItem(row_number, col_number, QTableWidgetItem(str(data)))

    def add(self):
        if all((self.name_input.text(), self.age_input.text(), self.position_input.text())):
            self.cursor.execute("INSERT INTO employees (name, age, position) VALUES (?, ?, ?)",
                                (self.name_input.text(), self.age_input.text(), self.position_input.text()))
            self.conn.commit()
            self.load_employees()
            self.reset_form()
            QMessageBox.information(self, "Succès", "Employé ajouté")
        else:
            QMessageBox.warning(self, "Erreur", "Tous les champs sont obligatoires")

    def update(self):
        if self.table.currentRow() != -1:
            row_id = int(self.table.item(self.table.currentRow(), 0).text())
            if all((self.name_input.text(), self.age_input.text(), self.position_input.text())):
                self.cursor.execute("UPDATE employees SET name = ?, age = ?, position = ? WHERE id = ?",
                                    (self.name_input.text(), self.age_input.text(), self.position_input.text(), row_id))
                self.conn.commit()
                self.load_employees()
                self.reset_form()
                QMessageBox.information(self, "Succès", "Employé mis à jour")
            else:
                QMessageBox.warning(self, "Erreur", "Tous les champs sont obligatoires")
        else:
            QMessageBox.warning(self, "Erreur", "Sélectionnez un employé")

    def delete(self):
        if self.table.currentRow() != -1:
            row_id = int(self.table.item(self.table.currentRow(), 0).text())
            self.cursor.execute("DELETE FROM employees WHERE id = ?", (row_id,))
            self.conn.commit()
            self.load_employees()
            self.reset_form()
            QMessageBox.information(self, "Succès", "Employé supprimé")
        else:
            QMessageBox.warning(self, "Erreur", "Sélectionnez un employé")

    def reset_form(self):
        self.name_input.clear()
        self.age_input.clear()
        self.position_input.clear()
        self.table.clearSelection()

    def select_employee(self):
        if self.table.currentRow() != -1:
            self.name_input.setText(self.table.item(self.table.currentRow(), 1).text())
            self.age_input.setText(self.table.item(self.table.currentRow(), 2).text())
            self.position_input.setText(self.table.item(self.table.currentRow(), 3).text())

    def import_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Ouvrir un fichier Excel", "", "Fichiers Excel (*.xlsx)")
        if file_name:
            df = pd.read_excel(file_name)
            self.table.setRowCount(0)
            for row_number, row_data in df.iterrows():
                self.table.insertRow(row_number)
                for col_number, data in enumerate(row_data):
                    self.table.setItem(row_number, col_number, QTableWidgetItem(str(data)))
            self.reset_form()

    def export_excel(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Enregistrer un fichier Excel", "", "Fichiers Excel (*.xlsx)")
        if file_name:
            df = pd.DataFrame(columns=["ID", "Nom", "Âge", "Poste"])
            for row_number in range(self.table.rowCount()):
                row_data = [self.table.item(row_number, col_number).text() for col_number in range(self.table.columnCount())]
                df.loc[row_number] = row_data
            df.to_excel(file_name, index=False)
            QMessageBox.information(self, "Succès", "Données exportées avec succès")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Employee()
    window.show()
    sys.exit(app.exec())
