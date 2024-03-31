from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton
from check_dialog import CheckDialog
from material_warehouse_dialog import MaterialWarehouseDialog
from product_warehouse_dialog import ProductWarehouseDialog
from settings_dialog import SettingsDialog
from ProductionPartyDialog import ProductionPartyDialog

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Главное окно")
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Создание виджетов для всех кнопок, включая скрытые
        buttons = [
            ("Склад готовой продукции", self.show_product_warehouse_dialog),
            ("Склад материалов", self.show_material_warehouse_dialog),
            ("Сверка", self.show_check_dialog),
            ("Настройки", self.show_settings_dialog),
            ("Создать партию", self.show_production_party_dialog),
            ("Кнопка 6", None),
            ("Кнопка 7", None),
            ("Кнопка 8", None)
        ]

        # Добавление всех кнопок в главное окно
        for text, func in buttons:
            button = QPushButton(text)
            if func:
                button.clicked.connect(func)
            layout.addWidget(button)

            # Скрытие дополнительных кнопок
            if text.startswith("Кнопка"):
                button.hide()

    def show_settings_dialog(self):
        settings_dialog = SettingsDialog()
        settings_dialog.exec_()

    def show_product_warehouse_dialog(self):
        product_dialog = ProductWarehouseDialog()
        product_dialog.exec_()

    def show_material_warehouse_dialog(self):
        material_dialog = MaterialWarehouseDialog()
        material_dialog.exec_()

    def show_check_dialog(self):
        check_dialog = CheckDialog()
        check_dialog.exec_()

    def show_production_party_dialog(self):
        production_party_dialog = ProductionPartyDialog()
        production_party_dialog.exec_()

if __name__ == "__main__":
    # Пример использования
    import sys
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())