from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStatusBar,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.plan_filler import PLAN_COLUMNS, build_cards_workbook, load_teacher_tables_from_rup


def _resource_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / relative_path
    return Path.cwd() / relative_path


def _to_cell_value(value: Any) -> Any:
    text = str(value).strip()
    if text == "":
        return None

    normalized = text.replace(",", ".")
    try:
        number = float(normalized)
    except ValueError:
        return text

    if number.is_integer():
        return int(number)
    return number


class TeacherSelectionDialog(QDialog):
    def __init__(self, teachers: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Выбор преподавателей")
        self.resize(680, 520)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        label = QLabel("Выберите преподавателей, для которых нужно формировать карточки:")
        root.addWidget(label)

        self.list_widget = QListWidget()
        for teacher in teachers:
            item = QListWidgetItem(teacher)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.list_widget.addItem(item)
        root.addWidget(self.list_widget, 1)

        controls = QHBoxLayout()
        select_all_button = QPushButton("Выбрать всех")
        clear_all_button = QPushButton("Снять выбор")
        select_all_button.clicked.connect(self._select_all)
        clear_all_button.clicked.connect(self._clear_all)
        controls.addWidget(select_all_button)
        controls.addWidget(clear_all_button)
        controls.addStretch()
        root.addLayout(controls)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _select_all(self) -> None:
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Checked)

    def _clear_all(self) -> None:
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)

    def selected_teachers(self) -> list[str]:
        result: list[str] = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                result.append(item.text())
        return result


class EditableDataFrameModel(QAbstractTableModel):
    def __init__(self) -> None:
        super().__init__()
        self._df = pd.DataFrame(columns=[key for key, _ in PLAN_COLUMNS])

    def set_dataframe(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df.index)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df.columns)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        if role in (Qt.DisplayRole, Qt.EditRole):
            value = self._df.iat[index.row(), index.column()]
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return ""
            return str(value)
        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        if not index.isValid() or role != Qt.EditRole:
            return False
        self._df.iat[index.row(), index.column()] = _to_cell_value(value)
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole
    ) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if section < len(PLAN_COLUMNS):
                return PLAN_COLUMNS[section][1]
            return None
        return str(section + 1)


class ExcelViewerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Teacher Card Builder")
        self.resize(1320, 800)
        self.setMinimumSize(1020, 660)

        icon_path = _resource_path("assets/app_icon.png")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self._rup_file: Path | None = None
        self._year_start = ""
        self._year_end = ""
        self._teacher_tables: dict[str, pd.DataFrame] = {}
        self._selected_teachers: list[str] = []

        self.table_model = EditableDataFrameModel()
        self._build_ui()
        self._apply_styles()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(18, 18, 18, 14)
        root_layout.setSpacing(12)

        title = QLabel("Teacher Card Builder")
        title.setObjectName("titleLabel")

        subtitle = QLabel(
            "1) Загрузите РУП  2) Выберите преподавателей  3) Исправьте данные  4) Создайте карточки."
        )
        subtitle.setObjectName("subtitleLabel")

        toolbar = QFrame()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setSpacing(10)

        self.open_rup_button = QPushButton("Загрузить РУП")
        self.open_rup_button.clicked.connect(self._load_rup_file)

        self.teacher_combo = QComboBox()
        self.teacher_combo.setMinimumWidth(320)
        self.teacher_combo.setEnabled(False)
        self.teacher_combo.currentTextChanged.connect(self._switch_teacher)

        self.create_cards_button = QPushButton("Создать карточки")
        self.create_cards_button.setEnabled(False)
        self.create_cards_button.clicked.connect(self._create_cards)

        toolbar_layout.addWidget(self.open_rup_button)
        toolbar_layout.addWidget(self.teacher_combo, 1)
        toolbar_layout.addWidget(self.create_cards_button)

        self.file_label = QLabel("РУП не выбран")
        self.file_label.setObjectName("mutedLabel")

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        left_panel = QFrame()
        left_panel.setObjectName("panel")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(14)

        stats_title = QLabel("Статистика")
        stats_title.setObjectName("panelTitle")

        stats_form = QFormLayout()
        stats_form.setContentsMargins(0, 0, 0, 0)
        stats_form.setSpacing(8)

        self.rows_value = QLabel("0")
        self.cols_value = QLabel("0")
        self.filled_value = QLabel("0")
        self.empty_value = QLabel("0")

        stats_form.addRow("Преподавателей:", self.rows_value)
        stats_form.addRow("Строк у выбранного:", self.cols_value)
        stats_form.addRow("Заполненных ячеек:", self.filled_value)
        stats_form.addRow("Пустых ячеек:", self.empty_value)

        left_layout.addWidget(stats_title)
        left_layout.addLayout(stats_form)
        left_layout.addStretch()

        right_panel = QFrame()
        right_panel.setObjectName("panel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.setSpacing(8)

        table_title = QLabel("Данные преподавателя (редактируемо)")
        table_title.setObjectName("panelTitle")

        self.table = QTableView()
        self.table.setModel(self.table_model)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.horizontalHeader().setStretchLastSection(True)

        right_layout.addWidget(table_title)
        right_layout.addWidget(self.table)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([320, 980])

        root_layout.addWidget(title)
        root_layout.addWidget(subtitle)
        root_layout.addWidget(toolbar)
        root_layout.addWidget(self.file_label)
        root_layout.addWidget(splitter, 1)

        self.setCentralWidget(root)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Готово к загрузке РУП")

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: #f3f6fb;
            }
            QLabel {
                color: #0f172a;
            }
            QLabel#titleLabel {
                font-size: 26px;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#subtitleLabel {
                color: #475569;
                font-size: 14px;
            }
            QLabel#mutedLabel {
                color: #475569;
                font-size: 13px;
            }
            QFrame#panel {
                background: #ffffff;
                border: 1px solid #d7deea;
                border-radius: 8px;
            }
            QLabel#panelTitle {
                color: #0f172a;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton {
                background: #1d4ed8;
                color: #ffffff;
                border: 0;
                border-radius: 8px;
                padding: 8px 14px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1e40af;
            }
            QPushButton:disabled {
                background: #94a3b8;
                color: #e2e8f0;
            }
            QComboBox {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 7px 10px;
                font-size: 13px;
            }
            QComboBox QAbstractItemView {
                color: #0f172a;
                background: #ffffff;
                selection-background-color: #dbeafe;
                selection-color: #0f172a;
            }
            QTableView {
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                background: #ffffff;
                color: #0f172a;
                gridline-color: #e2e8f0;
                selection-background-color: #dbeafe;
                selection-color: #0f172a;
            }
            QHeaderView::section {
                background: #e8eefc;
                color: #0f172a;
                border: 0;
                border-right: 1px solid #d7deea;
                border-bottom: 1px solid #d7deea;
                padding: 7px;
                font-weight: 600;
                font-size: 12px;
            }
            QTableCornerButton::section {
                background: #e8eefc;
                border: 1px solid #d7deea;
            }
            QStatusBar {
                background: #ffffff;
                color: #475569;
                border-top: 1px solid #d7deea;
            }
            """
        )

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def _load_rup_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл РУП",
            str(Path.home()),
            "Excel files (*.xlsx *.xlsm *.xls);;All files (*.*)",
        )
        if not file_path:
            return

        try:
            all_tables, (year_start, year_end) = load_teacher_tables_from_rup(file_path)
        except Exception as exc:
            self._show_error("Ошибка чтения РУП", str(exc))
            return

        teachers = sorted(all_tables.keys())
        dialog = TeacherSelectionDialog(teachers, self)
        if dialog.exec() != QDialog.Accepted:
            return

        selected_teachers = dialog.selected_teachers()
        if not selected_teachers:
            self._show_error("Нет выбора", "Нужно выбрать хотя бы одного преподавателя.")
            return

        self._rup_file = Path(file_path)
        self._year_start = year_start
        self._year_end = year_end
        self._selected_teachers = selected_teachers
        self._teacher_tables = {
            teacher: all_tables[teacher].copy(deep=True) for teacher in selected_teachers
        }

        self.teacher_combo.blockSignals(True)
        self.teacher_combo.clear()
        self.teacher_combo.addItems(self._selected_teachers)
        self.teacher_combo.setCurrentIndex(0)
        self.teacher_combo.blockSignals(False)
        self.teacher_combo.setEnabled(True)
        self.create_cards_button.setEnabled(True)

        self.file_label.setText(f"РУП: {self._rup_file.name} | Год: {year_start}-{year_end}")
        self._switch_teacher(self._selected_teachers[0])
        self.statusBar().showMessage("РУП загружен, преподаватели выбраны")

    def _switch_teacher(self, teacher_name: str) -> None:
        if not teacher_name or teacher_name not in self._teacher_tables:
            self.table_model.set_dataframe(pd.DataFrame(columns=[col for col, _ in PLAN_COLUMNS]))
            self._update_stats(None)
            return

        df = self._teacher_tables[teacher_name]
        self.table_model.set_dataframe(df)
        self._update_stats(df)
        self.statusBar().showMessage(
            f"Редактирование: {teacher_name} | строк: {len(df)}"
        )

    def _update_stats(self, df: pd.DataFrame | None) -> None:
        self.rows_value.setText(str(len(self._selected_teachers)))
        if df is None:
            self.cols_value.setText("0")
            self.filled_value.setText("0")
            self.empty_value.setText("0")
            return

        rows = len(df.index)
        cols = len(df.columns)
        empty_cells = int(df.isna().sum().sum()) if rows and cols else 0
        total_cells = rows * cols
        filled_cells = total_cells - empty_cells

        self.cols_value.setText(str(rows))
        self.filled_value.setText(str(filled_cells))
        self.empty_value.setText(str(empty_cells))

    def _create_cards(self) -> None:
        if not self._rup_file:
            self._show_error("Нет РУП", "Сначала загрузите РУП.")
            return
        if not self._selected_teachers:
            self._show_error("Нет преподавателей", "Сначала выберите преподавателей.")
            return

        template_path = self._find_template_in_project()
        if template_path is None:
            self._show_error(
                "Нет шаблона",
                "Не найден шаблон в проекте. Ожидается файл 'Пример.xlsx' "
                "или 'Шаблон и пример.xlsm' в корне проекта.",
            )
            return

        default_out = f"Карточки_{self._year_start}-{self._year_end}{template_path.suffix or '.xlsx'}"
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить карточки как",
            str(template_path.parent / default_out),
            "Excel files (*.xlsx *.xlsm)",
        )
        if not output_file:
            return

        try:
            filled = build_cards_workbook(
                rup_file=self._rup_file,
                template_file=template_path,
                output_file=output_file,
                teacher_tables=self._teacher_tables,
                selected_teachers=self._selected_teachers,
            )
        except Exception as exc:
            self._show_error("Ошибка создания карточек", str(exc))
            return

        lines = [f"{sheet}: {count} строк" for sheet, count in filled.items()]
        self.statusBar().showMessage("Карточки успешно созданы")
        QMessageBox.information(
            self,
            "Готово",
            "Карточки созданы и сохранены:\n" + "\n".join(lines),
        )

    def _find_template_in_project(self) -> Path | None:
        candidates = [
            Path.cwd() / "Пример.xlsx",
            Path.cwd() / "Шаблон и пример.xlsm",
        ]
        for path in candidates:
            if path.exists() and path.is_file():
                return path
        return None


def run_app() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Teacher Card Builder")
    window = ExcelViewerWindow()
    window.show()
    sys.exit(app.exec())
