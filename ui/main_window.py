# ui/main_window.py
# FULL REPLACE ENTIRE FILE

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QListWidget,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QFileDialog,
    QMessageBox,
    QHeaderView
)

from scraper.racing_australia_scraper import RacingAustraliaScraper
from engine.race_collector import RaceCollector
from engine.excel_exporter import ExcelExporter
from engine.search_engine import SearchEngine


class MainWindow(QWidget):

    def __init__(self):

        super().__init__()

        self.scraper = RacingAustraliaScraper()
        self.collector = RaceCollector()
        self.exporter = ExcelExporter()
        self.search_engine = SearchEngine()

        self.race_data = []
        self.selected_race = None

        self.setWindowTitle(
            "AM PRO v1.0"
        )

        self.resize(
            1800,
            900
        )

        self.build_ui()
        self.connect_events()

    def build_ui(self):

        self.main_layout = QVBoxLayout()

        self.url_label = QLabel(
            "Meeting Link"
        )

        self.url_box = QTextEdit()

        self.url_box.setPlaceholderText(
            "Paste Racing Australia Meeting Link"
        )

        self.load_button = QPushButton(
            "LOAD MEETING"
        )

        self.export_button = QPushButton(
            "EXPORT EXCEL"
        )

        self.search_label = QLabel(
            "Search Horse"
        )

        self.search_box = QLineEdit()

        self.search_box.setPlaceholderText(
            "Enter Horse Name"
        )

        self.search_button = QPushButton(
            "SEARCH"
        )

        self.race_list = QListWidget()

        self.info_label = QLabel()

        self.current_race_table = QTableWidget()

        # -------------------------
        # TABLE SETTINGS
        # -------------------------

        self.current_race_table.setWordWrap(
            False
        )

        self.current_race_table.setAlternatingRowColors(
            True
        )

        self.current_race_table.setHorizontalScrollMode(
            QTableWidget.ScrollPerPixel
        )

        self.current_race_table.setVerticalScrollMode(
            QTableWidget.ScrollPerPixel
        )

        self.current_race_table.horizontalHeader().setStretchLastSection(
            False
        )

        self.current_race_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents
        )

        self.current_race_table.verticalHeader().setVisible(
            False
        )

        self.main_layout.addWidget(
            self.url_label
        )

        self.main_layout.addWidget(
            self.url_box
        )

        self.main_layout.addWidget(
            self.load_button
        )

        self.main_layout.addWidget(
            self.export_button
        )

        search_layout = QHBoxLayout()

        search_layout.addWidget(
            self.search_label
        )

        search_layout.addWidget(
            self.search_box
        )

        search_layout.addWidget(
            self.search_button
        )

        self.main_layout.addLayout(
            search_layout
        )

        self.main_layout.addWidget(
            QLabel(
                "Race List"
            )
        )

        self.main_layout.addWidget(
            self.race_list
        )

        self.main_layout.addWidget(
            self.info_label
        )

        self.main_layout.addWidget(
            self.current_race_table
        )

        self.setLayout(
            self.main_layout
        )

    def connect_events(self):

        self.load_button.clicked.connect(
            self.load_meeting
        )

        self.export_button.clicked.connect(
            self.export_excel
        )

        self.search_button.clicked.connect(
            self.search_horse
        )

        self.race_list.itemClicked.connect(
            self.load_race
        )

    def load_meeting(self):

        try:

            meeting_url = (
                self.url_box
                .toPlainText()
                .strip()
            )

            if not meeting_url:
                return

            self.race_list.clear()

            self.race_data = (
                self.collector.collect_meeting(
                    meeting_url
                )
            )

            for race in self.race_data:

                self.race_list.addItem(
                    race.get(
                        "race_name",
                        ""
                    )
                )

            self.info_label.setText(
                f"Meeting Loaded\n\n"
                f"Races Found : "
                f"{len(self.race_data)}"
            )

        except Exception as e:

            self.info_label.setText(
                f"ERROR : {str(e)}"
            )

    def load_race(self, item):

        try:

            race_name = item.text()

            self.selected_race = None

            for race in self.race_data:

                if (
                    race.get(
                        "race_name",
                        ""
                    ) == race_name
                ):

                    self.selected_race = race
                    break

            if not self.selected_race:
                return

            header = self.selected_race.get(
                "header",
                {}
            )

            horses = self.selected_race.get(
                "horses",
                []
            )

            self.info_label.setText(
                f"""
Race : {race_name}

Place : {header.get('Place', '')}
Country : {header.get('Country', '')}
Distance : {header.get('Distance', '')}
Track : {header.get('Track Condition', '')}

Horse Count : {len(horses)}
"""
            )

            columns = [

                "Horse NO",
                "Horse Name",
                "Jockey Name",
                "Trainer Name",
                "Owner Name",

                "Distance",
                "Track Condition",

                "Barrier",

                "Final Weight",

                "Odds",

                "Hcp Rating"

            ]

            self.current_race_table.clear()

            self.current_race_table.setRowCount(
                len(horses)
            )

            self.current_race_table.setColumnCount(
                len(columns)
            )

            self.current_race_table.setHorizontalHeaderLabels(
                columns
            )

            for row, horse in enumerate(horses):

                for col, key in enumerate(columns):

                    value = str(
                        horse.get(
                            key,
                            ""
                        )
                    )

                    self.current_race_table.setItem(
                        row,
                        col,
                        QTableWidgetItem(
                            value
                        )
                    )

        except Exception as e:

            self.info_label.setText(
                f"ERROR : {str(e)}"
            )

    def search_horse(self):

        try:

            horse_name = (
                self.search_box.text()
                .strip()
            )

            if not horse_name:
                return

            history = (
                self.search_engine.get_history(
                    horse_name
                )
            )

            columns = [

                "Date",
                "Place",

                "Horse Name",

                "Jockey Name",
                "Trainer Name",
                "Owner Name",

                "Finishing Position",

                "Class",

                "Distance",

                "Track Condition",

                "Barrier",

                "Weight",

                "Odds",

                "Position @800",
                "Position @600",
                "Position @400",

                "Finishing Time",

                "Rating",

                "Run Details"

            ]

            self.current_race_table.clear()

            self.current_race_table.setRowCount(
                len(history)
            )

            self.current_race_table.setColumnCount(
                len(columns)
            )

            self.current_race_table.setHorizontalHeaderLabels(
                columns
            )

            for row, run in enumerate(history):

                for col, key in enumerate(columns):

                    value = str(
                        run.get(
                            key,
                            ""
                        )
                    )

                    self.current_race_table.setItem(
                        row,
                        col,
                        QTableWidgetItem(
                            value
                        )
                    )

            self.current_race_table.resizeColumnsToContents()

            self.info_label.setText(
                f"Horse : {horse_name}\n"
                f"Runs : {len(history)}"
            )

        except Exception as e:

            self.info_label.setText(
                f"ERROR : {str(e)}"
            )

    def export_excel(self):

        try:

            if not self.race_data:

                QMessageBox.warning(
                    self,
                    "AM PRO",
                    "Load Meeting First"
                )

                return

            print(self.race_data[0]["header"])
            
            header = self.race_data[0]["header"]

            date_text = (
                header.get(
                    "Date",
                    ""
                )
            )

            place_text = (
                header.get(
                    "Place",
                    ""
                )
            )

            parts = date_text.split()

            if len(parts) == 3:

                day = parts[0]

                month_map = {

                    "January":"Jan",
                    "February":"Feb",
                    "March":"Mar",
                    "April":"Apr",
                    "May":"May",
                    "June":"Jun",
                    "July":"Jul",
                    "August":"Aug",
                    "September":"Sep",
                    "October":"Oct",
                    "November":"Nov",
                    "December":"Dec"

                }

                month = month_map.get(
                    parts[1],
                    parts[1][:3]
                )   

                year = parts[2]

                default_name = (
                    f"{day}{month}{year}_{place_text}.xlsx"
                )

            else:

                default_name = (
                    f"{place_text}.xlsx"
                )

            file_name, _ = QFileDialog.getSaveFileName(

                self,

                "Save Excel",

                default_name,

                "Excel Files (*.xlsx)"

            )

            if not file_name:
                return

            self.exporter.export_meeting(
                template_path="templates/AM_PRO_Template.xlsx",
                output_path=file_name,
                race_collection=self.race_data
            )

            QMessageBox.information(
                self,
                "AM PRO",
                "Excel Export Complete"
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "AM PRO",
                str(e)
            )