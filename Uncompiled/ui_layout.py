from PyQt6.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QComboBox, QMessageBox
from PyQt6.QtGui import QFont, QPixmap
from PyQt6.QtCore import Qt
from excel_export import export_to_excel  # Import the Excel functionality
from camera import CameraHandler
from datetime import datetime

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle("Machine Dashboard")
        self.setGeometry(100, 100, 1600, 900)

        # Apply stylesheets for background and text color
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
            }
            QLabel, QTableWidget, QPushButton, QComboBox {
                color: black;
                font-size: 16px;
                font-family: Arial, sans-serif;
            }
            QTableWidget {
                background-color: white;
                gridline-color: black;
                color: black;
            }
            QHeaderView::section {
                background-color: lightgray;
                color: black;
            }
            QTableWidget QHeaderView::section {
                background-color: white;  /* Ensure white header */
                color: black;
                border: 1px solid black;
            }
            QComboBox {
                background-color: white;
                color: black;
                border: 1px solid black;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                color: black;
                selection-background-color: lightgray;
            }
            QPushButton {
                background-color: #e0e0e0;
                border: 1px solid black;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
        """)

        main_layout = QHBoxLayout()

        # Adjust left side layout (table and buttons)
        left_layout = QVBoxLayout()

        # Create the data table with all columns including Timestamp
        self.data_table = QTableWidget(24, 7, self)
        self.data_table.setHorizontalHeaderLabels(["LOT ID", "CBD", "Maker", "BMS", "Station","Direction", "Timestamp"])
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.data_table.setStyleSheet("color: black;")  # Set the text color to black
        left_layout.addWidget(self.data_table)

        # Add buttons for upload and delete
        button_layout = QHBoxLayout()

        self.upload_btn = QPushButton("Upload", self)
        self.upload_btn.setFixedSize(150, 50)
        self.upload_btn.setFont(QFont('Arial', 14))
        self.upload_btn.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.upload_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.delete_btn = QPushButton("Delete", self)
        self.delete_btn.setFixedSize(150, 50)
        self.delete_btn.setFont(QFont('Arial', 14))
        self.delete_btn.clicked.connect(self.delete_data)
        button_layout.addWidget(self.delete_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        left_layout.addLayout(button_layout)

        # Center layout for camera and dropdowns
        center_layout = QVBoxLayout()

        # Dropdown layout (Station, Direction, Camera)
        dropdown_layout = QHBoxLayout()

        # Station dropdown
        self.station_dropdown = QComboBox(self)
        self.station_dropdown.addItems([
            "After Assy", "BA-RF", "BS", "MOKU-AVI", "BI", 
            "SG Cleaning", "C Test", "CBUN", "JUNB-AVI", "JUNB-BAKE"
        ])
        dropdown_layout.addWidget(self.station_dropdown)

        # Direction dropdown
        self.direction_dropdown = QComboBox(self)
        self.direction_dropdown.addItems(["IN", "OUT"])
        dropdown_layout.addWidget(self.direction_dropdown)

        # Camera dropdown
        self.cam_select = QComboBox(self)
        self.cam_select.addItems(["Camera 1", "Camera 2"])
        self.cam_select.setFixedWidth(150)  # Shorten the camera dropdown
        dropdown_layout.addWidget(self.cam_select)

        center_layout.addLayout(dropdown_layout)  # Add dropdowns above the camera

        # Camera display (black box simulation)
        self.camera_label = QLabel(self)
        self.camera_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.camera_label.setStyleSheet("background-color: black;")  # Black background for the camera feed
        center_layout.addWidget(self.camera_label)

        # Initialize camera handler
        self.cam_handler = CameraHandler(self.camera_label)  

        # Add layouts to main layout
        main_layout.addLayout(left_layout)
        main_layout.addLayout(center_layout)

        main_layout.setStretch(0, 1)  # Left layout (table and buttons)
        main_layout.setStretch(1, 1)  # Center layout (camera)

        # Ensure no extra space at the top and bottom
        main_layout.setContentsMargins(0, 0, 0, 0)  # Set zero margins globally for the main layout
        main_layout.setSpacing(0)  # Set zero spacing for the main layout

        widget = QWidget()
        widget.setLayout(main_layout)
        self.setCentralWidget(widget)

        # Connect dropdown to camera start function
        self.cam_select.currentIndexChanged.connect(self.start_camera)

    def start_camera(self):
        cam_index = self.cam_select.currentIndex()
        self.cam_handler.start_camera(cam_index)

    def export_to_excel(self):
        export_to_excel(self.data_table)

    def delete_data(self):
        self.data_table.clearContents()
        print("Left table data deleted.")
