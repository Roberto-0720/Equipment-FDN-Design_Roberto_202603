"""
PyQt5 Main Window for Octagonal Footing Design Tool.
Two tabs: Input Data and Reaction Data.
"""

import os
import sys
import csv
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QGroupBox,
    QComboBox, QMessageBox, QFileDialog, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QSplitter, QFrame, QDialog,
    QDialogButtonBox, QStatusBar, QSizePolicy, QScrollArea, QApplication
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon

from utils.analysis import parse_reactions, run_analysis, export_analysis_xlsx


class AnalysisThread(QThread):
    """Background thread for running analysis."""
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, input_params, reactions, ds_mapping, use_end_node):
        super().__init__()
        self.input_params = input_params
        self.reactions = reactions
        self.ds_mapping = ds_mapping
        self.use_end_node = use_end_node
    
    def run(self):
        try:
            result = run_analysis(
                self.input_params,
                self.reactions,
                self.ds_mapping,
                self.use_end_node,
                progress_callback=lambda cur, tot, msg: self.progress.emit(cur, tot, msg)
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class DsEditDialog(QDialog):
    """Dialog for editing DS values for each load combination."""
    
    def __init__(self, lc_list, ds_mapping, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Ds Values for Load Combinations")
        self.setMinimumSize(500, 500)
        self.ds_mapping = ds_mapping.copy()
        
        layout = QVBoxLayout(self)
        
        info_label = QLabel("Set Soil Depth (Ds) for each Load Combination:")
        info_label.setFont(QFont('Segoe UI', 10))
        layout.addWidget(info_label)
        
        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['Load Case', 'Ds (m)'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setRowCount(len(lc_list))
        
        for i, lc in enumerate(lc_list):
            lc_item = QTableWidgetItem(str(lc))
            lc_item.setFlags(lc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, lc_item)
            
            ds_val = ds_mapping.get(lc, 1.0)
            ds_item = QTableWidgetItem(str(ds_val))
            self.table.setItem(i, 1, ds_item)
        
        layout.addWidget(self.table)
        
        # Set all button
        set_all_layout = QHBoxLayout()
        set_all_layout.addWidget(QLabel("Set all Ds to:"))
        self.set_all_input = QLineEdit("1.0")
        self.set_all_input.setMaximumWidth(80)
        set_all_layout.addWidget(self.set_all_input)
        set_all_btn = QPushButton("Apply to All")
        set_all_btn.clicked.connect(self._apply_all)
        set_all_layout.addWidget(set_all_btn)
        set_all_layout.addStretch()
        layout.addLayout(set_all_layout)
        
        # Import/Export CSV
        csv_layout = QHBoxLayout()
        import_btn = QPushButton("Import CSV")
        import_btn.clicked.connect(self._import_csv)
        csv_layout.addWidget(import_btn)
        export_btn = QPushButton("Export CSV")
        export_btn.clicked.connect(self._export_csv)
        csv_layout.addWidget(export_btn)
        csv_layout.addStretch()
        layout.addLayout(csv_layout)
        
        # OK/Cancel
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.lc_list = lc_list
    
    def _apply_all(self):
        try:
            val = float(self.set_all_input.text())
            for i in range(self.table.rowCount()):
                self.table.setItem(i, 1, QTableWidgetItem(str(val)))
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid Ds value.")
    
    def _import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import CSV", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, 'r') as f:
                    reader = csv.reader(f)
                    next(reader, None)  # Skip header
                    imported = {}
                    for row in reader:
                        if len(row) >= 2:
                            try:
                                lc = int(row[0])
                                ds = float(row[1])
                                imported[lc] = ds
                            except ValueError:
                                continue
                    
                    for i in range(self.table.rowCount()):
                        lc = self.lc_list[i]
                        if lc in imported:
                            self.table.setItem(i, 1, QTableWidgetItem(str(imported[lc])))
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to import CSV:\n{e}")
    
    def _export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export CSV", "load_case_ds.csv", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['LC', 'DS'])
                    for i in range(self.table.rowCount()):
                        lc = self.table.item(i, 0).text()
                        ds = self.table.item(i, 1).text()
                        writer.writerow([lc, ds])
                QMessageBox.information(self, "Success", f"Exported to {path}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to export CSV:\n{e}")
    
    def _on_ok(self):
        self.ds_mapping = {}
        for i in range(self.table.rowCount()):
            lc = self.lc_list[i]
            try:
                ds = float(self.table.item(i, 1).text())
            except (ValueError, AttributeError):
                ds = 1.0
            self.ds_mapping[lc] = ds
        self.accept()
    
    def get_ds_mapping(self):
        return self.ds_mapping


class MainWindow(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Octagonal Spread Footing Design Tool v1.0")
        self.setMinimumSize(900, 700)
        
        # Data
        self.reactions = []
        self.ds_mapping = {}
        self.analysis_result = None
        self.parsed_lc_list = []
        
        self._setup_ui()
        self._apply_styles()
    
    def _apply_styles(self):
        """Apply modern dark-ish professional stylesheet."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QTabWidget::pane {
                border: 1px solid #c0c0c0;
                background-color: white;
                border-radius: 4px;
            }
            QTabBar::tab {
                background: #e0e4e8;
                border: 1px solid #c0c0c0;
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 11px;
                font-weight: bold;
                color: #333;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom-color: white;
                color: #003366;
            }
            QTabBar::tab:hover {
                background: #d0d4d8;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 11px;
                color: #003366;
                border: 1px solid #b0b8c0;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QLineEdit {
                padding: 4px 8px;
                border: 1px solid #c0c0c0;
                border-radius: 3px;
                background-color: #fffef0;
                font-size: 10px;
            }
            QLineEdit:focus {
                border-color: #003366;
                background-color: #fffff0;
            }
            QTextEdit {
                border: 1px solid #c0c0c0;
                border-radius: 3px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 9px;
            }
            QPushButton {
                padding: 6px 16px;
                background-color: #003366;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                font-size: 10px;
            }
            QPushButton:hover {
                background-color: #004488;
            }
            QPushButton:pressed {
                background-color: #002244;
            }
            QPushButton:disabled {
                background-color: #999;
            }
            QPushButton#runBtn {
                background-color: #006633;
                font-size: 12px;
                padding: 10px 30px;
            }
            QPushButton#runBtn:hover {
                background-color: #008844;
            }
            QLabel {
                font-size: 10px;
            }
            QComboBox {
                padding: 4px 8px;
                border: 1px solid #c0c0c0;
                border-radius: 3px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                text-align: center;
                font-size: 10px;
            }
            QProgressBar::chunk {
                background-color: #003366;
                border-radius: 3px;
            }
            QStatusBar {
                background-color: #e0e4e8;
                font-size: 10px;
            }
            QTableWidget {
                gridline-color: #d0d0d0;
                font-size: 10px;
            }
            QHeaderView::section {
                background-color: #003366;
                color: white;
                padding: 4px;
                border: 1px solid #002244;
                font-weight: bold;
                font-size: 10px;
            }
        """)
    
    def _setup_ui(self):
        """Build the user interface."""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Title
        title = QLabel("OCTAGONAL SPREAD FOOTING ANALYSIS")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont('Segoe UI', 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #003366; padding: 5px; background-color: #d9e2f3; border-radius: 4px;")
        main_layout.addWidget(title)
        
        subtitle = QLabel("For Assumed Rigid Footing Base with Octagonal Pier | Based on PIP STE03350")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont('Segoe UI', 9))
        subtitle.setStyleSheet("color: #666; padding: 2px;")
        main_layout.addWidget(subtitle)
        
        # Tab widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Tab 1: Input Data
        self._create_input_tab()
        
        # Tab 2: Reaction Data
        self._create_reaction_tab()
        
        # Bottom controls
        bottom = QHBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumWidth(300)
        bottom.addWidget(self.progress_bar)
        
        bottom.addStretch()
        
        self.run_btn = QPushButton("▶  Run Analysis")
        self.run_btn.setObjectName("runBtn")
        self.run_btn.clicked.connect(self._run_analysis)
        bottom.addWidget(self.run_btn)
        
        main_layout.addLayout(bottom)
        
        # Status bar
        self.statusBar().showMessage("Ready. Fill in input data and paste reactions to begin.")
    
    def _create_input_tab(self):
        """Create Tab 1: Input Data."""
        tab = QWidget()
        scroll = QScrollArea()
        scroll.setWidget(tab)
        scroll.setWidgetResizable(True)
        layout = QVBoxLayout(tab)
        
        # Job Info
        job_group = QGroupBox("Job Information")
        job_layout = QGridLayout(job_group)
        
        self.job_name = QLineEdit()
        self.job_number = QLineEdit()
        self.subject = QLineEdit()
        self.originator = QLineEdit()
        self.checker = QLineEdit()
        
        job_layout.addWidget(QLabel("Job Name:"), 0, 0)
        job_layout.addWidget(self.job_name, 0, 1)
        job_layout.addWidget(QLabel("Subject:"), 0, 2)
        job_layout.addWidget(self.subject, 0, 3)
        job_layout.addWidget(QLabel("Job Number:"), 1, 0)
        job_layout.addWidget(self.job_number, 1, 1)
        job_layout.addWidget(QLabel("Originator:"), 1, 2)
        job_layout.addWidget(self.originator, 1, 3)
        job_layout.addWidget(QLabel("Checker:"), 2, 0)
        job_layout.addWidget(self.checker, 2, 1)
        
        layout.addWidget(job_group)
        
        # Use HBoxLayout for two columns
        params_layout = QHBoxLayout()
        
        # Footing Data
        ftg_group = QGroupBox("Footing Data")
        ftg_layout = QGridLayout(ftg_group)
        
        self.input_fields = {}
        
        footing_params = [
            ('Df', 'Ftg. Base Length, Df', '8.600', 'm'),
            ('Tf', 'Ftg. Base Thickness, Tf', '0.700', 'm'),
            ('Dp', 'Oct. Pier Length, Dp', '8.600', 'm'),
            ('hp', 'Oct. Pier Height, hp', '0.500', 'm'),
            ('gamma_c', 'Concrete Unit Wt., γc', '24.000', 'kN/m³'),
            ('Ds', 'Soil Depth, Ds (default)', '0.000', 'm'),
            ('gamma_s', 'Soil Unit Wt., γs', '18.000', 'kN/m³'),
            ('Kp', 'Pass. Press. Coef., Kp', '0.000', ''),
            ('mu', 'Coef. of Base Friction, μ', '0.400', ''),
            ('Q', 'Uniform Surcharge, Q', '0.000', 'kN/m²'),
            ('q_allow', 'SB Capacity, q_allow', '150.000', 'kN/m²'),
        ]
        
        for i, (key, label, default, unit) in enumerate(footing_params):
            ftg_layout.addWidget(QLabel(label), i, 0)
            field = QLineEdit(default)
            field.setMaximumWidth(100)
            field.setAlignment(Qt.AlignmentFlag.AlignRight)
            self.input_fields[key] = field
            ftg_layout.addWidget(field, i, 1)
            ftg_layout.addWidget(QLabel(unit), i, 2)
        
        params_layout.addWidget(ftg_group)
        
        # Node selection
        node_group = QGroupBox("Analysis Settings")
        node_layout = QGridLayout(node_group)
        
        node_layout.addWidget(QLabel("Node Selection:"), 0, 0)
        self.node_combo = QComboBox()
        self.node_combo.addItems(["End Node", "Start Node"])
        node_layout.addWidget(self.node_combo, 0, 1)
        
        node_layout.addWidget(QLabel("Output File:"), 1, 0)
        self.output_path = QLineEdit("analysis.xlsx")
        node_layout.addWidget(self.output_path, 1, 1)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_output)
        node_layout.addWidget(browse_btn, 1, 2)
        
        node_layout.setRowStretch(2, 1)
        params_layout.addWidget(node_group)
        
        layout.addLayout(params_layout)
        layout.addStretch()
        
        self.tabs.addTab(scroll, "📋 Input Data")
    
    def _create_reaction_tab(self):
        """Create Tab 2: Reaction Data."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Instructions
        info = QLabel(
            "Paste STAAD Pro reaction data below (columns: Beam, L/C, Node, Fx, Fy, Fz, Mx, My, Mz).\n"
            "Include both Node 1 and Node 2 rows. The tool will filter based on your End/Start selection."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #666; padding: 5px; background: #f8f8f8; border-radius: 3px;")
        layout.addWidget(info)
        
        # Text area for pasting
        self.reaction_text = QTextEdit()
        self.reaction_text.setPlaceholderText(
            "Paste reaction data here...\n"
            "Example:\n"
            "1  101 LC101  1  5377.911  0  0  0  0  0\n"
            "1  101 LC101  2  -3164.750  0  0  0  0  0\n"
            "..."
        )
        layout.addWidget(self.reaction_text, stretch=3)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        save_btn = QPushButton("💾  Save & Parse Reactions")
        save_btn.clicked.connect(self._save_reactions)
        btn_layout.addWidget(save_btn)
        
        self.edit_ds_btn = QPushButton("📝  Edit Ds Values")
        self.edit_ds_btn.clicked.connect(self._edit_ds_values)
        self.edit_ds_btn.setEnabled(False)
        btn_layout.addWidget(self.edit_ds_btn)
        
        clear_btn = QPushButton("🗑  Clear")
        clear_btn.setStyleSheet("background-color: #993333;")
        clear_btn.clicked.connect(lambda: self.reaction_text.clear())
        btn_layout.addWidget(clear_btn)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Parsed data preview table
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(['LC', 'LC Name', 'P (kN)', 'H (kN)', 'M (kN-m)', 'Ds (m)', 'Node'])
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.preview_table.setMaximumHeight(250)
        layout.addWidget(self.preview_table, stretch=1)
        
        self.tabs.addTab(tab, "📊 Reaction Data")
    
    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Analysis Report", "analysis.xlsx", "Excel Files (*.xlsx)")
        if path:
            self.output_path.setText(path)
    
    def _get_input_params(self):
        """Read and validate input parameters."""
        params = {}
        for key, field in self.input_fields.items():
            try:
                params[key] = float(field.text())
            except ValueError:
                raise ValueError(f"Invalid value for {key}: '{field.text()}'")
        return params
    
    def _save_reactions(self):
        """Parse pasted reaction data and show preview."""
        raw = self.reaction_text.toPlainText()
        if not raw.strip():
            QMessageBox.warning(self, "Warning", "No reaction data to parse. Please paste data first.")
            return
        
        use_end = self.node_combo.currentIndex() == 0
        self.reactions = parse_reactions(raw, use_end_node=use_end)
        
        if not self.reactions:
            QMessageBox.warning(self, "Warning", 
                "Could not parse any valid reaction data.\n\n"
                "Expected format (space or tab separated):\n"
                "Beam  L/C  Node  Fx  Fy  Fz  Mx  My  Mz")
            return
        
        # Extract unique LC numbers
        self.parsed_lc_list = [r['lc'] for r in self.reactions]
        
        # Initialize DS mapping (default 1.0)
        for lc in self.parsed_lc_list:
            if lc not in self.ds_mapping:
                self.ds_mapping[lc] = 1.0
        
        # Update preview table
        self.preview_table.setRowCount(len(self.reactions))
        from utils.analysis import compute_load_from_reaction
        for i, rxn in enumerate(self.reactions):
            P, H, M = compute_load_from_reaction(rxn)
            self.preview_table.setItem(i, 0, QTableWidgetItem(str(rxn['lc'])))
            self.preview_table.setItem(i, 1, QTableWidgetItem(rxn['lc_name']))
            self.preview_table.setItem(i, 2, QTableWidgetItem(f"{P:.3f}"))
            self.preview_table.setItem(i, 3, QTableWidgetItem(f"{H:.3f}"))
            self.preview_table.setItem(i, 4, QTableWidgetItem(f"{M:.3f}"))
            ds = self.ds_mapping.get(rxn['lc'], 1.0)
            self.preview_table.setItem(i, 5, QTableWidgetItem(f"{ds:.1f}"))
            self.preview_table.setItem(i, 6, QTableWidgetItem(str(rxn['node'])))
        
        self.edit_ds_btn.setEnabled(True)
        
        n = len(self.reactions)
        self.statusBar().showMessage(f"✓ Parsed {n} load combination(s). Click 'Edit Ds Values' to set soil depth per LC.")
        
        # Show DS editing dialog
        self._edit_ds_values()
    
    def _edit_ds_values(self):
        """Open dialog to edit Ds values for each LC."""
        if not self.parsed_lc_list:
            QMessageBox.warning(self, "Warning", "No load combinations parsed yet.")
            return
        
        dialog = DsEditDialog(self.parsed_lc_list, self.ds_mapping, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.ds_mapping = dialog.get_ds_mapping()
            
            # Update preview table Ds column
            for i, rxn in enumerate(self.reactions):
                ds = self.ds_mapping.get(rxn['lc'], 1.0)
                self.preview_table.setItem(i, 5, QTableWidgetItem(f"{ds:.1f}"))
            
            self.statusBar().showMessage("✓ Ds values updated. Ready to run analysis.")
    
    def _run_analysis(self):
        """Run the analysis."""
        # Validate inputs
        try:
            input_params = self._get_input_params()
        except ValueError as e:
            QMessageBox.warning(self, "Input Error", str(e))
            return
        
        if not self.reactions:
            QMessageBox.warning(self, "Warning", 
                "No reaction data loaded.\n"
                "Please go to 'Reaction Data' tab, paste data, and click 'Save & Parse'.")
            return
        
        # Disable run button
        self.run_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self.reactions))
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Running analysis...")
        
        use_end = self.node_combo.currentIndex() == 0
        
        # Run in background thread
        self.thread = AnalysisThread(input_params, self.reactions, self.ds_mapping, use_end)
        self.thread.progress.connect(self._on_progress)
        self.thread.finished.connect(self._on_finished)
        self.thread.error.connect(self._on_error)
        self.thread.start()
    
    def _on_progress(self, current, total, msg):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.statusBar().showMessage(f"Processing {msg} ({current}/{total})")
    
    def _on_finished(self, result):
        self.analysis_result = result
        self.run_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # Export to Excel
        output_path = self.output_path.text()
        if not output_path:
            output_path = "analysis.xlsx"
        
        # Make absolute path if relative
        if not os.path.isabs(output_path):
            output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', output_path)
            output_path = os.path.normpath(output_path)
        
        job_info = {
            'job_name': self.job_name.text(),
            'job_number': self.job_number.text(),
            'subject': self.subject.text(),
            'originator': self.originator.text(),
            'checker': self.checker.text(),
        }
        
        try:
            export_analysis_xlsx(result, output_path, job_info)
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to save Excel file:\n{e}")
            return
        
        # Show summary
        ratios = result.get('max_ratios', {})
        ctrl_lc = result.get('controlling_lc', 'N/A')
        r_max = ratios.get('Ratio_max', 0)
        status = "OK ✓" if r_max <= 1 else "NG ✗"
        
        msg = (
            f"Analysis Complete!\n\n"
            f"Controlling Load Case: {ctrl_lc}\n"
            f"Ratio Max: {r_max:.4f} ({status})\n"
            f"  - Overturning: {ratios.get('Ratio_OT', 0):.4f}\n"
            f"  - Sliding: {ratios.get('Ratio_SLD', 0):.4f}\n"
            f"  - Soil BC: {ratios.get('Ratio_SBC', 0):.4f}\n\n"
            f"Results saved to:\n{output_path}"
        )
        
        self.statusBar().showMessage(f"✓ Analysis complete. Ratio max = {r_max:.4f} ({status})")
        QMessageBox.information(self, "Analysis Complete", msg)
    
    def _on_error(self, error_msg):
        self.run_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Error during analysis.")
        QMessageBox.critical(self, "Analysis Error", f"An error occurred:\n{error_msg}")
