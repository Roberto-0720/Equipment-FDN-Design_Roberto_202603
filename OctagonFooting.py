"""
Octagonal Spread Footing Design Tool
=====================================
Main entry point for the application.

Based on PIP STE03350 - "Vertical Vessel Foundation Design Guide" (2007)
For Assumed Rigid Footing Base with Octagonal Pier
Supporting a Vertical Round Tank, Vessel, or Stack.
"""

import sys
import os

# Add the project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    
    # Set application-wide font
    font = QFont('Segoe UI', 9)
    app.setFont(font)
    
    # Set application properties
    app.setApplicationName("Octagonal Footing Design Tool")
    app.setOrganizationName("Engineering Tools")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
