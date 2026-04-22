"""
================================================================================
CUTTING STOCK OPTIMIZER - APPLICATION ENTRY POINT
================================================================================

This is the main entry point for the entire application. When you run:
    python main.py

This file:
1. Adds the src/ directory to Python's import path
2. Imports the UI module
3. Launches the graphical user interface

APPLICATION STRUCTURE:
- main.py (this file): Entry point and imports
- src/cutting_stock/ui.py: Complete graphical user interface
- src/cutting_stock/models.py: Data structures (CuttingJob, PipeAssignment, etc.)
- src/cutting_stock/utils.py: Optimization algorithm and helper functions
- src/cutting_stock/data.py: Sample test data (deprecated, not used by UI)

DEPENDENCIES:
- Python 3.7+ (required)
- tkinter: GUI framework (usually included with Python)
- No external packages needed

TO RUN:
    python main.py

This opens the Cutting Stock Optimizer window where users can:
1. Enter stock pipe length and blade width (kerf)
2. Add required cuts with quantities
3. Optionally add leftover pipes to reuse
4. Click "Compute Cutting Plan" to optimize
5. View results and material efficiency
6. Save/load cutting plans to CSV files
================================================================================
"""

import sys

# ADD SRC TO IMPORT PATH
# This allows importing from src/cutting_stock/ without installing as a package
sys.path.append('src')

# IMPORT UI MODULE
# This imports the main graphical interface
from cutting_stock.ui import main as run_ui


if __name__ == "__main__":
    # LAUNCH APPLICATION
    # Calls the main() function from ui.py which:
    # 1. Sets up DPI awareness on Windows
    # 2. Creates the Tkinter root window
    # 3. Creates the CuttingStockUI controller
    # 4. Starts the event loop (window stays open until user closes it)
    run_ui()
