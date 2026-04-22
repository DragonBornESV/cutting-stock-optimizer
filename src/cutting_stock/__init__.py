"""
================================================================================
CUTTING STOCK OPTIMIZER PACKAGE
================================================================================

This package provides a complete solution for the cutting stock optimization
problem using a graphical user interface.

MODULES:
- models.py: Data structures (CuttingJob, PipeAssignment, CutPiece)
- utils.py: Optimization algorithm and helper functions
- ui.py: Tkinter graphical user interface
- data.py: Sample test data (deprecated, for reference only)

MAIN ENTRY POINT:
    python main.py

TYPICAL WORKFLOW:
1. User launches application: python main.py
2. Enters cutting job parameters:
   - Stock pipe length
   - Blade width (kerf)
   - Required cuts with quantities
   - Optional leftover pipes to reuse
3. Clicks "Compute Cutting Plan"
4. Application displays:
   - Number of pipes to order
   - Material efficiency percentage
   - Visual representation of how cuts are assigned
   - Detailed breakdown of each pipe
5. User can save/load plans in CSV format

OPTIMIZATION ALGORITHM:
- Bin packing with kerf (blade width) accounting
- Prioritizes zero-remainder solutions (perfect fits)
- Uses leftover pipes before ordering new stock
- Minimizes material waste
- Maximizes material utilization

See individual module docstrings for detailed documentation.
================================================================================
"""
