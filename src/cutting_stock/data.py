"""
================================================================================
SAMPLE DATA MODULE (DEPRECATED)
================================================================================

This module provides sample/test data for the cutting stock optimizer.

NOTE: This module is NOT actively used by the application. All data entry
is now handled through the user interface (ui.py). This file is maintained
for reference and testing purposes only.

If you want to test the optimizer with sample data in code:
    from cutting_stock.data import get_sample_job
    from cutting_stock.utils import plan_cuts_for_job
    
    job = get_sample_job()
    assignments, new_pipes = plan_cuts_for_job(job)

See models.py for the CuttingJob data structure definition.
================================================================================
"""

from .models import CuttingJob


def get_sample_job() -> CuttingJob:
    """
    GET SAMPLE JOB: Return a sample CuttingJob for testing.
    
    This creates a realistic example cutting job with:
    - Multiple different cut sizes with various quantities
    - Standard pipe length (5000mm)
    - Leftover pipes available for reuse
    - Blade width/kerf specification
    - Leftover reuse enabled
    
    EXAMPLE PROBLEM:
    - Need to cut: 2x3000mm, 1x4000mm, 1x5000mm, 3x2000mm
    - Stock pipe: 5000mm each
    - Have leftover: 3x3500mm pipes
    - Blade width: 5mm
    
    This is a good test case because:
    - Mix of cut sizes
    - Some cuts fit perfectly in stock (5000mm)
    - Leftovers can be efficiently used
    - Multiple pipes likely needed
    
    Returns:
        CuttingJob object configured with sample data
    """
    return CuttingJob(
        # Required cuts as (label, length, quantity)
        cut_requirements=[
            ("3000mm", 3000, 2),  # Need 2 cuts of 3000mm
            ("4000mm", 4000, 1),  # Need 1 cut of 4000mm
            ("5000mm", 5000, 1),  # Need 1 cut of 5000mm (perfect fit!)
            ("2000mm", 2000, 3),  # Need 3 cuts of 2000mm
        ],
        # New pipe length when ordering stock
        stock_pipe_length=5000,
        # Leftover pipes as (label, length, quantity)
        leftover_pipes=[
            ("LeftOver 01", 3500, 3),  # Have 3 leftover pipes of 3500mm
        ],
        # Blade width (material lost per cut)
        kerf=5,
        # Whether to consider leftover pipes
        include_leftovers=True,
    )
