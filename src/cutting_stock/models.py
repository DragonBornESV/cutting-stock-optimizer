from dataclasses import dataclass, field
from typing import List, Tuple

"""
================================================================================
CUTTING STOCK OPTIMIZER - DATA MODELS MODULE
================================================================================

This module defines the core data structures used throughout the cutting stock
optimization system. All models are implemented as Python dataclasses for
clean, type-safe code.

KEY MODELS:
1. CutPiece: Represents a single cut needed from the stock
2. CuttingJob: Represents the complete input specification for one optimization
3. PipeAssignment: Represents a pipe with its assigned cuts (output of optimizer)

DATA FORMAT:
- All measurements are in MILLIMETERS (mm)
- Cut requirements are stored as tuples: (label, length, quantity)
- Leftover pipes are stored as tuples: (label, length, quantity)
- Kerf is the width of material lost during each cut (blade thickness)

EXAMPLE USAGE:
    job = CuttingJob(
        cut_requirements=[("Door", 1000, 2), ("Frame", 800, 1)],
        stock_pipe_length=5000,
        leftover_pipes=[("Leftover", 3500, 1)],
        kerf=5,
        include_leftovers=True
    )
================================================================================
"""


@dataclass
class CutPiece:
    """
    REPRESENTS A SINGLE REQUIRED CUT.
    
    This is the atomic unit of the cutting job - one specific piece that needs
    to be cut from the stock pipe.
    
    ATTRIBUTES:
    - id: Human-readable label for this cut (e.g., "Door Panel 01", "A")
    - length: Required length of this cut in millimeters
    
    NOTES:
    - Multiple identical CutPiece objects can exist if the same length/label
      is needed multiple times (stored via quantity in CuttingJob)
    - This is what the optimizer actually assigns to pipes
    """

    id: str
    length: int


@dataclass
class CuttingJob:
    """
    REPRESENTS ONE COMPLETE CUTTING JOB SPECIFICATION.
    
    This is the INPUT to the optimization algorithm. It contains everything
    needed to describe a cutting problem: what cuts are needed, how much
    stock is available, and what parameters to use.
    
    ATTRIBUTES:
    - cut_requirements: List of (label, length, quantity) tuples
      - label: Name/identifier for this cut type (user-provided or auto-generated)
      - length: Length needed in mm
      - quantity: How many cuts of this length are needed
      - Example: [("Door", 1000, 2), ("Frame", 800, 1)]
    
    - stock_pipe_length: Length of one standard pipe to purchase (in mm)
      - This is the "bin size" in bin packing terms
      - All new pipes ordered have this length
      - Example: 5000 mm standard stock length
    
    - leftover_pipes: List of (label, length, quantity) tuples for EXISTING stock
      - These are pipes already on hand that can be reused
      - The optimizer tries to use these BEFORE ordering new stock
      - Format same as cut_requirements
      - Example: [("LeftOver 01", 3500, 1), ("LeftOver 02", 2000, 2)]
    
    - kerf: Blade/cutting tool width in mm
      - Material lost during each cut due to blade thickness
      - Example: 5 mm kerf means each cut loses 5mm of material
      - Important for accurate calculations: actual usable length = cut length + kerf
      - If 3 cuts are made, total kerf loss = 2 * kerf (n-1 kerfs for n cuts)
    
    - include_leftovers: Boolean flag
      - If True: Optimizer can use leftover pipes to minimize new orders
      - If False: Optimizer only uses new pipes
      - Allows user to ignore leftovers in optimization
    
    DEFAULTS:
    - All lists default to empty (user must populate)
    - include_leftovers defaults to True (try to reuse existing stock)
    """

    cut_requirements: List[Tuple[str, int, int]] = field(default_factory=list)
    stock_pipe_length: int = 0
    leftover_pipes: List[Tuple[str, int, int]] = field(default_factory=list)
    kerf: int = 0
    include_leftovers: bool = True


@dataclass
class PipeAssignment:
    """
    REPRESENTS ONE PIPE WITH ITS ASSIGNED CUTS.
    
    This is the OUTPUT of the optimization algorithm. It represents a single
    pipe and what cuts will be made from it.
    
    ATTRIBUTES:
    - id: Unique identifier for this pipe (e.g., "pipe_1", "LeftOver 01")
    
    - source: Where this pipe comes from
      - "new": Newly purchased stock
      - "leftover": From the leftover_pipes list (reused existing stock)
    
    - original_length: Total length of this pipe before cutting (in mm)
      - For new pipes: equals stock_pipe_length from the job
      - For leftovers: equals the length specified in leftover_pipes
    
    - cuts: List of CutPiece objects assigned to this pipe
      - These are the cuts that will be made from this pipe
      - Stored in the order they should be cut
      - Example: [CutPiece("Door", 1000), CutPiece("Frame", 800)]
    
    - used_length: Total material used in this pipe (in mm)
      - Includes all cut lengths PLUS kerf losses between cuts
      - Calculated during optimization and finalization
      - Formula: sum(cut lengths) + (number_of_cuts - 1) * kerf
      - Example with 2 cuts (1000, 800) and kerf=5: 1000 + 5 + 800 = 1805
    
    - remaining_length: Unused material in this pipe (in mm)
      - Calculated automatically in __post_init__
      - Formula: original_length - used_length
      - 0 means pipe is completely used (zero remainder)
      - Positive value means material left over after all cuts
    
    LIFECYCLE:
    1. Created by optimizer with id, source, original_length
    2. remaining_length auto-calculated as original_length
    3. During optimization, cuts are added with add_cut_to_pipe()
    4. used_length is updated as cuts are added
    5. remaining_length is kept in sync with used_length changes
    
    EXAMPLE:
    pipe = PipeAssignment(id="pipe_1", source="new", original_length=5000)
    # At creation: remaining_length = 5000, used_length = 0, cuts = []
    # After adding cuts: might have remaining_length = 200 (waste)
    """

    id: str
    source: str
    original_length: int
    cuts: list[CutPiece] = field(default_factory=list)
    used_length: int = 0
    remaining_length: int = field(init=False)

    def __post_init__(self) -> None:
        """
        AUTOMATIC INITIALIZATION: Set remaining_length when pipe is created.
        
        This runs automatically after __init__. We use it to ensure
        remaining_length is always initialized to match original_length,
        even though the user doesn't pass it as a parameter.
        
        This is a dataclass feature - fields marked with field(init=False)
        won't be parameters, but __post_init__ runs after initialization,
        allowing us to set them based on other field values.
        """
        self.remaining_length = self.original_length