from .models import CutPiece, PipeAssignment
from typing import List, Tuple

"""
================================================================================
CUTTING STOCK OPTIMIZER - OPTIMIZATION ALGORITHM MODULE
================================================================================

This module contains the core optimization logic for the cutting stock problem.
It implements an intelligent bin-packing algorithm that:

1. Expands input data (quantities to individual items)
2. Prioritizes zero-remainder solutions (perfect fits)
3. Attempts to reuse existing leftover pipes before ordering new stock
4. Calculates material efficiency metrics

THE BIN PACKING PROBLEM:
- Items: CutPiece objects that need to be cut
- Bins: Pipes (either leftover or newly ordered stock)
- Objective: Minimize number of bins used, maximize material utilization
- Constraint: Kerf (blade width) reduces available material for each cut

STRATEGY:
1. Expand cut requirements: [("A", 100, 2)] -> [CutPiece("A", 100), CutPiece("A", 100)]
2. Try to fit cuts into leftover pipes first (before ordering new stock)
3. For each pipe, find the best combination of cuts that fit
4. Prioritize zero-remainder solutions (perfect fits with no waste)
5. Fall back to minimum-remainder solutions if perfect fit not found
6. Account for kerf (material lost during cuts)

KERF ACCOUNTING:
- When n cuts are made from one pipe, n-1 kerf segments are consumed
- Example: 3 cuts = 2 kerf separations
- Total used = cut1 + kerf + cut2 + kerf + cut3
================================================================================
"""


def expand_cut_requirements(cut_requirements: List[Tuple[str, int, int]]) -> List[CutPiece]:
    """
    EXPAND CUT REQUIREMENTS: Convert quantities to individual cut pieces.
    
    Takes compressed format and creates one CutPiece for each unit.
    
    INPUT FORMAT (compressed with quantities):
    [("A", 3000, 2), ("B", 2000, 1)]
    
    MEANING:
    - "A" with length 3000: need 2 cuts
    - "B" with length 2000: need 1 cut
    
    OUTPUT (expanded list, one per cut):
    [
        CutPiece(id="A", length=3000),
        CutPiece(id="A", length=3000),
        CutPiece(id="B", length=2000)
    ]
    
    WHY EXPAND?
    - Easier for the optimizer to work with individual cuts
    - Algorithm iterates over individual cuts, not quantities
    - Makes kerf accounting cleaner
    
    Args:
        cut_requirements: List of (label, length, quantity) tuples
    
    Returns:
        List of CutPiece objects, one per unit
    """
    cuts = []

    for label, length, quantity in cut_requirements:
        for _ in range(quantity):
            cuts.append(CutPiece(id=label, length=length))

    return cuts


def expand_leftover_pipes(leftover_pipes: List[Tuple[str, int, int]]) -> List[Tuple[str, int]]:
    """
    EXPAND LEFTOVER PIPES: Convert quantities to individual pipes.
    
    Similar to expand_cut_requirements but for leftover pipes.
    
    INPUT (compressed):
    [("Left Over 01", 3500, 3)]
    
    MEANING:
    - "Left Over 01" with length 3500: have 3 of them
    
    OUTPUT (expanded):
    [("Left Over 01", 3500), ("Left Over 01", 3500), ("Left Over 01", 3500)]
    
    Args:
        leftover_pipes: List of (label, length, quantity) tuples
    
    Returns:
        List of (label, length) tuples, one per pipe
    """
    pipes = []

    for label, length, quantity in leftover_pipes:
        pipes.extend([(label, length)] * quantity)

    return pipes


def get_initial_pipes(leftovers: List[Tuple[str, int]], include_leftovers: bool) -> List[Tuple[str, int]]:
    """
    GET INITIAL PIPES: Determine which existing pipes are available.
    
    This creates the starting inventory of pipes available to use
    before ordering new stock.
    
    LOGIC:
    - If include_leftovers is True: Return all leftover pipes (can reuse them)
    - If include_leftovers is False: Return empty list (ignore leftovers)
    
    This respects the user's choice on whether to consider existing stock.
    
    Args:
        leftovers: List of (label, length) tuples from expand_leftover_pipes
        include_leftovers: Boolean flag from user input
    
    Returns:
        List of (label, length) tuples available for use, or empty list
    """
    if include_leftovers:
        return leftovers.copy()
    return []


def can_fit_cut(pipe: PipeAssignment, cut_length: int, kerf: int) -> bool:
    """
    CAN FIT CUT: Check if a cut fits in the remaining space of a pipe.
    
    This is a quick validation check before attempting to add a cut.
    
    LOGIC:
    - If pipe is empty (no cuts yet): check if cut_length <= remaining_length
      (no kerf needed before first cut)
    - If pipe has cuts: check if cut_length + kerf <= remaining_length
      (need space for the cut PLUS the kerf that precedes it)
    
    KERF ACCOUNTING:
    When adding a cut to a pipe with existing cuts:
    - A kerf separator is needed between the last cut and this new cut
    - So we need: cut_length + kerf available space
    
    Args:
        pipe: The PipeAssignment to check
        cut_length: Length of the cut to add (in mm)
        kerf: Blade width (in mm)
    
    Returns:
        True if cut fits, False otherwise
    """
    if not pipe.cuts:
        # First cut: no kerf needed before it
        return cut_length <= pipe.remaining_length
    # Subsequent cuts: need space for kerf separator + cut
    return cut_length + kerf <= pipe.remaining_length


def add_cut_to_pipe(pipe: PipeAssignment, cut: CutPiece, kerf: int) -> None:
    """
    ADD CUT TO PIPE: Place a cut in a pipe and update length tracking.
    
    This method modifies the pipe in-place, adding a cut and recalculating
    used/remaining lengths.
    
    LOGIC:
    1. If pipe already has cuts, add kerf for the separator before this cut
    2. Add the cut length
    3. Recalculate remaining_length = original_length - used_length
    4. Add cut to the cuts list
    
    LENGTH TRACKING:
    - used_length: Cumulative material consumed (cuts + kerfs)
    - remaining_length: Material still available for more cuts
    
    EXAMPLE (5mm kerf):
    - Add first cut (1000): used = 1000, remaining = 4000
    - Add second cut (800): used = 1000 + 5 + 800 = 1805, remaining = 3195
    - Add third cut (500): used = 1805 + 5 + 500 = 2310, remaining = 2690
    
    Args:
        pipe: The PipeAssignment to modify
        cut: The CutPiece to add
        kerf: Blade width (in mm)
    """
    # Add kerf for separator if this is not the first cut
    if pipe.cuts:
        pipe.used_length += kerf
    
    # Add the cut length
    pipe.used_length += cut.length
    
    # Recalculate remaining space
    pipe.remaining_length = pipe.original_length - pipe.used_length
    
    # Add cut to the list
    pipe.cuts.append(cut)


def finalize_pipe_assignments(assignments: list[PipeAssignment], kerf: int) -> None:
    """
    FINALIZE PIPE ASSIGNMENTS: Account for the final kerf after last cut.
    
    After all cuts are assigned, we need to add kerf for the final cut that
    separates the remainder from the last usable cut.
    
    LOGIC:
    For each pipe that has cuts AND has remainder material:
    - Add kerf for the final cut that removes the remainder
    - Recalculate remaining_length
    
    WHY?
    During optimization, we don't account for the final kerf that's needed
    to separate the remainder. This happens at the very end, so we add it here.
    
    EXAMPLE:
    - Pipe with 3 cuts: kerf added between cuts during optimization
    - But the final kerf (to cut off remainder) is added here
    - So if remainder > 0: we add one more kerf to used_length
    
    Args:
        assignments: List of PipeAssignment objects to finalize
        kerf: Blade width (in mm)
    """
    for pipe in assignments:
        # Only if pipe has cuts and leftover material
        if pipe.cuts and pipe.remaining_length > 0:
            # Add the final kerf that cuts off the remainder
            pipe.used_length += kerf
            # Recalculate remaining (which will reduce by one kerf)
            pipe.remaining_length = pipe.original_length - pipe.used_length


def create_pipe_assignment(pipe_id: str, original_length: int, source: str) -> PipeAssignment:
    """
    CREATE PIPE ASSIGNMENT: Create a new pipe to fill with cuts.
    
    Factory function to create empty PipeAssignment objects.
    Encapsulates the creation logic so it's consistent throughout.
    
    Args:
        pipe_id: Unique identifier for this pipe
        original_length: Total length of the pipe (in mm)
        source: Either "new" (ordered) or "leftover" (existing)
    
    Returns:
        New PipeAssignment with no cuts, ready to be filled
    """
    return PipeAssignment(id=pipe_id, source=source, original_length=original_length)


def find_best_combination(available_cuts: list[CutPiece], pipe_length: int, kerf: int) -> tuple[list[CutPiece], int]:
    """
    FIND BEST COMBINATION: Determine optimal cuts for one pipe.
    
    This is the core optimization logic. It tries different combinations
    of cuts and selects the best one based on waste minimization.
    
    PRIORITY SYSTEM (in order):
    1. ZERO REMAINDER: Perfect fit with no waste (best)
    2. MINIMIZED REMAINDER: If no perfect fit, minimize leftover material
    3. MAXIMIZED CUTS: If same remainder, use more cuts (better utilization)
    
    ALGORITHM:
    - Uses recursive backtracking to try all combinations
    - Prunes branches when: perfect fit found (can stop searching)
    - Evaluates: remainder, quantity of cuts
    
    WHY RECURSIVE?
    - Tries all possible subsets of cuts that fit in the pipe
    - Recursive: add each possible cut and try more cuts in the remaining space
    - When a perfect fit is found, stops immediately (no need to check others)
    
    EXAMPLE (pipe 5000mm, kerf 5mm):
    - Available cuts: [1000, 1000, 800, 800, 500]
    - Perfect fit found: [1000, 1000, 800, 800, 300] uses exactly 5000
    - Algorithm: tries [1000], [1000, 1000], [1000, 1000, 800], etc.
    
    Args:
        available_cuts: List of cuts to choose from
        pipe_length: Total length available (in mm)
        kerf: Blade width (in mm)
    
    Returns:
        Tuple of (best_cuts, remainder_length)
        - best_cuts: List of selected CutPiece objects
        - remainder_length: Unused space in the pipe
    """
    best_cuts: list[CutPiece] = []
    best_remainder = pipe_length
    
    # RECURSIVE HELPER: Try different combinations of cuts
    def try_combinations(remaining_cuts: list[CutPiece], current_combo: list[CutPiece], used_length: int) -> None:
        nonlocal best_cuts, best_remainder
        
        # CALCULATE ACTUAL USED LENGTH (including kerfs)
        actual_used = used_length
        if len(current_combo) > 1:
            # If multiple cuts: kerf count = number of cuts - 1
            actual_used += (len(current_combo) - 1) * kerf
        
        remainder = pipe_length - actual_used
        
        # EVALUATE THIS COMBINATION: Is it better than current best?
        is_better = False
        if best_remainder > 0 and remainder == 0:
            # Found perfect fit (zero remainder) - always best
            is_better = True
        elif remainder >= 0:
            # Combination fits in pipe
            if best_remainder == 0 and remainder == 0:
                # Both are perfect fit: prefer more cuts (better utilization)
                is_better = len(current_combo) > len(best_cuts)
            elif best_remainder > 0 and remainder > 0:
                # Both have remainder: prefer smaller remainder, or more cuts if same remainder
                is_better = remainder < best_remainder or (remainder == best_remainder and len(current_combo) > len(best_cuts))
            elif best_remainder > 0:
                # Current is fit, best is not: current is better
                is_better = used_length > sum(cut.length for cut in best_cuts)
        
        if is_better:
            best_cuts = current_combo.copy()
            best_remainder = remainder
        
        # PRUNE SEARCH: Stop if we found perfect fit (can't do better)
        if best_remainder == 0:
            return
        
        # TRY ADDING EACH REMAINING CUT
        for i, cut in enumerate(remaining_cuts):
            new_used = used_length + cut.length
            # Calculate kerf needed if we add this cut
            new_kerf = (len(current_combo) * kerf) if len(current_combo) > 0 else 0
            
            # Only try if it fits
            if new_used + new_kerf <= pipe_length:
                new_combo = current_combo + [cut]
                # Remaining cuts: all except the one we're using
                new_remaining = remaining_cuts[:i] + remaining_cuts[i+1:]
                # Recurse: try more cuts
                try_combinations(new_remaining, new_combo, new_used)
    
    # Start recursive search with empty combination
    try_combinations(available_cuts, [], 0)
    return best_cuts, best_remainder


def assign_cuts_to_pipes(
    cuts: list[CutPiece],
    existing_pipes: List[Tuple[str, int]],
    stock_pipe_length: int,
    kerf: int,
) -> list[PipeAssignment]:
    """
    ASSIGN CUTS TO PIPES: Run the main optimization algorithm.
    
    This is the orchestration function that:
    1. Creates PipeAssignment objects for all existing pipes
    2. Tries to fit all cuts into existing pipes first
    3. Creates new pipes for cuts that don't fit anywhere
    4. Returns complete cutting plan
    
    OPTIMIZATION STRATEGY:
    - Prioritize using existing leftover pipes (before ordering new stock)
    - For each pipe, find the best combination of cuts using find_best_combination()
    - Keep adding pipes until all cuts are assigned
    
    ALGORITHM:
    1. Create PipeAssignment for each existing leftover pipe
    2. For each leftover pipe: find best cuts that fit, assign them, remove from list
    3. While cuts remain:
       a. Create a new PipeAssignment (new stock)
       b. Find best cuts that fit this pipe
       c. Assign them, remove from list
    4. Finalize all pipes (account for final kerf)
    
    Args:
        cuts: List of all CutPiece objects to assign
        existing_pipes: List of (label, length) tuples for leftover pipes
        stock_pipe_length: Length of new pipes to order
        kerf: Blade width
    
    Returns:
        List of PipeAssignment objects representing the complete cutting plan
    """
    # CREATE PIPE ASSIGNMENTS for existing/leftover pipes
    assignments: list[PipeAssignment] = [
        create_pipe_assignment(label, length, source="leftover")
        for label, length in existing_pipes
    ]
    
    # REMAINING CUTS to assign (gets smaller as we assign cuts)
    remaining_cuts = cuts.copy()
    
    # FILL EXISTING PIPES FIRST
    for pipe in assignments:
        if not remaining_cuts:
            break
        # Find best combination of remaining cuts that fit in this pipe
        best_cuts, _ = find_best_combination(remaining_cuts, pipe.original_length, kerf)
        # Assign the best cuts to this pipe
        for cut in best_cuts:
            add_cut_to_pipe(pipe, cut, kerf)
            # Remove assigned cuts from remaining
            remaining_cuts.remove(cut)
    
    # CREATE NEW PIPES for remaining cuts
    while remaining_cuts:
        # Create new pipe with standard length
        new_pipe_id = f"pipe_{len([p for p in assignments if p.source == 'new']) + 1}"
        new_pipe = create_pipe_assignment(new_pipe_id, stock_pipe_length, source="new")
        assignments.append(new_pipe)
        
        # Find best combination for this new pipe
        best_cuts, _ = find_best_combination(remaining_cuts, stock_pipe_length, kerf)
        # Assign cuts
        for cut in best_cuts:
            add_cut_to_pipe(new_pipe, cut, kerf)
            remaining_cuts.remove(cut)
    
    # FINALIZE: Add kerf for final cuts
    finalize_pipe_assignments(assignments, kerf)
    return assignments


def plan_cuts_for_job(job: "CuttingJob") -> tuple[list[PipeAssignment], int]:
    """
    PLAN CUTS FOR JOB: Main entry point for optimization.
    
    Takes a CuttingJob specification and returns the complete cutting plan.
    This is the function called by the UI to compute results.
    
    PROCESS:
    1. Expand cut requirements (quantities -> individual cuts)
    2. Expand leftover pipes (quantities -> individual pipes)
    3. Get initial pipe inventory (respect include_leftovers flag)
    4. Run assignment algorithm
    5. Count how many NEW pipes need to be ordered
    
    Args:
        job: CuttingJob object with all input specifications
    
    Returns:
        Tuple of:
        - assignments: List of PipeAssignment (the cutting plan)
        - new_pipe_count: How many NEW pipes to order (for ordering)
    """
    # Expand quantities to individual items
    cuts = expand_cut_requirements(job.cut_requirements)
    leftovers = expand_leftover_pipes(job.leftover_pipes)
    
    # Determine which existing pipes can be used
    initial_pipes = get_initial_pipes(leftovers, job.include_leftovers)
    
    # Run the optimization algorithm
    assignments = assign_cuts_to_pipes(cuts, initial_pipes, job.stock_pipe_length, job.kerf)
    
    # Count new pipes (for ordering)
    new_pipes = [pipe for pipe in assignments if pipe.source == "new"]
    
    return assignments, len(new_pipes)


def calculate_efficiency(assignments: list[PipeAssignment]) -> dict:
    """
    CALCULATE EFFICIENCY: Compute material utilization metrics.
    
    Analyzes the cutting plan and calculates how efficiently material
    was used, including waste and efficiency percentage.
    
    METRICS RETURNED:
    - total_pipes: Number of pipes used (new + leftover)
    - total_material: Total pipe length used (sum of all pipe lengths)
    - used_material: Material in actual cuts (just the cut portions)
    - effective_used: Material in cuts + kerf (total consumed)
    - waste_material: Remainders + inefficiency (total_material - effective_used)
    - efficiency: Percentage (effective_used / total_material * 100)
    
    INTERPRETATION:
    - Efficiency 95%: Very good, minimal waste
    - Efficiency 80%: Good, acceptable waste
    - Efficiency 60%: Fair, significant waste
    
    EXAMPLE (5mm kerf):
    - 2 pipes of 5000mm = 10000mm total
    - Cut 1000 + Cut 1000 + Cut 800 + Cut 800 = 3600mm
    - Kerf: 3 cuts from pipe 1, 1 cut from pipe 2 = 4 kerfs = 20mm
    - Effective used: 3600 + 20 = 3620mm
    - Waste: 10000 - 3620 = 6380mm
    - Efficiency: 36.2%
    
    Args:
        assignments: List of PipeAssignment objects from the optimizer
    
    Returns:
        Dictionary with efficiency metrics
    """
    # Sum all pipe lengths
    total_material = sum(pipe.original_length for pipe in assignments)
    
    # Sum just the cut portions (no kerf included)
    used_material = sum(cut.length for pipe in assignments for cut in pipe.cuts)
    
    # Sum used length which includes kerf
    effective_used = sum(pipe.used_length for pipe in assignments)
    
    # Waste = total - used (includes remainder pieces and kerf losses)
    waste_material = total_material - effective_used
    
    # Calculate efficiency percentage
    efficiency = (effective_used / total_material * 100) if total_material > 0 else 0
    
    return {
        "total_pipes": len(assignments),
        "total_material": total_material,
        "used_material": used_material,
        "effective_used": effective_used,
        "waste_material": waste_material,
        "efficiency": efficiency,
    }

