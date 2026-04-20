from .models import CutPiece, PipeAssignment


def expand_cut_requirements(cut_requirements: dict[int, int]) -> list[CutPiece]:
    """
    Converts {length: quantity} into a flat list of CutPiece objects.

    Example:
    {3000: 2, 2000: 1}
    ->
    [
        CutPiece(id="cut_01", length=3000),
        CutPiece(id="cut_02", length=3000),
        CutPiece(id="cut_03", length=2000)
    ]
    """
    cuts = []
    cut_number = 1

    for length, quantity in cut_requirements.items():
        for _ in range(quantity):
            cut_id = f"cut_{cut_number:02d}"
            cuts.append(CutPiece(id=cut_id, length=length))
            cut_number += 1

    return cuts


def expand_leftover_pipes(leftover_pipes: dict[int, int]) -> list[int]:
    """
    Converts {length: quantity} into a flat list of leftover pipes.

    Example:
    {3500: 3}
    -> [3500, 3500, 3500]
    """
    pipes = []

    for length, quantity in leftover_pipes.items():
        pipes.extend([length] * quantity)

    return pipes


def get_initial_pipes(leftovers: list[int], include_leftovers: bool) -> list[int]:
    """
    Returns the starting list of available pipes.
    Includes leftovers only if enabled.
    """
    if include_leftovers:
        return leftovers.copy()
    return []


def can_fit_cut(pipe: PipeAssignment, cut_length: int, kerf: int) -> bool:
    """Return whether a given cut length fits into this pipe assignment."""
    if not pipe.cuts:
        return cut_length <= pipe.remaining_length
    return cut_length + kerf <= pipe.remaining_length


def add_cut_to_pipe(pipe: PipeAssignment, cut: CutPiece, kerf: int) -> None:
    """Place a cut into a pipe assignment and update used/remaining lengths."""
    if pipe.cuts:
        pipe.used_length += kerf
    pipe.used_length += cut.length
    pipe.remaining_length = pipe.original_length - pipe.used_length
    pipe.cuts.append(cut)


def create_pipe_assignment(pipe_id: str, original_length: int, source: str) -> PipeAssignment:
    """Create a new pipe assignment for either leftover or new stock pipe."""
    return PipeAssignment(id=pipe_id, source=source, original_length=original_length)


def assign_cuts_to_pipes(
    cuts: list[CutPiece],
    existing_pipes: list[int],
    stock_pipe_length: int,
    kerf: int,
) -> list[PipeAssignment]:
    """Pack required cuts into existing leftovers first, then open new pipes."""
    assignments: list[PipeAssignment] = [
        create_pipe_assignment(f"leftover_{index + 1}", length, source="leftover")
        for index, length in enumerate(existing_pipes)
    ]

    for cut in sorted(cuts, key=lambda cut_item: cut_item.length, reverse=True):
        best_pipe: PipeAssignment | None = None
        best_remaining: int | None = None

        for pipe in assignments:
            if can_fit_cut(pipe, cut.length, kerf):
                if best_pipe is None or pipe.remaining_length < best_remaining:
                    best_pipe = pipe
                    best_remaining = pipe.remaining_length

        if best_pipe is None:
            new_pipe_id = f"pipe_{len([p for p in assignments if p.source == 'new']) + 1}"
            best_pipe = create_pipe_assignment(new_pipe_id, stock_pipe_length, source="new")
            assignments.append(best_pipe)

        add_cut_to_pipe(best_pipe, cut, kerf)

    return assignments


def plan_cuts_for_job(job: "CuttingJob") -> tuple[list[PipeAssignment], int]:
    """Create a cutting plan from a CuttingJob."""
    cuts = expand_cut_requirements(job.cut_requirements)
    leftovers = expand_leftover_pipes(job.leftover_pipes)
    initial_pipes = get_initial_pipes(leftovers, job.include_leftovers)
    assignments = assign_cuts_to_pipes(cuts, initial_pipes, job.stock_pipe_length, job.kerf)
    new_pipes = [pipe for pipe in assignments if pipe.source == "new"]
    return assignments, len(new_pipes)