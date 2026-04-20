from .models import CutPiece


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