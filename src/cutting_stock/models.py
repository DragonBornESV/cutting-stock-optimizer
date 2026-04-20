from dataclasses import dataclass, field
from typing import Dict


@dataclass
class CutPiece:
    """
    Represents one specific required cut.
    """

    id: str
    length: int


@dataclass
class CuttingJob:
    """
    Stores the input data for one cutting job.
    All values are in millimeters.
    """

    cut_requirements: Dict[int, int] = field(default_factory=dict)
    stock_pipe_length: int = 0
    leftover_pipes: Dict[int, int] = field(default_factory=dict)
    kerf: int = 0
    include_leftovers: bool = True


@dataclass
class PipeAssignment:
    """Represents one pipe used for cutting."""

    id: str
    source: str
    original_length: int
    cuts: list[CutPiece] = field(default_factory=list)
    used_length: int = 0
    remaining_length: int = field(init=False)

    def __post_init__(self) -> None:
        self.remaining_length = self.original_length