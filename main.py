from src.cutting_stock.data import get_sample_job
from src.cutting_stock.utils import plan_cuts_for_job


def main():
    job = get_sample_job()
    assignments, new_pipe_count = plan_cuts_for_job(job)

    print(f"Pipes to order = {new_pipe_count}")
    for index, pipe in enumerate(assignments, start=1):
        if not pipe.cuts:
            continue
        cut_entries = ", ".join(f"{cut.id}({cut.length})" for cut in pipe.cuts)
        source_label = "new" if pipe.source == "new" else "leftover"
        print(
            f"Pipe {index} ({source_label}, length {pipe.original_length}, remaining {pipe.remaining_length}): {cut_entries}"
        )

    print("\nJob summary:")
    print(f"  stock pipe length: {job.stock_pipe_length}")
    print(f"  kerf: {job.kerf}")
    print(f"  include leftovers: {job.include_leftovers}")


if __name__ == "__main__":
    main()