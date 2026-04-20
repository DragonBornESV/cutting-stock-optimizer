from src.cutting_stock.data import get_sample_job
from src.cutting_stock.utils import expand_cut_requirements
from src.cutting_stock.utils import expand_leftover_pipes
from src.cutting_stock.utils import get_initial_pipes


def main():
    job = get_sample_job()

    cuts = expand_cut_requirements(job.cut_requirements)

    print("Expanded cut requirements:")
    for cut in cuts:
        print(f"  {cut.id}: {cut.length}")

    leftovers = expand_leftover_pipes(job.leftover_pipes)
    print("Expanded leftovers:", leftovers)

    pipes = get_initial_pipes(leftovers, job.include_leftovers)
    print("Initial pipes:", pipes)

    print("Cut requirements:", job.cut_requirements)
    print("Stock pipe length:", job.stock_pipe_length)
    print("Leftover pipes:", job.leftover_pipes)
    print("Kerf:", job.kerf)


if __name__ == "__main__":
    main()