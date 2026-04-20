from .models import CuttingJob


def get_sample_job() -> CuttingJob:
    return CuttingJob(
        cut_requirements={
            3000: 2,
            4000: 1,
            5000: 1,
            2000: 3
        },
        stock_pipe_length=5000,
        
        leftover_pipes={
            3500: 3
        },
        
        kerf=5,
        include_leftovers=True
    )