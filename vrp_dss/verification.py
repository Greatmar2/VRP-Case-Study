from time import perf_counter
from typing import List, Dict, Tuple

from openpyxl import load_workbook

import data


def get_data_from_sheet(row: int) -> str:
    """Gets the output text file data from a certain row on the solve times summary sheet."""
    data_column = "G"

    workbook = load_workbook(filename="Solve Times Summary.xlsx", read_only=True, data_only=True)
    pass


def write_data_to_sheet(row: int, math_objective: float, meta_routes: str, meta_time: float, meta_objective: float):
    """Writes metaheuristic output data to the solve times summary sheet."""
    math_objective_col = "I"
    meta_routes_col = "J"
    meta_time_col = "K"
    meta_objective_col = "L"

    pass


def extract_data_from_output(math_output: str) -> Tuple[data.Data, Dict[int, List[List]]]:
    """Takes the mathematical output text and converts it to input data for the metaheuristic."""
    pass


if __name__ == "__main__":
    """Verify the metaheuristic against all mathematical instances."""
    start_row = 2
    end_row = 61

    for row in range(start_row, end_row):
        text_data = get_data_from_sheet(row)
        input_data, exact_solution = extract_data_from_output(text_data)
        data.run_data = input_data

        from main import Runner

        start_time = perf_counter()
        runner = Runner(5000, 1800, use_multiprocessing=False)
        best_solution = runner.run()
        end_time = perf_counter()

        from evaluate import evaluate_solution

        eval_results = evaluate_solution(exact_solution)

        write_data_to_sheet(eval_results["penalised_cost"], best_solution.pretty_route_output(), end_time - start_time,
                            best_solution.get_penalised_cost(1))
