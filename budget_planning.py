"""
Budget Planning Optimization using Linear Programming

This script optimizes budget allocation across projects to maximize ROI
while respecting budget constraints and management guidelines.
"""

import pandas as pd
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpStatus, value
import matplotlib.pyplot as plt


def load_projects():
    """Load project data from Excel file."""
    df_p = pd.read_excel("warehouses.xlsx", sheet_name="projects")

    for col in ["TURNOVER", "YEAR 1", "YEAR 2", "YEAR 3", "ROI"]:
        df_p[col] = df_p[col].fillna(0).astype(int)

    # Create project description linked to customer
    df_p["PROJECT DESCRIPTION"] = (
        df_p["PROJECT DESCRIPTION"].astype(str) + "-" + df_p["CUSTOMER"]
    )

    print(f"{len(df_p):,} projects in your list")
    return df_p


def optimize_roi(df_p, budget_limits):
    """Model 1: Maximize ROI without management constraints."""
    model = LpProblem("Budget_Planning_MaxROI", LpMaximize)

    projects = list(df_p["PROJECT DESCRIPTION"].values)
    roi = list(df_p["ROI"].values)

    # Decision variables: binary for each project
    P = LpVariable.dicts("P", range(len(projects)), lowBound=0, cat="Binary")

    # Objective: maximize total ROI
    model += lpSum([roi[i] * P[i] for i in range(len(projects))])

    # Constraints: budget limits per year
    for j in range(3):
        model += (
            lpSum(
                [P[i] * df_p.loc[i, f"YEAR {j+1}"] for i in range(len(projects))]
            )
            <= budget_limits[j]
        )

    model.solve()

    return model, P, projects


def optimize_with_guidelines(df_p, budget_limits, objective_minimums):
    """Model 2: Maximize ROI with management guidelines."""
    model = LpProblem("Budget_Planning_Guidelines", LpMaximize)

    projects = list(df_p["PROJECT DESCRIPTION"].values)
    roi = list(df_p["ROI"].values)

    P = LpVariable.dicts("P", range(len(projects)), lowBound=0, cat="Binary")

    # Objective: maximize total ROI
    model += lpSum([roi[i] * P[i] for i in range(len(projects))])

    # Budget constraints per year
    for j in range(3):
        model += (
            lpSum(
                [P[i] * df_p.loc[i, f"YEAR {j+1}"] for i in range(len(projects))]
            )
            <= budget_limits[j]
        )

    # Management objective constraints
    O = df_p["OPERATIONAL EXCELLENCE"].values * 1
    S = df_p["SUSTAINABILITY"].values * 1
    D = df_p["DIGITAL TRANSFORMATION"].values * 1

    model += (
        lpSum([P[i] * O[i] * df_p.loc[i, "TOTAL"] for i in range(len(projects))])
        >= objective_minimums[0]
    )
    model += (
        lpSum([P[i] * S[i] * df_p.loc[i, "TOTAL"] for i in range(len(projects))])
        >= objective_minimums[1]
    )
    model += (
        lpSum([P[i] * D[i] * df_p.loc[i, "TOTAL"] for i in range(len(projects))])
        >= objective_minimums[2]
    )

    model.solve()

    return model, P, projects


def display_results(model, P, df_p, budget_limits, model_name):
    """Display optimization results."""
    print("\n" + "=" * 60)
    print(f"RESULTS: {model_name}")
    print("=" * 60)
    print(f"Status: {LpStatus[model.status]}")
    print(f"Return of Investment = {int(value(model.objective)):,} Euros")

    max_budget = sum(budget_limits) / 1e6
    actual_budget = round(
        sum([P[i].varValue * df_p.loc[i, "TOTAL"] for i in range(len(P))]) / 1e6, 2
    )
    project_count = int(sum([P[i].varValue for i in range(len(P))]))

    print(
        f"{project_count}/{len(df_p)} Projects Accepted with Budget {actual_budget:,}/{max_budget:,} M€"
    )

    # Get selected projects
    selected = [i for i in range(len(P)) if P[i].varValue == 1]
    print(f"\nSelected projects: {len(selected)}")

    return selected


def main():
    """Main function to run budget planning optimization."""
    # Load data
    df_p = load_projects()

    # Budget limits per year (Year 1, Year 2, Year 3)
    budget_limits = [1250000, 1500000, 1750000]

    # Model 1: Maximize ROI only
    print("\n" + "-" * 60)
    print("Running Model 1: Maximize ROI")
    print("-" * 60)
    model1, P1, projects = optimize_roi(df_p, budget_limits)
    display_results(model1, P1, df_p, budget_limits, "Maximize ROI")

    # Model 2: With management guidelines
    print("\n" + "-" * 60)
    print("Running Model 2: With Management Guidelines")
    print("-" * 60)
    objective_minimums = [1000000, 1000000, 1000000]  # Min budget for each objective
    model2, P2, projects = optimize_with_guidelines(
        df_p, budget_limits, objective_minimums
    )
    display_results(model2, P2, df_p, budget_limits, "With Management Guidelines")


if __name__ == "__main__":
    main()
