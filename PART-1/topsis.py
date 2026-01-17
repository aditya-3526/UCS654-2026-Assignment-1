import sys
import pandas as pd
import numpy as np

def error(msg):
    print(f"Error: {msg}")
    sys.exit(1)

def topsis(input_file, weights, impacts, output_file):
    # Read input file
    try:
        data = pd.read_excel(input_file)
    except FileNotFoundError:
        error("Input file not found")

    if data.shape[1] < 3:
        error("Input file must contain at least 3 columns")

    # Criteria columns (2nd to last)
    criteria = data.iloc[:, 1:]

    # Check numeric values
    if not np.all(criteria.applymap(np.isreal)):
        error("All criteria values must be numeric")

    # Parse weights and impacts
    weights = weights.split(',')
    impacts = impacts.split(',')

    if len(weights) != criteria.shape[1]:
        error("Number of weights must equal number of criteria")

    if len(impacts) != criteria.shape[1]:
        error("Number of impacts must equal number of criteria")

    try:
        weights = np.array(weights, dtype=float)
    except:
        error("Weights must be numeric")

    for i in impacts:
        if i not in ['+', '-']:
            error("Impacts must be either + or -")

    # Step 1: Normalize
    norm = np.sqrt((criteria ** 2).sum())
    normalized = criteria / norm

    # Step 2: Weighted normalization
    weighted = normalized * weights

    # Step 3: Ideal best and worst
    ideal_best = []
    ideal_worst = []

    for i in range(len(impacts)):
        if impacts[i] == '+':
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    ideal_best = np.array(ideal_best)
    ideal_worst = np.array(ideal_worst)

    # Step 4: Distances
    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    # Step 5: TOPSIS score
    score = dist_worst / (dist_best + dist_worst)

    # Step 6: Rank
    rank = score.rank(ascending=False)

    # Output
    data['Topsis Score'] = score
    data['Rank'] = rank.astype(int)

    data.to_csv(output_file, index=False)
    print("TOPSIS analysis completed successfully.")

if __name__ == "__main__":
    if len(sys.argv) != 5:
        error("Usage: python topsis.py <InputFile> <Weights> <Impacts> <OutputFile>")

    topsis(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
