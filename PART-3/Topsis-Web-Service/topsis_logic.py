import pandas as pd
import numpy as np

def run_topsis(input_file, weights, impacts, output_file):
    data = pd.read_excel(input_file)

    if data.shape[1] < 3:
        raise Exception("Input file must contain at least 3 columns")

    criteria = data.iloc[:, 1:]

    if not np.all(criteria.applymap(np.isreal)):
        raise Exception("Criteria values must be numeric")

    weights = list(map(float, weights.split(',')))
    impacts = impacts.split(',')

    if len(weights) != criteria.shape[1]:
        raise Exception("Weights count mismatch")

    if len(impacts) != criteria.shape[1]:
        raise Exception("Impacts count mismatch")

    for i in impacts:
        if i not in ['+', '-']:
            raise Exception("Impacts must be + or -")

    norm = np.sqrt((criteria ** 2).sum())
    normalized = criteria / norm
    weighted = normalized * weights

    ideal_best, ideal_worst = [], []

    for i in range(len(impacts)):
        if impacts[i] == '+':
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    score = dist_worst / (dist_best + dist_worst)
    rank = score.rank(ascending=False)

    data['Topsis Score'] = score
    data['Rank'] = rank.astype(int)

    data.to_csv(output_file, index=False)
