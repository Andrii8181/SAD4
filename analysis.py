from scipy.stats import shapiro

def check_normality(data_columns):
    results = []
    for i, col in enumerate(data_columns):
        stat, p = shapiro(col)
        normal = "Нормальний розподіл" if p > 0.05 else "НЕ нормальний розподіл"
        results.append({
            "column": f"Фактор {i+1}",
            "W": round(stat, 4),
            "p": round(p, 4),
            "result": normal
        })
    return results
