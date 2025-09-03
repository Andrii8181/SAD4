from scipy.stats import shapiro

def check_normality(data):
    """
    Перевірка нормальності методом Шапіро-Вілка
    """
    W, p = shapiro(data)
    return {"W": W, "p": p, "normal": p > 0.05}
 
