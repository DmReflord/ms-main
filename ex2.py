from pulp import *

def solve_ex2():
    
    #Создаем функцию и задаем задачу, название задачи, LpMinimize - минимизация ЦФ
    prob = LpProblem("Ballast_Traffic", LpMinimize)
    
    # Переменные решения, название, lowBound - неотрицательность (>=0), cat='Continuous' - непрерывные переменные
    x11 = LpVariable("x11", lowBound=0, cat='Continuous')  # Объем балласта 1-го карьера на 1-ый участок 
    x12 = LpVariable("x12", lowBound=0, cat='Continuous')  # Объем балласта 1-го карьера на 2-ый участок 
    x21 = LpVariable("x21", lowBound=0, cat='Continuous')  # Объем балласта 2-го карьера на 1-ый участок
    x22 = LpVariable("x22", lowBound=0, cat='Continuous')  # Объем балласта 2-го карьера на 2-ый участок
    
    # Целевая функция
    prob += 10*x11 + 9*x12 + 4*x21 + 5*x22, "Total_Profit"
    
    # Ограничения
    prob += x11 + x12 <= 35, "Power first provider"
    prob += x21 + x22 <= 25, "Power second provider"
    prob += x11 + x21 >= 27, "Demand first field"
    prob += x12 + x22 >= 15, "Demand second field"
    
    # Решение задачи
    prob.solve()
    
    # Вывод результатов
    print("=" * 50)
    print("РЕШЕНИЕ")
    print("=" * 50)
    print(f"Статус: {LpStatus[prob.status]}")
    print(f"Минимальные затраты: {value(prob.objective):.2f} тыс. ден. ед.")
    print(f"\nОптимальные объемы:")
    for v in prob.variables():
        print(f"{v.name} = {v.varValue:.2f} тыс. м³")
    
    # Анализ дефицитных ресурсов
    """ print(f"\nАнализ ограничений:")
    for name, constraint in prob.constraints.items():
        print(f"{name}: {constraint.value():.2f} (свободный ресурс: {-constraint.value():.2f})")
 """

solve_ex2()