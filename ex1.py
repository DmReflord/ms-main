from pulp import *

def solve_with_pulp():
    
    #Создаем функцию и задаем задачу, название задачи, LpMaximize - максимизация ЦФ
    prob = LpProblem("Ballast_Production", LpMaximize)
    
    # Переменные решения, название, lowBound - неотрицательность (>=0), cat='Continuous' - непрерывные переменные
    x1 = LpVariable("x1", lowBound=0, cat='Continuous')  # Объем 1-го типа балласта
    x2 = LpVariable("x2", lowBound=0, cat='Continuous')  # Объем 2-го типа балласта  
    x3 = LpVariable("x3", lowBound=0, cat='Continuous')  # Объем 3-го типа балласта
    
    # Целевая функция
    prob += 6*x1 + 10*x2 + 12*x3, "Total_Profit"
    
    # Ограничения
    prob += 13*x1 + 27*x2 + 24*x3 <= 230, "Excavators"
    prob += 8*x1 + 4*x2 + 6*x3 <= 50, "Bulldozers"
    prob += 50*x1 + 30*x2 + 50*x3 <= 610, "Labor"
    prob += x2 <= 8, "Demand_x2"
    prob += x3 <= 5, "Demand_x3"
    
    # Решение задачи
    prob.solve()
    
    # Вывод результатов
    print("=" * 50)
    print("РЕШЕНИЕ")
    print("=" * 50)
    print(f"Статус: {LpStatus[prob.status]}")
    print(f"Максимальная прибыль: {value(prob.objective):.2f} тыс. ден. ед.")
    print(f"\nОптимальные объемы:")
    for v in prob.variables():
        print(f"{v.name} = {v.varValue:.2f} тыс. м³")
    
    # Анализ дефицитных ресурсов
    """ print(f"\nАнализ ограничений:")
    for name, constraint in prob.constraints.items():
        print(f"{name}: {constraint.value():.2f} (свободный ресурс: {-constraint.value():.2f})")
 """

solve_with_pulp()