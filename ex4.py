import numpy as np
from pulp import *

# Данные из контрольного примера
count = [0, 17, 34, 51, 68]  # Количество рабочих

CMR = np.array([
    [0,  0,  0,  0],
    [8,  9,  7, 6],
    [14, 16, 16, 10],
    [24, 25, 22, 18],
    [32, 33, 30, 24]
])

# Бинарные переменные
v = np.array([
    [LpVariable("v11", cat='Binary'), LpVariable("v12", cat='Binary'), 
     LpVariable("v13", cat='Binary'), LpVariable("v14", cat='Binary')],
    [LpVariable("v21", cat='Binary'), LpVariable("v22", cat='Binary'), 
     LpVariable("v23", cat='Binary'), LpVariable("v24", cat='Binary')],
    [LpVariable("v31", cat='Binary'), LpVariable("v32", cat='Binary'), 
     LpVariable("v33", cat='Binary'), LpVariable("v34", cat='Binary')],
    [LpVariable("v41", cat='Binary'), LpVariable("v42", cat='Binary'), 
     LpVariable("v43", cat='Binary'), LpVariable("v44", cat='Binary')],
    [LpVariable("v51", cat='Binary'), LpVariable("v52", cat='Binary'), 
     LpVariable("v53", cat='Binary'), LpVariable("v54", cat='Binary')]
])

C = 68  # Всего рабочих

problem = LpProblem('Maximize_CMR', LpMaximize)

# Целевая функция
profit = lpSum(CMR[i][j] * v[i][j] for i in range(5) for j in range(4))
problem += profit

# Ограничение на общее количество рабочих
problem += (lpSum(count[i] * v[i][j] for i in range(5) for j in range(4)) == C)

# Каждому объекту назначается ровно одна группа рабочих
for j in range(4):
    problem += lpSum(v[i][j] for i in range(5)) == 1

# Решение
status = problem.solve()

print("Матрица распределения:")
print("=" * 50)
print("Объекты →", end=" ")
for j in range(4):
    print(f"  {j+1}  ", end=" ")
print("\n" + "=" * 50)

for i in range(5):
    print(f"Группа {i} ({count[i]} раб.) |", end=" ")
    for j in range(4):
        val = v[i][j].varValue  # Изменил имя переменной с value на val
        print(f" {val:5.1f} ", end=" ")
    print()

print("=" * 50)
print("Status:", LpStatus[status])
print("Result:")
for i in range(5):
    print(f"Группа {i}:", end=" ")
    for j in range(4):
        print(f"{v[i][j].varValue:>4}", end=" ")
    print()

# Используем pulp.value() явно
print("Максимальный объем СМР:", pulp.value(problem.objective))

# Проверка распределения рабочих
total_workers = 0
total_cmr = 0
for j in range(4):
    for i in range(5):
        if v[i][j].varValue == 1:
            total_workers += count[i]
            total_cmr += CMR[i][j]
            print(f"Объект {j+1}: {count[i]} рабочих, СМР = {CMR[i][j]} тыс.руб")

print(f"Всего распределено рабочих: {total_workers}")
print(f"Суммарный объем СМР: {total_cmr} тыс.руб")