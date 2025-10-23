from pulp import *

def solve_assignment():
    # Создаем задачу - задача о назначениях
    prob = LpProblem("Brigade_Assignment", LpMinimize)
    
    # Создаем переменные - БИНАРНЫЕ (0 или 1)
    variables = {}
    for i in range(1, 5):  # 4 бригады
        for j in range(1, 5):  # 4 объекта
            variables[f'x{i}{j}'] = LpVariable(f"x{i}{j}", cat='Binary')
    
    # Целевая функция - минимизация времени
    prob += (30*variables['x11'] + 40*variables['x12'] + 50*variables['x13'] + 60*variables['x14'] +
             37*variables['x21'] + 47*variables['x22'] + 57*variables['x23'] + 58*variables['x24'] +
             27*variables['x31'] + 44*variables['x32'] + 49*variables['x33'] + 57*variables['x34'] +
             35*variables['x41'] + 37*variables['x42'] + 47*variables['x43'] + 63*variables['x44']), "Total_Time"
    
    # Ограничения: каждая бригада назначается на один объект
    prob += variables['x11'] + variables['x12'] + variables['x13'] + variables['x14'] == 1, "Brigade1_assignment"
    prob += variables['x21'] + variables['x22'] + variables['x23'] + variables['x24'] == 1, "Brigade2_assignment"
    prob += variables['x31'] + variables['x32'] + variables['x33'] + variables['x34'] == 1, "Brigade3_assignment"
    prob += variables['x41'] + variables['x42'] + variables['x43'] + variables['x44'] == 1, "Brigade4_assignment"
    
    # Ограничения: на каждый объект назначается одна бригада
    prob += variables['x11'] + variables['x21'] + variables['x31'] + variables['x41'] == 1, "Object1_assignment"
    prob += variables['x12'] + variables['x22'] + variables['x32'] + variables['x42'] == 1, "Object2_assignment"
    prob += variables['x13'] + variables['x23'] + variables['x33'] + variables['x43'] == 1, "Object3_assignment"
    prob += variables['x14'] + variables['x24'] + variables['x34'] + variables['x44'] == 1, "Object4_assignment"
    
    # Решение задачи
    prob.solve()
    
    # Вывод результатов
    print("=" * 60)
    print("ОПТИМАЛЬНОЕ РАСПРЕДЕЛЕНИЕ БРИГАД ПО ОБЪЕКТАМ")
    print("=" * 60)
    print(f"Статус: {LpStatus[prob.status]}")
    print(f"Минимальное суммарное время: {value(prob.objective)} единиц времени")
    
    print("\nНазначения бригад:")
    assignments = []
    for i in range(1, 5):
        for j in range(1, 5):
            if variables[f'x{i}{j}'].varValue == 1:
                time = 0
                if f'x{i}{j}' == 'x11': time = 30
                elif f'x{i}{j}' == 'x12': time = 40
                elif f'x{i}{j}' == 'x13': time = 50
                elif f'x{i}{j}' == 'x14': time = 60
                elif f'x{i}{j}' == 'x21': time = 37
                elif f'x{i}{j}' == 'x22': time = 47
                elif f'x{i}{j}' == 'x23': time = 57
                elif f'x{i}{j}' == 'x24': time = 58
                elif f'x{i}{j}' == 'x31': time = 27
                elif f'x{i}{j}' == 'x32': time = 44
                elif f'x{i}{j}' == 'x33': time = 49
                elif f'x{i}{j}' == 'x34': time = 57
                elif f'x{i}{j}' == 'x41': time = 35
                elif f'x{i}{j}' == 'x42': time = 37
                elif f'x{i}{j}' == 'x43': time = 47
                elif f'x{i}{j}' == 'x44': time = 63
                
                assignments.append(f"Бригада {i} → Объект {j} (время: {time})")
    
    for assignment in sorted(assignments):
        print(assignment)

# Альтернативный вариант с использованием списков (более компактный)
def solve_assignment_compact():
    prob = LpProblem("Brigade_Assignment", LpMinimize)
    
    # Матрица времени
    time_matrix = [
        [30, 40, 50, 60],  # Бригада 1
        [37, 47, 57, 58],  # Бригада 2
        [27, 44, 49, 57],  # Бригада 3
        [35, 37, 47, 63]   # Бригада 4
    ]
    
    # Создаем переменные
    x = [[LpVariable(f"x{i+1}{j+1}", cat='Binary') for j in range(4)] for i in range(4)]
    
    # Целевая функция
    prob += lpSum(time_matrix[i][j] * x[i][j] for i in range(4) for j in range(4))
    
    # Ограничения
    for i in range(4):  # Каждая бригада на одном объекте
        prob += lpSum(x[i][j] for j in range(4)) == 1
    
    for j in range(4):  # На каждый объект одна бригада
        prob += lpSum(x[i][j] for i in range(4)) == 1
    
    # Решение
    prob.solve()
    
    print("\n" + "=" * 60)
    print("КОМПАКТНОЕ РЕШЕНИЕ")
    print("=" * 60)
    print(f"Статус: {LpStatus[prob.status]}")
    print(f"Минимальное суммарное время: {value(prob.objective)} единиц времени")
    
    print("\nНазначения:")
    for i in range(4):
        for j in range(4):
            if x[i][j].varValue == 1:
                print(f"Бригада {i+1} → Объект {j+1} (время: {time_matrix[i][j]})")

# Запуск решений
if __name__ == "__main__":
    solve_assignment()
    solve_assignment_compact()