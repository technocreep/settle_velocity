import numpy as np
import math
import pandas as pd
import xlwings as xw 
from openpyxl import Workbook

  
    # INITIAL DATA
min_1 = input('Название и плотность (кг/м3) минерала #1 (Золото 17000): ').split(' ')
min_2 = input('Название и плотность (кг/м3) минерала #2 (Кварц 2650): ').split(' ')
rho_1 = float(min_1[1])      
rho_2 = float(min_2[1])       


w_percentage = float(input('Содержание твердого, %: '))   
krup0 = float(input('min крупность (мkм): '))
krup1 = float(input('max крупность (мkм): '))

w = w_percentage/100
m = 0.001
m0 = m*math.pow(10, 3*w)

p1 = rho_1
p11 = rho_2
p2 = 1050
g = 9.81

# TABLE SIZE nXn
while True:
    try:
        n = int(input('Размер таблицы, n-значений: '))
        break
    except Exception:
        print('n должно быть целое')
        


def round_000(i):
    return round(i*1000)/1000


# PARTICLE SIZE
size = np.array([int(i) for i in np.linspace(krup0, krup1, n)])

# INTERGROWTH DENSITY 
p3 = np.array([round_000(i) for i in (np.linspace(rho_1, rho_2, n))])

# DATA TABLE FOR SOLUTION
df = pd.DataFrame(columns=[i for i in p3], index=[i for i in size])

for i in p3:
    re2y = math.pi*(size/1000000)**3*(i-1050)*g*1050/(6*m0*m0)
    re = 0.084*re2y**0.811
    v0 = re*m0/((size/1000000)*p2)
    df[i] = v0

# EXPORT TO EXCELL FILE
    # Create file
filename = str(input('Назовите сохраняемый файл'))
wb = Workbook()
wb.save('{}.xlsx'.format(filename))
wb.close()
    # save to created file
wb = xw.Book('{}.xlsx'.format(filename))
sht = wb.sheets['sheet']
sht.range('A1').value = df
