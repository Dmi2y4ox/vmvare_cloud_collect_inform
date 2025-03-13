import os

# Получаем базовую директорию (где находится config.py)
base_dir = os.path.dirname(os.path.abspath(__file__))

# Файл загруженный из VMware Cloud
input_file = os.path.join(base_dir, '..', 'excel', 'data-export.xlsx')

# Файл, в который будет записана информация
output_file = os.path.join(base_dir, '..', 'excel', 'data-import.xlsx')