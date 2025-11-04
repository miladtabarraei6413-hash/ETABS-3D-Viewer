import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import numpy as np

# مسیر فایل اکسس خروجی ETABS
mdb_file = r"C:\Users\CLINET1\Documents\ETABS\Model.accdb"

# اتصال به دیتابیس اکسس
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={mdb_file};"
)
conn = pyodbc.connect(conn_str)

# خواندن جدول مختصات گره‌ها
nodes = pd.read_sql("SELECT * FROM [Point Object Connectivity]", conn)

# خواندن ستون‌ها، تیرها و (در صورت وجود) مهاربندها
columns = pd.read_sql("SELECT * FROM [Column Object Connectivity]", conn)
beams   = pd.read_sql("SELECT * FROM [Beam Object Connectivity]", conn)
# اگر جدول مهاربند دارید، همینطور:
# braces = pd.read_sql("SELECT * FROM [Brace Object Connectivity]", conn)

# ترسیم سازه سه‌بعدی
fig = plt.figure(figsize=(10,8))
ax = fig.add_subplot(111, projection='3d')

# تابع کمکی برای ترسیم المان‌ها
def plot_element(row, start_col, end_col, color, linewidth=2):
    start = nodes.loc[nodes['UniqueName'] == row[start_col]]
    end   = nodes.loc[nodes['UniqueName'] == row[end_col]]
    if not start.empty and not end.empty:
        ax.plot(
            [start['X'].values[0], end['X'].values[0]],
            [start['Y'].values[0], end['Y'].values[0]],
            [start['Z'].values[0], end['Z'].values[0]],
            color=color, linewidth=linewidth
        )

# ترسیم ستون‌ها
for _, row in columns.iterrows():
    plot_element(row, 'UniquePtI', 'UniquePtJ', color='red', linewidth=3)

# ترسیم تیرها
for _, row in beams.iterrows():
    plot_element(row, 'UniquePtI', 'UniquePtJ', color='blue', linewidth=2)

# اگر مهاربند دارید، مثل زیر اضافه کنید
# for _, row in braces.iterrows():
#     plot_element(row, 'UniquePtI', 'UniquePtJ', color='green', linewidth=2)

# برچسب محورها
ax.set_xlabel('X')
ax.set_ylabel('Y')
ax.set_zlabel('Z')
ax.set_title('ETABS Model Viewer')

# تنظیم مقیاس برابر برای تمام محورها
x_limits = ax.get_xlim3d()
y_limits = ax.get_ylim3d()
z_limits = ax.get_zlim3d()

max_range = np.array([x_limits[1]-x_limits[0], y_limits[1]-y_limits[0], z_limits[1]-z_limits[0]]).max()
mid_x = (x_limits[0] + x_limits[1]) * 0.5
mid_y = (y_limits[0] + y_limits[1]) * 0.5
mid_z = (z_limits[0] + z_limits[1]) * 0.5

ax.set_xlim3d(mid_x - max_range/2, mid_x + max_range/2)
ax.set_ylim3d(mid_y - max_range/2, mid_y + max_range/2)
ax.set_zlim3d(mid_z - max_range/2, mid_z + max_range/2)

plt.show()
