import os
import pandas as pd

print("启动")
# 获取当前脚本所在文件夹路径
current_folder = os.path.dirname(os.path.abspath(__file__))

# 指定要遍历的文件夹路径

folder_path = current_folder.replace('\\', '/')
## 使用os.listdir()函数获取该文件夹下所有文件名
files = os.listdir(folder_path)

# 使用os.path.splitext()函数获取每个文件的文件名和扩展名
# 然后使用os.path.basename()函数获取不带扩展名的文件名
file_names = [os.path.splitext(os.path.basename(file))[0] for file in files]

# 移除不需要的字符 "abc"
file_names = [name.replace('abc', '') for name in file_names]

# 将所有文件名存储到一个pandas DataFrame中
df_files = pd.DataFrame(file_names, columns=['File Name'])

# 将DataFrame输出到一个CSV文件中
df_files.to_csv('file_names.csv', index=False)

print("启动完成")