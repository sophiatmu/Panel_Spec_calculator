# Panel_Spec_calculator


## Requirements
```
pip install openpyxl
pip install numpy
```
## 主檔案main_v5.py

### 檔案中的路徑修改
```
wb = openpyxl.load_workbook('X098_v5.xlsm')  # 自己的初始檔案
save_wb = 'D://Muyun//cal_python//X098_v5_0606.xlsx'  # 要存的地方跟另存檔名
```

## Read_optical.py
用來寫入Optical資料
```
wb = openpyxl.load_workbook('X098_v4.xlsm')  # 需要存放資料到哪個excel表格
```

找到目錄中所有後綴為 '_R_Optical.xlsx'、'_G_Optical.xlsx、 '_B_Optical.xlsx' 的檔案
因此這個目錄只能有各1份RGB的optical資料
```
R_files = [filename for filename in os.listdir(directory) if filename.endswith('_R_Optical.xlsx')]
G_files = [filename for filename in os.listdir(directory) if filename.endswith('_G_Optical.xlsx')]
B_files = [filename for filename in os.listdir(directory) if filename.endswith('_B_Optical.xlsx')]

```
