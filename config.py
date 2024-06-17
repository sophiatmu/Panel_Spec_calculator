# config.py
import numpy as np

# 默认参数值
PixelSize = None
VResolution = None
HResolution = None
PanelShape = None
CPL = None
R_optical = None
G_optical = None
B_optical = None
opticals = None
max_yield_locate = None
matrix = None
voltage_data = None
color_filter = None
temp_matrix = None
Color_Matching = None
Color_Space = None
MatrixRGB = None
Beta = None
dutyRGB = None
Number = None
CFL = None
Efficiency = None
ArrayW = None
matrix_max_table = None
matrix_wr_table = None

# 工作表对象
workbook = None
OpticalData_sheet = None
home_sheet = None
Grayscale_sheet = None
Gamma_sheet = None
R_spectrum = None
G_spectrum = None
B_spectrum = None
Spectrums = None
CFData_sheet = None
CIE_sheet = None

def init_parameters():
    global PixelSize, VResolution, HResolution, PanelShape, CPL
    global R_optical, G_optical, B_optical, opticals
    global max_yield_locate, matrix, voltage_data, color_filter
    global temp_matrix, Color_Matching, Color_Space, MatrixRGB
    global Beta, dutyRGB, Number, CFL, Efficiency, ArrayW, matrix_max_table,matrix_wr_table
    global workbook,OpticalData_sheet,home_sheet,Grayscale_sheet,Gamma_sheet
    global R_spectrum,G_spectrum,B_spectrum,Spectrums,CFData_sheet,CIE_sheet
    PixelSize = 0
    VResolution = 0
    HResolution = 0
    PanelShape = 0
    CPL = 0
    R_optical = np.zeros((110, 4))
    G_optical = np.zeros((110, 4))
    B_optical = np.zeros((110, 4))
    opticals = [R_optical, G_optical, B_optical]
    max_yield_locate = np.zeros(3)
    matrix = np.zeros((3, 3))
    voltage_data = np.zeros((201, 3))
    color_filter = np.zeros((201, 3))
    temp_matrix = np.zeros((201, 3))
    Color_Matching = np.zeros((201, 3))
    Color_Space = np.zeros((3, 3))
    Color_Space[0, :] = 1
    MatrixRGB = np.zeros((3, 3))
    MatrixRGB[1, :] = 1
    Beta = np.ones(3)
    dutyRGB = np.zeros(3)
    Number = np.zeros(3)  # LED Number per Pixel
    CFL = np.ones(3)  # Color Filter Transparency
    Efficiency = np.zeros(3)  # Efficiency after OC (%)
    ArrayW = np.zeros((3, 1))
    matrix_max_table = np.zeros((6, 3))
    matrix_wr_table = np.zeros((6, 3))
    matrix_wr_table = np.zeros((6, 3))

def set_workbook(wb):
    global workbook, OpticalData_sheet, home_sheet, Grayscale_sheet
    global Gamma_sheet, R_spectrum, G_spectrum, B_spectrum, Spectrums
    global CFData_sheet, CIE_sheet

    workbook = wb
    OpticalData_sheet = wb['OpticalData']
    home_sheet = wb['HaHa']
    Grayscale_sheet = wb['Grayscale']
    Gamma_sheet = wb['Gamma']
    R_spectrum = wb['RSpectrum']
    G_spectrum = wb['GSpectrum']
    B_spectrum = wb['BSpectrum']
    Spectrums = [R_spectrum, G_spectrum, B_spectrum]
    CFData_sheet = wb['CFData']
    CIE_sheet = wb['CIE1931']