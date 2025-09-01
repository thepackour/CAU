import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import numpy
from scipy.interpolate import interp1d
import matplotlib.font_manager as fm
import mydata # 로컬 파일

load_wb = openpyxl.load_workbook(mydata.excel_path, data_only=True)
load_ws = load_wb['Sheet1']
ttf_path = mydata.ttf_path
font_name = fm.FontProperties(fname=ttf_path, size=10).get_name()

fe = fm.FontEntry(
    fname=ttf_path,
    name=font_name
)

fm.fontManager.ttflist.append(fe)
matplotlib.rcParams.update({'font.size': 8, 'font.family': font_name})

# 셀 입력칸 -------------------------------------------------------

x_input = [
    load_ws['C26': 'C30'],
    load_ws['C45': 'C49'],
    load_ws['C64': 'C68']
]
y_input = [
    load_ws['D26': 'D30'],
    load_ws['D45': 'D49'],
    load_ws['D64': 'D68']
]
x_labels = ["작용한 힘 (F) [N]",
            "회전축으로부터 작용점까지의 거리 (r_M) [cm]",
            "위치벡터와 작용한 힘이 이루는 각의 sine"]
y_label = "토크 (τ) [N·m]"

titles = ["작용한 힘과 토크의 관계 그래프",
          "회전축으로부터 힘의 작용점까지의 거리와 토크의 관계 그래프",
          "위치벡터와 작용한 힘이 이루는 각의 sine값과 토크의 관계 그래프"]

# 셀 입력칸 -------------------------------------------------------

for i in range(3):
    x_getcell = x_input[i]
    y_getcell = y_input[i]
    x_label = x_labels[i]
    title = titles[i]

    x_list = []
    y_list = []

    for x_cell in x_getcell:
        for x in x_cell:
            x_list.append(x.value)

    for y_cell in y_getcell:
        for y in y_cell:
            y_list.append(y.value)

    numpy_x = numpy.array(x_list)
    numpy_y = numpy.array(y_list)

    x_new = np.linspace(numpy_x.min(), numpy_x.max(), 500)

    f = interp1d(numpy_x, numpy_y, kind="quadratic")
    y_smooth=f(x_new)

    plt.subplot(3,1,i+1)
    plt.plot(x_list, y_list, 'r-', x_new, y_smooth, 'b--')
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    plt.title(title)
    plt.grid(True)

plt.tight_layout(h_pad=2)
plt.show()
