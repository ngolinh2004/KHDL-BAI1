import pandas as pd
import matplotlib.pyplot as plt

from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# ====================================
# STYLE BIỂU ĐỒ
# ====================================

plt.style.use("ggplot")

# ====================================
# ĐỌC FILE EXCEL
# ====================================

file_path = "TỔNG HỢP ĐIỂM K58KTP.xlsx"

df = pd.read_excel(file_path, header=None)

print("========== KIỂM TRA DỮ LIỆU ==========")
print(df.head())

# ====================================
# CẤU TRÚC FILE
# ====================================

# Dòng 0 : Tiêu đề
# Dòng 1 : MSSV
# Dòng 2 : Họ tên
# Dòng 3 : STT
# Dòng 4 : Trống
# Dòng 5 trở đi : Điểm các môn

# ====================================
# LẤY MSSV VÀ HỌ TÊN
# ====================================

mssv = df.iloc[1, 4:].values

ho_ten = df.iloc[2, 4:].values

# ====================================
# LẤY TÊN MÔN HỌC
# ====================================

mon_hoc = df.iloc[5:, 1].values

# ====================================
# LẤY DỮ LIỆU ĐIỂM
# ====================================

data = df.iloc[5:, 4:].T

# đặt tên cột
data.columns = mon_hoc

# ép kiểu string
data.columns = data.columns.astype(str)

# ====================================
# XỬ LÝ DỮ LIỆU
# ====================================

# chuyển sang số
data = data.apply(pd.to_numeric, errors='coerce')

# thay NaN bằng 0
data = data.fillna(0)

# ====================================
# CHUẨN HÓA DỮ LIỆU
# ====================================

scaler = StandardScaler()

data_scaled = scaler.fit_transform(data)

# ====================================
# KMEANS
# ====================================

kmeans = KMeans(
    n_clusters=3,
    random_state=42,
    n_init=10
)

clusters = kmeans.fit_predict(data_scaled)

# ====================================
# THÊM THÔNG TIN
# ====================================

data["Cum"] = clusters

data["MSSV"] = mssv

data["Họ tên"] = ho_ten

data["STT"] = range(1, len(data) + 1)

# ====================================
# ĐỔI TÊN NHÓM
# ====================================

ten_nhom = {
    0: "Giỏi",
    1: "Khá",
    2: "Yếu"
}

data["Tên nhóm"] = data["Cum"].map(ten_nhom)

# ====================================
# IN KẾT QUẢ
# ====================================

print("\n========== KẾT QUẢ PHÂN CỤM ==========\n")

for i in range(3):

    print(f"\n===== NHÓM {ten_nhom[i]} =====")

    group = data[data["Cum"] == i]

    for index, row in group.iterrows():

        print(
            "STT:", row["STT"],
            "| MSSV:", row["MSSV"],
            "| Họ tên:", row["Họ tên"]
        )

# ====================================
# XUẤT FILE EXCEL
# ====================================

ket_qua = data[[
    "STT",
    "MSSV",
    "Họ tên",
    "Cum",
    "Tên nhóm"
]]

output_excel = "ket_qua_phan_cum.xlsx"

ket_qua.to_excel(
    output_excel,
    index=False
)

print("\nĐã xuất file Excel")

# ====================================
# PCA GIẢM CHIỀU
# ====================================

pca = PCA(n_components=2)

pca_data = pca.fit_transform(data_scaled)

x = pca_data[:, 0]
y = pca_data[:, 1]

# ====================================
# ĐẾM SỐ LƯỢNG NHÓM
# ====================================

so_luong_nhom = data["Tên nhóm"].value_counts()

# ====================================
# BIỂU ĐỒ 1: KMEANS
# ====================================

plt.figure(figsize=(12, 8))

plt.scatter(
    x,
    y,
    c=data["Cum"],
    cmap="viridis",
    s=180,
    edgecolors="black"
)

# hiện STT
for i in range(len(data)):

    plt.annotate(
        str(data["STT"].iloc[i]),
        (x[i], y[i]),
        textcoords="offset points",
        xytext=(5, 5),
        fontsize=8,
        fontweight="bold"
    )

plt.title(
    "PHÂN CỤM KMEANS",
    fontsize=20,
    fontweight="bold"
)

plt.xlabel("PCA 1", fontsize=12)
plt.ylabel("PCA 2", fontsize=12)

plt.grid(True, linestyle="--", alpha=0.5)

plt.tight_layout()

# lưu ảnh
plt.savefig("bieudo_kmeans.png")

plt.close()

print("Đã lưu biểu đồ KMeans")

# ====================================
# BIỂU ĐỒ 2: BIỂU ĐỒ TRÒN
# ====================================

plt.figure(figsize=(8, 8))

plt.pie(
    so_luong_nhom,
    labels=so_luong_nhom.index,
    autopct='%1.1f%%',
    startangle=90
)

plt.title(
    "TỶ LỆ SINH VIÊN THEO NHÓM",
    fontsize=20,
    fontweight="bold"
)

plt.tight_layout()

plt.savefig("bieudo_tron.png")

plt.close()

print("Đã lưu biểu đồ tròn")

# ====================================
# BIỂU ĐỒ 3: BIỂU ĐỒ CỘT
# ====================================

plt.figure(figsize=(10, 7))

plt.bar(
    so_luong_nhom.index,
    so_luong_nhom.values
)

plt.title(
    "SỐ LƯỢNG SINH VIÊN MỖI NHÓM",
    fontsize=20,
    fontweight="bold"
)

plt.xlabel("Nhóm", fontsize=12)
plt.ylabel("Số lượng", fontsize=12)

plt.grid(True, linestyle="--", alpha=0.5)

plt.tight_layout()

plt.savefig("bieudo_cot.png")

plt.close()

print("Đã lưu biểu đồ cột")

# ====================================
# CHÈN ẢNH VÀO EXCEL
# ====================================

wb = load_workbook(output_excel)

ws = wb.active

# ====================================
# ẢNH 1: KMEANS
# ====================================

img1 = Image("bieudo_kmeans.png")

img1.width = 850
img1.height = 550

ws.add_image(img1, "H2")

# ====================================
# ẢNH 2: BIỂU ĐỒ TRÒN
# ====================================

img2 = Image("bieudo_tron.png")

img2.width = 500
img2.height = 400

ws.add_image(img2, "H35")

# ====================================
# ẢNH 3: BIỂU ĐỒ CỘT
# ====================================

img3 = Image("bieudo_cot.png")

img3.width = 600
img3.height = 400

ws.add_image(img3, "H60")

# ====================================
# LƯU FILE
# ====================================

wb.save(output_excel)

print("\nĐã chèn 3 biểu đồ vào Excel")

print("\nHOÀN THÀNH CHƯƠNG TRÌNH")