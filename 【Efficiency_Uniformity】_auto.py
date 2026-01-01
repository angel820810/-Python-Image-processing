from PIL import Image  # PIL 是 package，Image 是 module
import numpy as np
import pandas as pd
import tifffile as tiff
import xlsxwriter
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from matplotlib.colors import Normalize


file_path_ref = None
file_path_mea = None


###### 前置作業 ######




def reference_image_path():
    global file_path_ref

    path = filedialog.askopenfilename(
        title = "Select the 【reference image】",filetypes = [("All files", "*.*")])

    if not path:
        return
    
    file_path_ref = Path(path)
    ref_label_path.config(text = str(file_path_ref))



def measured_image_path():
    global file_path_mea

    path = filedialog.askopenfilename(
        title = "Select the 【measured image】",filetypes = [("All files", "*.*")])

    if not path:
        return
    
    file_path_mea = Path(path)
    mea_label_path.config(text = str(file_path_mea))


def run():
    global sub_dir
    
    if file_path_ref is None or file_path_mea is None:
        messagebox.showwarning("Warning", "Please select both reference and measured images.")
        return

    base_dir = Path("D:/Image analysis/") # 主資料夾路徑
    base_dir.mkdir(parents=True, exist_ok = True) # 若不存在就建立
    time_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") # 取得目前日期時間（避免 Windows 不允許的 :）
    sub_dir = base_dir / time_str # 子資料夾路徑
    sub_dir.mkdir()

    print(f"資料夾已建立：{sub_dir}")
    # .get()必須放在"使用者按 Run" 時才會執行的 function 裡
    vmin_input = vmin_var.get().strip()
    vmax_input = vmax_var.get().strip()

    vmin = float(vmin_input) if vmin_input else None
    vmax = float(vmax_input) if vmax_input else None


    try:
        L_ratio_mean, L_Uniformity_1, L_Uniformity_2, L_ratio_min, L_ratio_max = mono_CCD(file_path_ref, file_path_mea, sub_dir, vmax = vmax, vmin = vmin)

        effi_var.set(f"Efficiency: {L_ratio_mean}")  # ..set() 是用來創造一個可以放進"變數"的文字容器
        unif1_var.set(f"Uniformity I: {L_Uniformity_1}")
        unif2_var.set(f"Uniformity II: {L_Uniformity_2}")

        if not vmin_input:
            vmin_var.set(f"{L_ratio_min}")

        if not vmax_input:
            vmax_var.set(f"{L_ratio_max}")


        messagebox.showinfo("Done", "Successfully done!")
        
    except Exception as e:
        messagebox.showerror("Error", "Something wrong....")
        print(e)  # Debug 用（你之後可以拿掉）


def reset_state():
    global file_path_ref, file_path_mea

    file_path_ref = None
    file_path_mea = None

    ref_label_path.config(text="No file selected")
    mea_label_path.config(text="No file selected")

    effi_var.set("Efficiency: -")
    unif1_var.set("Uniformity I: -")
    unif2_var.set("Uniformity II: -")
    vmin_var.set("")
    vmax_var.set("")


####################################### 黑白 CCD ######################################
def mono_CCD(file_path_ref, file_path_mea, sub_dir, vmin = None, vmax= None):

    img_ref = Image.open(file_path_ref)
    img_mea = Image.open(file_path_mea)
    print(img_mea.mode)

    img_ref = np.array(img_ref).astype(float)
    img_mea = np.array(img_mea).astype(float)

    h = min(img_ref.shape[0], img_mea.shape[0])
    w = min(img_ref.shape[1], img_mea.shape[1])

    L_ref = img_ref[:h, :w]
    L_mea = img_mea[:h, :w]

    ROI = np.zeros((h, w), dtype = bool)
    ROI[:h, :w] = True

    valid_mask = (L_ref > 0) & ROI

    L_ratio = np.full((h, w), np.nan)
    L_ratio[valid_mask] = L_mea[valid_mask] / L_ref[valid_mask]

    L_ratio_mean = round(np.nanmean(L_ratio), 3)
    L_ratio_std = round(np.nanstd(L_ratio), 3)
    L_ratio_max = round(np.nanmax(L_ratio), 3)
    L_ratio_min = round(np.nanmin(L_ratio), 3)
    L_Uniformity_1 = L_ratio_std / L_ratio_mean
    L_Uniformity_2 = (L_ratio_max - L_ratio_mean) / L_ratio_mean
    L_Uniformity_1 = round(L_Uniformity_1, 3)
    L_Uniformity_2 = round(L_Uniformity_2, 3)

    L_ratio_1D = L_ratio[~np.isnan(L_ratio)]
    L_ratio_sorted = np.sort(L_ratio_1D)
    L_P5 = round(np.percentile(L_ratio_sorted, 5), 3)
    L_P95 = round(np.percentile(L_ratio_sorted, 95), 3)

    if vmin is None:
        vmin = L_ratio_min

    if vmax is None:
        vmax = L_ratio_max


    result_path = sub_dir

    # DataFrame 需要的是 list / array，不能直接放value
    Wavelength = ["Mono"]
    Analysis_df = pd.DataFrame({"Color": Wavelength,
                                "Mean (Efficiency)": [L_ratio_mean],
                                "Uniformity_I": [L_Uniformity_1],
                                "Uniformity_II": [L_Uniformity_2],
                                "std": [L_ratio_std],
                                "Max": [L_ratio_max],
                                "Min": [L_ratio_min],
                                "Darkest 5%": [L_P5],
                                "Brightes 5%": [L_P95]})

    Analysis_df.to_excel(result_path / "【Analysis】Efficiency & Uniformity.xlsx", index = False, header = True)


    df_L = pd.DataFrame(L_ratio)
    df_L.to_excel(result_path / "【L】Pixel_wise_division.xlsx", index = False, header = False)
    tiff.imwrite(result_path / "【L】Pixel_wise_division.tif", L_ratio.astype(np.float32))

    fig, ax = plt.subplots(figsize = (6, 5))
   # img_L = ax.imshow(L_ratio, cmap = "jet", vmin = vmin, vmax = vmax)
   # cbar = fig.colorbar(img_L, ax = ax)

    norm = Normalize(vmin=vmin, vmax=vmax)
    img_L = ax.imshow(L_ratio, cmap="jet", norm=norm)
    cbar = fig.colorbar(img_L, ax=ax)

    cbar.set_label("Efficiency (mea / ref)")
    ax.set_title("Monochromatic Efficiency Map")
    ax.axis("off")
    plt.savefig(result_path / "Monochromatic_efficiency_map.jpg", dpi = 300, bbox_inches = "tight")

    plt.close(fig)

    return L_ratio_mean, L_Uniformity_1, L_Uniformity_2, L_ratio_min, L_ratio_max


if __name__ == "__main__":
    root = tk.Tk()
    root.title("【Image Analysis】- Efficiency & Uniformity")
    root.geometry("600x360")
    
    # ===== 宣告變數 ====
    # tk.StringVar 建立一個給 tkinter 用的「可變字串變數」; value = ... 用來設定初始顯示文字
    effi_var = tk.StringVar(value = "Efficiency: -")  
    unif1_var = tk.StringVar(value = "Uniformity I: -")
    unif2_var = tk.StringVar(value = "Uniformity II: -")
    vmin_var = tk.StringVar(value = "")
    vmax_var = tk.StringVar(value = "")


    # 讓 column 0 可以伸縮（視窗拉伸時好看）
    root.columnconfigure(0, weight = 1)
    root.columnconfigure(1, weight = 1)
    root.columnconfigure(2, weight = 1)


    # ===== Reference Image =====
    label_ref_title = tk.Label(root, text = "【Reference Image】", font = ("Times New Roman", 11, "bold underline"))
    label_ref_title.grid(row = 0, column = 0, pady = (10, 5))

    ref_open_btn = tk.Button(root,text="open reference image", font = ("Times New Roman", 11), command = reference_image_path, width = 18)
    ref_open_btn.grid(row = 1, column = 0, pady = 5)

    ref_label_path = tk.Label(root, text="No file selected", fg="gray")
    ref_label_path.grid(row = 2, column = 0, pady = (0, 15))

    # ===== Measured Image =====
    label_mea_title = tk.Label(root, text = "【Measured Image】", font = ("Times New Roman", 11, "bold underline"))
    label_mea_title.grid(row = 0, column = 1, pady = (5, 5))

    mea_open_btn = tk.Button(root, text = "open measured image", font = ("Times New Roman", 11), command = measured_image_path, width = 18)
    mea_open_btn.grid(row = 1, column = 1, pady = 5)

    mea_label_path = tk.Label(root, text="No file selected", fg="gray")
    mea_label_path.grid(row = 2, column = 1, pady = (0, 10))


    # ===== Result outline =====
    result_title = tk.Label(root, text = "【Result】", font = ("Times New Roman", 14, "bold underline"))
    result_title.grid(row = 3, column = 0, columnspan = 2, pady = (15, 15))

    # 建立一個 Rresult analysis 專用的 frame
    result_frame = tk.Frame(root) 
   
    result_frame.grid(row=6, column=0, columnspan=2, pady=(5, 5), sticky="w")

    result_frame.columnconfigure(0, weight = 1)
    
    output_effi = tk.Label(result_frame, textvariable = effi_var, font = ("Times New Roman", 12))
    output_effi.grid(row = 0, column = 0, sticky = "w")

    output_unif_1 = tk.Label(result_frame, textvariable = unif1_var, font = ("Times New Roman", 12))
    output_unif_1.grid(row = 1, column = 0, sticky = "w")

    output_unif_2 = tk.Label(result_frame, textvariable = unif2_var, font = ("Times New Roman", 12))
    output_unif_2.grid(row = 2, column = 0, sticky = "w")


    # ===== Run the program =====
    btn_frame = tk.Frame(root)
    btn_frame.grid(row = 7, column = 0, columnspan = 2, pady = (10, 5))

    run_btn = tk.Button(btn_frame, text = "Run analysis", font = ("Times New Roman", 11), command = run, width = 10)
    run_btn.pack(side = "left", padx = 5)

    reset_btn = tk.Button(btn_frame, text = "Reset", font = ("Times New Roman", 11), command = reset_state, width = 10)
    reset_btn.pack(side = "left", padx = 5)

    # ===== Colormap 的 vmax 和 vmin =====
    cmap_label = tk.Label(root, text = "【Colormap scale bar】", font = ("Times New Roman", 11, "bold underline"))
    cmap_label.grid(row = 0, column = 2, pady = (10, 5), sticky = "w", padx = 10)
    
    cmap_frame = tk.Frame(root)
    cmap_frame.grid(row = 1, column = 2, rowspan = 1, sticky = "nw", padx = 10)
    cmap_frame.columnconfigure(0, weight=0)  # weight = 0 代表不要撐滿右邊空間
    cmap_frame.columnconfigure(1, weight=0)

    tk.Label(cmap_frame, text = "Min.:", font = ("Times New Roman", 11)).grid(row = 0, column = 0, sticky = "w", padx = (0, 5))
    vmin_entry = tk.Entry(cmap_frame, textvariable = vmin_var, width = 8)
    vmin_entry.grid(row = 0, column = 1, sticky = "w")

    tk.Label(cmap_frame, text = "Max.:", font = ("Times New Roman", 11)).grid(row = 1, column = 0, sticky = "w", padx = (0, 5))
    vmax_entry = tk.Entry(cmap_frame, textvariable = vmax_var, width = 8)
    vmax_entry.grid(row = 1, column = 1, sticky = "w")


    root.mainloop()