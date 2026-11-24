# -*- coding: utf-8 -*-
import os
import sys
import traceback
import multiprocessing
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed


def check_and_import():
    missing = []
    try:
        import OCP
    except ImportError:
        missing.append("cadquery-ocp")
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    if missing and not getattr(sys, "frozen", False):
        print("缺少必要的依赖库，正在自动安装:", ", ".join(missing))
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


def calc_step_dimensions(filepath):
    """
    使用 OpenCascade (OCP) 读取 STEP 文件并获取轴对齐包围盒（长 X、宽 Y、高 Z）及体积。
    返回: (料号, 完整文件名, 长_mm, 宽_mm, 高_mm, 体积_mm3, 实体数, 状态, 完整路径)
    """
    filename = os.path.basename(filepath)
    part_no = os.path.splitext(filename)[0]
    try:
        from OCP.STEPControl import STEPControl_Reader
        from OCP.IFSelect import IFSelect_RetDone
        from OCP.Bnd import Bnd_Box
        from OCP.BRepBndLib import BRepBndLib
        from OCP.GProp import GProp_GProps
        from OCP.BRepGProp import BRepGProp
        from OCP.TopExp import TopExp_Explorer
        from OCP.TopAbs import TopAbs_SOLID

        reader = STEPControl_Reader()
        status = reader.ReadFile(filepath)
        if status != IFSelect_RetDone:
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, "读取失败(STEP格式异常)", filepath)

        reader.TransferRoots()
        shape = reader.OneShape()
        if shape.IsNull():
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, "空模型(无几何体)", filepath)

        # 统计实体(SOLID)数量
        explorer = TopExp_Explorer(shape, TopAbs_SOLID)
        solid_count = 0
        while explorer.More():
            solid_count += 1
            explorer.Next()

        # 计算精确轴对齐包围盒 (X=长, Y=宽, Z=高)
        bbox = Bnd_Box()
        # useTriangulation=False 使用精确几何边界求包围盒
        BRepBndLib.Add_s(shape, bbox, False)

        if bbox.IsVoid():
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, solid_count, "警告: 包围盒为空", filepath)

        xmin, ymin, zmin, xmax, ymax, zmax = bbox.Get()
        length_x = max(0.0, xmax - xmin)
        width_y = max(0.0, ymax - ymin)
        height_z = max(0.0, zmax - zmin)

        # 计算体积 (mm³)
        props = GProp_GProps()
        BRepGProp.VolumeProperties_s(shape, props)
        vol_mm3 = abs(props.Mass())

        return (
            part_no,
            filename,
            round(length_x, 4),
            round(width_y, 4),
            round(height_z, 4),
            round(vol_mm3, 4),
            solid_count,
            "成功",
            filepath,
        )

    except Exception as e:
        return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, f"异常: {str(e)[:80]}", filepath)


def scan_stp_files(root_dir):
    """递归遍历当前文件夹及其所有子文件夹下的 .stp / .step 文件"""
    stp_files = []
    skip_dirs = {
        "$RECYCLE.BIN",
        "System Volume Information",
        "Windows",
        "Program Files",
        "Program Files (x86)",
        "ProgramData",
        ".git",
        "__pycache__",
        "build",
        "dist",
    }

    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if d not in skip_dirs and not d.startswith("$")]
        for fname in filenames:
            if fname.lower().endswith((".stp", ".step")):
                stp_files.append(os.path.join(dirpath, fname))
    return stp_files


def save_to_excel(records, output_path):
    """将最多 10000 条记录写入格式化良好的 Excel 文件"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "STP长宽高明细"

    headers = [
        "序号",
        "料号(文件名)",
        "完整文件名",
        "长 X (mm)",
        "宽 Y (mm)",
        "高 Z (mm)",
        "体积 (mm³)",
        "实体数量",
        "状态",
        "完整路径",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[1].height = 26

    data_font = Font(name="微软雅黑", size=10)
    warn_font = Font(name="微软雅黑", size=10, color="C00000")

    for idx, rec in enumerate(records, start=1):
        # rec: (part_no, filename, length_x, width_y, height_z, vol_mm3, solid_count, status, filepath)
        row_data = [idx, rec[0], rec[1], rec[2], rec[3], rec[4], rec[5], rec[6], rec[7], rec[8]]
        ws.append(row_data)
        row_num = idx + 1
        is_ok = rec[7] == "成功"

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_idx)
            c.font = data_font if is_ok else warn_font
            c.border = thin_border
            if col_idx in (1, 8, 9):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (4, 5, 6, 7):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.0000"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [8, 28, 32, 15, 15, 15, 18, 12, 22, 65]
    for i, w in enumerate(col_widths, start=1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{len(records) + 1}"
    wb.save(output_path)


def main():
    check_and_import()

    if getattr(sys, "frozen", False):
        current_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    # 搜索当前文件夹及其所有子文件夹；如果直接放在盘符根目录（如 D:\），则自动扫描整个磁盘
    scan_root = current_dir

    print("=" * 60)
    print(f"当前运行目录: {current_dir}")
    print(f"扫描范围: 该目录及所有子文件夹")
    print("=" * 60)
    print("正在检索 .stp / .step 文件，请稍候...")

    stp_files = scan_stp_files(scan_root)
    total_files = len(stp_files)
    print(f"检索完成！共发现 {total_files} 个 STEP 文件。")

    if total_files == 0:
        print("当前目录及子文件夹中未找到任何 .stp 或 .step 文件。")
        return

    max_workers = max(1, (os.cpu_count() or 4) - 1)
    print(f"启动 {max_workers} 个并行工作进程开始计算长宽高...")

    results = [None] * total_files
    completed = 0

    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {executor.submit(calc_step_dimensions, fp): i for i, fp in enumerate(stp_files)}
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            results[idx] = future.result()
            completed += 1
            if completed % 100 == 0 or completed == total_files:
                print(f"进度: {completed} / {total_files} ({completed * 100.0 / total_files:.1f}%)")

    results.sort(key=lambda x: x[0])

    batch_size = 10000
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = os.path.join(current_dir, f"STP长宽高统计结果_{timestamp}")
    os.makedirs(output_folder, exist_ok=True)

    num_batches = (total_files + batch_size - 1) // batch_size
    for b in range(num_batches):
        start_idx = b * batch_size
        end_idx = min((b + 1) * batch_size, total_files)
        batch_records = results[start_idx:end_idx]
        excel_name = f"STP长宽高明细_第{b + 1:02d}卷_({start_idx + 1}-{end_idx}).xlsx"
        excel_path = os.path.join(output_folder, excel_name)
        print(f"正在生成: {excel_name} ...")
        save_to_excel(batch_records, excel_path)

    print("=" * 60)
    print(f"全部处理完成！共生成 {num_batches} 个 Excel 文件。")
    print(f"输出目录: {output_folder}")
    print("=" * 60)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        main()
    except Exception:
        traceback.print_exc()
    input("\n按回车键退出程序...")