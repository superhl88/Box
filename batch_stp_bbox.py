# -*- coding: utf-8 -*-
import os
import sys
import traceback
import multiprocessing
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed

BATCH_SIZE = 10000  # 每个 Excel 表格记录的料号上限


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


def extract_bbox_xyz(shape):
    """提取轴对齐包围盒的长(X)、宽(Y)、高(Z)，兼容 CornerMin/CornerMax 与 Get()"""
    from OCP.Bnd import Bnd_Box
    from OCP.BRepBndLib import BRepBndLib
    from OCP.TopExp import TopExp_Explorer
    from OCP.TopAbs import TopAbs_SHELL, TopAbs_FACE

    bbox = Bnd_Box()
    BRepBndLib.Add_s(shape, bbox, False)
    if bbox.IsVoid():
        BRepBndLib.Add_s(shape, bbox, True)

    if bbox.IsVoid():
        for sub_type in (TopAbs_SHELL, TopAbs_FACE):
            exp = TopExp_Explorer(shape, sub_type)
            while exp.More():
                BRepBndLib.Add_s(exp.Current(), bbox, True)
                exp.Next()
            if not bbox.IsVoid():
                break

    if bbox.IsVoid():
        return None

    try:
        p_min = bbox.CornerMin()
        p_max = bbox.CornerMax()
        lx = max(0.0, float(p_max.X()) - float(p_min.X()))
        ly = max(0.0, float(p_max.Y()) - float(p_min.Y()))
        lz = max(0.0, float(p_max.Z()) - float(p_min.Z()))
        return lx, ly, lz
    except Exception:
        xmin, ymin, zmin, xmax, ymax, zmax = bbox.Get()
        return (
            max(0.0, float(xmax) - float(xmin)),
            max(0.0, float(ymax) - float(ymin)),
            max(0.0, float(zmax) - float(zmin)),
        )


def calc_step_dimensions(filepath):
    """
    读取 STEP 文件并获取长(X)、宽(Y)、高(Z)、体积及实体数量。
    返回: (料号, 完整文件名, 长_mm, 宽_mm, 高_mm, 体积_mm3, 实体数, 状态, 完整路径, 详细错误)
    """
    abs_path = os.path.abspath(filepath)
    filename = os.path.basename(abs_path)
    part_no = os.path.splitext(filename)[0]

    try:
        from OCP.STEPControl import STEPControl_Reader
        from OCP.IFSelect import IFSelect_RetDone
        from OCP.TopExp import TopExp_Explorer
        from OCP.TopAbs import TopAbs_SOLID

        reader = STEPControl_Reader()
        status = reader.ReadFile(abs_path)
        if status != IFSelect_RetDone:
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, "读取失败", abs_path, f"ReadFile status={status}")

        reader.TransferRoots()
        shape = reader.OneShape()
        if shape.IsNull():
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, "空模型", abs_path, "OneShape is Null")

        solid_count = 0
        explorer = TopExp_Explorer(shape, TopAbs_SOLID)
        while explorer.More():
            solid_count += 1
            explorer.Next()

        xyz = extract_bbox_xyz(shape)
        if xyz is None:
            return (part_no, filename, 0.0, 0.0, 0.0, 0.0, solid_count, "无有效边界", abs_path, "Bnd_Box is Void")
        length_x, width_y, height_z = xyz

        vol_mm3 = 0.0
        vol_note = ""
        try:
            from OCP.GProp import GProp_GProps
            from OCP.BRepGProp import BRepGProp

            props = GProp_GProps()
            BRepGProp.VolumeProperties_s(shape, props)
            vol_mm3 = abs(float(props.Mass()))
        except Exception as ve:
            vol_note = f" (体积计算失败: {str(ve).splitlines()[0]})"

        status_msg = ("成功" if solid_count > 0 else "成功(仅曲面/壳体)") + vol_note
        return (
            part_no,
            filename,
            round(length_x, 4),
            round(width_y, 4),
            round(height_z, 4),
            round(vol_mm3, 4),
            solid_count,
            status_msg,
            abs_path,
            "",
        )

    except Exception as e:
        tb_text = traceback.format_exc()
        short_err = str(e).splitlines()[0] if str(e) else repr(e)
        return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, f"异常: {short_err[:120]}", abs_path, tb_text)


def make_crash_record(filepath, reason="子进程异常终止/底层崩溃"):
    abs_path = os.path.abspath(filepath)
    filename = os.path.basename(abs_path)
    part_no = os.path.splitext(filename)[0]
    return (part_no, filename, 0.0, 0.0, 0.0, 0.0, 0, f"异常: {reason}", abs_path, reason)


def run_single_safe(filepath):
    """在独立单进程池中运行单个文件，防止该文件把整批任务拖垮"""
    try:
        with ProcessPoolExecutor(max_workers=1) as single_exec:
            fut = single_exec.submit(calc_step_dimensions, filepath)
            return fut.result(timeout=60)
    except Exception as e:
        return make_crash_record(filepath, f"单文件兜底异常: {str(e)[:100]}")


def process_files_resilient(stp_files, max_workers):
    """分块并行处理；若某块触发 BrokenProcessPool，则逐个安全处理"""
    total = len(stp_files)
    results = [None] * total
    completed = 0
    chunk_size = 50

    for start in range(0, total, chunk_size):
        end = min(start + chunk_size, total)
        chunk_indices = list(range(start, end))
        chunk_done = False

        try:
            with ProcessPoolExecutor(max_workers=max_workers) as executor:
                future_to_idx = {
                    executor.submit(calc_step_dimensions, stp_files[i]): i
                    for i in chunk_indices
                }
                for future in as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        results[idx] = future.result()
                    except Exception as fe:
                        results[idx] = make_crash_record(stp_files[idx], f"Future异常: {str(fe)[:100]}")
                    completed += 1
                    if completed % 50 == 0 or completed == total:
                        print(f"进度: {completed} / {total} ({completed * 100.0 / total:.1f}%)")
            chunk_done = True
        except Exception as pool_err:
            print(f"\n[提示] 批次 ({start + 1}-{end}) 检测到进程池异常 ({type(pool_err).__name__})，正在切换逐个安全模式...")

        if not chunk_done:
            for idx in chunk_indices:
                if results[idx] is None:
                    results[idx] = run_single_safe(stp_files[idx])
                    completed += 1
                    if completed % 20 == 0 or completed == total:
                        print(f"进度(安全模式): {completed} / {total} ({completed * 100.0 / total:.1f}%)")

    return results


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
    """将最多 BATCH_SIZE (10000) 条记录写入格式化良好的 Excel 文件"""
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
        row_data = [idx, rec[0], rec[1], rec[2], rec[3], rec[4], rec[5], rec[6], rec[7], rec[8]]
        ws.append(row_data)
        row_num = idx + 1
        is_ok = str(rec[7]).startswith("成功")

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

    col_widths = [8, 28, 32, 15, 15, 15, 18, 12, 26, 65]
    for i, w in enumerate(col_widths, start=1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{len(records) + 1}"
    wb.save(output_path)


def export_excel_batches(results, output_folder, batch_size=BATCH_SIZE):
    """按每 batch_size (10000) 个料号切分并导出独立的 Excel 文件"""
    total_files = len(results)
    num_batches = (total_files + batch_size - 1) // batch_size
    generated_files = []

    for b in range(num_batches):
        start_idx = b * batch_size
        end_idx = min((b + 1) * batch_size, total_files)
        batch_records = results[start_idx:end_idx]
        excel_name = f"STP长宽高明细_第{b + 1:02d}卷_({start_idx + 1}-{end_idx}).xlsx"
        excel_path = os.path.join(output_folder, excel_name)
        print(f"正在生成 [{b + 1}/{num_batches}]: {excel_name} (共 {len(batch_records)} 个料号)...")
        save_to_excel(batch_records, excel_path)
        generated_files.append(excel_path)

    return generated_files


def main():
    check_and_import()

    if getattr(sys, "frozen", False):
        current_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    scan_root = current_dir

    print("=" * 60)
    print(f"当前运行目录: {current_dir}")
    print(f"扫描范围: 该目录及所有子文件夹")
    print(f"分卷规则: 每 {BATCH_SIZE} 个料号生成 1 个 Excel 文件")
    print("=" * 60)
    print("正在检索 .stp / .step 文件，请稍候...")

    stp_files = scan_stp_files(scan_root)
    total_files = len(stp_files)
    print(f"检索完成！共发现 {total_files} 个 STEP 文件。")

    if total_files == 0:
        print("当前目录及子文件夹中未找到任何 .stp 或 .step 文件。")
        return

    cpu_cnt = os.cpu_count() or 2
    max_workers = max(1, min(4, cpu_cnt - 1)) if getattr(sys, "frozen", False) else max(1, min(6, cpu_cnt - 1))
    print(f"启动 {max_workers} 个并行工作进程开始计算长宽高...")

    results = process_files_resilient(stp_files, max_workers)
    results.sort(key=lambda x: x[0])

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = os.path.join(current_dir, f"STP长宽高统计结果_{timestamp}")
    os.makedirs(output_folder, exist_ok=True)

    err_items = [r for r in results if r[9]]
    if err_items:
        err_log_path = os.path.join(output_folder, "错误详情日志.txt")
        with open(err_log_path, "w", encoding="utf-8") as lf:
            for item in err_items[:50]:
                lf.write(f"文件: {item[8]}\n状态: {item[7]}\n{item[9]}\n" + ("-" * 50) + "\n")
        print(f"检测到 {len(err_items)} 条异常记录，首条异常摘要: {err_items[0][7]}")
        print(f"异常明细已写入: {err_log_path}")

    generated_files = export_excel_batches(results, output_folder, batch_size=BATCH_SIZE)

    print("=" * 60)
    print(f"全部处理完成！共生成 {len(generated_files)} 个 Excel 文件（每表最多 {BATCH_SIZE} 个料号）。")
    print(f"输出目录: {output_folder}")
    print("=" * 60)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        main()
    except Exception:
        traceback.print_exc()
    if sys.stdin and sys.stdin.isatty():
        input("\n按回车键退出程序...")