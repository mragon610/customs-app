import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.formula.translate import Translator
from copy import copy
from PIL import Image
import io

st.set_page_config(page_title="Hệ thống Kê khai Hải quan INV-PKL", layout="wide")

# --- HỆ THỐNG PHÂN TÍCH & GỢI Ý THÔNG MINH DỰA TRÊN TÊN TIẾNG VIỆT ---
def smart_suggest_customs(vietnamese_name):
    name_lc = str(vietnamese_name).strip().lower()
    
    if "trang trí" in name_lc or "đồ chơi" in name_lc or "spinner" in name_lc or "cốt cắm hoa" in name_lc or "xốp" in name_lc:
        if "nhựa" in name_lc or "cốt" in name_lc: return ("PLASTIC", "FACTORY", "PLASTIC DECORATION PARTS")
        if "xốp" in name_lc: return ("FOAM", "FACTORY", "FLORAL FOAM")
        return ("PLASTIC", "FACTORY", "PLASTIC TOY / DECORATION")
    if "dép" in name_lc: return ("PVC", "FACTORY", "PVC SANDAL")
    if "giày" in name_lc: return ("SYNTHETIC LEATHER", "FACTORY", "CASUAL SHOES")
    if "bốt" in name_lc: return ("SYNTHETIC LEATHER", "FACTORY", "BOOTS")
    if "khoác da" in name_lc: return ("ARTIFICIAL LEATHER", "FACTORY", "ARTIFICIAL LEATHER JACKET")
    if "khoác dạ" in name_lc or "dạ" in name_lc: return ("COTTON", "KNIT", "MEN JACKET")
    if "khoác nỉ" in name_lc or "hoodie" in name_lc: return ("COTTON", "KNIT", "LONG SLEEVES HOODIE")
    if "khoác lông" in name_lc: return ("POLYESTER", "KNIT", "FUR JACKET")
    if "khoác" in name_lc: return ("POLYESTER", "KNIT", "PUFFER JACKET")
    if "nỉ" in name_lc or "sweatshirt" in name_lc: return ("COTTON", "KNIT", "SWEATSHIRT")
    if "len" in name_lc: return ("COTTON", "KNIT", "LONG SLEEVES PULLOVER")
    if "thun" in name_lc or "cộc tay" in name_lc: return ("COTTON", "KNIT", "T-SHIRT")
    if "dài tay" in name_lc or "sơ mi" in name_lc: return ("COTTON", "KNIT", "LONG SLEEVES SHIRT")
    if "3 lỗ" in name_lc or "hai dây" in name_lc: return ("COTTON", "KNIT", "TANK TOP")
    if "jean ngắn" in name_lc or "short jean" in name_lc: return ("DENIM", "WOVEN", "DENIM SHORTS")
    if "jean" in name_lc: return ("DENIM", "WOVEN", "DENIM PANTS")
    if "quần short" in name_lc or "quần ngắn" in name_lc: return ("COTTON", "KNIT", "SHORTS")
    if "quần" in name_lc: return ("COTTON", "KNIT", "CASUAL PANTS")
    if "chân váy" in name_lc: return ("COTTON", "WOVEN", "SKIRT")
    if "váy" in name_lc or "đầm" in name_lc: return ("COTTON", "KNIT", "ONE PIECE DRESS")
    if "tóc giả" in name_lc or "lưới" in name_lc: return ("SYNTHETIC FIBER", "FACTORY", "SYNTHETIC HAIR WIG")
    if "lược" in name_lc: return ("PLASTIC", "FACTORY", "PLASTIC COMB")
    if "gấu bông" in name_lc or "móc khoá" in name_lc: return ("PLUSH FABRIC", "FACTORY", "STUFFED TOY KEYCHAIN")
    if "hoa" in name_lc or "giấy gói" in name_lc or "ruy băng" in name_lc or "dây" in name_lc: return ("PAPER", "FACTORY", "FLOWER PACKAGING MATERIAL")
    if "ga giường" in name_lc or "khăn" in name_lc: return ("COTTON", "WOVEN", "BED LINEN / TOWEL")

    return ("COTTON", "KNIT", "LONG SLEEVES")

# Khởi tạo bộ nhớ danh mục
if "custom_materials" not in st.session_state:
    st.session_state.custom_materials = ["COTTON", "PVC", "DENIM", "PLUSH FABRIC", "ARTIFICIAL LEATHER", "POLYESTER", "NYLON", "SILK", "PLASTIC", "FOAM", "PAPER", "SYNTHETIC FIBER"]
if "custom_weaves" not in st.session_state:
    st.session_state.custom_weaves = ["KNIT", "WOVEN", "FACTORY"]
if "all_declared_data" not in st.session_state:
    st.session_state.all_declared_data = {}
if "current_sheet_idx" not in st.session_state:
    st.session_state.current_sheet_idx = 0

# ================= HÀM XUẤT FILE EXCEL ĐỘNG & BẢO TOÀN 100% CÔNG THỨC/ĐỊNH DẠNG =================
def generate_excel_export(uploaded_file_bytes, case_order_list, all_items_by_case, sheet_to_makien, makien_to_sheet, size_map):
    # Dùng data_only=False để giữ nguyên toàn bộ CÔNG THỨC trong file gốc
    output_wb = openpyxl.load_workbook(io.BytesIO(uploaded_file_bytes), data_only=False)
    
    new_sheet_name = "INV-PKL-Ket-Qua"
    if new_sheet_name in output_wb.sheetnames:
        del output_wb[new_sheet_name]
        
    if "INV-PKL" not in output_wb.sheetnames:
        st.error("Không tìm thấy sheet INV-PKL gốc!")
        return None
        
    ws_out = output_wb.copy_worksheet(output_wb["INV-PKL"])
    ws_out.title = new_sheet_name
    
    # 0. CHIẾN THUẬT: PHÁ BỎ TOÀN BỘ MERGE TỪ DÒNG 13 TRỞ ĐI
    # Tránh triệt để lỗi Corrupted File của Excel khi chèn dòng qua vùng bị Merge
    ranges_to_unmerge = [mcr for mcr in ws_out.merged_cells.ranges if mcr.min_row >= 13]
    for mcr in ranges_to_unmerge:
        ws_out.unmerge_cells(range_string=str(mcr))
        
    # Hàm dò tìm tự động vị trí các Khối Kiện trong bảng mẫu
    def get_blocks(ws):
        blocks = []
        start = None
        for r in range(13, ws.max_row + 1):
            if ws.cell(r, 2).value == 'CASE No :':
                start = r
            elif ws.cell(r, 2).value == 'TOTAL' and start is not None:
                blocks.append({'start': start, 'end': r})
                start = None
        return blocks
        
    # Gom toàn bộ dữ liệu cần đổ
    cases_data = []
    for ma_kien in case_order_list:
        sheet_name = makien_to_sheet.get(ma_kien)
        if not sheet_name: continue
        
        if sheet_name in st.session_state.all_declared_data:
            c_data = st.session_state.all_declared_data[sheet_name]
            cases_data.append({'ma_kien': c_data['ma_kien'], 'ctns_size': c_data['ctns_size'], 'items': c_data['items']})
        else:
            raw_items = all_items_by_case[ma_kien]
            ctns_size = size_map.get(sheet_name, "50*40*25")
            items = []
            for i, it in enumerate(raw_items):
                items.append({
                    "No": i + 1, "STYLE_HDR": "NEW", "DESCRIPTION": "",
                    "ORIGIN": it["origin"], "STYLE_ITEM": "PLAIN", "QTY": it["qty"], "PRICE": 0.55
                })
            cases_data.append({'ma_kien': ma_kien, 'ctns_size': ctns_size, 'items': items})
            
    blocks = get_blocks(ws_out)
    if not blocks: return None
    
    # 1. THÊM HOẶC XÓA KHỐI KIỆN ĐỂ SỐ KIỆN BẰNG ĐÚNG THỰC TẾ
    while len(blocks) > len(cases_data):
        del_start = blocks[-1]['start']
        del_end = blocks[-1]['end']
        if ws_out.cell(del_end + 1, 2).value is None and ws_out.cell(del_end + 1, 1).value is None:
            del_end += 1 
        ws_out.delete_rows(del_start, del_end - del_start + 1)
        blocks = get_blocks(ws_out)
        
    while len(blocks) < len(cases_data):
        src_start = blocks[-1]['start']
        src_end = blocks[-1]['end']
        has_separator = False
        if ws_out.cell(src_end + 1, 2).value is None and ws_out.cell(src_end + 1, 1).value is None:
            src_end += 1; has_separator = True
            
        num_rows = src_end - src_start + 1
        insert_at = src_end + 1
        ws_out.insert_rows(insert_at, num_rows)
        
        for offset in range(num_rows):
            for c in range(1, 11):
                src_cell = ws_out.cell(src_start + offset, c)
                tgt_cell = ws_out.cell(insert_at + offset, c)
                
                # Copy và dịch chuyển Công Thức (Ctrl + D)
                if isinstance(src_cell.value, str) and src_cell.value.startswith('='):
                    tgt_cell.value = Translator(src_cell.value, origin=src_cell.coordinate).translate_formula(tgt_cell.coordinate)
                else:
                    tgt_cell.value = src_cell.value
                    
                # Copy format
                if src_cell.has_style:
                    tgt_cell.font = copy(src_cell.font)
                    tgt_cell.border = copy(src_cell.border)
                    tgt_cell.fill = copy(src_cell.fill)
                    tgt_cell.number_format = src_cell.number_format
                    tgt_cell.alignment = copy(src_cell.alignment)
            ws_out.row_dimensions[insert_at + offset].height = ws_out.row_dimensions[src_start + offset].height
            
        blocks = get_blocks(ws_out)
        
    # 2. ĐIỀU CHỈNH SỐ DÒNG BÊN TRONG MỖI KIỆN (Thêm/Bớt items bằng logic Ctrl+D)
    for i in range(len(cases_data) - 1, -1, -1):
        block = blocks[i]
        case = cases_data[i]
        target = len(case['items'])
        
        item_start = block['start'] + 2
        item_end = block['end'] - 1
        current = item_end - item_start + 1
        
        if target > current:
            # Chèn thêm dòng thiếu và Ctrl+D từ dòng cuối cùng của kiện
            diff = target - current
            insert_at = item_end + 1
            ws_out.insert_rows(insert_at, diff)
            for offset in range(diff):
                for c in range(1, 11):
                    src_cell = ws_out.cell(item_end, c) 
                    tgt_cell = ws_out.cell(insert_at + offset, c)
                    
                    if isinstance(src_cell.value, str) and src_cell.value.startswith('='):
                        tgt_cell.value = Translator(src_cell.value, origin=src_cell.coordinate).translate_formula(tgt_cell.coordinate)
                    else:
                        tgt_cell.value = src_cell.value
                        
                    if src_cell.has_style:
                        tgt_cell.font = copy(src_cell.font)
                        tgt_cell.border = copy(src_cell.border)
                        tgt_cell.fill = copy(src_cell.fill)
                        tgt_cell.number_format = src_cell.number_format
                        tgt_cell.alignment = copy(src_cell.alignment)
        elif target < current:
            diff = current - target
            ws_out.delete_rows(item_end - diff + 1, diff)
            
    # 3. ĐIỀN DỮ LIỆU CHÍNH THỨC & RE-MERGE LẠI CHO ĐẸP
    blocks = get_blocks(ws_out)
    for i, case in enumerate(cases_data):
        block = blocks[i]
        
        ws_out.cell(block['start'], 4, value=case['ma_kien'])
        ws_out.cell(block['start'], 8, value=case['ctns_size'])
        ws_out.cell(block['start'] + 1, 1, value=i + 1)
        
        items = case['items']
        r = block['start'] + 2
        for it in items:
            ws_out.cell(r, 2, value=it['No'])
            ws_out.cell(r, 3, value=it.get('STYLE_HDR', 'NEW'))
            ws_out.cell(r, 4, value=it.get('DESCRIPTION', ''))
            ws_out.cell(r, 5, value=it.get('ORIGIN', 'CN'))
            ws_out.cell(r, 6, value=it.get('STYLE_ITEM', 'PLAIN'))
            ws_out.cell(r, 7, value=int(it.get('QTY', 1)))
            ws_out.cell(r, 8, value=float(it.get('PRICE', 0.55)))
            
            # Đảm bảo công thức tính AMOUNT luôn chuẩn tuyệt đối
            ws_out.cell(r, 9, value=f"=G{r}*H{r}") 
            ws_out.cell(r, 10, value="USD")
            r += 1
            
        first_row = block['start'] + 2
        last_row = block['end'] - 1
        total_row = block['end']
        
        # Khôi phục công thức tổng nội bộ kiện
        ws_out.cell(total_row, 7, value=f"=SUM(G{first_row}:G{last_row})")
        ws_out.cell(total_row, 9, value=f"=SUM(I{first_row}:I{last_row})")
        
        # CHỐT HẠ: MERGE LẠI Ô TỪ B->F CỦA DÒNG TOTAL ĐÚNG NHƯ BẢN GỐC
        ws_out.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=6)
        ws_out.cell(total_row, 2).alignment = Alignment(horizontal='center', vertical='center')
        
    # 4. SỬA CHỮA HOÀN CHỈNH FOOTER BỊ SAI CÔNG THỨC (Cập nhật lại dải SUMIF tự động)
    data_start = blocks[0]['start']
    data_end = blocks[-1]['end']
    footer_start = data_end + 1
    
    for r in range(footer_start, ws_out.max_row + 1):
        for c in range(1, 11):
            val = ws_out.cell(r, c).value
            if isinstance(val, str) and "SUMIF" in val.upper():
                label = str(ws_out.cell(r, 6).value).strip()
                if c == 7: # Cột QTY tổng
                    if label == "TOTAL":
                        ws_out.cell(r, c, value=f'=SUMIF(B{data_start}:B{data_end},"TOTAL",G{data_start}:G{data_end})')
                    elif label in ["VN", "CN", "KR"]:
                        ws_out.cell(r, c, value=f'=SUMIF(E{data_start}:E{data_end},"{label}",G{data_start}:G{data_end})')
                elif c == 9: # Cột AMOUNT tổng
                    if label == "TOTAL":
                        ws_out.cell(r, c, value=f'=SUMIF(B{data_start}:B{data_end},"TOTAL",I{data_start}:I{data_end})')
                    elif label in ["VN", "CN", "KR"]:
                        ws_out.cell(r, c, value=f'=SUMIF(E{data_start}:E{data_end},"{label}",I{data_start}:I{data_end})')
                        
    buf = io.BytesIO()
    output_wb.save(buf)
    buf.seek(0)
    return buf

# ================= UI CHÍNH =================
st.title("📦 Hệ thống Tự động hóa Kê khai Hải quan INV-PKL")

uploaded_file = st.sidebar.file_uploader("Tải lên File Excel Template 4505 mới", type=["xlsx"])

if uploaded_file:
    # Dùng data_only=True ở đây CHỈ DÀNH RIÊNG cho việc ĐỌC & TÁCH DỮ LIỆU HIỂN THỊ TRÊN WEB (không lưu)
    uploaded_file_bytes = uploaded_file.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file_bytes), data_only=True)
    
    # 1. ĐỌC TAB TỔNG HỢP
    all_items_by_case = {}
    case_order_list = []
    try:
        sum_sheet_name = [s for s in wb.sheetnames if s.lower() in ["tổng hợp", "tong hop"]][0]
        summary_df = pd.read_excel(io.BytesIO(uploaded_file_bytes), sheet_name=sum_sheet_name, header=1)
        summary_df = summary_df.iloc[:, :6] 
        summary_df.columns = ["STT", "Ma_Kien", "Mat_Hang", "So_Luong", "Xuat_Xu", "Can_Nang"]
        summary_df["Ma_Kien"] = summary_df["Ma_Kien"].ffill()
        summary_df = summary_df.dropna(subset=["Ma_Kien", "Mat_Hang"])
        
        for index, row in summary_df.iterrows():
            ma_kien = str(row["Ma_Kien"]).strip()
            if pd.isna(ma_kien) or ma_kien == "nan": continue
            if ma_kien not in all_items_by_case:
                all_items_by_case[ma_kien] = []
                case_order_list.append(ma_kien)
            mat_hang = str(row["Mat_Hang"]).strip() if pd.notna(row["Mat_Hang"]) else ""
            if not mat_hang: continue
            qty = int(row["So_Luong"]) if pd.notna(row["So_Luong"]) and str(row["So_Luong"]).isdigit() else 1
            origin = str(row["Xuat_Xu"]).strip() if pd.notna(row["Xuat_Xu"]) else "CN"
            all_items_by_case[ma_kien].append({"desc_vn": mat_hang, "qty": qty, "origin": origin})
    except Exception as e:
        st.error(f"⚠️ Lỗi khi đọc tab Tổng hợp: {e}")

    case_sheets = [s for s in wb.sheetnames if s.lower() not in ["tổng hợp", "tong hop", "inv-pkl", "_log"]]
    
    # 2. MAP MÃ KIỆN VÀ TÌM SIZE CHO TẤT CẢ CÁC KIỆN
    sheet_to_makien = {}
    makien_to_sheet = {}
    size_map = {}
    for s_name in case_sheets:
        ws_tmp = wb[s_name]
        m_kien = str(ws_tmp.cell(row=3, column=4).value).strip()
        c_size = str(ws_tmp.cell(row=3, column=8).value).strip()
        if c_size == "None" or not c_size: c_size = "50*40*25"
        sheet_to_makien[s_name] = m_kien
        makien_to_sheet[m_kien] = s_name
        size_map[s_name] = c_size

    # --- SIDEBAR XUẤT FILE RAW/HOÀN CHỈNH BẤT CỨ LÚC NÀO ---
    st.sidebar.markdown("---")
    st.sidebar.success(f"Đã nạp {len(case_sheets)} kiện hàng.")
    
    export_buf = generate_excel_export(uploaded_file_bytes, case_order_list, all_items_by_case, sheet_to_makien, makien_to_sheet, size_map)
    if export_buf:
        st.sidebar.download_button(
            label="📥 TẢI FILE KẾT QUẢ NGAY", 
            data=export_buf, 
            file_name="INV-PKL_KetQua_ThongMinh.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.sidebar.markdown("*(Bản có sẵn dữ liệu định lượng, ô Cột Tiếng Anh sẽ để trống cho các kiện bạn chưa Bấm Lưu, bạn tải xuống xử lý tay vẫn dùng tốt!)*")
    st.sidebar.markdown("---")

    # 3. CHỌN SHEET ĐỂ LÀM VIỆC TRÊN UI
    if st.session_state.current_sheet_idx >= len(case_sheets): st.session_state.current_sheet_idx = len(case_sheets) - 1
    selected_sheet = st.sidebar.selectbox("Chọn kiện hàng để khai báo:", case_sheets, index=st.session_state.current_sheet_idx)
    if case_sheets.index(selected_sheet) != st.session_state.current_sheet_idx:
        st.session_state.current_sheet_idx = case_sheets.index(selected_sheet)
        st.rerun()

    ws_case = wb[selected_sheet]
    current_ma_kien = sheet_to_makien.get(selected_sheet, selected_sheet)
    current_ctns_size = size_map.get(selected_sheet, "50*40*25")
        
    items_in_case = all_items_by_case.get(current_ma_kien, [])
    if not items_in_case: items_in_case = [{"desc_vn": "Mẫu", "qty": 1, "origin": "CN"}]

    images_by_item = {}
    if hasattr(ws_case, '_images') and ws_case._images:
        for img in ws_case._images:
            try:
                row_idx = img.anchor._from.row + 1 
                item_index = (row_idx - 4) // 14
                if item_index >= 0:
                    if item_index not in images_by_item: images_by_item[item_index] = []
                    images_by_item[item_index].append(img._data())
            except Exception: pass

    st.header(f"📋 Đang xử lý Kiện: [{selected_sheet}] - Mã kiện: {current_ma_kien}")

    case_key = f"case_data_{selected_sheet}"
    if case_key not in st.session_state:
        if selected_sheet in st.session_state.all_declared_data:
            st.session_state[case_key] = st.session_state.all_declared_data[selected_sheet]["items"]
        else:
            init_list = []
            for i, it in enumerate(items_in_case):
                mat, meth, prod = smart_suggest_customs(it["desc_vn"])
                init_list.append({
                    "No": i + 1, "STYLE_HDR": "NEW", "MAT": mat, "METH": meth, "PROD": prod,
                    "ORIGIN": it["origin"], "STYLE_ITEM": "PLAIN",
                    "QTY": it["qty"], "PRICE": 0.55, "img_idx": i, "row_start": 4 + (i * 14), "vn_desc": it["desc_vn"]
                })
            st.session_state[case_key] = init_list

    current_items = st.session_state[case_key]

    for idx, it_state in enumerate(current_items):
        img_idx = it_state["img_idx"]
        vn_desc = it_state["vn_desc"]
        
        st.markdown("---")
        col_img, col_form = st.columns([1.2, 1.8], gap="medium")
        
        with col_img:
            st.markdown(f"### 🖼️ Món {idx+1}: **{vn_desc}**")
            item_imgs = images_by_item.get(img_idx, [])
            if item_imgs:
                img_cols = st.columns(min(len(item_imgs), 3))
                for idx_i, img_bytes in enumerate(item_imgs):
                    try:
                        img_pil = Image.open(io.BytesIO(img_bytes))
                        with img_cols[idx_i % len(img_cols)]:
                            st.image(img_pil, use_container_width=True)
                    except Exception: pass
            else:
                st.info("Không có ảnh đính kèm.")

        with col_form:
            st.markdown(f"### 📝 Khai báo Món {idx+1}")
            sug_mat, sug_meth, sug_prod = smart_suggest_customs(vn_desc)
            st.info(f"💡 **Gợi ý:** `{sug_mat} / {sug_meth} / {sug_prod}`")
            
            col_add, col_del = st.columns(2)
            with col_add:
                if st.button(f"➕ Tách dòng #{idx+1}", key=f"add_{case_key}_{idx}"):
                    new_item = current_items[idx].copy()
                    new_item["PROD"] = new_item["PROD"] + " (PARTS)"
                    current_items.insert(idx + 1, new_item)
                    for i_num, x in enumerate(current_items): x["No"] = i_num + 1
                    st.session_state[case_key] = current_items
                    st.rerun()
            with col_del:
                if st.button(f"🗑️ Xóa dòng #{idx+1}", key=f"del_{case_key}_{idx}"):
                    current_items.pop(idx)
                    for i_num, x in enumerate(current_items): x["No"] = i_num + 1
                    st.session_state[case_key] = current_items
                    st.rerun()

            c_m1, c_m2, c_m3 = st.columns(3)
            mat_opts = st.session_state.custom_materials + ["Khác..."]
            if it_state["MAT"] not in mat_opts: mat_opts.insert(0, it_state["MAT"])
            sel_mat = c_m1.selectbox("Chất liệu", mat_opts, index=mat_opts.index(it_state["MAT"]) if it_state["MAT"] in mat_opts else 0, key=f"mat_sel_{case_key}_{idx}")
            if sel_mat == "Khác...":
                new_mat = c_m1.text_input("Nhập chất liệu mới:", key=f"mat_txt_{case_key}_{idx}")
                if new_mat:
                    sel_mat = new_mat.upper().strip()
                    if sel_mat not in st.session_state.custom_materials:
                        st.session_state.custom_materials.append(sel_mat)
                        st.rerun()

            meth_opts = st.session_state.custom_weaves + ["Khác..."]
            if it_state["METH"] not in meth_opts: meth_opts.insert(0, it_state["METH"])
            sel_meth = c_m2.selectbox("Cách dệt", meth_opts, index=meth_opts.index(it_state["METH"]) if it_state["METH"] in meth_opts else 0, key=f"meth_sel_{case_key}_{idx}")
            if sel_meth == "Khác...":
                new_meth = c_m2.text_input("Nhập cách dệt mới:", key=f"meth_txt_{case_key}_{idx}")
                if new_meth:
                    sel_meth = new_meth.upper().strip()
                    if sel_meth not in st.session_state.custom_weaves:
                        st.session_state.custom_weaves.append(sel_meth)
                        st.rerun()

            txt_prod = c_m3.text_input("Tên SP (Tiếng Anh)", value=it_state["PROD"], key=f"prod_{case_key}_{idx}")
            final_desc_val = f"{sel_mat}/{sel_meth}/{txt_prod}"

            c1, c2, c3 = st.columns(3)
            sel_orig = c1.selectbox("Origin", ["CN", "VN"], index=0 if it_state["ORIGIN"]=="CN" else 1, key=f"orig_{case_key}_{idx}")
            qty_input = c2.number_input("Số lượng", value=int(it_state["QTY"]), min_value=1, key=f"qty_{case_key}_{idx}")
            price_input = c3.number_input("Đơn giá ($)", value=float(it_state["PRICE"]), min_value=0.01, step=0.05, key=f"prc_{case_key}_{idx}")

            if idx < len(current_items):
                current_items[idx].update({
                    "MAT": sel_mat, "METH": sel_meth, "PROD": txt_prod, "DESCRIPTION": final_desc_val,
                    "ORIGIN": sel_orig, "QTY": qty_input, "PRICE": price_input
                })

    st.session_state[case_key] = current_items

    # 4. THANH ĐIỀU HƯỚNG BÊN DƯỚI
    st.divider()
    col_prev, col_save, col_next = st.columns([1, 1.5, 1])
    with col_prev:
        if st.button("⬅️ Kiện trước", disabled=st.session_state.current_sheet_idx == 0):
            st.session_state.current_sheet_idx -= 1
            st.rerun()
    with col_save:
        if st.button("💾 GHI NHẬN KIỆN NÀY & NEXT", type="primary"):
            st.session_state.all_declared_data[selected_sheet] = {
                "ma_kien": current_ma_kien, "ctns_size": current_ctns_size, "items": current_items
            }
            st.success(f"✅ Đã ghi nhận kiện {current_ma_kien}!")
            if st.session_state.current_sheet_idx < len(case_sheets) - 1:
                st.session_state.current_sheet_idx += 1
            st.rerun()
    with col_next:
        if st.button("Bỏ qua / Kiện tiếp ➡️", disabled=st.session_state.current_sheet_idx >= len(case_sheets) - 1):
            st.session_state.current_sheet_idx += 1
            st.rerun()
