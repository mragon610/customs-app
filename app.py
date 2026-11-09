import streamlit as st
import pandas as pd
import openpyxl
from PIL import Image
import io

st.set_page_config(page_title="Hệ thống Kê khai Hải quan INV-PKL", layout="wide")

# --- HỆ THỐNG PHÂN TÍCH & GỢI Ý THÔNG MINH DỰA TRÊN TÊN TIẾNG VIỆT ---
def smart_suggest_customs(vietnamese_name):
    name_lc = str(vietnamese_name).strip().lower()
    
    # Đồ chơi / Đồ trang trí / Đồ nhựa
    if "trang trí" in name_lc or "đồ chơi" in name_lc or "spinner" in name_lc or "cốt cắm hoa" in name_lc or "xốp" in name_lc:
        if "nhựa" in name_lc or "cốt" in name_lc:
            return ("PLASTIC", "FACTORY", "PLASTIC DECORATION PARTS")
        if "xốp" in name_lc:
            return ("FOAM", "FACTORY", "FLORAL FOAM")
        return ("PLASTIC", "FACTORY", "PLASTIC TOY / DECORATION")
    
    # Giày dép & Đồ da
    if "dép" in name_lc:
        return ("PVC", "FACTORY", "PVC SANDAL")
    if "giày" in name_lc:
        return ("SYNTHETIC LEATHER", "FACTORY", "CASUAL SHOES")
    if "bốt" in name_lc:
        return ("SYNTHETIC LEATHER", "FACTORY", "BOOTS")
    
    # Áo khoác & Áo ấm
    if "khoác da" in name_lc:
        return ("ARTIFICIAL LEATHER", "FACTORY", "ARTIFICIAL LEATHER JACKET")
    if "khoác dạ" in name_lc or "dạ" in name_lc:
        return ("COTTON", "KNIT", "MEN JACKET")
    if "khoác nỉ" in name_lc or "hoodie" in name_lc:
        return ("COTTON", "KNIT", "LONG SLEEVES HOODIE")
    if "khoác lông" in name_lc:
        return ("POLYESTER", "KNIT", "FUR JACKET")
    if "khoác" in name_lc:
        return ("POLYESTER", "KNIT", "PUFFER JACKET")
    if "nỉ" in name_lc or "sweatshirt" in name_lc:
        return ("COTTON", "KNIT", "SWEATSHIRT")
    if "len" in name_lc:
        return ("COTTON", "KNIT", "LONG SLEEVES PULLOVER")
        
    # Quần áo các loại
    if "thun" in name_lc or "cộc tay" in name_lc:
        return ("COTTON", "KNIT", "T-SHIRT")
    if "dài tay" in name_lc or "sơ mi" in name_lc:
        return ("COTTON", "KNIT", "LONG SLEEVES SHIRT")
    if "3 lỗ" in name_lc or "hai dây" in name_lc:
        return ("COTTON", "KNIT", "TANK TOP")
    if "jean ngắn" in name_lc or "short jean" in name_lc:
        return ("DENIM", "WOVEN", "DENIM SHORTS")
    if "jean" in name_lc:
        return ("DENIM", "WOVEN", "DENIM PANTS")
    if "quần short" in name_lc or "quần ngắn" in name_lc:
        return ("COTTON", "KNIT", "SHORTS")
    if "quần" in name_lc:
        return ("COTTON", "KNIT", "CASUAL PANTS")
        
    # Váy & Đầm
    if "chân váy" in name_lc:
        return ("COTTON", "WOVEN", "SKIRT")
    if "váy" in name_lc or "đầm" in name_lc:
        return ("COTTON", "KNIT", "ONE PIECE DRESS")
        
    # Tóc giả & Phụ kiện
    if "tóc giả" in name_lc or "lưới" in name_lc:
        return ("SYNTHETIC FIBER", "FACTORY", "SYNTHETIC HAIR WIG")
    if "lược" in name_lc:
        return ("PLASTIC", "FACTORY", "PLASTIC COMB")
    if "gấu bông" in name_lc or "móc khoá" in name_lc:
        return ("PLUSH FABRIC", "FACTORY", "STUFFED TOY KEYCHAIN")
    if "hoa" in name_lc or "giấy gói" in name_lc or "ruy băng" in name_lc or "dây" in name_lc:
        return ("PAPER", "FACTORY", "FLOWER PACKAGING MATERIAL")
    if "ga giường" in name_lc or "khăn" in name_lc:
        return ("COTTON", "WOVEN", "BED LINEN / TOWEL")

    # Mặc định thông minh
    return ("COTTON", "KNIT", "LONG SLEEVES")

# Khởi tạo bộ nhớ danh mục tùy chọn có thể thêm/xóa linh hoạt
if "custom_materials" not in st.session_state:
    st.session_state.custom_materials = ["COTTON", "PVC", "DENIM", "PLUSH FABRIC", "ARTIFICIAL LEATHER", "POLYESTER", "NYLON", "SILK", "PLASTIC", "FOAM", "PAPER", "SYNTHETIC FIBER"]
if "custom_weaves" not in st.session_state:
    st.session_state.custom_weaves = ["KNIT", "WOVEN", "FACTORY"]

st.title("📦 Hệ thống Tự động hóa Kê khai Hải quan INV-PKL")

uploaded_file = st.sidebar.file_uploader("Tải lên File Excel Template 4505 mới", type=["xlsx"])

if "all_declared_data" not in st.session_state:
    st.session_state.all_declared_data = {}
if "current_sheet_idx" not in st.session_state:
    st.session_state.current_sheet_idx = 0

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    
    # 1. ĐỌC VÀ GOM NHÓM 100% TỪ TAB "Tổng Hợp" CỦA FILE MỚI
    all_items_by_case = {}
    case_order_list = []
    try:
        sum_sheet_name = [s for s in wb.sheetnames if s.lower() in ["tổng hợp", "tong hop"]][0]
        summary_df = pd.read_excel(uploaded_file, sheet_name=sum_sheet_name, header=1)
        summary_df = summary_df.iloc[:, :6] 
        summary_df.columns = ["STT", "Ma_Kien", "Mat_Hang", "So_Luong", "Xuat_Xu", "Can_Nang"]
        summary_df["Ma_Kien"] = summary_df["Ma_Kien"].ffill() # Kéo fill mã kiện chuẩn 100%
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
    if not case_sheets:
        st.error("⚠️ Không tìm thấy các sheet kiện hàng!")
        st.stop()

    if st.session_state.current_sheet_idx >= len(case_sheets):
        st.session_state.current_sheet_idx = len(case_sheets) - 1

    selected_sheet = st.sidebar.selectbox("Chọn kiện hàng:", case_sheets, index=st.session_state.current_sheet_idx)
    
    if case_sheets.index(selected_sheet) != st.session_state.current_sheet_idx:
        st.session_state.current_sheet_idx = case_sheets.index(selected_sheet)
        st.rerun()

    ws_case = wb[selected_sheet]
    current_ma_kien = str(ws_case.cell(row=3, column=4).value).strip()
    current_ctns_size = str(ws_case.cell(row=3, column=8).value).strip()
    if current_ctns_size == "None" or not current_ctns_size: current_ctns_size = "50*40*25"
        
    items_in_case = all_items_by_case.get(current_ma_kien, [])
    if not items_in_case:
        items_in_case = [{"desc_vn": "Mẫu", "qty": 1, "origin": "CN"}]

    # Trút hình ảnh theo tọa độ tịnh tiến
    images_by_item = {}
    if hasattr(ws_case, '_images') and ws_case._images:
        for img in ws_case._images:
            try:
                row_idx = img.anchor._from.row + 1 
                item_index = (row_idx - 4) // 14
                if item_index >= 0:
                    if item_index not in images_by_item: images_by_item[item_index] = []
                    images_by_item[item_index].append(img._data())
            except Exception:
                pass

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
        row_start = it_state["row_start"]
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
                    except Exception:
                        pass
            else:
                st.info("Không có ảnh đính kèm.")

        with col_form:
            st.markdown(f"### 📝 Khai báo Món {idx+1}")
            
            # --- HIỂN THỊ GỢI Ý CHUẨN HẢI QUAN ĐỂ THAM KHẢO ---
            sug_mat, sug_meth, sug_prod = smart_suggest_customs(vn_desc)
            st.info(f"💡 **Gợi ý chuẩn Hải quan Nhật Bản:** `{sug_mat} / {sug_meth} / {sug_prod}`")
            
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

            # --- 3 Ô LỰA CHỌN KÊ KHAI (CÓ HỖ TRỢ XÓA / THÊM TÙY CHỌN SAI CHÍNH TẢ) ---
            c_m1, c_m2, c_m3 = st.columns(3)
            
            # 1. Chất liệu
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
            elif sel_mat not in ["COTTON", "PVC", "DENIM", "PLUSH FABRIC", "ARTIFICIAL LEATHER", "POLYESTER", "NYLON", "SILK", "PLASTIC", "FOAM", "PAPER", "SYNTHETIC FIBER"]:
                if c_m1.button(f"❌ Xóa '{sel_mat}'", key=f"del_mat_{case_key}_{idx}"):
                    st.session_state.custom_materials.remove(sel_mat)
                    st.rerun()

            # 2. Cách dệt
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
            elif sel_meth not in ["KNIT", "WOVEN", "FACTORY"]:
                if c_m2.button(f"❌ Xóa '{sel_meth}'", key=f"del_meth_{case_key}_{idx}"):
                    st.session_state.custom_weaves.remove(sel_meth)
                    st.rerun()

            # 3. Tên sản phẩm tiếng Anh
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

    # 6. ĐIỀU HƯỚNG & XUẤT FILE MỚI HOÀN TOÀN ĐỘC LẬP
    st.divider()
    is_current_saved = selected_sheet in st.session_state.all_declared_data
    
    col_prev, col_save, col_next = st.columns([1, 1.5, 1])
    with col_prev:
        if st.button("⬅️ Kiện trước", disabled=st.session_state.current_sheet_idx == 0):
            st.session_state.current_sheet_idx -= 1
            st.rerun()
    with col_save:
        if st.button("💾 GHI NHẬN KIỆN NÀY", type="primary"):
            st.session_state.all_declared_data[selected_sheet] = {
                "ma_kien": current_ma_kien, "ctns_size": current_ctns_size, "items": current_items
            }
            st.success(f"✅ Đã ghi nhận kiện {current_ma_kien} thành công!")
            st.rerun()
    with col_next:
        if st.button("Kiện tiếp ➡️", disabled=st.session_state.current_sheet_idx >= len(case_sheets) - 1):
            st.session_state.current_sheet_idx += 1
            st.rerun()

    st.divider()
    unsigned_sheets = [s for s in case_sheets if s not in st.session_state.all_declared_data]
    if unsigned_sheets:
        st.warning(f"⚠️ Còn kiện chưa lưu: {', '.join(unsigned_sheets)}")
    else:
        st.success("🎉 Đã ghi nhận đủ toàn bộ các kiện hàng!")

    if st.button("🚀 XUẤT FILE MỚI HOÀN TOÀN ĐỘC LẬP"):
        if unsigned_sheets:
            st.error("❌ Vui lòng ghi nhận tất cả các kiện trước khi xuất file!")
        else:
            output_wb = openpyxl.load_workbook(uploaded_file)
            
            # Tạo tab mới hoàn toàn tên là "INV-PKL-Ket-Qua" để KHÔNG BAO GIỜ ghi đè lên tab INV-PKL gốc
            new_sheet_name = "INV-PKL-Ket-Qua"
            if new_sheet_name in output_wb.sheetnames:
                output_wb.remove(output_wb[new_sheet_name])
            ws_out = output_wb.create_sheet(new_sheet_name)
            
            # Sao chép phần header từ tab INV-PKL gốc sang tab kết quả mới
            if "INV-PKL" in output_wb.sheetnames:
                ws_orig = output_wb["INV-PKL"]
                for r in range(1, 13):
                    for c in range(1, 11):
                        ws_out.cell(row=r, column=c, value=ws_orig.cell(row=r, column=c).value)
            
            start_r = 13
            case_counter = 1
            for s_name in case_sheets:
                if s_name not in st.session_state.all_declared_data: continue
                c_data = st.session_state.all_declared_data[s_name]
                
                ws_out.cell(row=start_r, column=2, value="CASE No :")
                ws_out.cell(row=start_r, column=4, value=c_data["ma_kien"])
                ws_out.cell(row=start_r, column=6, value="CTNS SIZE :")
                ws_out.cell(row=start_r, column=8, value=c_data["ctns_size"])
                
                start_r += 1
                ws_out.cell(row=start_r, column=1, value=case_counter)
                headers = ["No.", "STYLE", "DESCRIPTION", "ORIGIN", "STYLE", "QTY", "PRICE", "AMOUNT", "CURRENCY"]
                for col_idx, h_text in enumerate(headers, 2):
                    ws_out.cell(row=start_r, column=col_idx, value=h_text)
                    
                start_r += 1
                first_row = start_r
                for it in c_data["items"]:
                    ws_out.cell(row=start_r, column=2, value=it["No"])
                    ws_out.cell(row=start_r, column=3, value=it["STYLE_HDR"])
                    ws_out.cell(row=start_r, column=4, value=it["DESCRIPTION"])
                    ws_out.cell(row=start_r, column=5, value=it["ORIGIN"])
                    ws_out.cell(row=start_r, column=6, value=it["STYLE_ITEM"])
                    ws_out.cell(row=start_r, column=7, value=int(it["QTY"]))
                    ws_out.cell(row=start_r, column=8, value=float(it["PRICE"]))
                    ws_out.cell(row=start_r, column=9, value=f"=G{start_r}*H{start_r}")
                    ws_out.cell(row=start_r, column=10, value="USD")
                    start_r += 1
                
                last_row = start_r - 1
                ws_out.cell(row=start_r, column=2, value="TOTAL")
                ws_out.cell(row=start_r, column=7, value=f"=SUM(G{first_row}:G{last_row})")
                ws_out.cell(row=start_r, column=8, value="PCS")
                ws_out.cell(row=start_r, column=9, value=f"=SUM(I{first_row}:I{last_row})")
                ws_out.cell(row=start_r, column=10, value="USD")
                
                start_r += 2
                case_counter += 1
                
            buf = io.BytesIO()
            output_wb.save(buf)
            buf.seek(0)
            st.download_button("📥 Tải xuống File Kết Quả Mới (.xlsx)", data=buf, file_name="INV-PKL_Ket_Qua_Moi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")