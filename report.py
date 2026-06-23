import streamlit as st
import pandas as pd

# --- ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="ระบบตรวจสอบข้อมูล", layout="wide")
st.title("ระบบตรวจสอบข้อมูลการสำรวจจราจร")


# ============================================================
# ฟังก์ชันช่วย
# ============================================================

def get_val(df, r, c_list):
    """อ่านค่าจาก DataFrame และแปลงวันที่เป็น DD-MM-YYYY"""
    for c in c_list:
        try:
            val = df.iat[r, c]
            if pd.isna(val) or str(val).strip() == "":
                continue
            if isinstance(val, pd.Timestamp):
                return val.strftime('%d-%m-%Y')
            try:
                return pd.to_datetime(val).strftime('%d-%m-%Y')
            except Exception:
                return str(val).strip()
        except IndexError:
            continue
    return ""


def get_multi_cells(df, r, c_start, c_end):
    """รวมค่าจากหลาย cell ในแถวเดียวกัน"""
    values = []
    for c in range(c_start, c_end + 1):
        val = df.iat[r, c]
        if pd.notna(val) and str(val).strip() != "":
            values.append(str(val).strip())
    return " ".join(values)


def check_anomaly(df_subset):
    """ตรวจสอบว่ามีค่าเกิน 20% ติดกัน 3 แถวหรือไม่"""
    is_anomaly = df_subset > 0.2
    rows_exceed = is_anomaly.any(axis=1)
    for i in range(len(rows_exceed) - 2):
        if rows_exceed.iloc[i:i + 3].all():
            return True
    return False


def color_diff(val):
    """ระบายสีตามค่าบวก/ลบ"""
    if isinstance(val, (int, float)):
        return f'color: {"red" if val < 0 else "green"}; font-weight: bold'
    return ''


def highlight_and_format(df_target):
    """แสดงเป็น % และ highlight ถ้าเกิน 20%"""
    def to_pct(x):
        try:
            return f"{float(x) * 100:.1f}%"
        except Exception:
            return x

    def check_color(val):
        try:
            num = float(str(val).replace('%', ''))
            return 'background-color: #ffcccc; color: #cc0000; font-weight: bold' if num > 20 else ''
        except Exception:
            return ''

    return df_target.map(to_pct).style.map(check_color)


def show_hourly_tables(df, start, end):
    """แสดงตารางเปรียบเทียบรายชั่วโมง วันที่ 1 vs วันที่ 2"""
    p1 = df.iloc[start:end, 4]
    p2 = df.iloc[start:end, 14]
    c1 = df.iloc[start:end, 8]
    c2 = df.iloc[start:end, 18]

    df_p = pd.DataFrame({
        "คน(ว1)": p1.values,
        "คน(ว2)": p2.values,
        "ผลต่างคน": p2.values - p1.values,
    })
    df_c = pd.DataFrame({
        "รถ(ว1)": c1.values,
        "รถ(ว2)": c2.values,
        "ผลต่างรถ": c2.values - c1.values,
    })

    c1_col, c2_col = st.columns(2)
    with c1_col:
        st.dataframe(df_p.style.map(color_diff, subset=["ผลต่างคน"]), use_container_width=True)
    with c2_col:
        st.dataframe(df_c.style.map(color_diff, subset=["ผลต่างรถ"]), use_container_width=True)


# ============================================================
# ส่วนที่ 1: กรอกข้อมูล
# ============================================================

st.header("1. กรอกข้อมูลการสำรวจ")

TEAM_MEMBERS = [
    "บวรพลภ์ สุนทราธนาทิพย์",
    "ทิวากรณ์ จันดาดี",
    "วิสุทธิ์ อำพันธ์พงศ์",
    "สวาท เพียรภูเขา",
    "สวาสดิ์ กันธินาม",
    "สาวิตรี พิมยนต์",
    "อนุสิทธิ์ ผลสวัสดิ์",
]

col1, col2 = st.columns(2)

with col1:
    code = st.text_input("รหัสจับตัวเลข")
    site_name = st.text_input("Site")
    capture_options = st.multiselect(
        "รูปแบบการจับ",
        ["แบบที่1", "แบบที่2", "แบบที่3", "แบบที่4", "แบบที่5", "แบบที่6", "คนขึ้นสะพาน", "คนลงสะพาน"],
    )
    capture_type2 = st.selectbox(
        "จับตัวเลขประเภทรถ",
        ["1,2", "1,2+3,4", "1,2+3,4+5,6", "1,2+3,4+5,6+7,8", "1,2+5,6+7,8"],
    )
    leader1_list = st.multiselect("ชื่อหัวหน้าทีม (Site 1)", TEAM_MEMBERS)
    leader2_list = st.multiselect("ชื่อหัวหน้าทีม (Site 2)", TEAM_MEMBERS)

with col2:
    date1 = st.date_input("วันที่สำรวจ 1")
    date2 = st.date_input("วันที่สำรวจ 2")
    pt1_1 = st.text_input("พนักงานเช้าวันที่ 1")
    pt1_2 = st.text_input("พนักงานดึกวันที่ 1")
    pt2_1 = st.text_input("พนักงานเช้าวันที่ 2")
    pt2_2 = st.text_input("พนักงานดึกวันที่ 2")


# ============================================================
# ส่วนที่ 2: อัปโหลดและตรวจสอบข้อมูลหลัก
# ============================================================

st.header("2. อัปโหลดไฟล์เพื่อตรวจสอบ")
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"], key="main")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=None)

        # --- เตรียมค่าเพื่อเปรียบเทียบ ---
        leader1_str = ", ".join(leader1_list)
        leader2_str = ", ".join(leader2_list)
        selected_capture_str = ", ".join(capture_options)
        date1_str = date1.strftime('%d-%m-%Y')
        date2_str = date2.strftime('%d-%m-%Y')

        checks = {
            "รหัสจับตัวเลข":      (str(code),               get_val(df, 0, [4, 9])),
            "Site":                (str(site_name),           get_val(df, 0, [14, 19])),
            "รูปแบบการจับ":        (selected_capture_str,     get_val(df, 1, [4, 9])),
            "จับตัวเลขประเภทรถ":  (str(capture_type2),       get_val(df, 1, [14, 19])),
            "หัวหน้าทีม 1":        (leader1_str,              get_val(df, 2, [4, 9])),
            "หัวหน้าทีม 2":        (leader2_str,              get_val(df, 2, [14, 19])),
            "วันที่ 1":            (date1_str,                get_val(df, 3, [4, 9])),
            "วันที่ 2":            (date2_str,                get_val(df, 3, [14, 19])),
            "P/T เช้าวันที่ 1":   (str(pt1_1),               get_val(df, 4, [4, 9])),
            "P/T ดึกวันที่ 1":    (str(pt1_2),               get_val(df, 5, [4, 9])),
            "P/T เช้าวันที่ 2":   (str(pt2_1),               get_val(df, 4, [14, 19])),
            "P/T ดึกวันที่ 2":    (str(pt2_2),               get_val(df, 5, [14, 19])),
        }

        # --- ผลการตรวจสอบ ---
        st.divider()
        st.subheader("ผลการตรวจสอบ")

        if not all([code, site_name, capture_options, leader1_list]):
            st.warning("⚠️ กรุณากรอกข้อมูลในช่อง Input ให้ครบถ้วนก่อนตรวจสอบไฟล์")
        else:
            errors = 0
            for label, (in_val, file_val) in checks.items():
                if in_val == file_val:
                    st.success(f"✅ {label}: ตรงกัน")
                else:
                    st.error(f"❌ {label}: ไม่ตรงกัน (กรอก: {in_val} / พบ: {file_val})")
                    errors += 1

            if errors == 0:
                st.balloons()
                st.success("ข้อมูลถูกต้องครบถ้วน!")
            else:
                st.error(f"พบข้อผิดพลาดทั้งหมด {errors} จุด")

        # --- รายละเอียดเพิ่มเติม ---
        st.divider()
        st.subheader("📊 รายละเอียดเพิ่มเติมจากไฟล์")

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.info("**ข้อมูล วันที่ 1**")
            st.write(f"รายละเอียด: {get_multi_cells(df, 40, 2, 7)}")
            st.write(f"เหตุการณ์พิเศษ: {get_multi_cells(df, 41, 2, 7)}")
            st.write(f"เวลาพัก: {get_multi_cells(df, 42, 2, 7)}")

        with col_s2:
            st.info("**ข้อมูล วันที่ 2**")
            st.write(f"รายละเอียด: {get_multi_cells(df, 40, 12, 19)}")
            st.write(f"เหตุการณ์พิเศษ: {get_multi_cells(df, 41, 12, 19)}")
            st.write(f"เวลาพัก: {get_multi_cells(df, 42, 12, 19)}")
            st.write(f"Backup: {get_multi_cells(df, 43, 12, 19)}")

        # ============================================================
        # ส่วนที่ 3: ตรวจสอบความผิดปกติ
        # ============================================================

        st.divider()
        st.subheader("⚠️ สรุปผลการตรวจสอบความผิดปกติ (เกิน 20%)")

        subsets = [
            ("เช้า",  df.iloc[8:16,  22:24]),
            ("บ่าย",  df.iloc[19:27, 22:24]),
            ("ดึก",   df.iloc[30:38, 22:24]),
        ]
        labels = ["คน", "รถ"]
        anomalies_report = []

        for period_name, data in subsets:
            for i in range(2):
                if check_anomaly(data.iloc[:, [i]]):
                    anomalies_report.append(f"ช่วง{period_name} - {labels[i]}")

        if anomalies_report:
            st.error("พบข้อมูลเกิน 20% ติดกัน 3 แถว ในจุดต่อไปนี้:")
            for item in anomalies_report:
                st.write(f"- 🚩 {item}")
        else:
            st.success("✅ ไม่พบความผิดปกติ (ข้อมูลอยู่ในเกณฑ์ปกติทั้งหมด)")

        # ============================================================
        # ส่วนที่ 4: ตารางเปรียบเทียบรายช่วง
        # ============================================================

        st.divider()
        st.subheader("📋 ตารางเปรียบเทียบข้อมูล (แดงหาก > 20%)")

        subset_morning = df.iloc[8:16,  22:24].copy()
        subset_afternoon = df.iloc[19:27, 22:24].copy()
        subset_night = df.iloc[30:38, 22:24].copy()

        subset_morning.columns   = ["คนเช้า",  "รถเช้า"]
        subset_afternoon.columns = ["คนบ่าย",  "รถบ่าย"]
        subset_night.columns     = ["คนดึก",   "รถดึก"]

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.dataframe(highlight_and_format(subset_morning),   use_container_width=True)
        with col_b:
            st.dataframe(highlight_and_format(subset_afternoon), use_container_width=True)
        with col_c:
            st.dataframe(highlight_and_format(subset_night),     use_container_width=True)

        # ============================================================
        # ส่วนที่ 5: ตารางรายชั่วโมง วันที่ 1 vs วันที่ 2
        # ============================================================

        st.divider()
        st.subheader("📊 ตารางเปรียบเทียบรายชั่วโมง: วันที่ 1 vs วันที่ 2")

        for label, start, end in [("เช้า", 8, 16), ("บ่าย", 19, 27), ("ดึก", 30, 38)]:
            st.markdown(f"**ช่วง{label}**")
            show_hourly_tables(df, start, end)

        # ============================================================
        # ส่วนที่ 6: เปรียบเทียบกับไซต์อื่น
        # ============================================================

        st.divider()
        st.header("🔗 เปรียบเทียบกับไซต์อื่น")
        uploaded_compare_file = st.file_uploader("อัปโหลดไฟล์ไซต์อื่น", type=["xlsx"], key="compare")

        if uploaded_compare_file:
            try:
                df2 = pd.read_excel(uploaded_compare_file, header=None)

                # --- ฟังก์ชันดึงยอดรวมคนและรถ ต่อวัน ต่อช่วง ---
                # layout: คน วันที่ 1 = col 4, รถ วันที่ 1 = col 8
                #         คน วันที่ 2 = col 14, รถ วันที่ 2 = col 18
                PERIODS = [
                    ("เช้า",  8,  16),
                    ("บ่าย", 19,  27),
                    ("ดึก",  30,  38),
                ]

                def sum_period(source_df, row_start, row_end, col):
                    try:
                        vals = pd.to_numeric(source_df.iloc[row_start:row_end, col], errors='coerce')
                        return vals.sum()
                    except Exception:
                        return 0

                def build_summary(source_df):
                    rows = []
                    for period, r0, r1 in PERIODS:
                        rows.append({
                            "ช่วง": period,
                            "คน ว1": sum_period(source_df, r0, r1, 4),
                            "รถ ว1": sum_period(source_df, r0, r1, 8),
                            "คน ว2": sum_period(source_df, r0, r1, 14),
                            "รถ ว2": sum_period(source_df, r0, r1, 18),
                        })
                    # แถวยอดรวม
                    total = {"ช่วง": "รวมทั้งหมด"}
                    for col in ["คน ว1", "รถ ว1", "คน ว2", "รถ ว2"]:
                        total[col] = sum(r[col] for r in rows)
                    rows.append(total)
                    return pd.DataFrame(rows).set_index("ช่วง")

                summary_main    = build_summary(df)
                summary_compare = build_summary(df2)

                site_main_name    = get_val(df,  0, [14, 19]) or "ไฟล์หลัก"
                site_compare_name = get_val(df2, 0, [14, 19]) or "ไฟล์เปรียบเทียบ"

                # --- แสดงตารางยอดรวม side-by-side ---
                st.subheader("📊 ยอดรวมคนและรถ ทั้งวันที่ 1 และวันที่ 2")

                col_m, col_c, col_d = st.columns(3)

                with col_m:
                    st.markdown(f"**🏠 {site_main_name}**")
                    st.dataframe(summary_main.style.format("{:.0f}"), use_container_width=True)

                with col_c:
                    st.markdown(f"**🏢 {site_compare_name}**")
                    st.dataframe(summary_compare.style.format("{:.0f}"), use_container_width=True)

                with col_d:
                    # ผลต่าง = ไฟล์เปรียบเทียบ - ไฟล์หลัก
                    diff = summary_compare - summary_main

                    def color_diff_df(val):
                        try:
                            num = float(val)
                            if num > 0:
                                return 'color: green; font-weight: bold'
                            elif num < 0:
                                return 'color: red; font-weight: bold'
                        except Exception:
                            pass
                        return ''

                    st.markdown(f"**📐 ผลต่าง ({site_compare_name} − {site_main_name})**")
                    st.dataframe(
                        diff.style.format("{:+.0f}").map(color_diff_df),
                        use_container_width=True,
                    )

                # --- สรุปเปอร์เซ็นต์ความต่าง ---
                st.subheader("📐 สรุปเปอร์เซ็นต์ความต่าง")

                pct = (diff / summary_main.replace(0, pd.NA) * 100).fillna(0)

                def highlight_pct(val):
                    try:
                        num = float(val.replace('%', '').replace('+', ''))
                        if abs(num) > 20:
                            return 'background-color: #ffcccc; color: #cc0000; font-weight: bold'
                        elif abs(num) > 10:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                    except Exception:
                        pass
                    return ''

                st.dataframe(
                    pct.style.format("{:+.1f}%").map(highlight_pct),
                    use_container_width=True,
                )
                st.caption("🔴 แดง = ต่างกันเกิน 20%  |  🟡 เหลือง = ต่างกันเกิน 10%")

            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์เปรียบเทียบ: {e}")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")