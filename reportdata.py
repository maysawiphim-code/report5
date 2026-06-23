import streamlit as st
import tempfile
import os
import io
from processor import process_files

st.set_page_config(
    page_title="ระบบสรุปผลการจับตัวเลข",
    page_icon="📊",
    layout="centered",
)

# ---- Custom CSS ----
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #1a237e 0%, #1565c0 100%);
    color: white;
    padding: 32px 28px 24px;
    border-radius: 16px;
    margin-bottom: 28px;
    text-align: center;
}
.main-header h1 { font-size: 28px; margin: 0 0 6px; font-weight: 700; }
.main-header p  { font-size: 15px; margin: 0; opacity: 0.85; }

.step-bar {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0;
    margin-bottom: 28px;
}
.step { text-align: center; min-width: 80px; }
.step-num {
    width: 32px; height: 32px;
    border-radius: 50%;
    display: inline-flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 14px;
    background: #e8eaf6; color: #1a237e;
    margin-bottom: 4px;
}
.step.done .step-num   { background: #1a237e; color: white; }
.step.active .step-num { background: #1565c0; color: white; }
.step-label { font-size: 12px; color: #666; }
.step-line  { flex: 1; height: 2px; background: #ddd; max-width: 60px; margin-bottom: 16px; }

.info-box {
    background: #fff8e1;
    border-left: 4px solid #f9a825;
    border-radius: 6px;
    padding: 12px 16px;
    font-size: 14px;
    margin-bottom: 16px;
    color: #5d4037;
}
.success-box {
    background: #e8f5e9;
    border: 1px solid #a5d6a7;
    border-radius: 12px;
    padding: 24px;
    text-align: center;
    margin-bottom: 20px;
}
.success-box h2 { color: #2e7d32; margin-bottom: 8px; }
.success-box p  { color: #555; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ---- Header ----
st.markdown("""
<div class="main-header">
  <h1>📊 ระบบสรุปผลการจับตัวเลข</h1>
  <p>อัพโหลดไฟล์สรุป 5 เพื่อสร้างสรุป 1 และสรุปสุดท้ายอัตโนมัติ</p>
</div>
""", unsafe_allow_html=True)

# ---- Session state ----
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'out1_bytes' not in st.session_state:
    st.session_state.out1_bytes = None
if 'out_final_bytes' not in st.session_state:
    st.session_state.out_final_bytes = None

step = st.session_state.step

# ---- Step bar ----
def step_cls(n):
    if n < step: return 'done'
    if n == step: return 'active'
    return ''

st.markdown(f"""
<div class="step-bar">
  <div class="step {step_cls(1)}">
    <div class="step-num">{'✓' if step > 1 else '1'}</div>
    <div class="step-label">อัพโหลด</div>
  </div>
  <div class="step-line"></div>
  <div class="step {step_cls(2)}">
    <div class="step-num">{'✓' if step > 2 else '2'}</div>
    <div class="step-label">ประมวลผล</div>
  </div>
  <div class="step-line"></div>
  <div class="step {step_cls(3)}">
    <div class="step-num">3</div>
    <div class="step-label">ดาวน์โหลด</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ---- Step 1: Upload ----
if step == 1:
    st.markdown("""
    <div class="info-box">
    💡 <strong>รูปแบบไฟล์ที่รองรับ:</strong> ไฟล์ .xlsx ที่มีโครงสร้างเหมือน <strong>สรุป 5</strong>
    — มีข้อมูล 2 วัน (คอลัมน์ G–K และ L–P) พร้อมข้อมูลคนและรถ
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "เลือกไฟล์สรุป 5 (.xlsx)",
        type=["xlsx"],
        label_visibility="visible",
    )

    template_file = st.file_uploader(
        "📋 อัพโหลดไฟล์ template สรุปสุดท้าย (.xlsx) — ถ้าไม่อัพโหลด ระบบจะใช้ template ที่มีอยู่",
        type=["xlsx"],
        label_visibility="visible",
        key="template_upload",
    )

    if uploaded:
        st.success(f"✅ เลือกไฟล์แล้ว: **{uploaded.name}**")
        if st.button("🚀 สร้างไฟล์รายงาน", type="primary", use_container_width=True):
            st.session_state.step = 2
            st.session_state.input_bytes = uploaded.read()
            st.session_state.input_name = uploaded.name
            st.session_state.template_bytes = template_file.read() if template_file else None
            st.rerun()

# ---- Step 2: Process ----
elif step == 2:
    with st.spinner("🔄 กำลังประมวลผลข้อมูล..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                in_path  = os.path.join(tmpdir, "input.xlsx")
                out1     = os.path.join(tmpdir, "สรุป1.xlsx")
                out_fin  = os.path.join(tmpdir, "สรุปสุดท้าย.xlsx")
                tmpl_path = None

                with open(in_path, 'wb') as f:
                    f.write(st.session_state.input_bytes)

                if st.session_state.get('template_bytes'):
                    tmpl_path = os.path.join(tmpdir, "template_final.xlsx")
                    with open(tmpl_path, 'wb') as f:
                        f.write(st.session_state.template_bytes)
                else:
                    # Use bundled template
                    tmpl_path = os.path.join(os.path.dirname(__file__), 'template_final.xlsx')

                process_files(in_path, out1, out_fin, template_path=tmpl_path)

                with open(out1, 'rb') as f:
                    st.session_state.out1_bytes = f.read()
                with open(out_fin, 'rb') as f:
                    st.session_state.out_final_bytes = f.read()

            st.session_state.step = 3
            st.rerun()

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            if st.button("↩ กลับไปอัพโหลดใหม่"):
                st.session_state.step = 1
                st.rerun()

# ---- Step 3: Download ----
elif step == 3:
    st.markdown("""
    <div class="success-box">
      <div style="font-size:52px">✅</div>
      <h2>สร้างไฟล์รายงานสำเร็จ!</h2>
      <p>กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์รายงานทั้ง 2 ไฟล์</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📋 ดาวน์โหลด สรุป 1",
            data=st.session_state.out1_bytes,
            file_name="สรุป1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption("ตารางสรุปผลเฉลี่ยคนและรถ")

    with col2:
        st.download_button(
            label="📊 ดาวน์โหลด สรุปสุดท้าย",
            data=st.session_state.out_final_bytes,
            file_name="สรุปสุดท้าย.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption("ตารางสรุปพร้อมข้อมูลรายผลัด")

    st.divider()
    if st.button("↩ อัพโหลดไฟล์ใหม่", use_container_width=True):
        for key in ['step','out1_bytes','out_final_bytes','input_bytes','template_bytes']:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state.step = 1
        st.rerun()
