import streamlit as st
from docxtpl import DocxTemplate
import io
import datetime
import os

# 設定網頁標題與佈局
st.set_page_config(page_title="房仲物調表系統", page_icon="🏠")

def main():
    st.title("🏠 房仲物調表 - 快速填寫系統")
    st.markdown("請依序填寫下方資料，完成後點擊最下方的按鈕即可生成 Word 檔。")

    # 檢查範本是否存在
    template_name = "template.docx"
    if not os.path.exists(template_name):
        st.error(f"❌ 找不到範本檔案：{template_name}")
        st.warning("請將 Word 範本上傳至與程式相同的資料夾中。")
        return

    # 使用 Form 表單，讓使用者一次填完再按送出
    with st.form("survey_form"):
        
        # ---區塊 1: 基本資料---
        st.subheader("📋 基本資料")
        c1, c2 = st.columns(2)
        with c1:
            casename = st.text_input("1. 案名 (必填，將作為檔名)", placeholder="例如：信義之星")
            address = st.text_input("2. 物件地址")
            pr = st.text_input("3. 售價 (萬元)")
            phone = st.text_input("35. 承辦人電話") # 依照您提供的順序，這原本在最後，移到基本資料比較順手
        with c2:
            community = st.text_input("16. 社區名稱")
            feature = st.text_area("34. 房屋特色", height=100)

        # ---區塊 2: 坪數資料---
        st.subheader("📐 坪數資料")
        c1, c2, c3 = st.columns(3)
        with c1:
            totalping = st.text_input("4. 總建坪")
            public_ping = st.text_input("7. 公設坪數")
            addpos = st.text_input("10. 增建位置")
        with c2:
            main_ping = st.text_input("5. 主建物坪數")
            parkingping = st.text_input("8. 車位坪數")
            land_ping = st.text_input("31. 土地面積(坪)")
        with c3:
            sub_ping = st.text_input("6. 附屬建物坪數")
            addping = st.text_input("9. 增建坪數")

        # ---區塊 3: 樓層與屋齡---
        st.subheader("🏢 樓層與屋況")
        c1, c2, c3 = st.columns(3)
        with c1:
            totalfloor = st.text_input("11. 總樓層")
            builddate = st.text_input("14. 建築完成日")
            seat = st.text_input("32. 房屋坐向")
        with c2:
            myfloor = st.text_input("12. 位於樓層")
            age = st.text_input("15. 屋齡")
            face = st.text_input("32. 房屋面向") # 注意：您的原始碼 seat 和 face 都是 32，這裡分開處理
        with c3:
            underfloor = st.text_input("13. 地下幾層")
            moto = st.text_input("33. 機車車位")

        # ---區塊 4: 格局細節---
        st.subheader("🛋️ 格局配置")
        # 使用 5 個欄位並排
        cols = st.columns(5)
        room = cols[0].text_input("26. 房")
        hall = cols[1].text_input("27. 廳")
        bath = cols[2].text_input("28. 衛")
        kitchen = cols[3].text_input("29. 廚")
        balcony = cols[4].text_input("30. 陽台")

        # ---區塊 5: 社區與周邊---
        st.subheader("🌳 社區與周邊環境")
        c1, c2 = st.columns(2)
        with c1:
            fee = st.text_input("17. 管理費")
            units = st.text_input("19. 同層戶數")
            park = st.text_input("21. 附近公園")
            school = st.text_input("23. 附近學校")
            wi = st.text_input("24. 面寬幾米")
        with c2:
            totalunits = st.text_input("18. 總戶數")
            elevators = st.text_input("20. 電梯數")
            market = st.text_input("22. 附近市場")
            le = st.text_input("25. 臨路幾米")

        st.markdown("---")
        # 送出按鈕
        submitted = st.form_submit_button("✨ 產生 Word 物調表", type="primary")

    # --- 處理送出後的邏輯 ---
    if submitted:
        if not casename.strip():
            st.error("⚠️ 請輸入「案名」，否則無法產生檔案！")
            return

        # 1. 整理資料 (對應您原本 Word 的變數名稱)
        context = {
            "casename": casename,
            "address": address,
            "pr": pr,
            "totalping": totalping,
            "main_ping": main_ping,
            "sub_ping": sub_ping,
            "public_ping": public_ping,
            "parkingping": parkingping,
            "addping": addping,
            "addpos": addpos,
            "totalfloor": totalfloor,
            "myfloor": myfloor,
            "underfloor": underfloor,
            "builddate": builddate,
            "age": age,
            "community": community,
            "fee": fee,
            "totalunits": totalunits,
            "units": units,
            "elevators": elevators,
            "park": park,
            "market": market,
            "school": school,
            "wi": wi,
            "le": le,
            "room": room,
            "hall": hall,
            "bath": bath,
            "kitchen": kitchen,
            "balcony": balcony,
            "land_ping": land_ping,
            "seat": seat,
            "face": face,
            "moto": moto,
            "feature": feature,
            "phone": phone,
            "date": datetime.date.today().strftime("%Y/%m/%d")
        }

        # 2. 產生檔案 (使用記憶體 BytesIO，不存到硬碟)
        try:
            doc = DocxTemplate(template_name)
            doc.render(context)

            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0) # 指標歸零

            # 3. 顯示下載按鈕
            output_filename = f"物調表_{casename.strip()}.docx"
            
            st.success(f"✅ 成功生成！請點擊下方按鈕下載檔案：")
            st.download_button(
                label="📥 點擊下載 Word 檔",
                data=bio,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{e}")
            st.info("請檢查 Word 範本內容是否正確。")

if __name__ == "__main__":
    main()