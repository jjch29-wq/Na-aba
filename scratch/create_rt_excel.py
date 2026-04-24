import pandas as pd
import os

# Define the data for RT Shooting Conditions
data = {
    '구분': [
        '기본 상세', '기본 상세', '기본 상세', '기본 상세',
        '촬영 파라미터', '촬영 파라미터', '촬영 파라미터', '촬영 파라미터',
        '품질 기준', '품질 기준', '품질 기준', '품질 기준',
        '작업 주의사항', '작업 주의사항', '작업 주의사항'
    ],
    '항목': [
        '검사 대상', '관경/두께', '용접 방법', '촬영 기술',
        '방사선원', '관전압 (kV)', '관전류 (mA)', '촬영 거리 (SFD)',
        '투과계 (IQI)', '필수 식별선', '필름 종별', '허용 농도',
        '이미지 분리', '마킹 배치', '표면 관리'
    ],
    '세부 사양': [
        'PIPE (Stainless Steel)', '1.5인치 / 1.5T', '자동용접 (Automatic TIG)', 'DWDI 타원촬영 (Elliptical)',
        'X-Ray (Portable)', '80 ~ 110 kV', '3 ~ 5 mA', '500 ~ 600 mm',
        'ASTM Set A (Wire Type)', '#5 (0.10mm) 이상', 'Fine Grain (IX50, D4급)', '2.0 ~ 3.5',
        '상/하부 비드 타원 분리 필수', '결함 판독 방해 금지', '스패터 및 오염 제거 후 촬영'
    ],
    '비고': [
        'SUS 배관 위주', '외경 38.1mm', '정밀 제어 요망', '90도 간격 2매 권장',
        '동위원소(Ir-192) 사용 금지', '박판용 저전압 유지', '장비 사양별 조정', '상의 선명도 우선',
        '절차서 및 규격 준수', '감도 2% 이내', '급미립자 필름 사용', '농도 중간값 권장',
        'Offset 각도 5~10도', '번호 및 날짜 식별', '의사지시 방지'
    ]
}

# Create DataFrame
df = pd.DataFrame(data)

# File path
save_path = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\RT_Shooting_Conditions_1.5T.xlsx'

# Save to Excel
try:
    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='RT_Conditions')
        
        # Adjust column width (optional but helps for readability)
        worksheet = writer.sheets['RT_Conditions']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 5
            worksheet.column_dimensions[chr(65+i)].width = max_len

    print(f"Excel file created successfully: {save_path}")
except Exception as e:
    print(f"Error creating excel: {e}")
