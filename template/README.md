# 워드 보고서 템플릿

## 개요

이 템플릿은 Python의 `python-docx` 라이브러리를 사용하여 다양한 유형의 워드 보고서를 자동 생성하는 예시 코드를 제공합니다.

## 파일 구조

```
make_report/
├── template/
│   ├── README.md               # 이 파일
│   └── code/
│       └── report.py           # 보고서 생성 코드
└── image/                      # 예시 이미지들
    ├── Cursor_6MnMx69wWR.png
    ├── Cursor_3yPrL9z45P.png
    ├── Cursor_azrBa7z9ZQ.png
    └── Cursor_cEi6qlPj9A.png
```

## 보고서 유형

### 1. 기본 튜토리얼 보고서

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.shared import RGBColor

doc = Document()

# 제목 추가
title = doc.add_heading('보고서 입니다.', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 요약 섹션
summary = doc.add_heading('요약', level=1)
summary.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 내용 추가
content = doc.add_paragraph()
content.add_run('요약을 할 예정입니다\n').bold = True
content.add_run('주요 내용을 여기에 작성합니다\n').font.size = Pt(11)
content.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 이미지 테이블 (1행 2열)
image_table = doc.add_table(rows=1, cols=2)
left_cell = image_table.cell(0, 0)
left_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
left_cell.paragraphs[0].add_run().add_picture(r'image\mul.png', width=Cm(8))

right_cell = image_table.cell(0, 1)
right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
right_cell.paragraphs[0].add_run().add_picture(r'image\mul2.png', width=Cm(8))

doc.save('tutorial_report.docx')
```

![기본 보고서 예시](../image/Cursor_6MnMx69wWR.png)

### 2. 데이터 분석 보고서

matplotlib 차트를 포함한 데이터 분석 보고서

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
import matplotlib.pyplot as plt
import io

doc = Document()

# 제목
title = doc.add_heading('데이터 분석 보고서', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 차트 생성
fig, ax = plt.subplots(figsize=(8, 6))
ax.plot([1, 2, 3, 4], [10, 20, 25, 30], 'o-')
ax.set_title('성능 향상 추이')

# 메모리에서 이미지 처리
img_buffer = io.BytesIO()
fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
img_buffer.seek(0)

# 워드에 차트 삽입
chart_para = doc.add_paragraph()
chart_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
chart_para.add_run().add_picture(img_buffer, width=Cm(12))

doc.save('data_analysis_report.docx')
```

![데이터 분석 보고서 예시](../image/Cursor_3yPrL9z45P.png)

### 3. 프로젝트 진행 보고서

테이블을 포함한 프로젝트 상황 보고서

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

doc = Document()

# 제목
title = doc.add_heading('프로젝트 진행 현황 보고서', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 진행 상황 테이블
doc.add_heading('진행 현황', level=1)
table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'

# 헤더
hdr_cells = table.rows[0].cells
hdr_cells[0].text = '작업 항목'
hdr_cells[1].text = '담당자'
hdr_cells[2].text = '진행률'
hdr_cells[3].text = '완료 예정일'

# 데이터 행 추가
tasks = [
    ['데이터 수집', '김개발', '100%', '2024-12-10'],
    ['모델 개발', '이분석', '80%', '2024-12-20'],
    ['테스트', '박검증', '30%', '2024-12-25']
]

for task in tasks:
    row_cells = table.add_row().cells
    for i, cell_text in enumerate(task):
        row_cells[i].text = cell_text

doc.save('project_progress_report.docx')
```

![프로젝트 진행 보고서 예시](../image/Cursor_azrBa7z9ZQ.png)

### 4. 실험 결과 보고서

상관관계 분석 결과를 포함한 실험 보고서

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import io

doc = Document()

# 제목
title = doc.add_heading('실험 결과 분석 보고서', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 실험 데이터
data = pd.DataFrame({
    'A': [1, 2, 3, 4, 5],
    'B': [2, 4, 6, 8, 10],
    'C': [5, 4, 3, 2, 1]
})

# 히트맵 생성
fig, ax = plt.subplots(figsize=(8, 6))
sns.heatmap(data.corr(), annot=True, cmap='coolwarm', center=0, ax=ax)
ax.set_title('변수 간 상관관계 히트맵')

# 워드에 히트맵 삽입
img_buffer = io.BytesIO()
fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
img_buffer.seek(0)

doc.add_heading('분석 결과', level=1)
result_para = doc.add_paragraph()
result_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
result_para.add_run().add_picture(img_buffer, width=Cm(12))

doc.save('experiment_result_report.docx')
```

![실험 결과 보고서 예시](../image/Cursor_cEi6qlPj9A.png)

## 필요한 라이브러리

```bash
pip install python-docx matplotlib seaborn pandas
```

## 기본 사용법

1. 원하는 보고서 유형의 코드를 복사
2. 필요에 따라 내용 및 이미지 경로 수정
3. 코드 실행하여 워드 문서 생성