# 📄 AI 보고서 자동 생성기

Python을 활용하여 Word 문서를 자동으로 생성하는 도구 모음입니다.

## 🚀 주요 기능

- **자동 Word 문서 생성**: Python을 통한 전문적인 보고서 작성
- **LLM 통합**: OpenAI API와 LM Studio를 활용한 지능형 문서 생성
- **템플릿 시스템**: 다양한 보고서 양식을 미리 정의하여 재사용 가능
- **학습 결과 시각화**: 머신러닝 모델의 학습 곡선과 성능 지표 자동 포함

## 📋 포함된 도구들

### 1. `word_with_summary.py`
- CNN 학습 결과를 종합한 상세 보고서 생성
- 요약(Executive Summary) 섹션 포함
- 학습 곡선 및 Confusion Matrix 자동 삽입
- PDF 변환 기능 내장

### 2. `make_word.py`
- 기본적인 Word 문서 생성 도구
- 간단한 보고서 템플릿 제공

### 3. `generate_training_curves.py`
- 머신러닝 학습 결과 시각화
- 정확도 및 손실 곡선 자동 생성

## 🛠️ 설치 및 사용법

### 필요한 라이브러리 설치
```bash
pip install python-docx docx2pdf langchain-openai python-dotenv
```

### 환경 변수 설정
`.env` 파일을 생성하고 다음 내용을 추가하세요:
```
LMSTUDIO=http://localhost:1234/v1
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-3.5-turbo
```

### 실행 방법
```bash
python word_with_summary.py
```

## 📁 프로젝트 구조

```
make_report/
├── word_with_summary.py    # 메인 보고서 생성기
├── make_word.py           # 기본 Word 생성 도구
├── generate_training_curves.py  # 학습 곡선 생성
├── templates/             # 보고서 템플릿들
├── Model/                 # 생성된 보고서 저장 폴더
└── README.md             # 이 파일
```

## 🎯 사용 사례

- **연구 보고서**: 머신러닝/딥러닝 실험 결과 문서화
- **프로젝트 리포트**: 기술 프로젝트 성과 보고서
- **학습 로그 분석**: 모델 학습 과정 시각화 및 문서화
- **업무 보고서**: 정형화된 비즈니스 문서 생성

## 📊 생성되는 보고서 예시

생성된 보고서에는 다음이 포함됩니다:
- 📈 학습 정확도 및 손실 곡선
- 🎯 Confusion Matrix 시각화
- 📋 성능 지표 요약 테이블
- 📝 상세한 실행 요약(Executive Summary)
- 📊 에포크별 학습 진행 상황

## 🤝 기여하기

이 프로젝트는 지속적으로 개선되고 있습니다. 새로운 템플릿이나 기능 추가에 대한 아이디어가 있으시면 언제든 기여해주세요!

## 📄 라이선스

이 프로젝트는 MIT 라이선스를 따릅니다.
