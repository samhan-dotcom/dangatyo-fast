# ⚡ 단가표 빠른 추출기

통신사 단가표 이미지를 AI(Claude)로 자동 추출하여 **Excel / CSV / Markdown**으로 다운로드하는 도구입니다.

🔗 **배포 URL**: https://dangatyo-fast.streamlit.app  
📁 **GitHub**: https://github.com/samhan-dotcom/dangatyo-fast

---

## 목차

1. [왜 만들었나](#왜-만들었나)
2. [버전 구성](#버전-구성)
3. [주요 기능](#주요-기능)
4. [파일 구조](#파일-구조)
5. [기술 스택 및 설계 결정](#기술-스택-및-설계-결정)
6. [설치 및 로컬 실행](#설치-및-로컬-실행)
7. [Streamlit Cloud 배포](#streamlit-cloud-배포)
8. [비용 구조](#비용-구조)
9. [템플릿 추가 방법](#템플릿-추가-방법)
10. [향후 개선 방향](#향후-개선-방향)

---

## 왜 만들었나

통신사 단가표(LGU+, KT, SKT 등)는 이미지나 PDF 형태로 배포되어, 수작업으로 Excel에 옮기는 작업이 반복적으로 발생했습니다.  
Claude의 Vision 기능을 활용해 이 과정을 자동화하고, 팀원 누구나 링크 하나로 사용할 수 있게 만든 도구입니다.

---

## 버전 구성

총 두 가지 버전이 존재합니다.

| 구분 | 빠른 버전 (`app_fast.py`) | 정밀 버전 (`app.py`) |
|------|--------------------------|----------------------|
| Claude 호출 | 1회 | 2~4회 (병렬) |
| 소요 시간 | 약 10~20초 | 약 30~60초 |
| 비용 | 약 50~120원/장 | 약 100~250원/장 |
| OCR 교차검증 | ❌ | ✅ (PaddleOCR) |
| 재검증 | ❌ | ✅ |
| 감사(Audit) | ❌ | ✅ |
| 배포 방식 | Streamlit Cloud (무료) | 로컬 랩탑 전용 |
| 접속 방법 | 링크로 누구나 | Cloudflare Tunnel URL |

> **정밀 버전은 PaddleOCR 때문에 Python 3.12 + 로컬 설치가 필요**해 클라우드 배포가 불가합니다.  
> Vercel·Heroku는 WebSocket 기반 Streamlit과 호환되지 않아 Streamlit Cloud를 선택했습니다.

---

## 주요 기능

- 📤 이미지 업로드 (PNG, JPG, WEBP)
- 🤖 Claude AI로 표 자동 추출 (JSON Schema 강제 출력)
- 🟡 확신도 낮은 셀 노란색 강조 표시 (`uncertain_cells`)
- ✏️ 결과 인라인 편집 (셀 더블클릭)
- 📥 다운로드: Excel (.xlsx) / CSV / Markdown
- 🕐 시작·종료 시간 및 소요시간 표시
- 🔢 API 토큰 사용량 및 원화 비용 표시
- 💳 Anthropic Console 바로가기 (잔여 크레딧 확인)
- 📋 통신사별 양식 템플릿 (LGU+, KT, SKT, 자동감지)

---

## 파일 구조

```
단가표앱-cloud/          ← Streamlit Cloud 배포용 (이 레포)
├── app_fast.py          ← 메인 앱 (빠른 버전)
├── templates.py         ← 통신사별 양식 정의
├── requirements.txt     ← 패키지 목록
├── .gitignore           ← .env, secrets 제외
└── README.md            ← 이 파일

단가표앱/                ← 로컬 랩탑 전용 (별도 폴더)
├── app.py               ← 정밀 버전 (OCR 포함)
├── app_fast.py          ← 빠른 버전 (로컬용)
├── templates.py         ← 공유
├── run_all.bat          ← 두 버전 + Cloudflare Tunnel 동시 시작
└── SETUP.md             ← 로컬 설치 가이드
```

---

## 기술 스택 및 설계 결정

### 언어 / 프레임워크
- **Python + Streamlit**: 데이터 처리에 익숙한 팀을 위해 Python 선택. Streamlit은 UI 코드 없이 빠르게 웹앱 제작 가능.

### AI 모델
- **claude-sonnet-4-6**: 속도와 비용의 균형. Opus는 너무 비싸고, Haiku는 표 추출 정확도가 낮았음.
- **Extended Thinking (adaptive)**: 숫자 판독 정확도 향상을 위해 사용. `effort: "low"`로 비용 최소화.
- **JSON Schema 강제 출력**: 자유 텍스트 대신 구조화된 JSON으로 받아 파싱 오류 제거.

### 이미지 전처리
```python
# preprocess_image() 함수
- 긴 변 기준 1500~2576px로 리사이즈
- UnsharpMask로 선명도 향상
- 대비 +15% 조정
```
단순 원본 업로드 대비 OCR 정확도가 눈에 띄게 개선되었음.

### 다운로드 캐싱
`st.session_state`에 xlsx/csv/md를 저장해 두고, 편집 내용이 바뀔 때만 재빌드합니다 (`build_key` 비교).  
이렇게 하지 않으면 Streamlit 특성상 버튼 클릭마다 파일을 재생성해 느려집니다.

### 비용 추적
```python
# _add_tokens(usage) 함수
# 세션 내 모든 API 호출의 토큰을 누산
st.session_state.usage_tokens = {"input": 0, "output": 0, "cache_read": 0, "cache_create": 0}
```
1달러 = 1,400원으로 환산해 원화 표시. 캐시 읽기(0.1x)·쓰기(1.25x) 할인도 반영.

### API 키 관리
로컬과 클라우드 양쪽을 지원하도록 우선순위를 두었습니다:
```python
def get_client():
    api_key = os.environ.get("ANTHROPIC_API_KEY")   # 1순위: .env
    if not api_key:
        api_key = st.secrets.get("ANTHROPIC_API_KEY")  # 2순위: Streamlit Secrets
```

---

## 설치 및 로컬 실행

### 사전 조건
- Python 3.12 (PaddleOCR 호환, 3.13+ 불가)
- `cloudflared` 설치 ([다운로드](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/))

### 설치
```bash
# 가상환경 생성 (반드시 Python 3.12로)
py -3.12 -m venv .venv
.venv\Scripts\activate

# 패키지 설치
pip install -r requirements.txt

# OCR 설치 (정밀 버전 사용 시)
pip install paddlepaddle paddleocr
```

### 환경변수 설정
`.env` 파일 생성:
```
ANTHROPIC_API_KEY=sk-ant-api03-...
```

### 실행
```bash
# 빠른 버전만
streamlit run app_fast.py --server.port 8502

# 두 버전 + Cloudflare 동시 시작 (Windows)
run_all.bat
```

`run_all.bat`은 아래를 자동 실행합니다:
1. 기존 프로세스 종료
2. 정밀 버전 (포트 8501) 시작
3. 빠른 버전 (포트 8502) 시작
4. Cloudflare Tunnel 2개 시작 → URL 출력

> **주의**: Cloudflare Tunnel은 `http://localhost` 대신 `http://127.0.0.1`을 사용해야 합니다.  
> `localhost`를 쓰면 IPv6(`[::1]`)로 해석되어 Streamlit(IPv4)와 연결 실패합니다.

---

## Streamlit Cloud 배포

### 구조
이 레포(`dangatyo-fast`) 자체가 Streamlit Cloud에 직접 연결되어 있습니다.  
`main` 브랜치에 push하면 자동으로 재배포됩니다.

### 코드 수정 → 배포 워크플로우
```bash
# 수정 후
git add .
git commit -m "변경 내용 설명"
git push
# → Streamlit Cloud가 자동으로 재빌드 (약 1~2분)
```

### API 키 설정
Streamlit Cloud → 앱 설정 → Secrets:
```toml
ANTHROPIC_API_KEY = "sk-ant-api03-..."
```

### 절전 모드
무료 플랜은 일정 시간 미접속 시 앱이 절전 상태가 됩니다.  
접속하면 "Zzzz" 화면이 뜨고, 버튼 클릭 후 30~60초면 재시작됩니다.  
**팀원에게 이 점을 미리 안내하세요.**

---

## 비용 구조

| 항목 | 단가 |
|------|------|
| claude-sonnet-4-6 입력 | $3.00 / 100만 토큰 |
| claude-sonnet-4-6 출력 | $15.00 / 100만 토큰 |
| 캐시 쓰기 | 입력의 1.25배 |
| 캐시 읽기 | 입력의 0.10배 |
| **장당 실비용 (빠른)** | **약 50~120원** |
| **장당 실비용 (정밀)** | **약 100~250원** |

잔여 크레딧 확인: https://console.anthropic.com/settings/billing

---

## 템플릿 추가 방법

새 통신사나 새 양식을 추가하려면 `templates.py`만 수정하면 됩니다.

```python
# templates.py 예시 구조
TEMPLATES = {
    "auto": {
        "display_name": "⊕ 자동 감지 (양식 모를 때)",
        "prompt": None,
    },
    "lgup": {
        "display_name": "📱 LGU+",
        "prompt": "...(LGU+ 전용 추출 지시사항)...",
    },
    # 새 양식 추가 예시:
    "mvno": {
        "display_name": "📡 알뜰폰 (MVNO)",
        "prompt": """
        이 이미지는 알뜰폰 단가표입니다.
        컬럼: 통신사, 요금제명, 월정액, 데이터, 특이사항
        ...(구체적인 추출 지시)...
        """,
    },
}
```

추가 후 `git push`하면 클라우드에 자동 반영됩니다.

---

## 향후 개선 방향

현재 알려진 한계와 개선 포인트입니다.

### 단기
- [ ] 여러 이미지 일괄 업로드 (현재 1장씩)
- [ ] 추출 결과 수동 저장/불러오기 (히스토리)
- [ ] 다크 모드 지원

### 중기  
- [ ] PDF 직접 업로드 (현재 스크린샷만 가능)
- [ ] 표 비교 기능 (이번달 vs 지난달 단가 차이 강조)
- [ ] 팀 공유용 결과 링크 생성

### 장기
- [ ] 자동화: 이메일/카카오톡으로 받은 단가표 자동 처리
- [ ] 데이터베이스 연동 (단가표 이력 관리)
- [ ] 단가 변동 알림

### AI 모델 관련
- 현재 `claude-sonnet-4-6` 사용 중
- 모델 업그레이드 시 `app_fast.py` 상단의 `MODEL` 변수만 변경하면 됩니다:
  ```python
  MODEL = "claude-sonnet-4-6"  # 이 줄만 수정
  ```
- `PRICES` 딕셔너리도 함께 업데이트 필요 (비용 계산용)

---

## 관련 링크

- [Anthropic Console](https://console.anthropic.com) — API 키 관리, 크레딧 확인
- [Streamlit Cloud](https://share.streamlit.io) — 배포 상태 확인, 재시작
- [Claude API 문서](https://docs.anthropic.com) — 모델/가격 정보
- [Streamlit 문서](https://docs.streamlit.io) — UI 컴포넌트 참고
