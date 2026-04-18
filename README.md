# PPT → PDF 변환기 (폰트 깨짐 방지, 1280×720)

각 슬라이드를 PowerPoint로 직접 렌더링해 **1280 픽셀 가로 고정** PNG로 내보낸 뒤,
img2pdf로 무손실 결합해 PDF를 만든다. 이 방식 덕분에 **글씨·폰트가 절대 깨지지 않는다**.

> 트레이드오프: 결과물은 이미지 기반 PDF이므로 텍스트 검색/복사가 불가능하다.

## 빠른 시작 (Windows 사용자용 — Python 불필요)

1. 이 저장소의 [**Releases**](../../releases/latest) 페이지로 이동
2. 최신 릴리스의 **`ppt2pdf.exe`** 다운로드
3. 더블클릭하여 실행 → PPT 파일을 드래그앤드롭 → **변환 시작**

> **Windows SmartScreen 경고가 뜨는 경우**
> 서명되지 않은 exe라 "Windows에서 PC를 보호했습니다" 안내가 표시될 수 있다.
> **추가 정보 → 실행** 을 클릭하면 정상 실행된다.

> **PowerPoint는 반드시 설치되어 있어야 한다.** 슬라이드 렌더링을 실제 PowerPoint에 위임하기 때문에, 설치된 폰트가 그대로 사용되어 글씨 깨짐이 발생하지 않는다.

## 필요 환경

- Windows 10 이상
- Microsoft PowerPoint 설치 (슬라이드 렌더링에 사용)
- Python 3.10+ (개발용). 빌드된 `.exe`를 쓰면 Python 불필요.

## 설치

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 실행

### GUI
```bat
python app.py
```

1. PPT 파일을 창에 드래그앤드롭 (또는 "찾아보기")
2. 필요하면 "용량 최소화 (JPEG 재압축)" 체크
3. **변환 시작** 클릭
4. 완료 후 PDF 용량이 표시되며 "폴더 열기"로 결과 확인

### CLI
```bat
python converter.py input.pptx output.pdf
python converter.py input.pptx output.pdf --jpeg
```

## 출력 사양

| 입력 PPT 비율 | 출력 이미지 크기 |
| ------------- | ---------------- |
| 16:9          | 1280 × 720       |
| 4:3           | 1280 × 960       |

- 30 슬라이드 기준 일반적으로 **5~8MB** 수준.
- 10MB 초과 시 GUI가 JPEG 재압축을 제안한다.

## 배포용 단일 exe 빌드

### 방법 1: GitHub Actions 자동 빌드 (권장)

저장소에 포함된 `.github/workflows/build.yml` 이 Windows 러너에서 자동으로 exe를 만든다. 로컬에 Windows가 없어도 된다.

| 트리거 | 결과물 | 위치 |
|--------|--------|------|
| `git tag v1.0.0 && git push --tags` | **GitHub Release + ppt2pdf.exe 첨부** | Releases 페이지 |
| Actions 탭 → "Run workflow" (수동) | 아티팩트 업로드 (30일 보관) | Actions 실행 상세 |
| `main` 브랜치 push | 아티팩트 업로드 (30일 보관) | Actions 실행 상세 |

**공식 배포 절차**:
```bash
git tag v1.0.0
git push --tags
# 3~5분 뒤 Releases 페이지에 ppt2pdf.exe 가 자동 게시됨
```

### 방법 2: Windows PC 로컬 빌드

```bat
pip install pyinstaller
build.bat
```

결과: `dist\ppt2pdf.exe` (단일 실행 파일)

## 문제 해결

- **"PowerPoint.Application" 디스패치 실패**: PowerPoint가 설치되어 있지 않거나
  COM 등록이 깨졌을 수 있다. PowerPoint를 먼저 한번 실행해 정상 기동되는지 확인.
- **한글 경로에서 실패**: 거의 없으나, 문제 시 영문 경로로 옮겨 재시도.
- **변환 중 PowerPoint 창이 잠깐 뜸**: COM 제약으로 완전한 hidden이 불가능한
  버전이 있다. 변환은 정상 진행된다.
