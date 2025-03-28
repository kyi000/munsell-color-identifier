# 먼셀 색상 자동 인식 프로그램

이 프로그램은 문화재 발굴조사를 위해 개발된 도구로, 사진이나 화면에서 마우스 커서가 위치한 색상의 먼셀 기호를 실시간으로 표시합니다.

## 주요 기능
- 화면에서 마우스 커서 위치의 색상을 실시간 감지
- 감지된 색상에 가장 가까운 먼셀 색상 코드 자동 표시
- 이미지 파일 불러오기 기능 지원
- 색상 히스토리 저장 및 내보내기 기능

## 설치 방법
1. Python 3.7 이상이 설치되어 있어야 합니다.
2. 다음 명령어로 필요한 라이브러리를 설치합니다:

```bash
pip install -r requirements.txt
```

## 사용 방법
1. 다음 명령어로 프로그램을 실행합니다:

```bash
python munsell_identifier.py
```

2. 프로그램이 실행되면 '파일 열기' 버튼을 클릭하여 발굴 현장에서 촬영한 이미지를 불러올 수 있습니다.
3. 이미지 위에 마우스를 움직이면 커서 옆에 해당 위치 색상의 먼셀 기호가 실시간으로 표시됩니다.
4. 관심 있는 색상 위치에서 클릭하면 해당 색상과 먼셀 기호가 히스토리에 저장됩니다.
5. '내보내기' 버튼을 통해 저장된 색상 정보를 CSV 파일로 저장할 수 있습니다.

## 개발자 정보
이 프로그램은 문화재 발굴조사 전문가들의 작업 효율성을 높이기 위해 개발되었습니다.
