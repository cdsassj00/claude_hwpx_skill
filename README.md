# HWPX Writer

템플릿 HWPX 파일의 스타일, 레이아웃, 폰트를 100% 보존하면서 텍스트 내용만 교체하여 새 한글(.hwpx) 문서를 생성하는 AI 에이전트용 스킬입니다.

## 이런 분들에게 유용합니다

- 공공기관 교육 시간표, 보고서 등 **정해진 양식**에 내용만 바꿔서 반복 생산하는 경우
- AI 에이전트(Claude Code, Cursor, Cline 등)로 **한글 문서를 자동 생성**하고 싶은 경우
- HWPX 파일의 XML 구조를 직접 건드리지 않고 **안전하게 텍스트만 교체**하고 싶은 경우

## 작동 원리

HWPX는 ZIP 압축된 XML 문서입니다. 이 스킬은 원본 HWPX의 `section0.xml` 안에 있는 `<hp:t>` 태그(텍스트)만 위치 기반으로 교체합니다.

- linesegarray (줄 배치 정보) 보존
- 셀 크기, 마진, 테두리 보존
- charPrIDRef (폰트, 크기, 색상, 볼드) 보존
- paraPrIDRef (정렬, 줄간격, 들여쓰기) 보존
- 테이블 구조, 셀 병합, borderFill 보존

**즉, 텍스트 외에는 원본과 바이트 단위로 동일합니다.**

## 빠른 시작

### 1. 템플릿의 텍스트 위치 확인

```bash
python scripts/build_hwpx.py --dump-texts 양식.hwpx > text_positions.json
```

출력 예시:
```json
[
  {"index": 0, "text": "<붙임>"},
  {"index": 1, "text": "「AI시대 교육시간표」"},
  {"index": 2, "text": "□"},
  {"index": 3, "text": " 교육기간 : 2026. 3. 31.(5H) "},
  ...
]
```

### 2. 교체할 내용 작성 (content_spec.json)

```json
{
  "mode": "replace",
  "texts": [
    "<붙임>",
    "「사이버보안 실무자 양성」교육시간표",
    "□",
    " 교육기간 : 2026. 5. 12.~13.(10H) ",
    "□",
    " 교육인원 : 30명 내외 ",
    null,
    null
  ]
}
```

| 값 | 의미 |
|----|------|
| `"새 텍스트"` | 해당 위치의 텍스트를 교체 |
| `null` | 원본 텍스트 유지 |
| `""` | 텍스트 삭제 (빈 칸으로 만들기) |

리스트가 전체 위치 수보다 짧으면 나머지는 원본 유지됩니다.

### 3. HWPX 생성

```bash
python scripts/build_hwpx.py 양식.hwpx content_spec.json 결과물.hwpx
```

### 4. 확인

생성된 `결과물.hwpx`를 한컴오피스에서 열어서 확인합니다.

## AI 에이전트에서 사용하기

### Claude Code

`~/.claude/skills/` 에 이 레포를 클론합니다:

```bash
git clone https://github.com/cdsassj00/claude_hwpx_skill ~/.claude/skills/hwpx-writer
```

이후 Claude Code에서 HWPX 템플릿과 함께 내용 작성을 요청하면 자동으로 스킬이 활성화됩니다.

### 기타 AI 에이전트 (Cursor, Cline 등)

`SKILL.md` 파일을 에이전트의 지시문(rules/instructions)에 포함시키면 동일하게 동작합니다.

## 스크립트 설명

| 스크립트 | 용도 |
|---------|------|
| `scripts/build_hwpx.py` | HWPX 빌드 (텍스트 위치 덤프 + 교체 생성) |
| `scripts/analyze_hwpx.py` | 템플릿의 스타일 시스템 분석 (JSON 출력) |

### build_hwpx.py 사용법

```bash
# 텍스트 위치 확인
python scripts/build_hwpx.py --dump-texts 양식.hwpx

# HWPX 생성 (replace 모드 - 기본)
python scripts/build_hwpx.py 양식.hwpx content_spec.json 결과물.hwpx
```

### analyze_hwpx.py 사용법

```bash
# 템플릿의 전체 스타일 맵 추출
python scripts/analyze_hwpx.py 양식.hwpx -o style_map.json
```

출력되는 JSON에는 다음 정보가 포함됩니다:
- `charProperties`: 글자 속성 (폰트, 크기, 색상, 볼드 등)
- `paraProperties`: 문단 속성 (정렬, 줄간격, 들여쓰기 등)
- `borderFills`: 테두리/배경 스타일
- `documentStructure`: 실제 문서 내용과 스타일 참조 관계

## 요구사항

- Python 3.7+
- 추가 패키지 없음 (표준 라이브러리만 사용)

## 제한사항

- **텍스트 교체 전용**: 행/열 추가, 이미지 삽입 등 구조 변경은 불가
- **동일 구조 필요**: 새 내용의 텍스트 슬롯 수가 템플릿과 같아야 최적 결과
- **다중 런(run) 주의**: 한 문단에 여러 스타일이 섞인 경우 각 런이 별도 인덱스로 잡힘. 원본 런 수를 맞춰야 스타일 보존됨

## 참고 자료

- [references/hwpx-structure.md](references/hwpx-structure.md): HWPX 파일 포맷 상세 구조
- [SKILL.md](SKILL.md): AI 에이전트용 워크플로우 지시문
