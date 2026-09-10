# harness

AI 에이전트(Claude Code · Aside · Codex 등)가 일할 때 따르는 규칙 저장소.
**프로젝트마다 규칙이 다르므로, 작업하는 프로젝트의 폴더만 읽는다.** 다른 프로젝트 규칙을 섞지 않는다.

```
AGENTS.md                         공통 — 레인 배치 · Effort · 작업 시작 루틴 (모든 프로젝트)
projects/
  batmaego/                       밭매고 · 밭매고맵 (농지 배치 시뮬레이션 서비스)
    AGENTS.md  CLAUDE.md  PROMPT.md
  thesis-seolhyang/               설향 딸기 영농창업계획 학위논문
    AGENTS.md  CLAUDE.md  PROMPT.md
    PPT_디자인스펙.md              16:9 팔레트 · 장 골격 좌표 · AI 티 빼기 체크리스트
    scripts/ppt_charts.py         차트 + KPI 카드 + 각주 헬퍼 (python-pptx)
    scripts/ppt_diet.py           글자 많은 슬라이드 → 메시지 + 표 구조
    templates/근거기록_템플릿.md   수치 출처 · 이유 · 변경이력
    templates/논문체제_번호대장_템플릿.md
```

## 사용법

1. 새 세션을 열고 **작업할 프로젝트의 `PROMPT.md`** 내용을 통째로 붙여넣는다.
2. 에이전트가 루트 `AGENTS.md` + 해당 프로젝트 `AGENTS.md`를 읽고, 작업마다 "목표 · Effort · 위임 계획" 세 줄을 먼저 제시한다.
3. 확인 후 진행.

| 프로젝트 | 붙여넣을 파일 |
|---|---|
| 밭매고 · 밭매고맵 | `projects/batmaego/PROMPT.md` |
| 설향 딸기 논문 | `projects/thesis-seolhyang/PROMPT.md` |

## 공통 원칙 요약 (루트 AGENTS.md)

- 메인 모델은 **설계 · 분해 · 검토 · 최종 판단**만. 구현 · 작성 · 반복 작업은 서브에이전트(Sonnet 기본, 무거운 건 Opus)에 위임. 5분 이내 단순 작업은 직접.
- Effort 기본 medium, 단순 low. high/xhigh는 사용자 승인 시에만.
- 작업 착수 전 세 줄: ① 목표 · 완료 기준 ② 권장 Effort · 이유 ③ 위임 계획.

## Aside에서의 적용

- 루트 `AGENTS.md` → `~/.aside/u/0/AGENTS.md` (계정 공통, 항상 로드)
- `projects/batmaego/AGENTS.md` → `~/.aside/u/0/skills/user/batmaego/SKILL.md` (밭매고 작업 시 자동 로드)
- `projects/thesis-seolhyang/AGENTS.md` → `~/Desktop/논문 준비 자료/.agents/skills/thesis-seolhyang/SKILL.md` (논문 워크스페이스 전용)
