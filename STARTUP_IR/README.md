# STARTUP_IR | Reusable IR Deck Template

기업별로 반복 활용할 수 있는 **16:9, 화이트·딥그린 기반 투자유치 IR 덱** 템플릿입니다. 인피니트리 IR 덱의 시각적 방향(결론형 제목, 녹색 강조, 이미지·차트 중심, 최소 텍스트)과 공유된 IR 강연 자료의 장표별 구조를 참고해 **기업 정보가 없는 범용 템플릿**으로 재구성했습니다.

## 시작하기

1. `templates/startup_ir_green_editable_12slides.pptx`를 PowerPoint에서 엽니다.
2. `[대괄호]`를 기업별 검증된 정보로 교체합니다.
3. 회사 로고, 제품 사진, 고객 사례, 검증된 시장 데이터, 출처를 삽입합니다.
4. `docs/QUALITY_CHECKLIST.md`에 따라 검증한 후 투자자에게 공유합니다.

모든 텍스트와 도형은 PPT에서 편집 가능하며, 시각적 예시용 도형은 기업의 실제 제품 사진이나 데이터 시각화로 교체할 수 있습니다. Pretendard 폰트가 설치되지 않은 PC에서는 대체 폰트로 표시될 수 있습니다.

## 슬라이드 구성 (12장)

| # | 장표 | 핵심 질문 |
|---|---|---|
| 01 | Cover | 어떤 기업이며 어떤 가치를 제공하는가? |
| 02 | Problem | 고객의 실질적 문제는 무엇인가? |
| 03 | Solution | 기존 방식보다 무엇이 나은가? |
| 04 | Traction | 무엇이 검증되었는가? |
| 05 | Market Size | 실제로 접근 가능한 시장은 얼마나 큰가? |
| 06 | Business Model | 누가 무엇에 돈을 지불하는가? |
| 07 | Go-to-Market | 어떻게 고객을 확보하고 전환하는가? |
| 08 | Competition | 어떤 구매 기준에서 경쟁 우위가 있는가? |
| 09 | Team | 왜 이 팀이 실행할 수 있는가? |
| 10 | Financial Plan | 매출·이익·현금흐름이 어떻게 성장하는가? |
| 11 | Scaleup / EXIT | 성장 및 회수 마일스톤은 무엇인가? |
| 12 | Use of Funds | 투자금을 어디에 쓰고 무엇을 달성하는가? |

## 재생성

```bash
pip install python-pptx
python scripts/build_template.py
```

위 명령은 `templates/startup_ir_green_editable_12slides.pptx`를 다시 생성합니다. PPTX를 직접 수정한 내용은 스크립트를 재실행하면 덮어써질 수 있으므로 기업별 결과물은 별도 폴더에 저장하세요.

## 기업별 프로젝트 운영 권장 구조

```
clients/
  company-name/
    sources/       # 원본 자료: 반드시 비공개 저장소 사용
    facts.md       # 수치, 증빙, 확인 상태
    deck.pptx      # 기업별 편집본
    exports/       # PDF / PNG
```

**보안:** 기업 IR, 계약서, 투자검토보고서, 개인정보, 고객사 실명 및 비공개 매출은 공개 GitHub 저장소에 올리지 마세요. 이 템플릿 저장소는 범용 디자인 요소만 포함합니다.

## 상태와 검증

- `[ ]`는 미입력 항목이며 **실제 기업의 확정 정보가 아닙니다**.
- 매출·영업이익·시장규모·효율개선율은 출처·기준연도·산식을 함께 기록합니다.
- 실제 실적, 시뮬레이션, 목표, 계획, 협의 중인 사항을 구분합니다.
- 이미지·로고는 사용 권한을 확인합니다.

자세한 규칙은 `docs/DESIGN_GUIDE.md`, `docs/SLIDE_CONTENT_GUIDE.md`, `docs/QUALITY_CHECKLIST.md` 참고.


Repository: https://github.com/Navion78/STARTUP_IR
