# PPT Style HTML

Source: `ppt-style-html.md` at the project root, copied unchanged below. The block that follows says how the skill applies it. Where the two differ, the skill's starters follow the block.

- **Canvas**: each slide is 1280×720 px, the same 16:9 ratio as the 1920×1080 example (section 1 allows a logical size with the same ratio). The starters use the html-ppt slide format (`section.slide`).
- **Animation**: off by default. The starters have no motion library; the optional presenter (`?present`) only switches slides.
- **Charts**: ECharts (style 2) uses the SVG renderer, so its figures are vector shapes and text in the page and in PowerPoint. Style 1 draws its bars from elements.
- **Diagrams**: Mermaid (style 3) draws SVG labels at 24px; in PowerPoint the diagram is one vector image.
- **Presets**: Standard = style 1; Executive = style 1 with KPI cards and inline SVG figures; Data Story = style 2; Developer = style 3; Document = style 4 (Marp; not converted by html-ppt).
- **Colours**: hex or rgb only. Tailwind v4 palette classes compile to `oklch()`, which the converter drops, so the backgrounds, bars and icons disappear in PowerPoint.
- **Body text**: at least 24px (18pt) on the 1280×720 canvas (the bottom of the body range in section 11).
- **Conversion**: html-ppt converts styles 1–3 to native text and shapes, except the style 3 Mermaid flow, which becomes one vector picture. The renders were checked in PowerPoint (SKILL.md, "Verified conversion").

---

## Purpose

이 스킬의 최종 목적은 입력된 콘텐츠를 **PPT로 안정적으로 변환할 수 있는 HTML 프레젠테이션**으로 만드는 것이다.

HTML 자체에서 화려한 웹 프레젠테이션을 구현하는 것보다, 최종적으로 PowerPoint로 변환했을 때 레이아웃, 텍스트, 이미지, 표, 차트, 도형이 최대한 원형에 가깝게 유지되는 것을 우선한다.

따라서 다음 우선순위를 따른다.

1. PPT 변환 안정성
2. 레이아웃 재현성
3. 가독성
4. 발표 목적에 맞는 시각적 구성
5. 유지보수성
6. 시각적 화려함

기술적으로 구현 가능하더라도 PPT 변환 과정에서 손실될 가능성이 높은 효과는 기본적으로 사용하지 않는다.

핵심 원칙:

> **HTML을 최종 결과물로 보는 것이 아니라, PPT를 생성하기 위한 중간 프레젠테이션 포맷으로 사용한다.**

---

# 1. Default Presentation Settings

사용자가 별도의 요청을 하지 않는 경우 반드시 다음 기본값을 사용한다.

## Aspect Ratio

**16:9**

기본 프레젠테이션 비율:

```text
1920 × 1080
```

또는 이에 대응하는 동일 비율의 논리적 슬라이드 크기를 사용한다.

---

## Background

기본 배경:

**White (#FFFFFF)**

전체 프레젠테이션은 기본적으로 밝은 배경을 사용한다.

사용자가 별도로 다크 테마나 특정 배경을 요청한 경우에만 변경한다.

---

## Default Visual Direction

기본 스타일:

- White background
- Clean
- Professional
- Minimal
- Business-friendly
- PPT-friendly

기본적으로 과도한 장식이나 웹사이트 스타일 UI는 사용하지 않는다.

---

# 2. Core Goal

최종 HTML은 다음 조건을 만족해야 한다.

- PowerPoint 슬라이드로 변환하기 쉬워야 한다.
- 각 슬라이드의 위치와 크기가 명확해야 한다.
- 텍스트와 도형이 변환 후에도 최대한 동일하게 보여야 한다.
- 이미지와 차트가 잘리지 않아야 한다.
- 요소가 서로 겹치지 않아야 한다.
- 슬라이드마다 독립적인 레이아웃을 가져야 한다.
- 브라우저 크기에 따라 슬라이드 내부 요소가 임의로 재배치되지 않아야 한다.

웹페이지의 반응형 레이아웃보다는 **고정된 프레젠테이션 캔버스**를 우선한다.

---

# 3. Core Workflow

HTML 프레젠테이션 생성 시 다음 순서로 작업한다.

## Step 1. 발표 목적 분석

다음 항목을 분석한다.

- 발표 목적
- 발표 대상
- 핵심 메시지
- 콘텐츠 양
- 데이터 비중
- 이미지 필요성
- 차트 필요성
- 다이어그램 필요성

---

## Step 2. PPT 변환 가능성 판단

디자인을 결정하기 전에 다음 질문을 먼저 판단한다.

- PPT로 변환했을 때 레이아웃이 유지되는가?
- 애니메이션 없이도 메시지가 전달되는가?
- SVG, PNG, 텍스트, CSS 도형 등으로 안정적으로 표현 가능한가?
- 브라우저 전용 기능에 의존하고 있지 않은가?

PPT 변환이 어려워지는 디자인은 피한다.

---

# 4. Presentation Stack Presets

기본적으로 다음 프리셋을 사용한다.

기존 웹 프레젠테이션 스택보다 **PPT 변환 안정성**을 우선한다.

---

## 4.1 Standard

### Stack

HTML + Tailwind CSS

또는

Reveal.js + Tailwind CSS

### Priority

★★★★★

### Recommended For

- 일반 발표
- 사내 보고
- 교육 자료
- 프로젝트 보고
- 일반적인 PPT

### Characteristics

- 단순한 구조
- 높은 변환 안정성
- 빠른 제작
- 정적인 레이아웃
- 다양한 슬라이드 디자인 구현 가능

별도 요청이 없으면 가장 먼저 고려한다.

---

## 4.2 Executive

### Stack

HTML + Tailwind CSS + SVG

필요한 경우 ECharts를 제한적으로 사용한다.

### Recommended For

- 경영진 보고
- 사업 계획
- 전략 보고
- 컨설팅 스타일 발표
- 투자 보고

### Characteristics

- 숫자 강조
- 비교
- KPI
- 차트
- 전략 구조
- 적절한 여백

기본 배경은 흰색을 유지한다.

---

## 4.3 Data Story

### Stack

HTML + Tailwind CSS + ECharts

### Recommended For

- 데이터 분석
- KPI
- 시장 분석
- 실적 보고
- 통계 자료

### Rule

ECharts는 최종 PPT 변환 과정에서 문제가 발생할 가능성이 있는 경우 SVG 또는 PNG 형태로 고정하여 사용한다.

차트의 시각적 효과보다 **데이터 가독성과 변환 안정성**을 우선한다.

---

## 4.4 Developer

### Stack

HTML + Tailwind CSS + Mermaid

또는

Slidev 기반 정적 슬라이드

### Recommended For

- AI
- 개발
- 기술 발표
- 시스템 구조
- API
- 아키텍처

### Rule

Mermaid 다이어그램은 가능하면 최종적으로 SVG로 렌더링하여 사용한다.

코드 블록은 HTML 텍스트 형태로 안정적으로 표현한다.

---

## 4.5 Document

### Stack

Marp + Custom CSS + Mermaid

### Recommended For

- 보고서
- 문서 기반 발표
- 교육 자료
- 빠른 PPT 제작

### Characteristics

- 빠른 제작
- 단순한 구조
- 높은 변환 호환성
- Markdown 기반

---

# 5. Avoid Overly Complex Stacks

다음 기술은 기본 선택으로 사용하지 않는다.

- Framer Motion
- GSAP 기반 복잡한 애니메이션
- CSS 3D
- impress.js
- WebGL
- Canvas 기반 인터랙션
- 복잡한 React 상태 기반 UI

사용자가 명시적으로 요구하거나 반드시 필요한 경우에만 제한적으로 사용한다.

---

# 6. Animation Policy

PPT 변환이 최종 목적이므로 애니메이션은 기본적으로 사용하지 않는다.

기본값:

```text
Animation = OFF
```

HTML 미리보기용으로 애니메이션을 사용하는 경우에도 콘텐츠 이해를 방해하지 않는 수준으로 제한한다.

허용 가능한 효과:

- Fade
- Simple appear
- Minimal transition

피해야 할 효과:

- 3D Rotation
- Perspective
- Continuous animation
- Parallax
- Bounce
- Complex morphing
- Scroll-triggered animation

애니메이션이 없어져도 슬라이드 의미가 유지되어야 한다.

---

# 7. Layout Philosophy

## Fixed Slide Canvas

모든 슬라이드는 고정된 16:9 캔버스를 기준으로 설계한다.

예:

```css
.slide {
  width: 1920px;
  height: 1080px;
  background: #ffffff;
}
```

또는 동일한 비율의 기준값을 사용한다.

슬라이드 내부 레이아웃은 가능한 한 고정 위치 기반으로 구성한다.

---

# 8. Safe Area

중요 콘텐츠는 슬라이드 가장자리에 붙이지 않는다.

권장 Safe Area:

- 좌우 약 5%
- 상하 약 5%

제목, 본문, 차트, 이미지가 슬라이드 경계와 너무 가까워지지 않도록 한다.

---

# 9. Slide Design Principles

## One Slide One Message

가능한 한 한 슬라이드에는 하나의 핵심 메시지만 둔다.

---

## Visual Hierarchy

우선순위:

1. 제목 또는 핵심 메시지
2. 핵심 숫자
3. 이미지 / 차트
4. 설명
5. 보조 정보

---

## Text Density

긴 문단은 슬라이드에 그대로 넣지 않는다.

다음 형태로 변환한다.

- 핵심 문장
- Bullet
- KPI
- 비교
- Process
- Timeline
- Diagram
- Chart

---

# 10. Typography

PPT 변환 안정성을 위해 일반적으로 사용 가능한 폰트를 우선한다.

한글 기본 추천:

- Pretendard
- Noto Sans KR
- 맑은 고딕
- Arial
- sans-serif

특수 폰트 사용은 최소화한다.

사용자 환경에서 사용할 수 없는 폰트에 지나치게 의존하지 않는다.

---

# 11. Font Size

16:9 기준 권장 크기:

## Main Title

44~60pt 수준

## Slide Title

30~40pt 수준

## Body

18~26pt 수준

## Caption

14~18pt 수준

작은 텍스트를 과도하게 사용하지 않는다.

---

# 12. Recommended Slide Structures

PPT 변환에 안정적인 다음 구조를 우선한다.

## Title Slide

- 제목
- 부제
- 날짜
- 발표자

---

## Section Divider

- 큰 제목
- 짧은 설명

---

## Title + Content

가장 기본적인 PPT 구조.

---

## Two Column

좌우 비교 또는 이미지 + 설명.

---

## Three Column

3개의 핵심 요소 비교.

---

## KPI

큰 숫자 중심.

---

## Comparison

Before / After

A / B

---

## Timeline

프로젝트 일정 또는 시간 흐름.

---

## Process

3~6단계 프로세스.

---

## Chart

차트 + 핵심 인사이트.

---

## Diagram

시스템 또는 구조 설명.

---

## Full Image

큰 이미지 + 최소 텍스트.

---

# 13. Chart Rules

차트는 최대한 단순하게 만든다.

추천:

- Bar
- Line
- Area
- Donut
- Scatter
- Waterfall

피해야 할 것:

- 복잡한 3D 차트
- 너무 많은 범례
- 지나친 색상
- 지나치게 작은 글씨

가능하면 차트를 SVG로 렌더링하여 PPT 변환 안정성을 높인다.

---

# 14. Table Rules

표는 PowerPoint에서 읽기 쉬운 구조로 만든다.

- Header 강조
- 행 높이 충분히 확보
- 지나치게 많은 열 금지
- 지나치게 많은 행 금지
- 필요하면 여러 슬라이드로 분할

---

# 15. Image Rules

이미지는 다음을 우선한다.

- PNG
- JPEG
- SVG

가능하면 고해상도 이미지를 사용한다.

이미지가 슬라이드에서 비정상적으로 늘어나지 않도록 `object-fit`을 적절히 사용한다.

---

# 16. Color Rules

기본 배경:

```text
#FFFFFF
```

기본 텍스트:

```text
#111111
```

보조 텍스트:

```text
#555555
```

강조 색상은 일반적으로 1~2개만 사용한다.

사용자가 브랜드 컬러를 제공하면 해당 색상을 우선한다.

---

# 17. Default Theme

사용자가 다른 스타일을 요청하지 않는 경우 다음 디자인을 기본값으로 한다.

### Background

White

### Style

Modern Business

### Layout

Clean Grid

### Typography

Sans-serif

### Accent

1 Primary Accent Color

### Decoration

Minimal

### Animation

None

### Aspect Ratio

16:9

---

# 18. Avoid Web-App-Like Design

PPT가 최종 목적이므로 다음 UI는 기본적으로 사용하지 않는다.

- 웹사이트 Navigation Bar
- 모바일 앱 스타일 UI
- 버튼 중심 인터페이스
- Hover에 의존하는 정보
- Dropdown
- Modal
- Scroll interaction
- Web dashboard UI를 그대로 복제한 구조

필요한 경우 PPT 스타일로 재해석한다.

---

# 19. Content Transformation

입력된 콘텐츠를 그대로 복사하지 않는다.

발표용으로 재구성한다.

예:

Input:

"매출은 전년 대비 32% 증가했으며 신규 고객 증가와 해외 매출 증가가 주요 원인이다."

Output:

```text
32%
YoY Revenue Growth

신규 고객 확대와
해외 매출 성장이 성장 견인
```

---

# 20. Story Structure

가능하면 다음 흐름으로 구성한다.

Opening

↓

Context

↓

Problem

↓

Insight

↓

Evidence

↓

Solution

↓

Impact

↓

Next Step

↓

Closing

단순히 문서를 페이지 단위로 잘라 슬라이드로 만들지 않는다.

---

# 21. Technology Selection Rule

가능한 한 가장 단순한 기술로 구현한다.

우선순위:

```text
HTML/CSS
↓
Tailwind
↓
SVG
↓
ECharts / Mermaid
↓
JavaScript
↓
Advanced Framework
```

React, Vue, GSAP 등의 사용은 반드시 필요한 경우에만 선택한다.

---

# 22. PPT Conversion First Rule

HTML에서 구현 가능한 기능이라도 PowerPoint로 변환했을 때 손실될 가능성이 높으면 사용하지 않는다.

다음 질문을 항상 확인한다.

> 이 요소가 PPT로 변환되어도 동일한 의미와 시각적 구조를 유지하는가?

아니라면 더 단순한 표현 방식으로 변경한다.

---

# 23. Output Quality Checklist

완료 후 반드시 확인한다.

## Slide

- 16:9 비율인가?
- 별도 요청이 없다면 흰색 배경인가?
- 모든 요소가 Safe Area 안에 있는가?
- 요소가 겹치지 않는가?

## Content

- 제목이 명확한가?
- 슬라이드마다 핵심 메시지가 있는가?
- 텍스트가 지나치게 많지 않은가?

## Typography

- 본문 글씨가 너무 작지 않은가?
- 폰트가 일관적인가?
- PPT 변환 후에도 사용할 수 있는 폰트인가?

## Images

- 이미지 비율이 깨지지 않는가?
- 이미지 해상도가 충분한가?

## Charts

- 범례와 숫자가 읽히는가?
- 차트가 과도하게 복잡하지 않은가?

## Conversion

- 애니메이션 없이도 의미가 유지되는가?
- HTML을 정적 슬라이드로 렌더링할 수 있는가?
- PPT 변환 과정에서 사라질 요소가 없는가?

---

# 24. Default Behavior

사용자가 별도의 지시를 하지 않은 경우 자동으로 다음 설정을 적용한다.

```text
Aspect Ratio: 16:9
Background: White
Theme: Modern Business
Style: Clean / Minimal
Animation: None
Layout: Fixed Slide Canvas
Primary Goal: PPT Conversion Compatibility
```

프레젠테이션 스타일을 선택해야 할 경우:

1. Standard
2. Executive
3. Data Story
4. Developer
5. Document

순으로 콘텐츠에 적합한 스타일을 판단한다.

과도하게 화려한 웹 프레젠테이션은 기본값으로 선택하지 않는다.

---

# 25. Final Principle

이 스킬의 성공 기준은 HTML 자체가 얼마나 화려한지가 아니다.

최종적으로 PPT로 변환했을 때 다음을 만족하는지가 가장 중요하다.

- 깔끔하다.
- 읽기 쉽다.
- 발표하기 좋다.
- 레이아웃이 유지된다.
- 수정하기 쉽다.
- 일반적인 PowerPoint 슬라이드처럼 자연스럽다.

최종 판단 기준:

> **“이 HTML을 PPT로 변환했을 때 실제 업무나 발표에서 바로 사용할 수 있는가?”**

그렇지 않다면 디자인을 단순화하고 PPT 변환 안정성을 우선한다.