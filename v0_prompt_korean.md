# v0.app 프롬프트 (한글 메뉴 + 반응형)

Design a modern industrial CNC machine monitoring dashboard UI.

Context:
- This is for a factory environment monitoring multiple FANUC CNC machines.
- The current app is a dense WinForms desktop app. It shows too many grids and buttons on one screen and feels old and cluttered.
- I want a clean, modern dashboard layout that I can later implement with WPF or WinUI 3 in C#.
- All menu items and labels should be in Korean (한글).
- The design should be responsive for desktop window resizing (not mobile web, but flexible desktop layout).
- **IMPORTANT**: The core features are manual offset adjustment and coordinate/macro data correction. Real-time monitoring is a secondary/auxiliary feature.

Requirements:
- Overall style: dark / dark-gray industrial theme with accent colors (green for normal, yellow for warning, red for alarm).
- Top header:
  - App title: "CNC 제어 센터"
  - Selected machine info (기계명, IP 주소, 연결 상태)
  - Quick action buttons (연결, 새로고침)
- Left sidebar:
  - **Primary navigation (main features):**
    - "수동 옵셋" (Manual Offset) - **CORE FEATURE**
    - "좌표계 보정" (Coordinate System Correction) - **CORE FEATURE**
    - "매크로 데이터 보정" (Macro Data Correction) - **CORE FEATURE**
  - **Secondary navigation (auxiliary features):**
    - "실시간 모니터링" (Real-time Monitoring)
    - "다축 모니터링" (Multi-Machine Monitoring)
    - "알람" (Alarms)
    - "로그" (Logs)
    - "설정" (Settings)
  - Machine selector dropdown at bottom (기계 선택)

- **Main "수동 옵셋" page layout (PRIMARY FEATURE):**
  - Top section: Selected machine info and tool selector
  - Main content area:
    - Left: Tool list table/grid
      - Columns: 공구번호, X축 마모, Z축 마모, 상태
      - Click to select tool
    - Center: Selected tool detail
      - Current wear values display (X축, Z축)
      - Input fields for manual adjustment (수동 보정값 입력)
      - Apply button (적용)
      - History log below (최근 보정 이력)
    - Right: PMC control buttons (Block Skip, Optional Stop)
  - Bottom: Operation log table (시간, 공구번호, 축, 입력값, 변경 전, 변경 후, 성공 여부)

- **Main "좌표계 보정" page layout (PRIMARY FEATURE):**
  - Top: Coordinate system selector (EXT, G54, G55, G56, G57, G58, G59)
  - Main content:
    - Left: Current coordinate values table
      - Columns: 좌표계, X축, Y축, Z축, C축
      - Read all button (전체 읽기)
    - Center: Correction input area
      - Input fields for each axis (X, Y, Z, C)
      - Range validation (-0.3 ~ +0.3)
      - Apply button (적용)
      - Clear button (초기화)
    - Right: Operation history
      - Scrollable log of coordinate changes
  - Bottom: Quick actions (저장, 불러오기)

- **Main "매크로 데이터 보정" page layout (PRIMARY FEATURE):**
  - Top: Macro variable selector (매크로 번호 선택)
  - Main content:
    - Left: Macro variable list/grid
      - Columns: 매크로 번호, 현재 값, 목표 값, 편차
      - Search/filter box
    - Center: Selected macro detail
      - Current value display
      - Target value input
      - Correction value input
      - Apply button (적용)
      - Sync interval setting (동기화 간격: 초 단위)
    - Right: Macro change history
      - Time, macro number, old value, new value, success status
  - Bottom: Batch operations (일괄 읽기, 일괄 쓰기, 동기화 시작/중지)

- **Secondary "실시간 모니터링" page layout (AUXILIARY FEATURE):**
  - Left panel: Machine list with status
  - Center panel: Selected machine detail
    - Status card (가동 상태, 프로그램명, 가공 수)
    - Tabs: "개요", "프로그램", "옵셋", "매크로", "PMC"
    - Key metrics cards (스핀들 부하, 이송 속도 등)
  - Right panel: Recent alarms/events

- **Secondary "다축 모니터링" page (AUXILIARY FEATURE):**
  - Grid of machine cards showing status, utilization, small charts

- **Secondary "로그" page (AUXILIARY FEATURE):**
  - Table with columns: 시간, 기계, 카테고리 (옵셋, 매크로, 수동 조정, 알람), 메시지
  - Filters and search bar

Design goals:
- **Primary focus**: Manual offset, coordinate system correction, and macro data correction pages should be the most prominent and easily accessible.
- Very clear visual hierarchy: correction input fields and apply buttons must be immediately visible and easy to use.
- Minimize visible borders; use cards, shadows, and spacing instead.
- Use consistent iconography for states (check, warning, error).
- Make it feel like a professional, modern industrial CNC control interface (not just monitoring).
- Responsive layout: panels should resize smoothly when window size changes (flexible grid/flexbox layout).
- All text labels, menu items, and UI elements should be in Korean.
- Input fields for corrections should be large, clear, and have proper validation feedback.
- Operation history/logs should be easily accessible but not overwhelming the main correction interface.

Output:
- Provide the page layout structure (sections, components) and styling hints.
- Use Tailwind or CSS utility-like descriptions so that the layout is clear.
- Focus on UX and visual hierarchy, especially for the three primary correction features.
- Ensure the layout is responsive for desktop window resizing scenarios.
- Emphasize the correction workflow: select machine → select tool/coordinate/macro → input correction value → apply → view result.
