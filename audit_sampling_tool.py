# -*- coding: utf-8 -*-
"""
표본감사(Audit Sampling) 통합 Template
=================================================================
전체 구성 (좌측 메뉴 4개 탭)
    1. Tool 사용가이드     <- 프로그램 내장 사용법 안내
    2. 표본감사유형선택    <- 기본정보·모집단·수행중요성 입력, 통계적/비통계적 선택
    3. 통계적표본감사      <- 위험요소 입력 → 표본크기 산정 → MUS 방식 결과평가 → 최종결론
    4. 비통계적표본감사    <- 위험요소 입력 → 표본크기 산정 → 비례식 결과평가 → 최종결론

3·4번 탭은 2단계에서 통계적/비통계적 중 하나를 선택해야 활성화되며, 감사위험모형
(TD = AR ÷ (IR×CR×AP1×AP2)) 계산과 표본크기 산정 로직을 공유합니다.

우측 하단 Reset 버튼 (2단계 탭)
    - 확인 팝업 후 2단계(기본정보 등) 입력값만 초기화 + 3·4단계 선택 잠금
    - 3·4단계의 「결과 엑셀 저장」에 2단계 정보가 함께 포함되므로 별도 Save 버튼은 없음
    - 다른 탭으로 이동해도 Reset을 누르거나 프로그램을 종료하지 않는 한 입력값은
      그대로 유지됨 (프레임을 파괴하지 않고 tkraise()로만 화면 전환하기 때문)

자세한 사용법은 프로그램 내 "1. Tool 사용가이드" 탭 또는 README.md 참고.

빌드 방법 (PyInstaller)
-----------------------------------------------
    pip install pyinstaller
    pyinstaller --onefile --windowed --name "AuditSamplingTool" audit_sampling_tool.py
"""

import json
import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

from gamma_pure import poisson_sample, reliability_factor

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ----------------------------------------------------------------------
# 감사위험모형(Audit Risk Model) 계수표
#   TD(설계 표본추출 발견위험) = AR ÷ (IR × CR × AP)
#   설계 신뢰수준(Confidence Level) = 1 − TD
#   * AP는 분석절차계수와 기타절차계수를 통합한 개념으로 사용
#   * 근거: AICPA Audit Sampling Guide의 위험모형 프레임워크(Table 4류)를
#     기반으로 자체 캘리브레이션한 값이며, 추후 조정 가능
# ----------------------------------------------------------------------
AR_BY_LISTED = {
    "상장": 0.05,
    "기타": 0.10,
}

IR_FACTORS = {
    "Other (Material Only / Applicable Assertion / 기타)": 0.2,
    "낮음": 0.5,
    "높음": 1.0,
}

CR_FACTORS = {
    "None": 1.0,
    "Limited": 0.5,
    "Normal": 0.15,
    "Extended": 0.1,
}

# AP1 = 분석절차계수 (기존 "분석·기타절차계수(AP)"의 옵션/값을 그대로 유지)
AP1_FACTORS = {
    "None": 1.0,
    "Less extensive": 0.7,
    "More extensive": 0.4,
}

# AP2 = 기타세부절차계수 (신규 항목, None / Less extensive만 지원)
AP2_FACTORS = {
    "None": 1.0,
    "Less extensive": 0.7,
}

# 표본선택방법 옵션 (통계적/비통계적 공용, 통계적은 임의(Haphazard) 제외)
SELECTION_METHODS_STATISTICAL = ["체계적 (Systematic)"]
SELECTION_METHODS_NONSTATISTICAL = ["체계적 (Systematic)", "무작위 (Random, IDEA only)", "임의 (Haphazard)"]

NONSTAT_PENALTY_FACTOR = 1.5  # 계층화 X + 무작위/임의 선택 시 표본크기 보정계수
MIN_SAMPLE_SIZE_FLOOR = 5     # 최소 표본크기(모든 계산 공통 하한, 잔여모집단 건수 이하로 제한)


def setup_ttk_styles():
    """ttk.Radiobutton 스타일 설정.

    Windows에서 classic tk.Radiobutton에 커스텀 bg를 지정하면 선택 여부와 무관하게
    인디케이터가 항상 채워진 것처럼 보이는 알려진 렌더링 버그가 있다. ttk.Radiobutton은
    테마 기반 렌더링이라 이 문제가 없으므로, 앱 전역에서 ttk.Radiobutton + 아래 스타일만
    사용한다.
    """
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "App.TRadiobutton",
        background=COLOR_SECTION_BODY,
        foreground=COLOR_TEXT_DARK,
        font=(FONT_NAME, 10),
    )
    style.map(
        "App.TRadiobutton",
        background=[("active", COLOR_SECTION_BODY)],
        foreground=[("active", COLOR_TEXT_DARK)],
    )


def make_radio(parent, text, value, variable, bold=False):
    """섹션 본문 배경(COLOR_SECTION_BODY) 위에 올라가는 표준 라디오버튼."""
    style_name = "App.TRadiobutton"
    rb = ttk.Radiobutton(parent, text=text, value=value, variable=variable, style=style_name)
    return rb

# ----------------------------------------------------------------------
# 색상 팔레트 (#16365C 남색을 Base로 하고, 포인트/강조색은 회색 톤 사용)
# ----------------------------------------------------------------------
COLOR_BG_MAIN = "#16365C"          # 남색 Base (바깥 배경)
COLOR_HEADER_BAR = "#1D4570"       # 최상단 바 배경 (Base보다 살짝 밝은 남색)
COLOR_ACCENT = "#5C6670"           # 포인트 회색 (테두리 등, 기존 파란 포인트 대체)

COLOR_SIDEBAR_BG = "#0E233C"       # 좌측 사이드바 배경 (Base보다 어둡게)
COLOR_SIDEBAR_ITEM = "#2D4A6C"     # 사이드바 비활성 메뉴 항목 (Base보다 살짝 밝게)
COLOR_SIDEBAR_ITEM_ACTIVE = "#5C6670"   # 사이드바 현재 선택된 메뉴 항목 (회색 포인트)
COLOR_SIDEBAR_ITEM_DISABLED = "#16365C"  # 사이드바 잠긴(비활성) 메뉴 항목 (Base와 동일)
COLOR_SIDEBAR_TEXT = "#E8EEF5"
COLOR_SIDEBAR_TEXT_DISABLED = "#5A6B7D"

COLOR_SECTION_BODY = "#E9ECEF"     # 섹션 본문 배경 - 옅은 회색
COLOR_SECTION_HEADER = "#B9C0C9"   # 섹션 제목줄 배경 - 회색

COLOR_TEXT_LIGHT = "#F2F6FB"
COLOR_TEXT_DARK = "#1B2733"
COLOR_TEXT_MUTED = "#4B5563"

COLOR_ENTRY_BG = "#FFFFFF"
COLOR_ENTRY_FG = "#16365C"

COLOR_BTN_SAVE = "#2E7D32"
COLOR_BTN_RESET = "#F9A825"

COLOR_WARNING = "#C62828"
COLOR_OK = "#1B5E20"
COLOR_REF_VALUE = "#16365C"        # 강조 표시값 텍스트도 Base 남색으로 통일

FONT_NAME = "맑은 고딕"


# ----------------------------------------------------------------------
# 공용 헬퍼 함수
# ----------------------------------------------------------------------
def parse_float(value, field_name=""):
    """문자열(쉼표 포함 가능)을 float로 변환. 실패 시 ValueError."""
    value = (value or "").strip().replace(",", "")
    if value == "":
        return 0.0
    try:
        return float(value)
    except ValueError:
        raise ValueError(f"'{field_name}' 항목에는 숫자만 입력해야 합니다. (입력값: {value})")


def format_with_commas(raw_text):
    """숫자만 추출해 천단위 콤마 문자열로 변환 (정수 기준)."""
    digits = "".join(ch for ch in raw_text if ch.isdigit())
    if digits == "":
        return ""
    digits = str(int(digits))
    return f"{int(digits):,}"


class LabeledEntry:
    """라벨 + 입력란(옵션: 천단위 콤마 자동 포맷)을 한 세트로 배치."""

    def __init__(self, parent, label_text, row, column=0, width=24, unit="",
                 columnspan_label=1, numeric=True,
                 on_keyrelease=None, on_focusout=None):
        self.numeric = numeric
        self.label = tk.Label(
            parent, text=label_text, bg=parent["bg"], fg=COLOR_TEXT_DARK,
            font=(FONT_NAME, 10), anchor="w", justify="left"
        )
        self.label.grid(row=row, column=column, sticky="w", padx=(10, 6), pady=6,
                         columnspan=columnspan_label)

        entry_col = column + columnspan_label
        self.var = tk.StringVar()
        self.entry = tk.Entry(
            parent, textvariable=self.var, width=width,
            bg=COLOR_ENTRY_BG, fg=COLOR_ENTRY_FG,
            insertbackground=COLOR_ENTRY_FG,
            relief="flat", highlightthickness=1,
            highlightbackground=COLOR_ACCENT, highlightcolor=COLOR_ACCENT,
            font=(FONT_NAME, 10),
            justify="right" if numeric else "left",
        )
        self.entry.grid(row=row, column=entry_col, sticky="w", padx=(0, 6), pady=6)

        if numeric:
            self.entry.bind("<KeyRelease>", self._on_key_release, add="+")
        if on_keyrelease:
            # 타이핑할 때마다 즉시 반영 (예: 잔여모집단 실시간 미리보기) - 팝업 없음
            self.entry.bind("<KeyRelease>", lambda e: on_keyrelease(), add="+")
        if on_focusout:
            # 입력을 마치고 다른 곳을 클릭/Tab 했을 때만 검증(팝업 등) 실행
            # -> 타이핑 도중 팝업이 반복해서 뜨는 것을 방지
            self.entry.bind("<FocusOut>", lambda e: on_focusout(), add="+")

        if unit:
            tk.Label(
                parent, text=unit, bg=parent["bg"], fg=COLOR_TEXT_MUTED,
                font=(FONT_NAME, 9)
            ).grid(row=row, column=entry_col + 1, sticky="w", pady=6)

    def _on_key_release(self, event):
        if event.keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                             "Shift_L", "Shift_R", "Tab"):
            return
        raw = self.var.get()
        cursor_from_end = len(raw) - self.entry.index(tk.INSERT)
        formatted = format_with_commas(raw)
        self.var.set(formatted)
        new_pos = max(len(formatted) - cursor_from_end, 0)
        self.entry.icursor(new_pos)

    def get(self):
        return self.var.get()

    def set(self, value):
        self.var.set(value)

    def reset(self):
        self.var.set("")


# ----------------------------------------------------------------------
# 메인 애플리케이션
# ----------------------------------------------------------------------
class AuditSamplingApp(tk.Tk):

    MENU_ITEMS = [
        ("guide", "1. Tool 사용가이드"),
        ("stage1", "2. 표본감사유형선택"),
        ("stage2", "3. 통계적표본감사"),
        ("stage3", "4. 비통계적표본감사"),
    ]

    def __init__(self):
        super().__init__()
        self.title("표본감사 Template (Audit Sampling Template)")
        self.configure(bg=COLOR_BG_MAIN)
        setup_ttk_styles()

        # 창 크기: 먼저 "복원 시 크기"(모니터 절반, 화면 중앙)를 지정해두고
        # 시작은 최대화 상태로 연다. 이렇게 해야 최대화 버튼을 눌러 창을
        # 작게 줄였을 때, 제목표시줄만 남고 찌그러지는 문제 없이 모니터
        # 절반 크기로 정상 복원된다.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        half_w, half_h = int(screen_w * 0.5), int(screen_h * 0.6)
        x = (screen_w - half_w) // 2
        y = (screen_h - half_h) // 2
        self.geometry(f"{half_w}x{half_h}+{x}+{y}")
        self.minsize(900, 600)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        # 사이드바 각 메뉴의 잠금 상태 (stage2/3는 1단계 완료 전까지 잠김)
        self.menu_enabled = {"guide": True, "stage1": True, "stage2": False, "stage3": False}
        self.menu_buttons = {}
        self.current_stage = "guide"

        # 탭(stage)별 스크롤 캔버스 등록소. 마우스 휠은 스크롤바 위가 아니라
        # 현재 화면에 보이는 탭(self.current_stage) 아무 곳에서나 굴려도
        # 그 탭의 캔버스가 스크롤되도록 전역으로 한 번만 바인딩한다.
        self.stage_canvases = {}

        self._build_header()
        self._build_body_layout()
        self._build_sidebar()
        self._build_guide_stage()
        self._build_stage1()
        self.stage2_tab = SamplingSubTab(self, "statistical")
        self.stage3_tab = SamplingSubTab(self, "nonstatistical")

        self._bind_global_scroll_and_focus_events()

        self._show_stage("guide")

    # ------------------------------------------------------------
    # 전역 이벤트: (1) 마우스 휠 스크롤 (2) 창 전환 시 콤보박스 드롭다운 잔상 제거
    # ------------------------------------------------------------
    def _bind_global_scroll_and_focus_events(self):
        # 어떤 탭이 보이든, 해당 탭의 스크롤 캔버스 위 임의의 지점을 클릭한 뒤
        # 마우스 휠만으로 위아래 이동이 가능하도록 전역(bind_all)으로 한 번만 바인딩.
        def _on_mousewheel(event):
            canvas = self.stage_canvases.get(self.current_stage)
            if canvas is not None:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_mousewheel_linux(direction):
            def _handler(event):
                canvas = self.stage_canvases.get(self.current_stage)
                if canvas is not None:
                    canvas.yview_scroll(direction, "units")
            return _handler

        self.bind_all("<MouseWheel>", _on_mousewheel)
        # Windows/Mac은 <MouseWheel>, 리눅스는 <Button-4>/<Button-5> 사용
        self.bind_all("<Button-4>", _on_mousewheel_linux(-1))
        self.bind_all("<Button-5>", _on_mousewheel_linux(1))

        # Alt+Tab 등으로 다른 창으로 전환했을 때, 열려 있던 ttk.Combobox의
        # 드롭다운 목록(별도의 override-redirect 팝업 창)이 화면 위에 잔상처럼
        # 남아있는 문제를 막기 위해, 창이 비활성화되는 시점에 강제로 닫아준다.
        self.bind("<Deactivate>", self._close_open_dropdowns, add="+")
        self.bind("<FocusOut>", self._on_root_focus_out, add="+")

    def _close_open_dropdowns(self, event=None):
        try:
            self.tk.call("ttk::combobox::Unpost")
        except tk.TclError:
            pass

    def _on_root_focus_out(self, event):
        # 위젯 간 포커스 이동(콤보박스 클릭 등)에서는 반복적으로 발생하므로,
        # 최상위 창(self) 자체가 포커스를 잃을 때(=다른 앱/창으로 전환)만 처리.
        if event.widget is self:
            self._close_open_dropdowns()

    # ------------------------------------------------------------
    # 레이아웃 뼈대
    # ------------------------------------------------------------
    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_HEADER_BAR, height=56)
        header.pack(fill="x", side="top")
        tk.Label(
            header, text="☰  표본감사(Audit Sampling) Template",
            bg=COLOR_HEADER_BAR, fg=COLOR_TEXT_LIGHT,
            font=(FONT_NAME, 14, "bold")
        ).pack(side="left", padx=18, pady=12)

    def _build_body_layout(self):
        self.body = tk.Frame(self, bg=COLOR_BG_MAIN)
        self.body.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(self.body, bg=COLOR_SIDEBAR_BG, width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.content_container = tk.Frame(self.body, bg=COLOR_BG_MAIN)
        self.content_container.pack(side="left", fill="both", expand=True)

        # 각 stage 프레임을 미리 만들어 겹쳐두고(tkraise), 파괴하지 않는다.
        # -> 탭 전환해도 입력값이 절대 사라지지 않음
        self.stage_frames = {}

    def _build_sidebar(self):
        tk.Label(
            self.sidebar, text="메뉴", bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_TEXT,
            font=(FONT_NAME, 10, "bold")
        ).pack(fill="x", padx=16, pady=(18, 8), anchor="w")

        for key, label in self.MENU_ITEMS:
            btn = tk.Label(
                self.sidebar, text=label, bg=COLOR_SIDEBAR_ITEM,
                fg=COLOR_SIDEBAR_TEXT, font=(FONT_NAME, 11),
                anchor="w", padx=16, pady=12, cursor="hand2"
            )
            btn.pack(fill="x", padx=10, pady=4)
            btn.bind("<Button-1>", lambda e, k=key: self._on_menu_click(k))
            self.menu_buttons[key] = btn

        self._refresh_sidebar_style()

        tk.Label(
            self.sidebar,
            text=("※ 2·3단계 메뉴는 1단계에서\n"
                  "   표본감사유형을 선택해야\n"
                  "   활성화됩니다."),
            bg=COLOR_SIDEBAR_BG, fg=COLOR_SIDEBAR_TEXT_DISABLED,
            font=(FONT_NAME, 8), justify="left", anchor="w"
        ).pack(fill="x", padx=16, pady=(20, 8))

    def _refresh_sidebar_style(self):
        for key, _ in self.MENU_ITEMS:
            btn = self.menu_buttons[key]
            if key == self.current_stage:
                btn.configure(bg=COLOR_SIDEBAR_ITEM_ACTIVE, fg="#FFFFFF",
                               font=(FONT_NAME, 11, "bold"))
            elif self.menu_enabled.get(key):
                btn.configure(bg=COLOR_SIDEBAR_ITEM, fg=COLOR_SIDEBAR_TEXT,
                               font=(FONT_NAME, 11))
            else:
                btn.configure(bg=COLOR_SIDEBAR_ITEM_DISABLED,
                               fg=COLOR_SIDEBAR_TEXT_DISABLED,
                               font=(FONT_NAME, 11))

    def _on_menu_click(self, key):
        if not self.menu_enabled.get(key):
            messagebox.showinfo(
                "잠긴 메뉴",
                "먼저 '1. 표본감사유형선택' 탭에서\n"
                "필수 입력값을 채우고 표본감사유형(통계적/비통계적)을\n"
                "선택해야 이 메뉴를 이용할 수 있습니다."
            )
            return
        self._show_stage(key)

    def _show_stage(self, key):
        self.current_stage = key
        self.stage_frames[key].tkraise()
        self._refresh_sidebar_style()

    # ------------------------------------------------------------
    # 준비중 화면 (2, 3단계 - 추후 실제 내용으로 교체 예정)
    # ------------------------------------------------------------
    def _build_guide_stage(self):
        outer = tk.Frame(self.content_container, bg=COLOR_BG_MAIN)
        outer.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.stage_frames["guide"] = outer

        canvas = tk.Canvas(outer, bg=COLOR_BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas, bg=COLOR_BG_MAIN)
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.stage_canvases["guide"] = canvas

        tk.Label(body, text="1. Tool 사용가이드", bg=COLOR_BG_MAIN, fg=COLOR_TEXT_LIGHT,
                  font=(FONT_NAME, 15, "bold")).pack(anchor="w", padx=18, pady=(14, 4))

        def make_section(title):
            outer2 = tk.Frame(body, bg=COLOR_BG_MAIN)
            outer2.pack(fill="x", padx=18, pady=(14, 0))
            header = tk.Frame(outer2, bg=COLOR_SECTION_HEADER)
            header.pack(fill="x")
            tk.Label(header, text=title, bg=COLOR_SECTION_HEADER, fg=COLOR_TEXT_DARK,
                      font=(FONT_NAME, 11, "bold"), anchor="w").pack(fill="x", padx=10, pady=6)
            b = tk.Frame(outer2, bg=COLOR_SECTION_BODY, padx=12, pady=10)
            b.pack(fill="x")
            return b

        def add_text(parent, text):
            tk.Label(parent, text=text, bg=parent["bg"], fg=COLOR_TEXT_DARK,
                      font=(FONT_NAME, 10), justify="left", anchor="w", wraplength=1000
                      ).pack(anchor="w", pady=3)

        f1 = make_section("전체 흐름")
        add_text(f1, "① '2. 표본감사유형선택' 탭에서 기본정보·모집단·수행중요성을 입력하고 "
                      "통계적/비통계적 표본감사유형을 선택합니다.")
        add_text(f1, "② 선택한 유형에 따라 '3. 통계적표본감사' 또는 '4. 비통계적표본감사' 메뉴가 "
                      "활성화됩니다.")
        add_text(f1, "③ 위험요소 정의 → 표본크기 산정 → 표본결과평가 → 최종결론 → 결과파일저장"
                      "(Excel) 순서로 진행합니다.")

        f2 = make_section("2단계: 표본감사유형선택 탭")
        add_text(f2, "· 기본정보(회사명/상장여부/수행중요성), A(테스트대상 정의), B(특정항목 선정) "
                      "순서로 입력합니다. C(잔여모집단)는 자동 계산됩니다.")
        add_text(f2, "· 특정항목 선정은 표본감사와 별도로 테스트할 예정인 항목에 대한 정보입니다.")
        add_text(f2, "· C(잔여모집단금액)가 수행중요성보다 작거나 같으면 표본감사 적용이 적정한지 "
                      "재검토하시기 바랍니다. 관련하여 경고 팝업이 생성되며, '예'를 선택해야 다음 "
                      "단계(표본감사유형 선택)로 진행할 수 있습니다.")
        add_text(f2, "· 통계적 또는 비통계적 중 하나를 선택하면 해당 메뉴만 활성화됩니다.")
        add_text(f2, "· 우측 하단 'Reset' 버튼을 누르면 1단계 입력내용이 초기화됩니다.")

        f3 = make_section("3·4단계: 통계적 / 비통계적 표본감사 탭 공통")
        add_text(f3, "① 위험요소 입력: 고유위험(IR)·통제위험(CR)·분석적절차(AP1)·기타세부절차(AP2) "
                      "계수를 선택한 뒤 「TD/신뢰수준 계산」 버튼을 누릅니다(감사위험(AR)은 1단계 "
                      "상장여부로 자동 결정됩니다).")
        add_text(f3, "   * 고유위험수준은 2302T 조서와 일관되게 작성하여야 함에 유의합니다.")
        add_text(f3, "   * 통제위험계수(CR)의 'None'은 통제의 운영효과성(Operating Effectiveness) "
                      "테스트를 수행하지 않은 경우 선택합니다(설계평가·walkthrough 수행 여부와 무관). "
                      "Limited/Normal/Extended는 운영효과성 테스트를 수행한 경우, 그 테스트 범위에 "
                      "따라 선택합니다.")
        add_text(f3, "   * 분석적절차수준은 비율분석(Ratio Test)·추세분석(Trend Test)·산업비교분석을 "
                      "수행한 경우 Less Extensive를, 합리성테스트(Reasonableness Test)·회귀분석테스트"
                      "(Regression Analysis)를 수행한 경우 More Extensive를 선택할 수 있습니다. 관련 "
                      "조서번호와 실제 수행된 절차를 입력해야 다음 단계로 진행할 수 있습니다.")
        add_text(f3, "② 표본크기 산정 파라미터: 예상왜곡표시율(%)을 입력하면 예상왜곡표시집계액· "
                      "허용오류율이 자동으로 표시됩니다. 모집단 성격(계층화 여부)과 표본선택방법도 "
                      "선택합니다.")
        add_text(f3, "   * 계층화 여부는 모집단의 금액 분포, 위험 특성 및 항목 간 변동성을 고려하여 "
                      "결정합니다. 모집단 내 금액 또는 위험도의 편차가 큰 경우에는 계층화를 적용하는 "
                      "것이 일반적으로 적절하며, 모집단이 전반적으로 동질적인 경우에는 비계층화 "
                      "모집단으로 표본을 선정할 수 있습니다.")
        add_text(f3, "③ '표본크기 계산' 버튼을 누르면 최소 표본크기와(해당 시) 표본간격이 계산됩니다. "
                      "표본크기는 최소 5건, 최대 잔여모집단 건수로 제한됩니다.")
        add_text(f3, "④ 표본결과평가")
        add_text(f3, "   * 통계적: 실제 테스트한 각 표본에서 발견된 오차금액을 기준으로 점증허용"
                      "(Incremental Allowance)을 반영하여 투사오류를 산정하므로, 반드시 표본별 테스트 "
                      "상세내역을 입력하여야 합니다. 실제 테스트한 표본의 장부금액 및 감사금액을 표에 "
                      "직접 입력하거나 엑셀에서 복사하여 붙여넣을 수 있습니다(20건 이하). 표본 건수가 "
                      "20건을 초과하는 경우에는 엑셀 템플릿을 다운로드하여 Template 컬럼에 필요한 정보를 "
                      "입력한 후 업로드해야 합니다. Template의 구조, 수식 및 컬럼은 수정할 수 없습니다.")
        add_text(f3, "   * 비통계적: 비례식으로 투사오류를 산정하므로, 실제 선정된 표본 금액/건수와 "
                      "표본에서 실제 발견된 오차 금액을 입력합니다.")
        add_text(f3, "⑤ 최종 총 투사오류가 수행중요성 이상이면 경고 팝업이 생성되며, 「추가수행한 "
                      "감사절차」를 반드시 기재하여야 다음 단계(결과 엑셀 저장)로 진행할 수 있습니다.")
        add_text(f3, "⑥ '결과 엑셀 저장' 버튼으로 입력값·계산결과를 정리한 엑셀 파일을 다운로드할 "
                      "수 있습니다.")

    # ------------------------------------------------------------
    # 1단계: 표본감사유형선택 탭
    # ------------------------------------------------------------
    def _build_stage1(self):
        outer = tk.Frame(self.content_container, bg=COLOR_BG_MAIN)
        outer.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.stage_frames["stage1"] = outer

        # 스크롤 가능한 캔버스 (입력 항목이 많으므로)
        canvas = tk.Canvas(outer, bg=COLOR_BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.body_frame = tk.Frame(canvas, bg=COLOR_BG_MAIN)

        self.body_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.body_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.stage_canvases["stage1"] = canvas

        self._build_section_basic_info()      # 기본정보 (1)(2)(3)
        self._build_section_A_population()    # A. (a)(b)(c)
        self._build_section_key_items()       # B. (a)(b)(c)
        self._build_section_remaining()       # C. (자동계산)
        self._build_section_stage2_selector()
        self._build_bottom_buttons(outer)

    def _make_section(self, title_text):
        outer = tk.Frame(self.body_frame, bg=COLOR_BG_MAIN)
        outer.pack(fill="x", padx=18, pady=(14, 0))

        header = tk.Frame(outer, bg=COLOR_SECTION_HEADER)
        header.pack(fill="x")
        tk.Label(
            header, text=title_text, bg=COLOR_SECTION_HEADER, fg=COLOR_TEXT_DARK,
            font=(FONT_NAME, 11, "bold"), anchor="w"
        ).pack(fill="x", padx=10, pady=6)

        body = tk.Frame(outer, bg=COLOR_SECTION_BODY, padx=6, pady=6)
        body.pack(fill="x")
        return body

    # ---- 기본정보: (1) 회사명 (2) 상장여부 (3) 수행중요성 ----
    def _build_section_basic_info(self):
        f = self._make_section("기본정보")

        self.company_name = LabeledEntry(
            f, "(1) 회사명", 0, numeric=False, width=30
        )

        tk.Label(
            f, text="(2) 상장여부", bg=f["bg"], fg=COLOR_TEXT_DARK,
            font=(FONT_NAME, 10)
        ).grid(row=1, column=0, sticky="w", padx=(10, 6), pady=6)

        self.listed_status = tk.StringVar(value="")
        radio_frame = tk.Frame(f, bg=f["bg"])
        radio_frame.grid(row=1, column=1, sticky="w", pady=6)
        make_radio(radio_frame, "상장", "상장", self.listed_status).pack(side="left", padx=(0, 14))
        make_radio(radio_frame, "기타", "기타", self.listed_status).pack(side="left")

        # (3) 수행중요성 - 별도 섹션이 아니라 기본정보 안으로 이동
        self.materiality = LabeledEntry(
            f, "(3) 수행중요성", 2, unit="원",
            on_focusout=self._validate_and_toggle_stage2
        )

    # ---- A. 테스트대상 정의 (a)(b)(c) ----
    def _build_section_A_population(self):
        f = self._make_section("A. 테스트대상 정의 (Total Population)")
        self.test_target = LabeledEntry(
            f, "(a) 테스트대상항목", 0, numeric=False, width=40
        )
        self.pop_amount = LabeledEntry(
            f, "(b) 테스트대상 총 금액(절대값)", 1, unit="원",
            on_keyrelease=self._recalculate_remaining,
            on_focusout=self._validate_and_toggle_stage2
        )
        self.pop_count = LabeledEntry(
            f, "(c) 테스트대상 총 건수", 2, unit="건",
            on_keyrelease=self._recalculate_remaining,
            on_focusout=self._validate_and_toggle_stage2
        )

    # ---- B. 특정항목 선정 ----
    def _build_section_key_items(self):
        f = self._make_section("B. 특정항목 선정 (Selection of Specific Items)")
        self.key_amount = LabeledEntry(
            f, "(a) 특정항목선정 금액(절대값)", 0, unit="원",
            on_keyrelease=self._recalculate_remaining,
            on_focusout=self._validate_and_toggle_stage2
        )
        self.key_count = LabeledEntry(
            f, "(b) 특정항목 선정 건수", 1, unit="건",
            on_keyrelease=self._recalculate_remaining,
            on_focusout=self._validate_and_toggle_stage2
        )
        self.key_actual_error = LabeledEntry(
            f, "(c) 특정항목 실제 발견 오차 금액", 2, unit="원"
        )

    # ---- (6) 잔여모집단 (자동계산) ----
    def _build_section_remaining(self):
        f = self._make_section("C. 잔여모집단 금액 및 건수 (자동계산)")
        tk.Label(
            f, text="잔여모집단 금액 / 건수", bg=f["bg"], fg=COLOR_TEXT_DARK,
            font=(FONT_NAME, 10)
        ).grid(row=0, column=0, sticky="w", padx=(10, 6), pady=8)
        self.remaining_display = tk.Label(
            f, text="- 원  /  - 건", bg=f["bg"], fg=COLOR_REF_VALUE,
            font=(FONT_NAME, 11, "bold")
        )
        self.remaining_display.grid(row=0, column=1, sticky="w", pady=8)
        tk.Label(
            f, text="(= 테스트대상 총액·총건수 − 특정항목선정 금액·건수)",
            bg=f["bg"], fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 8, "italic")
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 6))

    # ---- 2단계 선택 (통계적/비통계적) ----
    def _build_section_stage2_selector(self):
        f = self._make_section("2단계: 표본감사유형 선택")
        tk.Label(
            f, text="(1단계 필수 입력값 확인 후 활성화됩니다)",
            bg=f["bg"], fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 8, "italic")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 2))

        self.sampling_type = tk.StringVar(value="")
        self.radio_statistical = ttk.Radiobutton(
            f, text="통계적 표본감사", value="statistical",
            variable=self.sampling_type, style="App.TRadiobutton",
            state="disabled", command=self._on_sampling_type_selected
        )
        self.radio_statistical.grid(row=1, column=0, sticky="w", padx=(10, 20), pady=8)

        self.radio_nonstatistical = ttk.Radiobutton(
            f, text="비통계적 표본감사", value="nonstatistical",
            variable=self.sampling_type, style="App.TRadiobutton",
            state="disabled", command=self._on_sampling_type_selected
        )
        self.radio_nonstatistical.grid(row=1, column=1, sticky="w", pady=8)

        self.stage2_status_label = tk.Label(
            f, text="🔒 잠김", bg=f["bg"], fg=COLOR_WARNING, font=(FONT_NAME, 9, "bold")
        )
        self.stage2_status_label.grid(row=1, column=2, sticky="w", padx=(20, 10))

    # ---- 하단 Reset 버튼 ----
    def _build_bottom_buttons(self, parent):
        btn_bar = tk.Frame(parent, bg=COLOR_BG_MAIN)
        btn_bar.pack(fill="x", side="bottom", padx=18, pady=14)

        tk.Button(
            btn_bar, text="Reset", command=self._on_reset,
            bg=COLOR_BTN_RESET, fg="white", font=(FONT_NAME, 11, "bold"),
            relief="flat", padx=18, pady=8, cursor="hand2"
        ).pack(side="right")

    # ------------------------------------------------------------
    # 로직: 잔여모집단 자동계산
    # ------------------------------------------------------------
    def _get_remaining(self):
        # A(테스트대상 총금액/총건수)가 아직 비어있으면 "계산 불가(미입력)" 상태로 취급.
        # (parse_float은 빈칸을 0으로 처리하는데, 그러면 항상 "잔여모집단<=수행중요성"
        #  조건이 성립해버려 수행중요성만 입력해도 즉시 경고가 뜨는 버그가 있었음)
        if not self.pop_amount.get().strip() or not self.pop_count.get().strip():
            return None, None
        try:
            pop_amt = parse_float(self.pop_amount.get())
            pop_cnt = parse_float(self.pop_count.get())
            key_amt = parse_float(self.key_amount.get())
            key_cnt = parse_float(self.key_count.get())
            return pop_amt - key_amt, pop_cnt - key_cnt
        except ValueError:
            return None, None

    def _recalculate_remaining(self):
        """A/B 입력 중 실시간으로 C(잔여모집단) 표시만 갱신. 팝업 없음."""
        remain_amt, remain_cnt = self._get_remaining()
        if remain_amt is None:
            self.remaining_display.config(text="- 원  /  - 건")
            return
        self.remaining_display.config(
            text=f"{remain_amt:,.0f} 원  /  {remain_cnt:,.0f} 건"
        )

    # ------------------------------------------------------------
    # 로직: C(잔여모집단) 확정 시점(A/B/수행중요성 중 어느 필드든 포커스아웃)
    #       -> 잔여모집단 vs 수행중요성 비교 -> 팝업/2단계 활성화 판단
    # ------------------------------------------------------------
    def _validate_and_toggle_stage2(self):
        # 혹시 실시간 표시가 아직 안 갱신됐을 수 있으니 먼저 최신화
        self._recalculate_remaining()

        remain_amt, _ = self._get_remaining()
        try:
            pm = parse_float(self.materiality.get())
        except ValueError:
            pm = None

        if remain_amt is None or pm is None or pm == 0:
            # A/B/수행중요성 중 아직 입력이 불완전하면 2단계는 잠금 상태 유지
            return

        if remain_amt <= pm:
            # C(잔여모집단) <= 수행중요성 -> 경고 팝업
            proceed = messagebox.askyesno(
                "경고",
                "표본감사 모집단이 수행중요성과 같거나 작습니다.\n"
                "진행하시겠습니까?"
            )
            if proceed:
                self._enable_stage2_selector()
            else:
                self._disable_stage2_selector()
        else:
            # 정상 케이스: C(잔여모집단) > 수행중요성 -> 팝업 없이 바로 2단계 활성화
            self._enable_stage2_selector()

    def _enable_stage2_selector(self):
        self.radio_statistical.configure(state="normal")
        self.radio_nonstatistical.configure(state="normal")
        self.stage2_status_label.configure(text="✅ 선택 가능", fg=COLOR_OK)

    def _disable_stage2_selector(self):
        self.radio_statistical.configure(state="disabled")
        self.radio_nonstatistical.configure(state="disabled")
        self.sampling_type.set("")
        self.stage2_status_label.configure(text="🔒 잠김", fg=COLOR_WARNING)
        self.menu_enabled["stage2"] = False
        self.menu_enabled["stage3"] = False
        self._refresh_sidebar_style()

    def _on_sampling_type_selected(self):
        choice = self.sampling_type.get()
        if choice == "statistical":
            self.menu_enabled["stage2"] = True
            self.menu_enabled["stage3"] = False
        elif choice == "nonstatistical":
            self.menu_enabled["stage2"] = False
            self.menu_enabled["stage3"] = True
        self._refresh_sidebar_style()

    # ------------------------------------------------------------
    # Save / Reset
    # ------------------------------------------------------------
    def _on_reset(self):
        if not messagebox.askyesno("초기화 확인", "1단계 입력값을 모두 초기화하시겠습니까?"):
            return
        self.company_name.reset()
        self.listed_status.set("")
        self.test_target.reset()
        self.pop_amount.reset()
        self.pop_count.reset()
        self.key_amount.reset()
        self.key_count.reset()
        self.key_actual_error.reset()
        self.materiality.reset()
        self.remaining_display.config(text="- 원  /  - 건")
        self._disable_stage2_selector()
        messagebox.showinfo("초기화 완료", "1단계 입력값이 초기화되었습니다.")


class PasteableGrid:
    """엑셀에서 그대로 복사해 셀서식 변경 없이 붙여넣을 수 있는
    2열(장부금액/감사(확인)금액) 입력 그리드. 최대 max_rows(기본 20)행."""

    def __init__(self, parent, max_rows=20):
        self.max_rows = max_rows
        self.entries_book = []
        self.entries_audit = []

        container = tk.Frame(parent, bg=parent["bg"])
        container.pack(fill="x", padx=4, pady=4)

        header = tk.Frame(container, bg=parent["bg"])
        header.pack(fill="x")
        tk.Label(header, text="#", width=4, bg=parent["bg"], fg=COLOR_TEXT_MUTED,
                  font=(FONT_NAME, 9, "bold")).grid(row=0, column=0)
        tk.Label(header, text="장부금액 (Book Value)", width=20, bg=parent["bg"],
                  fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 9, "bold")).grid(row=0, column=1)
        tk.Label(header, text="감사(확인)금액 (Audited Value)", width=24, bg=parent["bg"],
                  fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 9, "bold")).grid(row=0, column=2)

        rows_frame = tk.Frame(container, bg=parent["bg"])
        rows_frame.pack(fill="x")

        for r in range(max_rows):
            tk.Label(rows_frame, text=str(r + 1), width=4, bg=parent["bg"],
                      fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 9)).grid(row=r, column=0, pady=1)
            e1 = tk.Entry(rows_frame, width=20, justify="right", font=(FONT_NAME, 9))
            e1.grid(row=r, column=1, padx=2, pady=1)
            e2 = tk.Entry(rows_frame, width=24, justify="right", font=(FONT_NAME, 9))
            e2.grid(row=r, column=2, padx=2, pady=1)
            e1.bind("<KeyRelease>", self._on_key_release, add="+")
            e2.bind("<KeyRelease>", self._on_key_release, add="+")
            e1.bind("<<Paste>>", lambda ev, rr=r, cc=0: self._on_paste(ev, rr, cc))
            e2.bind("<<Paste>>", lambda ev, rr=r, cc=1: self._on_paste(ev, rr, cc))
            self.entries_book.append(e1)
            self.entries_audit.append(e2)

    def _on_key_release(self, event):
        if event.keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                             "Shift_L", "Shift_R", "Tab"):
            return
        entry = event.widget
        raw = entry.get()
        cursor_from_end = len(raw) - entry.index(tk.INSERT)
        formatted = format_with_commas(raw)
        entry.delete(0, tk.END)
        entry.insert(0, formatted)
        new_pos = max(len(formatted) - cursor_from_end, 0)
        entry.icursor(new_pos)

    def _on_paste(self, event, row, col):
        try:
            clip = event.widget.clipboard_get()
        except tk.TclError:
            return "break"
        lines = [ln for ln in clip.replace("\r", "").split("\n") if ln != ""]
        truncated = False
        for i, line in enumerate(lines):
            rr = row + i
            if rr >= self.max_rows:
                truncated = True
                break
            cols = line.split("\t")
            for j, val in enumerate(cols):
                cc = col + j
                val = format_with_commas(val.strip())
                if cc == 0 and rr < self.max_rows:
                    self.entries_book[rr].delete(0, tk.END)
                    self.entries_book[rr].insert(0, val)
                elif cc == 1 and rr < self.max_rows:
                    self.entries_audit[rr].delete(0, tk.END)
                    self.entries_audit[rr].insert(0, val)
        if truncated:
            messagebox.showwarning(
                "행 초과",
                f"최대 {self.max_rows}행까지만 입력 가능합니다.\n"
                f"그 이상은 '엑셀 업로드' 기능을 이용해주세요."
            )
        return "break"

    def clear(self):
        for e in self.entries_book + self.entries_audit:
            e.delete(0, tk.END)

    def get_rows(self):
        rows = []
        for b, a in zip(self.entries_book, self.entries_audit):
            bv, av = b.get().strip(), a.get().strip()
            if bv == "" and av == "":
                continue
            try:
                bv = float(bv.replace(",", "")) if bv else 0.0
                av = float(av.replace(",", "")) if av else 0.0
            except ValueError:
                continue
            rows.append((bv, av))
        return rows


class SamplingSubTab:
    """'2. 통계적표본감사' / '3. 비통계적표본감사' 탭 공통 구현.

    mode: 'statistical' 또는 'nonstatistical'
    두 탭은 위험요소·표본크기 산정 로직은 100% 동일하고,
    표본결과평가 방식(정식 MUS 평가 vs 단순 비례식)만 다르다.
    """

    KEY = {"statistical": "stage2", "nonstatistical": "stage3"}
    TITLE = {"statistical": "2. 통계적표본감사", "nonstatistical": "3. 비통계적표본감사"}

    def __init__(self, app, mode):
        self.app = app
        self.mode = mode
        self.key = self.KEY[mode]
        self.uploaded_rows = None  # 엑셀 업로드된 표본결과(20행 초과 시)

        outer = tk.Frame(app.content_container, bg=COLOR_BG_MAIN)
        outer.place(relx=0, rely=0, relwidth=1, relheight=1)
        app.stage_frames[self.key] = outer

        canvas = tk.Canvas(outer, bg=COLOR_BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.body = tk.Frame(canvas, bg=COLOR_BG_MAIN)
        self.body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        app.stage_canvases[self.key] = canvas

        tk.Label(self.body, text=self.TITLE[mode], bg=COLOR_BG_MAIN, fg=COLOR_TEXT_LIGHT,
                  font=(FONT_NAME, 15, "bold")).pack(anchor="w", padx=18, pady=(14, 4))

        self._build_risk_section()
        self._build_size_param_section()
        self._build_size_result_section()
        if mode == "statistical":
            self._build_statistical_evaluation_section()
        else:
            self._build_nonstatistical_evaluation_section()
        self._build_final_conclusion_section()

    def _make_section(self, title_text):
        outer = tk.Frame(self.body, bg=COLOR_BG_MAIN)
        outer.pack(fill="x", padx=18, pady=(14, 0))
        header = tk.Frame(outer, bg=COLOR_SECTION_HEADER)
        header.pack(fill="x")
        tk.Label(header, text=title_text, bg=COLOR_SECTION_HEADER, fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 11, "bold"), anchor="w").pack(fill="x", padx=10, pady=6)
        body = tk.Frame(outer, bg=COLOR_SECTION_BODY, padx=6, pady=6)
        body.pack(fill="x")
        return body

    # ------------------------------------------------------------
    # ① 위험요소 입력 (Audit Risk Model)
    # ------------------------------------------------------------
    def _build_risk_section(self):
        f = self._make_section("① 위험요소 입력 (Audit Risk Model)")

        tk.Label(f, text="고유위험계수(IR)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        self.ir_var = tk.StringVar(value=list(IR_FACTORS.keys())[0])
        ttk.Combobox(f, textvariable=self.ir_var, values=list(IR_FACTORS.keys()),
                     state="readonly", width=40).grid(row=1, column=1, sticky="w", pady=6)

        tk.Label(f, text="통제위험계수(CR)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=2, column=0, sticky="w", padx=(10, 6), pady=6)
        self.cr_var = tk.StringVar(value=list(CR_FACTORS.keys())[0])
        ttk.Combobox(f, textvariable=self.cr_var, values=list(CR_FACTORS.keys()),
                     state="readonly", width=40).grid(row=2, column=1, sticky="w", pady=6)

        # ---- 분석절차계수(AP1) ----
        tk.Label(f, text="분석절차계수(AP1)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=3, column=0, sticky="w", padx=(10, 6), pady=6)
        self.ap1_var = tk.StringVar(value=list(AP1_FACTORS.keys())[0])
        ap1_combo = ttk.Combobox(f, textvariable=self.ap1_var, values=list(AP1_FACTORS.keys()),
                                   state="readonly", width=40)
        ap1_combo.grid(row=3, column=1, sticky="w", pady=6)

        # AP1이 None이 아닌 경우에만 표시되는 조서번호/수행절차 입력칸 (좌측정렬)
        self.ap1_detail_frame = tk.Frame(f, bg=f["bg"])
        self.ap1_detail_frame.grid(row=4, column=0, columnspan=3, sticky="w", padx=(10, 0))
        self.ap1_workpaper = LabeledEntry(
            self.ap1_detail_frame, "조서번호", 0, numeric=False, width=20)
        self.ap1_procedure = LabeledEntry(
            self.ap1_detail_frame, "수행절차", 1, numeric=False, width=55)
        self.ap1_detail_frame.grid_remove()

        def _toggle_ap1_detail(event=None):
            if self.ap1_var.get() == "None":
                self.ap1_detail_frame.grid_remove()
            else:
                self.ap1_detail_frame.grid()
        ap1_combo.bind("<<ComboboxSelected>>", _toggle_ap1_detail)

        # ---- 기타세부절차계수(AP2) ----
        tk.Label(f, text="기타세부절차계수(AP2)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=5, column=0, sticky="w", padx=(10, 6), pady=6)
        self.ap2_var = tk.StringVar(value=list(AP2_FACTORS.keys())[0])
        ap2_combo = ttk.Combobox(f, textvariable=self.ap2_var, values=list(AP2_FACTORS.keys()),
                                   state="readonly", width=40)
        ap2_combo.grid(row=5, column=1, sticky="w", pady=6)

        # AP2가 None이 아닌 경우에만 표시되는 조서번호/수행절차 입력칸 (좌측정렬)
        self.ap2_detail_frame = tk.Frame(f, bg=f["bg"])
        self.ap2_detail_frame.grid(row=6, column=0, columnspan=3, sticky="w", padx=(10, 0))
        self.ap2_workpaper = LabeledEntry(
            self.ap2_detail_frame, "조서번호", 0, numeric=False, width=20)
        self.ap2_procedure = LabeledEntry(
            self.ap2_detail_frame, "수행절차", 1, numeric=False, width=55)
        self.ap2_detail_frame.grid_remove()

        def _toggle_ap2_detail(event=None):
            if self.ap2_var.get() == "None":
                self.ap2_detail_frame.grid_remove()
            else:
                self.ap2_detail_frame.grid()
        ap2_combo.bind("<<ComboboxSelected>>", _toggle_ap2_detail)

        tk.Button(f, text="TD / 신뢰수준 계산", command=self._calc_td,
                  bg=COLOR_ACCENT, fg="white", font=(FONT_NAME, 9, "bold"),
                  relief="flat", padx=10, pady=4, cursor="hand2"
                  ).grid(row=7, column=0, sticky="w", padx=(10, 6), pady=(10, 6))

        self.td_display = tk.Label(f, text="TD: -  /  신뢰수준: -", bg=f["bg"],
                                     fg=COLOR_REF_VALUE, font=(FONT_NAME, 10, "bold"))
        self.td_display.grid(row=7, column=1, sticky="w", pady=(10, 6))

    def _calc_td(self):
        # AP1/AP2에서 'None'이 아닌 옵션을 선택한 경우, 조서번호·수행절차를
        # 반드시 입력해야만 다음 단계(TD/신뢰수준 계산 및 표본크기 계산)로
        # 진행할 수 있도록 차단한다.
        if self.ap1_var.get() != "None" and (
                not self.ap1_workpaper.get().strip() or not self.ap1_procedure.get().strip()):
            messagebox.showerror(
                "입력 필요",
                "분석절차계수(AP1)에서 'None'이 아닌 옵션을 선택한 경우,\n"
                "조서번호와 수행절차를 반드시 입력해야 다음 단계로 진행할 수 있습니다."
            )
            return
        if self.ap2_var.get() != "None" and (
                not self.ap2_workpaper.get().strip() or not self.ap2_procedure.get().strip()):
            messagebox.showerror(
                "입력 필요",
                "기타세부절차계수(AP2)에서 'None'이 아닌 옵션을 선택한 경우,\n"
                "조서번호와 수행절차를 반드시 입력해야 다음 단계로 진행할 수 있습니다."
            )
            return

        listed = self.app.listed_status.get()
        ar = AR_BY_LISTED.get(listed)
        if ar is None:
            messagebox.showwarning("입력 필요", "1단계에서 상장여부를 먼저 선택해주세요.")
            return
        ir = IR_FACTORS[self.ir_var.get()]
        cr = CR_FACTORS[self.cr_var.get()]
        ap1 = AP1_FACTORS[self.ap1_var.get()]
        ap2 = AP2_FACTORS[self.ap2_var.get()]
        td = ar / (ir * cr * ap1 * ap2)
        confidence = 1 - td
        note = "  ⚠ 신뢰수준≤0 → 최소표본(5건) 적용" if td >= 1 else ""
        self.td_display.config(text=f"TD: {td:.4f}  /  신뢰수준: {confidence:.4f}{note}")
        self._td = td
        self._confidence = confidence
        return td, confidence

    # ------------------------------------------------------------
    # ② 표본크기 산정 파라미터
    # ------------------------------------------------------------
    def _build_size_param_section(self):
        f = self._make_section("② 표본크기 산정 파라미터")

        self.expected_rate_entry = LabeledEntry(
            f, "예상왜곡표시율 (%) - 직접입력", 0, unit="%",
            on_keyrelease=self._update_expected_and_tolerable_display,
            on_focusout=self._update_expected_and_tolerable_display,
        )

        tk.Label(f, text="예상왜곡표시집계액 (자동계산 = 수행중요성 × 예상왜곡표시율)",
                  bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 9)
                  ).grid(row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        self.expected_amount_display = tk.Label(
            f, text="-", bg=f["bg"], fg=COLOR_REF_VALUE, font=(FONT_NAME, 10, "bold"))
        self.expected_amount_display.grid(row=1, column=1, sticky="w", pady=6)

        tk.Label(f, text="허용오류율 (자동계산 = 수행중요성 ÷ 잔여모집단금액)",
                  bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 9)
                  ).grid(row=2, column=0, sticky="w", padx=(10, 6), pady=6)
        self.tolerable_rate_display = tk.Label(
            f, text="-", bg=f["bg"], fg=COLOR_REF_VALUE, font=(FONT_NAME, 10, "bold"))
        self.tolerable_rate_display.grid(row=2, column=1, sticky="w", pady=6)

        tk.Label(f, text="모집단 성격 (계층화 여부)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=3, column=0, sticky="w", padx=(10, 6), pady=6)
        self.stratified_var = tk.StringVar(value="계층화되어 있음")
        radio_frame = tk.Frame(f, bg=f["bg"])
        radio_frame.grid(row=3, column=1, sticky="w", pady=6)
        make_radio(radio_frame, "계층화되어 있음", "계층화되어 있음", self.stratified_var
                    ).pack(side="left", padx=(0, 14))
        make_radio(radio_frame, "이외 (계층화되지 않음)", "이외", self.stratified_var
                    ).pack(side="left")

        tk.Label(f, text="표본선택방법", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=4, column=0, sticky="w", padx=(10, 6), pady=6)
        options = (SELECTION_METHODS_STATISTICAL if self.mode == "statistical"
                   else SELECTION_METHODS_NONSTATISTICAL)
        self.selection_method_var = tk.StringVar(value=options[0])
        ttk.Combobox(f, textvariable=self.selection_method_var, values=options,
                     state="readonly", width=30).grid(row=4, column=1, sticky="w", pady=6)
        if self.mode == "statistical":
            tk.Label(f, text="(통계적 표본감사는 임의 또는 무작위 선정을 지원하지 않습니다)",
                      bg=f["bg"], fg=COLOR_TEXT_MUTED, font=(FONT_NAME, 8, "italic")
                      ).grid(row=5, column=0, columnspan=2, sticky="w", padx=10)

    def _update_expected_and_tolerable_display(self):
        """예상왜곡표시율을 입력하는 즉시(버튼 클릭 없이) 예상왜곡표시집계액·허용오류율을
        실시간으로 갱신한다."""
        materiality_str = self.app.materiality.get()
        if not materiality_str:
            return
        materiality = parse_float(materiality_str)

        try:
            pE_pct = float(self.expected_rate_entry.get() or 0)
        except ValueError:
            return
        pE = pE_pct / 100.0
        expected_amount = materiality * pE
        self.expected_amount_display.config(text=f"{expected_amount:,.0f} 원")

        remain_amt, remain_cnt = self.app._get_remaining()
        if remain_amt is None or remain_amt <= 0:
            self.tolerable_rate_display.config(text="-")
            return
        pT = materiality / remain_amt
        self.tolerable_rate_display.config(text=f"{pT:.4%}")

    # ------------------------------------------------------------
    # ③ 최소 표본크기 산출결과
    # ------------------------------------------------------------
    def _build_size_result_section(self):
        f = self._make_section("③ 최소 표본크기 산출결과")

        tk.Button(f, text="표본크기 계산", command=self._calc_sample_size,
                  bg=COLOR_BTN_SAVE, fg="white", font=(FONT_NAME, 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2"
                  ).grid(row=0, column=0, sticky="w", padx=(10, 6), pady=8)

        self.sample_size_display = tk.Label(
            f, text="표본크기: -", bg=f["bg"], fg=COLOR_REF_VALUE, font=(FONT_NAME, 12, "bold"))
        self.sample_size_display.grid(row=0, column=1, sticky="w", pady=8)

        self.interval_display = tk.Label(
            f, text="표본간격: -", bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 10))
        self.interval_display.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 8))

    def _calc_sample_size(self):
        td_result = self._calc_td()
        if td_result is None:
            return
        td, confidence = td_result

        try:
            pE_pct = float(self.expected_rate_entry.get() or 0)
        except ValueError:
            messagebox.showerror("입력 오류", "예상왜곡표시율은 숫자로 입력해주세요.")
            return
        pE = pE_pct / 100.0

        materiality_str = self.app.materiality.get()
        remain_amt, remain_cnt = self.app._get_remaining()
        if remain_amt is None or not materiality_str:
            messagebox.showwarning("입력 필요", "1단계의 수행중요성 및 A/B 항목을 먼저 입력해주세요.")
            return
        materiality = parse_float(materiality_str)
        if remain_amt <= 0:
            messagebox.showerror("계산 불가", "잔여모집단금액이 0 이하입니다.")
            return

        expected_amount = materiality * pE
        pT = materiality / remain_amt
        # pE(사용자 입력, "수행중요성 대비" 비율)와 pT("잔여모집단 대비" 비율)는
        # 기준(분모)이 서로 다르므로, poisson_sample에 넘기기 전에 pE도
        # "잔여모집단 대비" 비율로 환산해서 단위를 맞춘다.
        # (엑셀 매크로 로직과 동일: 예상왜곡표시금액 ÷ 잔여모집단금액)
        pE_pop = expected_amount / remain_amt

        self.expected_amount_display.config(text=f"{expected_amount:,.0f} 원")
        self.tolerable_rate_display.config(text=f"{pT:.4%}")

        if pT <= 0 or pT >= 1 or pE_pop >= pT:
            messagebox.showerror(
                "계산 불가",
                "허용오류율이 예상왜곡표시율보다 커야 계산이 가능합니다.\n"
                "(예상왜곡표시율 / 수행중요성 / 모집단 입력값을 확인해주세요)"
            )
            return

        if td >= 1:
            # 신뢰수준이 0 이하로 설계됨 -> 다른 절차에서 위험을 이미 충분히
            # 커버했다고 보아 최소 표본(5건)만 요구 (매크로의 "최소범위 5" 로직과 동일)
            raw_n = MIN_SAMPLE_SIZE_FLOOR
        else:
            raw_n = poisson_sample(td, pE_pop, pT)

        # 계층화 X + 무작위/임의 선택 -> 1.5배 (비통계적 탭에만 적용)
        factor = 1.0
        if (self.mode == "nonstatistical"
                and self.stratified_var.get() == "이외"
                and self.selection_method_var.get() != SELECTION_METHODS_NONSTATISTICAL[0]):
            factor = NONSTAT_PENALTY_FACTOR
        adjusted_n = math.ceil(raw_n * factor)

        # 캡: 표본크기는 잔여모집단 건수를 초과할 수 없음
        pop_cap = int(remain_cnt) if remain_cnt else adjusted_n
        final_n = min(adjusted_n, pop_cap)
        # 하한: 최소 표본크기(5건) - 단, 잔여모집단 자체가 5건 미만이면 모집단 건수까지만
        final_n = max(final_n, min(MIN_SAMPLE_SIZE_FLOOR, pop_cap))
        final_n = max(final_n, 1)

        self.final_n = final_n

        # 표본간격: 통계적이면 항상, 비통계적이면 선택방법=체계적일 때만
        interval = None
        selection = self.selection_method_var.get()
        is_systematic = selection == SELECTION_METHODS_STATISTICAL[0] or \
                         selection == SELECTION_METHODS_NONSTATISTICAL[0]
        if self.mode == "statistical" or (self.mode == "nonstatistical" and is_systematic):
            interval = remain_amt / final_n

        self.interval = interval

        self.sample_size_display.config(text=f"표본크기: {final_n:,} 건")
        if interval is not None:
            self.interval_display.config(text=f"표본간격: {interval:,.0f} 원")
        else:
            self.interval_display.config(text="표본간격: 해당없음 (체계적 선정이 아님)")

    # ------------------------------------------------------------
    # ④-통계적 표본결과평가
    # ------------------------------------------------------------
    def _build_statistical_evaluation_section(self):
        f = self._make_section("④ 표본결과평가 (통계적 - MUS 정식 평가)")

        tk.Label(
            f, text="샘플테스트 상세내역 (20건 이하: 아래 표에 엑셀 그대로 복사/붙여넣기)",
            bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 9, "bold")
        ).pack(anchor="w", padx=10, pady=(4, 2))

        self.grid = PasteableGrid(f, max_rows=20)

        btn_row = tk.Frame(f, bg=f["bg"])
        btn_row.pack(fill="x", padx=10, pady=6)

        tk.Button(btn_row, text="그리드 초기화", command=self.grid.clear,
                  bg=COLOR_BTN_RESET, fg="white", font=(FONT_NAME, 9),
                  relief="flat", padx=8, pady=4, cursor="hand2").pack(side="left", padx=(0, 6))

        tk.Button(btn_row, text="엑셀 템플릿 다운로드", command=self._download_template,
                  bg=COLOR_ACCENT, fg="white", font=(FONT_NAME, 9),
                  relief="flat", padx=8, pady=4, cursor="hand2").pack(side="left", padx=(0, 6))

        tk.Button(btn_row, text="엑셀 업로드 (20건 초과 시)", command=self._upload_excel,
                  bg=COLOR_ACCENT, fg="white", font=(FONT_NAME, 9),
                  relief="flat", padx=8, pady=4, cursor="hand2").pack(side="left", padx=(0, 6))

        self.upload_status_label = tk.Label(
            f, text="", bg=f["bg"], fg=COLOR_OK, font=(FONT_NAME, 9, "italic"))
        self.upload_status_label.pack(anchor="w", padx=10, pady=(0, 6))

        tk.Button(f, text="표본결과평가 계산", command=self._calc_statistical_evaluation,
                  bg=COLOR_BTN_SAVE, fg="white", font=(FONT_NAME, 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(anchor="w", padx=10, pady=(6, 6))

        self.eval_result_display = tk.Label(
            f, text="", bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 10),
            justify="left", anchor="w")
        self.eval_result_display.pack(fill="x", padx=10, pady=(0, 8))

    def _download_template(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("오류", "openpyxl이 설치되어 있지 않습니다.")
            return
        path = filedialog.asksaveasfilename(
            title="엑셀 템플릿 저장", defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="샘플테스트내역_템플릿.xlsx",
        )
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "샘플테스트내역"
        headers = ["항목번호", "장부금액", "감사(확인)금액"]
        for col, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="153E6E")
            c.alignment = Alignment(horizontal="center")
        example_rows = [(1, 1000000, 950000), (2, 500000, 500000), (3, 2000000, 1800000)]
        for r, (no, book, audit) in enumerate(example_rows, start=2):
            ws.cell(row=r, column=1, value=no)
            ws.cell(row=r, column=2, value=book)
            ws.cell(row=r, column=3, value=audit)
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18
        try:
            wb.save(path)
            messagebox.showinfo(
                "저장 완료",
                f"템플릿이 저장되었습니다.\n{path}\n\n"
                "1행(헤더)은 그대로 두고, 2행부터 실제 데이터를 입력한 뒤\n"
                "'엑셀 업로드' 버튼으로 이 파일을 불러오시면 됩니다."
            )
        except Exception as e:
            messagebox.showerror("저장 실패", str(e))

    def _upload_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("오류", "openpyxl이 설치되어 있지 않습니다.")
            return
        path = filedialog.askopenfilename(
            title="샘플테스트내역 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
            rows = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                val1 = row[1] if len(row) > 1 else None
                val2 = row[2] if len(row) > 2 else None
                if val1 is None and val2 is None:
                    continue
                book = float(val1) if val1 not in (None, "") else 0.0
                audit = float(val2) if val2 not in (None, "") else 0.0
                rows.append((book, audit))
            if not rows:
                messagebox.showwarning(
                    "데이터 없음",
                    "2행부터 데이터를 찾지 못했습니다.\n"
                    "B열(장부금액)/C열(감사확인금액)에 값이 있는지 확인해주세요."
                )
                return
            self.uploaded_rows = rows
            self.upload_status_label.config(
                text=f"✅ 엑셀 업로드 완료: {len(rows)}건 (그리드 입력값보다 우선 적용됩니다)"
            )
        except Exception as e:
            import traceback
            messagebox.showerror("업로드 실패", f"{e}\n\n{traceback.format_exc()[-500:]}")

    def _calc_statistical_evaluation(self):
        if not hasattr(self, "_confidence"):
            msg = "먼저 ① TD/신뢰수준, ③ 표본크기를 계산해주세요."
            messagebox.showwarning("순서 오류", msg)
            self.eval_result_display.config(text=f"⚠ {msg}")
            return
        rows = self.uploaded_rows if self.uploaded_rows is not None else self.grid.get_rows()
        if not rows:
            msg = "투사오류 금액 계산을 위해 샘플테스트내역을 업로드하세요, 또는 입력하세요."
            messagebox.showwarning("입력 필요", msg)
            self.eval_result_display.config(text=f"⚠ {msg}")
            return

        interval = getattr(self, "interval", None)
        if not interval:
            msg = "먼저 ③ 표본크기(표본간격)를 계산해주세요."
            messagebox.showwarning("순서 오류", msg)
            self.eval_result_display.config(text=f"⚠ {msg}")
            return

        try:
            confidence = self._confidence

            large_items_actual_error = 0.0  # 장부금액이 표본간격을 초과하는 항목(자동 핵심항목 취급)
            normal_items = []  # (tainting, int_proj)
            for book, audit in rows:
                misstatement = book - audit
                if book > interval:
                    # 표본간격을 초과하는 개별항목 -> 비례투사 없이 실제오차 그대로 반영
                    large_items_actual_error += misstatement
                else:
                    tainting = (misstatement / book) if book > 0 else 0.0
                    int_proj = tainting * interval
                    normal_items.append((tainting, int_proj))

            # 기본정확도
            basic_precision = interval * reliability_factor(confidence, 0)

            # 순위(Tainting% 내림차순) -> 증분계수 반영
            normal_items.sort(key=lambda x: x[0], reverse=True)
            total_with_incremental = 0.0
            for rank, (tainting, int_proj) in enumerate(normal_items, start=1):
                r_n = reliability_factor(confidence, rank)
                r_n_1 = reliability_factor(confidence, rank - 1)
                incremental = r_n - r_n_1
                total_with_incremental += int_proj * incremental

            ulp = basic_precision + total_with_incremental + large_items_actual_error
            self.ulp = ulp

            key_error = parse_float(self.app.key_actual_error.get())
            final_total = ulp + key_error
            self.final_total_misstatement = final_total

            self.eval_result_display.config(
                text=(
                    f"기본정확도: {basic_precision:,.0f} 원\n"
                    f"점증허용 반영 투사오차 합계: {total_with_incremental:,.0f} 원\n"
                    f"표본간격 초과항목 실제오차 합계: {large_items_actual_error:,.0f} 원\n"
                    f"→ 오류상한(ULP, 잔여모집단): {ulp:,.0f} 원\n"
                    f"+ 특정항목 실제 발견 오차(1단계 B): {key_error:,.0f} 원\n"
                    f"= 최종 총 투사오류: {final_total:,.0f} 원"
                )
            )
            self._check_final_conclusion(final_total)
        except Exception as e:
            import traceback
            detail = traceback.format_exc()
            self.eval_result_display.config(text=f"⚠ 계산 중 오류 발생: {e}")
            messagebox.showerror("계산 오류", f"표본결과평가 계산 중 오류가 발생했습니다.\n\n{e}\n\n{detail[-500:]}")


    # ------------------------------------------------------------
    # ④-비통계적 표본결과평가
    # ------------------------------------------------------------
    def _build_nonstatistical_evaluation_section(self):
        f = self._make_section("④ 실제 선정 표본 및 테스트 결과")

        self.actual_sample_amount = LabeledEntry(
            f, "실제 선정된 표본 금액 (Sample Value)", 0, unit="원"
        )
        self.actual_sample_count = LabeledEntry(
            f, "실제 선정된 표본 건수 (Sample Count)", 1, unit="건"
        )
        self.actual_error_amount = LabeledEntry(
            f, "표본에서 실제 발견된 오차 금액", 2, unit="원"
        )

        tk.Label(f, text="오차 원인 및 성격평가 결과", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).grid(row=3, column=0, sticky="nw", padx=(10, 6), pady=6)
        self.error_cause_text = tk.Text(f, width=50, height=3, font=(FONT_NAME, 9))
        self.error_cause_text.grid(row=3, column=1, sticky="w", pady=6)

        tk.Button(f, text="투사오류 계산", command=self._calc_nonstatistical_evaluation,
                  bg=COLOR_BTN_SAVE, fg="white", font=(FONT_NAME, 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2"
                  ).grid(row=4, column=0, sticky="w", padx=(10, 6), pady=(10, 6))

        self.eval_result_display = tk.Label(
            f, text="", bg=f["bg"], fg=COLOR_TEXT_DARK, font=(FONT_NAME, 10),
            justify="left", anchor="w")
        self.eval_result_display.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 8))

    def _calc_nonstatistical_evaluation(self):
        remain_amt, _ = self.app._get_remaining()
        if remain_amt is None:
            msg = "1단계 A/B 항목을 먼저 입력해주세요."
            messagebox.showwarning("입력 필요", msg)
            self.eval_result_display.config(text=f"⚠ {msg}")
            return
        try:
            sample_value = parse_float(self.actual_sample_amount.get())
            actual_error = parse_float(self.actual_error_amount.get())
            if sample_value <= 0:
                msg = "투사오류 금액 계산을 위해 실제 선정된 표본 금액을 입력하세요."
                messagebox.showwarning("입력 필요", msg)
                self.eval_result_display.config(text=f"⚠ {msg}")
                return

            ratio_projected = actual_error / sample_value * remain_amt
            key_error = parse_float(self.app.key_actual_error.get())
            final_total = ratio_projected + key_error
            self.final_total_misstatement = final_total

            self.eval_result_display.config(
                text=(
                    f"비례식 투사오류(잔여모집단): {ratio_projected:,.0f} 원\n"
                    f"  (= 실제발견오차 {actual_error:,.0f} ÷ 표본금액 {sample_value:,.0f} × 잔여모집단 {remain_amt:,.0f})\n"
                    f"+ 특정항목 실제 발견 오차(1단계 B): {key_error:,.0f} 원\n"
                    f"= 최종 총 투사오류: {final_total:,.0f} 원"
                )
            )
            self._check_final_conclusion(final_total)
        except Exception as e:
            import traceback
            detail = traceback.format_exc()
            self.eval_result_display.config(text=f"⚠ 계산 중 오류 발생: {e}")
            messagebox.showerror("계산 오류", f"투사오류 계산 중 오류가 발생했습니다.\n\n{e}\n\n{detail[-500:]}")

    # ------------------------------------------------------------
    # ⑤ 최종 결론 (수행중요성 비교 및 경고)
    # ------------------------------------------------------------
    def _build_final_conclusion_section(self):
        f = self._make_section("⑤ 최종 결론")
        tk.Label(f, text="추가수행한 감사절차 (필요 시 기재)", bg=f["bg"], fg=COLOR_TEXT_DARK,
                  font=(FONT_NAME, 10)).pack(anchor="w", padx=10, pady=(6, 2))
        self.additional_procedure_text = tk.Text(f, width=70, height=4, font=(FONT_NAME, 9))
        self.additional_procedure_text.pack(anchor="w", padx=10, pady=(0, 8))

        tk.Button(f, text="결과 엑셀 저장", command=self._export_result_excel,
                  bg=COLOR_BTN_SAVE, fg="white", font=(FONT_NAME, 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2"
                  ).pack(anchor="w", padx=10, pady=(0, 10))

    def _check_final_conclusion(self, final_total):
        materiality = parse_float(self.app.materiality.get())
        self._additional_procedure_required = bool(materiality and final_total >= materiality)
        if self._additional_procedure_required:
            messagebox.showwarning(
                "경고",
                "최종 총 투사오류가 수행중요성 이상입니다.\n"
                "「추가수행한 감사절차」란에 실제 수행한 절차를 반드시 기재해야\n"
                "다음 단계(결과 엑셀 저장)로 진행할 수 있습니다."
            )

    # ------------------------------------------------------------
    # 결과 엑셀 저장 (기존 프로그램의 "적정성 검증 결과" 리포트 방식)
    # ------------------------------------------------------------
    def _export_result_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("오류", "openpyxl이 설치되어 있지 않습니다.")
            return
        if not hasattr(self, "final_total_misstatement"):
            messagebox.showwarning(
                "계산 필요",
                "먼저 표본결과평가를 계산한 뒤 저장해주세요."
            )
            return
        # 최종 총 투사오류가 수행중요성 이상인 경우, '추가수행한 감사절차'를
        # 반드시 입력해야만 결과 엑셀 저장(다음 단계)으로 진행할 수 있도록 차단한다.
        if getattr(self, "_additional_procedure_required", False):
            additional = self.additional_procedure_text.get("1.0", tk.END).strip()
            if not additional:
                messagebox.showerror(
                    "입력 필요",
                    "최종 총 투사오류가 수행중요성 이상입니다.\n"
                    "「추가수행한 감사절차」란에 실제 수행한 절차를 입력해야\n"
                    "결과 엑셀 저장이 가능합니다."
                )
                return

        default_name = f"{self.TITLE[self.mode].split('. ')[-1]}_적정성검증결과.xlsx"
        path = filedialog.asksaveasfilename(
            title="결과 엑셀 저장", defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")], initialfile=default_name,
        )
        if not path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "적정성검증결과"

            title_text = f"{self.TITLE[self.mode].split('. ')[-1]}(Sampling) 적정성 검증 결과"
            HEADER_FILL = PatternFill("solid", fgColor="153E6E")
            SECTION_FILL = PatternFill("solid", fgColor="B9C0C9")
            bold_white = Font(bold=True, color="FFFFFF", size=13)
            bold_dark = Font(bold=True, color="1B2733")
            normal = Font(color="1B2733")
            AMOUNT_FORMAT = "#,##0"  # 첫번째 이미지의 셀서식(숫자, 소수0자리, 1000단위 구분기호)

            row = 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            c = ws.cell(row=row, column=1, value=title_text)
            c.font = bold_white
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center")
            row += 1
            ws.cell(row=row, column=1, value="생성일시").font = normal
            ws.cell(row=row, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                     ).alignment = Alignment(horizontal="center")
            row += 2

            def section(title):
                nonlocal row
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
                cell = ws.cell(row=row, column=1, value=title)
                cell.font = bold_dark
                cell.fill = SECTION_FILL
                row += 1

            def line(label, value, is_amount=False, align=None):
                """label: A열 항목명 / value: B열 값
                is_amount=True면 숫자로 기록하고 첫번째 이미지 셀서식(#,##0) 적용하며
                기본적으로 우측정렬. False면 텍스트 그대로 기록하며 기본 가운데정렬.
                align을 명시하면 그 값으로 강제 지정 (예: '추가수행한 감사절차'=left).
                """
                nonlocal row
                if align is None:
                    align = "right" if is_amount else "center"
                ws.cell(row=row, column=1, value=label).font = normal
                v = ws.cell(row=row, column=2, value=value)
                v.font = normal
                v.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(align == "left"))
                if is_amount:
                    v.number_format = AMOUNT_FORMAT
                row += 1

            app = self.app
            remain_amt, remain_cnt = app._get_remaining()
            materiality = parse_float(app.materiality.get())

            section("① 기본정보")
            line("회사명", app.company_name.get())
            line("상장여부", app.listed_status.get())
            line("테스트대상항목", app.test_target.get())
            line("수행중요성(원)", materiality, is_amount=True)
            row += 1

            section("② 전체 모집단 / 특정항목 정보")
            line("테스트대상 총 금액(원)", parse_float(app.pop_amount.get()), is_amount=True)
            line("테스트대상 총 건수(건)", parse_float(app.pop_count.get()), is_amount=True)
            line("특정항목선정 금액(원)", parse_float(app.key_amount.get()), is_amount=True)
            line("특정항목선정 건수(건)", parse_float(app.key_count.get()), is_amount=True)
            line("특정항목 실제 발견 오차 금액(원)", parse_float(app.key_actual_error.get()), is_amount=True)
            line("잔여모집단 금액(원)", remain_amt, is_amount=True)
            line("잔여모집단 건수(건)", remain_cnt, is_amount=True)
            row += 1

            section("③ 위험요소 및 표본크기 산정 파라미터")
            line("고유위험계수(IR)", self.ir_var.get())
            line("통제위험계수(CR)", self.cr_var.get())
            line("분석절차계수(AP1)", self.ap1_var.get())
            if self.ap1_var.get() != "None":
                line("  - AP1 조서번호", self.ap1_workpaper.get())
                line("  - AP1 수행절차", self.ap1_procedure.get(), align="left")
            line("기타세부절차계수(AP2)", self.ap2_var.get())
            if self.ap2_var.get() != "None":
                line("  - AP2 조서번호", self.ap2_workpaper.get())
                line("  - AP2 수행절차", self.ap2_procedure.get(), align="left")
            line("TD", round(getattr(self, "_td", 0), 4))
            line("신뢰수준", round(getattr(self, "_confidence", 0), 4))
            line("예상왜곡표시율(%)", parse_float(self.expected_rate_entry.get()))
            # 예상왜곡표시집계액: "금액 + 원" 텍스트가 아니라 숫자만 우측(가운데정렬)컬럼에 기록
            expected_amount_val = parse_float(
                self.expected_amount_display["text"].replace("원", "").strip()
            )
            line("예상왜곡표시집계액(원)", expected_amount_val, is_amount=True)
            line("허용오류율", self.tolerable_rate_display["text"])
            line("모집단 성격(계층화 여부)", self.stratified_var.get())
            line("표본선택방법", self.selection_method_var.get())
            row += 1

            section("④ 최소 표본크기 산출결과")
            line("최소 표본크기(건)", getattr(self, "final_n", ""), is_amount=True)
            interval_val = getattr(self, "interval", None)
            if interval_val:
                line("표본간격(원)", interval_val, is_amount=True)
            else:
                line("표본간격", "해당없음 (체계적 선정이 아님)")
            row += 1

            section("⑤ 표본결과평가")
            if self.mode == "statistical":
                line("평가방식", "MUS 정식평가 (통계적)")
                line("오류상한(ULP, 잔여모집단)(원)", getattr(self, "ulp", 0), is_amount=True)
            else:
                line("평가방식", "비례식 투사 (비통계적)")
                line("실제 선정된 표본 금액(원)", parse_float(self.actual_sample_amount.get()), is_amount=True)
                line("실제 선정된 표본 건수(건)", parse_float(self.actual_sample_count.get()), is_amount=True)
                line("표본에서 실제 발견된 오차 금액(원)", parse_float(self.actual_error_amount.get()), is_amount=True)
            line("특정항목 실제 발견 오차(원)", parse_float(app.key_actual_error.get()), is_amount=True)
            line("최종 총 투사오류(원)", self.final_total_misstatement, is_amount=True)
            row += 1

            section("⑥ 최종 결론")
            exceeded = self.final_total_misstatement >= materiality if materiality else False
            line("수행중요성 초과 여부", "초과 (경고)" if exceeded else "이내 (정상)")
            additional = self.additional_procedure_text.get("1.0", tk.END).strip()
            line("추가수행한 감사절차", additional if additional else "(기재 없음)", align="left")

            ws.column_dimensions["A"].width = 34
            ws.column_dimensions["B"].width = 40

            wb.save(path)
            messagebox.showinfo("저장 완료", f"결과가 저장되었습니다.\n{path}")
        except Exception as e:
            import traceback
            messagebox.showerror("저장 실패", f"{e}\n\n{traceback.format_exc()[-500:]}")


if __name__ == "__main__":
    app = AuditSamplingApp()
    app.mainloop()
