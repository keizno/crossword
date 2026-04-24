#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
회사 교육용 크로스워드 퍼즐 생성기
Dark Theme / CustomTkinter GUI
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import random
import json
import os
import time
import copy
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io

# ── 한글 폰트 등록 ────────────────────────────────────────────────
def _register_korean_fonts():
    """나눔고딕 TTF를 ReportLab에 등록 (한글 PDF 출력용)"""
    candidates = [
        # 나눔고딕 (Ubuntu fonts-nanum 패키지)
        ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
         "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        # macOS
        ("/Library/Fonts/NanumGothic.ttf", None),
        # Windows
        ("C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/malgunbd.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont("KorFont",     regular))
                pdfmetrics.registerFont(TTFont("KorFont-Bold", bold if bold and os.path.exists(bold) else regular))
                from reportlab.lib.fonts import addMapping
                addMapping("KorFont", 0, 0, "KorFont")
                addMapping("KorFont", 1, 0, "KorFont-Bold")
                return True
            except Exception:
                pass
    return False

_KOR_FONT_OK = _register_korean_fonts()
KOR_FONT      = "KorFont"      if _KOR_FONT_OK else "Helvetica"
KOR_FONT_BOLD = "KorFont-Bold" if _KOR_FONT_OK else "Helvetica-Bold"

# ── Dark Theme 설정 ──────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── 색상 팔레트 ───────────────────────────────────────────────────
COLORS = {
    "bg":        "#1a1a2e",
    "panel":     "#16213e",
    "card":      "#0f3460",
    "accent":    "#e94560",
    "accent2":   "#533483",
    "gold":      "#f5a623",
    "correct":   "#2ecc71",
    "wrong":     "#e74c3c",
    "cell_bg":   "#1e2a45",
    "cell_text": "#e0e0e0",
    "cell_num":  "#a0c4ff",
    "border":    "#2d4a7a",
    "text":      "#e0e0e0",
    "subtext":   "#8899aa",
    "across":    "#1a4a7a",
    "down":      "#4a1a7a",
    "selected":  "#e94560",
    "highlight": "#2d6a9f",
    "empty":     "#0d1b2a",
}

# ════════════════════════════════════════════════════════════════
# 크로스워드 생성 엔진
# ════════════════════════════════════════════════════════════════
class CrosswordGenerator:
    def __init__(self, words_clues, grid_rows, grid_cols, max_attempts=50):
        self.words_clues = [(w.strip().upper(), c.strip()) for w, c in words_clues if 2 <= len(w.strip()) <= 15]
        self.rows = grid_rows
        self.cols = grid_cols
        self.max_attempts = max_attempts
        self.grid = []
        self.placed = []   # {word, clue, row, col, direction, number}

    def _empty_grid(self):
        return [['#'] * self.cols for _ in range(self.rows)]

    def _can_place(self, grid, word, r, c, direction):
        L = len(word)
        if direction == 'A':
            if c + L > self.cols: return False
            if c > 0 and grid[r][c-1] != '#': return False
            if c + L < self.cols and grid[r][c+L] != '#': return False
            for i, ch in enumerate(word):
                cell = grid[r][c+i]
                if cell != '#' and cell != ch: return False
                # 교차점 이외엔 같은 방향 인접 셀이 있으면 안 됨
                if cell == '#':
                    if r > 0 and grid[r-1][c+i] != '#': return False
                    if r < self.rows-1 and grid[r+1][c+i] != '#': return False
        else:  # 'D'
            if r + L > self.rows: return False
            if r > 0 and grid[r-1][c] != '#': return False
            if r + L < self.rows and grid[r+L][c] != '#': return False
            for i, ch in enumerate(word):
                cell = grid[r+i][c]
                if cell != '#' and cell != ch: return False
                if cell == '#':
                    if c > 0 and grid[r+i][c-1] != '#': return False
                    if c < self.cols-1 and grid[r+i][c+1] != '#': return False
        return True

    def _place_word(self, grid, word, r, c, direction):
        new_grid = copy.deepcopy(grid)
        if direction == 'A':
            for i, ch in enumerate(word):
                new_grid[r][c+i] = ch
        else:
            for i, ch in enumerate(word):
                new_grid[r+i][c] = ch
        return new_grid

    def _find_positions(self, grid, word, placed_count):
        positions = []
        # 첫 단어: 중앙 배치
        if placed_count == 0:
            r = self.rows // 2
            c = (self.cols - len(word)) // 2
            if self._can_place(grid, word, r, c, 'A'):
                positions.append((r, c, 'A', 10))
            r = (self.rows - len(word)) // 2
            c = self.cols // 2
            if self._can_place(grid, word, r, c, 'D'):
                positions.append((r, c, 'D', 10))
            return positions

        # 교차점 찾기
        for r in range(self.rows):
            for c in range(self.cols):
                if grid[r][c] == '#': continue
                ch = grid[r][c]
                for i, wch in enumerate(word):
                    if wch == ch:
                        # Across
                        tr, tc = r, c - i
                        if self._can_place(grid, word, tr, tc, 'A'):
                            positions.append((tr, tc, 'A', 5))
                        # Down
                        tr, tc = r - i, c
                        if self._can_place(grid, word, tr, tc, 'D'):
                            positions.append((tr, tc, 'D', 5))
        return positions

    def generate(self):
        best_grid, best_placed = None, []
        wc_list = list(self.words_clues)

        for attempt in range(self.max_attempts):
            random.shuffle(wc_list)
            grid = self._empty_grid()
            placed = []

            for word, clue in wc_list:
                positions = self._find_positions(grid, word, len(placed))
                if not positions: continue
                r, c, direction, _ = random.choice(positions)
                grid = self._place_word(grid, word, r, c, direction)
                placed.append({'word': word, 'clue': clue, 'row': r, 'col': c, 'direction': direction})

            if len(placed) > len(best_placed):
                best_placed = placed
                best_grid = copy.deepcopy(grid)
                if len(placed) >= min(len(wc_list), 12): break

        # 번호 부여
        if best_grid:
            self._assign_numbers(best_grid, best_placed)
        return best_grid, best_placed

    def _assign_numbers(self, grid, placed):
        numbered = {}
        num = 1
        for r in range(self.rows):
            for c in range(self.cols):
                if grid[r][c] == '#': continue
                starts = False
                for p in placed:
                    if p['row'] == r and p['col'] == c:
                        starts = True
                        break
                if starts:
                    numbered[(r, c)] = num
                    for p in placed:
                        if p['row'] == r and p['col'] == c:
                            p['number'] = num
                    num += 1
        self.numbered = numbered

# ════════════════════════════════════════════════════════════════
# 메인 앱
# ════════════════════════════════════════════════════════════════
class CrosswordApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🔤 크로스워드 퍼즐 메이커 Pro")
        self.geometry("1400x900")
        self.minsize(1100, 750)
        self.configure(fg_color=COLORS["bg"])

        # State
        self.words_data = []          # [(word, clue, difficulty)]
        self.current_grid = None
        self.current_placed = []
        self.numbered = {}            # (r,c) -> number
        self.cell_widgets = {}        # (r,c) -> Entry widget
        self.cell_state = {}          # (r,c) -> 'empty'|'correct'|'wrong'
        self.grid_rows = 15
        self.grid_cols = 15
        self.difficulty_filter = "전체"
        self.selected_direction = 'A'
        self.selected_word_idx = -1
        self.timer_running = False
        self.start_time = 0
        self.elapsed = 0
        self.score = 0
        self.hints_used = 0

        self._build_ui()

    # ── UI 구성 ────────────────────────────────────────────────
    def _build_ui(self):
        # 상단 타이틀 바
        self._build_titlebar()
        # 메인 컨테이너 (좌/중/우)
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=(0,10))
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        self._build_left_panel(main)
        self._build_center_panel(main)
        self._build_right_panel(main)

    def _build_titlebar(self):
        bar = ctk.CTkFrame(self, fg_color=COLORS["panel"], height=60, corner_radius=0)
        bar.pack(fill="x", pady=(0,2))
        bar.pack_propagate(False)

        ctk.CTkLabel(bar, text="🔤  크로스워드 퍼즐 메이커  Pro",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=COLORS["gold"]).pack(side="left", padx=20)

        # 우측 버튼들
        btn_frame = ctk.CTkFrame(bar, fg_color="transparent")
        btn_frame.pack(side="right", padx=10)

        self._icon_btn(btn_frame, "🖨️ 인쇄", self._print_puzzle).pack(side="right", padx=4)
        self._icon_btn(btn_frame, "🖼️ 이미지 저장", self._save_image).pack(side="right", padx=4)
        self._icon_btn(btn_frame, "📄 PDF 저장", self._export_pdf).pack(side="right", padx=4)
        self._icon_btn(btn_frame, "💾 저장", self._save_puzzle).pack(side="right", padx=4)
        self._icon_btn(btn_frame, "📂 불러오기", self._load_puzzle).pack(side="right", padx=4)

    def _icon_btn(self, parent, text, cmd, color=None):
        return ctk.CTkButton(parent, text=text, command=cmd,
                             width=110, height=32,
                             fg_color=color or COLORS["card"],
                             hover_color=COLORS["accent2"],
                             font=ctk.CTkFont(size=12))

    # ── 좌측 패널 ─────────────────────────────────────────────
    def _build_left_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color=COLORS["panel"], width=290, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,6), pady=0)
        left.pack_propagate(False)

        # 엑셀 로드
        sec = self._section(left, "📁 데이터 로드")
        self._icon_btn(sec, "📊 엑셀 파일 열기", self._load_excel,
                       color=COLORS["accent"]).pack(fill="x", pady=4)
        self.file_label = ctk.CTkLabel(sec, text="파일 미선택",
                                        text_color=COLORS["subtext"],
                                        font=ctk.CTkFont(size=11), wraplength=240)
        self.file_label.pack(pady=2)

        self.word_count_label = ctk.CTkLabel(sec, text="단어: 0개",
                                              font=ctk.CTkFont(size=12, weight="bold"),
                                              text_color=COLORS["gold"])
        self.word_count_label.pack(pady=2)

        # 난이도
        sec2 = self._section(left, "🎯 난이도 필터")
        self.diff_var = ctk.StringVar(value="전체")
        # 2×2 그리드로 배치해서 짤림 방지
        diff_grid = ctk.CTkFrame(sec2, fg_color="transparent")
        diff_grid.pack(fill="x", padx=4, pady=(2,6))
        diff_items = [("전체", "#2d4a7a"), ("하 ✅", "#1a5c2a"), ("중 ⚡", "#5c4a1a"), ("상 🔥", "#5c1a1a")]
        diff_vals  = ["전체", "하", "중", "상"]
        for idx, ((label, _), val) in enumerate(zip(diff_items, diff_vals)):
            rb = ctk.CTkRadioButton(diff_grid, text=label, variable=self.diff_var,
                                    value=val, command=self._on_diff_change,
                                    fg_color=COLORS["accent"],
                                    font=ctk.CTkFont(size=13))
            rb.grid(row=idx//2, column=idx%2, sticky="w", padx=8, pady=3)

        # 그리드 설정
        sec3 = self._section(left, "⚙️ 그리드 설정")
        row_frame = ctk.CTkFrame(sec3, fg_color="transparent")
        row_frame.pack(fill="x", pady=2)
        ctk.CTkLabel(row_frame, text="행:", width=40).pack(side="left")
        self.row_var = ctk.IntVar(value=15)
        ctk.CTkSlider(row_frame, from_=8, to=25, variable=self.row_var,
                      number_of_steps=17, width=140,
                      command=lambda v: self.row_val.configure(text=f"{int(v)}")).pack(side="left", padx=4)
        self.row_val = ctk.CTkLabel(row_frame, text="15", width=30)
        self.row_val.pack(side="left")

        col_frame = ctk.CTkFrame(sec3, fg_color="transparent")
        col_frame.pack(fill="x", pady=2)
        ctk.CTkLabel(col_frame, text="열:", width=40).pack(side="left")
        self.col_var = ctk.IntVar(value=15)
        ctk.CTkSlider(col_frame, from_=8, to=25, variable=self.col_var,
                      number_of_steps=17, width=140,
                      command=lambda v: self.col_val.configure(text=f"{int(v)}")).pack(side="left", padx=4)
        self.col_val = ctk.CTkLabel(col_frame, text="15", width=30)
        self.col_val.pack(side="left")

        # 단어 길이 필터
        len_frame = ctk.CTkFrame(sec3, fg_color="transparent")
        len_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(len_frame, text="단어 길이:").pack(side="left")
        self.min_len = ctk.IntVar(value=2)
        self.max_len = ctk.IntVar(value=7)
        ctk.CTkLabel(len_frame, text="  최소").pack(side="left")
        ctk.CTkOptionMenu(len_frame, variable=self.min_len,
                          values=[str(i) for i in range(2,9)], width=55).pack(side="left", padx=2)
        ctk.CTkLabel(len_frame, text="최대").pack(side="left")
        ctk.CTkOptionMenu(len_frame, variable=self.max_len,
                          values=[str(i) for i in range(2,16)], width=55).pack(side="left", padx=2)

        # 생성 버튼
        ctk.CTkButton(left, text="🎲  퍼즐 생성!", command=self._generate_puzzle,
                      height=44, fg_color=COLORS["accent"],
                      hover_color="#c0392b",
                      font=ctk.CTkFont(size=16, weight="bold"),
                      corner_radius=10).pack(fill="x", padx=12, pady=8)

        # 단어 목록 (접기/펼치기 + 정답 숨김)
        sec4_header = ctk.CTkFrame(left, fg_color=COLORS["card"], corner_radius=8)
        sec4_header.pack(fill="x", padx=10, pady=5)
        sec4_top = ctk.CTkFrame(sec4_header, fg_color="transparent")
        sec4_top.pack(fill="x")

        ctk.CTkLabel(sec4_top, text="📋 단어 목록",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["cell_num"]).pack(side="left", padx=8, pady=(6,2))

        self.word_list_visible = True
        self.show_answers_var = ctk.BooleanVar(value=True)

        toggle_frame = ctk.CTkFrame(sec4_top, fg_color="transparent")
        toggle_frame.pack(side="right", padx=6, pady=2)

        ctk.CTkCheckBox(toggle_frame, text="정답숨김", variable=self.show_answers_var,
                        onvalue=True, offvalue=False,
                        command=self._update_word_list,
                        font=ctk.CTkFont(size=10),
                        checkbox_width=16, checkbox_height=16,
                        fg_color=COLORS["accent"]).pack(side="left", padx=4)

        self.toggle_btn = ctk.CTkButton(toggle_frame, text="▲ 접기",
                                         command=self._toggle_word_list,
                                         width=60, height=22,
                                         fg_color=COLORS["accent2"],
                                         font=ctk.CTkFont(size=11))
        self.toggle_btn.pack(side="left", padx=2)

        self.word_list_container = ctk.CTkFrame(sec4_header, fg_color="transparent")
        self.word_list_container.pack(fill="both", expand=True)

        self.word_list_box = ctk.CTkTextbox(self.word_list_container, height=180,
                                             fg_color=COLORS["empty"],
                                             text_color=COLORS["cell_text"],
                                             font=ctk.CTkFont(size=11))
        self.word_list_box.pack(fill="both", expand=True, padx=6, pady=(0,6))

    def _toggle_word_list(self):
        if self.word_list_visible:
            self.word_list_container.pack_forget()
            self.toggle_btn.configure(text="▼ 펼치기")
            self.word_list_visible = False
        else:
            self.word_list_container.pack(fill="both", expand=True)
            self.toggle_btn.configure(text="▲ 접기")
            self.word_list_visible = True

    # ── 중앙 패널 (퍼즐 그리드) ───────────────────────────────
    def _build_center_panel(self, parent):
        center = ctk.CTkFrame(parent, fg_color=COLORS["bg"], corner_radius=0)
        center.grid(row=0, column=1, sticky="nsew", padx=3)
        center.rowconfigure(1, weight=1)
        center.columnconfigure(0, weight=1)

        # 상태 바
        status = ctk.CTkFrame(center, fg_color=COLORS["panel"], height=45, corner_radius=8)
        status.grid(row=0, column=0, sticky="ew", pady=(0,6))
        status.pack_propagate(False)
        status.columnconfigure((0,1,2,3,4), weight=1)

        self.timer_label = ctk.CTkLabel(status, text="⏱️  00:00",
                                         font=ctk.CTkFont(size=16, weight="bold"),
                                         text_color=COLORS["gold"])
        self.timer_label.grid(row=0, column=0, padx=10)

        self.score_label = ctk.CTkLabel(status, text="🏆  0점",
                                         font=ctk.CTkFont(size=16, weight="bold"),
                                         text_color=COLORS["correct"])
        self.score_label.grid(row=0, column=1, padx=10)

        self.progress_label = ctk.CTkLabel(status, text="📊  0/0",
                                            font=ctk.CTkFont(size=14),
                                            text_color=COLORS["text"])
        self.progress_label.grid(row=0, column=2, padx=10)

        self.hint_label = ctk.CTkLabel(status, text="💡  힌트: 0",
                                        font=ctk.CTkFont(size=14),
                                        text_color=COLORS["subtext"])
        self.hint_label.grid(row=0, column=3, padx=10)

        btn_bar = ctk.CTkFrame(status, fg_color="transparent")
        btn_bar.grid(row=0, column=4, padx=8)
        ctk.CTkButton(btn_bar, text="✅ 채점", command=self._check_answers,
                       width=72, height=28, fg_color=COLORS["correct"]).pack(side="left", padx=2)
        ctk.CTkButton(btn_bar, text="💡 힌트", command=self._give_hint,
                       width=72, height=28, fg_color=COLORS["gold"],
                       text_color="#000").pack(side="left", padx=2)
        ctk.CTkButton(btn_bar, text="🔄 초기화", command=self._reset_puzzle,
                       width=72, height=28, fg_color=COLORS["card"]).pack(side="left", padx=2)
        ctk.CTkButton(btn_bar, text="👁️ 정답", command=self._reveal_all,
                       width=72, height=28, fg_color=COLORS["accent2"]).pack(side="left", padx=2)

        # 퍼즐 캔버스 영역
        canvas_outer = ctk.CTkFrame(center, fg_color=COLORS["panel"], corner_radius=10)
        canvas_outer.grid(row=1, column=0, sticky="nsew")
        canvas_outer.rowconfigure(0, weight=1)
        canvas_outer.columnconfigure(0, weight=1)

        self.canvas_frame = ctk.CTkScrollableFrame(canvas_outer,
                                                    fg_color=COLORS["empty"],
                                                    corner_radius=8)
        self.canvas_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        self.puzzle_placeholder = ctk.CTkLabel(
            self.canvas_frame,
            text="퍼즐을 생성하려면\n왼쪽에서 엑셀 파일을 열고\n'퍼즐 생성!' 버튼을 클릭하세요.",
            font=ctk.CTkFont(size=18),
            text_color=COLORS["subtext"]
        )
        self.puzzle_placeholder.pack(expand=True, pady=80)

    # ── 우측 패널 (단서 목록) ────────────────────────────────
    def _build_right_panel(self, parent):
        right = ctk.CTkFrame(parent, fg_color=COLORS["panel"], width=280, corner_radius=12)
        right.grid(row=0, column=2, sticky="nsew", padx=(6,0))
        right.pack_propagate(False)

        ctk.CTkLabel(right, text="📖 단서",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=COLORS["gold"]).pack(pady=(12,4))

        # Across / Down 탭
        tab = ctk.CTkTabview(right, fg_color=COLORS["card"],
                              segmented_button_fg_color=COLORS["empty"],
                              segmented_button_selected_color=COLORS["accent"])
        tab.pack(fill="both", expand=True, padx=8, pady=4)
        self.across_tab = tab.add("→ 가로")
        self.down_tab   = tab.add("↓ 세로")

        self.across_clue_frame = ctk.CTkScrollableFrame(self.across_tab,
                                                         fg_color="transparent")
        self.across_clue_frame.pack(fill="both", expand=True)

        self.down_clue_frame = ctk.CTkScrollableFrame(self.down_tab,
                                                       fg_color="transparent")
        self.down_clue_frame.pack(fill="both", expand=True)

        # 정답 현황
        sec = self._section(right, "📊 현황")
        self.progress_bar = ctk.CTkProgressBar(sec, fg_color=COLORS["empty"],
                                                progress_color=COLORS["correct"])
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=4)
        self.progress_detail = ctk.CTkLabel(sec, text="아직 시작 전",
                                             text_color=COLORS["subtext"],
                                             font=ctk.CTkFont(size=12))
        self.progress_detail.pack()

        # 통계
        sec2 = self._section(right, "📈 통계")
        self.stats_label = ctk.CTkLabel(sec2, text="퍼즐을 생성해 주세요",
                                         text_color=COLORS["subtext"],
                                         font=ctk.CTkFont(size=11),
                                         justify="left")
        self.stats_label.pack(anchor="w", padx=4)

    # ── 헬퍼 ─────────────────────────────────────────────────
    def _section(self, parent, title):
        frame = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=8)
        frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame, text=title,
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["cell_num"]).pack(anchor="w", padx=8, pady=(6,2))
        return frame

    # ════════════════════════════════════════════════════════
    # 엑셀 로드
    # ════════════════════════════════════════════════════════
    def _load_excel(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path: return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            self.words_data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0] or not row[1]: continue
                clue = str(row[0]).strip()
                word = str(row[1]).strip().upper()
                diff = str(row[2]).strip() if len(row) > 2 and row[2] else "중"
                if word.isalpha():
                    self.words_data.append((word, clue, diff))

            fname = os.path.basename(path)
            self.file_label.configure(text=fname)
            self.word_count_label.configure(text=f"단어: {len(self.words_data)}개")
            self._update_word_list()
            messagebox.showinfo("로드 완료", f"총 {len(self.words_data)}개 단어를 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")

    def _update_word_list(self):
        self.word_list_box.configure(state="normal")
        self.word_list_box.delete("1.0", "end")
        filtered = self._get_filtered_words()
        hide = self.show_answers_var.get()
        for w, c, d in filtered:
            answer = "?" * len(w) if hide else w
            self.word_list_box.insert("end", f"[{d}] {answer} - {c}\n")
        self.word_list_box.configure(state="disabled")

    def _get_filtered_words(self):
        diff = self.diff_var.get()
        min_l = self.min_len.get()
        max_l = self.max_len.get()
        result = []
        for w, c, d in self.words_data:
            if diff != "전체" and d != diff: continue
            if not (min_l <= len(w) <= max_l): continue
            result.append((w, c, d))
        return result

    def _on_diff_change(self):
        self._update_word_list()

    # ════════════════════════════════════════════════════════
    # 퍼즐 생성
    # ════════════════════════════════════════════════════════
    def _generate_puzzle(self):
        filtered = self._get_filtered_words()
        if len(filtered) < 3:
            messagebox.showwarning("단어 부족",
                "최소 3개 이상의 단어가 필요합니다.\n엑셀 파일을 로드하거나 난이도 필터를 확인하세요.")
            return

        self.grid_rows = self.row_var.get()
        self.grid_cols = self.col_var.get()

        wc_list = [(w, c) for w, c, _ in filtered]
        gen = CrosswordGenerator(wc_list, self.grid_rows, self.grid_cols, max_attempts=60)

        # 생성 중 메시지
        self.puzzle_placeholder.configure(text="⏳ 퍼즐 생성 중...")
        self.update()

        grid, placed = gen.generate()
        self.numbered = gen.numbered if hasattr(gen, 'numbered') else {}

        if not grid or not placed:
            messagebox.showwarning("생성 실패", "단어 배치에 실패했습니다. 그리드 크기를 늘리거나 다시 시도하세요.")
            self.puzzle_placeholder.configure(text="생성 실패 - 다시 시도해주세요")
            return

        self.current_grid = grid
        self.current_placed = placed
        self.cell_state = {}
        self.hints_used = 0
        self.score = 0

        self._draw_puzzle()
        self._build_clue_list()
        self._update_stats()
        self._start_timer()

    # ════════════════════════════════════════════════════════
    # 그리드 그리기
    # ════════════════════════════════════════════════════════
    def _draw_puzzle(self):
        # 기존 위젯 제거
        for w in self.canvas_frame.winfo_children():
            w.destroy()
        self.cell_widgets = {}

        if not self.current_grid: return

        CELL_SIZE = max(34, min(52, 640 // max(self.grid_rows, self.grid_cols)))
        NUM_FONT_SIZE = max(7, CELL_SIZE // 4)
        LETTER_FONT_SIZE = max(11, int(CELL_SIZE * 0.45))

        outer = tk.Frame(self.canvas_frame, bg=COLORS["empty"])
        outer.pack(expand=True, padx=10, pady=10)

        for r in range(self.grid_rows):
            for c in range(self.grid_cols):
                cell_val = self.current_grid[r][c]
                if cell_val == '#':
                    lbl = tk.Frame(outer, width=CELL_SIZE, height=CELL_SIZE,
                                   bg=COLORS["empty"], bd=0)
                    lbl.grid(row=r, column=c, padx=1, pady=1)
                    lbl.grid_propagate(False)
                else:
                    # 셀 컨테이너 (Canvas 기반 → 번호+입력 겹침 없이 깔끔하게)
                    cell = tk.Canvas(outer, width=CELL_SIZE, height=CELL_SIZE,
                                     bg=COLORS["cell_bg"],
                                     bd=0, highlightthickness=1,
                                     highlightbackground=COLORS["border"])
                    cell.grid(row=r, column=c, padx=1, pady=1)

                    # ── 셀 번호 (좌상단, Canvas text로 그려서 항상 보임) ──
                    if (r, c) in self.numbered:
                        cell.create_text(
                            2, 1,
                            text=str(self.numbered[(r, c)]),
                            anchor="nw",
                            font=("Arial", NUM_FONT_SIZE, "bold"),
                            fill="#a0c4ff",
                            tags="num"
                        )

                    # ── 입력 Entry (번호 아래쪽 공간 차지) ──
                    num_offset = NUM_FONT_SIZE + 2 if (r, c) in self.numbered else 0
                    var = tk.StringVar()
                    entry = tk.Entry(cell, textvariable=var,
                                     font=("Arial", LETTER_FONT_SIZE, "bold"),
                                     bg=COLORS["cell_bg"],
                                     fg=COLORS["cell_text"],
                                     insertbackground=COLORS["accent"],
                                     bd=0, highlightthickness=0,
                                     justify="center",
                                     width=1)
                    # Entry를 번호 아래에 배치 (번호 없는 셀은 중앙)
                    ey = (CELL_SIZE + num_offset) // 2
                    cell.create_window(CELL_SIZE // 2, ey,
                                       window=entry,
                                       width=CELL_SIZE - 4,
                                       height=CELL_SIZE - num_offset - 4)

                    entry.bind("<KeyRelease>", lambda e, row=r, col=c, v=var: self._on_key(e, row, col, v))
                    entry.bind("<FocusIn>",    lambda e, row=r, col=c: self._on_focus(row, col))
                    var.trace_add("write", lambda *a, row=r, col=c, v=var: self._on_entry_change(row, col, v))

                    self.cell_widgets[(r, c)] = (entry, var, cell)

    def _on_key(self, event, row, col, var):
        val = var.get().upper()
        if len(val) > 1:
            var.set(val[-1])
        elif val and val.isalpha():
            var.set(val)
            self._move_next(row, col)
        elif event.keysym == 'BackSpace':
            var.set('')
        elif event.keysym in ('Left', 'Right', 'Up', 'Down'):
            self._navigate(row, col, event.keysym)

    def _move_next(self, row, col):
        # 현재 선택 방향으로 이동
        if self.selected_direction == 'A':
            next_cell = (row, col+1)
        else:
            next_cell = (row+1, col)
        if next_cell in self.cell_widgets:
            self.cell_widgets[next_cell][0].focus_set()

    def _navigate(self, row, col, key):
        moves = {'Left': (0,-1), 'Right': (0,1), 'Up': (-1,0), 'Down': (1,0)}
        dr, dc = moves.get(key, (0,0))
        target = (row+dr, col+dc)
        if target in self.cell_widgets:
            self.cell_widgets[target][0].focus_set()

    def _on_focus(self, row, col):
        self._highlight_word(row, col)

    def _on_entry_change(self, row, col, var):
        self._update_progress()

    def _set_cell_bg(self, r, c, color):
        """Canvas 셀과 그 안의 Entry 배경색을 함께 변경"""
        if (r, c) not in self.cell_widgets: return
        entry, var, canvas = self.cell_widgets[(r, c)]
        canvas.configure(bg=color)
        try: canvas.itemconfigure("num", fill="#a0c4ff")
        except: pass
        entry.configure(bg=color)

    def _highlight_word(self, row, col):
        # 모든 셀 초기화
        for (r, c) in self.cell_widgets:
            self._set_cell_bg(r, c, COLORS["cell_bg"])

        # 해당 셀이 포함된 단어 하이라이트
        for p in self.current_placed:
            word = p['word']
            wr, wc, wd = p['row'], p['col'], p['direction']
            cells = [(wr, wc+i) if wd == 'A' else (wr+i, wc) for i in range(len(word))]
            if (row, col) in cells:
                self.selected_direction = wd
                self.selected_word_idx = self.current_placed.index(p)
                for cr, cc in cells:
                    self._set_cell_bg(cr, cc, COLORS["highlight"])
                # 현재 셀은 더 강하게
                self._set_cell_bg(row, col, "#e94580")
                break

    # ════════════════════════════════════════════════════════
    # 단서 목록
    # ════════════════════════════════════════════════════════
    def _build_clue_list(self):
        for w in self.across_clue_frame.winfo_children(): w.destroy()
        for w in self.down_clue_frame.winfo_children(): w.destroy()

        across = sorted([p for p in self.current_placed if p['direction'] == 'A'], key=lambda x: x.get('number', 99))
        down   = sorted([p for p in self.current_placed if p['direction'] == 'D'], key=lambda x: x.get('number', 99))

        for p in across:
            self._clue_item(self.across_clue_frame, p)
        for p in down:
            self._clue_item(self.down_clue_frame, p)

    def _clue_item(self, parent, p):
        num = p.get('number', '?')
        word = p['word']
        clue = p['clue']
        direction = p['direction']
        length = len(word)

        frame = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=6)
        frame.pack(fill="x", pady=2, padx=2)
        frame.bind("<Button-1>", lambda e, pp=p: self._jump_to_word(pp))

        color = COLORS["across"] if direction == 'A' else COLORS["down"]
        ctk.CTkLabel(frame,
                     text=f"{'→' if direction=='A' else '↓'} {num}",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["gold"],
                     width=36).pack(side="left", padx=(6,2), pady=4)

        ctk.CTkLabel(frame,
                     text=f"{clue}  ({length}글자)",
                     font=ctk.CTkFont(size=12),
                     text_color=COLORS["text"],
                     anchor="w",
                     wraplength=180).pack(side="left", fill="x", expand=True, padx=4, pady=4)

    def _jump_to_word(self, p):
        wr, wc = p['row'], p['col']
        if (wr, wc) in self.cell_widgets:
            self.cell_widgets[(wr, wc)][0].focus_set()

    # ════════════════════════════════════════════════════════
    # 채점 / 힌트 / 초기화
    # ════════════════════════════════════════════════════════
    def _check_answers(self):
        if not self.current_placed: return
        correct_words = 0
        total_words = len(self.current_placed)

        for p in self.current_placed:
            word = p['word']
            wr, wc, wd = p['row'], p['col'], p['direction']
            user_word = ""
            cells = []
            for i in range(len(word)):
                cr, cc = (wr, wc+i) if wd == 'A' else (wr+i, wc)
                cells.append((cr, cc))
                if (cr, cc) in self.cell_widgets:
                    v = self.cell_widgets[(cr, cc)][1].get().upper()
                    user_word += v
                else:
                    user_word += '?'

            is_correct = user_word == word
            if is_correct:
                correct_words += 1

            for i, (cr, cc) in enumerate(cells):
                bg = COLORS["correct"] if is_correct else COLORS["wrong"]
                self._set_cell_bg(cr, cc, bg)

        self._stop_timer()
        base_score = correct_words * 10
        time_bonus = max(0, 300 - self.elapsed) // 10
        hint_penalty = self.hints_used * 5
        self.score = max(0, base_score + time_bonus - hint_penalty)
        self.score_label.configure(text=f"🏆  {self.score}점")

        # 완료 팝업
        elapsed_str = self._fmt_time(self.elapsed)
        msg = (f"✅ 채점 완료!\n\n"
               f"정답 단어: {correct_words}/{total_words}\n"
               f"소요 시간: {elapsed_str}\n"
               f"힌트 사용: {self.hints_used}회\n\n"
               f"🏆 최종 점수: {self.score}점")
        messagebox.showinfo("채점 결과", msg)
        self._update_progress()

    def _give_hint(self):
        if not self.current_placed: return

        # 커서가 있는 단어 우선, 없으면 전체에서 첫 번째 빈 셀
        idx = self.selected_word_idx
        if 0 <= idx < len(self.current_placed):
            ordered = [self.current_placed[idx]] + \
                      [p for i, p in enumerate(self.current_placed) if i != idx]
        else:
            ordered = self.current_placed

        for p in ordered:
            word = p['word']
            wr, wc, wd = p['row'], p['col'], p['direction']
            for i, ch in enumerate(word):
                cr, cc = (wr, wc+i) if wd == 'A' else (wr+i, wc)
                if (cr, cc) in self.cell_widgets:
                    v = self.cell_widgets[(cr, cc)][1].get().upper()
                    if v != ch:
                        self.cell_widgets[(cr, cc)][1].set(ch)
                        self._set_cell_bg(cr, cc, COLORS["gold"])
                        self.hints_used += 1
                        self.hint_label.configure(text=f"💡  힌트: {self.hints_used}")
                        return
        messagebox.showinfo("힌트", "모든 칸이 이미 채워졌습니다!")

    def _reset_puzzle(self):
        for (r, c) in list(self.cell_widgets.keys()):
            self.cell_widgets[(r, c)][1].set('')
            self._set_cell_bg(r, c, COLORS["cell_bg"])
        self.hints_used = 0
        self.hint_label.configure(text="💡  힌트: 0")
        self.score = 0
        self.score_label.configure(text="🏆  0점")
        self._start_timer()
        self._update_progress()

    def _reveal_all(self):
        if not messagebox.askyesno("정답 공개", "모든 정답을 보시겠습니까?\n(점수가 0점이 됩니다)"):
            return
        for p in self.current_placed:
            word = p['word']
            wr, wc, wd = p['row'], p['col'], p['direction']
            for i, ch in enumerate(word):
                cr, cc = (wr, wc+i) if wd == 'A' else (wr+i, wc)
                if (cr, cc) in self.cell_widgets:
                    self.cell_widgets[(cr, cc)][1].set(ch)
                    self._set_cell_bg(cr, cc, COLORS["accent2"])
        self.score = 0
        self.score_label.configure(text="🏆  0점")
        self._stop_timer()

    # ════════════════════════════════════════════════════════
    # 타이머
    # ════════════════════════════════════════════════════════
    def _start_timer(self):
        self.timer_running = True
        self.start_time = time.time()
        self.elapsed = 0
        self._tick()

    def _stop_timer(self):
        self.timer_running = False

    def _tick(self):
        if not self.timer_running: return
        self.elapsed = int(time.time() - self.start_time)
        self.timer_label.configure(text=f"⏱️  {self._fmt_time(self.elapsed)}")
        self.after(1000, self._tick)

    def _fmt_time(self, secs):
        return f"{secs//60:02d}:{secs%60:02d}"

    # ════════════════════════════════════════════════════════
    # 진행 상황 업데이트
    # ════════════════════════════════════════════════════════
    def _update_progress(self):
        if not self.current_placed: return
        total_cells = sum(len(p['word']) for p in self.current_placed)
        filled = 0
        for (r, c), (entry, var, cell) in self.cell_widgets.items():
            if var.get().strip(): filled += 1

        pct = filled / total_cells if total_cells else 0
        self.progress_bar.set(pct)
        self.progress_label.configure(text=f"📊  {filled}/{total_cells} 칸")
        self.progress_detail.configure(text=f"완성률 {pct*100:.0f}%")

    def _update_stats(self):
        total = len(self.current_placed)
        across = sum(1 for p in self.current_placed if p['direction'] == 'A')
        down   = total - across
        total_cells = sum(len(p['word']) for p in self.current_placed)
        words = [p['word'] for p in self.current_placed]
        avg_len = sum(len(w) for w in words) / len(words) if words else 0

        self.stats_label.configure(
            text=(f"가로 단어: {across}개\n"
                  f"세로 단어: {down}개\n"
                  f"총 단어: {total}개\n"
                  f"총 칸 수: {total_cells}개\n"
                  f"평균 단어 길이: {avg_len:.1f}자")
        )
        self.progress_label.configure(text=f"📊  0/{total_cells}")

    # ════════════════════════════════════════════════════════
    # 저장 / 불러오기
    # ════════════════════════════════════════════════════════
    def _save_puzzle(self):
        if not self.current_grid:
            messagebox.showwarning("저장 실패", "생성된 퍼즐이 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="퍼즐 저장"
        )
        if not path: return
        data = {
            "grid": self.current_grid,
            "placed": self.current_placed,
            "numbered": {f"{k[0]},{k[1]}": v for k, v in self.numbered.items()},
            "rows": self.grid_rows,
            "cols": self.grid_cols,
            "words_data": [{"word": w, "clue": c, "diff": d} for w, c, d in self.words_data],
        }
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("저장 완료", f"퍼즐이 저장되었습니다:\n{path}")

    def _load_puzzle(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="퍼즐 불러오기"
        )
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.current_grid = data["grid"]
            self.current_placed = data["placed"]
            self.numbered = {tuple(int(x) for x in k.split(',')): v
                             for k, v in data.get("numbered", {}).items()}
            self.grid_rows = data["rows"]
            self.grid_cols = data["cols"]
            # words_data 복원 (구버전 파일에는 없을 수 있으므로 get 사용)
            raw_words = data.get("words_data", [])
            if raw_words:
                self.words_data = [(item["word"], item["clue"], item.get("diff", "중"))
                                   for item in raw_words]
                fname = os.path.basename(path)
                self.file_label.configure(text=f"[JSON] {fname}")
                self.word_count_label.configure(text=f"단어: {len(self.words_data)}개")
                self._update_word_list()
            self.cell_state = {}
            self.hints_used = 0
            self._draw_puzzle()
            self._build_clue_list()
            self._update_stats()
            self._start_timer()
            messagebox.showinfo("불러오기 완료", "퍼즐을 불러왔습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"불러오기 실패:\n{e}")

    # ════════════════════════════════════════════════════════
    # 이미지 저장
    # ════════════════════════════════════════════════════════
    def _save_image(self):
        if not self.current_grid:
            messagebox.showwarning("저장 실패", "생성된 퍼즐이 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg")],
            title="이미지로 저장"
        )
        if not path: return
        img = self._render_puzzle_image()
        img.save(path)
        messagebox.showinfo("저장 완료", f"이미지가 저장되었습니다:\n{path}")

    def _render_puzzle_image(self, cell_size=40, show_answers=False):
        rows, cols = self.grid_rows, self.grid_cols
        margin = 20
        clue_width = 320
        W = cols * cell_size + 2 * margin + clue_width
        H = rows * cell_size + 2 * margin + 60

        img = Image.new('RGB', (W, H), color=(26, 26, 46))
        draw = ImageDraw.Draw(img)

        try:
            kor_ttf    = "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"
            kor_bold   = "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"
            font_big   = ImageFont.truetype(kor_bold, 16)
            font_small = ImageFont.truetype(kor_ttf,  11)
            font_num   = ImageFont.truetype(kor_ttf,  9)
            font_title = ImageFont.truetype(kor_bold, 20)
        except:
            try:
                font_big   = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
                font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 10)
                font_num   = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 9)
                font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
            except:
                font_big = font_small = font_num = font_title = ImageFont.load_default()

        # 제목
        draw.text((margin, 8), "CROSSWORD PUZZLE", fill=(245, 166, 35), font=font_title)

        # 그리드
        for r in range(rows):
            for c in range(cols):
                x = margin + c * cell_size
                y = 50 + r * cell_size
                val = self.current_grid[r][c]
                if val == '#':
                    draw.rectangle([x, y, x+cell_size, y+cell_size], fill=(13, 27, 42))
                else:
                    draw.rectangle([x, y, x+cell_size, y+cell_size],
                                   fill=(30, 42, 69), outline=(45, 74, 122), width=1)
                    if (r, c) in self.numbered:
                        draw.text((x+2, y+1), str(self.numbered[(r,c)]),
                                  fill=(160, 196, 255), font=font_num)
                    if show_answers:
                        draw.text((x + cell_size//2 - 5, y + cell_size//2 - 8),
                                  val, fill=(224, 224, 224), font=font_big)
                    else:
                        # 현재 입력값 표시
                        if (r, c) in self.cell_widgets:
                            user_val = self.cell_widgets[(r, c)][1].get().upper()
                            if user_val:
                                draw.text((x + cell_size//2 - 5, y + cell_size//2 - 8),
                                          user_val, fill=(224, 224, 224), font=font_big)

        # 단서 목록
        cx = margin + cols * cell_size + 15
        cy = 50
        draw.text((cx, cy), "→ 가로", fill=(245, 166, 35), font=font_big)
        cy += 24
        across = sorted([p for p in self.current_placed if p['direction'] == 'A'],
                        key=lambda x: x.get('number', 99))
        for p in across:
            txt = f"{p.get('number','?')}. {p['clue']} ({len(p['word'])})"
            draw.text((cx, cy), txt, fill=(200, 200, 200), font=font_small)
            cy += 16
            if cy > H - 60: break

        cy += 10
        draw.text((cx, cy), "↓ 세로", fill=(245, 166, 35), font=font_big)
        cy += 24
        down = sorted([p for p in self.current_placed if p['direction'] == 'D'],
                      key=lambda x: x.get('number', 99))
        for p in down:
            txt = f"{p.get('number','?')}. {p['clue']} ({len(p['word'])})"
            draw.text((cx, cy), txt, fill=(200, 200, 200), font=font_small)
            cy += 16
            if cy > H - 20: break

        return img

    # ════════════════════════════════════════════════════════
    # PDF 내보내기
    # ════════════════════════════════════════════════════════
    def _export_pdf(self):
        if not self.current_grid:
            messagebox.showwarning("저장 실패", "생성된 퍼즐이 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="PDF로 저장"
        )
        if not path: return
        try:
            self._create_pdf(path)
            messagebox.showinfo("저장 완료", f"PDF가 저장되었습니다:\n{path}")
        except Exception as e:
            messagebox.showerror("오류", f"PDF 생성 실패:\n{e}")

    def _create_pdf(self, path):
        """한글 폰트(나눔고딕)를 적용한 PDF 생성"""
        from reportlab.platypus import PageBreak

        PAGE_W, PAGE_H = A4
        doc = SimpleDocTemplate(path, pagesize=A4,
                                leftMargin=15*mm, rightMargin=15*mm,
                                topMargin=15*mm, bottomMargin=15*mm)

        # ── 스타일 (모두 KorFont 사용) ──────────────────────────
        title_style = ParagraphStyle(
            'kor_title',
            fontName=KOR_FONT_BOLD, fontSize=20,
            textColor=colors.HexColor('#0f3460'),
            spaceAfter=4, alignment=TA_CENTER,
            leading=26,
        )
        sub_style = ParagraphStyle(
            'kor_sub',
            fontName=KOR_FONT, fontSize=10,
            textColor=colors.HexColor('#555555'),
            spaceAfter=6, alignment=TA_CENTER,
        )
        heading_style = ParagraphStyle(
            'kor_heading',
            fontName=KOR_FONT_BOLD, fontSize=12,
            textColor=colors.HexColor('#0f3460'),
            spaceAfter=4, spaceBefore=6,
            borderPad=2,
        )
        clue_style = ParagraphStyle(
            'kor_clue',
            fontName=KOR_FONT, fontSize=9,
            textColor=colors.HexColor('#222222'),
            spaceAfter=2, leftIndent=12, leading=13,
        )
        answer_style = ParagraphStyle(
            'kor_answer',
            fontName=KOR_FONT_BOLD, fontSize=9,
            textColor=colors.HexColor('#0f3460'),
            spaceAfter=2, leftIndent=12,
        )
        num_cell_style = ParagraphStyle(
            'kor_numcell',
            fontName=KOR_FONT, fontSize=6,
            textColor=colors.HexColor('#003399'),
            leading=7,
        )

        story = []

        # ── 제목 ────────────────────────────────────────────────
        diff_label = {"전체": "전체", "하": "하(Easy)", "중": "중(Normal)", "상": "상(Hard)"}.get(
            self.diff_var.get(), "전체")
        story.append(Paragraph("크로스워드 퍼즐 (Crossword Puzzle)", title_style))
        story.append(Paragraph(
            f"난이도: {diff_label}  |  단어 수: {len(self.current_placed)}개  |  "
            f"그리드: {self.grid_rows}×{self.grid_cols}", sub_style))
        story.append(Spacer(1, 4*mm))

        # ── 퍼즐 그리드 테이블 ──────────────────────────────────
        rows, cols = self.grid_rows, self.grid_cols

        # 셀 크기: 가용 너비에 맞게 자동 조정
        avail_w = PAGE_W - 30*mm
        CELL_MM = min(14, avail_w / cols / mm)
        CELL = CELL_MM

        table_data = []
        for r in range(rows):
            row_data = []
            for c in range(cols):
                val = self.current_grid[r][c]
                if val == '#':
                    row_data.append('')
                else:
                    num_str = str(self.numbered[(r, c)]) if (r, c) in self.numbered else ''
                    row_data.append(Paragraph(num_str, num_cell_style))
            table_data.append(row_data)

        col_widths  = [CELL * mm] * cols
        row_heights = [CELL * mm] * rows
        grid_table  = Table(table_data, colWidths=col_widths, rowHeights=row_heights)

        ts = TableStyle([
            ('GRID',   (0,0), (-1,-1), 0.6, colors.HexColor('#2d4a7a')),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING',  (0,0), (-1,-1), 1),
            ('TOPPADDING',   (0,0), (-1,-1), 1),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0), (-1,-1), 0),
        ])
        for r in range(rows):
            for c in range(cols):
                if self.current_grid[r][c] == '#':
                    ts.add('BACKGROUND', (c,r), (c,r), colors.HexColor('#1a1a2e'))
                    ts.add('GRID', (c,r), (c,r), 0, colors.HexColor('#1a1a2e'))
                else:
                    ts.add('BACKGROUND', (c,r), (c,r), colors.white)
        grid_table.setStyle(ts)
        story.append(grid_table)
        story.append(Spacer(1, 6*mm))

        # ── 단서 목록 (가로 / 세로 나란히) ─────────────────────
        across = sorted([p for p in self.current_placed if p['direction'] == 'A'],
                        key=lambda x: x.get('number', 99))
        down   = sorted([p for p in self.current_placed if p['direction'] == 'D'],
                        key=lambda x: x.get('number', 99))

        def clue_paragraphs(word_list, symbol):
            items = [Paragraph(f"{symbol} 단서", heading_style)]
            for p in word_list:
                num = p.get('number', '?')
                txt = f"{num}. {p['clue']}  ({len(p['word'])}글자)"
                items.append(Paragraph(txt, clue_style))
            return items

        # 가로/세로를 두 컬럼으로 배치
        across_paras = clue_paragraphs(across, "→ 가로")
        down_paras   = clue_paragraphs(down,   "↓ 세로")

        # 두 컬럼 테이블에 넣기
        from reportlab.platypus import KeepInFrame
        half_w = (PAGE_W - 30*mm) / 2 - 4*mm
        col_a = KeepInFrame(half_w, 200*mm, across_paras, mode='shrink')
        col_d = KeepInFrame(half_w, 200*mm, down_paras,   mode='shrink')
        clue_table = Table([[col_a, col_d]],
                           colWidths=[half_w, half_w])
        clue_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING',  (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('LINEAFTER', (0,0), (0,-1), 0.5, colors.HexColor('#cccccc')),
        ]))
        story.append(clue_table)

        # ── 페이지 구분선 + 정답 키 ─────────────────────────────
        story.append(Spacer(1, 6*mm))
        story.append(Paragraph("─" * 80, ParagraphStyle('hr', fontName=KOR_FONT,
                                                          fontSize=6, textColor=colors.HexColor('#aaaaaa'))))
        story.append(Paragraph("✅  정답 (Answer Key)", heading_style))

        ans_rows = []
        for p in sorted(self.current_placed, key=lambda x: x.get('number', 99)):
            d_sym = '→' if p['direction'] == 'A' else '↓'
            ans_rows.append([
                Paragraph(f"{d_sym} {p.get('number','?')}", clue_style),
                Paragraph(p['clue'], clue_style),
                Paragraph(p['word'], answer_style),
            ])
        if ans_rows:
            ans_table = Table(ans_rows, colWidths=[14*mm, 100*mm, 35*mm])
            ans_table.setStyle(TableStyle([
                ('FONTNAME',        (0,0), (-1,-1), KOR_FONT),
                ('FONTSIZE',        (0,0), (-1,-1), 9),
                ('ROWBACKGROUNDS',  (0,0), (-1,-1),
                 [colors.HexColor('#f0f4ff'), colors.white]),
                ('GRID',            (0,0), (-1,-1), 0.3, colors.HexColor('#cccccc')),
                ('VALIGN',          (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING',     (0,0), (-1,-1), 4),
                ('TOPPADDING',      (0,0), (-1,-1), 2),
                ('BOTTOMPADDING',   (0,0), (-1,-1), 2),
            ]))
            story.append(ans_table)

        doc.build(story)

    # ════════════════════════════════════════════════════════
    # 인쇄
    # ════════════════════════════════════════════════════════
    def _print_puzzle(self):
        if not self.current_grid:
            messagebox.showwarning("인쇄 실패", "생성된 퍼즐이 없습니다.")
            return
        import tempfile, subprocess, sys
        tmp = tempfile.mktemp(suffix=".pdf")
        try:
            self._create_pdf(tmp)
            if sys.platform == "win32":
                os.startfile(tmp, "print")
            elif sys.platform == "darwin":
                subprocess.run(["lpr", tmp])
            else:
                # Linux
                result = subprocess.run(["which", "lpr"], capture_output=True)
                if result.returncode == 0:
                    subprocess.run(["lpr", tmp])
                else:
                    # PDF 뷰어로 열기
                    subprocess.run(["xdg-open", tmp])
            messagebox.showinfo("인쇄", "인쇄 대화상자가 열립니다.")
        except Exception as e:
            messagebox.showerror("인쇄 오류", f"인쇄 실패:\n{e}\n\nPDF로 저장 후 인쇄해주세요.")


# ════════════════════════════════════════════════════════════
# 실행
# ════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = CrosswordApp()
    app.mainloop()
