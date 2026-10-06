"""
N-Back Test with Face Stimuli
================================
Visual N-Back task with participant form + Excel export.
UI: English  |  Excel headers: Persian  |  Text inputs: Persian supported.
Includes in-app settings menu (gear icon, top-left).
Includes post-test restart prompt.
Includes Random / Manual stage ordering modes.
"""

import pygame
import random
import sys
import csv
import math
import json
import copy
from pathlib import Path
from datetime import datetime
from statistics import mean, NormalDist

# --- Excel support ---
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font as XLFont, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- Persian text rendering ---
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_PERSIAN = True
except ImportError:
    HAS_PERSIAN = False


# ============================================================
#  PERSIAN TEXT HELPERS
# ============================================================
def _has_persian(text):
    if not text:
        return False
    for ch in str(text):
        if ('\u0600' <= ch <= '\u06FF'
            or '\uFB50' <= ch <= '\uFDFF'
            or '\uFE70' <= ch <= '\uFEFF'):
            return True
    return False


def shape_persian(text):
    if text is None:
        return ""
    s = str(text)
    if not s:
        return s
    if not HAS_PERSIAN:
        return s
    if not _has_persian(s):
        return s
    try:
        return get_display(arabic_reshaper.reshape(s))
    except Exception:
        return s


# ============================================================
#  CONFIG
# ============================================================
CONFIG = {
    # --- N-back parameters ---
    "n_back": 2,
    "trials_per_stage": 22,
    "match_rate": 0.20,
    "target_folder_rate": 0.5,

    # --- Stage ordering mode ---
    "stage_mode": "random",          # "random" or "manual"
    "manual_stages": [
        {"target": "high",    "filler": "neutral", "target_rate": 0.5},
        {"target": "low",     "filler": "neutral", "target_rate": 0.5},
        {"target": "neutral", "filler": "high",    "target_rate": 0.5},
    ],

    # --- Timing (ms) ---
    "image_duration_ms": 2000,
    "rest_duration_ms": 1000,
    "fixation_duration_ms": 500,
    "feedback_duration_ms": 700,
    "inter_trial_pause_ms": 300,
    "post_response_delay_ms": 150,

    # --- Constraints ---
    "min_gap": 3,
    "min_match_gap": 3,

    # --- Scoring ---
    "score_correct": 1,
    "score_incorrect": -1,
    "score_timeout": -1,

    # --- Display ---
    "window_size": (1050, 760),
    "bg_color": (16, 18, 26),
    "panel_color": (26, 30, 42),
    "panel_light": (36, 42, 58),
    "border_color": (58, 66, 88),
    "text_color": (238, 240, 248),
    "muted_color": (150, 158, 180),
    "accent_color": (110, 180, 255),
    "success_color": (90, 220, 130),
    "error_color": (245, 105, 105),
    "warn_color": (245, 190, 80),
    "image_max_size": (560, 360),

    # --- Buttons ---
    "btn_size": (170, 52),
    "btn_gap": 30,
    "btn_radius": 16,
    "btn_match_fill": (110, 180, 255),
    "btn_match_border": (110, 180, 255),
    "btn_nomatch_fill": None,
    "btn_nomatch_border": (110, 180, 255),

    # --- Results ---
    "save_results": True,
    "results_dir": "results",
    "excel_filename": "participants.xlsx",
}


# ============================================================
#  EDITABLE KEYS / DEFAULTS
# ============================================================
EDITABLE_KEYS = [
    "n_back",
    "trials_per_stage",
    "match_rate",
    "target_folder_rate",
    "image_duration_ms",
    "rest_duration_ms",
    "fixation_duration_ms",
    "feedback_duration_ms",
    "inter_trial_pause_ms",
    "min_gap",
    "min_match_gap",
    "stage_mode",
    "manual_stages",
]

DEFAULT_CONFIG = copy.deepcopy({k: CONFIG[k] for k in EDITABLE_KEYS})


STAGE_DESIGNS = [
    {"target": "high",    "filler": "neutral", "label": "Attractive Faces"},
    {"target": "low",     "filler": "neutral", "label": "Unattractive Faces"},
    {"target": "neutral", "filler": "high",    "label": "Neutral Faces"},
]


FOLDERS = ["high", "low", "neutral"]

FOLDER_LABELS = {
    "high":    "Attractive Faces",
    "low":     "Unattractive Faces",
    "neutral": "Neutral Faces",
}

FOLDER_SHORT = {
    "high":    "Attractive",
    "low":     "Unattractive",
    "neutral": "Neutral",
}


# ============================================================
#  PATHS
# ============================================================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_DIR = get_base_dir()
IMAGES_DIR = BASE_DIR / "images"
SETTINGS_FILE = BASE_DIR / "settings.json"
FOLDER_PATHS = {
    "high":    IMAGES_DIR / "high",
    "neutral": IMAGES_DIR / "neutral",
    "low":     IMAGES_DIR / "low",
}


# ============================================================
#  SETTINGS PERSISTENCE
# ============================================================
def load_settings():
    if not SETTINGS_FILE.exists():
        return
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k, v in data.items():
            if k not in DEFAULT_CONFIG:
                continue

            if k == "stage_mode":
                CONFIG[k] = v if v in ("random", "manual") else "random"
                continue

            if k == "manual_stages":
                if (isinstance(v, list) and len(v) == 3
                        and all(isinstance(s, dict) for s in v)):
                    validated = []
                    for s in v:
                        target = s.get("target", "high")
                        filler = s.get("filler", "neutral")
                        rate = s.get("target_rate", 0.5)
                        if target not in FOLDER_LABELS:
                            target = "high"
                        if filler not in FOLDER_LABELS:
                            filler = "neutral"
                        try:
                            rate = float(rate)
                        except (TypeError, ValueError):
                            rate = 0.5
                        rate = max(0.10, min(0.90, rate))
                        validated.append({
                            "target": target,
                            "filler": filler,
                            "target_rate": rate,
                        })
                    CONFIG[k] = validated
                continue

            default = DEFAULT_CONFIG[k]
            if isinstance(default, bool):
                CONFIG[k] = bool(v)
            elif isinstance(default, float):
                CONFIG[k] = float(v)
            elif isinstance(default, int):
                CONFIG[k] = int(round(float(v)))
            else:
                CONFIG[k] = v
    except Exception as e:
        print(f"[Settings] Failed to load settings.json: {e}")


def save_settings():
    try:
        data = {k: CONFIG[k] for k in EDITABLE_KEYS if k in CONFIG}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except PermissionError:
        print(f"[Settings] Cannot write settings.json — file is locked.")
    except Exception as e:
        print(f"[Settings] Failed to save settings.json: {e}")


# ============================================================
#  SETTINGS SPEC
# ============================================================
SETTINGS_SPEC = [
    {
        "key": "image_duration_ms",
        "label": "Image Duration",
        "min": 500, "max": 5000, "step": 100,
        "format": lambda v: f"{v/1000:.1f} s",
    },
    {
        "key": "rest_duration_ms",
        "label": "Rest Between Images",
        "min": 100, "max": 3000, "step": 100,
        "format": lambda v: f"{v/1000:.1f} s",
    },
    {
        "key": "fixation_duration_ms",
        "label": "Fixation Duration",
        "min": 0, "max": 2000, "step": 100,
        "format": lambda v: f"{v/1000:.1f} s",
    },
    {
        "key": "feedback_duration_ms",
        "label": "Feedback Duration",
        "min": 0, "max": 2000, "step": 100,
        "format": lambda v: f"{v/1000:.1f} s",
    },
    {
        "key": "trials_per_stage",
        "label": "Trials per Stage",
        "min": 8, "max": 80, "step": 2,
        "format": lambda v: str(int(v)),
    },
    {
        "key": "n_back",
        "label": "N-Back Level",
        "min": 1, "max": 5, "step": 1,
        "format": lambda v: str(int(v)),
    },
    {
        "key": "match_rate",
        "label": "Match Rate",
        "min": 0.10, "max": 0.60, "step": 0.05,
        "format": lambda v: f"{v*100:.0f}%",
    },
    {
        "key": "min_match_gap",
        "label": "Min Match Gap",
        "min": 0, "max": 6, "step": 1,
        "format": lambda v: str(int(v)),
    },
]


# ============================================================
#  FONTS
# ============================================================
def _pick_font_name():
    available = pygame.font.get_fonts()
    for name in ("tahoma", "segoeui", "arial", "dejavusans"):
        if name in available:
            return name
    return None


def make_fonts():
    family = _pick_font_name()

    def F(size, bold=True):
        if family:
            return pygame.font.SysFont(family, size, bold=bold)
        return pygame.font.Font(None, size)

    return {
        "hero":   F(42),
        "title":  F(30),
        "huge":   F(24),
        "large":  F(21),
        "med":    F(19),
        "body":   F(17),
        "small":  F(15),
        "tiny":   F(13),
        "hint":   F(12, bold=False),
    }


# ============================================================
#  IMAGE LOADING
# ============================================================
def load_folder_images(folder_name):
    folder = FOLDER_PATHS[folder_name]
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")
    exts = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}
    images = sorted([p for p in folder.iterdir() if p.suffix.lower() in exts])
    if len(images) < 2:
        raise ValueError(f"Need at least 2 images in {folder}, found {len(images)}")
    return images


def load_image_scaled(path, max_size):
    img = pygame.image.load(str(path)).convert()
    w, h = img.get_size()
    scale = min(max_size[0] / w, max_size[1] / h, 1.0)
    return pygame.transform.smoothscale(img, (int(w * scale), int(h * scale)))


def preload_images(folder_images, max_size):
    cache = {}
    for paths in folder_images.values():
        for p in paths:
            if p not in cache:
                cache[p] = load_image_scaled(p, max_size)
    return cache


# ============================================================
#  SEQUENCE GENERATION
# ============================================================
def max_matches_for(n_trials, min_match_gap, n_back=0):
    if min_match_gap < 0:
        return max(0, n_trials - n_back)
    step = min_match_gap + 1
    available = n_trials - n_back
    if available <= 0:
        return 0
    return (available - 1) // step + 1


def validate_config(config):
    n = config["trials_per_stage"]
    mr = config["match_rate"]
    g = config["min_match_gap"]
    nb = config["n_back"]

    max_m = max_matches_for(n, g, nb)
    num_m = int(round(n * mr))

    if num_m > max_m:
        if n > 0:
            new_mr = max_m / n
            new_mr = math.floor(new_mr * 100) / 100
            if new_mr < 0.10:
                new_mr = 0.10
            print(
                f"[Config] Requested {num_m} matches but max is {max_m} "
                f"with these settings. Auto-adjusting match_rate "
                f"from {mr:.2f} to {new_mr:.2f}."
            )
            config["match_rate"] = new_mr
    return True


def _config_warning(config):
    n = config["trials_per_stage"]
    mr = config["match_rate"]
    g = config["min_match_gap"]
    nb = config["n_back"]

    max_m = max_matches_for(n, g, nb)
    num_m = int(round(n * mr))

    if num_m > max_m:
        return (f"Note: {num_m} matches requested, but only {max_m} fit "
                f"with these settings — will use {max_m}.")
    return None


def choose_match_positions(n_trials, n_back, num_matches, min_match_gap,
                            max_tries=5000):
    if num_matches == 0:
        return set()

    step = min_match_gap + 1
    candidates = list(range(n_back, n_trials))
    if len(candidates) < num_matches:
        return None

    for _ in range(max_tries):
        random.shuffle(candidates)
        selected = []
        for pos in candidates:
            if all(abs(pos - s) >= step for s in selected):
                selected.append(pos)
                if len(selected) == num_matches:
                    return set(selected)
    return None


def build_folder_sequence(num_trials, n_back, num_target, match_positions):
    num_matches = len(match_positions)

    max_target_matches = max(1, num_target // 2)
    num_target_matches = min(max_target_matches, max(1, num_matches // 2))

    match_list = list(match_positions)
    random.shuffle(match_list)
    target_matches = set(match_list[:num_target_matches])

    folders = [None] * num_trials

    for i in target_matches:
        folders[i] = "target"
        folders[i - n_back] = "target"

    for i in match_positions:
        if i in target_matches:
            continue
        if folders[i] is None:
            folders[i] = "filler"
        if folders[i - n_back] is None:
            folders[i - n_back] = "filler"

    target_count = sum(1 for f in folders if f == "target")
    if target_count > num_target:
        return None

    target_remaining = num_target - target_count
    available = [i for i in range(num_trials) if folders[i] is None]

    if len(available) < target_remaining:
        return None

    for i in random.sample(available, target_remaining):
        folders[i] = "target"

    for i in range(num_trials):
        if folders[i] is None:
            folders[i] = "filler"

    return folders


def assign_images(folders, match_positions, n_back, min_gap,
                  target_folder, filler_folder, folder_images):
    num_trials = len(folders)
    lookback = max(min_gap, n_back)

    target_pool = folder_images[target_folder]
    filler_pool = folder_images[filler_folder]

    seq = []
    for i in range(num_trials):
        folder = folders[i]
        pool = target_pool if folder == "target" else filler_pool

        if i in match_positions:
            src = seq[i - n_back]
            if src["folder"] != folder:
                raise ValueError(f"Folder mismatch at {i}")
            image = src["image"]

            recent = {seq[j]["image"]
                      for j in range(max(0, i - min_gap), i)
                      if j != i - n_back and seq[j]["folder"] == folder}
            if image in recent:
                raise ValueError(f"Match conflict at {i}")
        else:
            forbidden = {seq[j]["image"]
                         for j in range(max(0, i - lookback), i)
                         if seq[j]["folder"] == folder}
            candidates = [img for img in pool if img not in forbidden]
            if not candidates:
                forbidden = {seq[j]["image"]
                             for j in range(max(0, i - 1), i)
                             if seq[j]["folder"] == folder}
                candidates = [img for img in pool if img not in forbidden]
            if not candidates:
                candidates = list(pool)
            image = random.choice(candidates)

        seq.append({
            "folder": folder,
            "is_match": i in match_positions,
            "image": image,
        })

    return seq


def generate_stage_sequence(n_back, num_trials, target_folder, filler_folder,
                            folder_images, match_rate, target_rate,
                            min_gap, min_match_gap, max_attempts=5000):
    num_target = int(round(num_trials * target_rate))
    num_matches = int(round(num_trials * match_rate))

    max_m = max_matches_for(num_trials, min_match_gap, n_back)
    if num_matches > max_m:
        num_matches = max_m

    for _ in range(max_attempts):
        match_positions = choose_match_positions(
            num_trials, n_back, num_matches, min_match_gap
        )
        if match_positions is None:
            continue

        folders = build_folder_sequence(
            num_trials, n_back, num_target, match_positions
        )
        if folders is None:
            continue

        try:
            return assign_images(
                folders, match_positions, n_back, min_gap,
                target_folder, filler_folder, folder_images
            )
        except ValueError:
            continue

    raise RuntimeError(
        "Could not generate sequence.\n"
        "Try: more images per folder, lower match_rate, or lower min_gap."
    )


# ============================================================
#  DRAW HELPERS
# ============================================================
def draw_text(screen, text, font, color, center=None, topleft=None):
    surf = font.render(str(text), True, color)
    rect = surf.get_rect()
    if center is not None:
        rect.center = center
    elif topleft is not None:
        rect.topleft = topleft
    screen.blit(surf, rect)
    return rect


def draw_panel(screen, rect, color, border=None, radius=14, width=2):
    pygame.draw.rect(screen, color, rect, border_radius=radius)
    if border is not None:
        pygame.draw.rect(screen, border, rect, width=width, border_radius=radius)


def draw_button(screen, rect, text, font, fill_color, border_color,
                selected, hover):
    radius = CONFIG.get("btn_radius", 20)

    if fill_color is not None:
        if selected:
            fill = fill_color
        elif hover:
            fill = tuple(min(255, c + 30) for c in fill_color)
        else:
            fill = fill_color
    else:
        if selected:
            fill = (40, 70, 110)
        elif hover:
            fill = (30, 50, 80)
        else:
            fill = None

    if selected:
        border = (255, 255, 255)
    elif hover:
        border = tuple(min(255, c + 40) for c in border_color)
    else:
        border = border_color

    if fill is not None:
        pygame.draw.rect(screen, fill, rect, border_radius=radius)

    pygame.draw.rect(screen, border, rect, width=2, border_radius=radius)

    if fill_color is None and not selected:
        text_color = border_color
    else:
        text_color = (255, 255, 255)

    txt = font.render(str(text), True, text_color)
    screen.blit(txt, txt.get_rect(center=rect.center))


def draw_timer_bar(screen, rect, fraction, color_ok, color_bad):
    pygame.draw.rect(screen, (40, 44, 58), rect, border_radius=8)
    frac = max(0.0, min(1.0, fraction))
    color = color_ok if frac > 0.3 else color_bad
    if frac > 0:
        filled = pygame.Rect(rect.x, rect.y, int(rect.width * frac), rect.height)
        pygame.draw.rect(screen, color, filled, border_radius=8)
    pygame.draw.rect(screen, (70, 76, 96), rect, width=2, border_radius=8)


def draw_gear_icon(screen, center, radius, color):
    cx, cy = center
    teeth = 6
    tooth_r = max(2, radius // 5)
    for i in range(teeth):
        angle = 2 * math.pi * i / teeth - math.pi / 2
        tx = cx + math.cos(angle) * (radius + tooth_r)
        ty = cy + math.sin(angle) * (radius + tooth_r)
        pygame.draw.circle(screen, color, (int(tx), int(ty)), tooth_r)
    pygame.draw.circle(screen, color, (cx, cy), radius, 2)
    pygame.draw.circle(screen, color, (cx, cy), max(2, radius // 3), 2)


def draw_small_button(screen, rect, text, font, color, hover, enabled=True):
    if not enabled:
        fill = (36, 40, 52)
        border = (60, 66, 80)
        txt_color = (100, 108, 128)
    else:
        fill = (tuple(c // 2 for c in color) if hover
                else tuple(c // 3 for c in color))
        border = color
        txt_color = (255, 255, 255)

    pygame.draw.rect(screen, fill, rect, border_radius=16)
    pygame.draw.rect(screen, border, rect, width=2, border_radius=16)
    txt = font.render(text, True, txt_color)
    screen.blit(txt, txt.get_rect(center=rect.center))


def drain_events():
    pygame.event.clear()
    pygame.event.pump()


# ============================================================
#  FORM WIDGETS
# ============================================================
class TextInput:
    def __init__(self, label, x, y, width, height, numeric=False, max_len=60):
        self.label = label
        self.rect = pygame.Rect(x, y, width, height)
        self.value = ""
        self.active = False
        self.numeric = numeric
        self.max_len = max_len

    def is_click_inside(self, pos):
        return self.rect.collidepoint(pos)

    def activate(self):
        self.active = True

    def deactivate(self):
        self.active = False

    def handle_key(self, event):
        if not self.active:
            return False
        if event.key == pygame.K_BACKSPACE:
            self.value = self.value[:-1]
            return True
        if event.key in (pygame.K_TAB, pygame.K_RETURN, pygame.K_ESCAPE):
            return False
        if event.unicode and event.unicode.isprintable():
            if self.numeric and not event.unicode.isdigit():
                return True
            if len(self.value) < self.max_len:
                self.value += event.unicode
            return True
        return False

    def draw(self, screen, fonts, colors, time_ms):
        draw_text(screen, self.label, fonts["body"], colors["muted_color"],
                  topleft=(self.rect.x, self.rect.y - 26))

        border_color = (colors["accent_color"] if self.active
                        else colors["border_color"])
        pygame.draw.rect(screen, colors["panel_light"], self.rect,
                         border_radius=16)
        pygame.draw.rect(screen, border_color, self.rect,
                         width=2, border_radius=16)

        display_text = shape_persian(self.value)
        text_surf = fonts["body"].render(display_text, True, colors["text_color"])
        max_w = self.rect.width - 20
        clip_rect = pygame.Rect(self.rect.x + 10, self.rect.y,
                                max_w, self.rect.height)

        old_clip = screen.get_clip()
        screen.set_clip(clip_rect)
        screen.blit(text_surf, (self.rect.x + 10,
                                self.rect.centery - text_surf.get_height() // 2))
        screen.set_clip(old_clip)

        if self.active and (time_ms // 500) % 2 == 0:
            if _has_persian(self.value):
                cursor_x = self.rect.x + 10
            else:
                visible_w = min(text_surf.get_width(), max_w)
                cursor_x = self.rect.x + 10 + visible_w
            pygame.draw.line(screen, colors["text_color"],
                             (cursor_x, self.rect.y + 8),
                             (cursor_x, self.rect.bottom - 8), 2)


class GenderSelector:
    def __init__(self, label, x, y, width, height, gap=15):
        self.label = label
        self.options = ["Male", "Female"]
        self.selected = None
        self.rects = [
            pygame.Rect(x, y, width, height),
            pygame.Rect(x + width + gap, y, width, height),
        ]

    def handle_click(self, pos):
        for i, r in enumerate(self.rects):
            if r.collidepoint(pos):
                self.selected = self.options[i]
                return True
        return False

    def draw(self, screen, fonts, colors):
        draw_text(screen, self.label, fonts["body"], colors["muted_color"],
                  topleft=(self.rects[0].x, self.rects[0].y - 26))

        for i, r in enumerate(self.rects):
            is_selected = (self.selected == self.options[i])
            fill = colors["accent_color"] if is_selected else colors["panel_light"]
            border = colors["accent_color"] if is_selected else colors["border_color"]

            pygame.draw.rect(screen, fill, r, border_radius=16)
            pygame.draw.rect(screen, border, r, width=2, border_radius=16)

            txt_color = (255, 255, 255) if is_selected else colors["text_color"]
            txt = fonts["body"].render(self.options[i], True, txt_color)
            screen.blit(txt, txt.get_rect(center=r.center))


# ============================================================
#  STAGE CONFIG SCREEN (Manual Mode)
# ============================================================
def show_stage_config_screen(screen, config, fonts):
    drain_events()
    clock = pygame.time.Clock()
    W, H = config["window_size"]

    panel_w, panel_h = 1000, 500
    panel = pygame.Rect((W - panel_w) // 2, (H - panel_h) // 2,
                        panel_w, panel_h)

    stages = config["manual_stages"]

    stage_layouts = []
    rows_start = panel.y + 110
    row_h = 100

    btn_w, btn_h = 90, 34
    btn_gap = 6
    btn_radius = 16          

    for i in range(3):
        y = rows_start + i * row_h

        target_x = panel.x + 130
        target_btns = []
        for j, folder in enumerate(FOLDERS):
            r = pygame.Rect(target_x + j * (btn_w + btn_gap), y + 40,
                            btn_w, btn_h)
            target_btns.append((folder, r))

        filler_x = target_x + 3 * (btn_w + btn_gap) + 40
        filler_btns = []
        for j, folder in enumerate(FOLDERS):
            r = pygame.Rect(filler_x + j * (btn_w + btn_gap), y + 40,
                            btn_w, btn_h)
            filler_btns.append((folder, r))

        rate_x = filler_x + 3 * (btn_w + btn_gap) + 40
        minus_r = pygame.Rect(rate_x, y + 40, 34, btn_h)
        value_r = pygame.Rect(rate_x + 40, y + 40, 70, btn_h)
        plus_r  = pygame.Rect(rate_x + 116, y + 40, 34, btn_h)

        stage_layouts.append({
            "i": i, "y": y,
            "target_btns": target_btns,
            "filler_btns": filler_btns,
            "minus_r": minus_r,
            "value_r": value_r,
            "plus_r": plus_r,
        })

    back_rect = pygame.Rect(W // 2 - 100, panel.bottom - 68, 200, 48)

    while True:
        mouse_pos = pygame.mouse.get_pos()

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.KEYDOWN and event.key == pygame.K_ESCAPE:
                save_settings()
                return

            if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                for lay in stage_layouts:
                    st = stages[lay["i"]]

                    for folder, r in lay["target_btns"]:
                        if r.collidepoint(event.pos):
                            st["target"] = folder
                    for folder, r in lay["filler_btns"]:
                        if r.collidepoint(event.pos):
                            st["filler"] = folder

                    if lay["minus_r"].collidepoint(event.pos):
                        st["target_rate"] = max(0.10, round(st["target_rate"] - 0.05, 2))
                    if lay["plus_r"].collidepoint(event.pos):
                        st["target_rate"] = min(0.90, round(st["target_rate"] + 0.05, 2))

                if back_rect.collidepoint(event.pos):
                    save_settings()
                    return

        screen.fill(config["bg_color"])
        draw_panel(screen, panel, config["panel_color"],
                   border=config["border_color"], radius=22, width=2)

        draw_text(screen, "Configure Stages (Manual Mode)", fonts["title"],
                  config["accent_color"], center=(W // 2, panel.y + 45))
        pygame.draw.line(screen, config["border_color"],
                         (panel.x + 40, panel.y + 82),
                         (panel.right - 40, panel.y + 82), 2)

        for lay in stage_layouts:
            i = lay["i"]
            st = stages[i]
            y = lay["y"]

            draw_text(screen, f"Stage {i + 1}", fonts["large"],
                      config["text_color"], topleft=(panel.x + 40, y + 48))

            draw_text(screen, "Target", fonts["small"], config["muted_color"],
                      topleft=(lay["target_btns"][0][1].x, y + 20))
            for folder, r in lay["target_btns"]:
                sel = (st["target"] == folder)
                fill = config["accent_color"] if sel else config["panel_light"]
                border = config["accent_color"] if sel else config["border_color"]
                pygame.draw.rect(screen, fill, r, border_radius=btn_radius)
                pygame.draw.rect(screen, border, r, width=2,
                                 border_radius=btn_radius)
                tc = (255, 255, 255) if sel else config["text_color"]
                draw_text(screen, FOLDER_SHORT[folder], fonts["tiny"],
                          tc, center=r.center)

            draw_text(screen, "Filler", fonts["small"], config["muted_color"],
                      topleft=(lay["filler_btns"][0][1].x, y + 20))
            for folder, r in lay["filler_btns"]:
                sel = (st["filler"] == folder)
                fill = config["accent_color"] if sel else config["panel_light"]
                border = config["accent_color"] if sel else config["border_color"]
                pygame.draw.rect(screen, fill, r, border_radius=btn_radius)
                pygame.draw.rect(screen, border, r, width=2,
                                 border_radius=btn_radius)
                tc = (255, 255, 255) if sel else config["text_color"]
                draw_text(screen, FOLDER_SHORT[folder], fonts["tiny"],
                          tc, center=r.center)

            draw_text(screen, "Rate", fonts["small"], config["muted_color"],
                      topleft=(lay["minus_r"].x, y + 20))

            # Minus button
            mh = lay["minus_r"].collidepoint(mouse_pos)
            m_fill = ((tuple(c // 2 for c in config["accent_color"])) if mh
                      else tuple(c // 3 for c in config["accent_color"]))
            pygame.draw.rect(screen, m_fill, lay["minus_r"],
                             border_radius=btn_radius)
            pygame.draw.rect(screen, config["accent_color"], lay["minus_r"],
                             width=2, border_radius=btn_radius)
            draw_text(screen, "-", fonts["small"], (255, 255, 255),
                      center=lay["minus_r"].center)

            # Value box
            pygame.draw.rect(screen, config["panel_light"], lay["value_r"],
                             border_radius=btn_radius)
            pygame.draw.rect(screen, config["border_color"], lay["value_r"],
                             width=2, border_radius=btn_radius)
            draw_text(screen, f"{int(round(st['target_rate']*100))}%",
                      fonts["small"], config["text_color"],
                      center=lay["value_r"].center)

            # Plus button
            ph = lay["plus_r"].collidepoint(mouse_pos)
            p_fill = ((tuple(c // 2 for c in config["accent_color"])) if ph
                      else tuple(c // 3 for c in config["accent_color"]))
            pygame.draw.rect(screen, p_fill, lay["plus_r"],
                             border_radius=btn_radius)
            pygame.draw.rect(screen, config["accent_color"], lay["plus_r"],
                             width=2, border_radius=btn_radius)
            draw_text(screen, "+", fonts["small"], (255, 255, 255),
                      center=lay["plus_r"].center)

        # Back button
        bh = back_rect.collidepoint(mouse_pos)
        bf = (tuple(min(255, c + 30) for c in config["accent_color"])
              if bh else config["accent_color"])
        pygame.draw.rect(screen, bf, back_rect, border_radius=22)
        draw_text(screen, "Back", fonts["small"], (255, 255, 255),
                  center=back_rect.center)

        pygame.display.flip()
        clock.tick(60)


# ============================================================
#  SETTINGS SCREEN
# ============================================================
def show_settings_screen(screen, config, fonts):
    drain_events()
    clock = pygame.time.Clock()
    W, H = config["window_size"]

    panel_w, panel_h = 820, 720
    panel = pygame.Rect((W - panel_w) // 2, (H - panel_h) // 2,
                        panel_w, panel_h)

    row_h = 42
    row_gap = 4
    rows_start_y = panel.y + 100

    btn_w = 34
    btn_h = 34
    label_x = panel.x + 40
    controls_right = panel.right - 40
    value_box_w = 130
    value_box_x = controls_right - value_box_w - btn_w * 2 - 20
    minus_x = value_box_x + value_box_w + 10
    plus_x = minus_x + btn_w + 10

    rows = []
    for i, spec in enumerate(SETTINGS_SPEC):
        y = rows_start_y + i * (row_h + row_gap)
        rows.append({
            "spec": spec,
            "y": y,
            "minus_rect": pygame.Rect(minus_x, y, btn_w, btn_h),
            "value_rect": pygame.Rect(value_box_x, y, value_box_w, btn_h),
            "plus_rect":  pygame.Rect(plus_x, y, btn_w, btn_h),
        })

    # Stage mode row
    sm_y = rows_start_y + len(SETTINGS_SPEC) * (row_h + row_gap)
    sm_label_x = label_x
    sm_random_rect = pygame.Rect(value_box_x - 30, sm_y, 90, btn_h)
    sm_manual_rect = pygame.Rect(value_box_x + 70, sm_y, 90, btn_h)

    # Configure stages button row
    cfg_y = sm_y + row_h + row_gap
    cfg_rect = pygame.Rect(value_box_x - 30, cfg_y, 260, btn_h)

    reset_rect = pygame.Rect(panel.x + 40, panel.bottom - 68, 200, 50)
    save_rect  = pygame.Rect(panel.right - 240, panel.bottom - 68, 200, 50)

    while True:
        mouse_pos = pygame.mouse.get_pos()

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.KEYDOWN and event.key == pygame.K_ESCAPE:
                save_settings()
                return

            if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                for row in rows:
                    spec = row["spec"]
                    key = spec["key"]
                    val = config[key]

                    if row["minus_rect"].collidepoint(event.pos):
                        new_val = val - spec["step"]
                        if new_val >= spec["min"] - 1e-9:
                            config[key] = round(new_val, 4)
                    elif row["plus_rect"].collidepoint(event.pos):
                        new_val = val + spec["step"]
                        if new_val <= spec["max"] + 1e-9:
                            config[key] = round(new_val, 4)

                # Stage mode toggle
                if sm_random_rect.collidepoint(event.pos):
                    config["stage_mode"] = "random"
                elif sm_manual_rect.collidepoint(event.pos):
                    config["stage_mode"] = "manual"

                # Configure stages
                if (config["stage_mode"] == "manual"
                        and cfg_rect.collidepoint(event.pos)):
                    show_stage_config_screen(screen, config, fonts)
                    drain_events()
                    continue

                if reset_rect.collidepoint(event.pos):
                    for k, v in DEFAULT_CONFIG.items():
                        config[k] = copy.deepcopy(v)

                if save_rect.collidepoint(event.pos):
                    save_settings()
                    return

        screen.fill(config["bg_color"])
        draw_panel(screen, panel, config["panel_color"],
                   border=config["border_color"], radius=22, width=2)

        draw_text(screen, "Settings", fonts["title"], config["accent_color"],
                  center=(W // 2, panel.y + 42))
        pygame.draw.line(screen, config["border_color"],
                         (panel.x + 40, panel.y + 76),
                         (panel.right - 40, panel.y + 76), 2)

        for row in rows:
            spec = row["spec"]
            val = config[spec["key"]]

            draw_text(screen, spec["label"], fonts["body"],
                      config["text_color"],
                      topleft=(label_x, row["y"] + 6))

            at_min = val <= spec["min"] + 1e-9
            mh = row["minus_rect"].collidepoint(mouse_pos)
            draw_small_button(screen, row["minus_rect"], "-", fonts["small"],
                              config["accent_color"], mh,
                              enabled=not at_min)

            pygame.draw.rect(screen, config["panel_light"],
                             row["value_rect"], border_radius=16)
            pygame.draw.rect(screen, config["border_color"],
                             row["value_rect"], width=2, border_radius=16)
            draw_text(screen, spec["format"](val), fonts["body"],
                      config["text_color"], center=row["value_rect"].center)

            at_max = val >= spec["max"] - 1e-9
            ph = row["plus_rect"].collidepoint(mouse_pos)
            draw_small_button(screen, row["plus_rect"], "+", fonts["small"],
                              config["accent_color"], ph,
                              enabled=not at_max)

        # --- Stage mode row ---
        draw_text(screen, "Stage Mode", fonts["body"], config["text_color"],
                  topleft=(sm_label_x, sm_y + 6))

        is_random = (config["stage_mode"] == "random")
        rh = sm_random_rect.collidepoint(mouse_pos)
        mh2 = sm_manual_rect.collidepoint(mouse_pos)

        for rect, label, selected, hov in [
            (sm_random_rect, "Random", is_random, rh),
            (sm_manual_rect, "Manual", not is_random, mh2),
        ]:
            if selected:
                fill = config["accent_color"]
                border = config["accent_color"]
                tc = (255, 255, 255)
            elif hov:
                fill = (30, 50, 80)
                border = config["accent_color"]
                tc = config["accent_color"]
            else:
                fill = config["panel_light"]
                border = config["border_color"]
                tc = config["text_color"]
            pygame.draw.rect(screen, fill, rect, border_radius=16)
            pygame.draw.rect(screen, border, rect, width=2, border_radius=16)
            draw_text(screen, label, fonts["small"], tc, center=rect.center)

        # --- Configure stages button ---
        if not is_random:
            ch = cfg_rect.collidepoint(mouse_pos)
            cf = (tuple(min(255, c + 30) for c in config["accent_color"])
                  if ch else tuple(c // 2 for c in config["accent_color"]))
            pygame.draw.rect(screen, cf, cfg_rect, border_radius=16)
            pygame.draw.rect(screen, config["accent_color"], cfg_rect,
                             width=2, border_radius=16)
            draw_text(screen, "Configure Stages...", fonts["small"],
                      (255, 255, 255), center=cfg_rect.center)

        warn = _config_warning(config)
        if warn:
            draw_text(screen, warn, fonts["tiny"], config["warn_color"],
                      center=(W // 2, panel.bottom - 90))

        rh2 = reset_rect.collidepoint(mouse_pos)
        rfill = (70, 76, 94) if rh2 else (50, 54, 70)
        pygame.draw.rect(screen, rfill, reset_rect, border_radius=16)
        pygame.draw.rect(screen, (120, 128, 148), reset_rect,
                         width=2, border_radius=16)
        draw_text(screen, "Reset Defaults", fonts["body"],
                  config["text_color"], center=reset_rect.center)

        sh = save_rect.collidepoint(mouse_pos)
        sfill = (tuple(min(255, c + 30) for c in config["accent_color"])
                 if sh else config["accent_color"])
        pygame.draw.rect(screen, sfill, save_rect, border_radius=16)
        draw_text(screen, "Save & Close", fonts["body"], (255, 255, 255),
                  center=save_rect.center)

        pygame.display.flip()
        clock.tick(60)


# ============================================================
#  PARTICIPANT FORM
# ============================================================
def show_participant_form(screen, config, fonts):
    drain_events()
    clock = pygame.time.Clock()
    W, H = config["window_size"]

    panel_w, panel_h = 700, 560
    panel = pygame.Rect((W - panel_w) // 2, (H - panel_h) // 2,
                        panel_w, panel_h)

    fx = panel.x + 60
    fw = panel_w - 120
    fh = 40

    name_input = TextInput("Full Name", fx, panel.y + 130, fw, fh)
    field_input = TextInput("Field of Study", fx, panel.y + 215, fw, fh)
    age_input = TextInput("Age", fx, panel.y + 300, 170, fh,
                          numeric=True, max_len=3)
    gender = GenderSelector("Gender", fx + 240, panel.y + 300, 120, fh)

    inputs = [name_input, field_input, age_input]
    name_input.activate()

    submit_rect = pygame.Rect(W // 2 - 120, panel.bottom - 110, 240, 52)
    gear_rect = pygame.Rect(16, 16, 38, 38)
    error_message = ""

    while True:
        submit_requested = False

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                if gear_rect.collidepoint(event.pos):
                    show_settings_screen(screen, config, fonts)
                    drain_events()
                    continue

                if submit_rect.collidepoint(event.pos):
                    submit_requested = True
                    continue

                clicked_inside = False
                for inp in inputs:
                    if inp.is_click_inside(event.pos):
                        for other in inputs:
                            other.deactivate()
                        inp.activate()
                        clicked_inside = True
                        break

                if not clicked_inside:
                    gender.handle_click(event.pos)
                continue

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_TAB:
                    current_idx = -1
                    for i, inp in enumerate(inputs):
                        if inp.active:
                            current_idx = i
                            break
                    for inp in inputs:
                        inp.deactivate()
                    next_idx = (current_idx + 1) % len(inputs)
                    inputs[next_idx].activate()
                    continue

                if event.key == pygame.K_RETURN:
                    submit_requested = True
                    continue

                for inp in inputs:
                    if inp.handle_key(event):
                        break

        if submit_requested:
            error_message = ""
            if not name_input.value.strip():
                error_message = "Please enter your full name."
            elif not field_input.value.strip():
                error_message = "Please enter your field of study."
            elif not age_input.value.strip():
                error_message = "Please enter your age."
            elif not (5 <= int(age_input.value) <= 120):
                error_message = "Please enter a valid age (5-120)."
            elif gender.selected is None:
                error_message = "Please select your gender."

            if not error_message:
                return {
                    "full_name": name_input.value.strip(),
                    "field_of_study": field_input.value.strip(),
                    "age": int(age_input.value),
                    "gender": gender.selected,
                }

        now = pygame.time.get_ticks()
        screen.fill(config["bg_color"])
        draw_panel(screen, panel, config["panel_color"],
                   border=config["border_color"], radius=22, width=2)

        draw_text(screen, "Participant Information", fonts["title"],
                  config["accent_color"], center=(W // 2, panel.y + 50))
        pygame.draw.line(screen, config["border_color"],
                         (panel.x + 60, panel.y + 82),
                         (panel.right - 60, panel.y + 82), 2)

        for inp in inputs:
            inp.draw(screen, fonts, config, now)
        gender.draw(screen, fonts, config)

        hover = submit_rect.collidepoint(pygame.mouse.get_pos())
        btn_color = (tuple(min(255, c + 30) for c in config["accent_color"])
                     if hover else config["accent_color"])
        pygame.draw.rect(screen, btn_color, submit_rect, border_radius=16)
        draw_text(screen, "Start Test", fonts["med"], (255, 255, 255),
                  center=submit_rect.center)

        if error_message:
            draw_text(screen, error_message, fonts["body"],
                      config["error_color"],
                      center=(W // 2, panel.bottom - 145))

        draw_text(screen, "TAB: next field    Enter: start test",
                fonts["hint"], config["muted_color"],
                center=(W // 2, panel.bottom - 22))

        gear_hover = gear_rect.collidepoint(pygame.mouse.get_pos())
        gear_bg = (config["panel_light"] if gear_hover
                   else config["panel_color"])
        pygame.draw.rect(screen, gear_bg, gear_rect, border_radius=16)
        pygame.draw.rect(screen, config["border_color"], gear_rect,
                         width=1, border_radius=16)
        gear_color = (config["accent_color"] if gear_hover
                      else config["muted_color"])
        draw_gear_icon(screen, gear_rect.center, 11, gear_color)

        pygame.display.flip()
        clock.tick(60)


# ============================================================
#  MESSAGE SCREEN
# ============================================================
def show_message_screen(screen, config, fonts, title, subtitle, lines,
                         wait_for_key=True):
    drain_events()
    clock = pygame.time.Clock()
    W, H = config["window_size"]

    screen.fill(config["bg_color"])
    pygame.display.flip()

    while True:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if wait_for_key and event.type == pygame.KEYDOWN:
                if event.key in (pygame.K_SPACE, pygame.K_RETURN):
                    return

        screen.fill(config["bg_color"])

        panel_w, panel_h = 720, 440
        panel = pygame.Rect((W - panel_w) // 2, (H - panel_h) // 2,
                            panel_w, panel_h)
        draw_panel(screen, panel, config["panel_color"],
                   border=config["border_color"], radius=22, width=2)

        draw_text(screen, title, fonts["title"], config["accent_color"],
                  center=(W // 2, panel.y + 55))

        if subtitle:
            draw_text(screen, subtitle, fonts["body"], config["muted_color"],
                      center=(W // 2, panel.y + 92))

        sep_y = panel.y + 125
        pygame.draw.line(screen, config["border_color"],
                         (panel.x + 50, sep_y), (panel.right - 50, sep_y), 2)

        for i, line in enumerate(lines):
            draw_text(screen, line, fonts["med"], config["text_color"],
                      center=(W // 2, sep_y + 45 + i * 34))

        if wait_for_key:
            draw_text(screen, "Press  SPACE  to continue", fonts["body"],
                      config["accent_color"], center=(W // 2, panel.bottom - 40))

        pygame.display.flip()
        clock.tick(60)


# ============================================================
#  FIXATION
# ============================================================
def show_fixation(screen, config, font, duration_ms):
    if duration_ms <= 0:
        return
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()
    W, H = config["window_size"]

    while pygame.time.get_ticks() - start < duration_ms:
        screen.fill(config["bg_color"])
        draw_text(screen, "+", font, config["muted_color"],
                  center=(W // 2, H // 2))
        pygame.display.flip()
        clock.tick(60)
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()


# ============================================================
#  TRIAL
# ============================================================
def run_trial(screen, image, config, fonts, trial_num, total_trials,
              score, n_back, needs_response):
    drain_events()
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()

    responded = None
    rt_ms = None
    response_time = None

    image_dur = config["image_duration_ms"]
    rest_dur = config["rest_duration_ms"]
    total_dur = image_dur + rest_dur
    post_delay = config.get("post_response_delay_ms", 150)
    W, H = config["window_size"]

    btn_w, btn_h = config["btn_size"]
    gap = config["btn_gap"]
    btn_y = H - 115
    match_rect = pygame.Rect(W // 2 - btn_w - gap // 2, btn_y, btn_w, btn_h)
    nomatch_rect = pygame.Rect(W // 2 + gap // 2, btn_y, btn_w, btn_h)

    bar = pygame.Rect(220, 100, W - 440, 16)

    while True:
        now = pygame.time.get_ticks()
        elapsed = now - start

        if responded is not None and response_time is not None:
            if now - response_time >= post_delay:
                break
        else:
            if elapsed >= total_dur:
                break

        in_image_phase = elapsed < image_dur
        mouse_pos = pygame.mouse.get_pos()

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if needs_response and responded is None and in_image_phase:
                if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                    if match_rect.collidepoint(event.pos):
                        responded = "match"; rt_ms = elapsed; response_time = now
                    elif nomatch_rect.collidepoint(event.pos):
                        responded = "no_match"; rt_ms = elapsed; response_time = now
                elif event.type == pygame.KEYDOWN:
                    if event.key == pygame.K_LEFT:
                        responded = "match"; rt_ms = elapsed; response_time = now
                    elif event.key == pygame.K_RIGHT:
                        responded = "no_match"; rt_ms = elapsed; response_time = now

        screen.fill(config["bg_color"])

        header = pygame.Rect(35, 25, W - 70, 52)
        draw_panel(screen, header, config["panel_color"],
                   border=config["border_color"], radius=12, width=2)
        draw_text(screen, f"Trial  {trial_num} / {total_trials}",
                  fonts["body"], config["text_color"],
                  topleft=(header.x + 18, header.y + 16))
        draw_text(screen, f"N = {n_back}", fonts["body"],
                  config["accent_color"],
                  center=(header.centerx, header.y + 26))
        score_color = (config["success_color"] if score >= 0
                       else config["error_color"])
        draw_text(screen, f"Score  {score:+d}", fonts["body"], score_color,
                  topleft=(header.right - 130, header.y + 16))

        # ---- Status bar under the header ----
        if in_image_phase:
            if needs_response:
                remaining = (image_dur - elapsed) / image_dur
                draw_timer_bar(screen, bar, remaining,
                               config["success_color"], config["error_color"])
            else:
                # Warm-up trial: styled blue "MEMORIZE" strip
                strip_h = 28
                strip = pygame.Rect(
                    bar.x, bar.centery - strip_h // 2,
                    bar.width, strip_h
                )
                pygame.draw.rect(screen, config["panel_light"], strip,
                                 border_radius=14)
                pygame.draw.rect(screen, config["accent_color"], strip,
                                 width=2, border_radius=14)
                draw_text(screen,
                          "MEMORIZE  —  this image will be compared later",
                          fonts["small"], config["accent_color"],
                          center=strip.center)

        if in_image_phase:
            img_rect = image.get_rect(center=(W // 2, H // 2 - 20))
            frame = img_rect.inflate(14, 14)
            draw_panel(screen, frame, config["panel_light"],
                       border=config["border_color"], radius=10, width=2)
            screen.blit(image, img_rect)

            if needs_response:
                match_hover = match_rect.collidepoint(mouse_pos)
                nomatch_hover = nomatch_rect.collidepoint(mouse_pos)

                draw_button(screen, match_rect, "MATCH", fonts["small"],
                            fill_color=config["btn_match_fill"],
                            border_color=config["btn_match_border"],
                            selected=(responded == "match"),
                            hover=match_hover)
                draw_button(screen, nomatch_rect, "NO MATCH", fonts["small"],
                            fill_color=config["btn_nomatch_fill"],
                            border_color=config["btn_nomatch_border"],
                            selected=(responded == "no_match"),
                            hover=nomatch_hover)

                draw_text(screen,
                          "Click MATCH / NO MATCH   or   Left / Right arrow",
                          fonts["small"], config["muted_color"],
                          center=(W // 2, H - 42))
            else:
                draw_text(screen, "Watch carefully. No response needed.",
                          fonts["body"], config["muted_color"],
                          center=(W // 2, H - 52))
        else:
            draw_text(screen, "next image coming up...",
                      fonts["med"], config["muted_color"],
                      center=(W // 2, H // 2 - 20))

        pygame.display.flip()
        clock.tick(60)

    return responded, rt_ms


# ============================================================
#  FEEDBACK
# ============================================================
def show_feedback(screen, config, fonts, outcome, score_delta, duration_ms=700):
    if duration_ms <= 0:
        return
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()
    W, H = config["window_size"]

    label_map = {
        "hit":                ("HIT",             config["success_color"]),
        "correct_rejection":  ("CORRECT REJECT",  config["success_color"]),
        "miss":               ("MISS",            config["error_color"]),
        "false_alarm":        ("FALSE ALARM",     config["error_color"]),
        "timeout":            ("TOO SLOW",        config["warn_color"]),
    }
    label, color = label_map.get(outcome, ("?", config["text_color"]))

    while pygame.time.get_ticks() - start < duration_ms:
        screen.fill(config["bg_color"])
        draw_text(screen, label, fonts["hero"], color,
                  center=(W // 2, H // 2 - 30))
        if score_delta != 0:
            sign = "+" if score_delta > 0 else ""
            draw_text(screen, f"{sign}{score_delta}", fonts["title"], color,
                      center=(W // 2, H // 2 + 50))
        pygame.display.flip()
        clock.tick(60)
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()


# ============================================================
#  STAGE RUNNER
# ============================================================
def run_stage(screen, stage_design, config, folder_images, image_cache,
              fonts, stage_num, total_stages, running_score):
    target = stage_design["target"]
    filler = stage_design["filler"]
    label = stage_design["label"]
    n_back = config["n_back"]

    sequence = generate_stage_sequence(
        n_back=n_back,
        num_trials=config["trials_per_stage"],
        target_folder=target,
        filler_folder=filler,
        folder_images=folder_images,
        match_rate=config["match_rate"],
        target_rate=stage_design.get("target_rate",
                                     config["target_folder_rate"]),
        min_gap=config["min_gap"],
        min_match_gap=config["min_match_gap"],
    )

    show_message_screen(
        screen, config, fonts,
        title=f"Stage {stage_num} / {total_stages}",
        subtitle=label,
        lines=[
            f"The first {n_back} images are for memorization only.",
            f"From image {n_back + 1} onward:",
            "",
            f"Click MATCH if it equals the image from {n_back} step(s) ago.",
            "Click NO MATCH otherwise.",
            "Respond before the timer runs out.",
        ]
    )

    results = []
    stage_score = 0

    for i, trial in enumerate(sequence):
        is_warmup = (i < n_back)

        show_fixation(screen, config, fonts["huge"],
                      config["fixation_duration_ms"])

        responded, rt_ms = run_trial(
            screen, image_cache[trial["image"]], config, fonts,
            trial_num=i + 1,
            total_trials=len(sequence),
            score=running_score + stage_score,
            n_back=n_back,
            needs_response=not is_warmup,
        )

        if is_warmup:
            results.append({
                "stage": stage_num, "stage_label": label,
                "trial": i + 1, "folder": trial["folder"],
                "is_match": trial["is_match"],
                "responded": "warmup", "rt_ms": None,
                "outcome": "warmup", "score_delta": 0,
            })
            continue

        is_match = trial["is_match"]

        if responded is None:
            outcome = "timeout"
        elif is_match and responded == "match":
            outcome = "hit"
        elif is_match and responded == "no_match":
            outcome = "miss"
        elif not is_match and responded == "match":
            outcome = "false_alarm"
        else:
            outcome = "correct_rejection"

        if outcome in ("hit", "correct_rejection"):
            delta = config["score_correct"]
        elif outcome == "timeout":
            delta = config["score_timeout"]
        else:
            delta = config["score_incorrect"]

        stage_score += delta

        show_feedback(screen, config, fonts, outcome, delta,
                      duration_ms=config["feedback_duration_ms"])
        pygame.time.wait(config["inter_trial_pause_ms"])

        results.append({
            "stage": stage_num, "stage_label": label,
            "trial": i + 1, "folder": trial["folder"],
            "is_match": is_match,
            "responded": responded if responded is not None else "none",
            "rt_ms": rt_ms, "outcome": outcome,
            "score_delta": delta,
        })

    return results, stage_score


# ============================================================
#  STATS
# ============================================================
def compute_stats(results):
    results = [r for r in results if r["outcome"] != "warmup"]
    if not results:
        return {
            "n_trials": 0, "hits": 0, "misses": 0, "false_alarms": 0,
            "correct_rejections": 0, "timeouts": 0, "accuracy": 0.0,
            "hit_rate": 0.5, "fa_rate": 0.5, "d_prime": 0.0,
            "criterion": 0.0, "mean_rt_ms": None, "total_score": 0,
        }

    hits = sum(1 for r in results if r["outcome"] == "hit")
    misses = sum(1 for r in results if r["outcome"] == "miss")
    fas = sum(1 for r in results if r["outcome"] == "false_alarm")
    crs = sum(1 for r in results if r["outcome"] == "correct_rejection")
    timeouts = sum(1 for r in results if r["outcome"] == "timeout")

    n = len(results)
    accuracy = (hits + crs) / n if n else 0.0

    for r in results:
        if r["outcome"] == "timeout":
            if r["is_match"]:
                misses += 1
            else:
                crs += 1

    hit_rate = (hits + 0.5) / (hits + misses + 1)
    fa_rate = (fas + 0.5) / (fas + crs + 1)

    nd = NormalDist()
    d_prime = nd.inv_cdf(hit_rate) - nd.inv_cdf(fa_rate)
    criterion = -0.5 * (nd.inv_cdf(hit_rate) + nd.inv_cdf(fa_rate))

    rts = [r["rt_ms"] for r in results
           if r["outcome"] == "hit" and r["rt_ms"] is not None]
    mean_rt = mean(rts) if rts else None

    total_score = sum(r["score_delta"] for r in results)

    return {
        "n_trials": n, "hits": hits, "misses": misses,
        "false_alarms": fas, "correct_rejections": crs,
        "timeouts": timeouts, "accuracy": accuracy,
        "hit_rate": hit_rate, "fa_rate": fa_rate,
        "d_prime": d_prime, "criterion": criterion,
        "mean_rt_ms": mean_rt, "total_score": total_score,
    }


# ============================================================
#  RESULTS SCREEN
# ============================================================
def show_results_screen(screen, config, fonts, stage_stats, overall_stats):
    drain_events()
    W, H = config["window_size"]
    clock = pygame.time.Clock()

    ask_btn_w, ask_btn_h = 200, 60
    ask_gap = 40
    ask_btn_y = 610
    yes_rect = pygame.Rect(W // 2 - ask_btn_w - ask_gap // 2, ask_btn_y,
                           ask_btn_w, ask_btn_h)
    no_rect  = pygame.Rect(W // 2 + ask_gap // 2, ask_btn_y,
                           ask_btn_w, ask_btn_h)

    while True:
        mouse_pos = pygame.mouse.get_pos()

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                return "exit"

            if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                if yes_rect.collidepoint(event.pos):
                    return "restart"
                if no_rect.collidepoint(event.pos):
                    return "exit"

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_SPACE:
                    return "exit"
                if event.key == pygame.K_RETURN:
                    return "restart"

        screen.fill(config["bg_color"])

        draw_text(screen, "Test Complete", fonts["title"],
                  config["accent_color"], center=(W // 2, 40))

        score_panel = pygame.Rect(50, 75, W - 100, 90)
        draw_panel(screen, score_panel, config["panel_color"],
                   border=config["border_color"], radius=16, width=2)
        score_color = (config["success_color"]
                       if overall_stats["total_score"] >= 0
                       else config["error_color"])
        sign = "+" if overall_stats["total_score"] >= 0 else ""
        draw_text(screen, f"Total Score:  {sign}{overall_stats['total_score']}",
                  fonts["title"], score_color,
                  center=(W // 2, score_panel.centery))

        left_panel = pygame.Rect(50, 185, (W - 130) // 2, 300)
        draw_panel(screen, left_panel, config["panel_color"],
                   border=config["border_color"], radius=16, width=2)
        draw_text(screen, "Overall Statistics", fonts["large"],
                  config["accent_color"],
                  topleft=(left_panel.x + 18, left_panel.y + 16))

        y = left_panel.y + 65
        rows = [
            ("Accuracy",           f"{overall_stats['accuracy']*100:.1f}%"),
            ("d' (sensitivity)",   f"{overall_stats['d_prime']:.3f}"),
            ("Criterion",          f"{overall_stats['criterion']:.3f}"),
        ]
        if overall_stats["mean_rt_ms"] is not None:
            rows.append(("Mean RT (hits)", f"{overall_stats['mean_rt_ms']:.0f} ms"))
        rows.append(("Total trials",   f"{overall_stats['n_trials']}"))

        for label, value in rows:
            draw_text(screen, label, fonts["body"], config["muted_color"],
                      topleft=(left_panel.x + 18, y))
            draw_text(screen, value, fonts["med"], config["text_color"],
                      topleft=(left_panel.right - 140, y))
            y += 42

        right_panel = pygame.Rect(50 + (W - 130) // 2 + 30, 185,
                                   (W - 130) // 2, 300)
        draw_panel(screen, right_panel, config["panel_color"],
                   border=config["border_color"], radius=16, width=2)
        draw_text(screen, "Per-Stage Breakdown", fonts["large"],
                  config["accent_color"],
                  topleft=(right_panel.x + 18, right_panel.y + 16))

        y = right_panel.y + 65
        for label, st in stage_stats.items():
            draw_text(screen, label[:22], fonts["body"], config["text_color"],
                      topleft=(right_panel.x + 18, y))
            score_str = f"{st['total_score']:+d}   {st['accuracy']*100:.0f}%"
            draw_text(screen, score_str, fonts["body"], config["muted_color"],
                      topleft=(right_panel.right - 130, y))
            y += 38

            bar = pygame.Rect(right_panel.x + 18, y, right_panel.width - 36, 8)
            pygame.draw.rect(screen, (40, 44, 58), bar, border_radius=4)
            acc = st["accuracy"]
            bar_color = (config["success_color"] if acc >= 0.7
                         else config["warn_color"] if acc >= 0.5
                         else config["error_color"])
            filled = pygame.Rect(bar.x, bar.y, int(bar.width * acc), bar.height)
            if acc > 0:
                pygame.draw.rect(screen, bar_color, filled, border_radius=4)
            y += 28

        draw_text(screen, "Start a new test?",
                  fonts["huge"], config["text_color"],
                  center=(W // 2, 555))

        yes_hover = yes_rect.collidepoint(mouse_pos)
        no_hover  = no_rect.collidepoint(mouse_pos)

        blue_fill = config["btn_match_fill"]
        blue_border = config["btn_match_border"]

        draw_button(screen, yes_rect, "Yes", fonts["med"],
                    fill_color=blue_fill, border_color=blue_border,
                    selected=False, hover=yes_hover)

        draw_button(screen, no_rect, "No", fonts["med"],
                    fill_color=None, border_color=blue_border,
                    selected=False, hover=no_hover)

        draw_text(screen, "Enter: Yes    Space: No",
                  fonts["hint"], config["muted_color"],
                  center=(W // 2, 700))

        pygame.display.flip()
        clock.tick(60)


# ============================================================
#  SAVE RESULTS (CSV detail log)
# ============================================================
def save_results(all_results, config):
    if not config["save_results"] or not all_results:
        return
    out_dir = BASE_DIR / config["results_dir"]
    out_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"nback_{ts}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(all_results[0].keys()))
        writer.writeheader()
        writer.writerows(all_results)
    print(f"[CSV] Detailed trial log saved to: {path}")


# ============================================================
#  SAVE TO EXCEL (Persian headers — comprehensive stats)
# ============================================================
def _fmt_or_blank(v, digits=0):
    if v is None or v == "":
        return ""
    if digits == 0:
        return round(v, 0)
    return round(v, digits)

def _show_excel_locked_popup(screen, config, fonts, fallback_path):
    """نمایش پاپ‌آپ هشدار وقتی فایل Excel قفل باشه."""
    if screen is None:
        return
    clock = pygame.time.Clock()
    W, H = config["window_size"]
    panel_w, panel_h = 640, 280
    panel = pygame.Rect((W - panel_w) // 2, (H - panel_h) // 2,
                        panel_w, panel_h)
    btn_rect = pygame.Rect(W // 2 - 90, panel.bottom - 70, 180, 44)

    while True:
        mouse_pos = pygame.mouse.get_pos()
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                return
            if event.type == pygame.KEYDOWN and event.key in (
                    pygame.K_SPACE, pygame.K_RETURN, pygame.K_ESCAPE):
                return
            if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
                if btn_rect.collidepoint(event.pos):
                    return

        screen.fill(config["bg_color"])
        draw_panel(screen, panel, config["panel_color"],
                   border=config["warn_color"], radius=18, width=2)

        draw_text(screen, "Excel File Is Locked",
                  fonts["large"], config["warn_color"],
                  center=(W // 2, panel.y + 45))

        pygame.draw.line(screen, config["border_color"],
                         (panel.x + 30, panel.y + 75),
                         (panel.right - 30, panel.y + 75), 2)

        draw_text(screen, "The file participants.xlsx is open in Excel",
                  fonts["body"], config["text_color"],
                  center=(W // 2, panel.y + 105))
        draw_text(screen, "or another program, so it cannot be updated.",
                  fonts["body"], config["text_color"],
                  center=(W // 2, panel.y + 130))
        draw_text(screen, "Results were saved to a new file:",
                  fonts["body"], config["muted_color"],
                  center=(W // 2, panel.y + 168))
        draw_text(screen, fallback_path.name,
                  fonts["small"], config["accent_color"],
                  center=(W // 2, panel.y + 198))

        hover = btn_rect.collidepoint(mouse_pos)
        fill = (tuple(min(255, c + 30) for c in config["accent_color"])
                if hover else config["accent_color"])
        pygame.draw.rect(screen, fill, btn_rect, border_radius=16)
        draw_text(screen, "OK", fonts["med"], (255, 255, 255),
                  center=btn_rect.center)

        pygame.display.flip()
        clock.tick(60)

def save_to_excel(participant, stage_stats, overall_stats, config):
    if not HAS_OPENPYXL:
        print("[WARN] openpyxl not installed. Install with: pip install openpyxl")
        _save_to_csv_fallback(participant, stage_stats, overall_stats)
        return None

    path = BASE_DIR / config["excel_filename"]

    stage_label_fa = {
        "Attractive Faces":   "چهره‌های جذاب",
        "Unattractive Faces": "چهره‌های غیرجذاب",
        "Neutral Faces":      "چهره‌های خنثی",
    }
    gender_fa = {"Male": "مرد", "Female": "زن"}

    headers = [
        "زمان ثبت",
        "نام و نام خانوادگی",
        "رشته تحصیلی",
        "سن",
        "جنسیت",
        "سطح N",

        "امتیاز کل",
        "تعداد کل تریال‌ها",
        "Match صحیح (Hit)",
        "Miss (از دست رفته)",
        "هشدار نادرست (False Alarm)",
        "رد صحیح (Correct Rejection)",
        "بی‌پاسخ (Timeout)",
        "دقت کل (%)",
        "d' کل",
        "معیار پاسخ کل",
        "میانگین زمان واکنش Hit (ms)",
    ]

    for i in range(1, 4):
        headers.extend([
            f"مرحله {i} - نام",
            f"مرحله {i} - امتیاز",
            f"مرحله {i} - دقت (%)",
            f"مرحله {i} - Match صحیح",
            f"مرحله {i} - Miss",
            f"مرحله {i} - هشدار نادرست",
            f"مرحله {i} - رد صحیح",
            f"مرحله {i} - بی‌پاسخ",
            f"مرحله {i} - d'",
            f"مرحله {i} - میانگین RT (ms)",
        ])

    mean_rt = overall_stats["mean_rt_ms"]
    stage_list = list(stage_stats.items())

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        participant["full_name"],
        participant["field_of_study"],
        participant["age"],
        gender_fa.get(participant["gender"], participant["gender"]),
        config["n_back"],

        overall_stats["total_score"],
        overall_stats["n_trials"],
        overall_stats["hits"],
        overall_stats["misses"],
        overall_stats["false_alarms"],
        overall_stats["correct_rejections"],
        overall_stats["timeouts"],
        round(overall_stats["accuracy"] * 100, 2),
        round(overall_stats["d_prime"], 3),
        round(overall_stats["criterion"], 3),
        _fmt_or_blank(mean_rt, 0),
    ]

    for i in range(3):
        if i < len(stage_list):
            label, st = stage_list[i]
            label_display = stage_label_fa.get(label, label)
            st_rt = st["mean_rt_ms"]
            row.extend([
                label_display,
                st["total_score"],
                round(st["accuracy"] * 100, 2),
                st["hits"],
                st["misses"],
                st["false_alarms"],
                st["correct_rejections"],
                st["timeouts"],
                round(st["d_prime"], 3),
                _fmt_or_blank(st_rt, 0),
            ])
        else:
            row.extend([""] * 10)

    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
        if ws.max_column != len(headers):
            for col in range(ws.max_column, 0, -1):
                ws.cell(row=1, column=col).value = None
            for i, h in enumerate(headers, start=1):
                ws.cell(row=1, column=i, value=h)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Participants"
        for i, h in enumerate(headers, start=1):
            ws.cell(row=1, column=i, value=h)

    ws.sheet_view.rightToLeft = True

    header_fill = PatternFill("solid", fgColor="DDEBFF")
    header_align = Alignment(horizontal="center", vertical="center",
                             readingOrder=2, wrap_text=True)

    for cell in ws[1]:
        cell.font = XLFont(bold=True)
        cell.fill = header_fill
        cell.alignment = header_align

    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "A2"

    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(h) + 2)

    data_align = Alignment(horizontal="right", vertical="center",
                           readingOrder=2)
    new_row_idx = ws.max_row + 1

    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=new_row_idx, column=col_idx, value=value)
        cell.alignment = data_align

    try:
        highlight_fill = PatternFill("solid", fgColor="FFF4CC")
        for col in (7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17):
            ws.cell(row=new_row_idx, column=col).fill = highlight_fill
    except Exception:
        pass

    try:
        wb.save(path)
        print(f"[Excel] Participant row appended to: {path}")
        return path

    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = BASE_DIR / f"{path.stem}_locked_{ts}{path.suffix}"
        try:
            wb.save(fallback)
            print("[Excel] Original file is locked (open in Excel?).")
            print(f"[Excel] Saved to fallback file instead: {fallback}")
            return fallback
        except Exception as e2:
            print(f"[Excel] Fallback save also failed: {e2}")
            _save_to_csv_fallback(participant, stage_stats, overall_stats)
            return None

    except Exception as e:
        print(f"[Excel] Failed to save Excel: {e}")
        _save_to_csv_fallback(participant, stage_stats, overall_stats)
        return None


def _save_to_csv_fallback(participant, stage_stats, overall_stats):
    path = BASE_DIR / "participants.csv"
    headers = ["زمان ثبت", "نام و نام خانوادگی", "رشته تحصیلی", "سن", "جنسیت",
               "امتیاز کل", "دقت", "d_prime", "معیار",
               "Match صحیح", "Miss", "False Alarm", "Correct Rejection",
               "Timeout"]
    new_file = not path.exists()
    gender_fa = {"Male": "مرد", "Female": "زن"}
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(headers)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            participant["full_name"],
            participant["field_of_study"],
            participant["age"],
            gender_fa.get(participant["gender"], participant["gender"]),
            overall_stats["total_score"],
            round(overall_stats["accuracy"] * 100, 2),
            round(overall_stats["d_prime"], 3),
            round(overall_stats["criterion"], 3),
            overall_stats["hits"],
            overall_stats["misses"],
            overall_stats["false_alarms"],
            overall_stats["correct_rejections"],
            overall_stats["timeouts"],
        ])
    print(f"[CSV fallback] Participant row appended to: {path}")


# ============================================================
#  MAIN
# ============================================================
def main():
    load_settings()

    validate_config(CONFIG)

    pygame.init()
    pygame.display.set_caption("N-Back Test")
    screen = pygame.display.set_mode(CONFIG["window_size"])

    fonts = make_fonts()
    folder_images = {name: load_folder_images(name) for name in FOLDER_PATHS}
    image_cache = preload_images(folder_images, CONFIG["image_max_size"])

    while True:
        participant = show_participant_form(screen, CONFIG, fonts)

        validate_config(CONFIG)

        show_message_screen(
            screen, CONFIG, fonts,
            title="N-Back Test",
            subtitle="Cognitive Working-Memory Assessment",
            lines=[
                f"Each trial shows one image for {CONFIG['image_duration_ms']/1000:.1f} seconds.",
                f"Click MATCH if the current image equals",
                f"the image shown {CONFIG['n_back']} step(s) earlier.",
                "Click NO MATCH otherwise.",
                f"3 stages  -  {CONFIG['trials_per_stage']} trials each  -  random order",
            ]
        )

        if CONFIG.get("stage_mode", "random") == "manual":
            stages = []
            for cfg in CONFIG["manual_stages"]:
                stages.append({
                    "target": cfg["target"],
                    "filler": cfg["filler"],
                    "label": FOLDER_LABELS[cfg["target"]],
                    "target_rate": cfg["target_rate"],
                })
        else:
            stages = copy.deepcopy(STAGE_DESIGNS)
            random.shuffle(stages)

        all_results = []
        stage_stats = {}
        running_score = 0

        for idx, stage in enumerate(stages, start=1):
            results, stage_score = run_stage(
                screen, stage, CONFIG, folder_images, image_cache, fonts,
                stage_num=idx, total_stages=len(stages),
                running_score=running_score,
            )
            running_score += stage_score
            all_results.extend(results)
            stage_stats[stage["label"]] = compute_stats(results)

        overall = compute_stats(all_results)
        overall["total_score"] = running_score

        save_results(all_results, CONFIG)

        saved_path = save_to_excel(participant, stage_stats, overall, CONFIG)
        expected_path = BASE_DIR / CONFIG["excel_filename"]
        if saved_path is not None and saved_path != expected_path:
            _show_excel_locked_popup(screen, CONFIG, fonts, saved_path)

        action = show_results_screen(screen, CONFIG, fonts,
                                     stage_stats, overall)

        if action == "exit":
            break

    pygame.quit()


if __name__ == "__main__":
    main()