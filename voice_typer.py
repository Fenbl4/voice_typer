"""
Voice Typer — fast local voice-to-text with optional cloud fallback.

Two modes:
  Push-to-Talk: Hold hotkey to record, release to transcribe & paste.
  Voice Activated: Detects speech automatically, pastes on silence.

Runs in the system tray.
"""

import sys
import os
import threading
import io
import wave
import json
import time
import math
import struct
import re
import hashlib
import urllib.request
import ctypes
import ctypes.wintypes
import winreg
from datetime import datetime

import tkinter as tk
import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw

import pyaudio
import pyperclip
import pyautogui
from pynput import keyboard
from google import genai
from google.genai import types as genai_types

try:
    from groq import Groq
    GROQ_AVAILABLE = True
except ImportError:
    GROQ_AVAILABLE = False
    Groq = None

try:
    import numpy as np
    import onnx_asr
    import onnxruntime as ort
    LOCAL_ASR_AVAILABLE = True
except ImportError:
    np = None
    onnx_asr = None
    ort = None
    LOCAL_ASR_AVAILABLE = False


# =============================================================================
# Configuration
# =============================================================================

LOCAL_PROVIDER = "Local"
AVAILABLE_PROVIDERS = [LOCAL_PROVIDER, "Groq", "Gemini"]
DEFAULT_PROVIDER = LOCAL_PROVIDER
LOCAL_MODEL_NAME = "gigaam-v3-e2e-rnnt"
LOCAL_MODEL_LABEL = "GigaAM v3 E2E RNN-T (RU)"
LOCAL_MODEL_QUANTIZATION = "int8"
LOCAL_DIRECT_MAX_SECONDS = 23.0
LOCAL_MODEL_WAIT_SECONDS = 2.0
# Замер 2026-09-02 на 340 с речи: 4 потока распознают так же быстро, как 8 (14.8 с против 15.5 с),
# а без спин-ожидания свободные ядра не крутятся вхолостую между запусками модели.
LOCAL_ASR_THREADS = 4
LOCAL_VAD_OPTIONS = {
    "max_speech_duration_s": 22.0,
    "min_speech_duration_ms": 250.0,
    "min_silence_duration_ms": 300.0,
    "speech_pad_ms": 120.0,
}
# Живое распознавание во время записи Push-to-Talk: закрытые речевые отрезки распознаются,
# пока клавиша ещё зажата, и после отпускания остаётся только хвост длиной до ~22 с.
# Порог 20 с оставляет короткие записи целыми, как раньше: резать их незачем, а разрез на заминке
# внутри слова давал обрывок («бу-у-у» вместо «будет», 2026-09-02).
LIVE_MIN_PENDING_SECONDS = 20.0   # столько нераспознанного звука копится перед новой нарезкой
LIVE_POLL_SECONDS = 0.5
LIVE_MIN_GAP_SECONDS = 1.0        # пауза короче этой считается заминкой внутри фразы, а не границей
LIVE_TAIL_GUARD_SECONDS = 1.0     # отрезок закрыт, если после него записано хотя бы столько звука
LIVE_JOIN_TIMEOUT_SECONDS = 30.0
SETTINGS_VERSION = 2

LOCAL_MODEL_REPO = "istupakov/gigaam-v3-onnx"
LOCAL_MODEL_REVISION = "322c3b29492673eb7d0b434bfa9dfb8653e34d02"
LOCAL_MODEL_FILES = {
    "config.json": "0641fcf73f4af791c73f05083e38a658ff5dcbee3534a0c61396c10f4b97f0fe",
    "v3_e2e_rnnt_decoder.int8.onnx": "89014e134865615b91e037157e46e389b1271e6072460efc010ea08e61e23146",
    "v3_e2e_rnnt_encoder.int8.onnx": "4e0e076a6076cd110277e529b8ac8f32cd5297f7fbebad5341ae8ddb7d00817b",
    "v3_e2e_rnnt_joint.int8.onnx": "ade116563dbf66e503b0994efab6b5861412743e52bf31c39fc3fffa3783d5d1",
    "v3_e2e_rnnt_vocab.txt": "39abae20e692998290c574e606f11a9edef2902a1995463fcff63d1490cf22b7",
}
LOCAL_VAD_REPO = "istupakov/silero-vad-onnx"
LOCAL_VAD_REVISION = "b3e3ee3cce4c11ceb63b1a0b229d916069c1ddf6"
LOCAL_VAD_FILES = {
    "config.json": "1094039d370c82889582ba739a3d1caac5754c8b3a17a66a534200c9f72086e2",
    "silero_vad.onnx": "1a153a22f4509e292a94e67d6f9b85e8deb25b4988682b7e174c65279d8788e3",
}

GEMINI_MODELS = [
    "gemini-3-flash-preview",   # Gemini 3 Flash (newest, preview)
    "gemini-2.5-flash",        # Gemini 2.5 Flash (stable, default)
    "gemini-2.5-flash-lite",   # Gemini 2.5 Flash Lite (lightweight)
]
DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"

GROQ_MODELS = [
    "whisper-large-v3-turbo",  # $0.04/hr — fast + high quality (default)
    "whisper-large-v3",        # $0.111/hr — max quality
]
DEFAULT_GROQ_MODEL = "whisper-large-v3-turbo"

# Legacy aliases для обратной совместимости
AVAILABLE_MODELS = GEMINI_MODELS
DEFAULT_MODEL = DEFAULT_GEMINI_MODEL

SAMPLE_RATE = 16000
CHANNELS = 1
CHUNK_SIZE = 1024
AUDIO_FORMAT = pyaudio.paInt16

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
    EXE_PATH = sys.executable
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_PATH = sys.executable

SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
HISTORY_FILE = os.path.join(APP_DIR, "history.json")
APP_NAME = "VoiceTyper"
WINDOW_TITLE = "Voice Typer"
LOCAL_DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", APP_DIR), APP_NAME)
LOCAL_MODEL_DIR = os.path.join(LOCAL_DATA_DIR, "models", LOCAL_MODEL_NAME)
LOCAL_VAD_DIR = os.path.join(LOCAL_DATA_DIR, "models", "silero-vad")

BASE_SYSTEM_INSTRUCTION = (
    "You are a professional transcriber. Transcribe the provided audio with perfect "
    "grammar and punctuation. Output ONLY the raw transcribed text — no quotes, no "
    "explanations, no markdown formatting, no commentary. Fix obvious speech errors. "
    "Preserve the original language of the speaker. "
    "IMPORTANT: Always use the native script of the language — Russian and Ukrainian "
    "must be in Cyrillic, never transliterate to Latin characters."
)

VAD_SPEECH_THRESHOLD = 300
VAD_SILENCE_TIMEOUT = 1.5
VAD_MIN_SPEECH_DURATION = 0.3

GROQ_HALLUCINATION_TAILS = (
    "продолжение следует",
    "спасибо за просмотр",
    "субтитры сделал",
    "субтитры делал",
)

# Gemini free tier: no cooldown (was 6.0s)
MIN_API_INTERVAL = 0

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.02

# =============================================================================
# UI Localization
# =============================================================================

UI_STRINGS = {
    "English": {
        "subtitle": "AI voice transcription",
        "api_key_lbl": "API Key:",
        "paste_btn": "Paste", "show_btn": "Show", "hide_btn": "Hide",
        "settings_lbl": "Settings",
        "mode_lbl": "Mode:",
        "provider_lbl": "Provider:",
        "model_lbl": "Model:",
        "hotkey_lbl": "Hotkey:", "hotkey_change": "⌨ Change",
        "language_lbl": "Language:",
        "autostart_chk": "Start with Windows",
        "history_btn": "History", "logs_btn": "Logs",
        "copy_sel": "Copy", "select_all": "Select All",
        "hint_ptt": "Hold  [ {name} ]  to record   |   Release to transcribe & paste",
        "hint_vad": "Voice Activated — just speak, it records automatically",
        "status_ready": "Ready",
        "status_recording": "Recording...",
        "status_processing": "Processing...",
        "status_pasted": "Pasted!",
        "status_no_audio": "No audio.",
        "status_api_key": "Set API key first!",
        "status_api_error": "API error",
        "status_clipboard": "Clipboard error!",
        "status_empty": "Empty response.",
        "status_ready_in": "Ready in {n}s",
        "status_retry": "Retry {attempt}/{max} in {n}s...",
        "status_cancelled": "Cancelled — record again",
        "status_wait": "Please wait...",
        "status_model_loading": "Preparing local speech model...",
        "status_model_error": "Local speech model error",
        "status_fallback": "Local model unavailable. Using {provider}...",
        "local_key_hint": "No API key needed",
        "capture_prompt": "Press keys...",
        "no_history": "No transcriptions yet.",
        "no_logs": "No logs yet. Errors will appear here.",
    },
    "Russian": {
        "subtitle": "ИИ голосовой ввод",
        "api_key_lbl": "API ключ:",
        "paste_btn": "Вставить", "show_btn": "Показать", "hide_btn": "Скрыть",
        "settings_lbl": "Настройки",
        "mode_lbl": "Режим:",
        "provider_lbl": "Провайдер:",
        "model_lbl": "Модель:",
        "hotkey_lbl": "Хоткей:", "hotkey_change": "⌨ Изменить",
        "language_lbl": "Язык:",
        "autostart_chk": "Запуск с Windows",
        "history_btn": "История", "logs_btn": "Логи",
        "copy_sel": "Копировать", "select_all": "Выделить всё",
        "hint_ptt": "Держи  [ {name} ]  для записи   |   Отпусти для вставки",
        "hint_vad": "Голосовая активация — просто говори",
        "status_ready": "Готово",
        "status_recording": "Запись...",
        "status_processing": "Обработка...",
        "status_pasted": "Вставлено!",
        "status_no_audio": "Нет аудио.",
        "status_api_key": "Сначала введи API ключ!",
        "status_api_error": "Ошибка API",
        "status_clipboard": "Ошибка буфера!",
        "status_empty": "Пустой ответ.",
        "status_ready_in": "Готово через {n}с",
        "status_retry": "Повтор {attempt}/{max} через {n}с...",
        "status_cancelled": "Отменено — запиши снова",
        "status_wait": "Подождите...",
        "status_model_loading": "Подготовка локальной модели...",
        "status_model_error": "Ошибка локальной модели",
        "status_fallback": "Локальная модель недоступна. Используется {provider}...",
        "local_key_hint": "API ключ не нужен",
        "capture_prompt": "Нажмите...",
        "no_history": "Записей пока нет.",
        "no_logs": "Логов пока нет.",
    },
    "Ukrainian": {
        "subtitle": "ШІ голосовий ввід",
        "api_key_lbl": "API ключ:",
        "paste_btn": "Вставити", "show_btn": "Показати", "hide_btn": "Приховати",
        "settings_lbl": "Налаштування",
        "mode_lbl": "Режим:",
        "provider_lbl": "Провайдер:",
        "model_lbl": "Модель:",
        "hotkey_lbl": "Хоткей:", "hotkey_change": "⌨ Змінити",
        "language_lbl": "Мова:",
        "autostart_chk": "Запуск з Windows",
        "history_btn": "Історія", "logs_btn": "Логи",
        "copy_sel": "Копіювати", "select_all": "Виділити все",
        "hint_ptt": "Тримай  [ {name} ]  для запису   |   Відпусти для вставки",
        "hint_vad": "Голосова активація — просто говори",
        "status_ready": "Готово",
        "status_recording": "Запис...",
        "status_processing": "Обробка...",
        "status_pasted": "Вставлено!",
        "status_no_audio": "Немає аудіо.",
        "status_api_key": "Спочатку введи API ключ!",
        "status_api_error": "Помилка API",
        "status_clipboard": "Помилка буфера!",
        "status_empty": "Порожня відповідь.",
        "status_ready_in": "Готово через {n}с",
        "status_retry": "Повтор {attempt}/{max} через {n}с...",
        "status_cancelled": "Скасовано — запиши знову",
        "status_wait": "Зачекайте...",
        "status_model_loading": "Підготовка локальної моделі...",
        "status_model_error": "Помилка локальної моделі",
        "status_fallback": "Локальна модель недоступна. Використовується {provider}...",
        "local_key_hint": "API ключ не потрібен",
        "capture_prompt": "Натисніть...",
        "no_history": "Записів поки немає.",
        "no_logs": "Логів поки немає.",
    },
}

# =============================================================================
# Modifier key sets for pynput
# =============================================================================

_MODIFIER_KEYS = {
    keyboard.Key.ctrl_l, keyboard.Key.ctrl_r,
    keyboard.Key.alt_l, keyboard.Key.alt_r,
    keyboard.Key.shift_l, keyboard.Key.shift_r,
}

_SPECIAL_KEY_NAMES = {
    keyboard.Key.f1: "F1", keyboard.Key.f2: "F2", keyboard.Key.f3: "F3",
    keyboard.Key.f4: "F4", keyboard.Key.f5: "F5", keyboard.Key.f6: "F6",
    keyboard.Key.f7: "F7", keyboard.Key.f8: "F8", keyboard.Key.f9: "F9",
    keyboard.Key.f10: "F10", keyboard.Key.f11: "F11", keyboard.Key.f12: "F12",
    keyboard.Key.space: "Space", keyboard.Key.tab: "Tab",
    keyboard.Key.enter: "Enter", keyboard.Key.backspace: "Backspace",
    keyboard.Key.delete: "Delete", keyboard.Key.insert: "Insert",
    keyboard.Key.home: "Home", keyboard.Key.end: "End",
    keyboard.Key.page_up: "PageUp", keyboard.Key.page_down: "PageDown",
    keyboard.Key.up: "Up", keyboard.Key.down: "Down",
    keyboard.Key.left: "Left", keyboard.Key.right: "Right",
    keyboard.Key.pause: "Pause", keyboard.Key.scroll_lock: "ScrollLock",
    keyboard.Key.print_screen: "PrintScreen", keyboard.Key.caps_lock: "CapsLock",
    keyboard.Key.num_lock: "NumLock",
}


# =============================================================================
# Hotkey helpers
# =============================================================================

def _is_modifier(key) -> bool:
    return key in _MODIFIER_KEYS


def _modifier_name(key) -> str:
    if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
        return "Ctrl"
    if key in (keyboard.Key.alt_l, keyboard.Key.alt_r):
        return "Alt"
    if key in (keyboard.Key.shift_l, keyboard.Key.shift_r):
        return "Shift"
    return ""


def _key_to_name(key) -> str:
    if key in _SPECIAL_KEY_NAMES:
        return _SPECIAL_KEY_NAMES[key]
    if hasattr(key, "char") and key.char:
        return key.char.upper()
    if hasattr(key, "vk") and key.vk:
        return f"VK{key.vk}"
    return str(key)


def _parse_hotkey_str(hotkey_str: str):
    parts = [p.strip() for p in hotkey_str.split("+")]
    mod_names = {"Ctrl", "Alt", "Shift"}
    modifiers = set()
    key_parts = []
    for p in parts:
        if p in mod_names:
            modifiers.add(p)
        else:
            key_parts.append(p)
    key_name = "+".join(key_parts) if key_parts else ""
    key_obj = None
    for k, name in _SPECIAL_KEY_NAMES.items():
        if name.lower() == key_name.lower():
            key_obj = k
            break
    if key_obj is None and len(key_name) == 1:
        key_obj = keyboard.KeyCode.from_char(key_name.lower())
    if key_obj is None:
        key_obj = keyboard.Key.f8
    return key_obj, modifiers


# =============================================================================
# Helpers
# =============================================================================

TRAY_COLORS = {
    "idle":       (59, 142, 208, 255),   # CTK blue
    "recording":  (220, 53, 53, 255),    # red
    "processing": (230, 126, 14, 255),   # orange
}


def _create_tray_image(state: str = "idle"):
    """Тёмный квадрат + цветной внутренний квадрат — в стиле CTK иконки."""
    inner_color = TRAY_COLORS.get(state, TRAY_COLORS["idle"])
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    bg = (30, 33, 48, 255)  # тёмно-синий фон, как у CTK
    r = 10  # радиус скругления
    try:
        draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=bg)
        draw.rounded_rectangle([11, 11, size - 12, size - 12], radius=6, fill=inner_color)
    except AttributeError:
        # Pillow < 8.2 — без скруглений
        draw.rectangle([0, 0, size - 1, size - 1], fill=bg)
        draw.rectangle([11, 11, size - 12, size - 12], fill=inner_color)
    return img


def _generate_ico_file():
    """Генерировать icon.ico в стиле CTK (тёмный квадрат + синий внутренний)."""
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    bg_color = (30, 33, 48, 255)       # тёмно-синий фон, как у CTK
    inner_color = (59, 142, 208, 255)  # CTK синий
    frames = []
    for w, h in sizes:
        frame = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(frame)
        radius = max(2, w // 10)
        margin = max(2, w // 14)
        inner_r = max(2, w // 16)
        try:
            d.rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=bg_color)
            d.rounded_rectangle(
                [margin, margin, w - margin - 1, h - margin - 1],
                radius=inner_r, fill=inner_color,
            )
        except AttributeError:
            d.rectangle([0, 0, w - 1, h - 1], fill=bg_color)
            d.rectangle([margin, margin, w - margin - 1, h - margin - 1], fill=inner_color)
        frames.append(frame)
    frames[0].save(out, format="ICO", append_images=frames[1:], sizes=sizes)
    print(f"Icon saved: {out}")


def _rms(data: bytes) -> float:
    count = len(data) // 2
    if count == 0:
        return 0.0
    shorts = struct.unpack(f"<{count}h", data)
    return math.sqrt(sum(s * s for s in shorts) / count)


def _trim_pcm_silence(pcm_bytes: bytes) -> bytes:
    """Обрезать только крайние тишину и шум, сохранив небольшой запас вокруг речи."""
    if not pcm_bytes or np is None:
        return pcm_bytes

    samples = np.frombuffer(pcm_bytes, dtype="<i2")
    if samples.size < int(SAMPLE_RATE * 0.15):
        return b""

    frame_samples = int(SAMPLE_RATE * 0.02)
    frame_count = (samples.size + frame_samples - 1) // frame_samples
    padded = np.pad(samples, (0, frame_count * frame_samples - samples.size))
    frames = padded.reshape(frame_count, frame_samples).astype(np.float32)
    levels = np.sqrt(np.mean(frames * frames, axis=1))

    noise_floor = float(np.percentile(levels, 20))
    threshold = max(90.0, min(350.0, noise_floor * 2.5))
    active = np.flatnonzero(levels >= threshold)
    if active.size == 0:
        return b""

    lead_padding = int(SAMPLE_RATE * 0.12)
    tail_padding = int(SAMPLE_RATE * 0.18)
    start = max(0, int(active[0]) * frame_samples - lead_padding)
    end = min(samples.size, (int(active[-1]) + 1) * frame_samples + tail_padding)
    return samples[start:end].tobytes()


def _merge_close_segments(segments, min_gap: int, max_length: int) -> list:
    """Склеить соседние речевые отрезки с паузой короче min_gap: заминка внутри слова не режет фразу."""
    merged: list = []
    for start, end in segments:
        if merged and start - merged[-1][1] < min_gap and end - merged[-1][0] <= max_length:
            merged[-1] = (merged[-1][0], end)
        else:
            merged.append((start, end))
    return merged


def _clean_transcript(text: str, provider: str) -> str:
    """Нормализовать пробелы и убрать только известную отдельную концовку Whisper."""
    cleaned = re.sub(r"[\t\r\n ]+", " ", text or "").strip()
    if provider != "Groq" or not cleaned:
        return cleaned

    for phrase in GROQ_HALLUCINATION_TAILS:
        pattern = rf"(?:^|(?<=[.!?…]))\s*{re.escape(phrase)}[.!?…]*$"
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE).rstrip()
    return cleaned


def _file_sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _is_autostart_enabled() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_READ,
        )
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False


def _set_autostart(enable: bool):
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0, winreg.KEY_SET_VALUE,
    )
    if enable:
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{EXE_PATH}" --delayed')
    else:
        try:
            winreg.DeleteValue(key, APP_NAME)
        except FileNotFoundError:
            pass
    winreg.CloseKey(key)


def _get_window_title(user32, hwnd) -> str:
    length = user32.GetWindowTextLengthW(hwnd) + 1
    buf = ctypes.create_unicode_buffer(length)
    user32.GetWindowTextW(hwnd, buf, length)
    return buf.value


# =============================================================================
# Win32 constants
# =============================================================================

KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_V = 0x56

# Clipboard Win32 constants
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002


# =============================================================================
# Application
# =============================================================================

class _LiveTranscription:
    """Черновик распознавания, который копится во время записи Push-to-Talk."""

    def __init__(self, frames: list):
        self.frames = frames             # тот же список, куда пишет поток микрофона
        self.texts: list[str] = []       # тексты закрытых отрезков по порядку записи
        self.consumed_frames = 0         # сколько блоков с начала записи уже распознано
        self.failed = False
        self.stop = threading.Event()
        self.thread = None

    def usable(self) -> bool:
        return not self.failed


class VoiceTyperApp:

    def __init__(self):
        self.is_recording = False
        self._key_held = False
        self._processing_lock = threading.Lock()
        self.audio_frames: list[bytes] = []
        self.stream = None
        self._target_hwnd = None
        self._cooldown_until = 0.0
        self._user32 = ctypes.windll.user32
        self._listener = None
        self._tray_icon = None
        self._vad_stop = threading.Event()
        self._cancel_retry = threading.Event()
        self._shutdown = threading.Event()
        self._vad_thread = None
        self._record_thread = None
        self._live = None

        # Локальная модель грузится один раз и остаётся в памяти между диктовками.
        self._local_model = None
        self._local_vad_model = None
        self._local_long_model = None
        self._local_state = "not_started"
        self._local_error = ""
        self._local_ready = threading.Event()
        self._local_load_thread = None
        self._local_state_lock = threading.Lock()

        # Провайдеры транскрипции — Gemini и Groq хранятся независимо
        self._provider = DEFAULT_PROVIDER
        self._gemini_client = None
        self._groq_client = None
        self._gemini_api_key = ""
        self._groq_api_key = ""
        self._gemini_model = DEFAULT_GEMINI_MODEL
        self._groq_model = DEFAULT_GROQ_MODEL

        # Hotkey state
        self._hotkey_key = keyboard.Key.f8
        self._hotkey_modifiers: set = set()
        self._pressed_modifiers: set = set()
        self._capture_modifiers: set = set()

        # UI / mode state
        self._ui_lang = "English"
        self._mode = "Push-to-Talk"
        self._status_key = "status_ready"
        self._status_fmt = {}

        # History
        self._history: list[dict] = []
        self._load_history()

        # PyAudio с retry при автозапуске
        self.audio = None
        for attempt in range(3):
            try:
                self.audio = pyaudio.PyAudio()
                break
            except Exception:
                if attempt < 2:
                    time.sleep(5)
        if self.audio is None:
            self.audio = pyaudio.PyAudio()

        # GUI
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry("480x520")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._minimize_to_tray)

        self._build_gui()
        self._build_overlay()
        self._load_settings()
        self._apply_mode()
        self._start_tray()
        if self._provider == LOCAL_PROVIDER:
            self._start_local_model_load()

        self.root.withdraw()

    # =====================================================================
    # System Tray
    # =====================================================================

    def _start_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Show / Hide", self._tray_toggle, default=True),
            pystray.MenuItem("Exit", self._tray_exit),
        )
        self._tray_icon = pystray.Icon(APP_NAME, _create_tray_image(), WINDOW_TITLE, menu)
        threading.Thread(target=self._tray_icon.run, daemon=True).start()

    def _tray_toggle(self, icon=None, item=None):
        self.root.after(0, self._toggle_window)

    def _tray_exit(self, icon=None, item=None):
        self.root.after(0, self._quit_app)

    def _toggle_window(self):
        if self.root.state() == "withdrawn":
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        else:
            self.root.withdraw()

    def _minimize_to_tray(self):
        self.root.withdraw()

    def _quit_app(self):
        self._shutdown.set()
        self._vad_stop.set()
        self.is_recording = False
        if self._live:
            self._live.stop.set()
        if self._listener:
            self._listener.stop()
        if self._vad_thread and self._vad_thread.is_alive():
            self._vad_thread.join(timeout=2)
        if self._record_thread and self._record_thread.is_alive():
            self._record_thread.join(timeout=2)
        if self.stream:
            try:
                self.stream.stop_stream()
                self.stream.close()
            except Exception:
                pass
        if self.audio:
            self.audio.terminate()
        if self._tray_icon:
            self._tray_icon.stop()
        self.root.destroy()

    # =====================================================================
    # Localization helpers
    # =====================================================================

    def _s(self, key: str, **fmt) -> str:
        """Получить строку текущего языка по ключу."""
        text = UI_STRINGS.get(self._ui_lang, UI_STRINGS["English"]).get(key, key)
        return text.format(**fmt) if fmt else text

    def _apply_ui_language(self):
        """Обновить все виджеты под текущий язык."""
        self._lbl_subtitle.configure(text=self._s("subtitle"))
        self._lbl_api_key.configure(text=self._s("api_key_lbl"))
        self._btn_paste.configure(text=self._s("paste_btn"))
        self._btn_toggle.configure(text=self._s("hide_btn" if self._key_visible else "show_btn"))
        self._lbl_settings.configure(text=self._s("settings_lbl"))
        self._lbl_mode.configure(text=self._s("mode_lbl"))
        self._lbl_provider.configure(text=self._s("provider_lbl"))
        self._lbl_model.configure(text=self._s("model_lbl"))
        self._lbl_hotkey.configure(text=self._s("hotkey_lbl"))
        self.hotkey_capture_btn.configure(text=self._s("hotkey_change"))
        self._lbl_language.configure(text=self._s("language_lbl"))
        self._chk_autostart.configure(text=self._s("autostart_chk"))
        self._btn_history.configure(text=self._s("history_btn"))
        self._btn_logs.configure(text=self._s("logs_btn"))
        if self._provider == LOCAL_PROVIDER:
            self.api_key_entry.configure(placeholder_text=self._s("local_key_hint"))
        self._update_hint()
        self._refresh_status()

    def _refresh_status(self):
        """Обновить статус-лейбл при смене языка."""
        text = self._s(self._status_key, **self._status_fmt)
        self.status_label.configure(text=text)

    # =====================================================================
    # GUI
    # =====================================================================

    def _build_gui(self):
        pad = dict(padx=20)

        ctk.CTkLabel(
            self.root, text="Voice Typer",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(pady=(16, 2), **pad)
        self._lbl_subtitle = ctk.CTkLabel(
            self.root, text=self._s("subtitle"),
            font=ctk.CTkFont(size=12), text_color="gray",
        )
        self._lbl_subtitle.pack(pady=(0, 12), **pad)

        # API Key
        key_frame = ctk.CTkFrame(self.root)
        key_frame.pack(fill="x", **pad, pady=(0, 6))
        self._lbl_api_key = ctk.CTkLabel(key_frame, text=self._s("api_key_lbl"))
        self._lbl_api_key.pack(side="left", padx=(12, 8), pady=8)
        self.api_key_var = ctk.StringVar()
        self.api_key_entry = ctk.CTkEntry(
            key_frame, textvariable=self.api_key_var, show="*",
            placeholder_text=self._s("local_key_hint"), state="disabled",
        )
        self.api_key_entry.pack(side="left", fill="x", expand=True, pady=8)
        self.api_key_entry.bind("<FocusOut>", self._on_api_key_changed)
        self.api_key_entry.bind("<Return>", self._on_api_key_changed)
        self._btn_paste = ctk.CTkButton(
            key_frame, text=self._s("paste_btn"), width=60,
            command=self._paste_api_key,
        )
        self._btn_paste.pack(side="left", padx=(4, 0), pady=8)
        self._key_visible = False
        self._btn_toggle = ctk.CTkButton(
            key_frame, text=self._s("show_btn"), width=60,
            command=self._toggle_key_vis,
        )
        self._btn_toggle.pack(side="left", padx=(4, 12), pady=8)

        # Settings panel
        sf = ctk.CTkFrame(self.root)
        sf.pack(fill="x", **pad, pady=(0, 4))
        self._lbl_settings = ctk.CTkLabel(
            sf, text=self._s("settings_lbl"),
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self._lbl_settings.pack(pady=(8, 4), padx=12, anchor="w")

        # Mode
        r0 = ctk.CTkFrame(sf, fg_color="transparent")
        r0.pack(fill="x", padx=12, pady=(0, 4))
        self._lbl_mode = ctk.CTkLabel(r0, text=self._s("mode_lbl"), width=100, anchor="w")
        self._lbl_mode.pack(side="left")
        self.mode_var = ctk.StringVar(value="Push-to-Talk")
        ctk.CTkOptionMenu(
            r0, variable=self.mode_var,
            values=["Push-to-Talk", "Voice Activated"], width=160,
            command=self._on_mode_changed,
        ).pack(side="left")

        # Provider (Gemini / Groq)
        r_provider = ctk.CTkFrame(sf, fg_color="transparent")
        r_provider.pack(fill="x", padx=12, pady=(0, 4))
        self._lbl_provider = ctk.CTkLabel(r_provider, text=self._s("provider_lbl"), width=100, anchor="w")
        self._lbl_provider.pack(side="left")
        self.provider_var = ctk.StringVar(value=DEFAULT_PROVIDER)
        ctk.CTkOptionMenu(
            r_provider, variable=self.provider_var,
            values=AVAILABLE_PROVIDERS, width=160,
            command=self._on_provider_changed,
        ).pack(side="left")

        # Model (список зависит от провайдера — меняется через _on_provider_changed)
        r_model = ctk.CTkFrame(sf, fg_color="transparent")
        r_model.pack(fill="x", padx=12, pady=(0, 4))
        self._lbl_model = ctk.CTkLabel(r_model, text=self._s("model_lbl"), width=100, anchor="w")
        self._lbl_model.pack(side="left")
        self.model_var = ctk.StringVar(value=LOCAL_MODEL_LABEL)
        self._model_menu = ctk.CTkOptionMenu(
            r_model, variable=self.model_var,
            values=[LOCAL_MODEL_LABEL], width=220,
            command=self._on_model_changed,
        )
        self._model_menu.pack(side="left")

        # Hotkey — capture-режим
        self.hotkey_row = ctk.CTkFrame(sf, fg_color="transparent")
        self.hotkey_row.pack(fill="x", padx=12, pady=(0, 4))
        self._lbl_hotkey = ctk.CTkLabel(
            self.hotkey_row, text=self._s("hotkey_lbl"), width=100, anchor="w",
        )
        self._lbl_hotkey.pack(side="left")
        self.hotkey_var = ctk.StringVar(value="F8")
        self.hotkey_display = ctk.CTkLabel(
            self.hotkey_row, textvariable=self.hotkey_var,
            font=ctk.CTkFont(size=13, weight="bold"),
            width=110, anchor="w",
        )
        self.hotkey_display.pack(side="left", padx=(0, 8))
        self.hotkey_capture_btn = ctk.CTkButton(
            self.hotkey_row, text=self._s("hotkey_change"), width=110,
            command=self._start_hotkey_capture,
        )
        self.hotkey_capture_btn.pack(side="left")

        # Language (UI language)
        r2 = ctk.CTkFrame(sf, fg_color="transparent")
        r2.pack(fill="x", padx=12, pady=(0, 4))
        self._lbl_language = ctk.CTkLabel(r2, text=self._s("language_lbl"), width=100, anchor="w")
        self._lbl_language.pack(side="left")
        self.language_var = ctk.StringVar(value="English")
        ctk.CTkOptionMenu(
            r2, variable=self.language_var,
            values=["English", "Russian", "Ukrainian"], width=120,
            command=self._on_language_changed,
        ).pack(side="left")

        # Autostart
        r3 = ctk.CTkFrame(sf, fg_color="transparent")
        r3.pack(fill="x", padx=12, pady=(0, 8))
        self.autostart_var = ctk.BooleanVar(value=False)
        self._chk_autostart = ctk.CTkCheckBox(
            r3, text=self._s("autostart_chk"), variable=self.autostart_var,
            command=self._on_autostart_changed,
        )
        self._chk_autostart.pack(side="left")

        # Hint
        self.hint_label = ctk.CTkLabel(
            self.root, text="", font=ctk.CTkFont(size=13), text_color="gray",
        )
        self.hint_label.pack(pady=(6, 4), **pad)

        # Status
        self.status_label = ctk.CTkLabel(
            self.root, text=self._s("status_ready"),
            font=ctk.CTkFont(size=16, weight="bold"), text_color="#4CAF50",
        )
        self.status_label.pack(pady=(2, 6))

        # History + Logs
        btn_row = ctk.CTkFrame(self.root, fg_color="transparent")
        btn_row.pack(fill="x", **pad, pady=(0, 12))
        self._btn_history = ctk.CTkButton(
            btn_row, text=self._s("history_btn"), width=200,
            command=self._show_history,
        )
        self._btn_history.pack(side="left", expand=True, padx=(0, 6))
        self._btn_logs = ctk.CTkButton(
            btn_row, text=self._s("logs_btn"), width=200,
            command=self._show_logs,
        )
        self._btn_logs.pack(side="left", expand=True, padx=(6, 0))

    # =====================================================================
    # Overlay (History / Logs)
    # =====================================================================

    def _build_overlay(self):
        self._overlay_frame = ctk.CTkFrame(self.root, corner_radius=0)
        header = ctk.CTkFrame(self._overlay_frame, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(8, 4))
        self._overlay_title = ctk.CTkLabel(
            header, text="", font=ctk.CTkFont(size=16, weight="bold"),
        )
        self._overlay_title.pack(side="left")
        ctk.CTkButton(
            header, text="✕", width=32, height=28,
            command=self._close_overlay,
        ).pack(side="right")

        # Обычный tkinter.Text — надёжное выделение и Ctrl+C
        text_frame = tk.Frame(self._overlay_frame, bg="#2b2b2b")
        text_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        self._overlay_text = tk.Text(
            text_frame,
            font=("Consolas", 12),
            wrap="word",
            bg="#2b2b2b",
            fg="#DCE4EE",
            insertbackground="#DCE4EE",
            selectbackground="#1f538d",
            selectforeground="white",
            yscrollcommand=scrollbar.set,
            relief="flat",
            borderwidth=0,
            padx=4,
            pady=4,
        )
        self._overlay_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self._overlay_text.yview)
        # Readonly: блокировать редактирование, разрешить выделение и копирование
        self._overlay_text.bind("<Key>", self._on_overlay_key)
        self._overlay_text.bind("<Control-c>", self._on_overlay_copy)
        self._overlay_text.bind("<Control-a>", self._on_overlay_select_all)
        # Правая кнопка мыши — контекстное меню
        self._overlay_text.bind("<Button-3>", self._on_overlay_rclick)
        self._overlay_menu = tk.Menu(self._overlay_text, tearoff=0)

    def _on_overlay_key(self, event):
        """Readonly: блокировать редактирование, разрешить навигацию и Ctrl+комбинации."""
        # Пропустить модификаторы
        if event.keysym in ('Control_L', 'Control_R', 'Shift_L', 'Shift_R',
                            'Alt_L', 'Alt_R', 'Caps_Lock'):
            return
        # Ctrl+клавиша — обрабатываем по keycode (работает на ЛЮБОЙ раскладке)
        if event.state & 0x4:  # Control modifier
            if event.keycode == 67:  # физическая клавиша C
                self._do_copy()
                return "break"
            if event.keycode == 65:  # физическая клавиша A
                self._overlay_text.tag_add("sel", "1.0", "end")
                return "break"
            return
        # Навигация
        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End',
                            'Prior', 'Next'):
            return
        return "break"

    def _on_overlay_copy(self, event):
        """Ctrl+C (латинская раскладка)."""
        self._do_copy()
        return "break"

    def _on_overlay_select_all(self, event):
        """Ctrl+A (латинская раскладка)."""
        self._overlay_text.tag_add("sel", "1.0", "end")
        return "break"

    def _do_copy(self):
        """Скопировать выделенный текст в clipboard."""
        try:
            selected = self._overlay_text.get("sel.first", "sel.last")
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
        except tk.TclError:
            pass

    def _on_overlay_rclick(self, event):
        """Правая кнопка мыши — контекстное меню."""
        menu = self._overlay_menu
        menu.delete(0, "end")
        has_sel = False
        try:
            self._overlay_text.get("sel.first", "sel.last")
            has_sel = True
        except tk.TclError:
            pass
        menu.add_command(
            label=self._s("copy_sel") if has_sel else self._s("copy_sel"),
            command=self._copy_selection,
            state="normal" if has_sel else "disabled",
        )
        menu.add_command(
            label=self._s("select_all"),
            command=lambda: self._overlay_text.tag_add("sel", "1.0", "end"),
        )
        menu.tk_popup(event.x_root, event.y_root)

    def _copy_selection(self):
        """Скопировать выделенный текст (для контекстного меню)."""
        self._do_copy()

    def _show_history(self):
        self._overlay_title.configure(text=self._s("history_btn"))
        self._overlay_text.config(state="normal")
        self._overlay_text.delete("1.0", "end")
        if not self._history:
            self._overlay_text.insert("end", self._s("no_history"))
        else:
            for entry in reversed(self._history):
                self._overlay_text.insert("end", f"[{entry['time']}]  {entry['text']}\n\n")
        self._overlay_frame.place(x=0, y=0, relwidth=1, relheight=1)
        self._overlay_frame.lift()
        self._overlay_text.focus_set()

    def _show_logs(self):
        self._overlay_title.configure(text=self._s("logs_btn"))
        self._overlay_text.config(state="normal")
        self._overlay_text.delete("1.0", "end")
        try:
            log_path = os.path.join(APP_DIR, "debug.log")
            with open(log_path, "r", encoding="utf-8") as f:
                content = f.read()
            self._overlay_text.insert("end", content if content else self._s("no_logs"))
        except Exception:
            self._overlay_text.insert("end", self._s("no_logs"))
        self._overlay_text.see("end")
        self._overlay_frame.place(x=0, y=0, relwidth=1, relheight=1)
        self._overlay_frame.lift()
        self._overlay_text.focus_set()

    def _close_overlay(self):
        self._overlay_frame.place_forget()

    # =====================================================================
    # GUI helpers
    # =====================================================================

    def _paste_api_key(self):
        try:
            self.api_key_var.set(self.root.clipboard_get().strip())
            self._on_api_key_changed(None)
        except Exception:
            pass

    def _toggle_key_vis(self):
        self._key_visible = not self._key_visible
        self.api_key_entry.configure(show="" if self._key_visible else "*")
        self._btn_toggle.configure(text=self._s("hide_btn" if self._key_visible else "show_btn"))

    def _on_model_changed(self, value):
        if self._provider == LOCAL_PROVIDER:
            return
        if self._provider == "Groq":
            self._groq_model = value
        else:
            self._gemini_model = value
        self._log("info", f"{self._provider} model changed to: {value}")
        self._auto_save()

    def _on_provider_changed(self, value):
        self._provider = value
        self._log("info", f"Provider changed to: {value}")
        self._apply_provider_ui()
        if value == LOCAL_PROVIDER:
            self._start_local_model_load()
        self._auto_save()

    def _apply_provider_ui(self):
        if self._provider == LOCAL_PROVIDER:
            self._model_menu.configure(values=[LOCAL_MODEL_LABEL])
            self.model_var.set(LOCAL_MODEL_LABEL)
            self.api_key_var.set("")
            self.api_key_entry.configure(
                state="disabled", placeholder_text=self._s("local_key_hint")
            )
            self._btn_paste.configure(state="disabled")
            self._btn_toggle.configure(state="disabled")
            if self._local_state == "ready":
                self._set_status("status_ready", "#4CAF50")
            elif self._local_state == "failed":
                self._set_status("status_model_error", "#F44336")
            else:
                self._set_status("status_model_loading", "#FF9800", tray_state="processing")
            return

        self.api_key_entry.configure(state="normal")
        self._btn_paste.configure(state="normal")
        self._btn_toggle.configure(state="normal")
        if self._provider == "Groq":
            self._model_menu.configure(values=GROQ_MODELS)
            self.model_var.set(self._groq_model)
            self.api_key_var.set(self._groq_api_key)
            self.api_key_entry.configure(placeholder_text="Paste your Groq API key (gsk_...)")
            active_client = self._groq_client
        else:
            self._model_menu.configure(values=GEMINI_MODELS)
            self.model_var.set(self._gemini_model)
            self.api_key_var.set(self._gemini_api_key)
            self.api_key_entry.configure(placeholder_text="Paste your Gemini API key (AIza...)")
            active_client = self._gemini_client
        if active_client:
            self._set_status("status_ready", "#4CAF50")
        else:
            self._set_status("status_api_key", "#F44336")

    def _on_mode_changed(self, value):
        if value == "Voice Activated":
            self.hotkey_row.pack_forget()
        else:
            self.hotkey_row.pack(
                fill="x", padx=12, pady=(0, 4),
                after=self.hotkey_row.master.winfo_children()[1],
            )
        self._mode = value
        self._apply_mode()
        self._update_hint()
        self._auto_save()

    def _on_language_changed(self, value):
        self._ui_lang = value
        self._apply_ui_language()
        self._auto_save()

    def _on_autostart_changed(self):
        self._auto_save()

    def _on_api_key_changed(self, event):
        if self._provider == LOCAL_PROVIDER:
            return
        api_key = self.api_key_var.get().strip()
        if api_key:
            if self._provider == "Groq":
                self._groq_api_key = api_key
                self._configure_groq(api_key)
            else:
                self._gemini_api_key = api_key
                self._configure_gemini(api_key)
            self._set_status("status_ready", "#4CAF50")
        self._auto_save()

    def _auto_save(self):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "settings_version": SETTINGS_VERSION,
                    "provider": self.provider_var.get(),
                    "gemini_api_key": self._gemini_api_key,
                    "groq_api_key": self._groq_api_key,
                    "gemini_model": self._gemini_model,
                    "groq_model": self._groq_model,
                    "hotkey": self.hotkey_var.get(),
                    "language": self.language_var.get(),
                    "mode": self.mode_var.get(),
                }, f, indent=2)
        except IOError:
            pass
        try:
            _set_autostart(self.autostart_var.get())
        except Exception:
            pass

    def _update_hint(self):
        if self._mode == "Voice Activated":
            self.hint_label.configure(text=self._s("hint_vad"))
        else:
            name = self.hotkey_var.get()
            self.hint_label.configure(text=self._s("hint_ptt", name=name))

    def _set_status(self, key: str, color: str = "gray", tray_state: str = None, **fmt):
        """Установить статус по ключу локализации."""
        self._status_key = key
        self._status_fmt = dict(fmt)
        text = self._s(key, **fmt)
        self.status_label.configure(text=text, text_color=color)
        # Определяем состояние трея
        if tray_state:
            state = tray_state
        elif "recording" in key:
            state = "recording"
        elif "processing" in key or "pasted" in key or "ready_in" in key or "wait" in key:
            state = "processing"
        else:
            state = "idle"
        self._set_tray_state(state)

    def _set_tray_state(self, state: str):
        if self._tray_icon:
            try:
                self._tray_icon.icon = _create_tray_image(state)
            except Exception:
                pass

    def _ui(self, fn):
        self.root.after(0, fn)

    def _log(self, level: str, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        try:
            log_path = os.path.join(APP_DIR, "debug.log")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] [{level.upper()}] {msg}\n")
        except Exception:
            pass

    # =====================================================================
    # Hotkey capture
    # =====================================================================

    def _start_hotkey_capture(self):
        self.hotkey_var.set(self._s("capture_prompt"))
        self.hotkey_capture_btn.configure(state="disabled")
        self._capture_modifiers = set()

        def on_press(key):
            if key == keyboard.Key.esc:
                self.root.after(0, self._cancel_hotkey_capture)
                return False
            if _is_modifier(key):
                self._capture_modifiers.add(_modifier_name(key))
                return
            key_name = _key_to_name(key)
            mods = sorted(self._capture_modifiers)
            hotkey_str = "+".join(mods + [key_name]) if mods else key_name
            self.root.after(0, lambda k=key, m=mods, s=hotkey_str: self._apply_captured_hotkey(k, m, s))
            return False

        capture_listener = keyboard.Listener(on_press=on_press)
        capture_listener.daemon = True
        capture_listener.start()

    def _cancel_hotkey_capture(self):
        current = self.hotkey_var.get()
        if current == self._s("capture_prompt"):
            self.hotkey_var.set("F8")
        self.hotkey_capture_btn.configure(state="normal")

    def _apply_captured_hotkey(self, key, mods: list, hotkey_str: str):
        self._hotkey_key = key
        self._hotkey_modifiers = set(mods)
        self.hotkey_var.set(hotkey_str)
        self.hotkey_capture_btn.configure(state="normal")
        self._pressed_modifiers.clear()
        self._apply_mode()
        self._auto_save()
        self._update_hint()
        self._log("info", f"Hotkey changed to: {hotkey_str}")

    # =====================================================================
    # History persistence
    # =====================================================================

    def _load_history(self):
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    self._history = json.load(f)
            except (json.JSONDecodeError, IOError):
                self._history = []

    def _save_history_entry(self, text: str):
        entry = {"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "text": text}
        self._history.append(entry)
        if len(self._history) > 500:
            self._history = self._history[-500:]
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self._history, f, ensure_ascii=False, indent=1)
        except IOError:
            pass

    # =====================================================================
    # Settings persistence
    # =====================================================================

    def _load_settings(self):
        data = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass

        # Миграция старого формата: api_key → gemini_api_key, model → gemini_model
        if "api_key" in data and "gemini_api_key" not in data:
            data["gemini_api_key"] = data.pop("api_key")
        if "model" in data and "gemini_model" not in data:
            data["gemini_model"] = data.pop("model")

        self._gemini_api_key = data.get("gemini_api_key", "")
        self._groq_api_key = data.get("groq_api_key", "")

        # Миграция старых имён моделей Gemini
        gemini_model = data.get("gemini_model", DEFAULT_GEMINI_MODEL)
        if gemini_model == "gemini-3-flash":
            gemini_model = "gemini-3-flash-preview"
        if gemini_model in GEMINI_MODELS:
            self._gemini_model = gemini_model

        groq_model = data.get("groq_model", DEFAULT_GROQ_MODEL)
        if groq_model in GROQ_MODELS:
            self._groq_model = groq_model

        try:
            settings_version = int(data.get("settings_version", 0))
        except (TypeError, ValueError):
            settings_version = 0

        provider = data.get("provider", DEFAULT_PROVIDER)
        if provider not in AVAILABLE_PROVIDERS:
            provider = DEFAULT_PROVIDER
        # Однократный переход старых установок на бесплатный локальный движок.
        if settings_version < SETTINGS_VERSION:
            provider = LOCAL_PROVIDER
        self._provider = provider
        self.provider_var.set(provider)
        self._apply_provider_ui()

        hotkey_str = data.get("hotkey", "F8")
        self._hotkey_key, self._hotkey_modifiers = _parse_hotkey_str(hotkey_str)
        self.hotkey_var.set(hotkey_str)

        # Language = UI language; old "Auto" → "English"
        lang = data.get("language", "English")
        if lang not in UI_STRINGS:
            lang = "English"
        self.language_var.set(lang)
        self._ui_lang = lang

        mode = data.get("mode", "Push-to-Talk")
        if mode in ("Push-to-Talk", "Voice Activated"):
            self.mode_var.set(mode)
        self._mode = self.mode_var.get()

        self.autostart_var.set(_is_autostart_enabled())
        self._on_mode_changed(self._mode)

        # Конфигурируем оба клиента если ключи есть — пользователь может переключаться между ними
        if self._gemini_api_key:
            self._configure_gemini(self._gemini_api_key)
        if self._groq_api_key:
            self._configure_groq(self._groq_api_key)

        active_client = self._groq_client if provider == "Groq" else self._gemini_client
        if provider != LOCAL_PROVIDER and active_client:
            self._set_status("status_ready", "#4CAF50")

        self._apply_ui_language()
        self._update_hint()

    def _configure_gemini(self, api_key: str):
        try:
            self._gemini_client = genai.Client(api_key=api_key)
        except Exception as e:
            self._log("error", f"Gemini configure failed: {e}")
            self._gemini_client = None

    def _configure_groq(self, api_key: str):
        if not GROQ_AVAILABLE:
            self._log("error", "groq package not installed — run: pip install groq")
            self._groq_client = None
            return
        try:
            self._groq_client = Groq(api_key=api_key)
        except Exception as e:
            self._log("error", f"Groq configure failed: {e}")
            self._groq_client = None

    # =====================================================================
    # Local GigaAM model
    # =====================================================================

    def _start_local_model_load(self):
        with self._local_state_lock:
            if self._local_state in ("loading", "ready"):
                return
            self._local_state = "loading"
            self._local_error = ""
            self._local_ready.clear()

        if self._provider == LOCAL_PROVIDER:
            self._set_status("status_model_loading", "#FF9800", tray_state="processing")
        self._local_load_thread = threading.Thread(
            target=self._load_local_model,
            name="VoiceTyperLocalModel",
            daemon=True,
        )
        self._local_load_thread.start()

    def _ensure_local_model_files(self, directory, repo, revision, files):
        os.makedirs(directory, exist_ok=True)
        for filename, expected_hash in files.items():
            target = os.path.join(directory, filename)
            if os.path.isfile(target) and _file_sha256(target) == expected_hash:
                continue
            if self._shutdown.is_set():
                raise RuntimeError("application is shutting down")

            temp_path = target + ".download"
            url = f"https://huggingface.co/{repo}/resolve/{revision}/{filename}?download=true"
            self._log("info", f"Downloading local model file: {filename}")
            try:
                request = urllib.request.Request(url, headers={"User-Agent": "VoiceTyper/2"})
                with urllib.request.urlopen(request, timeout=60) as response, open(temp_path, "wb") as output:
                    while True:
                        chunk = response.read(1024 * 1024)
                        if not chunk:
                            break
                        if self._shutdown.is_set():
                            raise RuntimeError("application is shutting down")
                        output.write(chunk)

                actual_hash = _file_sha256(temp_path)
                if actual_hash != expected_hash:
                    raise RuntimeError(f"checksum mismatch for {filename}")
                os.replace(temp_path, target)
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass

    def _load_local_model(self):
        if not LOCAL_ASR_AVAILABLE:
            self._finish_local_model_error("onnx-asr or numpy package is not installed")
            return

        started = time.perf_counter()
        try:
            self._ensure_local_model_files(
                LOCAL_MODEL_DIR,
                LOCAL_MODEL_REPO,
                LOCAL_MODEL_REVISION,
                LOCAL_MODEL_FILES,
            )
            session_options = ort.SessionOptions()
            session_options.intra_op_num_threads = LOCAL_ASR_THREADS
            session_options.add_session_config_entry("session.intra_op.allow_spinning", "0")
            model = onnx_asr.load_model(
                LOCAL_MODEL_NAME,
                LOCAL_MODEL_DIR,
                quantization=LOCAL_MODEL_QUANTIZATION,
                providers=["CPUExecutionProvider"],
                sess_options=session_options,
            )

            long_model = None
            vad_model = None
            try:
                self._ensure_local_model_files(
                    LOCAL_VAD_DIR,
                    LOCAL_VAD_REPO,
                    LOCAL_VAD_REVISION,
                    LOCAL_VAD_FILES,
                )
                vad_model = onnx_asr.load_vad(
                    "silero",
                    LOCAL_VAD_DIR,
                    providers=["CPUExecutionProvider"],
                    sess_options=session_options,
                )
                long_model = model.with_vad(vad_model, **LOCAL_VAD_OPTIONS)
            except Exception as e:
                self._log("warn", f"Local VAD unavailable: {str(e)[:300]}")

            if self._shutdown.is_set():
                return
            with self._local_state_lock:
                self._local_model = model
                self._local_vad_model = vad_model
                self._local_long_model = long_model
                self._local_state = "ready"
                self._local_error = ""
                self._local_ready.set()

            elapsed = time.perf_counter() - started
            self._log("info", f"Local GigaAM ready in {elapsed:.2f}s")
            if not self._shutdown.is_set():
                self._ui(lambda: self._set_status("status_ready", "#4CAF50")
                         if self._provider == LOCAL_PROVIDER else None)
        except Exception as e:
            self._finish_local_model_error(str(e))

    def _finish_local_model_error(self, error: str):
        with self._local_state_lock:
            self._local_state = "failed"
            self._local_error = error[:500]
            self._local_ready.set()
        self._log("error", f"Local GigaAM load failed: {error[:500]}")
        if not self._shutdown.is_set():
            self._ui(lambda: self._set_status("status_model_error", "#F44336")
                     if self._provider == LOCAL_PROVIDER else None)

    # =====================================================================
    # Mode switching
    # =====================================================================

    def _apply_mode(self):
        self._vad_stop.set()
        if self._vad_thread and self._vad_thread.is_alive():
            self._vad_thread.join(timeout=3)

        if self._listener:
            self._listener.stop()
            self._listener = None

        self._pressed_modifiers.clear()

        if self._mode == "Voice Activated":
            self._vad_stop.clear()
            self._start_vad()
        else:
            self._start_hotkey_listener()

        self._log("info", f"Mode: {self._mode}")

    # =====================================================================
    # Push-to-Talk
    # =====================================================================

    def _start_hotkey_listener(self):
        self._listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release,
        )
        self._listener.daemon = True
        self._listener.start()

    def _on_key_press(self, key):
        if _is_modifier(key):
            self._pressed_modifiers.add(_modifier_name(key))
            return

        if key != self._hotkey_key:
            return
        if self._pressed_modifiers != self._hotkey_modifiers:
            return

        # Если cooldown — не начинать запись, показать сколько ждать
        if time.time() < self._cooldown_until:
            remaining = int(self._cooldown_until - time.time()) + 1
            self._ui(lambda r=remaining: self._set_status(
                "status_ready_in", "#FF9800", tray_state="processing", n=r
            ))
            return

        if self._key_held:
            return
        # Если идёт retry — отменить его, lock освободится сам
        if not self._processing_lock.acquire(blocking=False):
            self._cancel_retry.set()
            return
        self._processing_lock.release()

        self._target_hwnd = self._user32.GetForegroundWindow()
        self._log("info", f"PTT start, target: {_get_window_title(self._user32, self._target_hwnd)}")
        self._key_held = True
        self.is_recording = True
        self.audio_frames = []
        self._record_thread = threading.Thread(
            target=self._record_loop,
            name="VoiceTyperRecording",
            daemon=True,
        )
        self._record_thread.start()
        self._live = self._start_live_transcription(self.audio_frames)
        self._ui(lambda: self._set_status("status_recording", "#F44336"))

    def _on_key_release(self, key):
        if _is_modifier(key):
            self._pressed_modifiers.discard(_modifier_name(key))
            return
        if key != self._hotkey_key or not self._key_held:
            return
        self._key_held = False
        self.is_recording = False
        live, self._live = self._live, None
        self._ui(lambda: self._set_status("status_processing", "#FF9800"))
        threading.Thread(
            target=lambda: self._process_and_paste(wait_for_recording=True, live=live),
            name="VoiceTyperProcessing",
            daemon=True,
        ).start()

    # =====================================================================
    # Voice Activated (VAD)
    # =====================================================================

    def _start_vad(self):
        self._vad_thread = threading.Thread(target=self._vad_loop, daemon=True)
        self._vad_thread.start()
        self._log("info", "VAD started")

    def _vad_loop(self):
        try:
            stream = self.audio.open(
                format=AUDIO_FORMAT, channels=CHANNELS,
                rate=SAMPLE_RATE, input=True,
                frames_per_buffer=CHUNK_SIZE,
            )
        except Exception as e:
            self._log("error", f"VAD mic: {e}")
            self._ui(lambda: self._set_status("status_api_error", "#F44336"))
            return

        recording = False
        frames: list[bytes] = []
        silence_start = 0.0
        speech_start = 0.0
        target_hwnd = None

        try:
            while not self._vad_stop.is_set():
                try:
                    data = stream.read(CHUNK_SIZE, exception_on_overflow=False)
                except Exception:
                    continue

                level = _rms(data)

                if not recording:
                    if level > VAD_SPEECH_THRESHOLD:
                        fg = self._user32.GetForegroundWindow()
                        title = _get_window_title(self._user32, fg)
                        if title == WINDOW_TITLE:
                            continue
                        recording = True
                        frames = [data]
                        speech_start = time.time()
                        target_hwnd = fg
                        silence_start = 0.0
                        self._log("info", f"VAD speech detected, target: {title}")
                        self._ui(lambda: self._set_status("status_recording", "#F44336"))
                else:
                    frames.append(data)
                    if level < VAD_SPEECH_THRESHOLD:
                        if silence_start == 0.0:
                            silence_start = time.time()
                        elif time.time() - silence_start >= VAD_SILENCE_TIMEOUT:
                            recording = False
                            duration = time.time() - speech_start
                            if duration >= VAD_MIN_SPEECH_DURATION:
                                # Проверить cooldown перед обработкой
                                if time.time() < self._cooldown_until:
                                    remaining = int(self._cooldown_until - time.time()) + 1
                                    self._ui(lambda r=remaining: self._set_status(
                                        "status_ready_in", "#FF9800", tray_state="processing", n=r
                                    ))
                                    self._log("info", f"VAD: cooldown active, {remaining}s left — skipping")
                                else:
                                    f_copy = frames[:]
                                    h_copy = target_hwnd
                                    def _process(frames_=f_copy, hwnd_=h_copy):
                                        self._process_and_paste(frames=frames_, hwnd=hwnd_)
                                    self._ui(lambda: self._set_status("status_processing", "#FF9800"))
                                    threading.Thread(target=_process, daemon=True).start()
                            else:
                                self._ui(lambda: self._set_status("status_ready", "#4CAF50"))
                            frames = []
                            silence_start = 0.0
                    else:
                        silence_start = 0.0
        finally:
            stream.stop_stream()
            stream.close()

    # =====================================================================
    # Audio recording (Push-to-Talk only)
    # =====================================================================

    def _record_loop(self):
        try:
            self.stream = self.audio.open(
                format=AUDIO_FORMAT, channels=CHANNELS,
                rate=SAMPLE_RATE, input=True,
                frames_per_buffer=CHUNK_SIZE,
            )
            while self.is_recording:
                data = self.stream.read(CHUNK_SIZE, exception_on_overflow=False)
                self.audio_frames.append(data)
        except Exception as e:
            self._ui(lambda: self._set_status("status_api_error", "#F44336"))
            self._log("error", f"Record: {e}")
        finally:
            if self.stream:
                try:
                    self.stream.stop_stream()
                    self.stream.close()
                except Exception:
                    pass
                self.stream = None

    # =====================================================================
    # Live transcription during Push-to-Talk recording
    # =====================================================================

    def _start_live_transcription(self, frames: list):
        """Запустить распознавание закрытых отрезков, пока клавиша ещё зажата. None = обычный путь."""
        if self._provider != LOCAL_PROVIDER or np is None:
            return None
        if self._local_state != "ready" or self._local_model is None or self._local_vad_model is None:
            return None
        live = _LiveTranscription(frames)
        live.thread = threading.Thread(
            target=self._live_transcribe_loop,
            args=(live,),
            name="VoiceTyperLive",
            daemon=True,
        )
        live.thread.start()
        return live

    def _live_transcribe_loop(self, live: "_LiveTranscription"):
        try:
            while not live.stop.wait(LIVE_POLL_SECONDS):
                self._live_transcribe_step(live)
        except Exception as e:
            live.failed = True
            self._log("error", f"Live transcription failed: {str(e)[:300]}")

    def _local_speech_segments(self, waveform) -> list:
        """Границы речевых отрезков (в отсчётах) по Silero VAD с теми же настройками, что у длинного пути."""
        segments = self._local_vad_model.segment_batch(
            waveform[None, :],
            np.array([waveform.size], dtype=np.int64),
            SAMPLE_RATE,
            **LOCAL_VAD_OPTIONS,
        )
        return [(int(start), int(end)) for start, end in next(iter(segments))]

    def _live_transcribe_step(self, live: "_LiveTranscription"):
        """Распознать отрезки, после которых уже записана тишина; открытый хвост остаётся на потом."""
        base = live.consumed_frames
        pending = live.frames[base:len(live.frames)]
        if not pending:
            return
        if sum(len(f) for f in pending) < int(LIVE_MIN_PENDING_SECONDS * SAMPLE_RATE) * 2:
            return

        waveform = np.frombuffer(b"".join(pending), dtype="<i2").astype(np.float32) / 32768.0
        segments = _merge_close_segments(
            self._local_speech_segments(waveform),
            int(LIVE_MIN_GAP_SECONDS * SAMPLE_RATE),
            int(LOCAL_VAD_OPTIONS["max_speech_duration_s"] * SAMPLE_RATE),
        )
        guard = int(LIVE_TAIL_GUARD_SECONDS * SAMPLE_RATE)
        closed = [(s, e) for s, e in segments if e <= waveform.size - guard]
        if not closed:
            return

        frame_ends = np.cumsum([len(f) for f in pending])  # байтовые границы блоков микрофона
        started = time.perf_counter()
        done = 0
        for start, end in closed:
            if live.stop.is_set():
                break
            text = str(self._local_model.recognize(waveform[start:end], sample_rate=SAMPLE_RATE)).strip()
            if text:
                live.texts.append(text)
            live.consumed_frames = base + int(np.searchsorted(frame_ends, end * 2, side="right"))
            done += 1
        if done:
            self._log(
                "info",
                f"Live: +{done} segments asr={time.perf_counter() - started:.2f}s "
                f"consumed={live.consumed_frames * CHUNK_SIZE / SAMPLE_RATE:.1f}s texts={len(live.texts)}",
            )

    # =====================================================================
    # Transcription & paste
    # =====================================================================

    def _process_and_paste(self, frames=None, hwnd=None, wait_for_recording=False, live=None):
        if not self._processing_lock.acquire(blocking=False):
            self._log("warn", "Skipped: already processing")
            return
        try:
            if wait_for_recording:
                record_thread = self._record_thread
                if record_thread and record_thread is not threading.current_thread():
                    record_thread.join(timeout=2)
                    if record_thread.is_alive():
                        self._log("warn", "Recording thread did not stop before transcription")
            if live is not None:
                live.stop.set()
                if live.thread is not None and live.thread.is_alive():
                    live.thread.join(timeout=LIVE_JOIN_TIMEOUT_SECONDS)
                    if live.thread.is_alive():
                        live.failed = True
                        self._log("warn", "Live transcription did not stop, reprocessing whole recording")
            frames_to_process = list(self.audio_frames if frames is None else frames)
            target_hwnd = self._target_hwnd if hwnd is None else hwnd
            self._do_transcribe(frames_to_process, target_hwnd, live)
        finally:
            self._processing_lock.release()

    def _do_transcribe(self, frames, target_hwnd, live=None):
        if not frames:
            self._ui(lambda: self._set_status("status_no_audio", "#FF9800"))
            return

        provider = self._provider
        active_client = self._groq_client if provider == "Groq" else self._gemini_client
        if provider != LOCAL_PROVIDER and not active_client:
            self._ui(lambda: self._set_status("status_api_key", "#F44336"))
            return

        # Всё, что живое распознавание уже закрыло во время записи, повторно не распознаётся.
        prefix_text = ""
        tail_frames = frames
        if live is not None and live.usable() and live.consumed_frames > 0:
            prefix_text = " ".join(live.texts)
            tail_frames = frames[live.consumed_frames:]
            self._log(
                "info",
                f"Live: {len(live.texts)} texts ready before release, "
                f"tail={len(tail_frames) * CHUNK_SIZE / SAMPLE_RATE:.1f}s of {len(frames) * CHUNK_SIZE / SAMPLE_RATE:.1f}s",
            )

        pcm_bytes = _trim_pcm_silence(b"".join(tail_frames))
        if not pcm_bytes:
            if prefix_text:
                self._finish_transcription(prefix_text, LOCAL_PROVIDER, target_hwnd)
            else:
                self._ui(lambda: self._set_status("status_no_audio", "#FF9800"))
            return

        buf = io.BytesIO()
        with wave.open(buf, "wb") as wf:
            wf.setnchannels(CHANNELS)
            wf.setsampwidth(self.audio.get_sample_size(AUDIO_FORMAT))
            wf.setframerate(SAMPLE_RATE)
            wf.writeframes(pcm_bytes)
        audio_bytes = buf.getvalue()

        actual_provider = provider
        if provider == LOCAL_PROVIDER:
            text = self._call_local(pcm_bytes)
            if text is None:
                actual_provider, text = self._call_online_fallback(audio_bytes)
                if text is None:
                    if self._local_state == "loading":
                        self._ui(lambda: self._set_status(
                            "status_model_loading", "#FF9800", tray_state="processing"
                        ))
                    else:
                        self._ui(lambda: self._set_status("status_model_error", "#F44336"))
                    return
        elif provider == "Groq":
            text = self._call_groq(audio_bytes)
        else:
            text = self._call_gemini(audio_bytes)

        if text is None:
            if not prefix_text:
                return  # Статус уже установлен внутри _call_* (cancelled / api_error)
            self._log("warn", "Tail transcription failed, pasting the live part only")
            text = ""

        text = _clean_transcript(text, actual_provider)
        if prefix_text:
            text = f"{prefix_text} {text}".strip()
        self._finish_transcription(text, actual_provider, target_hwnd)

    def _finish_transcription(self, text: str, actual_provider: str, target_hwnd):
        if not text:
            self._ui(lambda: self._set_status("status_empty", "#FF9800"))
            return

        self._save_history_entry(text)
        self._log("info", f"Transcribed ({actual_provider}) {len(text)} chars: '{text[:60]}'")

        self.root.after(0, lambda: self._paste_pipeline(text, target_hwnd))

    def _call_local(self, pcm_bytes: bytes):
        """Локальное распознавание. None означает, что нужен запасной провайдер."""
        if self._local_state == "not_started":
            self._ui(self._start_local_model_load)
        if self._local_state != "ready":
            self._local_ready.wait(timeout=LOCAL_MODEL_WAIT_SECONDS)
        if self._local_state != "ready" or self._local_model is None:
            self._log("warn", f"Local GigaAM unavailable, state={self._local_state}")
            return None

        try:
            waveform = np.frombuffer(pcm_bytes, dtype="<i2").astype(np.float32) / 32768.0
            duration = waveform.size / SAMPLE_RATE
            started = time.perf_counter()
            if duration <= LOCAL_DIRECT_MAX_SECONDS:
                text = self._local_model.recognize(waveform, sample_rate=SAMPLE_RATE)
                segment_count = 1
            else:
                if self._local_long_model is None:
                    self._log("warn", "Local VAD unavailable for long transcription")
                    return None
                segments = list(self._local_long_model.recognize(waveform, sample_rate=SAMPLE_RATE))
                text = " ".join(segment.text.strip() for segment in segments if segment.text.strip())
                segment_count = len(segments)
            elapsed = time.perf_counter() - started
            self._log(
                "info",
                f"Local GigaAM: audio={duration:.2f}s asr={elapsed:.2f}s segments={segment_count}",
            )
            return str(text).strip()
        except Exception as e:
            self._log("error", f"Local GigaAM transcription failed: {str(e)[:500]}")
            return None

    def _call_online_fallback(self, audio_bytes: bytes):
        if self._groq_client:
            self._ui(lambda: self._set_status(
                "status_fallback", "#FF9800", tray_state="processing", provider="Groq"
            ))
            self._log("warn", "Local GigaAM fallback to Groq")
            return "Groq", self._call_groq(audio_bytes, language="ru")
        if self._gemini_client:
            self._ui(lambda: self._set_status(
                "status_fallback", "#FF9800", tray_state="processing", provider="Gemini"
            ))
            self._log("warn", "Local GigaAM fallback to Gemini")
            return "Gemini", self._call_gemini(audio_bytes)
        return LOCAL_PROVIDER, None

    def _call_gemini(self, audio_bytes: bytes):
        """Транскрипция через Gemini с fallback по моделям. None = все упали."""
        MAX_ROUNDS = 2
        ROUND_DELAY = 10
        self._cancel_retry.clear()

        models_to_try = list(GEMINI_MODELS)
        current_idx = 0
        for i, m in enumerate(models_to_try):
            if m == self._gemini_model:
                current_idx = i
                break
        models_to_try = models_to_try[current_idx:] + models_to_try[:current_idx]

        for round_num in range(MAX_ROUNDS):
            if round_num > 0:
                self._log("info", f"Gemini: all models exhausted, waiting {ROUND_DELAY}s before round {round_num+1}")
                for remaining in range(ROUND_DELAY, 0, -1):
                    if self._cancel_retry.is_set():
                        self._log("info", "Retry cancelled by user")
                        self._ui(lambda: self._set_status(
                            "status_cancelled", "#FF9800", tray_state="idle"))
                        return None
                    n = remaining
                    self._ui(lambda n=n: self._set_status(
                        "status_retry", "#FF9800", tray_state="processing",
                        attempt=round_num+1, n=n, max=MAX_ROUNDS,
                    ))
                    time.sleep(1)

            for model_name in models_to_try:
                if self._cancel_retry.is_set():
                    self._log("info", "Retry cancelled by user")
                    self._ui(lambda: self._set_status(
                        "status_cancelled", "#FF9800", tray_state="idle"))
                    return None
                try:
                    self._log("info", f"Trying Gemini model: {model_name}")
                    response = self._gemini_client.models.generate_content(
                        model=model_name,
                        contents=[
                            genai_types.Part.from_bytes(data=audio_bytes, mime_type="audio/wav"),
                            "Transcribe this audio.",
                        ],
                        config=genai_types.GenerateContentConfig(
                            system_instruction=BASE_SYSTEM_INSTRUCTION,
                        ),
                    )
                    text = response.text.strip()
                    self._cooldown_until = time.time() + MIN_API_INTERVAL
                    if model_name != self._gemini_model:
                        self._log("info", f"Gemini fallback succeeded on {model_name}")
                    return text
                except Exception as e:
                    err = str(e)[:200]
                    self._log("error", f"Gemini API error ({model_name}): {err}")
                    continue

        self._log("error", "All Gemini models exhausted after all rounds")
        self._ui(lambda: self._set_status("status_api_error", "#F44336"))
        self.root.after(2500, lambda: self._ui(
            lambda: self._set_status("status_ready", "#4CAF50", tray_state="idle")))
        return None

    def _call_groq(self, audio_bytes: bytes, language: str = None):
        """Транскрипция через Groq Whisper с fallback по моделям. None = все упали."""
        MAX_ROUNDS = 2
        ROUND_DELAY = 10
        self._cancel_retry.clear()

        models_to_try = list(GROQ_MODELS)
        if self._groq_model in models_to_try:
            idx = models_to_try.index(self._groq_model)
            models_to_try = models_to_try[idx:] + models_to_try[:idx]

        for round_num in range(MAX_ROUNDS):
            if round_num > 0:
                self._log("info", f"Groq: all models exhausted, waiting {ROUND_DELAY}s before round {round_num+1}")
                for remaining in range(ROUND_DELAY, 0, -1):
                    if self._cancel_retry.is_set():
                        self._log("info", "Retry cancelled by user")
                        self._ui(lambda: self._set_status(
                            "status_cancelled", "#FF9800", tray_state="idle"))
                        return None
                    n = remaining
                    self._ui(lambda n=n: self._set_status(
                        "status_retry", "#FF9800", tray_state="processing",
                        attempt=round_num+1, n=n, max=MAX_ROUNDS,
                    ))
                    time.sleep(1)

            for model_name in models_to_try:
                if self._cancel_retry.is_set():
                    self._log("info", "Retry cancelled by user")
                    self._ui(lambda: self._set_status(
                        "status_cancelled", "#FF9800", tray_state="idle"))
                    return None
                try:
                    self._log("info", f"Trying Groq model: {model_name}")
                    request = {
                        "file": ("audio.wav", audio_bytes),
                        "model": model_name,
                        "response_format": "text",
                        "temperature": 0,
                    }
                    if language:
                        request["language"] = language
                    response = self._groq_client.audio.transcriptions.create(
                        **request,
                    )
                    # Groq возвращает str для response_format="text", иначе object с .text
                    if isinstance(response, str):
                        text = response.strip()
                    else:
                        text = response.text.strip()
                    self._cooldown_until = time.time() + MIN_API_INTERVAL
                    if model_name != self._groq_model:
                        self._log("info", f"Groq fallback succeeded on {model_name}")
                    return text
                except Exception as e:
                    err = str(e)[:200]
                    self._log("error", f"Groq API error ({model_name}): {err}")
                    continue

        self._log("error", "All Groq models exhausted after all rounds")
        self._ui(lambda: self._set_status("status_api_error", "#F44336"))
        self.root.after(2500, lambda: self._ui(
            lambda: self._set_status("status_ready", "#4CAF50", tray_state="idle")))
        return None

    def _get_clipboard_text(self) -> str | None:
        """Прочитать текст из clipboard через Win32 API (main thread)."""
        kernel32 = ctypes.windll.kernel32
        user32 = self._user32
        # Retry: clipboard может быть заблокирован другим процессом
        for retry in range(5):
            try:
                if not user32.OpenClipboard(0):
                    time.sleep(0.05)
                    continue
                h = user32.GetClipboardData(CF_UNICODETEXT)
                if not h:
                    user32.CloseClipboard()
                    return ""  # Clipboard открыт но пуст — возвращаем пустую строку
                p = kernel32.GlobalLock(h)
                if not p:
                    user32.CloseClipboard()
                    return ""
                text = ctypes.wstring_at(p)
                kernel32.GlobalUnlock(h)
                user32.CloseClipboard()
                return text
            except Exception:
                try:
                    user32.CloseClipboard()
                except Exception:
                    pass
                time.sleep(0.05)
        # Win32 не смог — fallback на pyperclip
        try:
            return pyperclip.paste()
        except Exception:
            return None

    def _restore_clipboard(self, text: str):
        """Восстановить оригинальный текст в clipboard через Win32 API."""
        kernel32 = ctypes.windll.kernel32
        user32 = self._user32
        try:
            if not user32.OpenClipboard(0):
                self._log("warn", "Restore clipboard: OpenClipboard failed")
                return
            user32.EmptyClipboard()
            encoded = text.encode("utf-16-le") + b"\x00\x00"
            h = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(encoded))
            if not h:
                user32.CloseClipboard()
                return
            p = kernel32.GlobalLock(h)
            if not p:
                kernel32.GlobalFree(h)
                user32.CloseClipboard()
                return
            ctypes.memmove(p, encoded, len(encoded))
            kernel32.GlobalUnlock(h)
            user32.SetClipboardData(CF_UNICODETEXT, h)
            user32.CloseClipboard()
            self._log("info", f"Clipboard restored: {len(text)} chars")
        except Exception as e:
            self._log("error", f"Restore clipboard failed: {e}")
            try:
                user32.CloseClipboard()
            except Exception:
                pass

    def _paste_pipeline(self, text, hwnd):
        user32 = self._user32
        # Сохраняем оригинальный буфер обмена (None = clipboard пуст или не текст)
        original_clipboard = self._get_clipboard_text()
        self._log("info", f"Saved clipboard: {len(original_clipboard) if original_clipboard is not None else 'None'} chars")
        try:
            pyperclip.copy(text)
            self._log("info", "Clipboard: OK")
        except Exception as e:
            self._log("error", f"Clipboard FAILED: {e}")
            self._set_status("status_clipboard", "#F44336")
            return

        current_fg = user32.GetForegroundWindow()
        fg_title = _get_window_title(user32, current_fg)
        self._log("info", f"Current foreground: '{fg_title}' (hwnd={current_fg})")

        if hwnd and user32.IsWindow(hwnd):
            target_title = _get_window_title(user32, hwnd)
            self._log("info", f"Target: '{target_title}' (hwnd={hwnd})")
            if current_fg != hwnd:
                result = user32.SetForegroundWindow(hwnd)
                self._log("info", f"SetForegroundWindow: {'OK' if result else 'FAIL'}")
                self.root.after(150, lambda: self._send_paste_keys(hwnd, original_clipboard))
                return
        else:
            self._log("info", "No specific target, pasting to current foreground")

        self.root.after(50, lambda: self._send_paste_keys(hwnd, original_clipboard))

    def _send_paste_keys(self, hwnd, original_clipboard=None):
        user32 = self._user32
        fg = user32.GetForegroundWindow()
        fg_title = _get_window_title(user32, fg)
        self._log("info", f"Sending Ctrl+V → '{fg_title}' (hwnd={fg})")

        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(VK_V, 0, 0, 0)
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

        self._log("info", "Ctrl+V sent")
        # Восстанавливаем оригинальный буфер через 300ms (дать время Ctrl+V отработать)
        if original_clipboard is not None:
            self.root.after(300, lambda: self._restore_clipboard(original_clipboard))
        # Pasted — кратко показываем статус, затем Ready
        self._set_status("status_pasted", "#4CAF50", tray_state="idle")
        self.root.after(1500, lambda: self._set_status("status_ready", "#4CAF50", tray_state="idle"))

    def _cooldown_tick(self):
        remaining = self._cooldown_until - time.time()
        if remaining > 0.5:
            secs = int(remaining) + 1
            self._set_status("status_ready_in", "#FF9800", tray_state="processing", n=secs)
            self.root.after(1000, self._cooldown_tick)
        else:
            self._set_status("status_ready", "#4CAF50", tray_state="idle")

    # =====================================================================
    # Lifecycle
    # =====================================================================

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    if "--make-icon" in sys.argv:
        _generate_ico_file()
        sys.exit(0)

    if "--delayed" in sys.argv:
        time.sleep(10)

    app = VoiceTyperApp()
    app.run()
