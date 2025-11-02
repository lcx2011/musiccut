from __future__ import annotations
import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


# Try to locate VLC runtime (libvlc.dll) before importing python-vlc
_candidate_dirs = [Path(__file__).parent / "vlc", Path(__file__).parent]
env_vlc_path = os.environ.get("VLC_PATH")
if env_vlc_path:
    _candidate_dirs.append(Path(env_vlc_path))

for _candidate in _candidate_dirs:
    try:
        if (_candidate / "libvlc.dll").exists():
            os.environ["PATH"] = str(_candidate) + os.pathsep + os.environ.get("PATH", "")
            plugins_dir = _candidate / "plugins"
            if plugins_dir.exists():
                os.environ.setdefault("VLC_PLUGIN_PATH", str(plugins_dir))
            break
    except OSError:
        # Ignore invalid paths silently; python-vlc will report missing DLL later.
        continue

if sys.platform.startswith("win"):
    try:
        ctypes.windll.ole32.CoInitialize(None)
    except OSError:
        pass

from PyQt5.QtCore import Qt, QThread, QTimer, pyqtSignal, QSize
from PyQt5.QtGui import QCloseEvent, QIcon
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSlider,
    QSpinBox,
    QStatusBar,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QProgressBar,
    QWidget,
    QAbstractItemView,
)

import vlc


def format_millis(ms: int) -> str:
    total_seconds = max(0, ms // 1000)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def format_seconds(seconds: float) -> str:
    if seconds.is_integer():
        return str(int(seconds))
    return f"{seconds:.3f}".rstrip("0").rstrip(".")


def probe_media_info(path: Path, ffprobe_path: str = "ffprobe") -> tuple[int, Optional[int], Optional[int]]:
    """Return duration (seconds) and resolution using ffprobe; fallback to zeros if unavailable."""
    cmd = [
        ffprobe_path,
        "-v",
        "error",
        "-select_streams",
        "v:0",
        "-show_entries",
        "stream=width,height",
        "-show_entries",
        "format=duration",
        "-of",
        "json",
        str(path),
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    except (FileNotFoundError, subprocess.CalledProcessError):
        return 0, None, None

    try:
        data = json.loads(result.stdout or "{}")
    except json.JSONDecodeError:
        return 0, None, None

    duration = 0
    width: Optional[int] = None
    height: Optional[int] = None

    format_info = data.get("format", {})
    if "duration" in format_info:
        try:
            duration = max(0, int(float(format_info["duration"]) + 0.5))
        except (ValueError, TypeError):
            duration = 0

    streams = data.get("streams", [])
    if streams:
        stream0 = streams[0]
        width = stream0.get("width")
        height = stream0.get("height")
        try:
            width = int(width) if width is not None else None
        except (TypeError, ValueError):
            width = None
        try:
            height = int(height) if height is not None else None
        except (TypeError, ValueError):
            height = None

    return duration, width, height


@dataclass(order=True)
class TimeRange:
    start: int
    end: int

    def duration(self) -> int:
        return max(0, self.end - self.start)


@dataclass
class VideoItem:
    path: Path
    duration: int = 0  # seconds
    width: Optional[int] = None
    height: Optional[int] = None
    remove_segments: List[TimeRange] = field(default_factory=list)

    def normalize_segments(self) -> None:
        if self.duration <= 0:
            self.remove_segments = []
            return

        cleaned: List[TimeRange] = []
        for segment in sorted(self.remove_segments):
            start = max(0, min(segment.start, self.duration))
            end = max(0, min(segment.end, self.duration))
            if end <= start:
                continue
            if cleaned and start <= cleaned[-1].end:
                cleaned[-1].end = max(cleaned[-1].end, end)
            else:
                cleaned.append(TimeRange(start=start, end=end))
        self.remove_segments = cleaned

    def keep_segments(self) -> List[TimeRange]:
        self.normalize_segments()
        if self.duration <= 0:
            return []

        result: List[TimeRange] = []
        cursor = 0
        for segment in self.remove_segments:
            if segment.start > cursor:
                result.append(TimeRange(start=cursor, end=segment.start))
            cursor = max(cursor, segment.end)
        if cursor < self.duration:
            result.append(TimeRange(start=cursor, end=self.duration))
        return result

    def total_removed(self) -> int:
        self.normalize_segments()
        return sum(segment.duration() for segment in self.remove_segments)


@dataclass
class ProjectModel:
    videos: List[VideoItem] = field(default_factory=list)
    output_path: Optional[Path] = None

    def total_kept_duration(self) -> int:
        total = 0
        for video in self.videos:
            for segment in video.keep_segments():
                total += segment.duration()
        return total


class DurationSpinBox(QSpinBox):
    """Spin box that displays seconds as HH:MM:SS while storing raw seconds."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setRange(0, 24 * 3600 - 1)
        self.setSingleStep(1)

    def textFromValue(self, value: int) -> str:  # noqa: N802 - PyQt override naming
        hours = value // 3600
        minutes = (value % 3600) // 60
        seconds = value % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def valueFromText(self, text: str) -> int:  # noqa: N802
        parts = text.split(":")
        if len(parts) != 3:
            return 0
        hours, minutes, seconds = (int(part) for part in parts)
        return hours * 3600 + minutes * 60 + seconds


class VideoPlayerWidget(QWidget):
    positionChanged = pyqtSignal(int)
    durationChanged = pyqtSignal(int)
    stateChanged = pyqtSignal(bool)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._vlc_instance: Optional[vlc.Instance] = None
        self._player: Optional[vlc.MediaPlayer] = None
        self._duration_ms: int = 0

        self._video_frame = QFrame(self)
        self._video_frame.setFrameShape(QFrame.StyledPanel)
        self._video_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._video_frame)

        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(200)
        self._poll_timer.timeout.connect(self._query_playback_state)

    def open_media(self, path: Path) -> None:
        self._ensure_player()
        if self._player is None:
            return

        try:
            # Stop previous playback to ensure clean state
            self._player.stop()
        except Exception:
            pass

        media = self._vlc_instance.media_new(str(path))  # type: ignore[union-attr]
        self._player.set_media(media)

        handle = int(self._video_frame.winId())
        self._player.set_hwnd(handle)

        self._duration_ms = 0
        if self._player.play() == -1:
            # Try re-create player once more as a fallback
            try:
                self._player.release()
            except Exception:
                pass
            self._player = self._vlc_instance.media_player_new()  # type: ignore[union-attr]
            self._player.set_hwnd(handle)
            self._player.set_media(media)
            if self._player.play() == -1:
                QMessageBox.critical(self, "播放错误", "无法播放所选视频，请检查解码器。")
                return

        self._poll_timer.start()

    def _ensure_player(self) -> None:
        if self._vlc_instance is None:
            try:
                # Force software decoding and D3D9 output for Win7 stability
                self._vlc_instance = vlc.Instance(
                    "--no-video-title-show",
                    "--avcodec-hw=none",
                    "--vout=direct3d9",
                )
            except Exception as exc:  # pragma: no cover - informative popup
                QMessageBox.critical(self, "VLC 初始化失败", str(exc))
                return
        if self._player is None and self._vlc_instance is not None:
            self._player = self._vlc_instance.media_player_new()

    def play(self) -> None:
        if self._player:
            self._player.play()

    def pause(self) -> None:
        if self._player:
            self._player.pause()

    def stop(self) -> None:
        if self._player:
            self._player.stop()

    def set_time(self, ms: int) -> None:
        if self._player:
            self._player.set_time(ms)

    def current_time(self) -> int:
        return self._player.get_time() if self._player else 0

    def duration(self) -> int:
        return self._duration_ms

    def is_playing(self) -> bool:
        return bool(self._player and self._player.is_playing())

    def _query_playback_state(self) -> None:
        if self._player is None:
            self._poll_timer.stop()
            return

        current = self._player.get_time()
        if current >= 0:
            self.positionChanged.emit(current)

        length = self._player.get_length()
        if length > 0 and length != self._duration_ms:
            self._duration_ms = length
            self.durationChanged.emit(length)

        self.stateChanged.emit(self.is_playing())

    def close(self) -> None:
        self._poll_timer.stop()
        if self._player is not None:
            self._player.stop()
            self._player.release()
            self._player = None
        if self._vlc_instance is not None:
            self._vlc_instance.release()
            self._vlc_instance = None


@dataclass
class SelectionState:
    start_seconds: int = 0
    end_seconds: int = 0

    def clamp(self) -> None:
        if self.end_seconds < self.start_seconds:
            self.end_seconds = self.start_seconds


class PlaylistPanel(QWidget):
    selectionChanged = pyqtSignal(int)
    addRequested = pyqtSignal()
    removeRequested = pyqtSignal()
    moveUpRequested = pyqtSignal()
    moveDownRequested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.list_widget = QListWidget(self)
        self.list_widget.setSelectionMode(QAbstractItemView.SingleSelection)

        self.add_button = QPushButton("添加", self)
        self.remove_button = QPushButton("移除", self)
        self.up_button = QPushButton("上移", self)
        self.down_button = QPushButton("下移", self)

        button_row = QHBoxLayout()
        button_row.addWidget(self.add_button)
        button_row.addWidget(self.remove_button)
        button_row.addWidget(self.up_button)
        button_row.addWidget(self.down_button)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("播放列表", self))
        layout.addLayout(button_row)
        layout.addWidget(self.list_widget)

        self.list_widget.currentRowChanged.connect(self.selectionChanged.emit)
        self.add_button.clicked.connect(self.addRequested.emit)
        self.remove_button.clicked.connect(self.removeRequested.emit)
        self.up_button.clicked.connect(self.moveUpRequested.emit)
        self.down_button.clicked.connect(self.moveDownRequested.emit)

    def set_items(self, videos: List[VideoItem], selected_index: int) -> None:
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for idx, video in enumerate(videos):
            duration_text = format_millis(video.duration * 1000) if video.duration else "未知"
            item = QListWidgetItem(f"{idx + 1}. {video.path.name} ({duration_text})")
            self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)
        if 0 <= selected_index < len(videos):
            self.list_widget.setCurrentRow(selected_index)
        else:
            self.list_widget.setCurrentRow(-1)

    def current_index(self) -> int:
        return self.list_widget.currentRow()


class PlaybackPanel(QWidget):
    def __init__(self, player: VideoPlayerWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.player = player
        self._slider_sync_active = True

        self.play_toggle = QToolButton(self)
        self._load_icons()
        self.play_toggle.setIcon(self._icon_play)
        self.play_toggle.setIconSize(QSize(50, 50))

        self.time_label = QLabel("00:00:00 / 00:00:00", self)
        self.progress_slider = QSlider(Qt.Horizontal, self)
        self.progress_slider.setRange(0, 1000)

        button_row = QHBoxLayout()
        button_row.addWidget(self.play_toggle)
        button_row.addStretch(1)
        button_row.addWidget(self.time_label)

        layout = QVBoxLayout(self)
        layout.addWidget(self.progress_slider)
        layout.addLayout(button_row)

        self.play_toggle.clicked.connect(self._on_toggle_clicked)

        self.progress_slider.sliderPressed.connect(self._on_slider_pressed)
        self.progress_slider.sliderReleased.connect(self._on_slider_released)

        self.player.positionChanged.connect(self._on_position_changed)
        self.player.durationChanged.connect(self._on_duration_changed)
        self.player.stateChanged.connect(self._on_state_changed)

        self._duration_ms = 0

    def _on_slider_pressed(self) -> None:
        self._slider_sync_active = False

    def _on_slider_released(self) -> None:
        if self._duration_ms <= 0:
            return
        value = self.progress_slider.value() / 1000
        target_ms = int(self._duration_ms * value)
        self.player.set_time(target_ms)
        self._slider_sync_active = True

    def _on_toggle_clicked(self) -> None:
        if self.player.is_playing():
            self.player.pause()
        else:
            self.player.play()

    def _on_position_changed(self, current_ms: int) -> None:
        if self._duration_ms <= 0:
            return
        if self._slider_sync_active:
            ratio = current_ms / self._duration_ms
            self.progress_slider.blockSignals(True)
            self.progress_slider.setValue(int(ratio * 1000))
            self.progress_slider.blockSignals(False)
        self._update_time_label(current_ms)

    def _on_duration_changed(self, duration_ms: int) -> None:
        self._duration_ms = duration_ms
        self._update_time_label(self.player.current_time())

    def _update_time_label(self, current_ms: int) -> None:
        current_text = format_millis(current_ms)
        total_text = format_millis(self._duration_ms)
        self.time_label.setText(f"{current_text} / {total_text}")

    def _on_state_changed(self, playing: bool) -> None:
        self.play_toggle.setIcon(self._icon_pause if playing else self._icon_play)

    def _load_icons(self) -> None:
        icon_dir = Path(__file__).parent / "icon"
        play_path = icon_dir / "开始.svg"
        pause_path = icon_dir / "暂停.svg"
        # Fallback to text if icons missing
        self._icon_play = QIcon(str(play_path)) if play_path.exists() else QIcon()
        self._icon_pause = QIcon(str(pause_path)) if pause_path.exists() else QIcon()
        if self._icon_play.isNull() or self._icon_pause.isNull():
            # ensure some visible label when icon loading fails
            self.play_toggle.setText("▶")


class TrimPanel(QWidget):
    segmentsChanged = pyqtSignal()
    outputPathEdited = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.selection = SelectionState()
        self._duration_seconds = 0
        self._video: Optional[VideoItem] = None

        self.info_labels = {
            "文件": QLabel("-"),
            "分辨率": QLabel("-"),
            "时长": QLabel("00:00:00"),
        }

        self.in_spin = DurationSpinBox(self)
        self.out_spin = DurationSpinBox(self)
        self.out_spin.setValue(0)

        self.set_in_btn = QPushButton("设为起点", self)
        self.set_out_btn = QPushButton("设为终点", self)

        self.add_segment_btn = QPushButton("添加删除段", self)
        self.remove_segment_btn = QPushButton("删除选中段", self)
        self.clear_segment_btn = QPushButton("清空删除段", self)

        self.segment_list = QListWidget(self)
        self.segment_list.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.output_path_edit = QLineEdit(self)
        self.output_browse_btn = QPushButton("浏览", self)

        self.run_btn = QPushButton("开始剪辑", self)
        self.run_btn.setEnabled(False)
        self.cancel_btn = QPushButton("取消", self)
        self.cancel_btn.setEnabled(False)

        self.selection_label = QLabel("当前区间: 00:00:00", self)
        self.removed_label = QLabel("删除总时长: 00:00:00", self)
        self.kept_label = QLabel("保留总时长: 00:00:00", self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪")

        info_layout = QFormLayout()
        for label, widget in self.info_labels.items():
            info_layout.addRow(label + "：", widget)

        in_row = QHBoxLayout()
        in_row.addWidget(self.in_spin)
        in_row.addWidget(self.set_in_btn)

        out_row = QHBoxLayout()
        out_row.addWidget(self.out_spin)
        out_row.addWidget(self.set_out_btn)

        segment_button_row = QHBoxLayout()
        segment_button_row.addWidget(self.add_segment_btn)
        segment_button_row.addWidget(self.remove_segment_btn)
        segment_button_row.addWidget(self.clear_segment_btn)

        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path_edit)
        output_row.addWidget(self.output_browse_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(info_layout)
        layout.addSpacing(10)
        layout.addWidget(QLabel("剪辑起点", self))
        layout.addLayout(in_row)
        layout.addWidget(QLabel("剪辑终点", self))
        layout.addLayout(out_row)
        layout.addSpacing(5)
        layout.addWidget(QLabel("删除片段列表", self))
        layout.addLayout(segment_button_row)
        layout.addWidget(self.segment_list)
        layout.addSpacing(10)
        layout.addWidget(QLabel("输出文件", self))
        layout.addLayout(output_row)
        layout.addWidget(self.selection_label)
        layout.addWidget(self.removed_label)
        layout.addWidget(self.kept_label)
        layout.addWidget(self.progress_bar)
        layout.addSpacing(10)
        layout.addWidget(self.run_btn)
        layout.addWidget(self.cancel_btn)
        layout.addStretch(1)

        self.in_spin.valueChanged.connect(self._on_selection_changed)
        self.out_spin.valueChanged.connect(self._on_selection_changed)
        self.segment_list.currentRowChanged.connect(self._on_segment_selected)
        self.add_segment_btn.clicked.connect(self._add_segment)
        self.remove_segment_btn.clicked.connect(self._remove_selected_segments)
        self.clear_segment_btn.clicked.connect(self._clear_segments)
        self.output_path_edit.textChanged.connect(self.outputPathEdited.emit)

    def set_video(self, video: Optional[VideoItem]) -> None:
        self._video = video
        if video is None:
            for label in self.info_labels.values():
                label.setText("-")
            self.info_labels["时长"].setText("00:00:00")
            self._duration_seconds = 0
            self.in_spin.setRange(0, 0)
            self.out_spin.setRange(0, 0)
            self.in_spin.setValue(0)
            self.out_spin.setValue(0)
            self.selection = SelectionState()
            self.segment_list.clear()
            self.removed_label.setText("删除总时长: 00:00:00")
            self.kept_label.setText("保留总时长: 00:00:00")
            self.selection_label.setText("当前区间: 00:00:00")
            self.run_btn.setEnabled(False)
            return

        duration_ms = video.duration * 1000
        resolution = (
            f"{video.width}x{video.height}" if video.width and video.height else "未知"
        )
        self.info_labels["文件"].setText(video.path.name)
        self.info_labels["分辨率"].setText(resolution)
        self.info_labels["时长"].setText(format_millis(duration_ms))
        self._duration_seconds = max(0, video.duration)
        self.in_spin.setRange(0, self._duration_seconds)
        self.out_spin.setRange(0, self._duration_seconds)
        self.in_spin.setValue(0)
        self.out_spin.setValue(self._duration_seconds)
        self.selection = SelectionState(0, self._duration_seconds)
        self._refresh_segment_list()
        self._update_selection_label()
        self.run_btn.setEnabled(True)

    def set_start_from_ms(self, ms: int) -> None:
        self.in_spin.setValue(min(ms // 1000, self.out_spin.value()))

    def set_end_from_ms(self, ms: int) -> None:
        seconds = ms // 1000
        if seconds <= self.in_spin.value():
            seconds = min(self.in_spin.value() + 1, self._duration_seconds)
        self.out_spin.setValue(seconds)

    def _on_selection_changed(self) -> None:
        self.selection.start_seconds = self.in_spin.value()
        self.selection.end_seconds = self.out_spin.value()
        self.selection.clamp()
        if self.out_spin.value() != self.selection.end_seconds:
            self.out_spin.blockSignals(True)
            self.out_spin.setValue(self.selection.end_seconds)
            self.out_spin.blockSignals(False)
        self._update_selection_label()

    def _update_selection_label(self) -> None:
        selected = max(0, self.selection.end_seconds - self.selection.start_seconds)
        self.selection_label.setText(
            f"当前区间: {format_millis(selected * 1000)}"
        )

    def _on_segment_selected(self, row: int) -> None:
        if self._video is None or row < 0 or row >= len(self._video.remove_segments):
            return
        segment = self._video.remove_segments[row]
        self.in_spin.setValue(segment.start)
        self.out_spin.setValue(segment.end)

    def _add_segment(self) -> None:
        if self._video is None:
            return
        start = self.in_spin.value()
        end = self.out_spin.value()
        if end <= start:
            QMessageBox.warning(self, "无效区间", "终点必须晚于起点。")
            return
        self._video.remove_segments.append(TimeRange(start=start, end=end))
        self._video.normalize_segments()
        self._refresh_segment_list()
        self.segmentsChanged.emit()

    def _remove_selected_segments(self) -> None:
        if self._video is None:
            return
        selected_rows = sorted({index.row() for index in self.segment_list.selectedIndexes()}, reverse=True)
        if not selected_rows:
            return
        for row in selected_rows:
            if 0 <= row < len(self._video.remove_segments):
                self._video.remove_segments.pop(row)
        self._video.normalize_segments()
        self._refresh_segment_list()
        self.segmentsChanged.emit()

    def _clear_segments(self) -> None:
        if self._video is None:
            return
        self._video.remove_segments.clear()
        self._refresh_segment_list()
        self.segmentsChanged.emit()

    def _refresh_segment_list(self) -> None:
        self.segment_list.clear()
        if self._video is None:
            return
        self._video.normalize_segments()
        for segment in self._video.remove_segments:
            duration_text = format_millis(segment.duration() * 1000)
            item_text = (
                f"删除 {format_millis(segment.start * 1000)} - "
                f"{format_millis(segment.end * 1000)} (时长 {duration_text})"
            )
            self.segment_list.addItem(item_text)
        removed = self._video.total_removed()
        kept = max(0, self._duration_seconds - removed)
        self.removed_label.setText(f"删除总时长: {format_millis(removed * 1000)}")
        self.kept_label.setText(f"保留总时长: {format_millis(kept * 1000)}")

    def reset_progress(self) -> None:
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪")

    def set_progress(self, ratio: float, message: str) -> None:
        value = max(0, min(int(ratio * 1000), 1000))
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(message)


@dataclass
class FFmpegTask:
    input_path: Path
    output_path: Path
    start_seconds: int
    end_seconds: int
    total_duration: int


class FFmpegWorker(QThread):
    progress = pyqtSignal(float)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    # Matches time=00:01:23.45
    _time_pattern = re.compile(r"time=(\d+):(\d+):(\d+\.\d+)")

    def __init__(self, ffmpeg_path: str, task: FFmpegTask) -> None:
        super().__init__()
        self._ffmpeg_path = ffmpeg_path
        self._task = task
        self._process: Optional[subprocess.Popen[str]] = None
        self._cancel_requested = False

    def run(self) -> None:  # noqa: D401 - PyQt thread entry
        try:
            self._execute()
        except FileNotFoundError:
            self.finished.emit(False, "未找到 FFmpeg，请确认路径或已安装。")
        except Exception as exc:  # pragma: no cover - runtime error path
            self.finished.emit(False, f"执行失败: {exc}")

    def request_cancel(self) -> None:
        self._cancel_requested = True
        if self._process and self._process.poll() is None:
            self._process.terminate()

    def _execute(self) -> None:
        task = self._task
        duration = max(1, task.end_seconds - task.start_seconds)

        args = [
            self._ffmpeg_path,
            "-y",
            "-hide_banner",
            "-loglevel",
            "info",
            "-ss",
            str(task.start_seconds),
            "-to",
            str(task.end_seconds),
            "-i",
            str(task.input_path),
            "-c",
            "copy",
            str(task.output_path),
        ]

        self.message.emit("FFmpeg 命令: " + " ".join(args))

        self._process = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )

        assert self._process.stdout is not None  # for type checker
        for line in self._process.stdout:
            if self._cancel_requested:
                self.finished.emit(False, "任务已取消")
                return
            line = line.strip()
            self.message.emit(line)
            ratio = self._parse_progress(line, duration)
            if ratio is not None:
                self.progress.emit(ratio)

        return_code = self._process.wait()
        if self._cancel_requested:
            self.finished.emit(False, "任务已取消")
            return
        if return_code == 0:
            self.progress.emit(1.0)
            self.finished.emit(True, f"导出完成: {task.output_path}")
        else:
            self.finished.emit(False, f"FFmpeg 退出失败 (code {return_code})")

    def _parse_progress(self, line: str, duration_seconds: int) -> Optional[float]:
        match = self._time_pattern.search(line)
        if not match:
            return None

        hours = int(match.group(1))
        minutes = int(match.group(2))
        seconds = float(match.group(3))
        current_seconds = hours * 3600 + minutes * 60 + seconds
        ratio = current_seconds / max(1.0, duration_seconds)
        return min(1.0, max(0.0, ratio))


class MultiStepFFmpegWorker(QThread):
    progress = pyqtSignal(float)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, steps: List[tuple[List[str], float]], cleanup: Optional[List[Path]] = None) -> None:
        super().__init__()
        self._steps = steps
        self._cleanup = cleanup or []
        self._cancel = False

    def request_cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        try:
            total_weight = sum(weight for _, weight in self._steps) or 1.0
            acc = 0.0
            for idx, (args, weight) in enumerate(self._steps, start=1):
                if self._cancel:
                    self.finished.emit(False, "任务已取消")
                    return
                self.message.emit("执行: " + " ".join(args))
                proc = subprocess.Popen(
                    args,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="ignore",
                )
                # Best effort: simple spinner without fine-grained time parsing for each step
                while True:
                    if self._cancel:
                        proc.terminate()
                        self.finished.emit(False, "任务已取消")
                        return
                    line = proc.stdout.readline() if proc.stdout else ""
                    if not line and proc.poll() is not None:
                        break
                code = proc.wait()
                if code != 0:
                    self._do_cleanup()
                    self.finished.emit(False, f"步骤 {idx} 执行失败 (code {code})")
                    return
                acc += weight
                self.progress.emit(min(1.0, acc / total_weight))
            self._do_cleanup()
            self.finished.emit(True, "导出完成")
        except FileNotFoundError as e:
            self._do_cleanup()
            self.finished.emit(False, f"缺少可执行文件: {e}")
        except Exception as exc:
            self._do_cleanup()
            self.finished.emit(False, f"执行失败: {exc}")

    def _do_cleanup(self) -> None:
        for path in self._cleanup:
            try:
                if path.is_file():
                    if path.exists():
                        path.unlink()
            except Exception:
                pass


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("视频剪辑助手 (Python + FFmpeg)")
        self.resize(960, 600)

        self.player_widget = VideoPlayerWidget(self)
        self.playback_panel = PlaybackPanel(self.player_widget, self)
        self.trim_panel = TrimPanel(self)
        self.playlist_panel = PlaylistPanel(self)

        # Layout: playlist | player+controls | trim
        middle_column = QVBoxLayout()
        middle_column.addWidget(self.player_widget, stretch=1)
        middle_column.addWidget(self.playback_panel, stretch=0)

        middle_container = QWidget(self)
        middle_container.setLayout(middle_column)

        central_layout = QHBoxLayout()
        central_layout.addWidget(self.playlist_panel, stretch=2)
        central_layout.addWidget(middle_container, stretch=5)
        central_layout.addWidget(self.trim_panel, stretch=3)

        central_widget = QWidget(self)
        central_widget.setLayout(central_layout)
        self.setCentralWidget(central_widget)

        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

        self._project = ProjectModel()
        self._current_index: int = -1
        self._ffmpeg_thread: Optional[QThread] = None

        self._create_actions()
        self._create_toolbar()

        self.player_widget.durationChanged.connect(self._on_duration_changed)

        self.trim_panel.set_in_btn.clicked.connect(self._set_start_point)
        self.trim_panel.set_out_btn.clicked.connect(self._set_end_point)
        self.trim_panel.output_browse_btn.clicked.connect(self._browse_output)
        self.trim_panel.run_btn.clicked.connect(self._start_export)
        self.trim_panel.cancel_btn.clicked.connect(self._cancel_trim)
        self.trim_panel.segmentsChanged.connect(self._on_segments_changed)
        self.trim_panel.outputPathEdited.connect(self._on_output_path_edited)

        self.playlist_panel.selectionChanged.connect(self._on_playlist_selection_changed)
        self.playlist_panel.addRequested.connect(self._add_videos)
        self.playlist_panel.removeRequested.connect(self._remove_selected_video)
        self.playlist_panel.moveUpRequested.connect(self._move_up)
        self.playlist_panel.moveDownRequested.connect(self._move_down)

    def _create_actions(self) -> None:
        self.open_action = QAction("导入视频", self)
        self.open_action.triggered.connect(self._add_videos)

        self.exit_action = QAction("退出", self)
        self.exit_action.triggered.connect(self.close)

    def _create_toolbar(self) -> None:
        toolbar = QToolBar("Main", self)
        toolbar.setIconSize(toolbar.iconSize())
        toolbar.addAction(self.open_action)
        toolbar.addSeparator()
        toolbar.addAction(self.exit_action)
        self.addToolBar(Qt.TopToolBarArea, toolbar)

    def _add_videos(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "添加视频文件",
            str(Path.home()),
            "视频文件 (*.mp4 *.mov *.mkv *.avi *.flv *.wmv);;所有文件 (*.*)",
        )
        if not paths:
            return
        new_items: list[VideoItem] = []
        for p in paths:
            path = Path(p)
            if not path.exists():
                continue
            dur, w, h = probe_media_info(path)
            new_items.append(VideoItem(path=path, duration=dur, width=w, height=h))
        self._project.videos.extend(new_items)
        if self._current_index < 0 and self._project.videos:
            self._current_index = 0
        self._refresh_playlist()
        self._load_current_video()

    def _on_duration_changed(self, duration_ms: int) -> None:
        # Keep the time label via playback panel; metadata already set via model
        pass

    def _set_start_point(self) -> None:
        current_ms = self.player_widget.current_time()
        self.trim_panel.set_start_from_ms(current_ms)

    def _set_end_point(self) -> None:
        current_ms = self.player_widget.current_time()
        self.trim_panel.set_end_from_ms(current_ms)

    def _browse_output(self) -> None:
        suggested = self._project.output_path or (Path.home() / "output_trimmed.mp4")
        path_str, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出位置",
            str(suggested),
            "视频文件 (*.mp4)",
        )
        if path_str:
            self.trim_panel.output_path_edit.setText(path_str)
            self._project.output_path = Path(path_str)

    def _start_export(self) -> None:
        if not self._project.videos:
            QMessageBox.warning(self, "没有视频", "请先添加至少一个视频。")
            return
        output_path_text = self.trim_panel.output_path_edit.text().strip()
        if not output_path_text:
            self._browse_output()
            output_path_text = self.trim_panel.output_path_edit.text().strip()
        if not output_path_text:
            return
        output_path = Path(output_path_text)

        # Build steps: cut keep-segments for each video, then concat
        temp_dir = Path(tempfile.mkdtemp(prefix="trimmerge_"))
        steps: list[tuple[list[str], float]] = []
        temp_files: list[Path] = []
        list_file = temp_dir / "concat_list.txt"
        with list_file.open("w", encoding="utf-8") as f:
            pass

        # Decide whether to transcode each segment to unify codecs/parameters
        transcode = len(self._project.videos) > 1
        # Scheme C: target resolution = max width/height across inputs; fps fixed to 30
        target_w = 0
        target_h = 0
        for v in self._project.videos:
            if v.width and v.width > target_w:
                target_w = v.width
            if v.height and v.height > target_h:
                target_h = v.height
        if target_w <= 0 or target_h <= 0:
            target_w, target_h = 1280, 720
        fps_target = 30

        # Calculate weights as kept durations
        for v_idx, video in enumerate(self._project.videos):
            keeps = video.keep_segments() or []
            for k_idx, keep in enumerate(keeps):
                out_clip = temp_dir / f"clip_{v_idx:02d}_{k_idx:03d}.mp4"
                if transcode:
                    # Re-encode to uniform H.264/AAC and normalized geometry/fps
                    vf_chain = (
                        f"scale={target_w}:{target_h}:force_original_aspect_ratio=decrease,"
                        f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2,"
                        f"setsar=1,format=yuv420p,fps={fps_target}"
                    )
                    args = [
                        "ffmpeg",
                        "-y",
                        "-hide_banner",
                        "-loglevel",
                        "error",
                        "-ss",
                        str(keep.start),
                        "-to",
                        str(keep.end),
                        "-i",
                        str(video.path),
                        "-vf",
                        vf_chain,
                        "-c:v",
                        "libx264",
                        "-preset",
                        "veryfast",
                        "-crf",
                        "20",
                        "-c:a",
                        "aac",
                        "-ar",
                        "48000",
                        "-ac",
                        "2",
                        "-b:a",
                        "192k",
                        str(out_clip),
                    ]
                else:
                    # Single source: try stream copy for speed
                    args = [
                        "ffmpeg",
                        "-y",
                        "-hide_banner",
                        "-loglevel",
                        "error",
                        "-ss",
                        str(keep.start),
                        "-to",
                        str(keep.end),
                        "-i",
                        str(video.path),
                        "-c",
                        "copy",
                        str(out_clip),
                    ]
                steps.append((args, float(keep.duration())))
                temp_files.append(out_clip)
        # Write concat list
        with list_file.open("w", encoding="utf-8") as f:
            for p in temp_files:
                f.write(f"file '{p.as_posix()}'\n")

        concat_args = [
            "ffmpeg",
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(list_file),
            "-c",
            "copy",
            str(output_path),
        ]
        steps.append((concat_args, 1.0))

        self._run_multistep(steps, temp_files + [list_file])

    def _run_ffmpeg(self, task: "FFmpegTask") -> None:
        if self._ffmpeg_thread is not None:
            QMessageBox.information(self, "正在处理", "已有任务正在执行，请稍候。")
            return

        ffmpeg_path = "ffmpeg"
        self._ffmpeg_thread = FFmpegWorker(ffmpeg_path, task)

        self._ffmpeg_thread.progress.connect(self._on_ffmpeg_progress)
        self._ffmpeg_thread.message.connect(self.status_bar.showMessage)
        self._ffmpeg_thread.finished.connect(self._on_ffmpeg_finished)

        self.trim_panel.run_btn.setEnabled(False)
        self.trim_panel.cancel_btn.setEnabled(True)
        self.trim_panel.set_progress(0.0, "准备中…")
        self.status_bar.showMessage("开始执行 FFmpeg 裁剪…")

        self._ffmpeg_thread.start()

    def _on_ffmpeg_progress(self, ratio: float) -> None:
        percent = ratio * 100
        self.trim_panel.set_progress(ratio, f"处理中 {percent:.1f}%")

    def _on_ffmpeg_finished(self, success: bool, message: str) -> None:
        self.status_bar.showMessage(message, 5000)
        self.trim_panel.set_progress(1.0 if success else 0.0, message)
        self.trim_panel.run_btn.setEnabled(True)
        self.trim_panel.cancel_btn.setEnabled(False)
        self._ffmpeg_thread = None
        if success:
            QMessageBox.information(self, "剪辑完成", message)
        else:
            QMessageBox.warning(self, "剪辑失败", message)

    def _cancel_trim(self) -> None:
        if self._ffmpeg_thread is None:
            return
        # Try to cancel both worker types
        if isinstance(self._ffmpeg_thread, FFmpegWorker):
            self._ffmpeg_thread.request_cancel()
        elif isinstance(self._ffmpeg_thread, MultiStepFFmpegWorker):
            self._ffmpeg_thread.request_cancel()
        self.status_bar.showMessage("正在取消任务…")

    def _run_multistep(self, steps: List[tuple[List[str], float]], cleanup: List[Path]) -> None:
        if self._ffmpeg_thread is not None:
            QMessageBox.information(self, "正在处理", "已有任务正在执行，请稍候。")
            return
        worker = MultiStepFFmpegWorker(steps, cleanup)
        self._ffmpeg_thread = worker
        worker.progress.connect(self._on_ffmpeg_progress)
        worker.message.connect(self.status_bar.showMessage)
        worker.finished.connect(self._on_ffmpeg_finished)
        self.trim_panel.run_btn.setEnabled(False)
        self.trim_panel.cancel_btn.setEnabled(True)
        self.trim_panel.set_progress(0.0, "准备中…")
        self.status_bar.showMessage("开始执行合并导出…")
        worker.start()

    def _on_segments_changed(self) -> None:
        # Refresh kept/removed labels already handled in TrimPanel
        pass

    def _on_output_path_edited(self, text: str) -> None:
        self._project.output_path = Path(text) if text else None

    def _refresh_playlist(self) -> None:
        self.playlist_panel.set_items(self._project.videos, self._current_index)

    def _on_playlist_selection_changed(self, index: int) -> None:
        self._current_index = index
        self._load_current_video()

    def _load_current_video(self) -> None:
        if 0 <= self._current_index < len(self._project.videos):
            video = self._project.videos[self._current_index]
            self.trim_panel.set_video(video)
            self.player_widget.open_media(video.path)
        else:
            self.trim_panel.set_video(None)

    def _remove_selected_video(self) -> None:
        idx = self.playlist_panel.current_index()
        if 0 <= idx < len(self._project.videos):
            self._project.videos.pop(idx)
            if self._current_index >= len(self._project.videos):
                self._current_index = len(self._project.videos) - 1
            self._refresh_playlist()
            self._load_current_video()

    def _move_up(self) -> None:
        idx = self.playlist_panel.current_index()
        if 1 <= idx < len(self._project.videos):
            self._project.videos[idx - 1], self._project.videos[idx] = (
                self._project.videos[idx],
                self._project.videos[idx - 1],
            )
            self._current_index = idx - 1
            self._refresh_playlist()
            self.playlist_panel.set_items(self._project.videos, self._current_index)

    def _move_down(self) -> None:
        idx = self.playlist_panel.current_index()
        if 0 <= idx < len(self._project.videos) - 1:
            self._project.videos[idx + 1], self._project.videos[idx] = (
                self._project.videos[idx],
                self._project.videos[idx + 1],
            )
            self._current_index = idx + 1
            self._refresh_playlist()
            self.playlist_panel.set_items(self._project.videos, self._current_index)

    def closeEvent(self, event: QCloseEvent) -> None:
        self.player_widget.close()
        super().closeEvent(event)


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()


@dataclass
class FFmpegTask:
    input_path: Path
    output_path: Path
    start_seconds: int
    end_seconds: int
    total_duration: int


class FFmpegWorker(QThread):
    progress = pyqtSignal(float)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    # Matches time=00:01:23.45
    _time_pattern = re.compile(r"time=(\d+):(\d+):(\d+\.\d+)")

    def __init__(self, ffmpeg_path: str, task: FFmpegTask) -> None:
        super().__init__()
        self._ffmpeg_path = ffmpeg_path
        self._task = task
        self._process: Optional[subprocess.Popen[str]] = None
        self._cancel_requested = False

    def run(self) -> None:  # noqa: D401 - PyQt thread entry
        try:
            self._execute()
        except FileNotFoundError:
            self.finished.emit(False, "未找到 FFmpeg，请确认路径或已安装。")
        except Exception as exc:  # pragma: no cover - runtime error path
            self.finished.emit(False, f"执行失败: {exc}")

    def request_cancel(self) -> None:
        self._cancel_requested = True
        if self._process and self._process.poll() is None:
            self._process.terminate()

    def _execute(self) -> None:
        task = self._task
        duration = max(1, task.end_seconds - task.start_seconds)

        args = [
            self._ffmpeg_path,
            "-y",
            "-hide_banner",
            "-loglevel",
            "info",
            "-ss",
            str(task.start_seconds),
            "-to",
            str(task.end_seconds),
            "-i",
            str(task.input_path),
            "-c",
            "copy",
            str(task.output_path),
        ]

        self.message.emit("FFmpeg 命令: " + " ".join(args))

        self._process = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )

        assert self._process.stdout is not None  # for type checker
        for line in self._process.stdout:
            if self._cancel_requested:
                self.finished.emit(False, "任务已取消")
                return
            line = line.strip()
            self.message.emit(line)
            ratio = self._parse_progress(line, duration)
            if ratio is not None:
                self.progress.emit(ratio)

        return_code = self._process.wait()
        if self._cancel_requested:
            self.finished.emit(False, "任务已取消")
            return
        if return_code == 0:
            self.progress.emit(1.0)
            self.finished.emit(True, f"导出完成: {task.output_path}")
        else:
            self.finished.emit(False, f"FFmpeg 退出失败 (code {return_code})")

    def _parse_progress(self, line: str, duration_seconds: int) -> Optional[float]:
        match = self._time_pattern.search(line)
        if not match:
            return None

        hours = int(match.group(1))
        minutes = int(match.group(2))
        seconds = float(match.group(3))
        current_seconds = hours * 3600 + minutes * 60 + seconds
        ratio = current_seconds / max(1.0, duration_seconds)
        return min(1.0, max(0.0, ratio))
