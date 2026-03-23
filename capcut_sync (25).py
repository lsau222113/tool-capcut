#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CapCut Timeline Sync Tool
Sắp xếp clip ghép (video) khớp theo thứ tự audio từ trái sang phải
"""

import json
import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import glob
from datetime import datetime
import threading

# ── Màu sắc & style ────────────────────────────────────────────
BG       = "#0d0d14"
SURFACE  = "#13131e"
SURFACE2 = "#1c1c2e"
BORDER   = "#2a2a3e"
ACCENT   = "#5b6ef5"
ACCENT2  = "#f55b8e"
GOLD     = "#f5c842"
SUCCESS  = "#42f5a7"
WARN     = "#f5a742"
TEXT     = "#e8e8f0"
MUTED    = "#6b6b8a"


def find_capcut_projects():
    """Tìm tất cả project CapCut, không trùng lặp."""
    appdata     = os.environ.get("LOCALAPPDATA", "")
    userprofile = os.environ.get("USERPROFILE", "")
    username    = os.environ.get("USERNAME", "")

    # Tất cả đường dẫn có thể chứa CapCut projects
    candidates = set()
    for base in [appdata, os.path.join(userprofile, "AppData", "Local")]:
        if base:
            candidates.add(os.path.join(base, "CapCut", "User Data", "Projects", "com.lveditor.draft"))
            candidates.add(os.path.join(base, "Programs", "CapCut", "User Data", "Projects", "com.lveditor.draft"))

    # Quét thêm ổ đĩa khác nếu cần
    for drive in ["C:", "D:", "E:", "F:"]:
        p = os.path.join(drive, os.sep, "Users", username, "AppData", "Local",
                         "CapCut", "User Data", "Projects", "com.lveditor.draft")
        candidates.add(p)

    # Resolve realpath để loại trùng lặp do symlink / chữ hoa chữ thường
    seen_roots = set()
    valid_roots = []
    for c in candidates:
        if not os.path.isdir(c):
            continue
        real = os.path.normcase(os.path.realpath(c))
        if real not in seen_roots:
            seen_roots.add(real)
            valid_roots.append(c)

    # Quét từng project folder
    seen_projects = set()
    projects = []
    for root in valid_roots:
        try:
            entries = os.listdir(root)
        except Exception:
            continue
        for folder in entries:
            proj_path = os.path.join(root, folder)
            json_path = os.path.join(proj_path, "draft_content.json")
            if not os.path.isfile(json_path):
                continue

            # Dùng realpath để dedup project (tránh hiện 2 lần)
            real_json = os.path.normcase(os.path.realpath(json_path))
            if real_json in seen_projects:
                continue
            seen_projects.add(real_json)

            # Lấy tên project từ draft_meta_info.json
            name = folder
            meta_path = os.path.join(proj_path, "draft_meta_info.json")
            try:
                if os.path.isfile(meta_path):
                    with open(meta_path, "r", encoding="utf-8") as f:
                        meta = json.load(f)
                    name = meta.get("draft_name", "") or meta.get("name", "") or folder
            except Exception:
                pass

            mtime = os.path.getmtime(json_path)
            projects.append({
                "name":  name,
                "path":  proj_path,
                "json":  json_path,
                "mtime": mtime,
            })

    # Mới nhất lên đầu
    projects.sort(key=lambda x: x["mtime"], reverse=True)
    return projects


def get_material_name(material_id, materials):
    """Lấy tên của material theo id. Hỗ trợ cả clip ghép (compound clip)."""
    all_mats = []
    all_mats.extend(materials.get("videos", []))
    all_mats.extend(materials.get("audios", []))
    all_mats.extend(materials.get("sounds", []))

    for m in all_mats:
        if m.get("id") == material_id:
            # Ưu tiên material_name (tên clip ghép như "Clip ghép21")
            mat_name = m.get("material_name", "")
            if mat_name:
                return mat_name
            # Nếu có path (video thường)
            path = m.get("path", "") or m.get("file_Path", "")
            if path:
                return os.path.basename(path)
            # Fallback: name field
            return m.get("name", material_id)
    return str(material_id)


def get_material_duration(material_id, materials):
    """
    Lấy duration thực của clip ghép (compound clip).
    Với clip ghép, duration trong materials.videos là duration của clip ghép đó,
    KHÔNG phải duration của video gốc.
    """
    for m in materials.get("videos", []):
        if m.get("id") == material_id:
            return m.get("duration", 0)
    return 0


def analyze_and_sync(draft_data):
    """
    Sắp xếp và chỉnh tốc độ video clips khớp với audio:
    - Clip video thứ i ghép với audio thứ i (theo thứ tự trái→phải)
    - Clip xếp liền nhau từ start=0
    - target_timerange.duration = audio_duration  (clip dài bằng audio)
    - source_timerange giữ nguyên (không cắt thêm/bớt footage)
    - speed = source_duration / audio_duration    (tự động tăng/giảm tốc)
    Returns: (mapping_list, log_lines, stats)
    """
    logs = []
    tracks   = draft_data.get("tracks", [])
    materials= draft_data.get("materials", {})

    # Tìm video track chính (nhiều segment nhất)
    video_track  = None
    audio_tracks = []
    for track in tracks:
        t    = track.get("type", "")
        segs = track.get("segments", [])
        if t == "video":
            if video_track is None or len(segs) > len(video_track.get("segments",[])):
                video_track = track
        elif t == "audio":
            audio_tracks.append(track)

    if not video_track:
        logs.append(("ERR", "Không tìm thấy video track!"))
        return None, logs, {}

    video_segs = video_track.get("segments", [])

    # Audio segments — sort theo start (trái→phải)
    audio_segs = []
    for atrack in audio_tracks:
        for seg in atrack.get("segments", []):
            name = get_material_name(seg.get("material_id",""), materials)
            audio_segs.append({"seg": seg, "name": name})
    audio_segs.sort(key=lambda x: x["seg"].get("target_timerange",{}).get("start",0))

    # Video segments — sort theo thứ tự ổn định:
    # - Segment đã chạy tool: có _sync_order → dùng để giữ thứ tự cố định
    # - Segment mới thêm: không có _sync_order → append theo start, sau các segment cũ
    old_segs = [s for s in video_segs if s.get("_sync_order") is not None]
    new_segs = [s for s in video_segs if s.get("_sync_order") is None]

    if old_segs:
        old_segs_sorted = sorted(old_segs, key=lambda s: s.get("_sync_order", 9999))
        new_segs_sorted = sorted(new_segs, key=lambda s: s.get("target_timerange",{}).get("start", 0))
        video_segs_sorted = old_segs_sorted + new_segs_sorted
        logs.append(("INFO", f"Clip cũ (giữ thứ tự): {len(old_segs)}  |  Clip mới thêm: {len(new_segs)}"))
    else:
        video_segs_sorted = sorted(
            video_segs,
            key=lambda s: s.get("target_timerange",{}).get("start", 0)
        )
        logs.append(("INFO", "Lần đầu chạy — sort theo vị trí start hiện tại"))

    logs.append(("INFO", f"Video clips: {len(video_segs_sorted)}  |  Audio lines: {len(audio_segs)}"))
    logs.append(("---",""))

    logs.append(("INFO","AUDIO (trái→phải):"))
    for i,a in enumerate(audio_segs):
        dur = a["seg"].get("target_timerange",{}).get("duration",0)/1e6
        logs.append(("AUDIO", f"  [{i+1}] {a['name']}  {dur:.2f}s"))

    logs.append(("---",""))
    logs.append(("INFO","VIDEO hiện tại:"))
    for i,v in enumerate(video_segs_sorted):
        name     = get_material_name(v.get("material_id",""), materials)
        tgt_dur  = v.get("target_timerange",{}).get("duration",0)/1e6
        src_dur  = v.get("source_timerange",{}).get("duration",0)/1e6
        spd      = v.get("speed", 1.0)
        logs.append(("VIDEO", f"  [{i+1}] {name}  tgt={tgt_dur:.2f}s  src={src_dur:.2f}s  speed={spd:.3f}x"))

    logs.append(("---",""))

    n       = min(len(video_segs_sorted), len(audio_segs))
    mapping = []
    cursor  = 0

    for i in range(n):
        vseg  = video_segs_sorted[i]
        aseg  = audio_segs[i]
        vname = get_material_name(vseg.get("material_id",""), materials)
        aname = aseg["name"]

        audio_dur = aseg["seg"].get("target_timerange",{}).get("duration", 0)
        mat_id    = vseg.get("material_id", "")

        # ── Lấy footage duration gốc ───────────────────────────────────────
        # Thứ tự ưu tiên:
        # 1. _sync_orig_dur: giá trị gốc đã lưu từ lần chạy trước (LUÔN đúng)
        # 2. material.duration: lần đầu chạy
        # 3. source_timerange.duration: fallback cuối
        footage_dur = (
            vseg.get("_sync_orig_dur", 0)
            or get_material_duration(mat_id, materials)
            or vseg.get("source_timerange",{}).get("duration", 0)
        )

        src_start = vseg.get("source_timerange",{}).get("start", 0)

        if audio_dur > 0 and footage_dur > 0:
            new_speed = round(footage_dur / audio_dur, 6)
        else:
            new_speed = 1.0
        new_speed = max(0.1, min(100.0, new_speed))

        mapping.append({
            "index":          i,
            "segment_id":     vseg.get("id",""),
            "video_seg":      vseg,
            "video_name":     vname,
            "audio_name":     aname,
            "new_start":      cursor,
            "new_target_dur": audio_dur,
            "new_speed":      new_speed,
            "src_dur":        footage_dur,
            "src_start":      src_start,
            "leftover":       False,
        })
        logs.append(("PAIR",
            f"[{i+1}] {vname} ({footage_dur/1e6:.2f}s) ↔ {aname} ({audio_dur/1e6:.2f}s) "
            f"→ speed={new_speed:.3f}x"))
        cursor += audio_dur

    # Clip thừa — giữ speed gốc, xếp sau
    for j in range(n, len(video_segs_sorted)):
        vseg     = video_segs_sorted[j]
        vname    = get_material_name(vseg.get("material_id",""), materials)
        tgt_dur  = vseg.get("target_timerange",{}).get("duration",0)
        src_dur  = vseg.get("source_timerange",{}).get("duration",0)
        src_start= vseg.get("source_timerange",{}).get("start",0)
        spd      = vseg.get("speed", 1.0)
        mapping.append({
            "index":          n + j - n,
            "segment_id":     vseg.get("id",""),
            "video_seg":      vseg,
            "video_name":     vname,
            "audio_name":     None,
            "new_start":      cursor,
            "new_target_dur": tgt_dur,
            "new_speed":      spd,
            "src_dur":        src_dur,
            "src_start":      src_start,
            "leftover":       True,
        })
        logs.append(("THỪA", f"  {vname}  start={cursor/1e6:.2f}s  speed={spd:.3f}x (giữ nguyên)"))
        cursor += tgt_dur

    stats = {
        "video_count": len(video_segs_sorted),
        "audio_count": len(audio_segs),
        "matched":     n,
        "leftover":    len(video_segs_sorted) - n,
    }
    return mapping, logs, stats


def apply_sync(draft_data, mapping, video_track):
    """
    Áp dụng mapping:
    - target_timerange: start mới, duration = audio_dur
    - source_timerange: giữ nguyên (footage gốc không đổi)
    - speed: src_dur / audio_dur (tự động tăng/giảm tốc)
    - Cập nhật speed material trong materials.speeds
    - Cập nhật duration tổng project
    """
    import copy
    new_draft = copy.deepcopy(draft_data)

    # Tìm video track bằng id
    orig_track_id = video_track.get("id", "")
    new_vtrack = None
    for track in new_draft.get("tracks", []):
        if track.get("type") == "video":
            if orig_track_id and track.get("id") == orig_track_id:
                new_vtrack = track
                break
    if not new_vtrack:
        for track in new_draft.get("tracks", []):
            if track.get("type") == "video":
                if new_vtrack is None or len(track.get("segments",[])) > len(new_vtrack.get("segments",[])):
                    new_vtrack = track
    if not new_vtrack:
        return None

    seg_by_id = {s.get("id",""): s for s in new_vtrack["segments"]}

    # Build lookup: speed material id → object (để cập nhật speed value)
    speeds_by_id = {}
    for spd in new_draft.get("materials",{}).get("speeds",[]):
        speeds_by_id[spd.get("id","")] = spd

    new_segments = []
    for m in mapping:
        seg_id = m.get("segment_id","")
        seg    = copy.deepcopy(seg_by_id.get(seg_id) or m["video_seg"])

        footage_dur = m["src_dur"]

        # Lưu thứ tự cố định vào segment (dùng cho lần chạy sau)
        seg["_sync_order"] = m["index"]

        # Lưu footage duration gốc (chỉ lần đầu)
        if not seg.get("_sync_orig_dur"):
            seg["_sync_orig_dur"] = footage_dur

        # 2. target_timerange: vị trí và độ dài mới trên timeline
        if "target_timerange" not in seg:
            seg["target_timerange"] = {}
        seg["target_timerange"]["start"]    = m["new_start"]
        seg["target_timerange"]["duration"] = m["new_target_dur"]

        # 3. source_timerange: luôn = footage gốc (đọc từ _sync_orig_dur)
        orig_footage = seg["_sync_orig_dur"]
        seg["source_timerange"] = {
            "start":    0,
            "duration": orig_footage,
        }

        # 4. speed
        seg["speed"] = m["new_speed"]

        # 5. Cập nhật speed material trong extra_material_refs
        for ref_id in seg.get("extra_material_refs", []):
            if ref_id in speeds_by_id:
                speeds_by_id[ref_id]["speed"] = m["new_speed"]
                speeds_by_id[ref_id]["mode"]  = 0
                break

        new_segments.append(seg)

    new_vtrack["segments"] = new_segments

    # Cập nhật duration tổng project
    if mapping:
        last = mapping[-1]
        new_draft["duration"] = last["new_start"] + last["new_target_dur"]

    return new_draft


# ══════════════════════════════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CapCut Timeline Sync Tool")
        self.geometry("920x700")
        self.minsize(800, 580)
        self.configure(bg=BG)

        self.draft_data   = None
        self.mapping      = None
        self.video_track  = None
        self.selected_proj = None

        self._build_ui()
        self._scan_projects()

    # ── UI Builder ─────────────────────────────────────────────
    def _build_ui(self):
        self._style()

        # Header
        hdr = tk.Frame(self, bg=SURFACE, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🎬  CapCut Timeline Sync",
                 font=("Segoe UI", 16, "bold"), fg=ACCENT, bg=SURFACE).pack(side="left", padx=20)
        tk.Label(hdr, text="sắp xếp clip ghép khớp theo thứ tự audio",
                 font=("Segoe UI", 9), fg=MUTED, bg=SURFACE).pack(side="left")

        # Workflow banner
        workflow = tk.Frame(self, bg="#1a1a10", pady=7)
        workflow.pack(fill="x")
        tk.Label(workflow,
                 text="📋 Quy trình đúng:  "
                      "① TẮT ĐỒNG BỘ CLOUD trong CapCut (Settings → tắt Sync)  →  "
                      "② Lưu (Ctrl+S) rồi ĐÓNG project về trang chủ  →  "
                      "③ Mở tool → Phân tích → Áp dụng  →  "
                      "④ Mở lại project trong CapCut",
                 font=("Segoe UI", 8), fg=GOLD, bg="#1a1a10").pack(side="left", padx=16)

        # Main pane
        pane = tk.PanedWindow(self, orient="horizontal", bg=BG, sashwidth=4, sashrelief="flat")
        pane.pack(fill="both", expand=True, padx=12, pady=12)

        # ── Left: project list ──────────────────────────────────
        left = tk.Frame(pane, bg=BG)
        pane.add(left, minsize=260)

        tk.Label(left, text="PROJECTS", font=("Segoe UI", 8, "bold"),
                 fg=MUTED, bg=BG).pack(anchor="w", pady=(0, 6))

        list_frame = tk.Frame(left, bg=SURFACE, bd=0, highlightthickness=1,
                              highlightbackground=BORDER)
        list_frame.pack(fill="both", expand=True)

        sb = tk.Scrollbar(list_frame, bg=SURFACE2, troughcolor=BG)
        sb.pack(side="right", fill="y")

        self.proj_list = tk.Listbox(
            list_frame, bg=SURFACE, fg=TEXT, selectbackground=ACCENT,
            selectforeground="white", font=("Segoe UI", 9),
            bd=0, highlightthickness=0, activestyle="none",
            yscrollcommand=sb.set, cursor="hand2"
        )
        self.proj_list.pack(fill="both", expand=True)
        sb.config(command=self.proj_list.yview)
        self.proj_list.bind("<<ListboxSelect>>", self._on_select_proj)

        btn_scan = tk.Button(left, text="↺  Quét lại", command=self._scan_projects,
                             bg=SURFACE2, fg=MUTED, font=("Segoe UI", 8),
                             bd=0, padx=8, pady=6, cursor="hand2",
                             activebackground=BORDER, activeforeground=TEXT)
        btn_scan.pack(fill="x", pady=(6, 0))

        # ── Right: main panel ───────────────────────────────────
        right = tk.Frame(pane, bg=BG)
        pane.add(right, minsize=500)

        # Stats row
        stats_row = tk.Frame(right, bg=BG)
        stats_row.pack(fill="x", pady=(0, 10))

        self.stat_vars = {}
        for key, label, color in [
            ("video",   "Video clips", ACCENT),
            ("audio",   "Audio lines", GOLD),
            ("matched", "Khớp",        SUCCESS),
            ("leftover","Thừa",        ACCENT2),
        ]:
            card = tk.Frame(stats_row, bg=SURFACE, bd=0, highlightthickness=1,
                            highlightbackground=BORDER)
            card.pack(side="left", fill="x", expand=True, padx=(0, 6))
            tk.Label(card, text=label.upper(), font=("Segoe UI", 7, "bold"),
                     fg=MUTED, bg=SURFACE).pack(anchor="w", padx=10, pady=(8, 0))
            v = tk.StringVar(value="—")
            self.stat_vars[key] = v
            tk.Label(card, textvariable=v, font=("Segoe UI", 20, "bold"),
                     fg=color, bg=SURFACE).pack(anchor="w", padx=10, pady=(0, 8))

        # Mapping preview
        hdr_row = tk.Frame(right, bg=BG)
        hdr_row.pack(fill="x", pady=(0,4))
        tk.Label(hdr_row, text="MAPPING  VIDEO ↔ AUDIO", font=("Segoe UI", 8, "bold"),
                 fg=MUTED, bg=BG).pack(side="left")
        tk.Label(hdr_row, text="(chọn video clip rồi nhấn ↑↓ để đổi thứ tự)",
                 font=("Segoe UI", 7), fg=MUTED, bg=BG).pack(side="left", padx=8)

        map_outer = tk.Frame(right, bg=SURFACE, bd=0, highlightthickness=1,
                             highlightbackground=BORDER)
        map_outer.pack(fill="both", expand=True)

        # Nút ↑↓ bên phải
        btn_move = tk.Frame(map_outer, bg=SURFACE, width=28)
        btn_move.pack(side="right", fill="y", padx=(0,2))
        btn_move.pack_propagate(False)
        tk.Button(btn_move, text="↑", command=self._move_up,
                  bg=SURFACE2, fg=TEXT, font=("Segoe UI", 10, "bold"),
                  bd=0, cursor="hand2", activebackground=ACCENT,
                  activeforeground="white").pack(fill="x", pady=(8,2))
        tk.Button(btn_move, text="↓", command=self._move_down,
                  bg=SURFACE2, fg=TEXT, font=("Segoe UI", 10, "bold"),
                  bd=0, cursor="hand2", activebackground=ACCENT,
                  activeforeground="white").pack(fill="x", pady=(0,2))
        tk.Button(btn_move, text="⟳", command=self._reset_order,
                  bg=SURFACE2, fg=MUTED, font=("Segoe UI", 9),
                  bd=0, cursor="hand2", activebackground=BORDER,
                  activeforeground=TEXT).pack(fill="x", pady=(12,2))

        # Treeview
        cols = ("pos", "audio", "arrow", "video", "speed", "status")
        self.tree = ttk.Treeview(map_outer, columns=cols, show="headings",
                                 selectmode="browse", height=12)
        self.tree.heading("pos",    text="#")
        self.tree.heading("audio",  text="Audio (cố định)")
        self.tree.heading("arrow",  text="")
        self.tree.heading("video",  text="Clip ghép → kéo thả đổi thứ tự")
        self.tree.heading("speed",  text="Speed")
        self.tree.heading("status", text="")
        self.tree.column("pos",    width=30,  anchor="center", stretch=False)
        self.tree.column("audio",  width=150, anchor="w")
        self.tree.column("arrow",  width=24,  anchor="center", stretch=False)
        self.tree.column("video",  width=170, anchor="w")
        self.tree.column("speed",  width=70,  anchor="center", stretch=False)
        self.tree.column("status", width=50,  anchor="center", stretch=False)

        self.tree.tag_configure("pair",     background=SURFACE,  foreground=TEXT)
        self.tree.tag_configure("leftover", background=SURFACE2, foreground=ACCENT2)
        self.tree.tag_configure("even",     background="#161620")
        self.tree.tag_configure("selected_vid", background=ACCENT, foreground="white")

        vsb = ttk.Scrollbar(map_outer, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Drag-drop trong treeview
        self.tree.bind("<ButtonPress-1>",   self._on_drag_start)
        self.tree.bind("<B1-Motion>",        self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>",  self._on_drag_release)
        self._drag_item = None

        # Log
        tk.Label(right, text="LOG", font=("Segoe UI", 8, "bold"),
                 fg=MUTED, bg=BG).pack(anchor="w", pady=(10, 4))

        self.log = scrolledtext.ScrolledText(
            right, height=7, bg=SURFACE, fg=MUTED,
            font=("Consolas", 8), bd=0, state="disabled",
            insertbackground=TEXT
        )
        self.log.pack(fill="x")

        # Tag colours in log
        self.log.tag_config("INFO",  foreground=ACCENT)
        self.log.tag_config("AUDIO", foreground=GOLD)
        self.log.tag_config("VIDEO", foreground=TEXT)
        self.log.tag_config("PAIR",  foreground=SUCCESS)
        self.log.tag_config("THỪA",  foreground=ACCENT2)
        self.log.tag_config("ERR",   foreground=ACCENT2)
        self.log.tag_config("OK",    foreground=SUCCESS)
        self.log.tag_config("---",   foreground=BORDER)

        # Buttons
        btn_row = tk.Frame(right, bg=BG)
        btn_row.pack(fill="x", pady=(10, 0))

        self.btn_analyze = tk.Button(
            btn_row, text="🔍  Phân tích", state="disabled",
            command=self._analyze,
            bg=ACCENT, fg="white", font=("Segoe UI", 10, "bold"),
            bd=0, padx=20, pady=10, cursor="hand2",
            activebackground="#7b8ff7", activeforeground="white"
        )
        self.btn_analyze.pack(side="left", padx=(0, 8))

        self.btn_apply = tk.Button(
            btn_row, text="✅  Áp dụng vào CapCut", state="disabled",
            command=self._apply,
            bg=SUCCESS, fg=BG, font=("Segoe UI", 10, "bold"),
            bd=0, padx=20, pady=10, cursor="hand2",
            activebackground="#42dab8", activeforeground=BG
        )
        self.btn_apply.pack(side="left")

        tk.Button(
            btn_row, text="📁", command=self._open_folder,
            bg=SURFACE2, fg=TEXT, font=("Segoe UI", 12),
            bd=0, padx=8, pady=8, cursor="hand2",
            activebackground=BORDER, activeforeground=TEXT
        ).pack(side="left", padx=(6,0))

        tk.Label(btn_row, text="(tự backup → ghi thẳng vào project)",
                 font=("Segoe UI", 8), fg=MUTED, bg=BG).pack(side="left", padx=6)

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background=SURFACE, foreground=TEXT,
                        fieldbackground=SURFACE, rowheight=26,
                        font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
                        background=SURFACE2, foreground=MUTED,
                        font=("Segoe UI", 8, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACCENT2)])
        style.configure("Vertical.TScrollbar",
                        background=SURFACE2, troughcolor=BG, relief="flat", width=8)

    # ── Project scanning ────────────────────────────────────────
    def _scan_projects(self):
        self.proj_list.delete(0, "end")
        self._projects = find_capcut_projects()
        if not self._projects:
            self.proj_list.insert("end", "  (Không tìm thấy project nào)")
            self._log("ERR", "Không tìm thấy project CapCut nào.")
            self._log("INFO", "Thử mở CapCut lên rồi quét lại.")
        else:
            for p in self._projects:
                mtime = datetime.fromtimestamp(p["mtime"]).strftime("%d/%m %H:%M")
                self.proj_list.insert("end", f"  {p['name']}  [{mtime}]")
            self._log("INFO", f"Tìm thấy {len(self._projects)} project (không trùng lặp)")

    def _on_select_proj(self, event):
        sel = self.proj_list.curselection()
        if not sel or not self._projects:
            return
        idx = sel[0]
        if idx >= len(self._projects):
            return
        proj = self._projects[idx]
        self.selected_proj = proj

        # Load JSON
        try:
            with open(proj["json"], "r", encoding="utf-8") as f:
                self.draft_data = json.load(f)
            self._log("INFO", f"Đã load: {proj['name']}")
            self._log("INFO", f"  → {proj['json']}")
            self.btn_analyze.config(state="normal")
            # Reset
            self.mapping = None
            self.btn_apply.config(state="disabled")
            self._clear_tree()
            for k in self.stat_vars:
                self.stat_vars[k].set("—")
        except Exception as e:
            self._log("ERR", f"Lỗi đọc JSON: {e}")
            self.draft_data = None

    # ── Analysis ────────────────────────────────────────────────
    def _analyze(self):
        if not self.selected_proj:
            return
        self._clear_tree()
        self._log_clear()

        # Luôn reload từ disk (quan trọng khi chạy lần 2+)
        try:
            with open(self.selected_proj["json"], "r", encoding="utf-8") as f:
                self.draft_data = json.load(f)
            self._log("INFO", "Reload file từ disk thanh cong")
        except Exception as e:
            self._log("ERR", f"Loi reload file: {e}")
            return

        mapping, logs, stats = analyze_and_sync(self.draft_data)

        for tag, msg in logs:
            if msg:
                self._log(tag, msg)

        if mapping is None:
            messagebox.showerror("Lỗi", "Không phân tích được. Xem log để biết chi tiết.")
            return

        # Update stats
        self.stat_vars["video"].set(str(stats["video_count"]))
        self.stat_vars["audio"].set(str(stats["audio_count"]))
        self.stat_vars["matched"].set(str(stats["matched"]))
        self.stat_vars["leftover"].set(str(stats["leftover"]))

        # Store mapping TRƯỚC khi gọi _refresh_tree
        tracks = self.draft_data.get("tracks", [])
        self.video_track = None
        for t in tracks:
            if t.get("type") == "video":
                if self.video_track is None or len(t.get("segments", [])) > len(self.video_track.get("segments", [])):
                    self.video_track = t

        self.mapping       = mapping
        self._orig_mapping = list(mapping)

        # Populate tree
        self._refresh_tree()

        self.btn_apply.config(state="normal")
        self._log("OK", f"✅ Phân tích xong! {stats['matched']} cặp khớp, {stats['leftover']} clip thừa.")
        self._log("INFO", "💡 Chọn dòng rồi nhấn ↑↓ để đổi thứ tự video nếu cần.")

    def _refresh_tree(self):
        """Vẽ lại bảng mapping từ self.mapping hiện tại."""
        if not self.mapping:
            return
        self._clear_tree()
        audio_segs_sorted = self._get_audio_segs_sorted()
        for i, m in enumerate(self.mapping):
            if not m["leftover"] and i < len(audio_segs_sorted):
                audio_dur = audio_segs_sorted[i]["seg"].get("target_timerange",{}).get("duration",0)
                src_dur   = m.get("src_dur", 0)
                speed     = round(src_dur / audio_dur, 3) if audio_dur > 0 and src_dur > 0 else 1.0
                speed     = max(0.1, min(100.0, speed))
                speed_txt = f"{speed:.3f}x"
            else:
                speed_txt = f"{m.get('new_speed', 1.0):.3f}x"

            tag    = "leftover" if m["leftover"] else ("pair" if i % 2 == 0 else "even")
            status = "🟡 thừa" if m["leftover"] else "✅"
            self.tree.insert("", "end", iid=str(i),
                values=(
                    i + 1,
                    m.get("audio_name") or "(không có audio)",
                    "→",
                    m.get("video_name", ""),
                    speed_txt,
                    status,
                ),
                tags=(tag,)
            )

    def _get_audio_segs_sorted(self):
        """Trả về audio segments sort theo start."""
        if not self.draft_data:
            return []
        tracks = self.draft_data.get("tracks", [])
        materials = self.draft_data.get("materials", {})
        audio_segs = []
        for track in tracks:
            if track.get("type") == "audio":
                for seg in track.get("segments", []):
                    name = get_material_name(seg.get("material_id",""), materials)
                    audio_segs.append({"seg": seg, "name": name})
        audio_segs.sort(key=lambda x: x["seg"].get("target_timerange",{}).get("start",0))
        return audio_segs

    def _get_selected_idx(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _move_up(self):
        idx = self._get_selected_idx()
        if idx is None or idx <= 0 or idx >= len(self.mapping):
            return
        self.mapping[idx], self.mapping[idx-1] = self.mapping[idx-1], self.mapping[idx]
        self._rebuild_mapping_audio()
        self._refresh_tree()
        self.tree.selection_set(str(idx-1))
        self.tree.see(str(idx-1))

    def _move_down(self):
        idx = self._get_selected_idx()
        if idx is None or idx < 0 or idx >= len(self.mapping) - 1:
            return
        self.mapping[idx], self.mapping[idx+1] = self.mapping[idx+1], self.mapping[idx]
        self._rebuild_mapping_audio()
        self._refresh_tree()
        self.tree.selection_set(str(idx+1))
        self.tree.see(str(idx+1))

    def _reset_order(self):
        """Reset về thứ tự gốc từ analyze."""
        self.mapping = list(self._orig_mapping)
        self._rebuild_mapping_audio()
        self._refresh_tree()
        self._log("INFO", "Đã reset về thứ tự gốc.")

    def _rebuild_mapping_audio(self):
        """Sau khi đổi thứ tự video, gán lại audio_name và tính lại speed/start."""
        audio_segs = self._get_audio_segs_sorted()
        cursor = 0
        for i, m in enumerate(self.mapping):
            if not m["leftover"] and i < len(audio_segs):
                aseg      = audio_segs[i]
                audio_dur = aseg["seg"].get("target_timerange",{}).get("duration",0)
                src_dur   = m["src_dur"]
                speed     = round(src_dur / audio_dur, 6) if audio_dur > 0 else 1.0
                speed     = max(0.1, min(100.0, speed))
                m["audio_name"]     = aseg["name"]
                m["new_start"]      = cursor
                m["new_target_dur"] = audio_dur
                m["new_speed"]      = speed
                cursor += audio_dur
            else:
                m["audio_name"] = None
                m["new_start"]  = cursor
                cursor += m["new_target_dur"]

    # ── Drag-drop ────────────────────────────────────────────────
    def _on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self._drag_item = item
            self.tree.selection_set(item)

    def _on_drag_motion(self, event):
        if not self._drag_item:
            return
        target = self.tree.identify_row(event.y)
        if target and target != self._drag_item:
            self.tree.selection_set(target)

    def _on_drag_release(self, event):
        if not self._drag_item:
            return
        target = self.tree.identify_row(event.y)
        if target and target != self._drag_item:
            try:
                from_idx = int(self._drag_item)
                to_idx   = int(target)
                # Di chuyển item từ from_idx đến to_idx
                item = self.mapping.pop(from_idx)
                self.mapping.insert(to_idx, item)
                self._rebuild_mapping_audio()
                self._refresh_tree()
                self.tree.selection_set(str(to_idx))
            except Exception:
                pass
        self._drag_item = None

    # ── Apply ───────────────────────────────────────────────────
    def _check_capcut_running(self):
        """Kiểm tra CapCut có đang chạy không."""
        import subprocess
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq CapCut.exe", "/NH"],
                capture_output=True, text=True, timeout=3
            )
            return "CapCut.exe" in result.stdout
        except Exception:
            return False

    def _kill_capcut(self):
        """Tắt toàn bộ tiến trình CapCut."""
        import subprocess
        import time
        # Tắt tất cả process liên quan đến CapCut
        capcut_processes = ["CapCut.exe", "CapCutService.exe", "CapCutUpdate.exe"]
        for proc in capcut_processes:
            try:
                subprocess.run(["taskkill", "/F", "/IM", proc],
                               capture_output=True, timeout=5)
            except Exception:
                pass
        # Chờ tắt hẳn (tối đa 10 giây)
        for _ in range(20):
            time.sleep(0.5)
            if not self._check_capcut_running():
                return True
        return False

    def _apply(self):
        if not self.mapping or not self.selected_proj:
            return

        # Kiểm tra CapCut có đang mở không
        capcut_open = self._check_capcut_running()
        if capcut_open:
            ans = messagebox.askyesno(
                "⚠️ CapCut đang mở!",
                "CapCut đang chạy — nếu bạn áp dụng ngay, CapCut sẽ ghi đè file khi bạn\n"
                "đóng project và mất toàn bộ thay đổi của tool!\n\n"
                "📋 Quy trình đúng:\n"
                "  1. Trong CapCut: lưu project (Ctrl+S)\n"
                "  2. ĐÓNG project (về trang chủ CapCut)\n"
                "  3. Quay lại tool → Phân tích lại → Áp dụng\n"
                "  4. Mở lại project trong CapCut\n\n"
                "Vẫn muốn tiếp tục (không khuyến nghị)?",
                icon="warning"
            )
            if not ans:
                return

        # Kiểm tra file đã được lưu chưa (so sánh mtime với lần load)
        json_path = self.selected_proj["json"]
        current_mtime = os.path.getmtime(json_path)
        loaded_mtime  = self.selected_proj.get("mtime", 0)
        if current_mtime > loaded_mtime + 2:  # file đã bị sửa sau khi tool load
            ans = messagebox.askyesno(
                "⚠️ File đã thay đổi!",
                f"File draft_content.json đã bị sửa đổi sau khi tool load\n"
                f"(có thể CapCut vừa auto-save).\n\n"
                f"Nên Phân tích lại trước khi Áp dụng.\n\n"
                f"Vẫn muốn áp dụng với data cũ?",
                icon="warning"
            )
            if not ans:
                self._log("WARN", "Hãy nhấn Phân tích lại trước!")
                return

        # Nếu CapCut đang chạy → hỏi tắt luôn
        if self._check_capcut_running():
            ans = messagebox.askyesno(
                "⚠️ CapCut đang chạy ngầm!",
                "CapCut đang chạy ngầm và sẽ ghi đè file ngay sau khi tool sửa!\n\n"
                "Tool cần TẮT CapCut trước khi ghi file.\n\n"
                "Cho phép tool tắt CapCut ngay bây giờ?",
                icon="warning"
            )
            if not ans:
                self._log("WARN", "Hủy — cần tắt CapCut trước khi áp dụng!")
                return
            self._log("INFO", "Đang tắt CapCut...")
            ok = self._kill_capcut()
            if ok:
                self._log("OK", "✅ Đã tắt CapCut thành công!")
            else:
                self._log("WARN", "⚠️ Không tắt được CapCut hoàn toàn, vẫn tiếp tục...")

        confirm = messagebox.askyesno(
            "Xác nhận",
            f"Sắp xếp lại {len(self.mapping)} clip trong project:\n\n"
            f"  {self.selected_proj['name']}\n\n"
            f"File gốc sẽ bị ghi đè (không tạo backup).\nTiếp tục?"
        )
        if not confirm:
            return

        try:
            proj_dir = os.path.dirname(json_path)

            # Apply
            new_draft = apply_sync(self.draft_data, self.mapping, self.video_track)
            if new_draft is None:
                messagebox.showerror("Lỗi", "Không tìm thấy video track khi áp dụng!")
                return

            # Ghi atomic: ghi ra .tmp rồi rename (tránh CapCut đọc file đang ghi dở)
            tmp_path = json_path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(new_draft, f, ensure_ascii=False)
                f.flush()
            os.replace(tmp_path, json_path)

            # Ghi đè nội dung mới vào TẤT CẢ file restore của CapCut
            # (thay vì xóa — để dù CapCut restore từ chúng cũng ra bản mới)
            new_content_bytes = json.dumps(new_draft, ensure_ascii=False).encode("utf-8")
            draft_fname = os.path.basename(json_path)
            overwritten = []
            deleted = []
            for fname in os.listdir(proj_dir):
                fpath = os.path.join(proj_dir, fname)
                if not os.path.isfile(fpath) or fname == draft_fname:
                    continue
                fl = fname.lower()
                is_restore_file = (
                    fl.endswith(".bak") or
                    fl.endswith(".backup") or
                    fl.endswith(".tmp") or
                    fl.startswith("draft_content.json.") or
                    ("backup" in fl and fl.endswith(".json"))
                )
                if is_restore_file:
                    try:
                        # Thử ghi đè nội dung mới vào file
                        with open(fpath, "w", encoding="utf-8") as rf:
                            json.dump(new_draft, rf, ensure_ascii=False)
                        overwritten.append(fname)
                        self._log("INFO", f"Ghi đè: {fname}")
                    except Exception:
                        # Nếu không ghi được thì xóa
                        try:
                            os.remove(fpath)
                            deleted.append(fname)
                            self._log("INFO", f"Xóa: {fname}")
                        except Exception as ex:
                            self._log("WARN", f"Không xử lý được {fname}: {ex}")
                elif fname == ".locked":
                    try:
                        os.remove(fpath)
                        deleted.append(fname)
                        self._log("INFO", f"Xóa .locked")
                    except Exception:
                        pass
            if overwritten or deleted:
                self._log("INFO", f"Đã xử lý {len(overwritten)+len(deleted)} file: ghi đè={overwritten}, xóa={deleted}")

            # Xác minh
            with open(json_path, "r", encoding="utf-8") as f:
                verify = json.load(f)
            vtrack = None
            for t in verify.get("tracks", []):
                if t.get("type") == "video":
                    if vtrack is None or len(t.get("segments",[])) > len(vtrack.get("segments",[])):
                        vtrack = t
            segs = vtrack.get("segments", []) if vtrack else []
            has_sync = any(s.get("_sync_order") is not None for s in segs)
            if has_sync:
                self._log("OK", f"✅ Xác minh OK! {len(segs)} segments")
            else:
                self._log("WARN", "⚠️ Không tìm thấy _sync_order — kiểm tra lại!")

            self.selected_proj["mtime"] = os.path.getmtime(json_path)
            self._log("OK", "✅ Hoàn tất! Mở lại project trong CapCut để thấy kết quả.")

            messagebox.showinfo(
                "Thành công! 🎉",
                "Đã ghi draft_content.json mới thành công!\n\n"
                "👉 Bây giờ mở lại project trong CapCut.\n\n"
                "⚠️ Nếu CapCut vẫn hiện bản cũ:\n"
                "   • Tắt đồng bộ cloud trong Settings của CapCut\n"
                "   • Hoặc tắt hoàn toàn CapCut rồi mở lại"
            )

        except Exception as e:
            self._log("ERR", f"Lỗi khi ghi file: {e}")
            messagebox.showerror("Lỗi", str(e))

    def _open_folder(self):
        """Mở thư mục chứa file project trong Explorer."""
        if self.selected_proj:
            import subprocess
            subprocess.Popen(f'explorer /select,"{self.selected_proj["json"]}"')
        else:
            messagebox.showinfo("Thông báo", "Chưa chọn project nào.")

    # ── Helpers ─────────────────────────────────────────────────
    def _clear_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

    def _log_clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _log(self, tag, msg):
        self.log.config(state="normal")
        if tag == "---":
            self.log.insert("end", "─" * 60 + "\n", "---")
        else:
            line = f"[{tag}]  {msg}\n"
            self.log.insert("end", line, tag)
        self.log.see("end")
        self.log.config(state="disabled")


# ── Entry point ────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
