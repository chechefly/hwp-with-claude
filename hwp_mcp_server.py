#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hwp_mcp — 한글(HWP) 파일을 Claude가 직접 읽고 수정할 수 있게 해주는 MCP 서버.

Windows + 한/글(Hangul Office) 설치 환경에서만 동작합니다.
한/글 자동화(COM) 엔진을 '숨김 모드'로 하나만 띄워 재사용하며,
COM 스레드 친화성(STA)을 위해 모든 한글 작업을 전용 워커 스레드에서 실행합니다.
"""

import atexit
import contextlib
import datetime
import os
import queue
import shutil
import sys
import threading
import time
from typing import Optional, Dict, Any, List

import pythoncom
import win32com.client as win32
import win32com.client.dynamic as win32dyn  # 순수 동적 디스패치(gencache/stdout 출력 방지)
try:
    import win32clipboard
except Exception:  # pragma: no cover
    win32clipboard = None
from pydantic import BaseModel, Field, ConfigDict
from mcp.server.fastmcp import FastMCP, Image

__version__ = "0.6.0"
# 업데이트 확인용. GitHub 저장소 만든 뒤 아래 CHANGE_ME를 본인 GitHub 사용자명으로 바꾸세요.
_UPDATE_URL = "https://raw.githubusercontent.com/chechefly/hwp-with-claude/main/version.json"

mcp = FastMCP("hwp_mcp")


_SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
_SEC_DLL = os.path.join(_SERVER_DIR, "FilePathCheckerModule.dll")
# 표시 모드: 환경변수 HWP_MCP_VISIBLE=1 이면 한/글 창을 띄운 채로 작업(실시간 확인용)
_VISIBLE = os.environ.get("HWP_MCP_VISIBLE", "0") == "1"


def _ensure_security_module() -> None:
    """한/글 자동화 보안모듈을 레지스트리에 등록(파일접근 승인창 제거). 이미 있으면 통과.
    관리자 권한/ regsvr32 불필요 — HKCU에 DLL 경로만 기록."""
    if not os.path.exists(_SEC_DLL):
        return
    try:
        import winreg
        key_path = r"Software\HNC\HwpAutomation\Modules"
        reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        # 이미 올바르게 등록돼 있으면 건드리지 않음
        try:
            k = winreg.OpenKey(reg, key_path, 0, winreg.KEY_READ)
            val, _ = winreg.QueryValueEx(k, "FilePathCheckerModule")
            winreg.CloseKey(k)
            if val and os.path.exists(val):
                return
        except FileNotFoundError:
            pass
        k = winreg.CreateKey(reg, key_path)
        winreg.SetValueEx(k, "FilePathCheckerModule", 0, winreg.REG_SZ, _SEC_DLL)
        winreg.CloseKey(k)
    except Exception:
        pass


def _backup_file(path: str, tag: str = "save") -> Optional[str]:
    """대상 파일을 같은 폴더의 .hwp_backups/ 에 타임스탬프 사본으로 보관(가역성 보장).
    존재하지 않으면 None. 실패해도 예외 대신 None(작업 자체를 막지 않되 로깅 목적)."""
    try:
        if not path or not os.path.exists(path):
            return None
        d = os.path.join(os.path.dirname(os.path.abspath(path)), ".hwp_backups")
        os.makedirs(d, exist_ok=True)
        base = os.path.basename(path)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(d, f"{base}.{tag}_{stamp}.bak")
        # 동일 타임스탬프 충돌 방지
        n = 1
        while os.path.exists(dst):
            dst = os.path.join(d, f"{base}.{tag}_{stamp}_{n}.bak")
            n += 1
        shutil.copy2(path, dst)
        return dst
    except Exception:
        return None


def _list_backups(path: str) -> List[str]:
    """대상 파일의 백업 목록(최신순)."""
    d = os.path.join(os.path.dirname(os.path.abspath(path)), ".hwp_backups")
    if not os.path.isdir(d):
        return []
    base = os.path.basename(path)
    items = [os.path.join(d, f) for f in os.listdir(d) if f.startswith(base + ".")]
    items.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return items


def _clear_clipboard() -> None:
    """클립보드를 비운다. 셀 Copy가 빈 셀에서 아무것도 못 담으면 '이전 클립보드 내용'이
    그대로 읽혀 유출되므로(예: shell 명령·개인정보), 읽기 전에 반드시 비운다."""
    if win32clipboard is None:
        return
    for _ in range(5):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            return
        except Exception:
            time.sleep(0.03)


def _clip_text() -> str:
    if win32clipboard is None:
        return ""
    for _ in range(8):
        try:
            win32clipboard.OpenClipboard()
            try:
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            except Exception:
                data = ""
            win32clipboard.CloseClipboard()
            return data or ""
        except Exception:
            time.sleep(0.05)
    return ""

# ---------------------------------------------------------------------------
# HWP 워커: 전용 STA 스레드에서 HwpObject 하나를 관리
# ---------------------------------------------------------------------------

class HwpWorker:
    """모든 한/글 COM 호출을 단일 STA 스레드로 직렬화하는 워커."""

    def __init__(self) -> None:
        self._tasks: "queue.Queue" = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._hwp = None
        self._current_path: Optional[str] = None

    # 내부: 스레드 루프
    def _run(self) -> None:
        pythoncom.CoInitialize()
        try:
            while True:
                fn, result_box, done = self._tasks.get()
                if fn is None:  # 종료 신호
                    break
                try:
                    result_box["value"] = fn()
                except Exception as e:  # noqa: BLE001
                    result_box["error"] = e
                finally:
                    done.set()
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn):
        """워커 스레드에서 fn()을 실행하고 결과를 반환(예외는 재발생)."""
        result_box: Dict[str, Any] = {}
        done = threading.Event()
        self._tasks.put((fn, result_box, done))
        done.wait()
        if "error" in result_box:
            raise result_box["error"]
        return result_box.get("value")

    # 내부: 엔진 확보 (없으면 생성)
    def _ensure_engine(self):
        if self._hwp is None:
            _ensure_security_module()  # 보안 승인창 제거(레지스트리 보장)
            # win32com이 stdout에 찍을 수 있는 메시지가 MCP(JSON-RPC) 스트림을 오염시키지
            # 않도록, 엔진 초기화 동안 stdout을 stderr로 우회.
            with contextlib.redirect_stdout(sys.stderr):
                # 순수 동적 디스패치: gencache 자동생성/stdout 출력 방지
                hwp = win32dyn.Dispatch("HWPFrame.HwpObject")
                try:
                    hwp.SetMessageBoxMode(0x00000020)
                except Exception:
                    pass
                # 보안모듈 등록(레지스트리에 DLL 등록돼 있으면 True → 승인창 안뜸)
                for mod in ("FilePathCheckerModule", "FilePathCheckerModuleExample"):
                    try:
                        if hwp.RegisterModule("FilePathCheckDLL", mod):
                            break
                    except Exception:
                        pass
            self._hwp = hwp
            self._set_visible(_VISIBLE)
        return self._hwp

    def _set_visible(self, visible: bool) -> None:
        try:
            self._hwp.XHwpWindows.Item(0).Visible = visible
        except Exception:
            pass

    @property
    def current_path(self) -> Optional[str]:
        return self._current_path

    # ----- 실제 동작들 (모두 워커 스레드에서 호출됨) -----

    def op_open(self, path: str) -> str:
        hwp = self._ensure_engine()
        if not hwp.Open(path, "", "forceopen:true"):
            raise RuntimeError(f"파일을 열지 못했습니다: {path}")
        self._set_visible(_VISIBLE)
        self._current_path = path
        return hwp.GetTextFile("TEXT", "")

    def op_new(self) -> None:
        hwp = self._ensure_engine()
        hwp.Run("FileNew")
        self._set_visible(_VISIBLE)
        self._current_path = None

    def op_read(self) -> str:
        hwp = self._require_doc()
        return hwp.GetTextFile("TEXT", "")

    def op_replace(self, find: str, replace: str) -> int:
        hwp = self._require_doc()
        before = hwp.GetTextFile("TEXT", "")
        act = hwp.CreateAction("AllReplace")
        pset = act.CreateSet()
        pset.SetItem("FindString", find)
        pset.SetItem("ReplaceString", replace)
        pset.SetItem("IgnoreMessage", 1)
        pset.SetItem("Direction", 0)
        pset.SetItem("WholeWordOnly", 0)
        pset.SetItem("UseWildCards", 0)
        pset.SetItem("MatchCase", 1)
        pset.SetItem("ReplaceMode", 1)
        act.Execute(pset)
        after = hwp.GetTextFile("TEXT", "")
        # 대략적인 치환 횟수(원문 기준 등장 횟수)
        return before.count(find)

    def op_insert(self, text: str, at_end: bool) -> None:
        hwp = self._require_doc()
        if at_end:
            hwp.MovePos(3, 0, 0)  # 문서 맨 끝으로 이동
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        pset.SetItem("Text", text)
        act.Execute(pset)

    def op_list_fields(self) -> list:
        hwp = self._require_doc()
        raw = hwp.GetFieldList(0, 0) or ""
        names = [n for n in raw.split("\x02") if n]
        # 중복 인덱스 표기(name{{0}}) 정리
        seen = []
        for n in names:
            base = n.split("{{")[0]
            if base not in seen:
                seen.append(base)
        return seen

    def op_fill_fields(self, values: Dict[str, str]) -> Dict[str, bool]:
        hwp = self._require_doc()
        result = {}
        for name, val in values.items():
            try:
                hwp.PutFieldText(name, val)
                result[name] = True
            except Exception:
                result[name] = False
        return result

    def op_save(self) -> str:
        hwp = self._require_doc()
        if not self._current_path:
            raise RuntimeError("현재 문서에 경로가 없습니다. save_as를 사용하세요.")
        # Save(save_if_dirty) — 인자 필요. 안전하게 SaveAs로 현재 경로에 저장(포맷 자동판별).
        fmt = "HWPX" if self._current_path.lower().endswith(".hwpx") else "HWP"
        if not hwp.SaveAs(self._current_path, fmt, ""):
            raise RuntimeError(f"저장 실패: {self._current_path}")
        return self._current_path

    def op_save_as(self, path: str, fmt: str) -> str:
        hwp = self._require_doc()
        if not hwp.SaveAs(path, fmt, ""):
            raise RuntimeError(f"저장 실패: {path}")
        if fmt.upper() == "HWP":
            self._current_path = path
        return path

    def op_export_pdf(self, path: str) -> str:
        hwp = self._require_doc()
        if not hwp.SaveAs(path, "PDF", ""):
            raise RuntimeError("PDF 내보내기 실패")
        return path

    def op_close_doc(self) -> None:
        """현재 문서만 닫고 엔진은 유지(다음 파일을 위해 재사용)."""
        if self._hwp is not None:
            try:
                self._hwp.Run("FileClose")
            except Exception:
                pass
            self._current_path = None

    def op_quit(self) -> None:
        """엔진을 완전히 종료(백그라운드 프로세스 정리). 서버 종료 시 호출."""
        if self._hwp is not None:
            try:
                self._hwp.Run("FileClose")
            except Exception:
                pass
            try:
                self._hwp.Quit()
            except Exception:
                pass
            self._hwp = None
            self._current_path = None

    def op_show(self, show: bool) -> str:
        """현재 한/글 창을 표시/숨김. 실시간으로 눈으로 확인하고 싶을 때 사용."""
        hwp = self._ensure_engine()
        self._set_visible(show)
        try:
            if show:
                hwp.XHwpWindows.Item(0).Visible = True
        except Exception:
            pass
        return "표시" if show else "숨김"

    # ---------- 표(table) ----------
    def _cell_addr(self):
        ki = self._hwp.KeyIndicator()
        s = ki[8] if len(ki) > 8 else "?"
        return s.split(")")[0].lstrip("(") if ")" in s else s

    def _goto_table(self, index):
        """index(1-base)번째 표에 진입, 캐럿을 A1에 둔다."""
        hwp = self._require_doc()
        ctrl = hwp.HeadCtrl
        ti = 0
        target = None
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                ti += 1
                if ti == index:
                    target = ctrl
                    break
            ctrl = ctrl.Next
        if target is None:
            raise RuntimeError(f"표 #{index} 를 찾을 수 없습니다.")
        hwp.SetPosBySet(target.GetAnchorPos(0))
        hwp.FindCtrl()
        hwp.HAction.Run("ShapeObjTableSelCell")
        hwp.HAction.Run("Cancel")

    def _read_current_cell(self):
        hwp = self._hwp
        _clear_clipboard()               # 스테일 클립보드 유출 방지(빈 셀 → "" 보장)
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("Copy")
        t = _clip_text()
        hwp.HAction.Run("Cancel")
        return t

    def _count_tables(self):
        hwp = self._require_doc()
        ctrl = hwp.HeadCtrl
        n = 0
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                n += 1
            ctrl = ctrl.Next
        return n

    def op_table_map(self, index):
        """표의 {셀주소: 텍스트} 지도. 병합셀 재방문은 첫 값 유지."""
        self._goto_table(index)
        result = {}
        order = []
        guard = 0
        while guard < 800:
            guard += 1
            a = self._cell_addr()
            if a not in result:
                result[a] = self._read_current_cell().strip()
                order.append(a)
            if not self._hwp.HAction.Run("TableRightCell"):
                break
        self._hwp.HAction.Run("Cancel")
        return [{"addr": a, "text": result[a]} for a in order]

    def op_set_cells(self, index, mapping, align=None):
        """{addr: text} 여러 셀을 한 번의 표 진입으로 채운다. 반환: 못 채운 주소 목록."""
        hwp = self._require_doc()
        self._goto_table(index)
        remaining = dict(mapping)
        guard = 0
        align_act = {"left": "ParagraphShapeAlignLeft",
                     "center": "ParagraphShapeAlignCenter",
                     "right": "ParagraphShapeAlignRight"}.get(align)
        while remaining and guard < 1200:
            guard += 1
            a = self._cell_addr()
            if a in remaining:
                text = remaining.pop(a)
                hwp.HAction.Run("SelectAll")   # 셀 내부 전체 선택
                hwp.HAction.Run("Delete")
                if align_act:
                    hwp.HAction.Run(align_act)
                if text:
                    ac = hwp.CreateAction("InsertText")
                    ps = ac.CreateSet()
                    ps.SetItem("Text", text)
                    ac.Execute(ps)
                hwp.HAction.Run("Cancel")
            if not hwp.HAction.Run("TableRightCell"):
                break
        hwp.HAction.Run("Cancel")
        return list(remaining.keys())

    # ---------- 체크박스 ----------
    def op_check(self, labels, mark="☑"):
        """각 label에 대해 '□'+label 을 mark+label 로 치환. 반환: {label: 건수}."""
        hwp = self._require_doc()
        out = {}
        for label in labels:
            find = "□" + label
            repl = mark + label
            before = hwp.GetTextFile("TEXT", "")
            act = hwp.CreateAction("AllReplace")
            p = act.CreateSet()
            p.SetItem("FindString", find)
            p.SetItem("ReplaceString", repl)
            p.SetItem("IgnoreMessage", 1)
            p.SetItem("Direction", 0)
            p.SetItem("ReplaceMode", 1)
            p.SetItem("MatchCase", 1)
            act.Execute(p)
            out[label] = before.count(find)
        return out

    def op_check_after(self, keyword, mark="☑", position="after"):
        """keyword(예:'동의함')를 찾아, 그 근처의 특수 네모(☐)를 지우고 mark 삽입.
        position: 'after'(단어 뒤 네모, 예: '동의함 ☐') | 'before'(단어 앞 네모, 예: '☐ 동의함').
        문서 끝→처음 되돌이(wrap)를 감지해 정지. 반환: 처리한 개수."""
        hwp = self._require_doc()
        hwp.HAction.Run("MoveDocBegin")
        pset = hwp.HParameterSet.HFindReplace
        seen = set()
        n = 0
        for _ in range(60):
            hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = keyword
            pset.IgnoreMessage = 1
            pset.Direction = 0
            if not hwp.HAction.Execute("RepeatFind", pset.HSet):
                break
            pos = tuple(hwp.GetPos())
            if pos in seen:
                break
            seen.add(pos)
            if position == "before":
                # '☐ 동의함' : keyword 앞의 네모를 선택
                hwp.HAction.Run("Cancel")
                hwp.HAction.Run("MoveLeft")       # 선택 해제 → keyword 앞
                hwp.HAction.Run("MoveLeft")       # 공백 건너
                hwp.HAction.Run("MoveSelRight")   # 박스 1글자 선택
            else:
                hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("MoveRight")      # keyword 뒤 공백 건너 박스 앞
                hwp.HAction.Run("MoveSelRight")   # 박스 1글자 선택
            ac = hwp.CreateAction("InsertText")
            ps = ac.CreateSet()
            ps.SetItem("Text", mark)
            ac.Execute(ps)
            n += 1
        return n

    def op_check_by_label(self, labels, mark="☑"):
        """옵션 라벨(예:'동의함','운전자금')의 인접 네모를 체크(범용).
        □(U+25A1)은 AllReplace로, ☐(U+2610)은 Find-이동으로 처리 → 두 종류 모두 커버.
        박스가 라벨 앞이든 뒤든 자동 시도. 반환: {라벨: 결과}."""
        hwp = self._require_doc()
        out = {}
        for label in labels:
            # 1) □(U+25A1): 라벨 앞/뒤 4패턴 AllReplace
            did = 0
            for find, repl in [("□" + label, mark + label),
                               ("□ " + label, mark + " " + label),
                               (label + "□", label + mark),
                               (label + " □", label + " " + mark)]:
                before = hwp.GetTextFile("TEXT", "")
                cnt = before.count(find)
                if cnt:
                    act = hwp.CreateAction("AllReplace")
                    p = act.CreateSet()
                    p.SetItem("FindString", find)
                    p.SetItem("ReplaceString", repl)
                    p.SetItem("IgnoreMessage", 1)
                    p.SetItem("Direction", 0)
                    p.SetItem("ReplaceMode", 1)
                    p.SetItem("MatchCase", 1)
                    act.Execute(p)
                    did += cnt
                    break
            if did:
                out[label] = f"OK(□×{did})"
                continue
            # 2) ☐(U+2610): 라벨 뒤 → 앞 순으로 Find-이동
            n = self.op_check_after(label, mark, "after")
            if not n:
                n = self.op_check_after(label, mark, "before")
            out[label] = f"OK(☐×{n})" if n else "박스 못 찾음"
        return out

    def op_fill_by_label(self, fills, direction="auto", nth=1):
        """라벨 텍스트를 Find로 찾아 인접 칸/같은 칸에 값을 채운다(read_text가 보는 라벨을 앵커로).
        table_map 셀읽기가 실패하는 폼에도 동작.
        - direction: 'auto'(오른쪽→아래→인라인 자동 시도, 권장) | 'right' | 'below' | 'inline'
        - fills 값은 문자열이거나 {value, direction, nth, near} 딕셔너리(라벨별 세밀 제어).
          near: 이 텍스트를 먼저 찾은 뒤 그 다음의 라벨을 매칭(중복 라벨 구분용).
        - 공백 유연 매칭(라벨 공백 변형도 시도). 반환: {라벨: 'OK(방향)' | 실패이유}."""
        hwp = self._require_doc()

        def _insert(text):
            ac = hwp.CreateAction("InsertText")
            ps = ac.CreateSet()
            ps.SetItem("Text", text)
            ac.Execute(ps)

        def _do_find(s):
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = s
            pset.IgnoreMessage = 1
            pset.Direction = 0
            return hwp.HAction.Execute("RepeatFind", pset.HSet)

        def _find(label, nth_, near):
            # 공백 유연: 원본 → 공백제거 → 단일공백 정규화 순으로 시도
            variants = list(dict.fromkeys([label, label.replace(" ", ""), " ".join(label.split())]))
            for v in variants:
                hwp.HAction.Run("MoveDocBegin")
                if near and not _do_find(near):
                    continue
                ok = False
                for _ in range(max(1, nth_)):
                    ok = _do_find(v)
                    if not ok:
                        break
                if ok:
                    return True
            return False

        def _fill_dir(d, value):
            # 호출 시 라벨이 방금 Find로 선택된 상태라고 가정
            if d == "inline":
                hwp.HAction.Run("MoveRight")   # 선택 해제(라벨 끝) → 뒤에 이어붙임
                _insert(value)
                return True
            hwp.HAction.Run("Cancel")
            mv = {"right": "TableRightCell", "below": "TableLowerCell"}.get(d)
            if not mv or not hwp.HAction.Run(mv):
                return False
            hwp.HAction.Run("SelectAll")       # 값 칸 내용 비우기
            hwp.HAction.Run("Delete")
            _insert(value)
            return True

        out = {}
        for label, spec in fills.items():
            if isinstance(spec, dict):
                value = str(spec.get("value", ""))
                d = spec.get("direction", direction)
                nth_ = int(spec.get("nth", nth))
                near = spec.get("near")
            else:
                value, d, nth_, near = str(spec), direction, nth, None
            dirs = ["right", "below", "inline"] if d == "auto" else [d]
            result = "라벨 못 찾음"
            for dd in dirs:
                if not _find(label, nth_, near):
                    result = "라벨 못 찾음"
                    break  # 못 찾으면 방향 바꿔도 소용없음
                if _fill_dir(dd, value):
                    result = f"OK({dd})"
                    break
                result = "인접 칸 없음(표 아님?)"
            out[label] = result
        return out

    def _require_doc(self):
        if self._hwp is None:
            raise RuntimeError("열려 있는 문서가 없습니다. 먼저 hwp_open 또는 hwp_new를 호출하세요.")
        return self._hwp


_worker = HwpWorker()


@atexit.register
def _cleanup_engine() -> None:
    """서버 종료 시 숨겨진 한/글 엔진을 정리(백그라운드 프로세스 잔류 방지)."""
    try:
        _worker.submit(_worker.op_quit)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 입력 모델
# ---------------------------------------------------------------------------

class OpenInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    path: str = Field(..., description="열 .hwp/.hwpx 파일의 절대 경로 (예: 'C:\\\\docs\\\\form.hwp')")
    make_backup: bool = Field(default=True, description="열 때 원본을 .hwp_backups/에 자동 백업(가역성). 기본 True")


class ReplaceInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    find: str = Field(..., description="찾을 문자열 (예: '{{name}}')", min_length=1)
    replace: str = Field(..., description="바꿀 문자열 (예: '홍길동'). 빈 문자열이면 삭제.")


class FillInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    replacements: Dict[str, str] = Field(
        ..., description="여러 자리표시자를 한 번에 치환. 예: {'{{name}}':'홍길동', '{{date}}':'2026-07-24'}"
    )


class InsertInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    text: str = Field(..., description="삽입할 텍스트. 줄바꿈은 \\r\\n 사용.", min_length=1)
    at_end: bool = Field(default=True, description="True면 문서 맨 끝에, False면 현재 커서 위치에 삽입")


class FillFieldsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    values: Dict[str, str] = Field(
        ..., description="누름틀(필드) 이름→값. 예: {'주소':'서울시...', '성명':'홍길동'}"
    )


class SaveAsInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    path: str = Field(..., description="저장할 절대 경로")
    format: str = Field(default="HWP", description="저장 형식: 'HWP', 'HWPX', 'PDF' 중 하나")


# ---------------------------------------------------------------------------
# 도구(tool)들
# ---------------------------------------------------------------------------

@mcp.tool(
    name="hwp_open",
    annotations={"title": "한글 파일 열기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_open(params: OpenInput) -> str:
    """지정한 .hwp/.hwpx 파일을 (숨김 모드로) 열고 전체 텍스트를 반환합니다.

    편집 작업(hwp_replace, hwp_fill, hwp_insert 등)을 하기 전에 먼저 호출하세요.
    반환값으로 현재 문서 내용을 확인해 어떤 자리표시자/필드를 채울지 판단할 수 있습니다.

    Returns:
        str: 열린 문서의 전체 텍스트. 실패 시 'Error: ...'
    """
    if not os.path.isabs(params.path):
        return "Error: 절대 경로를 입력하세요."
    if not os.path.exists(params.path):
        return f"Error: 파일이 존재하지 않습니다: {params.path}"
    try:
        bk = _backup_file(params.path, "open") if params.make_backup else None
        text = _worker.submit(lambda: _worker.op_open(params.path))
        head = f"[열림] {params.path}"
        if bk:
            head += f"\n[자동 백업] {bk}"
        return f"{head}\n\n--- 문서 텍스트 ---\n{text}"
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_new",
    annotations={"title": "새 한글 문서", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_new() -> str:
    """새 빈 한글 문서를 만듭니다. 저장하려면 hwp_save_as를 사용하세요."""
    try:
        _worker.submit(_worker.op_new)
        return "[새 문서 생성됨] 내용을 채운 뒤 hwp_save_as로 저장하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_read_text",
    annotations={"title": "현재 문서 텍스트 읽기", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_read_text() -> str:
    """현재 열려 있는 문서의 전체 텍스트를 반환합니다."""
    try:
        return _worker.submit(_worker.op_read)
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_replace",
    annotations={"title": "찾아 바꾸기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_replace(params: ReplaceInput) -> str:
    """문서 전체에서 문자열을 찾아 모두 바꿉니다(서식 유지).

    양식 채우기에 가장 많이 쓰입니다. 예: 템플릿에 '{{성명}}'을 넣어두고
    hwp_replace(find='{{성명}}', replace='홍길동').

    Returns:
        str: 치환 결과 요약. 실패 시 'Error: ...'
    """
    try:
        n = _worker.submit(lambda: _worker.op_replace(params.find, params.replace))
        if n == 0:
            return f"'{params.find}' 을(를) 찾지 못했습니다 (치환 0건)."
        return f"'{params.find}' → '{params.replace}' : 약 {n}건 치환 완료. (저장하려면 hwp_save)"
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_fill",
    annotations={"title": "양식 일괄 채우기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_fill(params: FillInput) -> str:
    """여러 자리표시자를 한 번에 치환합니다(서식 유지).

    서류 양식 자동 작성의 핵심 도구입니다.
    예: replacements={'{{성명}}':'홍길동','{{생년월일}}':'1990-01-01','{{금액}}':'1,000,000원'}

    Returns:
        str: 항목별 치환 건수 요약. 실패 시 'Error: ...'
    """
    try:
        lines = []
        for find, replace in params.replacements.items():
            n = _worker.submit(lambda f=find, r=replace: _worker.op_replace(f, r))
            mark = "✓" if n > 0 else "✗(못찾음)"
            lines.append(f"  {mark} '{find}' → '{replace}' ({n}건)")
        return "[일괄 채우기 결과]\n" + "\n".join(lines) + "\n\n저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_insert_text",
    annotations={"title": "텍스트 삽입", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_insert_text(params: InsertInput) -> str:
    """문서에 텍스트를 삽입합니다(기본: 문서 맨 끝)."""
    try:
        _worker.submit(lambda: _worker.op_insert(params.text, params.at_end))
        pos = "맨 끝" if params.at_end else "현재 커서 위치"
        return f"[{pos}에 삽입 완료] 저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_list_fields",
    annotations={"title": "누름틀 필드 목록", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_list_fields() -> str:
    """현재 문서에 정의된 누름틀(필드) 이름 목록을 반환합니다.

    관공서/회사 양식이 누름틀로 만들어진 경우, 이 목록을 보고 hwp_fill_fields로 채웁니다.
    """
    try:
        fields = _worker.submit(_worker.op_list_fields)
        if not fields:
            return "누름틀 필드가 없습니다. (자리표시자 방식이라면 hwp_read_text 후 hwp_fill 사용)"
        return "필드 목록:\n" + "\n".join(f"  - {f}" for f in fields)
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_fill_fields",
    annotations={"title": "누름틀 필드 채우기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_fill_fields(params: FillFieldsInput) -> str:
    """누름틀(필드) 이름별로 값을 채웁니다. 먼저 hwp_list_fields로 이름을 확인하세요."""
    try:
        res = _worker.submit(lambda: _worker.op_fill_fields(params.values))
        lines = [f"  {'✓' if ok else '✗'} {name}" for name, ok in res.items()]
        return "[필드 채우기 결과]\n" + "\n".join(lines) + "\n\n저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_save",
    annotations={"title": "저장", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_save() -> str:
    """현재 문서를 원래 파일에 덮어 저장합니다(덮어쓰기 전 자동 백업 → 가역적)."""
    try:
        bk = _backup_file(_worker.current_path or "", "presave")
        path = _worker.submit(_worker.op_save)
        msg = f"[저장 완료] {path}"
        if bk:
            msg += f"\n[덮어쓰기 전 백업] {bk}"
        return msg
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_save_as",
    annotations={"title": "다른 이름/형식으로 저장", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_save_as(params: SaveAsInput) -> str:
    """현재 문서를 지정 경로/형식으로 저장합니다. format: HWP / HWPX / PDF."""
    fmt = params.format.upper()
    if fmt not in ("HWP", "HWPX", "PDF"):
        return "Error: format은 HWP, HWPX, PDF 중 하나여야 합니다."
    if not os.path.isabs(params.path):
        return "Error: 절대 경로를 입력하세요."
    try:
        # 기존 파일을 덮어쓰는 경우 백업(가역성)
        bk = _backup_file(params.path, "presave") if os.path.exists(params.path) else None
        path = _worker.submit(lambda: _worker.op_save_as(params.path, fmt))
        msg = f"[저장 완료 / {fmt}] {path}"
        if bk:
            msg += f"\n[덮어쓰기 전 백업] {bk}"
        return msg
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_close",
    annotations={"title": "현재 문서 닫기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_close() -> str:
    """현재 문서를 닫습니다. 한/글 엔진은 다음 파일을 위해 백그라운드에 유지됩니다
    (엔진은 MCP 서버가 종료될 때 자동으로 정리됩니다)."""
    try:
        _worker.submit(_worker.op_close_doc)
        return "[문서 닫힘] 엔진은 대기 상태로 유지됩니다."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_status",
    annotations={"title": "상태 확인", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_status() -> str:
    """현재 열려 있는 문서 경로 등 상태를 반환합니다."""
    path = _worker.current_path
    if _worker._hwp is None:
        return "열린 문서 없음 (엔진 미기동)."
    return f"현재 문서: {path or '(새 문서 - 미저장)'}"


# ---------------------------------------------------------------------------
# 표/체크박스 입력 모델 + 도구
# ---------------------------------------------------------------------------

class TableMapInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    table_index: int = Field(..., description="표 번호(1부터). 문서 앞에서부터 순서.", ge=1)


class SetCellsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    table_index: int = Field(..., description="표 번호(1부터)", ge=1)
    cells: Dict[str, str] = Field(
        ..., description="{셀주소: 값}. 셀주소는 스프레드시트식(A1,B2,...). 예: {'C4':'홍길동','G9':'900101-1234567'}"
    )
    align: Optional[str] = Field(default=None, description="정렬: 'left'|'center'|'right' 또는 생략")


class CheckInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    labels: List[str] = Field(
        ..., description="체크할 □ 뒤 라벨 목록. 예: ['연소득 6,000만원 이하',' 운영자금']. '□'+라벨 형태로 매칭됨."
    )
    mark: str = Field(default="☑", description="채울 기호(기본 ☑)")


class CheckAfterInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    keyword: str = Field(..., description="이 텍스트 근처의 특수 네모(☐)를 체크. 예: '동의함'", min_length=1)
    mark: str = Field(default="☑", description="채울 기호(기본 ☑)")
    position: str = Field(default="after", description="네모 위치: 'after'(단어 뒤, '동의함 ☐') | 'before'(단어 앞, '☐ 동의함')")


@mcp.tool(
    name="hwp_table_map",
    annotations={"title": "표 셀 지도 읽기", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_table_map(params: TableMapInput) -> str:
    """지정한 표의 모든 셀을 {주소: 텍스트}로 반환합니다(빈 칸 = 채울 대상 파악용).

    표를 채우기 전에 먼저 호출해 라벨 셀과 빈 값 셀의 주소를 확인하세요.
    병합된 셀은 대표 주소 한 번만 나타납니다.

    Returns:
        str: JSON 배열 [{"addr":"A1","text":"..."}, ...]. 실패 시 'Error: ...'
    """
    try:
        import json
        data = _worker.submit(lambda: _worker.op_table_map(params.table_index))
        n = _worker.submit(_worker._count_tables)
        return f"[표 #{params.table_index} / 문서 내 표 총 {n}개]\n" + json.dumps(data, ensure_ascii=False, indent=1)
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_set_cells",
    annotations={"title": "표 셀 채우기", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_set_cells(params: SetCellsInput) -> str:
    """표의 여러 셀을 주소로 지정해 한 번에 채웁니다(기존 내용은 지우고 새로 입력).

    빈 표 양식 채우기의 핵심 도구입니다. 먼저 hwp_table_map으로 주소를 확인하세요.
    예: table_index=6, cells={'C4':'믿음상회','C9':'홍길동'}, align='center'

    Returns:
        str: 결과 요약(못 채운 주소 포함). 실패 시 'Error: ...'
    """
    if params.align not in (None, "left", "center", "right"):
        return "Error: align은 left/center/right 중 하나이거나 생략해야 합니다."
    try:
        miss = _worker.submit(lambda: _worker.op_set_cells(params.table_index, params.cells, params.align))
        filled = len(params.cells) - len(miss)
        msg = f"[표 #{params.table_index}] {filled}/{len(params.cells)}개 셀 입력 완료."
        if miss:
            msg += f" 못 찾은 주소: {miss}"
        return msg + "\n저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_check",
    annotations={"title": "네모 체크박스(□) 체크", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_check(params: CheckInput) -> str:
    """'□' 체크박스를 체크(☑)합니다. 각 라벨에 대해 '□'+라벨 → '☑'+라벨 로 치환.

    예: labels=['연소득 6,000만원 이하',' 운영자금',' 전,월세']
    주의: 같은 라벨이 여러 곳에 있으면 모두 체크됩니다(위치 구분이 필요하면 hwp_set_cells 사용).

    Returns:
        str: 라벨별 체크 건수. 실패 시 'Error: ...'
    """
    try:
        res = _worker.submit(lambda: _worker.op_check(params.labels, params.mark))
        lines = [f"  {'✓' if c else '✗(못찾음)'} □{lbl} ({c}건)" for lbl, c in res.items()]
        return "[체크 결과]\n" + "\n".join(lines) + "\n\n저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_check_after",
    annotations={"title": "특정 단어 뒤 네모(☐) 체크", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_check_after(params: CheckAfterInput) -> str:
    """특정 단어 바로 뒤의 특수 네모(☐, U+2610)를 체크(☑)합니다.

    일반 치환이 안 되는 동의서식 네모에 사용합니다.
    예: keyword='동의함' → 모든 '동의함' 뒤 네모를 ☑ 로 (동의하지 않음은 그대로).

    Returns:
        str: 처리 개수. 실패 시 'Error: ...'
    """
    try:
        n = _worker.submit(lambda: _worker.op_check_after(params.keyword, params.mark, params.position))
        return f"['{params.keyword}' 뒤 네모 {n}개 체크 완료]\n저장하려면 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


class CheckByLabelInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    labels: List[str] = Field(
        ..., description="체크할 옵션 텍스트 목록(네모 옆 단어). 예: ['동의함','운전자금']. 박스가 앞/뒤 어디든, □·☐ 두 종류 자동 처리."
    )
    mark: str = Field(default="☑", description="채울 기호(기본 ☑)")


@mcp.tool(
    name="hwp_check_by_label",
    annotations={"title": "라벨로 체크박스 체크(범용·권장)", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_check_by_label(params: CheckByLabelInput) -> str:
    """옵션 라벨 옆의 네모를 찾아 체크(☑)합니다. ★체크박스의 기본(권장) 도구★

    □(U+25A1)·☐(U+2610) 두 종류, 박스가 라벨 앞이든 뒤든 자동으로 처리합니다.
    (기존 hwp_check=□ 뒤만, hwp_check_after=☐ 뒤만 → 이걸로 통합)
    예: labels=['동의함','운전자금','전자상거래 업종']
    같은 라벨이 여러 곳이면 전부 체크됩니다.

    Returns:
        str: 라벨별 결과. 실패 시 'Error: ...'
    """
    try:
        res = _worker.submit(lambda: _worker.op_check_by_label(params.labels, params.mark))
        lines = [f"  {'✓' if str(v).startswith('OK') else '✗'} {k} → {v}" for k, v in res.items()]
        return "[라벨 기반 체크 결과]\n" + "\n".join(lines) + "\n\n확인은 hwp_render, 저장은 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


class FillByLabelInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    fills: Dict[str, Any] = Field(
        ..., description="{라벨: 값}. 값은 문자열이거나 {value, direction, nth, near} 딕셔너리(라벨별 세밀 제어). "
                         "라벨은 hwp_read_text에 보이는 그대로(공백 포함). near는 '이 텍스트 다음의 라벨'로 중복 라벨을 구분."
    )
    direction: str = Field(default="auto", description="값 위치: 'auto'(오른쪽→아래→인라인 자동, 권장) | 'right' | 'below' | 'inline'")
    nth: int = Field(default=1, description="같은 라벨이 여러 개면 몇 번째(1부터). 값 딕셔너리의 nth로 라벨별 지정 가능", ge=1)


@mcp.tool(
    name="hwp_fill_by_label",
    annotations={"title": "라벨로 채우기(범용·권장)", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_fill_by_label(params: FillByLabelInput) -> str:
    """라벨 텍스트를 찾아 그 인접 칸에 값을 채웁니다. ★양식 채우기의 기본(권장) 도구★

    table_map(셀 주소)이 셀을 못 읽는 폼에서도 동작합니다 — 사람처럼 "라벨 옆 칸"에 씁니다.
    셀 주소·병합·컨테이너 타입을 신경 쓸 필요가 없습니다.

    예: fills={'기 업 체 명':'위빌리브','대  표  자':'이선호','종 업 원 수':'5'}
    - 라벨은 hwp_read_text에 보이는 그대로(공백 포함) 넣으세요.
    - 값이 라벨 오른쪽이 아니라 아래면 direction='below', 같은 칸 안(라벨 뒤)이면 'inline'.
    - 채운 뒤 hwp_render로 눈으로 확인하세요.

    Returns:
        str: 라벨별 결과(OK / 라벨 못 찾음 / 인접 셀 없음). 실패 시 'Error: ...'
    """
    if params.direction not in ("auto", "right", "below", "inline"):
        return "Error: direction은 auto/right/below/inline 중 하나여야 합니다."
    try:
        res = _worker.submit(lambda: _worker.op_fill_by_label(params.fills, params.direction, params.nth))
        lines = [f"  {'✓' if str(v).startswith('OK') else '✗'} {k} → {v}" for k, v in res.items()]
        return "[라벨 기반 채우기 결과]\n" + "\n".join(lines) + "\n\n확인은 hwp_render, 저장은 hwp_save 를 호출하세요."
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


class ShowInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    show: bool = Field(default=True, description="True면 한/글 창을 화면에 표시, False면 숨김")


class RenderInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    pages: str = Field(default="all", description="렌더할 페이지. 'all' 또는 '1', '2-4' 형식")
    dpi: int = Field(default=110, description="해상도(기본 110). 글자 확인엔 100~130 권장", ge=60, le=300)
    out_dir: Optional[str] = Field(default=None, description="PNG 저장 폴더(절대경로). 생략 시 임시폴더. 접근 제한 환경에서 지정.")
    return_image: bool = Field(default=True, description="True면 이미지를 결과에 직접 담아 반환(Cowork/클라우드에서도 바로 보임). False면 파일 경로만.")


@mcp.tool(
    name="hwp_render",
    annotations={"title": "현재 문서를 이미지로 렌더(눈으로 확인)", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": False, "openWorldHint": False},
)
def hwp_render(params: RenderInput):
    """현재 문서를 PDF로 내보낸 뒤 지정 페이지를 PNG로 렌더링합니다.

    기본(return_image=True)은 **이미지를 결과에 직접 담아 반환**하므로, 로컬이든
    Cowork(클라우드)든 Claude가 바로 눈으로 봅니다(별도 Read 불필요). 채운 뒤 검증하는 표준 절차.
    return_image=False면 저장된 PNG 파일 경로만 반환.

    Returns:
        이미지 콘텐츠 목록(+요약) 또는 파일 경로. 실패 시 'Error: ...'
    """
    try:
        import tempfile
        try:
            import fitz  # PyMuPDF
        except Exception:
            return "Error: PyMuPDF(fitz)가 없습니다. 'pip install PyMuPDF' 후 사용하세요."
        outdir = params.out_dir or os.path.join(tempfile.gettempdir(), "hwp_mcp_render")
        os.makedirs(outdir, exist_ok=True)
        pdf = os.path.join(outdir, "current.pdf")
        _worker.submit(lambda: _worker.op_export_pdf(pdf))
        doc = fitz.open(pdf)
        total = doc.page_count
        # 페이지 범위 파싱
        sel = params.pages.strip().lower()
        if sel in ("all", ""):
            idxs = list(range(total))
        elif "-" in sel:
            a, b = sel.split("-", 1)
            idxs = list(range(int(a) - 1, min(int(b), total)))
        else:
            idxs = [int(sel) - 1]
        idxs = [i for i in idxs if 0 <= i < total]
        paths = []
        for i in idxs:
            p = os.path.join(outdir, f"page{i + 1:02d}.png")
            doc[i].get_pixmap(dpi=params.dpi).save(p)
            paths.append(p)
        doc.close()
        if not params.return_image:
            return (f"[{total}페이지 중 {len(paths)}페이지 렌더 완료] "
                    f"아래 PNG를 Read로 열어 확인하세요:\n" + "\n".join(paths))
        # 이미지를 결과에 직접 담아 반환(Cowork/클라우드에서도 바로 보임) — 페이로드 과다 방지 위해 상한
        MAX_IMG = 4
        note = f"[{total}페이지 중 {len(paths)}페이지 렌더]"
        if len(paths) > MAX_IMG:
            note += f" — 이미지는 앞 {MAX_IMG}페이지만 첨부(전체 PNG는 {outdir} 에 저장). 특정 페이지는 pages로 지정."
        out = [note]
        for p in paths[:MAX_IMG]:
            with open(p, "rb") as f:
                out.append(Image(data=f.read(), format="png"))
        return out
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_show_window",
    annotations={"title": "한글 창 표시/숨김", "readOnlyHint": False,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_show_window(params: ShowInput) -> str:
    """작업 중인 한/글 창을 화면에 표시하거나 숨깁니다(실시간 확인용).

    기본은 숨김 모드로 동작하지만, 진행 상황을 눈으로 보고 싶을 때 show=True 로 띄울 수 있습니다.
    """
    try:
        st = _worker.submit(lambda: _worker.op_show(params.show))
        return f"[한/글 창 {st}]"
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


class BackupPathInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    path: str = Field(..., description="백업 목록을 조회할 원본 파일의 절대 경로")


class RestoreInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    backup_path: str = Field(..., description="복원할 .bak 파일의 절대 경로 (hwp_list_backups로 확인)")
    target_path: str = Field(..., description="복원 대상(덮어쓸) 원본 경로")


@mcp.tool(
    name="hwp_list_backups",
    annotations={"title": "백업 목록", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
)
def hwp_list_backups(params: BackupPathInput) -> str:
    """지정 파일의 자동 백업 목록(최신순)을 반환합니다. 되돌리기(hwp_restore_backup) 전에 확인하세요."""
    try:
        items = _list_backups(params.path)
        if not items:
            return f"백업이 없습니다: {params.path}"
        lines = [f"  {i+1}. {p}" for i, p in enumerate(items)]
        return f"[{params.path} 백업 {len(items)}개 / 최신순]\n" + "\n".join(lines)
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_restore_backup",
    annotations={"title": "백업으로 되돌리기", "readOnlyHint": False,
                 "destructiveHint": True, "idempotentHint": False, "openWorldHint": False},
)
def hwp_restore_backup(params: RestoreInput) -> str:
    """백업(.bak)을 원본 경로로 복원(되돌리기)합니다. 복원 전, 현재 대상 파일도 안전을 위해 한 번 더 백업합니다.

    주의: 대상 파일이 현재 열려 있다면 먼저 hwp_close 후 복원하고, 필요하면 다시 hwp_open 하세요.
    """
    try:
        if not os.path.exists(params.backup_path):
            return f"Error: 백업 파일이 없습니다: {params.backup_path}"
        # 복원 직전 현재 상태도 백업(2중 안전 → 복원 자체도 되돌릴 수 있음)
        pre = _backup_file(params.target_path, "prerestore") if os.path.exists(params.target_path) else None
        # 대상이 엔진에 열려 있으면 파일 잠금이 걸리므로, 잠기면 엔진을 놓아주고 재시도
        last_err = None
        for attempt in range(5):
            try:
                shutil.copy2(params.backup_path, params.target_path)
                last_err = None
                break
            except PermissionError as pe:
                last_err = pe
                if attempt == 0:
                    _worker.submit(_worker.op_quit)  # 엔진 종료로 파일 잠금 해제
                time.sleep(0.3)
        if last_err is not None:
            return ("Error: 대상 파일이 사용 중이라 복원하지 못했습니다. "
                    "hwp_close 후 다시 시도하거나, 한/글에서 파일을 닫아주세요.")
        msg = f"[복원 완료] {params.backup_path}\n  → {params.target_path}"
        if pre:
            msg += f"\n[복원 전 상태 백업] {pre}"
        return msg
    except Exception as e:  # noqa: BLE001
        return f"Error: {e}"


@mcp.tool(
    name="hwp_version",
    annotations={"title": "버전/업데이트 확인", "readOnlyHint": True,
                 "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
def hwp_version() -> str:
    """현재 hwp-with-claude 버전을 반환하고, GitHub에서 최신 버전이 있는지 확인합니다."""
    msg = f"hwp with claude — 현재 버전 v{__version__}"
    if "CHANGE_ME" in _UPDATE_URL:
        return msg + "\n(업데이트 확인 URL 미설정)"
    try:
        import urllib.request, json as _json
        with urllib.request.urlopen(_UPDATE_URL, timeout=4) as r:
            latest = _json.loads(r.read().decode("utf-8")).get("version", "")
        if latest and latest != __version__:
            return msg + f"\n🔔 새 버전 v{latest} 있음! install.bat를 다시 실행해 업데이트하세요."
        return msg + "\n최신 버전입니다."
    except Exception:
        return msg + "\n(업데이트 확인 실패 — 오프라인이거나 저장소 접근 불가)"


if __name__ == "__main__":
    mcp.run()
