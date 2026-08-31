# -*- coding: utf-8 -*-
"""hwpx_engine — HWPX(zip+XML) 직접 편집 엔진.

한/글(COM) 조종 없이 파일 구조를 직접 읽고 쓴다. 표 구조(셀 주소·병합·내용)가
XML에 명시돼 있어 '폼마다 되냐 안 되냐'가 없다. 커서·Find·클립보드 미사용.

역할 분담: 편집=이 엔진(유일), 한/글(COM)=hwp↔hwpx 변환·PDF 렌더 전용.
"""
from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS = {"hp": HP}


def _norm(s: str) -> str:
    """라벨 비교용 정규화: 모든 공백 제거."""
    return re.sub(r"\s+", "", s or "")


class Cell:
    def __init__(self, tc, table_index: int):
        self.tc = tc
        self.table_index = table_index
        a = tc.find("hp:cellAddr", NS)
        s = tc.find("hp:cellSpan", NS)
        self.row = int(a.get("rowAddr"))
        self.col = int(a.get("colAddr"))
        self.colspan = int(s.get("colSpan"))
        self.rowspan = int(s.get("rowSpan"))

    # -- 읽기 ---------------------------------------------------------------
    def text(self) -> str:
        return "".join(t.text or "" for t in self.tc.iter(f"{{{HP}}}t"))

    # -- 쓰기 ---------------------------------------------------------------
    def set_text(self, value: str) -> None:
        """셀 내용을 value로 교체. 첫 문단의 첫 run(폼 기본 charPr 보유)을 재사용하고
        나머지 t는 제거 → 폼 기본 서식(검정)으로 들어간다."""
        sub = self.tc.find("hp:subList", NS)
        paras = sub.findall("hp:p", NS)
        first_run = None
        for p in paras:
            for run in p.findall("hp:run", NS):
                for t in list(run.findall("hp:t", NS)):
                    run.remove(t)
                if first_run is None:
                    first_run = run
        if first_run is None:  # run이 하나도 없는 셀(드묾) — 문단에 새 run
            p = paras[0] if paras else ET.SubElement(sub, f"{{{HP}}}p")
            first_run = ET.SubElement(p, f"{{{HP}}}run")
        t = ET.SubElement(first_run, f"{{{HP}}}t")
        t.text = value

    def replace_text(self, find: str, repl: str, count: int = 0) -> int:
        """셀 전체 텍스트 기준 치환 — ★run(서식조각) 경계에 걸쳐 있어도 동작.
        count=0이면 전부. 반환: 치환 횟수."""
        ts = [t for t in self.tc.iter(f"{{{HP}}}t")]
        if not ts:
            return 0
        texts = [t.text or "" for t in ts]
        full = "".join(texts)
        if find not in full:
            return 0
        done = 0
        pos = 0
        while True:
            if count and done >= count:
                break
            i = full.find(find, pos)
            if i < 0:
                break
            j = i + len(find)
            # i..j 구간이 걸친 t들을 찾아 재분배
            off = 0
            spans = []  # (t_idx, local_start, local_end)
            for k, s in enumerate(texts):
                lo, hi = off, off + len(s)
                if hi > i and lo < j:
                    spans.append((k, max(0, i - lo), min(len(s), j - lo)))
                off = hi
            k0, s0, _ = spans[0]
            _, _, eN = spans[-1]
            kN = spans[-1][0]
            if k0 == kN:
                texts[k0] = texts[k0][:s0] + repl + texts[k0][eN:]
            else:
                texts[k0] = texts[k0][:s0] + repl
                for k, _, _ in spans[1:-1]:
                    texts[k] = ""
                texts[kN] = texts[kN][eN:]
            full = "".join(texts)
            pos = i + len(repl)
            done += 1
        for t, s in zip(ts, texts):
            t.text = s
        return done


class HwpxDoc:
    """열기 → 편집(셀/치환) → 저장. 원본 zip의 다른 항목은 바이트 그대로 보존."""

    SECTION_RE = re.compile(r"^Contents/section(\d+)\.xml$")

    def __init__(self, path: str):
        self.path = path
        self._zin = zipfile.ZipFile(path)
        # 원본 XML 선두의 네임스페이스 접두사 보존 등록
        self.sections: Dict[str, ET.Element] = {}
        for name in self._zin.namelist():
            if self.SECTION_RE.match(name):
                raw = self._zin.read(name).decode("utf-8")
                for pfx, uri in re.findall(r'xmlns:([a-zA-Z0-9]+)="([^"]+)"', raw[:4000]):
                    ET.register_namespace(pfx, uri)
                self.sections[name] = ET.fromstring(raw)
        self._cells: Optional[List[Cell]] = None

    # -- 구조 ---------------------------------------------------------------
    def cells(self) -> List[Cell]:
        if self._cells is None:
            out = []
            ti = 0
            for name in sorted(self.sections):
                for tbl in self.sections[name].iter(f"{{{HP}}}tbl"):
                    for tc in tbl.iter(f"{{{HP}}}tc"):
                        out.append(Cell(tc, ti))
                    ti += 1
            self._cells = out
        return self._cells

    def table_count(self) -> int:
        return 1 + max((c.table_index for c in self.cells()), default=-1)

    def table_map(self) -> List[Dict[str, Any]]:
        """표별 셀 목록: [{table,row,col,colspan,rowspan,text}] — 항상 정확(파일이 곧 진실)."""
        return [{"table": c.table_index, "row": c.row, "col": c.col,
                 "colspan": c.colspan, "rowspan": c.rowspan, "text": c.text()}
                for c in self.cells()]

    def full_text(self) -> str:
        """문서 전체 텍스트(문단 단위 줄바꿈 — read_text용)."""
        lines = []
        for name in sorted(self.sections):
            for p in self.sections[name].iter(f"{{{HP}}}p"):
                s = "".join(t.text or "" for t in p.iter(f"{{{HP}}}t"))
                lines.append(s)
        return "\n".join(lines)

    def set_cell_addr(self, table: int, row: int, col: int, value: str) -> bool:
        """표/행/열 주소로 셀 내용 교체(주소는 table_map이 알려준 값)."""
        for c in self.cells():
            if c.table_index == table and c.row == row and c.col == col:
                c.set_text(value)
                return True
        return False

    # -- 라벨 기반 채우기 -----------------------------------------------------
    def _grid(self, ti: int) -> Dict[Tuple[int, int], Cell]:
        return {(c.row, c.col): c for c in self.cells() if c.table_index == ti}

    def find_label_cells(self, label: str) -> List[Cell]:
        """셀 텍스트가 라벨과 일치(공백 무시)하는 셀들. 완전일치 우선, 없으면 부분일치."""
        nl = _norm(label)
        exact = [c for c in self.cells() if _norm(c.text()) == nl]
        if exact:
            return exact
        return [c for c in self.cells() if nl and nl in _norm(c.text())]

    def value_cell(self, label_cell: Cell, direction: str = "auto") -> Optional[Cell]:
        """라벨 셀의 값 칸: right=병합 감안한 바로 오른쪽, below=바로 아래."""
        g = self._grid(label_cell.table_index)

        def right():
            want = label_cell.col + label_cell.colspan
            return g.get((label_cell.row, want)) or next(
                (g[(r, c)] for (r, c) in sorted(g) if r == label_cell.row and c >= want), None)

        def below():
            want = label_cell.row + label_cell.rowspan
            return g.get((want, label_cell.col)) or next(
                (g[(r, c)] for (r, c) in sorted(g) if c == label_cell.col and r >= want), None)

        if direction == "right":
            return right()
        if direction == "below":
            return below()
        # auto: 오른쪽 우선, 없으면 아래
        return right() or below()

    def fill_by_label(self, fills: Dict[str, Any], direction: str = "auto",
                      nth: int = 1, table: Optional[int] = None) -> Dict[str, str]:
        """{라벨: 값} 채우기. 값은 str 또는 {value,direction,nth,table} dict.
        table 지정 시 그 표(0부터)의 라벨만 매칭(예: 빈 양식=0, 작성예시=1).
        반환: {라벨: 'OK(r,c)' | 실패이유} — XML은 쓰면 확실히 들어가므로 결과가 곧 사실."""
        out = {}
        for label, spec in fills.items():
            if isinstance(spec, dict):
                value = str(spec.get("value", ""))
                d = spec.get("direction", direction)
                n = int(spec.get("nth", nth))
                tb = spec.get("table", table)
            else:
                value, d, n, tb = str(spec), direction, nth, table
            cands = self.find_label_cells(label)
            if tb is not None:
                cands = [c for c in cands if c.table_index == int(tb)]
            if not cands:
                out[label] = "라벨 못 찾음"
                continue
            if n > len(cands):
                out[label] = f"라벨 {len(cands)}개뿐(nth={n})"
                continue
            lc = cands[n - 1]
            if d == "inline":
                lc.replace_text(lc.text(), lc.text() + value, count=1) if lc.text() else lc.set_text(value)
                out[label] = f"OK(inline,t{lc.table_index} r{lc.row}c{lc.col})"
                continue
            vc = self.value_cell(lc, d)
            if vc is None:
                out[label] = "인접 칸 없음"
                continue
            vc.set_text(value)
            out[label] = f"OK(t{vc.table_index} r{vc.row}c{vc.col})"
        return out

    # -- 체크박스 ------------------------------------------------------------
    BOXES = ("□", "☐", "■")

    def check_by_label(self, labels: List[str], mark: str = "☑",
                       table: Optional[int] = None) -> Dict[str, str]:
        """라벨 앞/뒤의 빈 네모(□/☐)나 [ ]를 체크 문자로 '교체'. run 경계 무관."""
        def _label_box_style(full: str) -> bool:
            """이 셀이 '라벨 뒤에 박스'(예: '동의함 ☐  동의하지 않음 ☐') 구조인가?
            셀 안 모든 박스 앞(공백 무시)이 글자로 끝날 때만 True.
            ('□ 운영자금 □ 시설자금'처럼 박스가 앞에 오는 구조에서 '라벨 뒤 박스' 패턴을
            돌리면 다음 항목의 박스를 훔쳐 체크하므로 반드시 구조를 판별한다.)"""
            ms = list(re.finditer(r"[□☐]|\[\s{1,2}\]", full))
            if not ms:
                return False
            for mm in ms:
                j = mm.start() - 1
                while j >= 0 and full[j] in " \t":
                    j -= 1
                ch = full[j] if j >= 0 else ""
                if not (ch.isalnum() or "가" <= ch <= "힣" or ch in ")%"):
                    return False
            # 결정적 판별: '라벨 뒤 박스' 구조는 마지막 박스 '뒤'에 글자가 없다
            # ('…동의하지 않음 [ ]'는 박스로 끝남 / '운영자금 □ 시설자금 □ 기술개발자금'은 라벨로 끝남)
            return not full[ms[-1].end():].strip()

        out = {}
        for label in labels:
            done = 0
            kinds = []
            for c in self.cells():
                if table is not None and c.table_index != int(table):
                    continue
                # ★패턴은 배타적으로: 셀에서 한 패턴이 맞으면 나머지는 시도하지 않는다.
                before_pats = []
                after_pats = []
                for box in ("□", "☐"):
                    before_pats += [(f"{box} {label}", f"{mark} {label}", "□"),
                                    (f"{box}{label}", f"{mark}{label}", "□")]
                    after_pats += [(f"{label} {box}", f"{label} {mark}", "□"),
                                   (f"{label}{box}", f"{label}{mark}", "□")]
                before_pats += [(f"[ ] {label}", f"[√] {label}", "[]"),
                                (f"[  ] {label}", f"[√] {label}", "[]"),
                                (f"[ ]{label}", f"[√]{label}", "[]")]
                after_pats += [(f"{label} [ ]", f"{label} [√]", "[]"),
                               (f"{label}[ ]", f"{label}[√]", "[]")]
                patterns = list(before_pats)
                if _label_box_style(c.text()):
                    patterns += after_pats  # '라벨 뒤 박스' 구조가 확실할 때만
                for f, r, kind in patterns:
                    n = c.replace_text(f, r)
                    if n:
                        done += n
                        kinds.append(kind)
                        break  # 이 셀은 처리 완료 — 다른 패턴 재시도 금지
            if done:
                out[label] = f"OK({'+'.join(sorted(set(kinds)))}×{done})"
            else:
                # 못 찾음 vs 이미 체크됨 구분(정직 보고)
                already = any(
                    (f"{mark}{label}" in _t or f"{mark} {label}" in _t
                     or f"{label} {mark}" in _t or f"{label}{mark}" in _t or f"[√] {label}" in _t)
                    for _t in (c.text() for c in self.cells()
                               if table is None or c.table_index == int(table)))
                out[label] = "이미 체크됨" if already else "박스 못 찾음(수동 체크 필요)"
        return out

    # -- 일반 치환 ------------------------------------------------------------
    def replace_all(self, find: str, repl: str) -> int:
        n = 0
        for c in self.cells():
            n += c.replace_text(find, repl)
        # 표 밖 본문 문단도 처리
        for name in sorted(self.sections):
            root = self.sections[name]
            in_cell_ts = {id(t) for c in self.cells() for t in c.tc.iter(f"{{{HP}}}t")}
            ts = [t for t in root.iter(f"{{{HP}}}t") if id(t) not in in_cell_ts]
            for t in ts:
                if t.text and find in t.text:
                    n += t.text.count(find)
                    t.text = t.text.replace(find, repl)
        return n

    # -- 저장 ----------------------------------------------------------------
    def save(self, out_path: Optional[str] = None) -> str:
        """저장. 원본 zip 항목은 바이트 그대로, section*.xml만 교체.
        mimetype은 무압축 첫 항목 유지(포맷 규약)."""
        out_path = out_path or self.path
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w") as zout:
            for item in self._zin.infolist():
                data = self._zin.read(item.filename)
                if item.filename in self.sections:
                    data = ET.tostring(self.sections[item.filename], encoding="unicode").encode("utf-8")
                comp = zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED
                zout.writestr(item, data, comp)
        if out_path == self.path:
            self._zin.close()
        shutil.move(tmp, out_path)
        if out_path == self.path:
            self._zin = zipfile.ZipFile(self.path)
        return out_path

    def close(self):
        try:
            self._zin.close()
        except Exception:
            pass
