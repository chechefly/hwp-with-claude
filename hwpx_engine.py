# -*- coding: utf-8 -*-
"""hwpx_engine — HWPX(zip+XML) 직접 편집 엔진.

한/글(COM) 조종 없이 파일 구조를 직접 읽고 쓴다. 표 구조(셀 주소·병합·내용)가
XML에 명시돼 있어 '폼마다 되냐 안 되냐'가 없다. 커서·Find·클립보드 미사용.

역할 분담: 편집=이 엔진(유일), 한/글(COM)=hwp↔hwpx 변환·PDF 렌더 전용.
"""
from __future__ import annotations

import copy
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH = "http://www.hancom.co.kr/hwpml/2011/head"
NS = {"hp": HP, "hh": HH}


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
    def set_text(self, value: str, clean=None) -> None:
        """셀 내용을 value로 교체.
        ★value에 줄바꿈(\\n)이 있으면 줄마다 문단을 만든다(여러 줄 칸·줄바꿈 정리용).
        ★clean(orig_charpr_id)->new_id 콜백이 오면 삽입 run의 글자모양을 '정리된 서식'
          (검정·강조 해제)으로 교체 — 폼 칸에 걸린 이상한 색/굵기를 안 물려받는다."""
        lines = str(value).split("\n")
        sub = self.tc.find("hp:subList", NS)
        paras = sub.findall("hp:p", NS)
        # 기존 문단은 첫 번째만 남기고 제거(내용 통째 교체 — 어색한 줄바꿈 정리에도 사용)
        for p in paras[1:]:
            sub.remove(p)
        first = paras[0] if paras else ET.SubElement(sub, f"{{{HP}}}p")
        for run in first.findall("hp:run", NS):
            for t in list(run.findall("hp:t", NS)):
                run.remove(t)
        run0 = first.find("hp:run", NS)
        if run0 is None:  # run이 하나도 없는 셀(드묾)
            run0 = ET.SubElement(first, f"{{{HP}}}run")
        if clean is not None:
            cid = clean(run0.get("charPrIDRef"))
            if cid is not None:
                run0.set("charPrIDRef", cid)
        t = ET.SubElement(run0, f"{{{HP}}}t")
        t.text = lines[0]
        # 2번째 줄부터: 첫 문단을 복제(문단모양·글자모양 유지)해 텍스트만 바꿔 추가
        for ln in lines[1:]:
            np_ = copy.deepcopy(first)
            for run in np_.findall("hp:run", NS):
                for tt in list(run.findall("hp:t", NS)):
                    run.remove(tt)
            rr = np_.find("hp:run", NS)
            tt = ET.SubElement(rr, f"{{{HP}}}t")
            tt.text = ln
            sub.append(np_)

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
        # header.xml — 글자모양(charPr) 정의부. '정리된 서식'(검정·강조해제) 생성에 사용.
        self.header: Optional[ET.Element] = None
        self._header_name = "Contents/header.xml"
        if self._header_name in self._zin.namelist():
            raw = self._zin.read(self._header_name).decode("utf-8")
            for pfx, uri in re.findall(r'xmlns:([a-zA-Z0-9]+)="([^"]+)"', raw[:4000]):
                ET.register_namespace(pfx, uri)
            self.header = ET.fromstring(raw)
        self._clean_cache: Dict[str, Optional[str]] = {}
        self._cells: Optional[List[Cell]] = None

    def _clean_charpr(self, orig_id: Optional[str]) -> Optional[str]:
        """orig_id 글자모양을 복제해 '정리된 서식'(검정·바탕없음·굵기/기울임/밑줄 해제)으로
        만들어 header에 등록하고 새 id를 반환. 폼 칸의 이상한 색/강조를 안 물려받게 한다.
        (글꼴·크기는 폼 정의를 유지 — 문서 통일감)"""
        if self.header is None or orig_id is None:
            return None
        if orig_id in self._clean_cache:
            return self._clean_cache[orig_id]
        props = self.header.find(f".//{{{HH}}}charProperties")
        src = None
        if props is not None:
            for cp in props.findall(f"{{{HH}}}charPr"):
                if cp.get("id") == str(orig_id):
                    src = cp
                    break
        if src is None or props is None:
            self._clean_cache[orig_id] = None
            return None
        new = copy.deepcopy(src)
        new.set("textColor", "#000000")
        new.set("shadeColor", "none")
        for tag in ("bold", "italic", "outline", "shadow", "emboss", "engrave",
                    "supscript", "subscript"):
            for el in new.findall(f"{{{HH}}}{tag}"):
                new.remove(el)
        ul = new.find(f"{{{HH}}}underline")
        if ul is not None:
            ul.set("type", "NONE")
        new_id = str(max((int(cp.get("id", 0)) for cp in props.findall(f"{{{HH}}}charPr")), default=0) + 1)
        new.set("id", new_id)
        props.append(new)
        props.set("itemCnt", str(len(props.findall(f"{{{HH}}}charPr"))))
        self._clean_cache[orig_id] = new_id
        return new_id

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

    def set_cell_addr(self, table: int, row: int, col: int, value: str,
                      text_style: str = "reset") -> bool:
        """표/행/열 주소로 셀 내용 교체(주소는 table_map이 알려준 값)."""
        clean = self._clean_charpr if text_style == "reset" else None
        for c in self.cells():
            if c.table_index == table and c.row == row and c.col == col:
                c.set_text(value, clean=clean)
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
                      nth: int = 1, table: Optional[int] = None,
                      text_style: str = "reset") -> Dict[str, str]:
        """{라벨: 값} 채우기. 값은 str 또는 {value,direction,nth,table} dict.
        table 지정 시 그 표(0부터)의 라벨만 매칭(예: 빈 양식=0, 작성예시=1).
        반환: {라벨: 'OK(r,c)' | 실패이유} — XML은 쓰면 확실히 들어가므로 결과가 곧 사실.
        text_style='reset'(기본): 삽입값을 검정·강조해제 서식으로(칸의 이상한 색/굵기 무시)."""
        clean = self._clean_charpr if text_style == "reset" else None
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
                if lc.text():
                    lc.replace_text(lc.text(), lc.text() + value, count=1)
                else:
                    lc.set_text(value, clean=clean)
                out[label] = f"OK(inline,t{lc.table_index} r{lc.row}c{lc.col})"
                continue
            vc = self.value_cell(lc, d)
            if vc is None:
                out[label] = "인접 칸 없음"
                continue
            vc.set_text(value, clean=clean)
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

    # -- 이미지 삽입 ----------------------------------------------------------
    @staticmethod
    def _img_size_px(path: str) -> Tuple[int, int]:
        """PNG/JPG 픽셀 크기(의존성 없이 헤더 파싱)."""
        import struct
        with open(path, "rb") as f:
            head = f.read(26)
            if head[:8] == b"\x89PNG\r\n\x1a\n":
                w, h = struct.unpack(">II", head[16:24])
                return w, h
            if head[:2] == b"\xff\xd8":  # JPEG
                f.seek(2)
                while True:
                    b = f.read(1)
                    if not b:
                        break
                    if b != b"\xff":
                        continue
                    marker = f.read(1)
                    if marker in (b"\xc0", b"\xc1", b"\xc2", b"\xc3"):
                        f.read(3)
                        h, w = struct.unpack(">HH", f.read(4))
                        return w, h
                    if marker in (b"\xd8", b"\xd9") or b"\xd0" <= marker <= b"\xd7":
                        continue
                    (ln,) = struct.unpack(">H", f.read(2))
                    f.seek(ln - 2, 1)
        raise ValueError("PNG/JPG만 지원합니다: " + path)

    def insert_image(self, table: int, row: int, col: int, image_path: str,
                     width_mm: float = 25.0, keep_text: bool = True,
                     after_text: Optional[str] = None) -> str:
        """이미지(PNG/JPG)를 지정 셀 안에 '글자처럼'(treatAsChar) 삽입 — 셀 흐름을 따르므로
        표 레이아웃이 어긋나지 않는다. 서명/도장용. width_mm로 폭 지정(높이는 비율 자동).
        keep_text=False면 셀 기존 텍스트를 지우고 이미지만 남긴다."""
        import os as _os
        target = None
        for c in self.cells():
            if c.table_index == table and c.row == row and c.col == col:
                target = c
                break
        if target is None:
            return f"셀 없음: t{table} r{row}c{col}"
        px_w, px_h = self._img_size_px(image_path)
        # HWPUNIT = 1/7200 inch. 픽셀은 96dpi 가정(px*75), 표시 크기는 mm 지정.
        org_w, org_h = px_w * 75, px_h * 75
        cur_w = int(width_mm / 25.4 * 7200)
        cur_h = int(cur_w * px_h / px_w)
        # 바이너리 id 부여 + 저장 대기 목록에 등록
        ext = _os.path.splitext(image_path)[1].lower().lstrip(".") or "png"
        ext = "jpg" if ext in ("jpeg", "jpg") else "png"
        n = 1
        existing = set(self._zin.namelist()) | {p for p, _ in getattr(self, "_new_bins", [])}
        while any(x.startswith(f"BinData/inserted{n}.") for x in existing):
            n += 1
        bin_name = f"BinData/inserted{n}.{ext}"
        item_id = f"inserted{n}"
        if not hasattr(self, "_new_bins"):
            self._new_bins = []
        with open(image_path, "rb") as f:
            self._new_bins.append((bin_name, f.read()))
        # content.hpf 매니페스트 등록(save 때 문자열 치환으로 주입)
        if not hasattr(self, "_new_items"):
            self._new_items = []
        mt = "image/png" if ext == "png" else "image/jpeg"
        self._new_items.append(f'<opf:item id="{item_id}" href="{bin_name}" media-type="{mt}" isEmbeded="1"/>')
        # hp:pic 요소(글자취급) 생성 — 셀 첫 문단 run에 추가
        sca = cur_w / org_w if org_w else 1.0
        pic_xml = (
            f'<hp:pic xmlns:hp="{HP}" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
            f'id="{2000000000 + n}" zOrder="{n}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" '
            f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" '
            f'instid="{2000000000 + n}" reverse="0">'
            f'<hp:offset x="0" y="0"/><hp:orgSz width="{org_w}" height="{org_h}"/>'
            f'<hp:curSz width="{cur_w}" height="{cur_h}"/><hp:flip horizontal="0" vertical="0"/>'
            f'<hp:rotationInfo angle="0" centerX="{cur_w // 2}" centerY="{cur_h // 2}" rotateimage="1"/>'
            f'<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
            f'<hc:scaMatrix e1="{sca:.6f}" e2="0" e3="0" e4="0" e5="{sca:.6f}" e6="0"/>'
            f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>'
            f'<hc:img binaryItemIDRef="{item_id}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
            f'<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{org_w}" y="0"/>'
            f'<hc:pt2 x="{org_w}" y="{org_h}"/><hc:pt3 x="0" y="{org_h}"/></hp:imgRect>'
            f'<hp:imgClip left="0" right="{org_w}" top="0" bottom="{org_h}"/>'
            f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:imgDim dimwidth="{org_w}" dimheight="{org_h}"/><hp:effects/>'
            f'<hp:sz width="{cur_w}" widthRelTo="ABSOLUTE" height="{cur_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
            f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" '
            f'vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic>'
        )
        pic = ET.fromstring(pic_xml)
        if not keep_text:
            target.set_text("")
        sub = target.tc.find("hp:subList", NS)
        # after_text: 셀 안 그 텍스트가 있는 '문단'을 찾아 그 줄 뒤에 삽입(서명란의 '서명 또는 날인' 옆 등)
        p = None
        if after_text:
            for cand in sub.findall("hp:p", NS):
                if after_text.replace(" ", "") in "".join(
                        t.text or "" for t in cand.iter(f"{{{HP}}}t")).replace(" ", ""):
                    p = cand
                    break
            if p is None:
                return f"셀 안에서 '{after_text}' 문단을 못 찾음"
        if p is None:
            p = sub.find("hp:p", NS)
        run = p.findall("hp:run", NS)[-1] if p.findall("hp:run", NS) else ET.SubElement(p, f"{{{HP}}}run")
        run.append(pic)
        return f"OK(t{table} r{row}c{col}, {width_mm}mm, {bin_name})"

    # -- 표 행 추가/삭제 -------------------------------------------------------
    def _tables(self) -> List[ET.Element]:
        out = []
        for name in sorted(self.sections):
            out.extend(self.sections[name].iter(f"{{{HP}}}tbl"))
        return out

    def _invalidate(self):
        self._cells = None

    def add_rows(self, table: int, after_row: int, count: int = 1) -> str:
        """after_row 행을 템플릿으로 복제해 바로 아래에 count개 행 추가.
        서식·높이를 템플릿에서 그대로 물려받아 표 모양이 통일된다.
        템플릿 행 위에서 내려오는 세로병합(rowspan)은 자동 확장."""
        tbls = self._tables()
        if table >= len(tbls):
            return f"표 없음: t{table}"
        tbl = tbls[table]
        trs = tbl.findall(f"{{{HP}}}tr")
        tmpl = None
        for tr in trs:
            tcs = tr.findall(f"{{{HP}}}tc")
            if tcs and int(tcs[0].find("hp:cellAddr", NS).get("rowAddr")) == after_row:
                tmpl = tr
                break
        if tmpl is None:
            return f"행 없음: r{after_row}"
        # 템플릿 행 셀에 세로병합 시작이 있으면 거부(복제 시 구조 깨짐)
        for tc in tmpl.findall(f"{{{HP}}}tc"):
            if int(tc.find("hp:cellSpan", NS).get("rowSpan")) != 1:
                return f"r{after_row}에 세로병합 셀이 있어 템플릿으로 쓸 수 없음(다른 행 지정)"
        # 위에서 내려와 템플릿 행을 관통/포함하는 rowspan은 +count 확장
        for tc in tbl.iter(f"{{{HP}}}tc"):
            a, s = tc.find("hp:cellAddr", NS), tc.find("hp:cellSpan", NS)
            r0, rs = int(a.get("rowAddr")), int(s.get("rowSpan"))
            if rs > 1 and r0 <= after_row < r0 + rs:
                s.set("rowSpan", str(rs + count))
        # 이후 행들의 rowAddr 시프트
        for tc in tbl.iter(f"{{{HP}}}tc"):
            a = tc.find("hp:cellAddr", NS)
            r0 = int(a.get("rowAddr"))
            if r0 > after_row:
                a.set("rowAddr", str(r0 + count))
        # 복제 삽입
        idx = list(tbl).index(tmpl)
        for i in range(count):
            new_tr = copy.deepcopy(tmpl)
            for tc in new_tr.findall(f"{{{HP}}}tc"):
                tc.find("hp:cellAddr", NS).set("rowAddr", str(after_row + 1 + i))
                # 내용 비우기(서식 유지)
                sub = tc.find("hp:subList", NS)
                for p in sub.findall("hp:p", NS)[1:]:
                    sub.remove(p)
                first = sub.find("hp:p", NS)
                if first is not None:
                    for run in first.findall("hp:run", NS):
                        for t in list(run.findall("hp:t", NS)):
                            run.remove(t)
            tbl.insert(idx + 1 + i, new_tr)
        tbl.set("rowCnt", str(int(tbl.get("rowCnt", len(trs))) + count))
        self._invalidate()
        return f"OK(r{after_row} 아래 {count}행 추가)"

    def delete_rows(self, table: int, row: int, count: int = 1) -> str:
        """row부터 count개 행 삭제. 삭제 행에 세로병합 시작 셀이 있으면 거부(정직).
        위에서 내려오는 rowspan은 자동 축소."""
        tbls = self._tables()
        if table >= len(tbls):
            return f"표 없음: t{table}"
        tbl = tbls[table]
        victims = []
        for tr in tbl.findall(f"{{{HP}}}tr"):
            tcs = tr.findall(f"{{{HP}}}tc")
            if not tcs:
                continue
            r0 = int(tcs[0].find("hp:cellAddr", NS).get("rowAddr"))
            if row <= r0 < row + count:
                for tc in tcs:
                    if int(tc.find("hp:cellSpan", NS).get("rowSpan")) != 1:
                        return f"r{r0}에 세로병합 시작 셀이 있어 삭제 불가(수동 처리 필요)"
                victims.append(tr)
        if not victims:
            return f"행 없음: r{row}"
        # 관통 rowspan 축소
        for tc in tbl.iter(f"{{{HP}}}tc"):
            a, s = tc.find("hp:cellAddr", NS), tc.find("hp:cellSpan", NS)
            r0, rs = int(a.get("rowAddr")), int(s.get("rowSpan"))
            if rs > 1 and r0 < row and r0 + rs > row:
                overlap = min(r0 + rs, row + count) - row
                s.set("rowSpan", str(rs - overlap))
        for tr in victims:
            tbl.remove(tr)
        for tc in tbl.iter(f"{{{HP}}}tc"):
            a = tc.find("hp:cellAddr", NS)
            r0 = int(a.get("rowAddr"))
            if r0 >= row + count:
                a.set("rowAddr", str(r0 - len(victims)))
        tbl.set("rowCnt", str(int(tbl.get("rowCnt", 0)) - len(victims)))
        self._invalidate()
        return f"OK({len(victims)}행 삭제)"

    def scale_row_heights(self, table: int, factor: float,
                          rows: Optional[List[int]] = None) -> str:
        """표 행 높이를 factor배로(페이지 맞춤용 손잡이). rows 지정 시 그 행들만.
        내용이 많아 자동으로 늘어난 행은 한/글이 다시 계산하므로 '최소 높이'를 줄이는 효과."""
        tbls = self._tables()
        if table >= len(tbls):
            return f"표 없음: t{table}"
        n = 0
        for tc in tbls[table].iter(f"{{{HP}}}tc"):
            r0 = int(tc.find("hp:cellAddr", NS).get("rowAddr"))
            if rows is not None and r0 not in rows:
                continue
            sz = tc.find("hp:cellSz", NS)
            sz.set("height", str(max(500, int(int(sz.get("height")) * factor))))
            n += 1
        return f"OK({n}셀 높이 x{factor})"

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
            new_items = "".join(getattr(self, "_new_items", []))
            for item in self._zin.infolist():
                data = self._zin.read(item.filename)
                if item.filename in self.sections:
                    data = ET.tostring(self.sections[item.filename], encoding="unicode").encode("utf-8")
                elif item.filename == self._header_name and self.header is not None:
                    data = ET.tostring(self.header, encoding="unicode").encode("utf-8")
                elif item.filename == "Contents/content.hpf" and new_items:
                    # 삽입 이미지의 매니페스트 항목 주입
                    s = data.decode("utf-8")
                    data = s.replace("</opf:manifest>", new_items + "</opf:manifest>", 1).encode("utf-8")
                comp = zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED
                zout.writestr(item, data, comp)
            for bin_name, blob in getattr(self, "_new_bins", []):
                zout.writestr(bin_name, blob, zipfile.ZIP_DEFLATED)
        if out_path == self.path:
            self._zin.close()
        shutil.move(tmp, out_path)
        if out_path == self.path:
            self._zin = zipfile.ZipFile(self.path)
            # 새 바이너리/매니페스트는 이제 파일에 반영됐으므로 중복 주입 방지
            self._new_bins = []
            self._new_items = []
        return out_path

    def close(self):
        try:
            self._zin.close()
        except Exception:
            pass
