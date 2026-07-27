# -*- coding: utf-8 -*-
"""hwp with claude - 설치(등록) 스크립트.
이 파일이 있는 폴더의 hwp_mcp_server.py를 Claude에 등록한다(경로 자동 감지).
파이썬 자동설치는 install.ps1이 담당하며, 이 스크립트는 그 이후 실행된다."""
import sys, os, json, subprocess, shutil, datetime, platform

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "hwp_mcp_server.py")
PY = sys.executable
LOG = os.path.join(HERE, "install_result.txt")
_lines = []
STATUS = []  # (기호, 설명)

def out(*a):
    s = " ".join(str(x) for x in a)
    _lines.append(s)
    try:
        print(s)
    except Exception:
        pass

def mark(ok, label):
    STATUS.append(("[✓]" if ok else "[!]", label))

def claude_running():
    """Claude 데스크톱/코드가 실행 중인지 확인(설정 적용엔 재시작 필요)."""
    try:
        r = subprocess.run(["tasklist"], capture_output=True, text=True)
        return "claude" in (r.stdout or "").lower()
    except Exception:
        return False

def server_entry():
    return {"command": PY, "args": [SERVER], "env": {}}

def merge_hwp(cfg_path, create_if_missing):
    if not os.path.exists(cfg_path):
        if not create_if_missing:
            return False, "파일 없음(건너뜀)"
        d = {}
    else:
        try:
            d = json.load(open(cfg_path, encoding="utf-8"))
        except Exception as e:
            return False, f"읽기 실패: {e}"
        shutil.copy(cfg_path, cfg_path + ".bak_" +
                    datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    d.setdefault("mcpServers", {})["hwp"] = server_entry()
    os.makedirs(os.path.dirname(cfg_path), exist_ok=True)
    json.dump(d, open(cfg_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return True, "등록됨"

def main():
    out("hwp with claude 설치 시작...")
    out("파이썬:", PY)

    # 0) 환경
    if platform.system() != "Windows":
        out("[중단] Windows 전용입니다."); mark(False, "Windows"); return
    mark(True, "Windows")
    if not os.path.exists(SERVER):
        out("[중단] hwp_mcp_server.py 없음. 압축을 제대로 풀었는지 확인."); return

    # 1) 패키지
    out("패키지 설치 중 (pywin32, mcp, PyMuPDF)...")
    try:
        subprocess.run([PY, "-m", "pip", "install", "--upgrade",
                        "pywin32", "mcp", "PyMuPDF"], check=False)
        mark(True, "파이썬 패키지 설치")
    except Exception as e:
        out("[경고] 패키지 설치 오류:", e); mark(False, "파이썬 패키지 설치")

    # 2) 한/글
    hangul_ok = False
    try:
        import win32com.client as w
        h = w.Dispatch("HWPFrame.HwpObject"); h.Quit(); hangul_ok = True
    except Exception:
        hangul_ok = False
    mark(hangul_ok, "한/글 설치 확인" if hangul_ok else "한/글 미설치 → 한/글부터 설치하세요")

    # 3) Claude Code 등록
    ok, _ = merge_hwp(os.path.expanduser("~/.claude.json"), create_if_missing=True)
    mark(ok, "Claude Code 등록 (~/.claude.json)")

    # 4) Claude Desktop/Cowork 등록
    dpath = os.path.join(os.environ.get("APPDATA", ""), "Claude", "claude_desktop_config.json")
    ok2, msg2 = merge_hwp(dpath, create_if_missing=False)
    mark(ok2, "Claude Desktop/Cowork 등록" if ok2 else f"Claude Desktop 등록 건너뜀({msg2})")

    # 5) 스킬
    src_skill = os.path.join(HERE, "skills", "hwp-with-claude", "SKILL.md")
    if os.path.exists(src_skill):
        dst_dir = os.path.expanduser("~/.claude/skills/hwp-with-claude")
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copy(src_skill, os.path.join(dst_dir, "SKILL.md"))
        mark(True, "스킬 설치")
    else:
        mark(False, "스킬 파일 없음(건너뜀)")

    # ---- 결과 요약 ----
    out("")
    out("=" * 50)
    out("  hwp with claude - 설치 결과")
    out("=" * 50)
    for sym, label in STATUS:
        out(f"  {sym} {label}")

    running = claude_running()
    out("")
    out("━" * 50)
    out("  ★ 다음 할 일 (이거 안 하면 도구가 안 떠요!)")
    out("━" * 50)
    out("  1) Claude를 완전히 종료했다가 다시 켜세요.")
    out("     - Claude Code: 재시작 또는 새 세션 시작")
    out("     - Claude 데스크톱: 트레이(우측하단) 아이콘 우클릭 → 종료 → 다시 실행")
    out("  2) 새 채팅에서 확인:  \"hwp_status 실행해줘\"")
    out("  3) 그다음:  \"C:\\경로\\양식.hwp 열어서 채워줘\"")
    if running:
        out("")
        out("  ⚠️ 지금 Claude가 켜져 있습니다 → 반드시 껐다 켜야 적용됩니다.")
    if not hangul_ok:
        out("")
        out("  ⚠️ 한/글(HWP)이 없어 보입니다. 한/글 설치 후 사용하세요.")
    out("")
    out("  [AI 어시스턴트에게] 이 설치를 대신 실행했다면, 사용자에게 위")
    out("  '재시작'을 반드시 안내하고, 재시작 후 hwp_status로 작동을 확인시키세요.")
    out("=" * 50)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        out("[EXCEPTION]", repr(e))
    try:
        open(LOG, "w", encoding="utf-8").write("\n".join(_lines))
    except Exception:
        pass
