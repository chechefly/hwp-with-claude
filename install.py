# -*- coding: utf-8 -*-
"""hwp with claude - 자동 설치 스크립트.
이 파일이 있는 폴더의 hwp_mcp_server.py를 Claude에 등록한다(경로 자동 감지)."""
import sys, os, json, subprocess, shutil, datetime, platform

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "hwp_mcp_server.py")
PY = sys.executable  # 이 스크립트를 실행 중인 파이썬 = 서버 실행에 쓸 파이썬
LOG = os.path.join(HERE, "install_result.txt")
_lines = []

def out(*a):
    s = " ".join(str(x) for x in a)
    _lines.append(s)
    try:
        print(s)
    except Exception:
        pass

def step(title):
    out("")
    out("== " + title + " ==")

def server_entry():
    return {"command": PY, "args": [SERVER], "env": {}}

def merge_hwp(cfg_path, create_if_missing):
    """설정 파일에 hwp 항목을 병합(기존 항목 보존). 반환: (성공?, 메시지)"""
    if not os.path.exists(cfg_path):
        if not create_if_missing:
            return False, f"파일 없음(건너뜀): {cfg_path}"
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
    return True, f"등록됨: {cfg_path}"

def main():
    out("=========================================")
    out("  hwp with claude - 설치")
    out("=========================================")
    out("파이썬:", PY)
    out("서버:", SERVER)

    # 0) 환경 체크
    step("환경 확인")
    if platform.system() != "Windows":
        out("[중단] 이 도구는 Windows 전용입니다 (한/글이 Windows 전용).")
        return
    out("OK: Windows")
    if not os.path.exists(SERVER):
        out("[중단] hwp_mcp_server.py를 찾을 수 없습니다. 압축을 제대로 풀었는지 확인하세요.")
        return

    # 1) 의존성 설치
    step("필요한 파이썬 패키지 설치 (pywin32, mcp, PyMuPDF)")
    try:
        subprocess.run([PY, "-m", "pip", "install", "--upgrade",
                        "pywin32", "mcp", "PyMuPDF"], check=False)
        out("OK: 패키지 설치 시도 완료")
    except Exception as e:
        out("[경고] 패키지 설치 중 오류:", e)

    # 2) 한/글 설치 여부 확인(소프트)
    step("한/글(HWP) 설치 확인")
    try:
        import win32com.client  # noqa
        try:
            import pythoncom  # noqa
            import win32com.client as w
            h = w.Dispatch("HWPFrame.HwpObject")
            h.Quit()
            out("OK: 한/글 자동화 사용 가능")
        except Exception:
            out("[경고] 한/글이 설치돼 있지 않거나 자동화 불가. 한/글 설치 후 다시 시도하세요.")
    except Exception:
        out("[정보] pywin32 확인 필요(위 설치 단계 참고)")

    # 3) Claude Code 등록 (~/.claude.json)
    step("Claude Code 등록 (~/.claude.json)")
    ok, msg = merge_hwp(os.path.expanduser("~/.claude.json"), create_if_missing=True)
    out(("OK: " if ok else "건너뜀: ") + msg)

    # 4) Claude Desktop/Cowork 등록 (claude_desktop_config.json)
    step("Claude Desktop/Cowork 등록")
    dpath = os.path.join(os.environ.get("APPDATA", ""), "Claude", "claude_desktop_config.json")
    ok2, msg2 = merge_hwp(dpath, create_if_missing=False)
    out(("OK: " if ok2 else "건너뜀: ") + msg2)
    out("  ※ Desktop/Cowork에서 쓰려면: 모든 Claude 완전종료 후 이 설치 실행 → Desktop만 재시작")

    # 5) 스킬 설치 (선택)
    step("스킬 설치 (~/.claude/skills)")
    src_skill = os.path.join(HERE, "skills", "hwp-with-claude", "SKILL.md")
    if os.path.exists(src_skill):
        dst_dir = os.path.join(os.path.expanduser("~/.claude/skills/hwp-with-claude"))
        os.makedirs(dst_dir, exist_ok=True)
        shutil.copy(src_skill, os.path.join(dst_dir, "SKILL.md"))
        out("OK: 스킬 설치됨")
    else:
        out("건너뜀: 스킬 파일 없음")

    # 완료
    out("")
    out("=========================================")
    out("  설치 완료! 다음 순서로 사용하세요")
    out("=========================================")
    out("1. Claude 데스크톱 앱을 완전히 종료 후 다시 실행")
    out("2. 새 채팅에서:  이 한글 파일 열어서 채워줘  (파일 경로와 함께)")
    out("3. 끝! 문제가 있으면 install_result.txt 내용을 공유하세요.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        out("[EXCEPTION]", repr(e))
    try:
        open(LOG, "w", encoding="utf-8").write("\n".join(_lines))
    except Exception:
        pass
