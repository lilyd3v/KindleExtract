import sys
import os
import subprocess
import psutil
import yaml

if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 경우
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(__file__)

CONFIG_PATH = os.path.join(BASE_DIR, "config.yml")

PROCDUMP = "procdump.exe"
KRFKEYEXTRACTOR = "KRFKeyExtractor.exe"

KINDLE_CONTENT = None
BOOK_PATH = None
LOCAL_APPDATA = os.environ["LOCALAPPDATA"]
APPLICATION_PATH = os.path.join(LOCAL_APPDATA, "Amazon", "Kindle", "application")


def ensure_config():
    # config.yml이 없으면 생성
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write(
                """# KindleExtract 설정 파일
# 아래 변수에 실제 경로를 입력하세요.
KINDLE_CONTENT: ""  # 원본 킨들 폴더 경로
BOOK_PATH: ""  # 복사될 폴더 경로

# 예시:
# KINDLE_CONTENT: "C:/Users/username/OneDrive/문서/My Kindle Content"
# BOOK_PATH: "C:/KindleExtract"
"""
            )
        print(f"config.yml 파일이 생성되었습니다: {CONFIG_PATH}\n설정 파일을 편집한 후 다시 실행하세요.")
        input("Press Enter to exit...")
        sys.exit()
    else:
        # config.yml이 있으면 값 읽어서 글로벌 변수에 할당
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        global KINDLE_CONTENT, BOOK_PATH
        KINDLE_CONTENT = str(config.get("KINDLE_CONTENT", "")).strip()
        BOOK_PATH = str(config.get("BOOK_PATH", "")).strip()


def copy_books():
    # Kindle 콘텐츠 폴더가 존재하는지 확인
    if not os.path.exists(KINDLE_CONTENT):
        raise FileNotFoundError(f"Kindle 콘텐츠 폴더를 찾을 수 없습니다: {KINDLE_CONTENT}\nconfig.yml의 KINDLE_CONTENT 경로를 확인하세요.")

    # BOOK_PATH가 없으면 생성
    if not os.path.exists(BOOK_PATH):
        os.makedirs(BOOK_PATH, exist_ok=True)

    # PowerShell을 이용해 관리자 권한으로 폴더 복사
    cmd = (
        f'Copy-Item -Path "{KINDLE_CONTENT}\\*" -Destination "{BOOK_PATH}" -Recurse -Force'
    )
    subprocess.run(["powershell", "-Command", cmd], check=True)
    print(f"Kindle 콘텐츠 복사 완료: {BOOK_PATH}")


def dump_process():
    # 실행 중인 Kindle.exe PID 찾기
    pid = None
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'].lower() == "kindle.exe":
            pid = proc.info['pid']
            break

    if pid is None:
        raise RuntimeError("Kindle.exe 프로세스를 찾을 수 없습니다.")

    if not os.path.exists(PROCDUMP):
        raise FileNotFoundError(f"{PROCDUMP} 를 찾을 수 없습니다.")

    # 덤프 생성
    try:
        subprocess.run([PROCDUMP, "-accepteula", "-ma", "-o", str(pid),
                        os.path.join(APPLICATION_PATH, "Kindle")],
                       check=False)
    except subprocess.CalledProcessError as e:
        print(f"ProcDump 경고: {e}, 무시합니다.")

    print(f"메모리 덤프 완료: {os.path.join(APPLICATION_PATH, 'Kindle.DMP')}\n")


def generate_key():
    key_extractor = os.path.join(APPLICATION_PATH, KRFKEYEXTRACTOR)
    dmp_path = os.path.join(APPLICATION_PATH, "Kindle.DMP")
    keyfile_path = os.path.join(APPLICATION_PATH, "kindlekey.file")

    if not os.path.exists(key_extractor):
        # 현재 스크립트 폴더에 KRFKeyExtractor.exe가 있으면 복사
        local_exe = os.path.join(os.path.dirname(__file__), KRFKEYEXTRACTOR)
        if os.path.exists(local_exe):
            import shutil
            shutil.copy2(local_exe, key_extractor)
            print(f"{local_exe} 를 {key_extractor} 로 복사했습니다.")
        else:
            raise FileNotFoundError(f"{key_extractor} 와 {local_exe} 둘 다 찾을 수 없습니다.")

    if not os.path.exists(dmp_path):
        raise FileNotFoundError(f"{dmp_path} 를 찾을 수 없습니다.")

    # KRFKeyExtractor 실행
    subprocess.run([key_extractor, dmp_path, BOOK_PATH, keyfile_path], check=True)
    print(f"KRFKeyExtractor 실행 완료: {keyfile_path}\n")


def main():
    ensure_config()

    try:
        copy_books()
        dump_process()
        generate_key()
        print("모든 작업이 완료되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


if __name__ == "__main__":
    main()
    input("Press Enter to exit...")
