
import sys
import subprocess

# 필요한 라이브러리 목록
REQUIRED_PACKAGES = [
    "google-generativeai",
    "openpyxl",
    "pypdf",
    "pywin32"
]

def install_packages():
    """필요한 패키지들을 확인하고 설치합니다."""
    print("필요한 라이브러리 설치를 시작합니다...")
    
    # 파이썬 실행 파일 경로를 사용하여 pip를 정확히 호출
    python_executable = sys.executable

    for package in REQUIRED_PACKAGES:
        try:
            print(f"--- '{package}' 설치 확인 및 진행 ---")
            # pip를 사용하여 패키지 설치
            subprocess.check_call([python_executable, "-m", "pip", "install", package])
            print(f"'{package}' 설치 완료 또는 이미 설치됨\n")
        except subprocess.CalledProcessError as e:
            print(f"오류: '{package}' 설치에 실패했습니다. pip가 설치되어 있는지, 인터넷 연결이 정상인지 확인하세요.")
            print(f"에러 내용: {e}")
            # 하나의 패키지라도 실패하면 스크립트 중단
            return False
        except FileNotFoundError:
            print("오류: 'pip' 명령을 찾을 수 없습니다. Python 또는 pip가 시스템 경로에 올바르게 설정되었는지 확인하세요.")
            return False
            
    print("모든 필수 라이브러리가 준비되었습니다.")
    return True

if __name__ == "__main__":
    install_packages()
