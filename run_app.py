import streamlit.web.cli as stcli
import os, sys

def resolve_path(path):
    basedir = getattr(sys, '_MEIPASS', os.getcwd())
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # Substitua 'app_evasao.py' pelo seu arquivo principal se for outro
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app_evasao.py"),
        "--global.developmentMode=false",
    ]
    sys.exit(stcli.main())