"""Run Pywrangler from the project environment on Windows, Linux, or CI."""
from pathlib import Path
import shutil
import subprocess
import sys


ROOT=Path(__file__).resolve().parent.parent


def executable():
    candidates=(ROOT/'.venv'/'Scripts'/'pywrangler.exe',ROOT/'.venv'/'bin'/'pywrangler')
    for candidate in candidates:
        if candidate.is_file():
            return [str(candidate)]
    installed=shutil.which('pywrangler')
    if installed:
        return [installed]
    uv=shutil.which('uv')
    if uv:
        return [uv,'run','pywrangler']
    raise SystemExit('Pywrangler is unavailable. Run `uv sync` before building or deploying.')


if __name__=='__main__':
    raise SystemExit(subprocess.call([*executable(),*sys.argv[1:]],cwd=ROOT))
