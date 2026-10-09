"""Run the prototype, compile the short thesis, and check manuscript formatting."""
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def main():
    commands = [
        [sys.executable, "-X", "utf8", "prototype.py"],
        [sys.executable, "-X", "utf8", "scripts/build.py"],
        [sys.executable, "-X", "utf8", "scripts/check_format.py", "--document"],
    ]
    for command in commands:
        result = subprocess.run(command, cwd=ROOT)
        if result.returncode:
            return result.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
