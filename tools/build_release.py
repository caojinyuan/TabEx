"""Stage, validate and atomically promote a Windows release."""

import argparse
import os
from pathlib import Path
import subprocess
import sys
import tempfile

if __package__:
    from .validate_runtime import validate
else:
    from validate_runtime import validate


ROOT = Path(__file__).resolve().parents[1]


def build(root=ROOT, install=True, output=None, run=None, verify=None):
    root = Path(root).resolve()
    output = Path(output).resolve() if output else root / 'TabExplorer.exe'
    run = run or subprocess.run
    verify = verify or validate
    python = [sys.executable, '-X', 'utf8']
    if install:
        run(python + ['-m', 'pip', 'install', '-r', str(root / 'requirements-build.txt')], cwd=root, check=True)
    run(python + ['-W', 'ignore::DeprecationWarning', '-m', 'unittest', 'discover', '-s', 'tests'],
        cwd=root, check=True)
    with tempfile.TemporaryDirectory(prefix='.tabex-build-', dir=root) as temporary:
        stage = Path(temporary)
        command = python + [
            '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed',
            '--name', 'TabExplorer', '--distpath', str(stage / 'dist'),
            '--workpath', str(stage / 'work'), '--specpath', str(stage),
            '--add-data', str(root / 'icons') + ';icons',
        ]
        icon = root / 'icons' / 'TabExplorer.ico'
        if icon.exists():
            command.extend(['--icon', str(icon)])
        run(command + [str(root / 'TabEx.py')], cwd=root, check=True)
        candidate = stage / 'dist' / 'TabExplorer.exe'
        if not candidate.is_file():
            raise FileNotFoundError(candidate)
        report = root / 'artifacts' / 'release-runtime.json'
        result = verify([str(candidate)], report, seconds=10, tabs=10)
        if not result.get('ok') or not result.get('frozen'):
            raise RuntimeError('Packaged runtime validation did not pass')
        output.parent.mkdir(parents=True, exist_ok=True)
        os.replace(candidate, output)
    return output


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skip-install', action='store_true')
    parser.add_argument('--output', type=Path)
    options = parser.parse_args()
    output = build(install=not options.skip_install, output=options.output)
    print(f'Validated release: {output}')


if __name__ == '__main__':
    main()