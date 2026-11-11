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
MSVC_DLLS = (
    'msvcp140.dll', 'msvcp140_1.dll', 'msvcp140_2.dll',
    'msvcp140_atomic_wait.dll', 'msvcp140_codecvt_ids.dll',
    'vcruntime140.dll', 'vcruntime140_1.dll', 'concrt140.dll',
)


def collect_msvc_runtime(directory=None):
    import pefile

    directory = Path(directory) if directory else Path(os.environ['SystemRoot']) / 'System32'
    binaries = []
    for name in MSVC_DLLS:
        path = directory / name
        if not path.is_file():
            raise RuntimeError(f'Missing {path}. Install Microsoft Visual C++ 2015-2022 '
                               'Redistributable x64 (14.42 or newer) before building.')
        with pefile.PE(str(path)) as image:
            versions = getattr(image, 'VS_FIXEDFILEINFO', [])
            version = ((versions[0].FileVersionMS >> 16, versions[0].FileVersionMS & 0xffff)
                       if versions else (0, 0))
            if image.FILE_HEADER.Machine != 0x8664 or version < (14, 42):
                raise RuntimeError(f'{path} must be an x64 Visual C++ runtime DLL, version 14.42 '
                                   'or newer. Update Microsoft Visual C++ Redistributable x64.')
        # Explicit inputs take precedence over Qt's older DLLs, including its private copies.
        binaries.extend((path, destination) for destination in ('.', r'PyQt5\Qt5\bin'))
    return binaries


def build(root=ROOT, install=True, output=None, run=None, verify=None):
    root = Path(root).resolve()
    output = Path(output).resolve() if output else root / 'TabExplorer.exe'
    run = run or subprocess.run
    verify = verify or validate
    python = [sys.executable, '-X', 'utf8']
    if install:
        run(python + ['-m', 'pip', 'install', '-r', str(root / 'requirements' / 'requirements-build.txt')],
            cwd=root, check=True)
    run(python + ['-W', 'ignore::DeprecationWarning', '-m', 'unittest', 'discover', '-s', 'tests'],
        cwd=root, check=True)
    bundled_tools = [root / 'tools' / name for name in ('es.exe', 'rg.exe', '7zr.exe')]
    for tool in bundled_tools:
        if not tool.is_file():
            raise FileNotFoundError(f'Bundled tool not found: {tool}')
    runtime_binaries = collect_msvc_runtime()
    with tempfile.TemporaryDirectory(prefix='.tabex-build-', dir=root) as temporary:
        stage = Path(temporary)
        command = python + [
            '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed',
            '--name', 'TabExplorer', '--distpath', str(stage / 'dist'),
            '--workpath', str(stage / 'work'), '--specpath', str(stage),
            '--add-data', str(root / 'icons') + ';icons',
            '--collect-data', 'rapidocr',
            '--collect-submodules', 'rapidocr.inference_engine.onnxruntime',
            '--runtime-hook', str(root / 'tools' / 'pyi_rth_ocr.py'),
        ]
        command.extend(arg for source, destination in runtime_binaries
                       for arg in ('--add-binary', str(source) + ';' + destination))
        command.extend(arg for tool in bundled_tools for arg in ('--add-data', str(tool) + ';tools'))
        icon = root / 'icons' / 'TabExplorer.ico'
        if icon.exists():
            command.extend(['--icon', str(icon)])
        run(command + [str(root / 'TabEx.py')], cwd=root, check=True)
        candidate = stage / 'dist' / 'TabExplorer.exe'
        if not candidate.is_file():
            raise FileNotFoundError(candidate)
        report = root / 'artifacts' / 'release-runtime.json'
        result = verify([str(candidate)], report, seconds=10, tabs=10)
        if not result.get('ok') or not result.get('frozen') or not result.get('ocr', {}).get('ok'):
            raise RuntimeError('Packaged runtime/OCR validation did not pass')
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