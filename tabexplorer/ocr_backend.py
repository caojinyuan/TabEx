"""Qt-free RapidOCR process entry: keep ONNX and Qt native runtimes isolated."""

from pathlib import Path

from .debuglog import debug_print
from .i18n import _set_app_language, tr


MODEL_FILES = {
    'Det': 'PP-OCRv6_det_small.onnx',
    'Rec': 'PP-OCRv6_rec_small.onnx',
    'Cls': 'ch_ppocr_mobile_v2.0_cls_mobile.onnx',
}


class OcrError(RuntimeError):
    pass


def _load_engine():
    try:
        import rapidocr
        from rapidocr import RapidOCR
    except ImportError as error:
        raise OcrError(tr("缺少 RapidOCR 组件，请重新运行依赖安装脚本；EXE 用户请重新安装完整版本。")) from error
    models = Path(rapidocr.__file__).parent / 'models'
    params = {
        'Global.log_level': 'error',
        'Global.max_side_len': 4096,
        'EngineConfig.onnxruntime.intra_op_num_threads': 2,
        'EngineConfig.onnxruntime.inter_op_num_threads': 1,
    }
    for stage, name in MODEL_FILES.items():
        path = models / name
        if not path.is_file():
            raise OcrError(tr("缺少离线 OCR 模型：{}。请重新安装完整版本，不会自动联网下载。").format(name))
        # Explicit paths bypass RapidOCR's automatic model downloader.
        params[f'{stage}.model_path'] = str(path)
    return RapidOCR(params=params)


def _format_result(result):
    if result.boxes is None or result.txts is None:
        return ''
    cells = []
    for box, text in zip(result.boxes, result.txts):
        if not text.strip():
            continue
        left = min(float(point[0]) for point in box)
        top = min(float(point[1]) for point in box)
        bottom = max(float(point[1]) for point in box)
        cells.append((left, (top + bottom) / 2, max(1.0, bottom - top), text.strip()))
    rows = []
    for cell in sorted(cells, key=lambda value: value[1]):
        if rows and abs(cell[1] - rows[-1][0][1]) <= min(cell[2], rows[-1][0][2]) * 0.5:
            rows[-1].append(cell)
        else:
            rows.append([cell])
    return '\n'.join('\t'.join(cell[3] for cell in sorted(row)) for row in rows)


def recognize_process(image_bytes, connection, ui_language):
    _set_app_language(ui_language)
    try:
        import cv2
        import numpy as np
        engine = _load_engine()
        image = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
        if image is None:
            raise OcrError(tr("无法识别空图片。"))
        connection.send((_format_result(engine(image, use_cls=False)), ''))
    except Exception as error:
        debug_print(f"[OCR] inference failed: {type(error).__name__}: {error}")
        connection.send(('', str(error) if isinstance(error, OcrError)
                         else tr("文字识别失败：{}").format(error)))
    finally:
        connection.close()
