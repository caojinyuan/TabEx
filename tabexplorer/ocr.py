"""Bundled RapidOCR models, isolated CPU inference and cancellable Qt workers."""

import asyncio
import multiprocessing

from PyQt5.QtCore import QBuffer, QIODevice, QThread, Qt, pyqtSignal
from PyQt5.QtGui import QImage

from .debuglog import debug_print
from . import i18n
from .i18n import tr
from .ocr_backend import OcrError, recognize_process


OCR_TIMEOUT = 60
LANGUAGES = (('auto', '中英文自动识别'),)


def _prepare_image(image, maximum=4096):
    if image.isNull():
        raise OcrError(tr("无法识别空图片。"))
    if max(image.width(), image.height()) > maximum:
        image = image.scaled(maximum, maximum, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    return image.convertToFormat(QImage.Format_RGB888)


async def recognize_image(image, language='auto'):
    if language not in dict(LANGUAGES):
        raise OcrError(tr("不支持的识别语言：{}").format(language))
    image = _prepare_image(image)
    buffer = QBuffer()
    buffer.open(QIODevice.WriteOnly)
    if not image.save(buffer, 'PNG'):
        raise OcrError(tr("无法编码 OCR 图片。"))
    context = multiprocessing.get_context('spawn')
    receiver, sender = context.Pipe(duplex=False)
    process = context.Process(target=recognize_process,
                              args=(bytes(buffer.data()), sender, i18n._app_language), daemon=True)
    try:
        process.start()
        sender.close()
        try:
            while not receiver.poll():
                if not process.is_alive():
                    raise OcrError(tr("OCR 进程异常退出，请重试。"))
                await asyncio.sleep(0.05)
            text, error = receiver.recv()
        except (EOFError, BrokenPipeError) as error:
            debug_print(f"[OCR] child exited with code {process.exitcode}")
            raise OcrError(tr("OCR 进程异常退出，请重试。") + f" ({process.exitcode})") from error
        if error:
            raise OcrError(error)
        return text
    finally:
        if process.pid is not None:
            if process.is_alive():
                process.terminate()
            process.join()
            process.close()
        receiver.close()
        sender.close()


class OcrWorker(QThread):
    completed = pyqtSignal(str, str)  # text, error

    def __init__(self, image, language, parent=None):
        super().__init__(parent)
        self._image = image.copy()
        self._language = language

    async def _recognize(self):
        task = asyncio.create_task(recognize_image(self._image, self._language))
        try:
            while not task.done():
                if self.isInterruptionRequested():
                    raise asyncio.CancelledError
                await asyncio.wait((task,), timeout=0.05)
            return await task
        finally:
            if not task.done():
                task.cancel()
                try:
                    await task
                except asyncio.CancelledError:
                    pass

    def run(self):
        try:
            text = asyncio.run(asyncio.wait_for(self._recognize(), OCR_TIMEOUT))
            if not self.isInterruptionRequested():
                self.completed.emit(text, '')
        except asyncio.CancelledError:
            return
        except TimeoutError:
            self.completed.emit('', tr("文字识别超时，请缩小选区后重试。"))
        except Exception as error:
            debug_print(f"[OCR] recognition failed: {type(error).__name__}: {error}")
            self.completed.emit('', str(error) if isinstance(error, OcrError)
                                else tr("文字识别失败：{}").format(error))
