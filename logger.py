"""
Logger utilities
"""

import logging
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtCore import QDateTime
from PyQt5.QtGui import QTextCursor, QColor

# Custom log level for unknown/unexpected commands
UNKNOWN_CMD = 45  # above ERROR (40)
logging.addLevelName(UNKNOWN_CMD, 'UNKNOWN_CMD')


def unknown_cmd(self, message, *args, **kwargs):
    if self.isEnabledFor(UNKNOWN_CMD):
        self._log(UNKNOWN_CMD, message, args, **kwargs)


logging.Logger.unknown_cmd = unknown_cmd



class ConsoleHandler(logging.Handler):
    """Custom logging handler that writes to a QTextEdit widget"""
    
    def __init__(self, text_widget: QTextEdit):
        super().__init__()
        self.text_widget = text_widget
        self.visible_levels = {logging.DEBUG, logging.INFO, logging.WARNING}

    def set_level_visible(self, level: int, visible: bool):
        """Show or hide messages of the given level in the console."""
        if visible:
            self.visible_levels.add(level)
        else:
            self.visible_levels.discard(level)

    def emit(self, record):
        """Emit a log record to the text widget"""
        try:
            if record.levelno in (logging.DEBUG, logging.INFO, logging.WARNING) \
                    and record.levelno not in self.visible_levels:
                return
            msg = self.format(record)
            # Add timestamp
            timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
            formatted_msg = f"[{timestamp}] {msg}\n"
            
            # Append to text widget
            self.text_widget.moveCursor(QTextCursor.End)
            
            # Color based on level
            if record.levelno == UNKNOWN_CMD:
                self.text_widget.moveCursor(QTextCursor.End)
                import html
                safe_msg = html.escape(formatted_msg).replace('\n', '<br/>')
                self.text_widget.insertHtml(
                    f'<span style="color:red; font-weight:bold;">{safe_msg}</span>'
                )
                self.text_widget.setTextColor(QColor(0, 0, 0))
            else:
                if record.levelno >= logging.ERROR:
                    self.text_widget.setTextColor(QColor(255, 0, 0))
                elif record.levelno >= logging.WARNING:
                    self.text_widget.setTextColor(QColor(255, 165, 0))
                elif record.levelno == logging.DEBUG:
                    self.text_widget.setTextColor(QColor(128, 128, 128))
                else:
                    self.text_widget.setTextColor(QColor(0, 0, 0))

                self.text_widget.insertPlainText(formatted_msg)
                self.text_widget.setTextColor(QColor(0, 0, 0))
            
            # Keep only last 1000 lines
            doc = self.text_widget.document()
            if doc.blockCount() > 1000:
                cursor = QTextCursor(doc)
                cursor.movePosition(QTextCursor.Start)
                cursor.select(QTextCursor.BlockUnderCursor)
                cursor.removeSelectedText()
        except Exception:
            self.handleError(record)


class FilteringStreamHandler(logging.StreamHandler):
    """StreamHandler that can independently show/hide WARNING, INFO and DEBUG messages."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.visible_levels = {logging.DEBUG, logging.INFO, logging.WARNING}

    def set_level_visible(self, level: int, visible: bool):
        """Show or hide messages of the given level."""
        if visible:
            self.visible_levels.add(level)
        else:
            self.visible_levels.discard(level)

    def emit(self, record):
        if record.levelno in (logging.DEBUG, logging.INFO, logging.WARNING) \
                and record.levelno not in self.visible_levels:
            return
        super().emit(record)


class FilteringFileHandler(logging.FileHandler):
    """FileHandler that can independently show/hide WARNING, INFO and DEBUG messages."""

    def __init__(self, filename, *args, **kwargs):
        super().__init__(filename, *args, **kwargs)
        self.visible_levels = {logging.DEBUG, logging.INFO, logging.WARNING}

    def set_level_visible(self, level: int, visible: bool):
        """Show or hide messages of the given level."""
        if visible:
            self.visible_levels.add(level)
        else:
            self.visible_levels.discard(level)

    def emit(self, record):
        if record.levelno in (logging.DEBUG, logging.INFO, logging.WARNING) \
                and record.levelno not in self.visible_levels:
            return
        super().emit(record)


def setup_logger(name):
    """Setup logger with console handler"""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Skip adding a StreamHandler if the root logger already has one
    # (i.e. the main window has set up a shared FilteringStreamHandler).
    root_has_stream = any(
        isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler)
        for h in logging.getLogger().handlers
    )

    if not logger.handlers and not root_has_stream:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(levelname)s - %(name)s - %(message)s')
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger
