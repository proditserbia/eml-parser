"""EML file parser for ELEMENTS Info: Status projekata emails."""

import email
import email.message
import logging
import re
from email.header import decode_header
from html.parser import HTMLParser
from pathlib import Path
from typing import List, Optional, Tuple

from models import EmailMetadata, ParsedReport, VolumeUsage

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# HTML → plain-text fallback
# ---------------------------------------------------------------------------

class _HtmlToText(HTMLParser):
    """Minimal HTML-to-plain-text extractor (no external dependencies)."""

    def __init__(self) -> None:
        super().__init__()
        self._parts: List[str] = []
        self._skip_tags = {"script", "style", "head"}
        self._current_skip = 0

    def handle_starttag(self, tag: str, attrs) -> None:  # type: ignore[override]
        if tag.lower() in self._skip_tags:
            self._current_skip += 1
        if tag.lower() in ("br", "p", "div", "li", "tr"):
            self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:  # type: ignore[override]
        if tag.lower() in self._skip_tags:
            self._current_skip -= 1

    def handle_data(self, data: str) -> None:
        if self._current_skip == 0:
            self._parts.append(data)

    def get_text(self) -> str:
        return "".join(self._parts)


def _html_to_text(html: str) -> str:
    """Convert HTML to plain text using the built-in HTMLParser."""
    parser = _HtmlToText()
    try:
        parser.feed(html)
    except Exception:
        # Malformed HTML – return whatever was collected so far.
        pass
    return parser.get_text()


# ---------------------------------------------------------------------------
# MIME header decoder
# ---------------------------------------------------------------------------

def _decode_mime_header(raw: Optional[str]) -> str:
    """Decode a MIME-encoded header value to a plain Unicode string."""
    if not raw:
        return ""
    parts: List[str] = []
    for chunk, charset in decode_header(raw):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(chunk)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Text body extraction from MIME message
# ---------------------------------------------------------------------------

def _get_body(msg: email.message.Message) -> Tuple[str, str]:
    """
    Return *(plain_text, source)* where *source* is ``"plain"`` or ``"html"``.

    Tries ``text/plain`` first; falls back to ``text/html``.
    """
    plain_text: Optional[str] = None
    html_text: Optional[str] = None

    for part in msg.walk():
        ctype = part.get_content_type()
        if part.get_content_disposition() == "attachment":
            continue
        charset = part.get_content_charset() or "utf-8"
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        decoded = payload.decode(charset, errors="replace")
        if ctype == "text/plain" and plain_text is None:
            plain_text = decoded
        elif ctype == "text/html" and html_text is None:
            html_text = decoded

    if plain_text:
        return plain_text, "plain"
    if html_text:
        logger.info("Plain text part not found; falling back to HTML extraction.")
        return _html_to_text(html_text), "html"
    return "", "none"


# ---------------------------------------------------------------------------
# Section & item extraction helpers
# ---------------------------------------------------------------------------

# Patterns for the Volume Usage line, e.g.:
#   BeeGFS: 92.6 TB free (73% used, changed by -4%)
_VOLUME_PATTERN = re.compile(
    r"(?P<name>[^:]+):\s*"
    r"(?P<free>[0-9.,]+\s*[TGMK]?B?)\s*free\s*"
    r"\((?P<used>[0-9.,]+%)\s*used,\s*changed by\s*(?P<change>[+-]?[0-9.,]+%)\)",
    re.IGNORECASE,
)

_SECTION_HEADERS = {
    "volume_usage": re.compile(
        r"^[\s\-*•#]*volume\s+usage[\s\-:]*$", re.IGNORECASE
    ),
    "active": re.compile(
        r"^[\s\-*•#]*active\s+workspaces?[\s\-:]*$", re.IGNORECASE
    ),
    "old": re.compile(
        r"^[\s\-*•#]*workspaces?\s+with\s+last\s+login\s+older\s+than\s+90\s+days?[\s\-:]*$",
        re.IGNORECASE,
    ),
}

# A workspace item line: starts with a bullet, dash, asterisk, or is a
# non-empty line that does *not* look like a section header.
_BULLET_RE = re.compile(r"^[\s\-*•·–]+(.+)$")


def _normalize(text: str) -> str:
    """Collapse runs of whitespace and strip."""
    return re.sub(r"\s+", " ", text).strip()


def _classify_line(line: str) -> Optional[str]:
    """Return the section key if *line* is a section header, else None."""
    stripped = line.strip()
    for key, pattern in _SECTION_HEADERS.items():
        if pattern.match(stripped):
            return key
    return None


def _extract_workspace_items(lines: List[str]) -> List[str]:
    """
    Pull workspace names from a list of raw text lines.

    Accepts both bullet-prefixed lines and plain non-empty lines.
    Stops when a new section header is detected.
    """
    items: List[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if _classify_line(stripped) is not None:
            # Hit a new section – stop collecting.
            break
        m = _BULLET_RE.match(stripped)
        if m:
            items.append(_normalize(m.group(1)))
        else:
            items.append(_normalize(stripped))
    return items


def _parse_volume_line(line: str) -> Optional[VolumeUsage]:
    """Try to parse a Volume Usage line; return None if it does not match."""
    m = _VOLUME_PATTERN.search(line)
    if not m:
        return None
    return VolumeUsage(
        volume_name=_normalize(m.group("name")),
        free_space=_normalize(m.group("free")),
        used_percent=_normalize(m.group("used")),
        change_percent=_normalize(m.group("change")),
    )


# ---------------------------------------------------------------------------
# Main parser class
# ---------------------------------------------------------------------------

class EmlParser:
    """Parse ELEMENTS Info status emails from .eml files."""

    def __init__(self) -> None:
        self._last_source: str = "none"

    @property
    def last_source(self) -> str:
        """The content source used in the last parse: ``"plain"``, ``"html"``, or ``"none"``."""
        return self._last_source

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def parse_file(self, path: str | Path) -> ParsedReport:
        """
        Parse an .eml file and return a :class:`ParsedReport`.

        Raises:
            FileNotFoundError: If the file does not exist.
            ValueError: If the file cannot be parsed as an email.
        """
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"EML датотека није пронађена: {path}")

        raw = path.read_bytes()
        try:
            msg = email.message_from_bytes(raw)
        except Exception as exc:
            raise ValueError(f"Грешка при читању EML датотеке: {exc}") from exc

        return self._parse_message(msg)

    def parse_bytes(self, raw: bytes) -> ParsedReport:
        """Parse raw .eml bytes and return a :class:`ParsedReport`."""
        try:
            msg = email.message_from_bytes(raw)
        except Exception as exc:
            raise ValueError(f"Грешка при читању EML датотеке: {exc}") from exc
        return self._parse_message(msg)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _parse_message(self, msg: email.message.Message) -> ParsedReport:
        metadata = self._extract_metadata(msg)
        body, source = _get_body(msg)
        self._last_source = source

        if not body:
            logger.warning("Тело поруке је празно; враћамо празан извештај.")
            return ParsedReport.from_parts(metadata, VolumeUsage(), [], [])

        volume, active, old = self._parse_body(body)
        return ParsedReport.from_parts(metadata, volume, active, old)

    def _extract_metadata(self, msg: email.message.Message) -> EmailMetadata:
        return EmailMetadata(
            from_addr=_decode_mime_header(msg.get("From", "")),
            to_addr=_decode_mime_header(msg.get("To", "")),
            date=_decode_mime_header(msg.get("Date", "")),
            subject=_decode_mime_header(msg.get("Subject", "")),
        )

    def _parse_body(
        self, body: str
    ) -> Tuple[VolumeUsage, List[str], List[str]]:
        """
        Split the email body into sections and extract structured data.

        Returns *(volume_usage, active_workspaces, old_workspaces)*.
        """
        lines = body.splitlines()

        volume = VolumeUsage()
        active_workspaces: List[str] = []
        old_workspaces: List[str] = []

        current_section: Optional[str] = None
        section_lines: List[str] = []

        def _flush(section: str, collected: List[str]) -> None:
            nonlocal volume, active_workspaces, old_workspaces
            if section == "volume_usage":
                for ln in collected:
                    v = _parse_volume_line(ln)
                    if v:
                        volume = v
                        break
            elif section == "active":
                active_workspaces = _extract_workspace_items(collected)
            elif section == "old":
                old_workspaces = _extract_workspace_items(collected)

        for line in lines:
            key = _classify_line(line)
            if key is not None:
                # Flush the previous section before starting a new one.
                if current_section is not None:
                    _flush(current_section, section_lines)
                current_section = key
                section_lines = []
            elif current_section is not None:
                section_lines.append(line)

        # Flush the last open section.
        if current_section is not None:
            _flush(current_section, section_lines)

        return volume, active_workspaces, old_workspaces


# ---------------------------------------------------------------------------
# Simple self-test helper (run directly: python parser.py)
# ---------------------------------------------------------------------------

def _self_test() -> None:
    """Quick smoke-test using a synthetic email body."""
    sample_eml = b"""\
From: noreply@elements.local
To: admin@example.com
Date: Mon, 01 Apr 2024 10:00:00 +0200
Subject: ELEMENTS Info: Status projekata
Content-Type: text/plain; charset=utf-8

## Volume Usage
BeeGFS: 92.6 TB free (73% used, changed by -4%)

## Active Workspaces
- ProjectAlpha
- ProjectBeta
- ProjectGamma

## Workspaces with last login older than 90 days
- OldProject1
- OldProject2
"""
    p = EmlParser()
    report = p.parse_bytes(sample_eml)

    assert report.from_addr == "noreply@elements.local", repr(report.from_addr)
    assert report.volume_name == "BeeGFS", repr(report.volume_name)
    assert report.free_space == "92.6 TB", repr(report.free_space)
    assert report.used_percent == "73%", repr(report.used_percent)
    assert report.change_percent == "-4%", repr(report.change_percent)
    assert "ProjectAlpha" in report.active_workspaces, report.active_workspaces
    assert "OldProject1" in report.old_workspaces, report.old_workspaces

    print("Самотест прошао успешно ✓")
    print(f"  Волумен: {report.volume_name} — {report.free_space} слободно "
          f"({report.used_percent} искоришћено, промена {report.change_percent})")
    print(f"  Активни радни простори: {report.active_workspaces}")
    print(f"  Стари радни простори: {report.old_workspaces}")


if __name__ == "__main__":
    _self_test()
