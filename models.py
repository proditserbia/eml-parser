"""Data models for the EML parser application."""

from dataclasses import dataclass, field
from typing import List


@dataclass
class EmailMetadata:
    """Metadata extracted from an email header."""
    from_addr: str = ""
    to_addr: str = ""
    date: str = ""
    subject: str = ""


@dataclass
class VolumeUsage:
    """Storage volume usage information."""
    volume_name: str = ""
    free_space: str = ""
    used_percent: str = ""
    change_percent: str = ""


@dataclass
class ParsedReport:
    """Complete parsed report combining metadata, volume usage, and workspace lists."""
    from_addr: str = ""
    to_addr: str = ""
    date: str = ""
    subject: str = ""
    volume_name: str = ""
    free_space: str = ""
    used_percent: str = ""
    change_percent: str = ""
    active_workspaces: List[str] = field(default_factory=list)
    old_workspaces: List[str] = field(default_factory=list)

    @classmethod
    def from_parts(
        cls,
        metadata: EmailMetadata,
        volume: VolumeUsage,
        active_workspaces: List[str],
        old_workspaces: List[str],
    ) -> "ParsedReport":
        """Construct a ParsedReport from its component parts."""
        return cls(
            from_addr=metadata.from_addr,
            to_addr=metadata.to_addr,
            date=metadata.date,
            subject=metadata.subject,
            volume_name=volume.volume_name,
            free_space=volume.free_space,
            used_percent=volume.used_percent,
            change_percent=volume.change_percent,
            active_workspaces=active_workspaces,
            old_workspaces=old_workspaces,
        )
