from pydantic import BaseModel

from . import church_config as cfg


class HymnRef(BaseModel):
    number: str
    title: str
    source: str


class StaffMember(BaseModel):
    name: str
    title: str


class OrderOfWorship(BaseModel):
    date: str  # ISO 8601: "2026-03-22"
    serviceTitle: str = ""
    themeImageFilename: str | None = None

    # Staff
    speaker: StaffMember = StaffMember(
        name=cfg.DEFAULT_SPEAKER_NAME, title=cfg.DEFAULT_SPEAKER_TITLE
    )
    worshipLeader: StaffMember = StaffMember(
        name=cfg.DEFAULT_WORSHIP_LEADER_NAME, title=cfg.DEFAULT_WORSHIP_LEADER_TITLE
    )

    # Hymns
    praiseHymn1: HymnRef | None = None
    praiseHymn2: HymnRef | None = None
    doxology: HymnRef | None = None
    creed: HymnRef | None = None
    prayerHymn: HymnRef | None = None
    closingHymn: HymnRef | None = None

    # Sermon
    scripture: str = ""
    sermonTitle: str = ""
    sermonSubtitle: str = ""
    speakerShortName: str = cfg.DEFAULT_SPEAKER_SHORT
