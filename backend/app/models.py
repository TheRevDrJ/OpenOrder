from pydantic import AliasChoices, BaseModel, ConfigDict, Field


class HymnRef(BaseModel):
    number: str
    title: str
    source: str


class StaffMember(BaseModel):
    name: str
    title: str


class OrderOfWorship(BaseModel):
    # Lets the field be set by its own name as well as by an alias below.
    model_config = ConfigDict(populate_by_name=True)

    date: str  # ISO 8601: "2026-03-22"
    serviceTitle: str = ""

    # The ONE picture for this ONE service. Named "hero" and not "theme" because a
    # THEME here is the palette and typography a whole church uses (app/churches.py) —
    # two different things wore one word and every read cost a lookup.
    #
    # ⛔ READS EITHER KEY, WRITES ONLY THE NEW ONE. Services saved before 2026-09-24
    # hold `themeImageFilename` on disk, and a bare rename would make them stop
    # finding their image SILENTLY — a blank slide, not an error. The alias is what
    # makes the rename cost the user nothing, so do not "tidy" it away; it has to
    # outlive every saved service that predates it.
    heroImageFilename: str | None = Field(
        default=None,
        validation_alias=AliasChoices("heroImageFilename", "themeImageFilename"),
    )

    # ⛔ THERE IS NO `speaker` OR `worshipLeader` STRUCT, AND THAT IS ON PURPOSE
    # (BUG-004 and BUG-005, 2026-09-24). Both were collected by the form for months
    # and rendered nowhere. The ONE speaker value that reaches a bulletin is
    # `speakerShortName` below, which fills `{{SPEAKER}}`; everything else about a
    # church's staff is typed into its Word template, which owns those lines.
    # ⚠ Old saved services still carry both keys; pydantic ignores them and they
    # drop on the next save.

    # Hymns
    praiseHymn1: HymnRef | None = None
    praiseHymn2: HymnRef | None = None
    doxology: HymnRef | None = None
    creed: HymnRef | None = None
    prayerHymn: HymnRef | None = None
    liturgicalPrayer: HymnRef | None = HymnRef(
        number="895", title="The Lord's Prayer Former Methodist Text", source="umh-services"
    )
    closingHymn: HymnRef | None = None

    # Whether this Sunday includes Holy Communion. Adds one slide between the
    # Word and the Sending; false for most Sundays.
    communion: bool = False

    # Sermon
    scripture: str = ""
    scriptureTranslation: str = "BSB"
    sermonTitle: str = ""
    sermonSubtitle: str = ""
    speakerShortName: str = ""
