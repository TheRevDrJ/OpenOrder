export interface HymnRef {
  number: string
  title: string
  source: string
}

export interface HymnSearchResult {
  number: string
  title: string
  source: string
  slide_count: number
  file: string
}

export interface StaffMember {
  name: string
  title: string
}

export interface OrderOfWorship {
  date: string
  serviceTitle: string
  themeImageFilename: string | null

  speaker: StaffMember
  worshipLeader: StaffMember

  praiseHymn1: HymnRef | null
  praiseHymn2: HymnRef | null
  doxology: HymnRef | null
  creed: HymnRef | null
  prayerHymn: HymnRef | null
  closingHymn: HymnRef | null

  scripture: string
  sermonTitle: string
  sermonSubtitle: string
  speakerShortName: string
  offertoryNote: string
}

export function emptyOrder(date: string): OrderOfWorship {
  return {
    date,
    serviceTitle: '',
    themeImageFilename: null,
    speaker: { name: 'Rev. Dr. Jonathan Mellette', title: 'Lead Pastor' },
    worshipLeader: { name: 'Heather Davis', title: 'Worship Leader' },
    praiseHymn1: null,
    praiseHymn2: null,
    doxology: null,
    creed: null,
    prayerHymn: null,
    closingHymn: null,
    scripture: '',
    sermonTitle: '',
    sermonSubtitle: '',
    speakerShortName: 'Dr. Mellette',
    offertoryNote: '',
  }
}
