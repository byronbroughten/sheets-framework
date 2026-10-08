// Every Table's layout and ID conventions; head-row offsets count up from the header row, base 0.
export const tableLayout = {
  idDelimiter: ":",
  idHeader: "ID",
  nameHeader: "Name",
  headerZoneDepth: 4,
  headRowOffsets: {
    header: 0,
    action: 1,
    groupHeading2: 1,
    groupHeading1: 2,
    columnId: 3,
  },
} as const;
