export const Rfc4180 = {
  // Records split on newlines and fields on the delimiter, with RFC 4180 quoting.
  records(data: string, delimiter: string): string[][] {
    const records: string[][] = [];
    let record: string[] = [];
    let field = "";
    let isQuoted = false;
    for (let i = 0; i < data.length; i++) {
      const char = data[i];
      if (isQuoted) {
        if (char !== '"') {
          field += char;
        } else if (data[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          isQuoted = false;
        }
        continue;
      }
      if (char === '"' && field === "") {
        isQuoted = true;
        continue;
      }
      if (data.startsWith(delimiter, i)) {
        record.push(field);
        field = "";
        i += delimiter.length - 1;
        continue;
      }
      if (char === "\n") {
        records.push([...record, field]);
        record = [];
        field = "";
        continue;
      }
      field += char;
    }
    records.push([...record, field]);
    return records;
  },
};
