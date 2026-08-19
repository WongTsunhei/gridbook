export function parseSpreadsheetClipboard(input: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < input.length; index += 1) {
    const character = input[index];
    if (quoted) {
      if (character === '"' && input[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        cell += character;
      }
      continue;
    }
    if (character === '"' && cell === "") {
      quoted = true;
    } else if (character === "\t") {
      row.push(cell);
      cell = "";
    } else if (character === "\r" && input[index + 1] === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      index += 1;
    } else if (character === "\n" || character === "\r") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += character;
    }
  }
  row.push(cell);
  rows.push(row);

  if (rows.length > 1 && rows.at(-1)?.length === 1 && rows.at(-1)?.[0] === "") {
    rows.pop();
  }
  return rows.length ? rows : [[""]];
}

/** Read the table fragment placed on the clipboard by Excel and WPS. */
export function parseHtmlSpreadsheetClipboard(input: string): string[][] {
  const document = new DOMParser().parseFromString(input, "text/html");
  const rows = Array.from(document.querySelectorAll("table tr"));
  if (!rows.length) return [];
  return rows.map((row) => Array.from(row.querySelectorAll("th, td"), (cell) =>
    (cell.textContent ?? "").replaceAll("\u00a0", " ").replaceAll("\r\n", "\n"),
  ));
}

export function serializeSpreadsheetClipboard(rows: unknown[][]): string {
  const escape = (value: unknown) => {
    const text = value === null || value === undefined ? "" : String(value);
    return /[\t\r\n"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  return rows.map((row) => row.map(escape).join("\t")).join("\r\n");
}
