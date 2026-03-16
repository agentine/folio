/**
 * Convert a column number (1-based) to a letter string: 1 -> 'A', 27 -> 'AA'.
 */
export function colToLetter(col: number): string {
  let result = '';
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

/**
 * Convert a column letter string to a number: 'A' -> 1, 'AA' -> 27.
 */
export function letterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col;
}

/**
 * Parse a cell address like 'A1' into { col: 1, row: 1 }.
 */
export function parseAddress(address: string): { col: number; row: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  return {
    col: letterToCol(match[1].toUpperCase()),
    row: parseInt(match[2], 10),
  };
}

/**
 * Format a { col, row } into a cell address like 'A1'.
 */
export function formatAddress(col: number, row: number): string {
  return colToLetter(col) + row;
}

/**
 * Parse a range like 'A1:B2' into start and end addresses.
 */
export function parseRange(range: string): { start: { col: number; row: number }; end: { col: number; row: number } } {
  const parts = range.split(':');
  if (parts.length !== 2) {
    throw new Error(`Invalid range: ${range}`);
  }
  return {
    start: parseAddress(parts[0]),
    end: parseAddress(parts[1]),
  };
}
