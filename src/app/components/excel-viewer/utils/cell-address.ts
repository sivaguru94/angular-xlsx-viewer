/**
 * Converts a 0-based column index to a letter (e.g., 0 → "A", 25 → "Z", 26 → "AA").
 */
export function columnToLetter(col: number): string {
  let letter = '';
  let temp = col;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Converts a column letter to a 0-based index (e.g., "A" → 0, "Z" → 25, "AA" → 26).
 */
export function letterToColumn(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Parses a cell address string (e.g., "A1", "BC42") into 0-based row/col indices.
 */
export function parseAddress(address: string): { row: number; col: number } {
  const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  return {
    col: letterToColumn(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}
