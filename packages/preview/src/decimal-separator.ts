const separators = new Map<string, string>();

/**
 * The character a decimal tab stop aligns on for a run's language
 * (`<a:rPr lang>`): `,` for `de-DE`, `.` for `en-US`. Runs without a
 * language, or with one Intl does not know, align on `.`. Mac PowerPoint
 * uses the run language too, not the UI or system locale (native capture
 * 2026-10-07: Set Proofing Language to German moved alignment to `,`).
 */
export function decimalSeparatorOf(lang: string | null | undefined): string {
  if (!lang) return '.';
  let separator = separators.get(lang);
  if (separator === undefined) {
    separator = '.';
    try {
      separator =
        // Number text in decks is typed with Latin digits; native-digit
        // locales such as ar-EG would otherwise report `٫`.
        new Intl.NumberFormat(lang, { numberingSystem: 'latn' })
          .formatToParts(1.5)
          .find((part) => part.type === 'decimal')?.value ?? '.';
    } catch (error) {
      // Decks carry tags Intl rejects, such as Office's `x-none`.
      if (!(error instanceof RangeError)) throw error;
    }
    separators.set(lang, separator);
  }
  return separator;
}
