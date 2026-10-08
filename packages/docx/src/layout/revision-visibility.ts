/** ECMA-376 §17.13.5.14/§17.13.5.22: final view omits deleted and
 * moved-away occurrences. All acquisition and retained-coordinate paths use
 * the same visibility decision; markup view retains their original ownership. */
export function revisionIsOmitted(kind: string | undefined, showTrackedChanges: boolean | undefined): boolean {
  return showTrackedChanges !== true && (kind === 'deletion' || kind === 'moveFrom');
}
