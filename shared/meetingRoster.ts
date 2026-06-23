/**
 * Single source of truth for the H&S meeting attendee roster.
 *
 * The admin meeting-history signature view renders signatures keyed by these
 * exact names, and attendance rows store these exact strings. The Teams "Sign"
 * tab resolves the signed-in user to one of these entries so a self-signed
 * signature lines up with the admin view. Keep this list authoritative — both
 * the admin page and the server import it.
 */

export interface RosterMember {
  name: string;
  role: string;
}

export const meetingRoster: { management: RosterMember[]; glaziers: RosterMember[] } = {
  management: [
    { name: 'Hoani Hunt', role: 'Company Director' },
    { name: 'Simon Hubbard', role: 'Health & Safety Coordinator' },
    { name: 'James Waites', role: 'Glazing Supervisor' },
    { name: 'Emma White', role: 'Administrator' },
  ],
  glaziers: [
    { name: 'Kevin Young', role: 'Glazier' },
    { name: 'Ryan Newman', role: 'Glazier' },
    { name: 'Daniel Conlan', role: 'Glazier' },
    { name: "Struan O'Donnell", role: 'Glazier' },
    { name: 'Sam Chang', role: 'Glazier' },
  ],
};

export const allRosterMembers: RosterMember[] = [
  ...meetingRoster.management,
  ...meetingRoster.glaziers,
];

/**
 * Normalise a person's name for tolerant matching: strip diacritics, unify
 * apostrophe variants, lowercase, and collapse whitespace. This lets an Azure AD
 * display name like "Struan O’Donnell" match the roster's "Struan O'Donnell".
 */
function normaliseName(value: string): string {
  return value
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’`´]/g, "'")
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Resolve a roster member from one or more candidate names (e.g. a matched
 * staff record name, then the Microsoft Graph display name). Returns the first
 * roster entry that matches any candidate, or null if none match.
 */
export function findRosterMember(
  ...candidates: (string | null | undefined)[]
): RosterMember | null {
  for (const candidate of candidates) {
    if (!candidate) continue;
    const target = normaliseName(candidate);
    if (!target) continue;
    const hit = allRosterMembers.find((m) => normaliseName(m.name) === target);
    if (hit) return hit;
  }
  return null;
}
