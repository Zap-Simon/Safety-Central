/**
 * Centralized date parsing utilities for handling various SharePoint date formats
 */

/**
 * Parse various date formats from SharePoint and return a consistent Date object
 * Handles:
 * - ISO 8601: "2025-07-08T12:00:00Z"
 * - ISO with milliseconds: "2025-07-08T12:00:00.000Z"
 * - SharePoint REST: "2025-07-08T11:00:00Z" (different timezone offsets)
 * - Date only: "2025-07-08"
 * - Invalid/null dates
 */
export function parseSharePointDate(dateInput: string | null | undefined): Date | null {
  if (!dateInput || dateInput === 'unknown-meeting') {
    return null;
  }

  try {
    const date = new Date(dateInput);
    
    // Check if date is valid
    if (isNaN(date.getTime())) {
      console.warn(`Invalid date format: ${dateInput}`);
      return null;
    }
    
    return date;
  } catch (error) {
    console.error(`Error parsing date: ${dateInput}`, error);
    return null;
  }
}

/**
 * Format date for display in UI
 * @param date - Date object or string
 * @param format - 'full' | 'date-only' | 'meeting'
 */
export function formatDisplayDate(date: Date | string | null, format: 'full' | 'date-only' | 'meeting' = 'full'): string {
  if (!date) return 'No Date';
  
  const parsedDate = typeof date === 'string' ? parseSharePointDate(date) : date;
  if (!parsedDate) return 'Invalid Date';

  // For meeting date headers, use UTC date components to avoid timezone conversion issues
  if (format === 'meeting') {
    const year = parsedDate.getUTCFullYear();
    const month = parsedDate.getUTCMonth();
    const day = parsedDate.getUTCDate();
    const dayOfWeek = new Date(year, month, day).toLocaleDateString('en-NZ', { weekday: 'long' });
    const monthName = new Date(year, month, day).toLocaleDateString('en-NZ', { month: 'long' });
    
    return `${dayOfWeek}, ${day} ${monthName} ${year}`;
  }

  const options: Intl.DateTimeFormatOptions = {
    timeZone: 'Pacific/Auckland', // New Zealand timezone for non-meeting formats
    weekday: format === 'meeting' ? 'long' : undefined,
    year: 'numeric',
    month: format === 'date-only' ? '2-digit' : 'long',
    day: 'numeric'
  };

  if (format === 'full') {
    options.hour = '2-digit';
    options.minute = '2-digit';
  }

  return parsedDate.toLocaleDateString('en-NZ', options);
}

/**
 * Get a normalized date key for grouping items by date only (ignoring time)
 * This ensures all items from the same day are grouped together
 */
export function getDateGroupKey(dateInput: string | Date | null): string {
  if (!dateInput) return 'unknown-meeting';
  
  const date = typeof dateInput === 'string' ? parseSharePointDate(dateInput) : dateInput;
  if (!date) return 'unknown-meeting';

  // Use UTC date components to match the display formatting and avoid timezone issues
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Compare two dates to see if they're on the same day (ignoring time)
 */
export function isSameDay(date1: string | Date | null, date2: string | Date | null): boolean {
  const parsed1 = typeof date1 === 'string' ? parseSharePointDate(date1) : date1;
  const parsed2 = typeof date2 === 'string' ? parseSharePointDate(date2) : date2;
  
  if (!parsed1 || !parsed2) return false;
  
  return parsed1.toDateString() === parsed2.toDateString();
}

/**
 * Calculate the next Tuesday from a given date
 */
export function getNextTuesday(fromDate: Date): Date {
  const date = new Date(fromDate);
  const dayOfWeek = date.getDay();
  
  let daysUntilTuesday;
  if (dayOfWeek <= 2) { // Sunday, Monday, or Tuesday
    daysUntilTuesday = 2 - dayOfWeek;
  } else { // Wednesday through Saturday
    daysUntilTuesday = 9 - dayOfWeek;
  }
  
  // If it's Tuesday after noon, move to next Tuesday
  if (dayOfWeek === 2 && date.getHours() >= 12) {
    daysUntilTuesday = 7;
  }
  
  date.setDate(date.getDate() + daysUntilTuesday);
  date.setHours(10, 0, 0, 0); // Set to 10 AM
  
  return date;
}

/**
 * Get meeting status based on date
 */
export function getMeetingStatus(dateString: string | null): {
  isUpcoming: boolean;
  isUnknown: boolean;
  isPast: boolean;
} {
  if (!dateString || dateString === 'unknown-meeting') {
    return { isUpcoming: false, isUnknown: true, isPast: false };
  }
  
  const meetingDate = parseSharePointDate(dateString);
  if (!meetingDate) {
    return { isUpcoming: false, isUnknown: true, isPast: false };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const meeting = new Date(meetingDate);
  meeting.setHours(0, 0, 0, 0);
  
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  return {
    isUpcoming: meeting >= today,
    isUnknown: false,
    isPast: meeting < today
  };
}