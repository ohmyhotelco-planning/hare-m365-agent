export type SearchRange = {
  since: string;
  until: string;
  startDateTime: string;
  endDateTimeExclusive: string;
  timeZone: string;
  days: number;
  usedDefaultLookback: boolean;
  notice: string;
};

const datePattern = /^\d{4}-\d{2}-\d{2}$/;

export function resolveSearchRange(
  since: string | undefined,
  until: string | undefined,
  defaultLookbackDays: number,
  now = new Date(),
  timeZone = "Asia/Seoul"
): SearchRange {
  validateTimeZone(timeZone);
  const resolvedUntil = until ? validateDateOnly(until, "until") : dateInTimeZone(now, timeZone);
  const usedDefaultLookback = !since;
  const resolvedSince = since
    ? validateDateOnly(since, "since")
    : addCalendarDays(resolvedUntil, -(defaultLookbackDays - 1));

  if (resolvedSince > resolvedUntil) {
    throw new Error("since must be on or before until.");
  }

  const endExclusiveDate = addCalendarDays(resolvedUntil, 1);
  const days = calendarDayDifference(resolvedSince, resolvedUntil) + 1;
  const notice = usedDefaultLookback
    ? `기간 미지정: 최근 ${days}일(${resolvedSince} ~ ${resolvedUntil}, ${timeZone})을 조회했습니다.`
    : `요청 기간: ${resolvedSince} ~ ${resolvedUntil}(${days}일, ${timeZone})을 조회했습니다.`;

  return {
    since: resolvedSince,
    until: resolvedUntil,
    startDateTime: zonedMidnightToUtc(resolvedSince, timeZone).toISOString(),
    endDateTimeExclusive: zonedMidnightToUtc(endExclusiveDate, timeZone).toISOString(),
    timeZone,
    days,
    usedDefaultLookback,
    notice
  };
}

function validateDateOnly(value: string, optionName: string): string {
  if (!datePattern.test(value)) {
    throw new Error(`${optionName} must use YYYY-MM-DD format.`);
  }

  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.toISOString().slice(0, 10) !== value) {
    throw new Error(`${optionName} is not a valid calendar date.`);
  }
  return value;
}

function validateTimeZone(timeZone: string): void {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone }).format(new Date());
  } catch {
    throw new Error(`Invalid IANA time zone: ${timeZone}`);
  }
}

function dateInTimeZone(value: Date, timeZone: string): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(value);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function addCalendarDays(value: string, days: number): string {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day + days)).toISOString().slice(0, 10);
}

function calendarDayDifference(from: string, to: string): number {
  return Math.floor((Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / 86_400_000);
}

function zonedMidnightToUtc(dateOnly: string, timeZone: string): Date {
  const [year, month, day] = dateOnly.split("-").map(Number);
  const utcGuess = Date.UTC(year, month - 1, day);
  let candidate = utcGuess;

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23"
    }).formatToParts(new Date(candidate));
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    const representedAsUtc = Date.UTC(
      Number(values.year),
      Number(values.month) - 1,
      Number(values.day),
      Number(values.hour),
      Number(values.minute),
      Number(values.second)
    );
    const correction = representedAsUtc - utcGuess;
    if (correction === 0) break;
    candidate -= correction;
  }

  return new Date(candidate);
}
