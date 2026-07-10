export type SearchRange = {
  since: string;
  until: string;
  days: number;
  usedDefaultLookback: boolean;
  notice: string;
};

const datePattern = /^\d{4}-\d{2}-\d{2}$/;

export function resolveSearchRange(
  since: string | undefined,
  until: string | undefined,
  defaultLookbackDays: number,
  now = new Date()
): SearchRange {
  const resolvedUntil = until ? parseDateOnly(until, "until") : startOfUtcDay(now);
  const usedDefaultLookback = !since;
  const resolvedSince = since
    ? parseDateOnly(since, "since")
    : addUtcDays(resolvedUntil, -(defaultLookbackDays - 1));

  if (resolvedSince.getTime() > resolvedUntil.getTime()) {
    throw new Error("since must be on or before until.");
  }

  const sinceText = formatDateOnly(resolvedSince);
  const untilText = formatDateOnly(resolvedUntil);
  const days = Math.floor((resolvedUntil.getTime() - resolvedSince.getTime()) / 86_400_000) + 1;
  const notice = usedDefaultLookback
    ? `기간 미지정: 최근 ${days}일(${sinceText} ~ ${untilText})을 조회했습니다.`
    : `요청 기간: ${sinceText} ~ ${untilText}(${days}일)을 조회했습니다.`;

  return {
    since: sinceText,
    until: untilText,
    days,
    usedDefaultLookback,
    notice
  };
}

function parseDateOnly(value: string, optionName: string): Date {
  if (!datePattern.test(value)) {
    throw new Error(`${optionName} must use YYYY-MM-DD format.`);
  }

  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (formatDateOnly(date) !== value) {
    throw new Error(`${optionName} is not a valid calendar date.`);
  }
  return date;
}

function startOfUtcDay(value: Date): Date {
  return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
}

function addUtcDays(value: Date, days: number): Date {
  return new Date(value.getTime() + days * 86_400_000);
}

function formatDateOnly(value: Date): string {
  return value.toISOString().slice(0, 10);
}
