export const AccountRegistry = {
  matchesFilter(acc: { status?: string | null }, filter: string): boolean {
    if (filter === "ACTIVE") {
      return (acc.status || "ACTIVE").toUpperCase() === "ACTIVE";
    }
    return true;
  },
};

export const RevenueRegistry = {
  isExcludedFromRevenue(acc: { status?: string | null }): boolean {
    const st = (acc.status || "").toUpperCase();
    return st === "BLOCKED" || st === "CLOSED" || st === "TERMINATED";
  },
};
