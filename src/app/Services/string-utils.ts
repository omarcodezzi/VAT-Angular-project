export function normalizeHeader(s: string): string {
  return (s ?? "")
    .toString()
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")      // remove spaces
    .replace(/[_\-]/g, "");   // remove _ and -
}

// Levenshtein distance (small + fast)
export function levenshtein(a: string, b: string): number {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[m][n];
}

export function bestSuggestion(target: string, candidates: string[]): string | undefined {
  const t = normalizeHeader(target);  // Normalize the target (expected header)
  let best = { name: candidates[0], score: Infinity };

  // Loop through the possible candidates to find the best match
  for (const c of candidates) {
    const score = levenshtein(t, normalizeHeader(c));  // Use levenshtein or similar to calculate the similarity
    if (score < best.score) best = { name: c, score };
  }

  // Return the best match if it's within a reasonable threshold
  return best.score <= 2 ? best.name : undefined;  // Adjust threshold as needed
}


