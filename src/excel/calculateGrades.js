import { num } from '../utils/format.js';

export function classify(score) {
  const s = num(score) ?? 0;
  if (s >= 9) return 'Xuất Sắc';
  if (s >= 8) return 'Giỏi';
  if (s >= 6.5) return 'Khá';
  if (s >= 5) return 'Trung Bình';
  return 'Yếu';
}

// RANK.EQ-equivalent (competition ranking: ties share a rank, next rank skips ahead)
export function rankScores(students, key) {
  const vals = students.map((s) => num(s[key]) ?? -Infinity);
  students.forEach((s, i) => {
    const v = vals[i];
    s[key + 'Rank'] = 1 + vals.filter((x) => x > v).length;
  });
}
