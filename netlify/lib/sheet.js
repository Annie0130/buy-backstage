// 試算表把儲存格開頭的 = + - @ 當成公式。先加一個 '，寫進去就只是文字。
export function sheetCell(value) {
  const s = String(value ?? '');
  if (s.length && '=+-@'.includes(s[0])) return `'${s}`;
  return s;
}
