// Show Excel formulas the way Danish Excel expects them: Danish function names,
// semicolons between arguments and a decimal comma.
const NAMES = {
  "COUNTIFS": "TÆL.HVISER", "COUNTIF": "TÆL.HVIS", "COUNTA": "TÆLV", "COUNT": "TÆL",
  "SUMIFS": "SUM.HVISER", "SUMIF": "SUM.HVIS", "AVERAGEIFS": "MIDDEL.HVISER", "AVERAGEIF": "MIDDEL.HVIS", "AVERAGE": "MIDDEL",
  "VLOOKUP": "LOPSLAG", "XLOOKUP": "XOPSLAG", "INDEX": "INDEKS", "MATCH": "SAMMENLIGN",
  "LEFT": "VENSTRE", "RIGHT": "HØJRE", "MID": "MIDT", "LEN": "LÆNGDE", "FIND": "FIND", "SUBSTITUTE": "UDSKIFT",
  "VALUE": "VÆRDI", "TEXT": "TEKST", "CONCATENATE": "SAMMENKÆDNING", "UNIQUE": "ENTYDIGE", "ROWS": "RÆKKER",
  "IFERROR": "HVIS.FEJL", "ISNUMBER": "ER.TAL", "IF": "HVIS", "MIN": "MIN", "MAX": "MAKS",
  "ROUND": "AFRUND", "MROUND": "MAFRUND", "YEAR": "ÅR", "MONTH": "MÅNED", "WEEKDAY": "UGEDAG",
};
const RX = new RegExp("\\b(" + Object.keys(NAMES).sort((a, b) => b.length - a.length).join("|") + ")(?=\\()", "g");

function formula(f) {
  let out = "", inStr = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"') inStr = !inStr;
    if (!inStr && ch === ",") out += ";";
    else if (!inStr && ch === "." && /\d/.test(f[i - 1] || "") && /\d/.test(f[i + 1] || "")) out += ",";
    else out += ch;
  }
  // Function names outside string literals only
  return out.split(/("[^"]*")/).map((part, i) => (i % 2 ? part : part.replace(RX, (m) => NAMES[m]))).join("");
}

// A formula starts at "=" followed by a letter and ends where its parentheses close
// and the next character is whitespace or the end of the text (quotes are respected).
function formulaEnd(s, i) {
  let depth = 0, inStr = false;
  for (let j = i + 1; j < s.length; j++) {
    const ch = s[j];
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === "(") depth++;
    else if (!inStr && ch === ")") depth--;
    if (!inStr && depth <= 0 && /\s/.test(s[j + 1] || " ")) return j + 1;
  }
  return s.length;
}

export function toDanish(text) {
  const s = String(text);
  let out = "", i = 0;
  while (i < s.length) {
    if (s[i] === "=" && /[A-Za-z]/.test(s[i + 1] || "")) {
      const end = formulaEnd(s, i);
      const f = s.slice(i, end), tail = f.match(/[.,;]*$/)[0];
      out += formula(f.slice(0, f.length - tail.length)) + tail;
      i = end;
    } else out += s[i++];
  }
  return out.replace("Data → From Text/CSV", "Data → Fra tekst/CSV").replace("Data → Remove Duplicates", "Data → Fjern dubletter");
}
