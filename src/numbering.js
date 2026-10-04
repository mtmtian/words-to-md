import { attr, child, descendants, isW, kids, on } from './ooxml.js';

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];

// 1–9999 in Chinese numerals with 零 for skipped digits: 十一, 二十, 一百零五.
function chineseNumber(n, digits, units, shortTens) {
  if (n < 1 || n > 9999) return String(n);
  let text = '', gap = false;
  [...String(n)].forEach((digit, i, all) => {
    if (digit === '0') { gap = true; return; }
    if (gap && text) text += digits[0];
    gap = false;
    text += digits[digit] + units[all.length - 1 - i];
  });
  return shortTens && n < 20 && n >= 10 ? text.slice(1) : text;
}
const chinese = n => chineseNumber(n, '零一二三四五六七八九', ['', '十', '百', '千'], true);
const roman = n => ROMAN.reduce((text, [value, letters]) => { for (; n >= value; n -= value) text += letters; return text; }, '');
const letter = n => n < 1 ? String(n) : String.fromCharCode(97 + (n - 1) % 26).repeat(Math.ceil(n / 26));
const enclosed = (first, fallback) => n => n >= 1 && n <= 20 ? String.fromCodePoint(first + n - 1) : fallback(n);
const fullWidth = n => String(n).replace(/\d/g, digit => String.fromCharCode(0xff10 + Number(digit)));

// Word number formats (ST_NumberFormat) used in Chinese and Western documents; anything else falls back to decimal.
const FORMATS = {
  decimal: String, decimalHalfWidth: String, decimalZero: n => String(n).padStart(2, '0'),
  decimalFullWidth: fullWidth, decimalFullWidth2: fullWidth,
  upperLetter: n => letter(n).toUpperCase(), lowerLetter: letter, upperRoman: n => roman(n).toUpperCase(), lowerRoman: roman,
  chineseCounting: chinese, chineseCountingThousand: chinese, taiwaneseCounting: chinese, japaneseCounting: chinese,
  chineseLegalSimplified: n => chineseNumber(n, '零壹贰叁肆伍陆柒捌玖', ['', '拾', '佰', '仟'], false),
  ideographDigital: n => String(n).replace(/\d/g, digit => '〇一二三四五六七八九'[digit]),
  ideographTraditional: n => '甲乙丙丁戊己庚辛壬癸'[(n - 1) % 10] || String(n),
  ideographZodiac: n => '子丑寅卯辰巳午未申酉戌亥'[(n - 1) % 12] || String(n),
  decimalEnclosedCircle: enclosed(0x2460, String), decimalEnclosedCircleChinese: enclosed(0x2460, String),
  decimalEnclosedParen: enclosed(0x2474, n => `(${n})`), decimalEnclosedFullstop: enclosed(0x2488, n => `${n}.`),
  none: () => ''
};

// Restores automatic numbering labels ("一、", "（一）", "1.1", "①") as Word displays them. Each list instance (w:num)
// keeps its own counters, and using a level restarts deeper levels unless their w:lvlRestart says otherwise.
// Returns label(numId, ilvl, styleId), to be called once per paragraph in document order.
export function createNumbering(numberingDoc, warn) {
  const abstracts = new Map(), styleLinked = new Map(), nums = new Map(), counters = new Map();
  for (const el of descendants(numberingDoc, 'abstractNum')) {
    const id = attr(el, 'abstractNumId');
    abstracts.set(id, { levels: new Map(kids(el).filter(lvl => isW(lvl, 'lvl')).map(lvl => [attr(lvl, 'ilvl'), lvl])), link: attr(child(el, 'numStyleLink'), 'val') });
    if (attr(child(el, 'styleLink'), 'val')) styleLinked.set(attr(child(el, 'styleLink'), 'val'), id);
  }
  for (const el of descendants(numberingDoc, 'num')) {
    nums.set(attr(el, 'numId'), { abstractId: attr(child(el, 'abstractNumId'), 'val'), overrides: new Map(kids(el).filter(o => isW(o, 'lvlOverride')).map(o => [attr(o, 'ilvl'), o])) });
  }
  // A list-style reference (numStyleLink) points at the definition that declares the matching styleLink.
  const definition = num => {
    const abstract = abstracts.get(num.abstractId);
    return abstract?.link ? abstracts.get(styleLinked.get(abstract.link)) || abstract : abstract;
  };
  const level = (num, depth) => {
    const override = num.overrides.get(String(depth));
    const lvl = child(override, 'lvl') || definition(num)?.levels.get(String(depth));
    return lvl && {
      start: Number(attr(child(override, 'startOverride'), 'val') || attr(child(lvl, 'start'), 'val') || 0),
      format: attr(child(lvl, 'numFmt'), 'val') || 'decimal',
      text: attr(child(lvl, 'lvlText'), 'val'),
      restart: attr(child(lvl, 'lvlRestart'), 'val'),
      legal: on(child(lvl, 'isLgl')),
      suffix: attr(child(lvl, 'suff'), 'val') || 'tab'
    };
  };
  // Without an explicit level, the level linked to the paragraph style (w:pStyle) applies.
  const styleLevel = (num, styleId) => [...(definition(num)?.levels || [])].find(([, lvl]) => attr(child(lvl, 'pStyle'), 'val') === styleId)?.[0];
  const format = (name, n) => {
    if (FORMATS[name]) return FORMATS[name](n);
    warn(`部分编号格式（${name}）按阿拉伯数字输出。`);
    return String(n);
  };

  return (numId, ilvl, styleId) => {
    if (!numId || numId === '0') return '';
    const num = nums.get(numId);
    if (!num) { warn('缺少编号定义，部分自动编号未还原。'); return ''; }
    const depth = Number(ilvl || styleLevel(num, styleId) || 0);
    const current = level(num, depth);
    if (!current) return '';
    if (current.format === 'bullet') { warn('项目符号未保留，仅保留段落文字。'); return ''; }
    const counts = counters.get(numId) || [];
    counters.set(numId, counts);
    counts[depth] = counts[depth] === undefined ? current.start : counts[depth] + 1;
    for (let deeper = depth + 1; deeper < 9; deeper++) {
      const restart = level(num, deeper)?.restart;
      if (!restart || restart !== '0' && depth < Number(restart)) counts[deeper] = undefined;
    }
    const label = current.text.replace(/%([1-9])/g, (_, n) => {
      const shown = n - 1, lvl = shown === depth ? current : level(num, shown);
      return format(current.legal ? 'decimal' : lvl?.format || 'decimal', counts[shown] ?? lvl?.start ?? 0);
    });
    // The default tab after a number is layout; after CJK punctuation such as "一、" or "（一）" no space is added.
    const gap = current.suffix === 'nothing' || current.suffix === 'tab' && /[\u3000-\u303f\uff00-\uffef]$/.test(label) ? '' : ' ';
    return label && label + gap;
  };
}
