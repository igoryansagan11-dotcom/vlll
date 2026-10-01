// Собирает тренажёры печатных букв в .docx из картинок make_images.py.
// Использование: node build_docx.js <папка с картинками> <папка для .docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, HeightRule, VerticalAlign, PageBreak,
  Footer, PageNumber, LevelFormat, HeadingLevel,
} = require("docx");

const IMG = process.argv[2] || "build/img";
const OUT = process.argv[3] || "тренажеры";
const M = JSON.parse(fs.readFileSync(path.join(IMG, "manifest.json"), "utf8"));

const FONT = "Arial";
const MARGIN = 850; // ~1,5 см
const PX = (cm) => Math.round((cm / 2.54) * 96);

// ---------- общие элементы ----------

const img = (row, after = 80, keepNext = false) =>
  new Paragraph({
    keepNext,
    spacing: { before: 0, after },
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(path.join(IMG, row.file)),
      transformation: { width: PX(row.width_cm), height: PX(row.height_cm) },
    })],
  });

const p = (text, opts = {}) =>
  new Paragraph({
    alignment: opts.align,
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
    keepNext: opts.keepNext,
    numbering: opts.bullet ? { reference: "bullets", level: 0 } : undefined,
    children: [].concat(text).map((t) =>
      typeof t === "string" ? new TextRun({ text: t, size: opts.size ?? 24, bold: opts.bold, color: opts.color }) : t),
  });

const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, children: [new TextRun(text)] });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
const b = (text) => new TextRun({ text, bold: true, size: 24 });
const t = (text) => new TextRun({ text, size: 24 });

function makeDoc(title, children) {
  return new Document({
    creator: "Тренажёр печатных букв",
    title,
    styles: {
      default: { document: { run: { font: FONT, size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: FONT, color: "1F3864" },
          paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: FONT, color: "1F3864" },
          paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 1 } },
      ],
    },
    numbering: {
      config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] }],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, footer: 400 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: title + " — стр. ", size: 18, color: "808080" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "808080" })] })] }) },
      children,
    }],
  });
}

async function save(name, doc) {
  fs.mkdirSync(OUT, { recursive: true });
  fs.writeFileSync(path.join(OUT, name), await Packer.toBuffer(doc));
  console.log("готово:", name);
}

const howTo = (extra = []) => [
  h2("Как заниматься"),
  p("Распечатайте листы (подойдёт обычный чёрно-белый принтер). Пишите ручкой или простым карандашом.", { bullet: true }),
  p([b("Чёрная буква"), t(" в начале строки — образец. "), b("Серые буквы"), t(" — обведите точно по линиям. Пустое место — пишите сами, глядя на образец.")], { bullet: true }),
  p("Разлиновка: верхняя сплошная линия, пунктир посередине и основная (нижняя) линия.", { bullet: true }),
  p([b("Заглавные"), t(" буквы пишутся от верхней линии до основной.")], { bullet: true }),
  p([b("Строчные"), t(" — от пунктира до основной линии. У б — «хвостик» до верхней линии; у д, р, у, ф, ц, щ — элементы уходят ниже основной линии.")], { bullet: true }),
  p("Печатные буквы пишутся раздельно, без наклона и без соединений. Прямые линии ведите сверху вниз и слева направо, овалы — против часовой стрелки.", { bullet: true }),
  p("Расстояние между буквами — одинаковое, между словами — примерно ширина буквы «о».", { bullet: true }),
  p("Лучше 15–20 минут каждый день, чем час раз в неделю. Пишите медленно: аккуратность важнее скорости, скорость придёт сама.", { bullet: true }),
  ...extra,
];

// ---------- подсказки по написанию ----------

const HINT_UP = {
  "А": "1) наклонная от вершины вниз влево; 2) наклонная от вершины вниз вправо; 3) перекладина чуть ниже середины.",
  "Б": "1) вертикаль сверху вниз; 2) верхняя перекладина вправо; 3) от середины — полуовал вправо, замкнуть внизу.",
  "В": "1) вертикаль; 2) верхний полуовал до середины; 3) нижний полуовал — чуть шире верхнего.",
  "Г": "1) вертикаль сверху вниз; 2) перекладина сверху вправо.",
  "Д": "1) левая сторона «домика» от верха вниз; 2) верхняя перекладина; 3) правая вертикаль вниз; 4) нижняя полоса с короткими «ножками» под строку.",
  "Е": "1) вертикаль; 2) три перекладины слева направо: верхняя, средняя (чуть короче) и нижняя.",
  "Ё": "Пишется как Е, в конце — две точки сверху.",
  "Ж": "1) вертикаль посередине; 2) две наклонные сверху к середине (слева и справа); 3) две наклонные от середины вниз.",
  "З": "1) верхний полуовал (как у цифры 3); 2) нижний полуовал — чуть крупнее верхнего.",
  "И": "1) левая вертикаль; 2) правая вертикаль; 3) наклонная: снизу слева — вверх направо.",
  "Й": "Пишется как И, в конце сверху — маленькая дужка (краткая).",
  "К": "1) вертикаль; 2) наклонная от середины вверх вправо; 3) наклонная от середины вниз вправо.",
  "Л": "1) левая наклонная от верха вниз влево; 2) правая вертикаль от верха вниз.",
  "М": "1) левая вертикаль; 2) наклонная вниз к середине; 3) наклонная вверх; 4) правая вертикаль.",
  "Н": "1) левая вертикаль; 2) правая вертикаль; 3) перекладина посередине.",
  "О": "Овал против часовой стрелки, начиная сверху. Замкните ровно, без углов.",
  "П": "1) левая вертикаль; 2) верхняя перекладина; 3) правая вертикаль вниз.",
  "Р": "1) вертикаль; 2) полуовал справа сверху до середины.",
  "С": "Дуга против часовой стрелки: начните справа сверху, ведите влево, вниз и закончите справа внизу.",
  "Т": "1) верхняя перекладина; 2) вертикаль ровно посередине.",
  "У": "1) короткая наклонная слева сверху к середине; 2) длинная наклонная справа сверху вниз влево до основной линии.",
  "Ф": "1) овал посередине; 2) вертикаль сверху вниз через центр овала.",
  "Х": "1) наклонная из левого верхнего угла в правый нижний; 2) наклонная из правого верхнего в левый нижний.",
  "Ц": "1) левая вертикаль; 2) правая вертикаль; 3) нижняя перекладина и маленький хвостик справа под строку.",
  "Ч": "1) короткая линия слева сверху до середины, закруглите вправо; 2) правая вертикаль сверху до основной линии.",
  "Ш": "1)–3) три вертикали на одинаковом расстоянии; 4) нижняя перекладина.",
  "Щ": "Пишется как Ш, затем маленький хвостик справа под строку.",
  "Ъ": "1) короткая перекладина сверху влево; 2) вертикаль вниз; 3) полуовал справа в нижней половине.",
  "Ы": "1) вертикаль; 2) полуовал в нижней половине (как у Ь); 3) отдельная вертикаль справа.",
  "Ь": "1) вертикаль; 2) полуовал справа в нижней половине.",
  "Э": "1) дуга по часовой стрелке: сверху слева — вправо — вниз — влево; 2) перекладина посередине.",
  "Ю": "1) вертикаль; 2) короткая перекладина посередине; 3) овал справа.",
  "Я": "1) полуовал слева сверху (как зеркальная Р); 2) вертикаль справа; 3) наклонная «ножка» от середины вниз влево.",
};

const SAME = "Как заглавная, только высотой от пунктира до основной линии.";
const HINT_LO = {
  "а": "1) дуга сверху и палочка справа вниз; 2) маленький овал внизу слева, примыкающий к палочке.",
  "б": "1) «хвостик» от верхней линии вниз влево; 2) переходит в округлый овал в нижней части строки.",
  "д": "Как заглавная Д, только ниже: «домик» до пунктира, ножки уходят под основную линию.",
  "е": "1) горизонталь посередине слева направо; 2) дуга вверх и против часовой стрелки, снизу справа не замыкать.",
  "ё": "Пишется как е, в конце — две точки сверху.",
  "й": "Пишется как и, в конце — краткая сверху.",
  "р": "1) вертикаль от пунктира вниз под основную линию; 2) овал справа в верхней части.",
  "у": "1) наклонная слева сверху к основанию; 2) длинная наклонная справа сверху вниз — под основную линию.",
  "ф": "1) вертикаль от верхней линии до самого низа под строкой; 2) овал по центру — от пунктира до основной линии.",
  "ц": "Как заглавная Ц, только ниже; хвостик уходит под основную линию.",
  "щ": "Как заглавная Щ, только ниже; хвостик уходит под основную линию.",
};

function letterSections(list, hints, fallback) {
  const out = [];
  list.forEach((item, i) => {
    if (i > 0 && i % 3 === 0) out.push(pageBreak());
    const L = item.letter;
    out.push(new Paragraph({ keepNext: true, spacing: { before: i % 3 ? 240 : 0, after: 40 },
      children: [new TextRun({ text: `${L.toUpperCase()} ${L.toLowerCase()}`, bold: true, size: 40, color: "1F3864" }),
        new TextRun({ text: "    " + (hints[L] || fallback), size: 21, color: "404040" })] }));
    item.rows.forEach((r, k) => out.push(img(r, k === 2 ? 0 : 60)));
  });
  return out;
}

// ---------- 1. Заглавные ----------

async function upper() {
  await save("01_Заглавные_печатные_буквы.docx", makeDoc("Заглавные печатные буквы", [
    h1("Тренажёр 1. Заглавные печатные буквы"),
    p("33 буквы русского алфавита. На каждую букву — строка для обводки, строка «обведи три и продолжи сам» и строка для самостоятельного письма. Рядом с буквой — порядок написания."),
    ...howTo([
      h2("Весь алфавит"),
      p(M.upper.map((x) => x.letter).join("  "), { size: 40, bold: true, after: 200 }),
      p("Совет: сначала пройдите буквы, похожие по написанию: Г Т П Н И Ш Ц Щ (прямые линии) → А Л Д М Х Ж К (наклонные) → О С Э Ф Ю З (овалы и дуги) → Р В Б Ь Ъ Ы Я Е Ё Й Ч У.", { size: 22, color: "404040" }),
    ]),
    pageBreak(),
    ...letterSections(M.upper, HINT_UP, ""),
  ]));
}

// ---------- 2. Строчные ----------

async function lower() {
  await save("02_Строчные_печатные_буквы.docx", makeDoc("Строчные печатные буквы", [
    h1("Тренажёр 2. Строчные печатные буквы"),
    p("Большинство строчных печатных букв повторяют заглавные, только они ниже — от пунктира до основной линии. Особое внимание: а, б, д, е, р, у, ф, ц, щ — они отличаются от заглавных или выходят за пунктир / основную линию."),
    ...howTo([
      h2("Весь алфавит"),
      p(M.lower.map((x) => x.letter).join("  "), { size: 40, bold: true, after: 200 }),
    ]),
    pageBreak(),
    ...letterSections(M.lower, HINT_LO, SAME),
  ]));
}

// ---------- 3. Пары, слоги, слова ----------

async function words() {
  const c = [
    h1("Тренажёр 3. Пары букв, слоги и слова"),
    p("Когда отдельные буквы получаются, переходите к сочетаниям. Следите, чтобы все буквы в слове были одной высоты и стояли на основной линии, а расстояния между ними — одинаковые."),
    h2("Заглавная и строчная рядом"),
    p("Слева — образец, дальше — обведите. На пустой строке напишите эти пары сами.", { size: 22, color: "404040" }),
    ...M.pairs.flatMap((r, i) => [...(i === 6 ? [pageBreak()] : []), img(r.trace, 20, true), img(r.blank, 100)]),
    pageBreak(),
    h2("Слоги"),
    p("Обведите слоги, на пустой строке под ними напишите их сами.", { size: 22, color: "404040" }),
  ];
  M.syllables.forEach((s) => { c.push(img(s.trace, 20, true), img(s.blank, 120)); });
  c.push(pageBreak(), h2("Слова"),
    p("Чёрным — образец, серым — обведите. На пустой строке напишите эти слова сами.", { size: 22, color: "404040" }));
  M.words.forEach((w, i) => {
    c.push(img(w.trace, 20, true), img(w.blank, 120));
  });
  await save("03_Слоги_и_слова.docx", makeDoc("Слоги и слова", c));
}

// ---------- 4. Цифры и предложения ----------

async function sentences() {
  const c = [
    h1("Тренажёр 4. Цифры и предложения"),
    h2("Цифры"),
    p("Цифры пишутся той же высоты, что и заглавные буквы.", { size: 22, color: "404040" }),
    ...M.digits.map((r) => img(r)),
    h2("Предложения"),
    p("Обведите предложение, затем перепишите его на пустой строке. Начинайте с заглавной буквы, в конце ставьте точку. Последние три предложения содержат почти все буквы алфавита — это лучшая проверка.", { size: 22, color: "404040" }),
  ];
  M.sentences.forEach((s, i) => {
    s.trace.forEach((r) => c.push(img(r, 40, true)));
    s.blank.forEach((r, k) => c.push(img(r, k === s.blank.length - 1 ? 240 : 40, k < s.blank.length - 1)));
  });
  await save("04_Цифры_и_предложения.docx", makeDoc("Цифры и предложения", c));
}

// ---------- 5. Бланки с клетками ----------

const CELLS = 22;
const CELL = Math.floor((11906 - 2 * MARGIN) / CELLS); // DXA
const cellBorder = { style: BorderStyle.SINGLE, size: 6, color: "7F7F7F" };

function cellRow(text, color) {
  const chars = [...(text || "")];
  return new TableRow({
    height: { value: 560, rule: HeightRule.EXACT },
    children: Array.from({ length: CELLS }, (_, i) => new TableCell({
      width: { size: CELL, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
        children: chars[i] && chars[i] !== " "
          ? [new TextRun({ text: chars[i], size: 32, color, font: FONT })] : [] })],
    })),
  });
}

const grid = (rows) => new Table({
  width: { size: CELL * CELLS, type: WidthType.DXA },
  columnWidths: Array(CELLS).fill(CELL),
  rows,
});

const label = (text) => p(text, { size: 20, bold: true, after: 40, before: 160, keepNext: true, color: "404040" });
const GRAY = "A6A6A6";

async function forms() {
  const fields = [
    ["ФАМИЛИЯ", ["ИВАНОВА"]],
    ["ИМЯ", ["МАРИЯ"]],
    ["ОТЧЕСТВО", ["СЕРГЕЕВНА"]],
    ["ДАТА РОЖДЕНИЯ", ["05.03.1990"]],
    ["МЕСТО РОЖДЕНИЯ", ["Г. НИЖНИЙ НОВГОРОД"]],
    ["АДРЕС", ["УЛ. ЛЕНИНА, Д. 10, КВ. 5", "Г. МОСКВА"]],
    ["ТЕЛЕФОН", ["+7 900 123-45-67"]],
  ];
  const up = M.upper.map((x) => x.letter).join("");
  const c = [
    h1("Тренажёр 5. Заполнение бланков"),
    p("В анкетах, заявлениях и квитанциях обычно просят писать ЗАГЛАВНЫМИ ПЕЧАТНЫМИ буквами, по одной в каждой клетке."),
    p("Буква стоит в центре клетки и не касается её краёв.", { bullet: true }),
    p("Между словами оставляйте одну пустую клетку.", { bullet: true }),
    p("Точки, запятые и дефисы тоже занимают клетку.", { bullet: true }),
    p("Пишите чёрной или синей ручкой, без исправлений — испорченный бланк лучше взять новый.", { bullet: true }),
    p("В каждом поле: верхняя строка — образец серым для обводки, ниже — пустые строки для самостоятельной записи.", { bullet: true }),
    h2("Алфавит и цифры в клетках"),
    grid([cellRow(up.slice(0, CELLS), GRAY), cellRow("", GRAY),
      cellRow(up.slice(CELLS), GRAY), cellRow("", GRAY),
      cellRow("0 1 2 3 4 5 6 7 8 9", GRAY), cellRow("", GRAY)]),
    pageBreak(),
    h2("Образец анкеты"),
  ];
  for (const [name, lines] of fields) {
    c.push(label(name));
    const rows = [];
    lines.forEach((l) => rows.push(cellRow(l, GRAY)));
    lines.forEach(() => rows.push(cellRow("", GRAY)));
    c.push(grid(rows));
  }
  c.push(pageBreak(), h2("Ваши данные"),
    p("Заполните своими данными. Сначала напишите на черновике, посчитайте клетки — и только потом переносите.", { size: 22, color: "404040" }));
  for (const name of ["ФАМИЛИЯ", "ИМЯ", "ОТЧЕСТВО", "ДАТА РОЖДЕНИЯ", "МЕСТО РОЖДЕНИЯ", "АДРЕС", "ТЕЛЕФОН", "ЭЛЕКТРОННАЯ ПОЧТА"]) {
    c.push(label(name), grid([cellRow(""), cellRow("")]));
  }
  c.push(pageBreak(), h2("Свободная тренировка"));
  c.push(grid(Array.from({ length: 25 }, () => cellRow(""))));
  await save("05_Заполнение_бланков.docx", makeDoc("Заполнение бланков", c));
}

(async () => {
  await upper();
  await lower();
  await words();
  await sentences();
  await forms();
})();
