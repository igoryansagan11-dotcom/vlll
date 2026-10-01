"""Рисует строки прописей (PNG) для тренажёров печатных букв.

Каждая строка — разлиновка (верхняя линия, пунктирная середина, основная линия)
с образцом чёрным цветом и светло-серыми буквами для обводки.
Результат: папка с PNG и manifest.json, который читает build_docx.js.
"""
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

DPI = 300
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
ROW_WIDTH_CM = 17.5

INK = (0, 0, 0)
TRACE = (175, 175, 175)
LINE_TOP = (150, 150, 150)
LINE_MID = (170, 170, 170)
LINE_BASE = (80, 80, 80)
MARGIN_LINE = (225, 140, 140)

_probe = ImageFont.truetype(FONT_PATH, 1000)
CAP_RATIO = -_probe.getbbox("Н", anchor="ls")[1] / 1000
X_RATIO = -_probe.getbbox("х", anchor="ls")[1] / 1000


def cm(v):
    return int(round(v / 2.54 * DPI))


class Row:
    def __init__(self, cap_cm):
        self.H = cm(cap_cm)
        self.font = ImageFont.truetype(FONT_PATH, int(round(self.H / CAP_RATIO)))
        self.top = int(self.H * 0.45)
        self.base = self.top + self.H
        self.mid = self.base - int(round(self.font.size * X_RATIO))
        self.height = self.base + int(self.H * 0.36)
        self.width = cm(ROW_WIDTH_CM)
        self.gap = int(self.H * 0.55)
        self.left = int(self.H * 0.25)
        self.img = Image.new("RGB", (self.width, self.height), "white")
        self.draw = ImageDraw.Draw(self.img)
        self._rule()
        self.x = self.left

    def _rule(self):
        d, w = self.draw, self.width
        t = max(2, self.H // 60)
        d.line([(0, self.top), (w, self.top)], fill=LINE_TOP, width=t)
        dash = max(10, self.H // 10)
        for x in range(0, w, dash * 2):
            d.line([(x, self.mid), (x + dash, self.mid)], fill=LINE_MID, width=t)
        d.line([(0, self.base), (w, self.base)], fill=LINE_BASE, width=t + 1)

    def advance(self, text):
        return self.font.getlength(text)

    def fits(self, text):
        return self.x + self.advance(text) <= self.width - self.left

    def put(self, text, kind):
        if kind != "blank":
            color = INK if kind == "sample" else TRACE
            self.draw.text((self.x, self.base), text, font=self.font, fill=color, anchor="ls")
        self.x += self.advance(text) + self.gap

    def separator(self):
        """Красная вертикальная черта после образца, как в прописях."""
        x = int(self.x - self.gap / 2)
        self.draw.line([(x, self.top - self.H // 6), (x, self.base + self.H // 6)],
                       fill=MARGIN_LINE, width=max(2, self.H // 60))

    def fill(self, text, kind="trace", limit=None):
        n = 0
        while self.fits(text) and (limit is None or n < limit):
            self.put(text, kind)
            n += 1

    def save(self, path):
        self.img.save(path, dpi=(DPI, DPI))
        return {"file": os.path.basename(path),
                "width_cm": self.width / DPI * 2.54,
                "height_cm": self.height / DPI * 2.54}


UPPER = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
LOWER = UPPER.lower()

SYLLABLES = [
    "МА МО МУ МЫ МИ МЕ",
    "ПА ПО ПУ ПЫ ПИ ПЕ",
    "ЛА ЛО ЛУ ЛЫ ЛИ ЛЕ",
    "ра ро ру ры ри ре",
    "на но ну ны ни не",
    "са со су сы си се",
]
WORDS = [
    "МАМА", "ПАПА", "ДОМ", "КОТ", "СОК", "ЛЕС",
    "РЫБА", "ЖУК", "ЁЖ", "ЩИ", "ЧАЙ", "ЮЛА",
    "мир", "сад", "хлеб", "окно", "цветок", "шкаф",
    "подъезд", "семья", "яблоко", "вода", "друг", "фото",
    "Москва", "Россия", "Иван", "Анна", "Елена", "Пётр",
]
SENTENCES = [
    "Мама мыла раму.",
    "У нас дома живёт кот.",
    "Сегодня хорошая погода.",
    "Я учусь писать печатными буквами.",
    "Съешь же ещё этих мягких французских булок, да выпей чаю.",
    "Широкая электрификация южных губерний даст мощный толчок подъёму сельского хозяйства.",
    "В чащах юга жил бы цитрус? Да, но фальшивый экземпляр!",
]
DIGITS = "0123456789"


def letter_rows(out, letter, cap_cm, prefix):
    rows = []
    r = Row(cap_cm)
    r.put(letter, "sample"); r.separator(); r.fill(letter)
    rows.append(r.save(f"{out}/{prefix}_1.png"))

    r = Row(cap_cm)
    r.put(letter, "sample"); r.separator(); r.fill(letter, limit=3)
    rows.append(r.save(f"{out}/{prefix}_2.png"))

    r = Row(cap_cm)
    r.put(letter, "sample"); r.separator()
    rows.append(r.save(f"{out}/{prefix}_3.png"))
    return rows


def wrap(text, row_factory):
    """Разбивает фразу на куски, каждый из которых помещается в одну строку."""
    probe = row_factory()
    limit = probe.width - 2 * probe.left
    lines, cur = [], ""
    for w in text.split():
        cand = (cur + " " + w).strip()
        if probe.advance(cand) <= limit:
            cur = cand
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def main(out):
    os.makedirs(out, exist_ok=True)
    m = {"upper": [], "lower": [], "pairs": [], "syllables": [], "words": [],
         "digits": [], "sentences": []}

    for i, ch in enumerate(UPPER):
        m["upper"].append({"letter": ch, "rows": letter_rows(out, ch, 1.25, f"up{i:02d}")})
    for i, ch in enumerate(LOWER):
        m["lower"].append({"letter": ch, "rows": letter_rows(out, ch, 1.25, f"lo{i:02d}")})

    # Пары «Аа»: образец трёх пар, затем обводка этих пар до конца строки
    for i in range(0, len(UPPER), 3):
        pairs = [c + c.lower() for c in UPPER[i:i + 3]]
        r = Row(0.9)
        for pr in pairs:
            r.put(pr, "sample")
        r.separator()
        k = 0
        while r.fits(pairs[k % len(pairs)]):
            r.put(pairs[k % len(pairs)], "trace")
            k += 1
        m["pairs"].append({"trace": r.save(f"{out}/pair{i:02d}_t.png"),
                           "blank": Row(0.9).save(f"{out}/pair{i:02d}_b.png")})

    for i, s in enumerate(SYLLABLES):
        r = Row(0.9)
        for syl in s.split():
            r.put(syl, "trace")
        m["syllables"].append({"trace": r.save(f"{out}/syl{i:02d}_t.png")})
        r = Row(0.9)
        m["syllables"][-1]["blank"] = r.save(f"{out}/syl{i:02d}_b.png")

    for i in range(0, len(WORDS), 2):
        group = WORDS[i:i + 2]
        r = Row(0.85)
        k = 0
        while r.fits(group[k % len(group)]):
            r.put(group[k % len(group)], "sample" if k < len(group) else "trace")
            if k == len(group) - 1:
                r.separator()
            k += 1
        m["words"].append({"words": group, "trace": r.save(f"{out}/word{i:02d}_t.png")})
        r = Row(0.85)
        m["words"][-1]["blank"] = r.save(f"{out}/word{i:02d}_b.png")

    r = Row(1.0)
    for d in DIGITS:
        r.put(d, "sample")
    m["digits"].append(r.save(f"{out}/dig_sample.png"))
    for k in range(2):
        r = Row(1.0)
        for d in DIGITS:
            r.put(d, "trace")
        m["digits"].append(r.save(f"{out}/dig_t{k}.png"))
    for k in range(2):
        m["digits"].append(Row(1.0).save(f"{out}/dig_b{k}.png"))

    for i, s in enumerate(SENTENCES):
        parts = wrap(s, lambda: Row(0.75))
        entry = {"text": s, "trace": [], "blank": []}
        for j, p in enumerate(parts):
            r = Row(0.75)
            r.put(p, "trace")
            entry["trace"].append(r.save(f"{out}/sent{i:02d}_{j}_t.png"))
        for j in range(len(parts)):
            entry["blank"].append(Row(0.75).save(f"{out}/sent{i:02d}_{j}_b.png"))
        m["sentences"].append(entry)

    with open(f"{out}/manifest.json", "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "build/img")
