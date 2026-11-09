"""Build data.json + resized images for the T21 Asok Vivobarefoot lineup page.

Prices come only from the stock workbook ("Discout 20%" and "Normal Price" sheets).
Product copy/images come from vivobarefoot.co.jp (Shopify products.json) and
vivobarefoot.com/rw (English copy, transcribed into gl_text.json; images by SKU path).
"""
import json, re, io, sys, collections, html, urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import openpyxl
from PIL import Image

ROOT = Path(__file__).parent
SCR = Path(sys.argv[1])  # scratchpad with jp_all.json, gl_text.json, gl_imgs.json
XLSX = Path(r"C:/Users/Barefootinc TH/Downloads/イベント/T21Asok/VivoT21Asok Stock.xlsx")

jp = json.load(open(SCR / "jp_all.json", encoding="utf-8"))
gl_text = json.load(open(SCR / "gl_text.json", encoding="utf-8"))
gl_imgs = json.load(open(SCR / "gl_imgs.json", encoding="utf-8"))

jp_by_handle = {p["handle"]: p for p in jp}
jp_idx = collections.defaultdict(list)
for p in jp:
    for v in p["variants"]:
        m = re.match(r"Vivo(\d{6})-(\d{2})", v.get("sku") or "")
        if m:
            jp_idx[(m.group(1), m.group(2))].append((p, v))
jp_by_model = {}
for (mo, _), hits in jp_idx.items():
    jp_by_model.setdefault(mo, hits[0][0])

# model code -> (display model key, category)
CAT = {
    "MOTUS STRENGTH": "training", "MOTUS STRENGTH II": "training", "MOTUS FLEX": "training",
    "MOTUS FLEX NATURAL": "training", "PRIMUS LITE 3.5": "training", "PRIMUS LITE IV": "training",
    "PRIMUS LITE IV ALL WEATHER": "training", "PRIMUS LITE KNIT": "training",
    "PRIMUS LITE KNIT NATURAL": "training", "PRIMUS FLOW": "running", "PRIMUS FLOW KNIT": "running",
    "PRIMUS TRAIL KNIT FG": "trail", "PRIMUS TRAIL III ALL WEATHER FG": "trail",
    "PRIMUS TRAIL FLOW": "trail", "PRIMUS TRAIL FG 3.5": "trail",
    "GOBI II SNEAKER PREMIUM LEATHER": "everyday", "GOBI III SNEAKER LEATHER": "everyday",
    "GEO COURT IV": "everyday", "RA IV": "everyday", "SENSUS II": "everyday",
    "SENSUS YIN SLIDE": "everyday",
}
# global-site slug stem for English copy
GL = {
    "MOTUS STRENGTH II": "motus-strength-ii", "GOBI III SNEAKER LEATHER": "gobi-iii-sneaker-leather",
    "SENSUS II": "sensus-ii-grounding", "PRIMUS FLOW KNIT": "primus-flow-knit",
    "PRIMUS TRAIL FG 3.5": "primus-trail-fg-3-5", "RA IV": "ra-iv", "PRIMUS LITE IV": "primus-lite-iv",
    "PRIMUS TRAIL FLOW": "primus-trail-flow", "SENSUS YIN SLIDE": "sensus-yin-slide",
    "PRIMUS LITE KNIT NATURAL": "primus-lite-knit-natural", "MOTUS FLEX NATURAL": "motus-flex-natural",
    "PRIMUS TRAIL III ALL WEATHER FG": "primus-trail-iii-all-weather-fg",
    "GOBI II SNEAKER PREMIUM LEATHER": "gobi-ii-sneaker-premium-leather",
}
GL_SLUG = {  # global product URL per gender
    "SENSUS II": "sensus-ii-grounding", "RA IV": "ra-iv",
}


def model_name(desc):
    d = re.sub(r"\s+", " ", desc.strip().upper())
    m = re.match(r"(.+?) (WOMEN|MEN)\b", d)
    base = m.group(1)
    if base == "GOBI II SNEAKER LEATHER":
        base = "GOBI II SNEAKER PREMIUM LEATHER"
    return base, ("W" if m.group(2) == "WOMEN" else "M")


def colour_from_desc(desc):
    d = re.sub(r"\s+", " ", desc.strip())
    return re.split(r"\b(?:WOMEN|MEN|Women|Men)\b", d, maxsplit=1)[1].strip().title().replace("-", "/")


def clean_jp(body):
    """Split JP body_html into tagline, intro paragraphs and feature bullets."""
    body = re.sub(r"<meta[^>]*>", "", body or "")
    body = re.sub(r'<div[^>]*id="gtx-trans".*', "", body, flags=re.S)
    parts = re.split(r"<hr[^>]*>", body)
    fi = next((i for i, p in enumerate(parts) if "特徴" in p), len(parts))
    head = " ".join(parts[:fi]) if fi else parts[0]
    tail = " ".join(parts[fi:])
    def txt(s):
        s = re.sub(r"<br\s*/?>", "\n", s)
        s = re.sub(r"<[^>]+>", "", s)
        return re.sub(r"[ \t]+", " ", html.unescape(s).replace("\xa0", " ")).strip()
    lines = [l.strip() for l in txt(re.sub(r"</(p|h\d|div)>", "\n", head)).split("\n")]
    lines = [l for l in lines if l]
    tagline, intro = "", lines
    if lines and len(lines[0]) < 60:
        tagline, intro = lines[0], lines[1:]
    intro = "".join(intro)
    feats = []
    for li in re.findall(r"<li[^>]*>(.*?)</li>", tail, re.S):
        segs = [s.strip() for s in txt(li).split("\n") if s.strip()]
        if not segs or re.match(r"^重量", segs[0]):
            continue
        title, rest = (segs[0], " ".join(segs[1:])) if len(segs) > 1 and len(segs[0]) < 40 else ("", " ".join(segs))
        if title.startswith("ヴィーガン"):
            rest = "主要素材は動物由来ではない合成素材。"
        if len(rest) > 140:
            rest = rest[:rest.rfind("。", 0, 140) + 1] or rest[:140] + "…"
        feats.append({"t": title, "d": rest})
    weight = ""
    m = re.search(r"重量[^0-9A-Z]*((?:WOMENS|MENS)?[^<]{0,30}?\d+(?:\.\d)?\s*g)", txt(tail))
    if m:
        weight = m.group(1).strip()
    return tagline, intro, feats[:7], weight


def img_src(code, ang):
    return f"https://www.vivobarefoot.com/media/catalog/product/{code[0]}/{code[1]}/{code}_{ang}.jpg"


# ---------- read workbook ----------
wb = openpyxl.load_workbook(XLSX, data_only=True)
models = collections.OrderedDict()
for sheet, tier in (("Discout 20%", "sale"), ("Normal Price", "full")):
    for r in wb[sheet].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        sku, desc, qty, barcode, price = r[0], r[1], r[2] or 0, r[3], r[4]
        sale = r[5] if tier == "sale" else None
        mo, co, size = sku.split("-")
        name, g = model_name(desc)
        M = models.setdefault(name, {"name": name, "cat": CAT[name], "codes": {}, "cw": collections.OrderedDict()})
        M["codes"][g] = mo
        key = f"{mo}-{co}"
        cw = M["cw"].setdefault(key, {"code": key, "g": g, "tier": tier, "price": price, "sale": sale,
                                       "colour": colour_from_desc(desc), "sizes": []})
        cw["sizes"].append({"s": size[1:], "q": qty, "sku": sku, "bc": str(barcode)})

# ---------- enrich ----------
jobs = []
for name, M in models.items():
    # JP copy: prefer product matching this model code, women's first
    jp_p = None
    for g in ("W", "M"):
        if g in M["codes"] and M["codes"][g] in jp_by_model:
            jp_p = jp_by_model[M["codes"][g]]
            break
    note = ""
    if jp_p is None and name == "SENSUS II":
        jp_p = jp_by_handle["sensus-ii-womens"]
        note = "説明と画像は現行の Sensus II Grounding(公式サイト掲載)のもの"
    tag, intro, feats, jpw = clean_jp(jp_p["body_html"]) if jp_p else ("", "", [], "")
    M.update(jp_title=(jp_p["title"].split("/")[1].strip() if jp_p and "/" in jp_p["title"] else ""),
             tagline=tag, intro_ja=intro, feats=feats, weight_ja=jpw, note=note,
             jp_url=f"https://vivobarefoot.co.jp/products/{jp_p['handle']}" if jp_p else "")
    en = gl_text.get(GL.get(name, ""), {})
    M["en"] = en
    stem = GL.get(name)
    M["gl_url"] = {g: f"https://www.vivobarefoot.com/rw/{stem}-{'womens' if g == 'W' else 'mens'}"
                   for g in M["codes"]} if stem else {}
    M["cw"] = list(M["cw"].values())
    for cw in M["cw"]:
        angles = gl_imgs.get(cw["code"]) or []
        srcs = [("gl", a, img_src(cw["code"], a)) for a in angles if a in ("side", "pair", "back", "inside")]
        if not srcs:
            mo, co = cw["code"].split("-")
            hits = jp_idx.get((mo, co))
            if not hits and name == "SENSUS II":  # Tan of Sensus II Grounding as reference
                p = jp_by_handle["sensus-ii-womens" if cw["g"] == "W" else "sensus-ii-mens"]
                hits = [(p, v) for v in p["variants"] if v["option1"] == "TAN"]
                cw["ref_img"] = True
            if hits:
                p = hits[0][0]; vids = {v["id"] for _, v in hits}
                im = [i for i in p["images"] if set(i["variant_ids"]) & vids]
                if im:
                    srcs = [("jp", "side", im[0]["src"])]
        cw["imgs"] = []
        for src, a, url in srcs:
            fn = f"img/{cw['code']}_{a}.jpg"
            cw["imgs"].append(fn)
            jobs.append((url, ROOT / fn))
        cw["img_src"] = srcs[0][0] if srcs else None


def fetch(job):
    url, out = job
    if out.exists():
        return
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    data = urllib.request.urlopen(req, timeout=60).read()
    im = Image.open(io.BytesIO(data))
    if im.mode in ("RGBA", "P", "LA"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255)); bg.paste(im, mask=im.split()[-1]); im = bg
    im = im.convert("RGB")
    im.thumbnail((720, 720))
    im.save(out, "JPEG", quality=80, optimize=True, progressive=True)

with ThreadPoolExecutor(12) as ex:
    list(ex.map(fetch, jobs))

# Regular (always-stocked) lineup = the K Village "SALE 10% OFF" poster, 7-30 Sep 2026.
# Model-colour codes, gender is the first digit (2 = Women, 3 = Men).
REGULAR = {
    "209253-01", "309253-01", "209253-04", "309253-04", "209253-07", "309253-05",  # Primus Lite IV
    "209576-03", "309576-03",                                                      # Primus Lite Knit Natural
    "309600-03",                                                                   # Primus Trail FG 3.5
    "309123-07", "309123-02", "209123-02", "209123-01",                            # Gobi III Sneaker Leather
    "309648-05",                                                                   # Motus Strength II
    "209062-01",                                                                   # Geo Court IV
    "309553-01",                                                                   # Sensus II
    "305102-01",                                                                   # Ra IV
}
all_codes = {c["code"] for m in models.values() for c in m["cw"]}
missing = REGULAR - all_codes
assert not missing, f"regular codes not in workbook: {missing}"

# The page is for product knowledge: list price only, no event quantities or barcodes.
for m in models.values():
    for c in m["cw"]:
        c["regular"] = c["code"] in REGULAR
        for k in ("sizes", "sale", "tier"):
            c.pop(k, None)

out = list(models.values())
json.dump(out, open(ROOT / "data.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(out), "models,", sum(len(m["cw"]) for m in out), "colourways,",
      sum(c["regular"] for m in out for c in m["cw"]), "regular,", len(jobs), "images")
for m in out:
    for c in m["cw"]:
        if not c["imgs"]:
            print("NO IMAGE", c["code"], m["name"], c["colour"])
    if not m["intro_ja"]:
        print("NO JP COPY", m["name"])
