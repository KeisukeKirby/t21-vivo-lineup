"""Inline data.json into template.html.

index.html      -> Artifact version (the Artifact host adds the html/head skeleton)
site/index.html -> standalone version for Vercel (own skeleton, noindex) + site/img
"""
import json, shutil
from pathlib import Path

root = Path(__file__).parent
data = json.load(open(root / "data.json", encoding="utf-8"))
tpl = (root / "template.html").read_text(encoding="utf-8")
js = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", r"<\/")
page = tpl.replace("/*DATA*/", js)
(root / "index.html").write_text(page, encoding="utf-8")

site = root / "site"
site.mkdir(exist_ok=True)
head = ('<!doctype html>\n<html lang="ja">\n<head>\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
        '<meta name="robots" content="noindex">\n'
        '<style>html{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n')
# charset/title/meta/link/style from the template stay in <head>; the rest goes into <body>
split = page.index("</style>") + len("</style>")
(site / "index.html").write_text(head + page[:split] + "\n</head>\n<body>\n" + page[split:] + "\n</body>\n</html>\n",
                                 encoding="utf-8")
shutil.copytree(root / "img", site / "img", dirs_exist_ok=True)
print("index.html + site/index.html built,", len(list((site / "img").iterdir())), "images")
