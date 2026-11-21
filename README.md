# Vivobarefoot lineup

Barefoot Inc が扱う Vivobarefoot の商品ページです(日本語 / English / ไทย)。

- 公開URL: https://vivo-lineup-amber.vercel.app
- Vercel: チーム `SHIMA's projects` / プロジェクト `vivo-lineup`
- このリポジトリの `main` に変更を入れると、Vercel が自動で本番に反映します(1〜2分)

## ファイルの場所

| 場所 | 中身 |
|---|---|
| `public/index.html` | ページ本体。画面(HTML・CSS・JavaScript)と商品データ(`const DATA = [...]`)が1ファイルに入っています |
| `public/img/` | 商品画像 |
| `legacy-2026-11/` | 2026-11-09 時点の古い元ファイル(Excel から作る手順)。現行ページとは機能が違うので、参考用です |

`public/` の中身だけが公開されます。

## 編集のしかた

1. このリポジトリを手元に取得します(`git clone https://github.com/KeisukeKirby/t21-vivo-lineup.git`)
2. ブランチを切ります(例: `git switch -c fix-price`)
3. `public/index.html` を編集します。商品の追加・価格・説明の修正は `const DATA` の中です。Claude Code に「`public/index.html` の DATA で ○○ を直して」と頼むのが早いです
4. パソコンで `public/index.html` をブラウザで開いて確認します(スマホ幅は 375px 付近で確認)
5. GitHub に push して Pull Request を作り、`main` にマージします。マージすると本番に反映されます

ブランチ(`main` 以外)を push すると、Vercel がプレビュー用のページも作ります(チームメンバーだけが見られます)。

## 気をつけること

- 画像を追加するときは `public/img/` に置き、`DATA` からは `img/ファイル名` で参照します
- スマホ幅で3つのタブが収まるよう、`<style id="mobile-fix">` を入れてあります。消さないでください
- 価格はイベント在庫表の通常価格です。数量や原価は載せません(公開ページのため)
- Vercel の無料枠ではなく Pro プランのチームですが、画像の大きさには気をつけてください(1枚あたり 100KB 前後が目安)
