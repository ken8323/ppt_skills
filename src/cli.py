"""CLI エントリポイント: ppt-skills generate <config.json> <output.pptx>"""
from __future__ import annotations
import argparse
import json
import sys


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="ppt-skills",
        description="JSON 設定から PowerPoint を生成します",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    gen = sub.add_parser("generate", help="PPTX を生成する")
    gen.add_argument("config", help="設定 JSON ファイルのパス")
    gen.add_argument("output", help="出力 .pptx ファイルのパス")
    gen.add_argument("--no-lint", action="store_true", help="Lint チェックをスキップ")
    gen.add_argument("--no-validate", action="store_true", help="バリデーションをスキップ")

    scaffold_cmd = sub.add_parser("scaffold", help="テンプレートから設定 JSON を生成する")
    scaffold_cmd.add_argument("template", help="テンプレート名")
    scaffold_cmd.add_argument("output", nargs="?", default="-", help="出力先 JSON ファイル (省略時は stdout)")
    scaffold_cmd.add_argument("--list", action="store_true", help="利用可能なテンプレート一覧を表示")

    list_cmd = sub.add_parser("list-icons", help="利用可能なアイコン一覧を表示する")

    thumb = sub.add_parser("thumbnail", help="PPTX の各スライドを PNG サムネイルに変換する")
    thumb.add_argument("pptx", help="入力 .pptx ファイルのパス")
    thumb.add_argument("output_dir", help="PNG を保存するディレクトリ")
    thumb.add_argument("--dpi", type=int, default=150, help="解像度 (LibreOffice 使用時, 既定 150)")
    thumb.add_argument("--slide", type=int, default=None, metavar="N", help="特定スライドのみ (0 始まり)")

    args = parser.parse_args()

    if args.command == "generate":
        from src.generator import generate_pptx
        with open(args.config, encoding="utf-8") as f:
            config = json.load(f)
        generate_pptx(
            config, args.output,
            lint=not args.no_lint,
            validate=not args.no_validate,
        )
        print(f"生成完了: {args.output}", file=sys.stderr)

    elif args.command == "scaffold":
        from src.scaffold import scaffold, list_templates
        if args.list:
            for name in list_templates():
                print(name)
            return
        config = scaffold(args.template)
        text = json.dumps(config, ensure_ascii=False, indent=2)
        if args.output == "-":
            print(text)
        else:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"スキャフォールド完了: {args.output}", file=sys.stderr)

    elif args.command == "list-icons":
        from src.components.icon import list_icons
        for name in list_icons():
            print(name)

    elif args.command == "thumbnail":
        from src.thumbnail import generate_thumbnails
        paths = generate_thumbnails(
            args.pptx, args.output_dir,
            dpi=args.dpi,
            slide_index=args.slide,
        )
        for p in paths:
            print(str(p))
