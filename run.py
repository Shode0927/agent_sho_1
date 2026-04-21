"""
run.py
エントリーポイント。コマンドライン引数でモードを切り替える。

使い方:
  python run.py analyze          # 文体プロファイル生成（初回のみ）
  python run.py analyze /path    # ドキュメントフォルダを指定
  python run.py                  # 対話モードでレポート生成
  python run.py "課題テキスト"   # 課題を直接渡してレポート生成
  python run.py "課題" --pdf     # PDF形式で出力
"""

import sys
import os

def main():
    args = sys.argv[1:]

    # ─── analyze モード ───────────────────────────────────
    if args and args[0] == "analyze":
        from style_analyzer import run as analyze_run
        docs_dir = args[1] if len(args) > 1 else None
        analyze_run(docs_dir)
        return

    # ─── レポート生成モード ───────────────────────────────
    from report_agent import run_agent, interactive_mode

    output_format = "pdf" if "--pdf" in args else "docx"
    clean_args = [a for a in args if not a.startswith("--")]

    if clean_args:
        # 課題をコマンドライン引数から受け取る
        assignment = " ".join(clean_args)
        run_agent(assignment, output_format=output_format)
    else:
        # 対話モード
        interactive_mode()


if __name__ == "__main__":
    main()
