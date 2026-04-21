"""
report_agent.py
中沢頌の文体でレポートを書くメインエージェント。
style_profile.md を読み込み、課題の指示に従ってレポートを生成する。
"""

import json
import os
from pathlib import Path

import anthropic

from config import STYLE_PROFILE_PATH, OUTPUT_DIR, ANTHROPIC_API_KEY, MODEL
from tools import ALL_TOOLS, dispatch_tool

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ──────────────────────────────────────────────
# システムプロンプト
# ──────────────────────────────────────────────

BASE_SYSTEM_PROMPT = """\
あなたは「中沢頌（なかざわしょう）」の第二の人格として、彼の代わりに大学のレポートや課題を執筆するAIエージェントです。

## あなたの役割
- 中沢頌として、応用理工学類の課題・レポートを書く
- 彼の文体・表現・論理展開を忠実に再現する
- 物理、化学、数学、生物、実験レポートなど幅広い理工系科目に対応する

## 文体プロファイル
{style_profile}

## レポート作成の手順
1. **課題を理解する** - 何が求められているかを把握する
2. **調査する** - web_search で参考文献・定義・数値を調べる
3. **計算する** - run_python で数値計算・シミュレーションを行う
4. **グラフを作る** - generate_graph または run_python で図を生成する
5. **数式を整える** - format_latex で LaTeX 式を確定する
6. **文書を書く** - write_document で Word/PDF を出力する

## 重要なルール
- 文体は必ず style_profile に従う（中沢頌らしい文章）
- 数式は LaTeX 形式で書き、式番号を振る
- 参考文献は末尾にまとめる（番号引用方式）
- グラフには必ずキャプションと軸ラベルをつける
- 考察では誤差・不確かさ・仮定の限界を述べる
- 出力ファイルのパスを最後に必ず明示する
"""

def load_style_profile() -> str:
    if os.path.isfile(STYLE_PROFILE_PATH):
        with open(STYLE_PROFILE_PATH, encoding="utf-8") as f:
            return f.read()
    return (
        "※ 文体プロファイルがまだ生成されていません。"
        "style_analyzer.py を実行してください。\n"
        "それまでは「理工系大学生らしい、簡潔で論理的な常体（だ・である調）」で書いてください。"
    )


def build_system_prompt() -> str:
    profile = load_style_profile()
    return BASE_SYSTEM_PROMPT.format(style_profile=profile)


# ──────────────────────────────────────────────
# エージェントループ
# ──────────────────────────────────────────────

def run_agent(assignment: str, output_format: str = "docx", verbose: bool = True) -> str:
    """
    課題の指示を受け取り、レポートを生成してファイルパスを返す。

    Parameters
    ----------
    assignment   : str  - 課題の指示（何のレポートを書くか）
    output_format: str  - "docx" or "pdf"
    verbose      : bool - ツール呼び出しをコンソールに表示するか
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    system_prompt = build_system_prompt()

    # 課題メッセージにフォーマット指示を追加
    user_content = (
        f"{assignment}\n\n"
        f"【出力形式】{output_format.upper()}ファイルで出力すること。"
        f"著者名は「中沢頌」とする。"
    )

    messages = [{"role": "user", "content": user_content}]

    if verbose:
        print("\n" + "=" * 60)
        print("📝 レポートエージェント起動")
        print("=" * 60)
        print(f"課題: {assignment[:80]}{'...' if len(assignment) > 80 else ''}")
        print(f"出力形式: {output_format.upper()}")
        print("=" * 60 + "\n")

    generated_filepath = None
    iteration = 0
    max_iterations = 30

    while iteration < max_iterations:
        iteration += 1

        response = client.messages.create(
            model=MODEL,
            max_tokens=8192,
            system=system_prompt,
            tools=ALL_TOOLS,
            messages=messages,
        )

        # アシスタントメッセージを履歴に追加
        messages.append({"role": "assistant", "content": response.content})

        # テキスト出力を表示
        for block in response.content:
            if hasattr(block, "type") and block.type == "text" and block.text.strip():
                if verbose:
                    print(f"[Claude] {block.text}\n")

        # ツール呼び出し処理
        if response.stop_reason == "tool_use":
            tool_results = []

            for block in response.content:
                if not (hasattr(block, "type") and block.type == "tool_use"):
                    continue

                tool_name = block.name
                tool_input = block.input
                tool_use_id = block.id

                if verbose:
                    print(f"🔧 {tool_name}({json.dumps({k: str(v)[:80] for k, v in tool_input.items()}, ensure_ascii=False)})")

                try:
                    result_str = dispatch_tool(tool_name, tool_input)
                    result_data = json.loads(result_str)

                    # 生成されたファイルパスを記録
                    if tool_name == "write_document" and result_data.get("status") == "success":
                        generated_filepath = result_data.get("filepath")

                    if verbose:
                        status = result_data.get("status", "?")
                        icon = "✅" if status == "success" else "❌"
                        if tool_name == "run_python":
                            out = result_data.get("stdout", "")[:200]
                            print(f"  {icon} stdout: {out}")
                        elif tool_name == "web_search":
                            n = len(result_data.get("results", []))
                            print(f"  {icon} {n}件の検索結果")
                        elif tool_name == "generate_graph":
                            fp = result_data.get("filepath", "")
                            print(f"  {icon} 保存先: {fp}")
                        elif tool_name == "write_document":
                            fp = result_data.get("filepath", "")
                            print(f"  {icon} 保存先: {fp}")
                        else:
                            print(f"  {icon} {str(result_data)[:100]}")

                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": tool_use_id,
                        "content": result_str,
                    })

                except Exception as e:
                    error_msg = json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False)
                    if verbose:
                        print(f"  ❌ エラー: {e}")
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": tool_use_id,
                        "content": error_msg,
                        "is_error": True,
                    })

            messages.append({"role": "user", "content": tool_results})

        elif response.stop_reason == "end_turn":
            if verbose:
                print("\n" + "=" * 60)
                print("✅ レポート生成完了")
                if generated_filepath:
                    print(f"📄 出力ファイル: {generated_filepath}")
                print("=" * 60)
            break

        else:
            if verbose:
                print(f"[停止] stop_reason={response.stop_reason}")
            break

    return generated_filepath or ""


# ──────────────────────────────────────────────
# 対話モード
# ──────────────────────────────────────────────

def interactive_mode():
    print("\n🎓 中沢頌レポートエージェント（対話モード）")
    print("課題を入力してください。終了するには 'quit' と入力。\n")

    while True:
        print("-" * 60)
        assignment = input("📋 課題内容: ").strip()
        if assignment.lower() in ("quit", "exit", "q"):
            print("終了します。")
            break
        if not assignment:
            continue

        fmt = input("📁 出力形式 [docx/pdf] (デフォルト: docx): ").strip().lower()
        if fmt not in ("docx", "pdf"):
            fmt = "docx"

        run_agent(assignment, output_format=fmt)


if __name__ == "__main__":
    interactive_mode()
