import os

# ========== パス設定 ==========
# OneDriveのドキュメントフォルダ（WSL経由 or ローカル実行時に変更）
DOCS_DIR = os.environ.get(
    "DOCS_DIR",
    "/mnt/c/Users/shonk/OneDrive/ドキュメント"
)

STYLE_PROFILE_PATH = os.path.join(os.path.dirname(__file__), "style_profile.md")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")

# ========== Anthropic設定 ==========
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = "claude-opus-4-7"

# ========== 読み込むファイルの拡張子 ==========
SUPPORTED_EXTENSIONS = [".docx", ".pdf", ".xlsx", ".txt", ".md"]

# ========== スタイル解析 ==========
# 一度に解析するファイル数（大きすぎるとコンテキスト超過）
BATCH_SIZE = 10
# 1ファイルから読み込む最大文字数
MAX_CHARS_PER_FILE = 3000

# ========== レポート出力 ==========
# "docx" or "pdf"
DEFAULT_OUTPUT_FORMAT = "docx"
