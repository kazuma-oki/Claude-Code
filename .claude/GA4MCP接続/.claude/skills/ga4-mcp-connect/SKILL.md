---
name: ga4-mcp-connect
description: GA4 MCPサーバーをClaude Code（Windows/VSCode拡張）に接続するセットアップ手順。「GA4に接続したい」「analytics-mcpを設定したい」「GA4のMCPが動かない」などの場合に使用する。
---

## 前提条件

- Python 3.x + pipx インストール済み
- GCPプロジェクトが存在する
- GA4プロパティが存在する

---

## 手順

### 1. GCP側の準備

**APIを有効化**（Google Cloud Console）:
- Google Analytics Admin API
- Google Analytics Data API

**サービスアカウントを作成**:
1. IAM > サービスアカウント > 「作成」
2. JSONキーをダウンロードして安全な場所に保存

**GA4でアクセス許可**:
1. GA4管理画面 > プロパティのアクセス管理
2. サービスアカウントのメールアドレスを「閲覧者」で追加

---

### 2. analytics-mcpのインストール

```powershell
pipx install analytics-mcp
pipx ensurepath
```

インストール先: `C:\Users\<USER>\.local\bin\analytics-mcp.exe`

---

### 3. MCPサーバーの登録

**重要**: `~/.claude/settings.json` に直接書いても読まれない。  
Claude CLIは `~/.claude.json` を設定ファイルとして使うため、`claude mcp add` コマンドで登録する。

```powershell
# Claudeのexeパス（Windows app package）
$claude = "C:\Users\<USER>\AppData\Local\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude-code\<VERSION>\claude.exe"

# ユーザースコープで登録（グローバルに有効）
& $claude mcp add --scope user google-analytics `
  "C:\Users\<USER>\.local\bin\analytics-mcp.exe" `
  -e GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\key.json" `
  -e GOOGLE_PROJECT_ID="your-project-id"
```

スコープの選択:
- `--scope user` : ユーザー全体に適用（推奨）
- `--scope local` : カレントディレクトリのみ

---

### 4. 接続確認

```powershell
& $claude mcp list
# → google-analytics: ... - ✓ Connected が表示されればOK
```

---

### 5. VSCode拡張の再起動

VSCodeを完全に閉じて再度開く（`Developer: Reload Window` では不十分な場合がある）。

再起動後、チャットで以下のツールが利用可能になる:
- `mcp__google-analytics__get_account_summaries` — アカウント・プロパティ一覧
- `mcp__google-analytics__run_report` — カスタムレポート実行
- `mcp__google-analytics__run_realtime_report` — リアルタイムレポート
- `mcp__google-analytics__run_funnel_report` — ファネルレポート
- `mcp__google-analytics__run_conversions_report` — コンバージョンレポート
- `mcp__google-analytics__get_property_details` — プロパティ詳細
- `mcp__google-analytics__get_custom_dimensions_and_metrics` — カスタム指標
- `mcp__google-analytics__list_google_ads_links` — Google広告リンク
- `mcp__google-analytics__list_property_annotations` — アノテーション一覧

---

## 使用例

```
「GA4のプロパティ一覧を見せて」
「先週のセッション数を見せて」
「今月のページ別PVランキングTop10」
「デバイスカテゴリ別のコンバージョン数（過去30日）」
```

---

## トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| `/mcp` が使えない | VSCode拡張ではCLIコマンド非対応 | `claude mcp list` をターミナルで実行 |
| `mcp list` で「No servers configured」 | `settings.json` に書いても読まれない | `claude mcp add` で登録し直す |
| ツールがチャットに出てこない | 拡張が未再起動 | VSCodeを完全に再起動 |
| `analytics-mcp` が見つからない | PATHが未更新 | フルパスを使う or 新しいターミナルを開く |

---

## 実績環境

- OS: Windows 11 Pro
- Claude Code: VSCode拡張 v2.1.149
- analytics-mcp: v0.6.0
- Python: 3.12.4 / pipx: 1.12.0
- サービスアカウントキー: `C:\Users\kazug\.claude\GA4MCP接続\.claude\ga4-key\iron-axon-162706-69c5fb3aef9f.json`
- GCPプロジェクト: iron-axon-162706
- GA4プロパティ: n2p.co.jp_GA4 (305939297)
