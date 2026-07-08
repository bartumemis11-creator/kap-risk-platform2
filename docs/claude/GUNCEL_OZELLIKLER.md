# Claude Code — Güncel Özellikler ve Doğrulanmış Gerçekler (2026)

CLAUDE.md'den @import ile bağlanır. Kaynak: resmî Claude Code dokümanı + `claude-api`.

## Güncel model ID'leri (yeni kod için)

| Model | ID | Fiyat (in/out /1M) |
|---|---|---|
| Opus 4.8 (varsayılan) | `claude-opus-4-8` | $5 / $25 |
| Sonnet 5 | `claude-sonnet-5` | $3 / $15 (tanıtım $2/$10, 31.08.2026'ya kadar) |
| Fable 5 (en yetenekli) | `claude-fable-5` | $10 / $50 |
| Haiku 4.5 | `claude-haiku-4-5` | $1 / $5 |

`claude-opus-4-7` / `claude-sonnet-4-6` hâlâ aktif ama önceki nesil.

## Hafıza seviyeleri ve @import

- Kullanıcı: `~/.claude/CLAUDE.md` · Proje: `<repo>/CLAUDE.md` · Yerel: `CLAUDE.local.md` (git dışı)
- `@docs/x.md` ile başka dosyayı CLAUDE.md'ye gömersin (200 satır sınırını aşmadan detay taşı).
- Oturum içi hızlı hafıza: mesaj başına `#`.

## .mcp.json (proje kökü)

- Transport: `stdio` (yerel süreç), `http`, `sse`. Sırlar `${ENV}` ile.
- Scope: `local` (varsayılan), `project` (paylaşılır), `user` (tüm projeler).
- Yetki: `/mcp` interaktif menü.

## Hook bloklama (yaygın hata)

- `"blocking": true` **YOK**. PreToolUse'da **exit code 2** tool'u bloklar (stderr modele döner).
- Alternatif: JSON stdout `{"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "deny"}}`.
- Exit 0 = izin (varsayılan).

## İzinler (settings.json)

- Eşleşme **önek tabanlı**: `Bash(git push:*)`. `Bash(*prod*)` gibi ortadan/baştan joker **çalışmaz**.
- Kesin/desen tabanlı bloklama için `deny` yerine **PreToolUse hook + exit 2**.
- settings.json için resmî `$schema` URL'i yok.

## plugin.json

- Sadece metadata (name, version, description, author, license). `skills`/`agents`/`commands` klasörleri
  **otomatik keşfedilir**; elle dizi listelemek gerekmez. Dağıtım: `marketplace.json` + `/plugin`.

## Subagent / komut

- Subagent dosyaları `.claude/agents/`. `tools:` hem YAML listesi hem virgüllü string olur.
- Komut placeholder'ları: `$ARGUMENTS`, `$1`, `$2`. `/compact <talimat>` geçerli.
- Yanlış komutlar: `/save` yok → **`/export`**; `claude --auto-compact 30` yok.

## Diğer güncel özellikler

- **Output styles** (`/output-style`, `.claude/output-styles/`)
- **Statusline** (`statusLine` ayarı)
- **Checkpoints / geri sarma** (Esc-Esc)
- **Extended thinking** ("think" / "ultrathink" tetikleyicileri)
- **Background bash** (Ctrl-B)
