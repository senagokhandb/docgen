"""What the chat UI and the MCP service share about a conversation with the agent.

A conversation is a chain of messages; the deliverables and the scratch/ of earlier
messages come along to the next one (uploads/ and /workspace/previous/<n>/). This module
builds the prompt of a message, reads the agent's stream-json events into steps, and
cleans file names. Standard library only: the UI (ui/server.py) imports it as well.
"""
import json
import re
import unicodedata
from pathlib import Path

EVENTS = ".docgen-events.jsonl"  # stream-json events; entrypoint.sh writes them (DOCGEN_OUT_FILE)
PREVIOUS = "/workspace/previous"  # scratch/ of earlier messages, mounted read-only (QWEN.md knows this path)
HISTORY_TURNS = 6

OFFICE = {".docx", ".dotx", ".docm", ".pptx", ".potx", ".pptm", ".xlsx", ".xltx", ".xlsm",
          ".doc", ".ppt", ".xls", ".rtf", ".odt", ".odp", ".ods"}
SOURCES = {".csv", ".txt", ".md", ".json", ".png", ".jpg", ".jpeg"}  # inputs only (data, logo, notes)
UPLOADABLE = OFFICE | {".pdf"} | SOURCES
DELIVERABLE = OFFICE | {".pdf", ".csv"}


def clip(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[:n] + f" … (+{len(s) - n} karakter)"


def clean_name(name):
    name = unicodedata.normalize("NFC", str(name)).replace("\\", "/").split("/")[-1]
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)
    name = re.sub(r"\s+", "_", name.strip()).lstrip(".")
    name = re.sub(r"[<>:\"|?*'`$;&]", "_", name)
    p = Path(name or "dosya")
    return p.stem[:120] + p.suffix.lower()[:10]


def unique_name(taken, name):
    if name not in taken:
        return name
    p, i = Path(name), 2
    while f"{p.stem}-v{i}{p.suffix}" in taken:
        i += 1
    return f"{p.stem}-v{i}{p.suffix}"


def no_flags(text):
    """History must not carry a flag over: an old '--onay' is not consent for this message."""
    return re.sub(r"--(onay|visual-qa)\b", r"\1", str(text or ""))


def entries(path):
    """Top-level names in a directory (folders end with '/'); [] if it is missing."""
    try:
        return sorted(p.name + ("/" if p.is_dir() else "") for p in path.iterdir() if not p.name.startswith("."))
    except (FileNotFoundError, NotADirectoryError):
        return []


# ---- agent events -> steps --------------------------------------------------
def read_events(path):
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return []
    events = []
    for line in text.splitlines():
        if line.strip():
            try:
                events.append(json.loads(line))
            except json.JSONDecodeError:
                pass  # the last line may still be being written
    return events


def events_from_stdout(text):
    """Images built before DOCGEN_OUT_FILE print one JSON array when the run ends."""
    i = text.find("[{")
    if i < 0:
        return [json.loads(x) for x in text.splitlines() if x.startswith("{") and x.rstrip().endswith("}")]
    try:
        return json.loads(text[i:])
    except json.JSONDecodeError:  # cut-off output: keep every complete event
        events, dec, pos, s = [], json.JSONDecoder(), 1, text[i:]
        while pos < len(s):
            while pos < len(s) and s[pos] in " ,\n\r\t":
                pos += 1
            try:
                obj, pos = dec.raw_decode(s, pos)
            except json.JSONDecodeError:
                break
            events.append(obj)
        return events


def short_path(p):
    p = str(p or "")
    for pre, rep in (("/home/agent/.qwen/skills/", "skills/"), ("/opt/docgen/", ""), ("/workspace/", "")):
        if p.startswith(pre):
            return rep + p[len(pre):]
    return p


def tool_step(name, inp):
    if name == "run_shell_command":
        return {"k": "shell", "label": clip(inp.get("command"), 800), "note": clip(inp.get("description"), 200)}
    if name == "read_file":
        return {"k": "read", "label": short_path(inp.get("file_path") or inp.get("absolute_path"))}
    if name == "write_file":
        return {"k": "write", "label": short_path(inp.get("file_path")), "detail": clip(inp.get("content"), 3000)}
    if name in ("edit", "replace"):
        return {"k": "edit", "label": short_path(inp.get("file_path")), "detail": clip(inp.get("new_string"), 3000)}
    if name == "skill":
        return {"k": "skill", "label": str(inp.get("skill") or "")}
    if name in ("glob", "grep_search"):
        return {"k": "search", "label": str(inp.get("pattern") or "")}
    return {"k": "tool", "label": f"{name} {clip(json.dumps(inp, ensure_ascii=False), 300)}"}


def result_text(content):
    if isinstance(content, list):
        return "\n".join(c.get("text", "") if isinstance(c, dict) else str(c) for c in content)
    return str(content or "")


def summarize(events):
    """-> (steps, result|None). result: reply, is_error, error, duration_ms, num_turns."""
    steps, by_id, result = [], {}, None
    for e in events:
        kind = e.get("type")
        msg = e.get("message") if isinstance(e.get("message"), dict) else {}
        for c in msg.get("content") or []:
            if not isinstance(c, dict):
                continue
            if kind == "assistant" and c.get("type") == "text" and str(c.get("text") or "").strip():
                steps.append({"k": "text", "label": clip(c["text"].strip(), 2000)})
            elif kind == "assistant" and c.get("type") == "tool_use":
                step = tool_step(c.get("name"), c.get("input") or {})
                by_id[c.get("id")] = step
                steps.append(step)
            elif kind == "user" and c.get("type") == "tool_result" and c.get("tool_use_id") in by_id:
                step = by_id[c["tool_use_id"]]
                step["ok"] = not c.get("is_error")
                step["out"] = clip(result_text(c.get("content")).strip(), 1500)
        if kind == "result":
            err = e.get("error")
            result = {"reply": str(e.get("result") or "").strip(), "is_error": bool(e.get("is_error")),
                      "error": (err.get("message") if isinstance(err, dict) else err) or "",
                      "duration_ms": e.get("duration_ms"), "num_turns": e.get("num_turns")}
    if result and steps and steps[-1]["k"] == "text" and steps[-1]["label"] == clip(result["reply"], 2000):
        steps.pop()  # the final answer is shown as the reply
    return steps, result


# ---- prompt -----------------------------------------------------------------
def build_prompt(sess, turn, work=None):
    """sess: {"turns": [...], "files": [...]}; turn: the new message.

    A turn has n, text, status, reply, error, outputs (and confirm, active for the new
    one); a file has name, origin ("upload" | "output") and turn. work: {n: entries of
    message n's scratch/}, mounted at PREVIOUS/<n>/ by the caller.
    """
    request = turn["text"] + (" --onay" if turn.get("confirm") else "")
    history = [t for t in sess["turns"] if t["n"] < turn["n"] and t["status"] != "running"][-HISTORY_TURNS:]
    files = [f for f in sess["files"] if f["turn"] is not None and f["turn"] <= turn["n"]]
    work = work or {}
    if not history and not files and not work:
        return request
    lines = ["[Sohbet arayüzü] Bu istek süren bir sohbetin parçasıdır.", ""]
    if history:
        lines.append("Önceki mesajlar (yalnızca bağlam; bunları yeniden yapma):")
        for t in history:
            lines.append(f"{t['n']}. Kullanıcı: {no_flags(t['text'])}")
            if t["status"] != "done":  # say it plainly: an error text read as "my answer" looks like finished work
                why = t.get("error") or ("kullanıcı durdurdu" if t["status"] == "cancelled" else t["status"])
                lines.append(f"   Bu mesaj tamamlanmadı — {clip(no_flags(why), 300)}")
            elif t.get("reply"):
                lines.append(f"   Senin yanıtın: {clip(no_flags(t['reply']), 1200)}")
            if t.get("outputs"):
                lines.append(f"   Teslim edilen dosyalar: {', '.join(t['outputs'])}")
        lines.append("")
    if work:
        lines.append("Önceki mesajların çalışma dosyaları (o mesajın /workspace/scratch/'ı; salt okunur):")
        for t in sess["turns"]:
            if t["n"] not in work:
                continue
            if t["status"] != "done":
                what = " (tamamlanmadı: yarım, doğrulanmamış iş)"
            elif t.get("outputs"):
                what = f" (teslim: {', '.join(t['outputs'])})"
            else:
                what = ""
            names = work[t["n"]]
            more = f" … (+{len(names) - 12})" if len(names) > 12 else ""
            lines.append(f"- {PREVIOUS}/{t['n']}/{what}: {', '.join(names[:12])}{more}")
        lines.append("Bir teslimatın kaynağı (PDF'in render edildiği HTML, belgeyi kuran script) burada duruyorsa "
                     "kullanıcı o kaynağı da yüklemiş sayılır. Kaynak yalnızca aynı mesajın teslim ettiği dosyaya "
                     "aittir: düzenlenecek dosyayı teslim eden mesajın klasöründe kaynak yoksa daha eski bir mesajın "
                     "kaynağını kullanma (dosya o zamandan beri değişmiş olabilir).")
        lines.append("")
    if files:
        lines.append("/workspace/uploads/ içindeki dosyalar (salt okunur):")
        for f in files:
            if f["origin"] == "output":
                how = f"{f['turn']}. mesajda senin teslim ettiğin dosya"
            elif f["turn"] == turn["n"]:
                how = "kullanıcı bu mesajla yükledi"
            else:
                how = f"kullanıcı {f['turn']}. mesajda yükledi"
            lines.append(f"- {f['name']} — {how}")
        if turn.get("active"):
            lines.append(f"Kullanıcının ekranda açık tuttuğu dosya: {turn['active']} (istek bir dosyada "
                         "değişiklik istiyor ama dosyayı adıyla söylemiyorsa bu dosyayı kastediyor).")
        if any(f["origin"] == "output" for f in files):
            lines.append("Önceki teslimatların da uploads/ altında: onları düzenlemek, yüklenmiş bir dosyayı "
                         "düzenlemekle aynıdır (yerinde düzenle, sonucu /workspace/outputs/ altına yaz).")
        lines.append("")
    lines += ["Yeni istek:", request]
    return "\n".join(lines)
