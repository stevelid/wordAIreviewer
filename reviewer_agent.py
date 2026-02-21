import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path


DEFAULT_SYSTEM_PROMPT = """You are an expert Acoustic Engineering Reviewer. Your task is to review the provided acoustic report draft (formatted as Annotated Markdown) and make corrections for technical accuracy, clarity, standards compliance, and professional tone.

# EDITING FORMAT: SEARCH/REPLACE BLOCKS
You must output ALL your proposed changes using strict SEARCH/REPLACE blocks.

To change text, you must output a block formatted exactly like this:
<<<<<<< SEARCH
[Exact original text to be replaced, including any [ID] tags]
=======
[New modified text, keeping the [ID] tags intact]
>>>>>>> REPLACE

## CRITICAL RULES FOR SEARCH/REPLACE BLOCKS:
1. THE SEARCH BLOCK MUST BE EXACT: The text inside the `<<<<<<< SEARCH` section must be a literal, character-for-character copy of the original text from the document. Do not omit punctuation or spacing.
2. ALWAYS INCLUDE ID TAGS: You must include the `[P#]` or `[T#.R#.C#]` coordinate tags in BOTH the SEARCH and REPLACE blocks. The system relies on these tags to locate the edits.
3. NO LAZY EDITS: Never use placeholders like "..." or "/* rest of paragraph */". You must write out the entire modified paragraph or row in the REPLACE block.
4. KEEP EDITS ATOMIC: Do not bundle multiple distant paragraphs into a single block. Use a separate SEARCH/REPLACE block for each specific paragraph or table row you are editing.

## HANDLING TABLES
You will see tables annotated with cell coordinates, like `| [T2.R1.C2] 55 |`.
- NEVER attempt to redraw, reformat, or replace an entire table.
- To edit table data, you must target the specific row or specific cell using a SEARCH/REPLACE block.

Example of a correct table edit:
<<<<<<< SEARCH
| [T2.R3.C1] Pos 3 | [T2.R3.C2] 61 | [T2.R3.C3] High wind |
=======
| [T2.R3.C1] Pos 3 | [T2.R3.C2] 58 | [T2.R3.C3] High wind |
>>>>>>> REPLACE

## INLINE FORMATTING
You may use basic HTML tags `<b>`, `<i>`, `<sub>`, and `<sup>` in your REPLACE blocks to apply formatting.
Example: `[T1.R1.C1] 55 dB L<sub>Aeq, T</sub>`

## ADDING COMMENTS
If you need to ask the author a question, request clarification, or explain a complex change, you can attach a comment to a paragraph or cell.
Do this by appending `[COMMENT: Your message]` inside the REPLACE block.

For paragraphs, put it at the end of the text:
<<<<<<< SEARCH
[P15] The measurments was taken at 2:00 AM.
=======
[P15] The acoustic measurements were taken at 02:00. [COMMENT: Please confirm this was 02:00 and not 14:00.]
>>>>>>> REPLACE

For tables, put the comment INSIDE the specific cell's pipe delimiters:
<<<<<<< SEARCH
| [T1.R2.C2] 85 |
=======
| [T1.R2.C2] 85 [COMMENT: This exceeds the threshold.] |
>>>>>>> REPLACE

Review the document below and output your recommended SEARCH/REPLACE blocks."""


def load_system_prompt():
    """Loads prompt.txt from this script directory; falls back to default prompt."""
    prompt_path = Path(__file__).resolve().parent / "prompt.txt"
    if prompt_path.exists():
        prompt_text = read_text_file(prompt_path).strip()
        if prompt_text:
            return prompt_text
    return DEFAULT_SYSTEM_PROMPT


BLOCK_PATTERN = re.compile(
    r"<<<<<<< SEARCH\s*\n(.*?)\n=======\n(.*?)\n>>>>>>> REPLACE",
    re.DOTALL,
)
BRACKET_PATTERN = re.compile(r"\[([^\]]+)\]")
VALID_TARGET_PATTERN = re.compile(
    r"^(?:P\d+|T\d+(?:\.(?:H|R\d+)(?:\.C\d+)?)?)$",
    re.IGNORECASE,
)
CELL_TARGET_PATTERN = re.compile(r"^T\d+\.(?:H|R\d+)\.C\d+$", re.IGNORECASE)
ID_TAG_PATTERN = re.compile(
    r"\[(?:P\d+|T\d+(?:\.(?:H|R\d+)(?:\.C\d+)?)?)\]\s*",
    re.IGNORECASE,
)
COMMENT_PATTERN = re.compile(r"\[COMMENT:\s*(.*?)\s*\]", re.IGNORECASE | re.DOTALL)


def normalize_newlines(text):
    return text.replace("\r\n", "\n").replace("\r", "\n")


def get_temp_dir(custom_temp_dir):
    if custom_temp_dir:
        return Path(custom_temp_dir).expanduser().resolve()
    return Path(os.environ.get("TEMP", tempfile.gettempdir())) / "claude_review"


def resolve_dsm_input(file_arg, temp_dir):
    candidate = Path(file_arg).expanduser()
    if candidate.exists():
        dsm_path = candidate.resolve()
    else:
        if candidate.suffix.lower() == ".md":
            dsm_path = (temp_dir / candidate.name).resolve()
        else:
            dsm_path = (temp_dir / f"{file_arg}_dsm.md").resolve()

    if not dsm_path.exists():
        raise FileNotFoundError(f"DSM markdown not found: {dsm_path}")

    stem = dsm_path.stem
    if stem.lower().endswith("_dsm"):
        stem = stem[:-4]
    return dsm_path, stem


def read_text_file(path):
    try:
        return Path(path).read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Fallback for files exported by VBA's default ANSI encoder
        return Path(path).read_text(encoding="windows-1252", errors="replace")


def write_text_file(path, content):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(content, encoding="utf-8")


def write_json_file(path, payload):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def normalize_target(raw_target):
    target = raw_target.strip()
    return target.upper()


def is_valid_target(target):
    return bool(VALID_TARGET_PATTERN.match(target))


def is_cell_target(target):
    return bool(CELL_TARGET_PATTERN.match(target))


def extract_targets(text):
    targets = []
    seen = set()
    for match in BRACKET_PATTERN.finditer(text):
        candidate = normalize_target(match.group(1))
        if is_valid_target(candidate) and candidate not in seen:
            seen.add(candidate)
            targets.append(candidate)
    return targets


def strip_id_tags(text):
    cleaned = ID_TAG_PATTERN.sub("", text)
    return cleaned.strip()


def is_markdown_separator_line(line):
    body = line.strip()
    if not body.startswith("|"):
        return False
    body = body.strip("|").replace(" ", "")
    if not body:
        return False
    return all(ch in "-:" for ch in body)


def split_unescaped_pipes(row_body):
    return re.split(r"(?<!\\)\|", row_body)


def parse_markdown_row(line):
    raw = line.strip()
    if not raw.startswith("|"):
        return []
    row_body = raw.strip("|")
    raw_cells = split_unescaped_pipes(row_body)
    parsed = []
    for raw_cell in raw_cells:
        cell_text = raw_cell.strip()
        match = re.match(r"^\[([^\]]+)\]\s*(.*)$", cell_text)
        if match:
            cell_id = normalize_target(match.group(1))
            value = match.group(2).replace("\\|", "|").strip()
            parsed.append({"id": cell_id if is_valid_target(cell_id) else None, "text": value})
        else:
            parsed.append({"id": None, "text": cell_text.replace("\\|", "|")})
    return parsed


def extract_single_data_row(block_text):
    lines = [line.strip() for line in normalize_newlines(block_text).split("\n") if line.strip()]
    row_lines = []
    for line in lines:
        if not line.startswith("|"):
            continue
        if is_markdown_separator_line(line):
            continue
        row_lines.append(line)
    if len(row_lines) == 1:
        return row_lines[0]
    return None


def fix_markdown_formatting(text):
    """Converts stray markdown formatting into the HTML tags your VBA expects."""
    # Convert **bold** to <b>bold</b> (using negative lookbehind to avoid escaping)
    text = re.sub(r'(?<!\\)\*\*(.+?)(?<!\\)\*\*', r'<b>\1</b>', text)
    # Convert *italic* to <i>italic</i>
    text = re.sub(r'(?<!\\)\*(.+?)(?<!\\)\*', r'<i>\1</i>', text)
    return text


def make_replace_text_call(target, find_text, replace_text, block_index):
    return {
        "tool": "replace_text",
        "target": target,
        "args": {
            "find": find_text,  # Keep exact to match document
            "replace": fix_markdown_formatting(replace_text),  # Sanitize new text
        },
        "explanation": f"SEARCH/REPLACE block {block_index}",
    }


def make_add_comment_call(target, comment_text):
    return {
        "tool": "add_comment",
        "target": target,
        "args": {"text": comment_text},
    }


def convert_row_block(search_row, replace_row, block_index, warnings):
    search_cells = parse_markdown_row(search_row)
    replace_cells = parse_markdown_row(replace_row)
    if not search_cells or not replace_cells:
        return []

    if len(search_cells) != len(replace_cells):
        warnings.append(
            f"Block {block_index}: row cell counts differ ({len(search_cells)} vs {len(replace_cells)}); skipped row diff."
        )
        return []

    tool_calls = []
    for idx, (left, right) in enumerate(zip(search_cells, replace_cells), start=1):
        left_id = left.get("id")
        right_id = right.get("id")
        if not left_id or not right_id:
            continue
        if left_id != right_id:
            warnings.append(f"Block {block_index}: row cell {idx} ID mismatch ({left_id} vs {right_id}); skipped.")
            continue
        if not is_cell_target(left_id):
            continue

        find_text = left.get("text", "").strip()
        replace_text = right.get("text", "").strip()

        comment_match = COMMENT_PATTERN.search(replace_text)
        if comment_match:
            comment_text = comment_match.group(1).strip()
            if comment_text:
                tool_calls.append(make_add_comment_call(left_id, comment_text))
            replace_text = COMMENT_PATTERN.sub("", replace_text).strip()

        if find_text != replace_text:
            tool_calls.append(make_replace_text_call(left_id, find_text, replace_text, block_index))

    return tool_calls


def convert_generic_block(search_text, replace_text, block_index, warnings):
    targets = extract_targets(search_text)
    if not targets:
        warnings.append(f"Block {block_index}: no valid [ID] target found; skipped.")
        return []

    target = targets[0]
    if target.startswith("T") and (target.endswith(".H") or re.match(r"^T\d+\.R\d+$", target, re.IGNORECASE)):
        warnings.append(f"Block {block_index}: row-level target {target} needs row diff handling; skipped.")
        return []

    find_text = strip_id_tags(search_text)
    replace_value = strip_id_tags(replace_text)
    tool_calls = []

    comment_match = COMMENT_PATTERN.search(replace_value)
    if comment_match:
        comment_text = comment_match.group(1).strip()
        if comment_text:
            tool_calls.append(make_add_comment_call(target, comment_text))
        replace_value = COMMENT_PATTERN.sub("", replace_value).strip()

    if find_text != replace_value and find_text:
        tool_calls.append(make_replace_text_call(target, find_text, replace_value, block_index))
    elif not find_text and not tool_calls:
        warnings.append(f"Block {block_index}: empty find text after stripping IDs; skipped.")

    return tool_calls


def parse_search_replace_blocks(llm_output):
    text = normalize_newlines(llm_output)
    return list(BLOCK_PATTERN.finditer(text))


def build_tool_calls_from_response(llm_output):
    matches = parse_search_replace_blocks(llm_output)
    warnings = []
    tool_calls = []

    for idx, match in enumerate(matches, start=1):
        search_text = match.group(1).strip("\n")
        replace_text = match.group(2).strip("\n")

        search_row = extract_single_data_row(search_text)
        replace_row = extract_single_data_row(replace_text)

        if search_row and replace_row:
            row_calls = convert_row_block(search_row, replace_row, idx, warnings)
            if row_calls:
                tool_calls.extend(row_calls)
                continue

        generic_calls = convert_generic_block(search_text, replace_text, idx, warnings)
        tool_calls.extend(generic_calls)

    return tool_calls, warnings, len(matches)


def run_cmd(cmd, input_text, cwd, stream_output=False):
    if not stream_output:
        proc = subprocess.run(
            cmd,
            input=input_text,
            text=True,
            capture_output=True,
            cwd=str(cwd),
            encoding="utf-8",      # Force UTF-8 reading
            errors="replace"       # Silently replace broken characters instead of crashing
        )
        return proc.returncode, proc.stdout, proc.stderr

    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        cwd=str(cwd),
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )

    stdout_chunks = []
    stderr_chunks = []

    def _pump(stream, sink, out_stream):
        try:
            for line in iter(stream.readline, ""):
                sink.append(line)
                out_stream.write(line)
                out_stream.flush()
        finally:
            stream.close()

    t_out = threading.Thread(target=_pump, args=(proc.stdout, stdout_chunks, sys.stdout), daemon=True)
    t_err = threading.Thread(target=_pump, args=(proc.stderr, stderr_chunks, sys.stderr), daemon=True)
    t_out.start()
    t_err.start()

    try:
        if proc.stdin is not None:
            proc.stdin.write(input_text)
            proc.stdin.close()
    except BrokenPipeError:
        pass

    returncode = proc.wait()
    t_out.join()
    t_err.join()
    return returncode, "".join(stdout_chunks), "".join(stderr_chunks)


def run_codex(markdown_text, model, cwd):
    if shutil.which("codex") is None:
        raise RuntimeError("codex CLI not found on PATH.")

    system_prompt = load_system_prompt()
    prompt = (
        f"{system_prompt}\n\n"
        "The document to review is below. Output only SEARCH/REPLACE blocks.\n\n"
        f"{markdown_text}"
    )

    with tempfile.NamedTemporaryFile(prefix="codex_last_", suffix=".txt", delete=False) as tmp:
        last_message_path = Path(tmp.name)

    try:
        cmd = ["codex", "exec", "-", "--skip-git-repo-check", "--output-last-message", str(last_message_path)]
        if model:
            cmd.extend(["--model", model])

        code, stdout, stderr = run_cmd(cmd, prompt, cwd)
        if code != 0:
            raise RuntimeError(f"codex exec failed ({code}): {stderr.strip() or stdout.strip()}")

        if last_message_path.exists():
            response = last_message_path.read_text(encoding="utf-8").strip()
            if response:
                return response
        if stdout.strip():
            return stdout.strip()
        raise RuntimeError("codex returned no output.")
    finally:
        if last_message_path.exists():
            last_message_path.unlink()


def run_claude(markdown_text, model, cwd):
    if shutil.which("claude") is None:
        raise RuntimeError("claude CLI not found on PATH.")

    system_prompt = load_system_prompt()
    cmd = [
        "claude",
        "-p",
        "--output-format",
        "text",
        "--system-prompt",
        system_prompt,
    ]
    if model:
        cmd.extend(["--model", model])

    code, stdout, stderr = run_cmd(cmd, markdown_text, cwd, stream_output=True)
    if code != 0:
        raise RuntimeError(f"claude -p failed ({code}): {stderr.strip() or stdout.strip()}")
    if stdout.strip():
        return stdout.strip()
    raise RuntimeError("claude returned no output.")


def run_llm(markdown_text, runner, model, cwd):
    runner = runner.lower()
    errors = []

    if runner in ("auto", "codex"):
        try:
            return run_codex(markdown_text, model, cwd), "codex"
        except Exception as exc:
            errors.append(f"codex: {exc}")
            if runner == "codex":
                raise RuntimeError(errors[-1])

    if runner in ("auto", "claude"):
        try:
            return run_claude(markdown_text, model, cwd), "claude"
        except Exception as exc:
            errors.append(f"claude: {exc}")
            if runner == "claude":
                raise RuntimeError(errors[-1])

    if errors:
        raise RuntimeError(" ; ".join(errors))
    raise RuntimeError(f"Unsupported runner: {runner}")


def parse_args():
    parser = argparse.ArgumentParser(
        description="V5 reviewer orchestrator: DSM markdown -> SEARCH/REPLACE -> V4 tool_calls JSON."
    )
    parser.add_argument("--file", required=True, help="Document stem or path to _dsm.md file.")
    parser.add_argument(
        "--runner",
        choices=["auto", "codex", "claude"],
        default="auto",
        help="LLM runner for automatic mode (default: auto).",
    )
    parser.add_argument("--model", default="", help="Optional model override for codex/claude.")
    parser.add_argument("--temp-dir", default="", help="Override temp exchange directory.")
    parser.add_argument("--manual-output-file", default="", help="Parse saved LLM output text file.")
    parser.add_argument(
        "--manual-paste",
        action="store_true",
        help="Read LLM output from stdin until EOF and parse it (manual mode).",
    )
    parser.add_argument("--raw-output", default="", help="Optional path for raw LLM response output.")
    parser.add_argument("--toolcalls-out", default="", help="Optional path for output toolcalls JSON.")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.manual_output_file and args.manual_paste:
        print("ERROR: Use either --manual-output-file or --manual-paste, not both.", file=sys.stderr)
        return 2

    temp_dir = get_temp_dir(args.temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        dsm_path, stem = resolve_dsm_input(args.file, temp_dir)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    try:
        print(f"[1/4] Reading DSM markdown: {dsm_path}")
        markdown_text = read_text_file(dsm_path)
    except Exception as exc:
        print(f"ERROR: failed reading DSM markdown: {exc}", file=sys.stderr)
        return 2

    raw_output_path = Path(args.raw_output) if args.raw_output else (temp_dir / f"{stem}_llm_response.txt")
    toolcalls_path = Path(args.toolcalls_out) if args.toolcalls_out else (temp_dir / f"{stem}_toolcalls.json")

    manual_mode = bool(args.manual_output_file or args.manual_paste)
    llm_output = ""
    used_runner = "manual"

    try:
        if args.manual_output_file:
            print(f"[2/4] Reading manual LLM output file: {args.manual_output_file}")
            llm_output = read_text_file(args.manual_output_file)
        elif args.manual_paste:
            print("[2/4] Reading manual LLM output from stdin...")
            llm_output = sys.stdin.read()
            if not llm_output.strip():
                raise RuntimeError("No input received on stdin for --manual-paste.")
        else:
            print(f"[2/4] Running local LLM via runner='{args.runner}'...")
            llm_output, used_runner = run_llm(markdown_text, args.runner, args.model.strip(), dsm_path.parent)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    try:
        print(f"[3/4] Writing raw response: {raw_output_path}")
        write_text_file(raw_output_path, llm_output)
    except Exception as exc:
        print(f"ERROR: failed writing raw output file: {exc}", file=sys.stderr)
        return 1

    print("[4/4] Parsing SEARCH/REPLACE blocks and building tool calls...")
    tool_calls, warnings, block_count = build_tool_calls_from_response(llm_output)
    payload = {"tool_calls": tool_calls}

    try:
        write_json_file(toolcalls_path, payload)
    except Exception as exc:
        print(f"ERROR: failed writing tool calls file: {exc}", file=sys.stderr)
        return 1

    mode_label = "manual" if manual_mode else f"runner:{used_runner}"
    print(f"Mode: {mode_label}")
    print(f"DSM markdown: {dsm_path}")
    print(f"Blocks parsed: {block_count}")
    print(f"Tool calls generated: {len(tool_calls)}")
    print(f"Raw output: {raw_output_path}")
    print(f"Tool calls JSON: {toolcalls_path}")

    if warnings:
        print("Warnings:")
        for warning in warnings:
            print(f"- {warning}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
