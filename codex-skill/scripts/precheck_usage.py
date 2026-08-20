from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(description="Create or update a per-study token usage ledger.")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--study", required=True)
    parser.add_argument("--session", required=True)
    parser.add_argument("--role", required=True, choices=["director", "research_ethics", "factcheck", "editor", "other"])
    parser.add_argument("--input", type=int)
    parser.add_argument("--cached-input", type=int, default=0)
    parser.add_argument("--output", type=int)
    parser.add_argument("--unavailable-reason")
    parser.add_argument("--source", default="local_rollout_final_total")
    args = parser.parse_args()
    payload = {"schema_version": 1, "study": args.study, "entries": []}
    if args.out.exists():
        payload = json.loads(args.out.read_text(encoding="utf-8"))
    payload["entries"] = [e for e in payload.get("entries", []) if not (e.get("session") == args.session and e.get("role") == args.role)]
    if args.unavailable_reason:
        entry = {
            "recorded_at": datetime.now(timezone.utc).isoformat(timespec="seconds"), "session": args.session,
            "role": args.role, "status": "unavailable", "reason": args.unavailable_reason, "source": args.source,
        }
    else:
        if args.input is None or args.output is None:
            parser.error("--input and --output are required unless --unavailable-reason is provided")
        entry = {
        "recorded_at": datetime.now(timezone.utc).isoformat(timespec="seconds"), "session": args.session,
        "role": args.role, "input_tokens": args.input, "cached_input_tokens": args.cached_input,
        "output_tokens": args.output, "total_tokens": args.input + args.output, "source": args.source,
        }
    payload["entries"].append(entry)
    totals = {key: sum(int(e.get(key, 0)) for e in payload["entries"]) for key in ("input_tokens", "cached_input_tokens", "output_tokens", "total_tokens")}
    totals["unavailable_entries"] = sum(e.get("status") == "unavailable" for e in payload["entries"])
    totals["complete"] = totals["unavailable_entries"] == 0
    payload["totals"] = totals
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(totals, ensure_ascii=False))


if __name__ == "__main__":
    main()
