"""Generates one photorealistic image per card in cards.json using OpenAI gpt-image-1.

Usage:
  python generate_images.py              # all cards still missing an image
  python generate_images.py --limit 4    # only the first 4 cards (in page/position order)
  python generate_images.py --only agua leite   # only these keys (regenerates if --force)
  python generate_images.py --force ...  # overwrite existing images

Reads the API key from the OPENAI_API_KEY environment variable (never stored in files).
"""
import argparse, base64, json, os, pathlib, sys, time

HERE = pathlib.Path(__file__).resolve().parent
MODEL = "gpt-image-1"
SIZE = "1024x1024"
MAX_TRIES = 3


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, help="only the first N cards")
    ap.add_argument("--only", nargs="+", metavar="KEY", help="only these card keys")
    ap.add_argument("--force", action="store_true", help="overwrite existing images")
    args = ap.parse_args()

    if not os.environ.get("OPENAI_API_KEY"):
        sys.exit("OPENAI_API_KEY is not set. Set it in the environment and run again.")

    from openai import OpenAI
    client = OpenAI()

    data = json.loads((HERE / "cards.json").read_text(encoding="utf-8"))
    cards = sorted(data["cards"], key=lambda c: (c["page"], c["position"]))
    if args.only:
        cards = [c for c in cards if c["key"] in set(args.only)]
    if args.limit:
        cards = cards[: args.limit]

    total, failed = len(cards), []
    for i, c in enumerate(cards, 1):
        out = HERE / c["image_file"]
        name = out.name
        if out.exists() and not args.force:
            print(f"{i}/{total} {name} (já existe, saltado)")
            continue
        out.parent.mkdir(parents=True, exist_ok=True)
        for attempt in range(1, MAX_TRIES + 1):
            try:
                r = client.images.generate(model=MODEL, prompt=c["image_prompt"], size=SIZE, n=1)
                tmp = out.with_suffix(".part")
                tmp.write_bytes(base64.b64decode(r.data[0].b64_json))
                tmp.replace(out)  # atomic: no half-written PNG if interrupted
                print(f"{i}/{total} {name}")
                break
            except Exception as e:
                print(f"{i}/{total} {name} — erro (tentativa {attempt}/{MAX_TRIES}): {e}")
                if attempt < MAX_TRIES:
                    time.sleep(5 * attempt)
        else:
            failed.append(name)

    if failed:
        print("\nFalharam:", ", ".join(failed))
        sys.exit(1)
    print("\nConcluído.")


if __name__ == "__main__":
    main()
