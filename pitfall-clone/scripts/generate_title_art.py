from __future__ import annotations

import argparse
import base64
import os
from pathlib import Path

from openai import OpenAI


DEFAULT_PROMPT = """Use case: illustration-story
Asset type: retro game title screen key art
Primary request: Create a faux Atari-era adventure box-art painting for the game Temple Runaway.
Scene/backdrop: moonlit jungle temple at night, weathered wooden platforms in the foreground, ruined stone temple silhouette, humid mist, dramatic sky.
Subject: Dan the Hobbit front and center as a heroic but slightly comic adventurer, with curly brown hair, expressive face, pale shirt, brown waistcoat, short hobbit proportions, bare feet. Include Paul the angry crab, Dave the black goat, Mark the demented wasp, and several glowing golden clams with blue eyes.
Style/medium: polished painted poster illustration, pseudo-photoreal, deliberately much more dramatic and refined than the actual game graphics, inspired by late 1970s to early 1980s adventure game box art.
Composition/framing: landscape poster composition, Dan dominant in the center, monsters staged around him in action poses, clams sparkling, strong depth and theatrical staging, clean negative space near the top and bottom for game title and start prompt overlays.
Lighting/mood: dramatic moonlight, warm reflected glow from the clams, adventurous, campy, overpromised, slightly absurd heroic seriousness.
Color palette: deep jungle teals and greens, warm moon gold, rich wood browns, ember reds, antique parchment highlights.
Materials/textures: painted foliage, mist, weathered wood grain, polished shell highlights, furry goat coat, glossy crab shell, translucent insect wings.
Constraints: no embedded text, no logos, no UI, no watermark, no border. Make Dan, the crab, the goat, and the wasp visually distinct and readable.
Avoid: modern sci-fi elements, realistic photography, generic fantasy armor, extra characters, blur, muddy composition.
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate high-resolution title art with the OpenAI Images API.")
    parser.add_argument(
        "--prompt-file",
        type=Path,
        default=Path("docs/title-art-prompt.txt"),
        help="Optional text file to override the default prompt.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path("output/imagegen/title-box-art-api.png"),
        help="Output image path.",
    )
    parser.add_argument("--size", default="1536x1024", help="Image size accepted by the API.")
    parser.add_argument("--quality", default="high", help="Generation quality.")
    return parser.parse_args()


def load_prompt(path: Path) -> str:
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    return DEFAULT_PROMPT.strip()


def main() -> None:
    args = parse_args()
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise SystemExit("OPENAI_API_KEY is not set. Set it in your environment, then rerun this script.")

    prompt = load_prompt(args.prompt_file)
    client = OpenAI(api_key=api_key)

    result = client.images.generate(
        model="gpt-image-1",
        prompt=prompt,
        size=args.size,
        quality=args.quality,
    )

    image_b64 = result.data[0].b64_json
    if not image_b64:
        raise SystemExit("The API response did not contain image data.")

    output_path = args.out.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(base64.b64decode(image_b64))
    print(output_path)


if __name__ == "__main__":
    main()
