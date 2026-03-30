# Room Authoring

Room content is authored in `src/game/data/rooms.ts`.

## Platforms

Platforms use center-based coordinates:

- `x`, `y`: center point
- `width`, `height`: rendered and logical size

The platform top is `y - height / 2`. Ground monsters should usually sit with their feet close to that top edge.

## Ladders

Ladders also use center-based coordinates:

- `x`, `y`: center point
- `width`, `height`: visual ladder size

The gameplay ladder grab zone is intentionally taller and wider than the art, especially at the top, so Dan can climb down from a platform without pixel-perfect alignment.

## Hazards

Hazards use these important fields:

- `x`, `y`: initial sprite position
- `width`, `height`: rendered size and movement-bound input
- `minX`, `maxX`: coarse lane hint, not the final movement span

At room load, `GameScene` resolves the platform that supports the hazard and clamps the hazard's runtime travel to a safe platform span. Author `minX` and `maxX` as the intended lane, then let the runtime clamp prevent edge overhang.

### Ground monsters

- Place `paul_crab` and `dave_goat` so their bottoms are near the platform top.
- Keep their lane centered on the platform they should patrol.
- Do not hand-tune lanes right up to platform edges unless the danger is part of the design.

### Flying monsters

- Place `mark_wasp` at the desired baseline flight height.
- `swoopDepth` controls vertical amplitude.
- `swoopRate` controls vertical tempo.
- Horizontal motion still uses a platform-derived safe lane, so keep the wasp visually associated with one platform span.

## Spawn points

Spawn points are safety-critical.

- `default` should always land Dan on supported floor
- `fromLeft` and `fromRight` should land on the correct side of a room transition without placing Dan in a pit or inside a ladder
- Use the same platform-top rule as player landing logic when choosing `y`

If a room starts with instant death, inspect spawn coordinates before changing physics.
