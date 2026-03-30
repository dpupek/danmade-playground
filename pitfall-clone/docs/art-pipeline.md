# Art Pipeline

All art is generated in `src/game/scenes/BootScene.ts`. There are no external sprite sheets in the current build.

## Current generated textures

- `player-stand`
- `player-jump`
- `treasure`
- `paul-crab`
- `dave-goat`
- `mark-wasp`
- ladder and ground textures

## Working style

Favor silhouette-first shapes that read at small sizes:

- Dan should read as a hobbit at gameplay scale, not only when zoomed in
- monsters should have distinct outlines before color details matter
- ladders should read as climbable structures, not translucent bars
- relics should pop against the background

## Gameplay coupling

When changing art, also check:

- `setDisplaySize` calls in `GameScene`
- spawn heights and platform snap
- ladder grab range
- hazard safe travel inset

Changing a sprite from a block to a more expressive shape often shifts where players expect its feet, claws, or wings to be.

## Practical art rules

- Keep player art centered unless the movement code is updated to account for an off-center silhouette.
- Keep ground monster feet visually near the bottom of the sprite box.
- Keep flying monster silhouettes balanced enough that horizontal flip still looks natural.
- Prefer a few high-contrast colors over muddy detail.
