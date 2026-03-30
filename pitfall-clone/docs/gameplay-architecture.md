# Gameplay Architecture

Temple Runaway is a small Phaser game, but the gameplay rules are intentionally split across data, state, and scene code.

## Core split

- `src/game/data/rooms.ts`: author rooms as data. Platforms, ladders, hazards, relics, exits, and spawn points live here.
- `src/game/state/RunState.ts`: own run-level rules. Score, lives, timer, collected relic ids, room transitions, win, and loss live here.
- `src/game/scenes/BootScene.ts`: generate textures and boot-time assets in code.
- `src/game/scenes/GameScene.ts`: render the current room and run the moment-to-moment movement, climbing, hazard motion, and overlap checks.

## Player movement

The player does not rely on Phaser Arcade bodies for floor contact. Dan uses a custom controller that:

- applies velocity and gravity manually
- resolves platform landing by finding a supporting platform near the player's feet
- uses a larger ladder grab zone than the rendered ladder art
- swaps generated textures based on grounded, climbing, and jump state

This is intentional. Earlier versions trusted Arcade collisions more directly and repeatedly regressed into falling through floors or hovering above platforms.

## Hazard movement

Hazards are data-driven, but their runtime behavior is scene-owned.

- `minX` and `maxX` in room data are lane hints
- the scene resolves a supporting platform for each hazard at room load
- the scene computes `safeMinX` and `safeMaxX` from platform geometry and monster width
- movement behaviors use the resolved safe bounds instead of the raw authored lane

This keeps room authoring simple while preventing monsters from visibly walking off ledges.

## Collision model

- Player-vs-platform: custom support checks
- Player-vs-ladder: rectangle overlap against an expanded ladder zone
- Player-vs-hazard: rectangle overlap against the monster sprite bounds
- Player-vs-relic: rectangle overlap against the relic sprite bounds

## Practical rule

When a visual change seems to affect gameplay, inspect both art and runtime geometry. In this project, render size, spawn position, ladder reach, and support bounds are tightly connected.
