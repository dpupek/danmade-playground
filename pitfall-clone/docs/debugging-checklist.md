# Debugging Checklist

Use this file when a gameplay regression appears after art, room, or movement changes.

## Player falls through floor

Check these first:

- `src/game/scenes/GameScene.ts`: `handleMovement`, `resolvePlatformLanding`, `findSupportingPlatform`
- `src/game/data/rooms.ts`: player spawn point and platform coordinates for the current room
- player render size vs. support assumptions in `GAME_CONFIG.player`

Questions to answer:

- Is the player feet position close enough to the platform top?
- Is the support inset too strict?
- Did a spawn point place Dan over empty space?

## Player hovers above platform after landing

Check:

- `resolvePlatformLanding` snap position
- player texture/display size
- platform top calculation

If the sprite changed recently, verify the art proportions still match the collision assumptions.

## Ladder is hard to enter or impossible to descend from platform

Check:

- `findActiveLadder` in `src/game/scenes/GameScene.ts`
- ladder top reach vs. rendered ladder height
- player x-centering while climbing
- ladder coordinates in `src/game/data/rooms.ts`

Bias toward a forgiving top-entry zone. Strict ladders feel broken fast.

## Monster walks off platform

Check:

- `resolveHazardTravelBounds`
- `findHazardPlatform`
- monster display width
- room-authored `minX` and `maxX`

Remember that authored lane bounds are only hints. The runtime safe span should keep the visible sprite supported.

## Sprite looks wrong but collisions seem wrong too

Check both:

- `src/game/scenes/BootScene.ts` for the texture shape
- `src/game/scenes/GameScene.ts` for display size, snap position, lane bounds, and overlap checks

In this project, visual tuning and gameplay tuning often need to move together.
