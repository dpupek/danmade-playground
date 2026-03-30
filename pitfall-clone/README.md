# Temple Runaway

Temple Runaway is a browser-based Pitfall-style action platformer built with Phaser and TypeScript. The first playable is a short mini-adventure with five connected rooms, treasure collection, moving hazards, ladders, pits, score, lives, and a countdown timer.

## Run locally

1. Install Node.js.
2. From `C:\Sandbox\danmade\danmade-playground\pitfall-clone`, run `npm install`.
3. Start the dev server with `npm run dev`.
4. Run tests with `npm test`.

## Controls

- Arrow keys or `WASD`: move and climb
- `Space`: jump
- `P`: pause
- `R`: restart the run

## Notes

- All art is generated in code with simple retro shapes and colors.
- Sound effects are oscillator-based and do not rely on external audio assets.
- Room content is data-driven through typed room definitions in `src/game/data/rooms.ts`.

## Architecture

- Player movement and platform support are handled by custom scene logic in `src/game/scenes/GameScene.ts`, not Phaser Arcade physics.
- Hazard lanes start from room-authored `minX` and `maxX`, then clamp to platform-derived safe travel bounds when a room loads.
- Generated textures for Dan, ladders, relics, and monsters live in `src/game/scenes/BootScene.ts`.
- Run state, scoring, treasure persistence, lives, and timer flow through `src/game/state/RunState.ts`.

## Docs

- `docs/gameplay-architecture.md`
- `docs/room-authoring.md`
- `docs/debugging-checklist.md`
- `docs/art-pipeline.md`
