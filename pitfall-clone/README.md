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
