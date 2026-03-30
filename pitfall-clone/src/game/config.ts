export interface GameConfig {
  title: string;
  world: {
    width: number;
    height: number;
    floorY: number;
  };
  physics: {
    gravityY: number;
    runSpeed: number;
    climbSpeed: number;
    jumpVelocity: number;
    hazardSpeed: number;
  };
  run: {
    lives: number;
    timerSeconds: number;
    treasureScore: number;
  };
  player: {
    width: number;
    height: number;
  };
  palette: Record<string, string>;
}

export const GAME_CONFIG: GameConfig = {
  title: "Temple Runaway",
  world: {
    width: 960,
    height: 540,
    floorY: 468,
  },
  physics: {
    gravityY: 1100,
    runSpeed: 210,
    climbSpeed: 140,
    jumpVelocity: -460,
    hazardSpeed: 90,
  },
  run: {
    lives: 3,
    timerSeconds: 140,
    treasureScore: 250,
  },
  player: {
    width: 24,
    height: 40,
  },
  palette: {
    sky: "#15304c",
    canopy: "#2d6a4f",
    earth: "#5c3d2e",
    ladder: "#d9a441",
    player: "#f4d35e",
    playerAccent: "#ee964b",
    treasure: "#9ef01a",
    log: "#bc6c25",
    scorpion: "#ff006e",
    vine: "#55dde0",
    text: "#fefae0",
    shadow: "#0b132b",
  },
};
