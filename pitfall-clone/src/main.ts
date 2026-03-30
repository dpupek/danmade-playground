import Phaser from "phaser";
import "./styles.css";
import { GAME_CONFIG } from "./game/config";
import { BootScene } from "./game/scenes/BootScene";
import { EndScene } from "./game/scenes/EndScene";
import { GameScene } from "./game/scenes/GameScene";
import { TitleScene } from "./game/scenes/TitleScene";

const config: Phaser.Types.Core.GameConfig = {
  type: Phaser.AUTO,
  width: GAME_CONFIG.world.width,
  height: GAME_CONFIG.world.height,
  parent: "app",
  backgroundColor: GAME_CONFIG.palette.sky,
  pixelArt: true,
  physics: {
    default: "arcade",
    arcade: {
      gravity: { x: 0, y: GAME_CONFIG.physics.gravityY },
      debug: false,
    },
  },
  scale: {
    mode: Phaser.Scale.FIT,
    autoCenter: Phaser.Scale.CENTER_BOTH,
  },
  scene: [BootScene, TitleScene, GameScene, EndScene],
};

new Phaser.Game(config);
