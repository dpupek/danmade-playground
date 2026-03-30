import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class TitleScene extends Phaser.Scene {
  private hasStarted = false;

  constructor() {
    super("title");
  }

  create(): void {
    const { width, height } = this.scale;

    this.add.rectangle(width / 2, height / 2, width, height, 0x15304c);
    this.add.rectangle(width / 2, 120, width, 140, 0x2d6a4f, 0.8);

    this.add
      .text(width / 2, 92, GAME_CONFIG.title, {
        fontFamily: "Courier New",
        fontSize: "40px",
        color: GAME_CONFIG.palette.text,
      })
      .setOrigin(0.5);

    this.add
      .text(
        width / 2,
        220,
        "Dash through temple rooms, collect five relics,\nand escape before the jungle clock runs out.",
        {
          fontFamily: "Courier New",
          fontSize: "22px",
          align: "center",
          color: GAME_CONFIG.palette.text,
        },
      )
      .setOrigin(0.5);

    this.add
      .text(
        width / 2,
        330,
        "Controls\nArrow Keys or WASD: Move / Climb\nSpace: Jump\nP: Pause\nR: Restart Run",
        {
          fontFamily: "Courier New",
          fontSize: "20px",
          align: "center",
          color: GAME_CONFIG.palette.text,
        },
      )
      .setOrigin(0.5);

    this.add
      .text(width / 2, 455, "Press Enter or Space to Begin", {
        fontFamily: "Courier New",
        fontSize: "24px",
        color: "#f4d35e",
      })
      .setOrigin(0.5)
      .setInteractive({ useHandCursor: true })
      .on("pointerdown", () => this.startGame());

    this.input.keyboard?.once("keydown", () => this.startGame());
    this.input.once("pointerdown", () => this.startGame());
  }

  private startGame(): void {
    if (this.hasStarted) {
      return;
    }

    this.hasStarted = true;
    this.scene.start("game");
  }
}
