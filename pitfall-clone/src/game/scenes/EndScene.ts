import Phaser from "phaser";
import { GAME_CONFIG } from "../config";
import type { EndScenePayload } from "../types";

export class EndScene extends Phaser.Scene {
  constructor() {
    super("end");
  }

  create(data: EndScenePayload): void {
    const { width, height } = this.scale;
    const didWin = data.outcome === "won";

    this.add.rectangle(width / 2, height / 2, width, height, 0x0b132b);
    this.add
      .text(width / 2, 140, didWin ? "Temple Escaped" : "Run Lost", {
        fontFamily: "Courier New",
        fontSize: "42px",
        color: didWin ? GAME_CONFIG.palette.treasure : GAME_CONFIG.palette.scorpion,
      })
      .setOrigin(0.5);

    this.add
      .text(
        width / 2,
        250,
        `Score ${data.runState.score}\nRelics ${data.runState.collectedTreasureIds.length}/5\nLives ${data.runState.lives}`,
        {
          fontFamily: "Courier New",
          fontSize: "24px",
          align: "center",
          color: GAME_CONFIG.palette.text,
        },
      )
      .setOrigin(0.5);

    this.add
      .text(
        width / 2,
        390,
        didWin
          ? "You outran the ruins and carried every relic home."
          : "The temple keeps its prizes for now.",
        {
          fontFamily: "Courier New",
          fontSize: "22px",
          align: "center",
          color: GAME_CONFIG.palette.text,
        },
      )
      .setOrigin(0.5);

    this.add
      .text(width / 2, 470, "Press R to retry or T for title", {
        fontFamily: "Courier New",
        fontSize: "22px",
        color: "#f4d35e",
      })
      .setOrigin(0.5);

    this.input.keyboard?.on("keydown-R", () => this.scene.start("game"));
    this.input.keyboard?.on("keydown-T", () => this.scene.start("title"));
  }
}
