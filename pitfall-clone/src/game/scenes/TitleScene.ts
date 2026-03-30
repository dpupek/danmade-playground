import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class TitleScene extends Phaser.Scene {
  private hasStarted = false;

  constructor() {
    super("title");
  }

  create(): void {
    const { width, height } = this.scale;

    const art = this.add.image(width / 2, height / 2, "title-box-art");
    const scale = Math.max(width / art.width, height / art.height);
    art.setScale(scale);

    const vignetteTop = this.add.rectangle(width / 2, 60, width, 130, 0x091018, 0.42);
    vignetteTop.setDepth(2);
    const vignetteBottom = this.add.rectangle(width / 2, height - 46, width, 118, 0x091018, 0.5);
    vignetteBottom.setDepth(2);

    this.drawOverlay(width, height);

    this.input.keyboard?.once("keydown", () => this.startGame());
    this.input.once("pointerdown", () => this.startGame());
  }

  private drawOverlay(width: number, height: number): void {
    this.add
      .text(width / 2, 58, GAME_CONFIG.title, {
        fontFamily: "Georgia",
        fontSize: "56px",
        color: "#f8efd1",
        stroke: "#27170f",
        strokeThickness: 10,
        shadow: { color: "#000000", blur: 0, fill: true, offsetX: 4, offsetY: 5 },
      })
      .setOrigin(0.5)
      .setDepth(10);

    this.add
      .text(width / 2, 106, "Dan the Hobbit And The Temple Menagerie", {
        fontFamily: "Georgia",
        fontSize: "24px",
        color: "#f4d35e",
        stroke: "#27170f",
        strokeThickness: 6,
      })
      .setOrigin(0.5)
      .setDepth(10);

    const panel = this.add.rectangle(width / 2, height - 62, 770, 98, 0x0b132b, 0.74);
    panel.setStrokeStyle(3, 0xf4d35e, 0.65);
    panel.setDepth(10);

    this.add
      .text(
        width / 2,
        height - 80,
        "Snatch the five golden clams, dodge Paul, Dave, and Mark,\nand escape the temple before time runs out.",
        {
          fontFamily: "Courier New",
          fontSize: "20px",
          align: "center",
          color: GAME_CONFIG.palette.text,
          wordWrap: { width: 700 },
        },
      )
      .setOrigin(0.5)
      .setDepth(11);

    const prompt = this.add
      .text(width / 2, height - 30, "Press Any Key Or Click To Begin", {
        fontFamily: "Courier New",
        fontSize: "24px",
        color: "#ffe082",
        stroke: "#0b132b",
        strokeThickness: 6,
      })
      .setOrigin(0.5)
      .setDepth(11)
      .setInteractive({ useHandCursor: true })
      .on("pointerdown", () => this.startGame());

    this.tweens.add({
      targets: prompt,
      alpha: { from: 0.68, to: 1 },
      scaleX: { from: 0.98, to: 1.03 },
      scaleY: { from: 0.98, to: 1.03 },
      duration: 780,
      yoyo: true,
      repeat: -1,
      ease: "Sine.easeInOut",
    });

    const glintA = this.add.image(94, 72, "glint-8").setScale(0.55).setDepth(11);
    const glintB = this.add.image(width - 116, 92, "glint-4").setScale(0.42).setDepth(11);
    [glintA, glintB].forEach((glint, index) => {
      this.tweens.add({
        targets: glint,
        alpha: { from: 0.35, to: 1 },
        scaleX: { from: glint.scaleX * 0.7, to: glint.scaleX * 1.15 },
        scaleY: { from: glint.scaleY * 0.7, to: glint.scaleY * 1.15 },
        angle: index === 0 ? 18 : -18,
        duration: 720 + index * 100,
        yoyo: true,
        repeat: -1,
        repeatDelay: 500 + index * 180,
        ease: "Sine.easeInOut",
      });
    });
  }

  private startGame(): void {
    if (this.hasStarted) {
      return;
    }

    this.hasStarted = true;
    this.scene.start("game");
  }
}
