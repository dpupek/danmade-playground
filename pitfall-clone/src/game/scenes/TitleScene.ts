import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class TitleScene extends Phaser.Scene {
  private hasStarted = false;

  constructor() {
    super("title");
  }

  create(): void {
    const { width, height } = this.scale;

    this.drawBackdrop(width, height);
    this.drawStage(width, height);
    this.drawCast(width, height);
    this.drawUi(width, height);

    this.input.keyboard?.once("keydown", () => this.startGame());
    this.input.once("pointerdown", () => this.startGame());
  }

  private drawBackdrop(width: number, height: number): void {
    this.add.rectangle(width / 2, height / 2, width, height, 0x10263d);
    this.add.rectangle(width / 2, 92, width, 168, 0x214e44, 0.75);
    this.add.rectangle(width / 2, 165, width + 80, 130, 0x2f6f5b, 0.42);
    this.add.rectangle(width / 2, 240, width + 120, 150, 0x173a2e, 0.55);

    for (let i = 0; i < 7; i += 1) {
      this.add
        .ellipse(80 + i * 145, 78 + (i % 2) * 10, 160, 66, 0x3a7a61, 0.5)
        .setAngle(i % 2 === 0 ? -8 : 7);
    }

    const moon = this.add.circle(width - 126, 86, 34, 0xfef3c7, 0.9);
    moon.setStrokeStyle(4, 0xffe082, 0.35);

    for (let i = 0; i < 10; i += 1) {
      const glintKey = i % 3 === 0 ? "glint-8" : "glint-4";
      const star = this.add.image(84 + i * 84, 48 + (i % 4) * 22, glintKey);
      const baseScale = glintKey === "glint-8" ? 0.45 : 0.26;
      star.setScale(baseScale);
      star.setAlpha(glintKey === "glint-8" ? 0.55 : 0.35);
      this.tweens.add({
        targets: star,
        alpha: { from: star.alpha, to: 1 },
        scaleX: { from: baseScale * 0.7, to: baseScale * 1.25 },
        scaleY: { from: baseScale * 0.7, to: baseScale * 1.25 },
        angle: star.angle + 15,
        duration: 900 + i * 40,
        delay: i * 120,
        yoyo: true,
        repeat: -1,
        ease: "Sine.easeInOut",
      });
    }
  }

  private drawStage(width: number, height: number): void {
    const floor = this.add.container(0, 0);
    const woodBase = this.add.tileSprite(width / 2, height - 82, width + 40, 132, "wood-plank-b");
    woodBase.tilePositionX = 9;
    const woodOverlay = this.add.tileSprite(width / 2, height - 110, width + 40, 76, "wood-plank-c");
    woodOverlay.tilePositionX = 3;
    woodOverlay.setAlpha(0.92);
    const topLip = this.add.rectangle(width / 2, height - 146, width + 40, 8, 0xd4a373, 0.65);
    const shadow = this.add.rectangle(width / 2, height - 20, width + 40, 24, 0x25160f, 0.85);

    floor.add([woodBase, woodOverlay, topLip, shadow]);

    const leftPlank = this.add.tileSprite(154, height - 214, 220, 26, "wood-plank-a");
    leftPlank.tilePositionX = 6;
    leftPlank.setAngle(-6);
    const rightPlank = this.add.tileSprite(width - 176, height - 228, 240, 26, "wood-plank-a");
    rightPlank.tilePositionX = 13;
    rightPlank.setAngle(5);
    const centerPlank = this.add.tileSprite(width / 2, height - 264, 196, 24, "wood-plank-c");
    centerPlank.tilePositionX = 17;

    this.add.rectangle(154, height - 199, 220, 4, 0xe0b07c, 0.55).setAngle(-6);
    this.add.rectangle(width - 176, height - 213, 240, 4, 0xe0b07c, 0.55).setAngle(5);
    this.add.rectangle(width / 2, height - 276, 196, 4, 0xe0b07c, 0.55);
  }

  private drawCast(width: number, height: number): void {
    const dan = this.add.image(width / 2 - 20, height - 214, "player-stand");
    dan.setDisplaySize(170, 280);
    dan.setFlipX(false);

    const paul = this.add.image(148, height - 132, "paul-crab");
    paul.setDisplaySize(170, 128);
    paul.setFlipX(false);

    const dave = this.add.image(width - 178, height - 168, "dave-goat-scream");
    dave.setDisplaySize(190, 138);
    dave.setFlipX(true);

    const mark = this.add.image(width / 2 + 190, 182, "mark-wasp");
    mark.setDisplaySize(150, 94);
    mark.setAngle(-12);

    this.tweens.add({
      targets: mark,
      y: mark.y - 16,
      angle: 10,
      duration: 1400,
      yoyo: true,
      repeat: -1,
      ease: "Sine.easeInOut",
    });

    this.createRelicCluster(width, height);
  }

  private createRelicCluster(width: number, height: number): void {
    const positions = [
      { x: width / 2 - 186, y: height - 198, scale: 0.88 },
      { x: width / 2 + 106, y: height - 238, scale: 1.04 },
      { x: width / 2 + 270, y: height - 172, scale: 0.82 },
    ];

    positions.forEach((entry, index) => {
      const clam = this.add.image(entry.x, entry.y, "treasure");
      clam.setDisplaySize(58 * entry.scale, 58 * entry.scale);
      clam.setAngle(index === 1 ? -8 : index === 2 ? 10 : -12);
      clam.setDepth(7);

      const mainGlint = this.add.image(entry.x + 28, entry.y - 24, "glint-8");
      mainGlint.setScale(0.88 * entry.scale);
      mainGlint.setDepth(8);
      const sideGlint = this.add.image(entry.x - 24, entry.y + 6, "glint-4");
      sideGlint.setScale(0.56 * entry.scale);
      sideGlint.setDepth(8);

      [mainGlint, sideGlint].forEach((glint, glintIndex) => {
        this.tweens.add({
          targets: glint,
          alpha: { from: 0.25, to: 1 },
          scaleX: { from: glint.scaleX * 0.45, to: glint.scaleX * 1.18 },
          scaleY: { from: glint.scaleY * 0.45, to: glint.scaleY * 1.18 },
          angle: glint.angle + (glintIndex === 0 ? 18 : -14),
          duration: 520 + glintIndex * 120,
          yoyo: true,
          repeat: -1,
          repeatDelay: 700 + index * 110 + glintIndex * 80,
          delay: 160 + index * 180 + glintIndex * 130,
          ease: "Sine.easeOut",
        });
      });
    });
  }

  private drawUi(width: number, height: number): void {
    this.add
      .text(width / 2, 82, GAME_CONFIG.title, {
        fontFamily: "Georgia",
        fontSize: "54px",
        color: GAME_CONFIG.palette.text,
        stroke: "#0b132b",
        strokeThickness: 8,
      })
      .setOrigin(0.5)
      .setDepth(20);

    this.add
      .text(width / 2, 132, "Dan the Hobbit Vs. The Temple Menagerie", {
        fontFamily: "Georgia",
        fontSize: "24px",
        color: "#f4d35e",
        stroke: "#0b132b",
        strokeThickness: 6,
      })
      .setOrigin(0.5)
      .setDepth(20);

    const panel = this.add.rectangle(width / 2, height - 70, 730, 92, 0x0b132b, 0.62);
    panel.setStrokeStyle(3, 0xf4d35e, 0.55);
    panel.setDepth(19);

    this.add
      .text(width / 2, height - 86, "Snatch the five golden clams, dodge Paul, Dave, and Mark, and escape the temple before time runs out.", {
        fontFamily: "Courier New",
        fontSize: "20px",
        color: GAME_CONFIG.palette.text,
        align: "center",
        wordWrap: { width: 660 },
      })
      .setOrigin(0.5)
      .setDepth(20);

    const prompt = this.add
      .text(width / 2, height - 36, "Press Any Key Or Click To Begin", {
        fontFamily: "Courier New",
        fontSize: "24px",
        color: "#ffe082",
        stroke: "#0b132b",
        strokeThickness: 6,
      })
      .setOrigin(0.5)
      .setDepth(20)
      .setInteractive({ useHandCursor: true })
      .on("pointerdown", () => this.startGame());

    this.tweens.add({
      targets: prompt,
      alpha: { from: 0.7, to: 1 },
      scaleX: { from: 0.98, to: 1.03 },
      scaleY: { from: 0.98, to: 1.03 },
      duration: 760,
      yoyo: true,
      repeat: -1,
      ease: "Sine.easeInOut",
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
