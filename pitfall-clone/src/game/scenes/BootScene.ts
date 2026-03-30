import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class BootScene extends Phaser.Scene {
  constructor() {
    super("boot");
  }

  create(): void {
    this.createWoodTexture("wood-plank-a", 0);
    this.createWoodTexture("wood-plank-b", 1);
    this.createWoodTexture("wood-plank-c", 2);
    this.createGlintTexture("glint-4", 4);
    this.createGlintTexture("glint-8", 8);
    this.createLadderTexture();
    this.createPlayerTexture("player-stand", false);
    this.createPlayerTexture("player-jump", true);
    this.createTreasureTexture();
    this.createPaulTexture();
    this.createDaveTexture("dave-goat", false);
    this.createDaveTexture("dave-goat-scream", true);
    this.createMarkTexture();
    this.scene.start("title");
  }

  private createWoodTexture(key: string, variant: number): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const shadow = Phaser.Display.Color.HexStringToColor("#3b2418").color;
    const base = Phaser.Display.Color.HexStringToColor(variant === 0 ? "#8a5a37" : variant === 1 ? "#7a4d31" : "#94613b").color;
    const highlight = Phaser.Display.Color.HexStringToColor("#c38b56").color;
    const mid = Phaser.Display.Color.HexStringToColor("#a06d45").color;
    const dark = Phaser.Display.Color.HexStringToColor("#5e3b27").color;
    const knot = Phaser.Display.Color.HexStringToColor("#4a2c1d").color;

    graphics.fillStyle(base, 1);
    graphics.fillRect(0, 0, 32, 16);

    graphics.fillStyle(highlight, 1);
    graphics.fillRect(0, 0, 32, 2);
    graphics.fillRect(0, 2, 32, 1);

    graphics.fillStyle(mid, 1);
    for (let y = 4; y <= 10; y += 3) {
      graphics.fillRect(0, y, 32, 1);
    }

    graphics.fillStyle(dark, 1);
    graphics.fillRect(0, 12, 32, 2);
    graphics.fillRect(0, 14, 32, 2);

    const seams = variant === 0 ? [0, 10, 21, 31] : variant === 1 ? [0, 12, 20, 31] : [0, 8, 18, 31];
    graphics.fillStyle(shadow, 1);
    seams.forEach((x) => {
      graphics.fillRect(x, 0, 1, 16);
    });

    graphics.fillStyle(highlight, 0.8);
    seams.slice(1, -1).forEach((x) => {
      graphics.fillRect(x + 1, 1, 1, 10);
    });

    graphics.fillStyle(knot, 1);
    if (variant === 0) {
      graphics.fillRect(6, 7, 2, 2);
      graphics.fillRect(24, 5, 3, 2);
    } else if (variant === 1) {
      graphics.fillRect(15, 8, 3, 2);
      graphics.fillRect(26, 6, 2, 2);
    } else {
      graphics.fillRect(9, 5, 3, 2);
      graphics.fillRect(21, 8, 2, 2);
    }

    graphics.fillStyle(highlight, 0.45);
    const grainOffsets = variant === 0 ? [3, 13, 22] : variant === 1 ? [5, 16, 25] : [7, 14, 23];
    grainOffsets.forEach((x, index) => {
      graphics.fillRect(x, 4 + index, 5, 1);
      graphics.fillRect(x - 2, 9 + (index % 2), 6, 1);
    });

    graphics.generateTexture(key, 32, 16);
    graphics.destroy();
  }

  private createGlintTexture(key: string, points: 4 | 8): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const core = Phaser.Display.Color.HexStringToColor("#fff8c2").color;
    const halo = Phaser.Display.Color.HexStringToColor("#ffe066").color;

    graphics.fillStyle(halo, 0.55);
    graphics.fillRect(7, 1, 2, 14);
    graphics.fillRect(1, 7, 14, 2);

    if (points === 8) {
      graphics.fillRect(3, 3, 2, 2);
      graphics.fillRect(11, 3, 2, 2);
      graphics.fillRect(3, 11, 2, 2);
      graphics.fillRect(11, 11, 2, 2);
      graphics.fillRect(5, 5, 1, 1);
      graphics.fillRect(10, 5, 1, 1);
      graphics.fillRect(5, 10, 1, 1);
      graphics.fillRect(10, 10, 1, 1);
    }

    graphics.fillStyle(core, 1);
    graphics.fillRect(7, 4, 2, 8);
    graphics.fillRect(4, 7, 8, 2);
    graphics.fillRect(6, 6, 4, 4);

    if (points === 8) {
      graphics.fillRect(4, 4, 2, 2);
      graphics.fillRect(10, 4, 2, 2);
      graphics.fillRect(4, 10, 2, 2);
      graphics.fillRect(10, 10, 2, 2);
    }

    graphics.generateTexture(key, 16, 16);
    graphics.destroy();
  }

  private createPlayerTexture(key: string, jumping: boolean): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const hair = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.playerHair).color;
    const skin = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.playerSkin).color;
    const vest = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.playerVest).color;
    const shirt = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.playerShirt).color;
    const feet = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.playerFeet).color;
    const outline = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.shadow).color;

    graphics.fillStyle(outline, 1);
    graphics.fillRect(4, 1, 8, 4);
    graphics.fillRect(3, 5, 10, 4);
    graphics.fillRect(4, 9, 8, 5);
    if (jumping) {
      graphics.fillRect(2, 13, 3, 2);
      graphics.fillRect(11, 13, 3, 2);
    } else {
      graphics.fillRect(5, 14, 2, 2);
      graphics.fillRect(9, 14, 2, 2);
    }

    graphics.fillStyle(hair, 1);
    graphics.fillRect(5, 1, 6, 2);
    graphics.fillRect(4, 2, 8, 2);
    graphics.fillRect(3, 4, 2, 2);
    graphics.fillRect(11, 4, 2, 2);

    graphics.fillStyle(skin, 1);
    graphics.fillRect(5, 4, 6, 4);

    graphics.fillStyle(shirt, 1);
    graphics.fillRect(5, 9, 6, 2);
    graphics.fillRect(4, 10, 8, 2);

    graphics.fillStyle(vest, 1);
    graphics.fillRect(4, 11, 3, 3);
    graphics.fillRect(9, 11, 3, 3);
    graphics.fillRect(6, 12, 4, 2);

    graphics.fillStyle(feet, 1);
    if (jumping) {
      graphics.fillRect(2, 13, 3, 2);
      graphics.fillRect(11, 13, 3, 2);
    } else {
      graphics.fillRect(5, 14, 2, 2);
      graphics.fillRect(9, 14, 2, 2);
    }

    graphics.generateTexture(key, 16, 16);
    graphics.destroy();
  }

  private createLadderTexture(): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const rail = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.ladder).color;
    const rung = Phaser.Display.Color.HexStringToColor("#f6bd60").color;
    const shadow = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.shadow).color;

    graphics.fillStyle(shadow, 1);
    graphics.fillRect(1, 0, 3, 16);
    graphics.fillRect(12, 0, 3, 16);
    graphics.fillRect(3, 3, 10, 2);
    graphics.fillRect(3, 8, 10, 2);
    graphics.fillRect(3, 13, 10, 2);

    graphics.fillStyle(rail, 1);
    graphics.fillRect(2, 0, 2, 16);
    graphics.fillRect(12, 0, 2, 16);

    graphics.fillStyle(rung, 1);
    graphics.fillRect(3, 3, 10, 1);
    graphics.fillRect(3, 8, 10, 1);
    graphics.fillRect(3, 13, 10, 1);

    graphics.generateTexture("ladder", 16, 16);
    graphics.destroy();
  }

  private createTreasureTexture(): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const shell = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.treasure).color;
    const shellHighlight = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.treasureAccent).color;
    const eye = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.treasureEye).color;
    const shadow = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.shadow).color;

    graphics.fillStyle(shadow, 1);
    graphics.fillRect(3, 5, 10, 6);
    graphics.fillRect(4, 11, 8, 2);

    graphics.fillStyle(shell, 1);
    graphics.fillRect(4, 5, 8, 5);
    graphics.fillRect(5, 10, 6, 2);
    graphics.fillRect(6, 4, 4, 1);

    graphics.fillStyle(shellHighlight, 1);
    graphics.fillRect(5, 6, 6, 1);
    graphics.fillRect(6, 8, 4, 1);
    graphics.fillRect(5, 10, 1, 2);
    graphics.fillRect(10, 10, 1, 2);

    graphics.fillStyle(eye, 1);
    graphics.fillRect(6, 7, 1, 1);
    graphics.fillRect(9, 7, 1, 1);

    graphics.generateTexture("treasure", 16, 16);
    graphics.destroy();
  }

  private createPaulTexture(): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const shell = Phaser.Display.Color.HexStringToColor("#7a1f1f").color;
    const claw = Phaser.Display.Color.HexStringToColor("#b23a48").color;
    const eye = Phaser.Display.Color.HexStringToColor("#ffe066").color;
    const shadow = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.shadow).color;

    graphics.fillStyle(shadow, 1);
    graphics.fillRect(3, 5, 10, 7);
    graphics.fillRect(1, 3, 3, 4);
    graphics.fillRect(12, 3, 3, 4);
    graphics.fillRect(2, 12, 2, 3);
    graphics.fillRect(5, 12, 2, 3);
    graphics.fillRect(9, 12, 2, 3);
    graphics.fillRect(12, 12, 2, 3);

    graphics.fillStyle(shell, 1);
    graphics.fillRect(4, 5, 8, 6);
    graphics.fillRect(5, 4, 6, 1);
    graphics.fillRect(5, 11, 6, 1);

    graphics.fillStyle(claw, 1);
    graphics.fillRect(1, 3, 3, 3);
    graphics.fillRect(12, 3, 3, 3);
    graphics.fillRect(2, 2, 1, 1);
    graphics.fillRect(13, 2, 1, 1);

    graphics.fillStyle(eye, 1);
    graphics.fillRect(6, 3, 1, 2);
    graphics.fillRect(9, 3, 1, 2);

    graphics.generateTexture("paul-crab", 16, 16);
    graphics.destroy();
  }

  private createDaveTexture(key: string, screaming: boolean): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const fur = Phaser.Display.Color.HexStringToColor("#1f1f1f").color;
    const horn = Phaser.Display.Color.HexStringToColor("#bcb8b1").color;
    const eye = Phaser.Display.Color.HexStringToColor("#ef233c").color;
    const hoof = Phaser.Display.Color.HexStringToColor("#6c584c").color;
    const mouth = Phaser.Display.Color.HexStringToColor("#d62828").color;
    const shadow = Phaser.Display.Color.HexStringToColor(GAME_CONFIG.palette.shadow).color;

    graphics.fillStyle(shadow, 1);
    graphics.fillRect(2, 5, 11, 6);
    graphics.fillRect(10, 3, 4, 4);
    graphics.fillRect(3, 11, 2, 3);
    graphics.fillRect(10, 11, 2, 3);

    graphics.fillStyle(fur, 1);
    graphics.fillRect(3, 5, 9, 5);
    graphics.fillRect(10, 4, 3, 3);
    graphics.fillRect(8, 3, 2, 2);

    graphics.fillStyle(horn, 1);
    graphics.fillRect(10, 2, 1, 2);
    graphics.fillRect(12, 2, 1, 2);

    graphics.fillStyle(eye, 1);
    graphics.fillRect(11, 5, 1, 1);

    if (screaming) {
      graphics.fillStyle(mouth, 1);
      graphics.fillRect(12, 6, 2, 2);
      graphics.fillRect(11, 7, 1, 1);
    }

    graphics.fillStyle(hoof, 1);
    graphics.fillRect(3, 12, 2, 2);
    graphics.fillRect(10, 12, 2, 2);

    graphics.generateTexture(key, 16, 16);
    graphics.destroy();
  }

  private createMarkTexture(): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    const wing = Phaser.Display.Color.HexStringToColor("#6c757d").color;
    const body = Phaser.Display.Color.HexStringToColor("#2b2d42").color;
    const stripe = Phaser.Display.Color.HexStringToColor("#d4a373").color;
    const eye = Phaser.Display.Color.HexStringToColor("#ff006e").color;
    const stinger = Phaser.Display.Color.HexStringToColor("#8d99ae").color;

    graphics.fillStyle(wing, 0.9);
    graphics.fillRect(2, 3, 5, 3);
    graphics.fillRect(9, 3, 5, 3);
    graphics.fillRect(3, 1, 4, 2);
    graphics.fillRect(9, 1, 4, 2);

    graphics.fillStyle(body, 1);
    graphics.fillRect(6, 4, 4, 7);
    graphics.fillRect(5, 6, 6, 2);

    graphics.fillStyle(stripe, 1);
    graphics.fillRect(6, 6, 4, 1);
    graphics.fillRect(6, 8, 4, 1);

    graphics.fillStyle(eye, 1);
    graphics.fillRect(6, 4, 1, 1);
    graphics.fillRect(9, 4, 1, 1);

    graphics.fillStyle(stinger, 1);
    graphics.fillRect(7, 11, 2, 2);
    graphics.fillRect(7, 13, 1, 1);
    graphics.fillRect(8, 13, 1, 1);

    graphics.generateTexture("mark-wasp", 16, 16);
    graphics.destroy();
  }
}
