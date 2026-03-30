import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class BootScene extends Phaser.Scene {
  constructor() {
    super("boot");
  }

  create(): void {
    this.createTexture("ground", 16, 16, GAME_CONFIG.palette.earth);
    this.createLadderTexture();
    this.createPlayerTexture("player-stand", false);
    this.createPlayerTexture("player-jump", true);
    this.createTexture("treasure", 16, 16, GAME_CONFIG.palette.treasure);
    this.createTexture("log", 16, 16, GAME_CONFIG.palette.log);
    this.createTexture("scorpion", 16, 16, GAME_CONFIG.palette.scorpion);
    this.scene.start("title");
  }

  private createTexture(key: string, width: number, height: number, color: string): void {
    const graphics = this.make.graphics({ x: 0, y: 0 }, false);
    graphics.fillStyle(Phaser.Display.Color.HexStringToColor(color).color, 1);
    graphics.fillRect(0, 0, width, height);
    graphics.generateTexture(key, width, height);
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
}
