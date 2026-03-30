import Phaser from "phaser";
import { GAME_CONFIG } from "../config";

export class BootScene extends Phaser.Scene {
  constructor() {
    super("boot");
  }

  create(): void {
    this.createTexture("ground", 16, 16, GAME_CONFIG.palette.earth);
    this.createTexture("ladder", 16, 16, GAME_CONFIG.palette.ladder);
    this.createTexture("player", 16, 16, GAME_CONFIG.palette.player);
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
}
