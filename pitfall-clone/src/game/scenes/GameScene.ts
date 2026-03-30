import Phaser from "phaser";
import { RetroSfx } from "../audio/RetroSfx";
import { GAME_CONFIG } from "../config";
import { getRoomDefinition } from "../data/rooms";
import { RunStateController } from "../state/RunState";
import type {
  HazardDefinition,
  LadderDefinition,
  PlatformDefinition,
  RoomDefinition,
  RoomId,
  RunState,
  ScenePayload,
  TreasureDefinition,
} from "../types";

type PlayerSprite = Phaser.GameObjects.Image;

type HazardView = {
  config: HazardDefinition;
  sprite: Phaser.GameObjects.Image;
  direction: 1 | -1;
};

type TreasureView = {
  config: TreasureDefinition;
  sprite: Phaser.GameObjects.Image;
};

export class GameScene extends Phaser.Scene {
  private controller = new RunStateController();
  private runState!: RunState;
  private player!: PlayerSprite;
  private cursors!: Phaser.Types.Input.Keyboard.CursorKeys;
  private keys!: Record<"left" | "right" | "up" | "down" | "jump" | "pause" | "restart", Phaser.Input.Keyboard.Key>;
  private roomLabel!: Phaser.GameObjects.Text;
  private hudLabel!: Phaser.GameObjects.Text;
  private statusLabel!: Phaser.GameObjects.Text;
  private sfx!: RetroSfx;
  private isPaused = false;
  private isClimbing = false;
  private isRespawning = false;
  private currentRoom!: RoomDefinition;
  private platformViews: Phaser.GameObjects.Rectangle[] = [];
  private ladderViews: Phaser.GameObjects.Graphics[] = [];
  private hazardViews: HazardView[] = [];
  private treasureViews: TreasureView[] = [];
  private vx = 0;
  private vy = 0;
  private isGrounded = false;
  private activeLadder: LadderDefinition | null = null;
  private facing: 1 | -1 = 1;

  constructor() {
    super("game");
  }

  create(data: ScenePayload): void {
    this.controller = new RunStateController(data.runState);
    this.runState = this.controller.snapshot;
    this.sfx = new RetroSfx(this);

    this.drawBackdrop();
    this.createUi();
    this.createInput();

    this.player = this.add.image(0, 0, "player-stand");
    this.player.setDisplaySize(GAME_CONFIG.player.width, GAME_CONFIG.player.height);
    this.player.setDepth(5);

    this.loadRoom(this.runState.currentRoomId, "default");
  }

  update(_time: number, delta: number): void {
    if (this.isPaused) {
      this.statusLabel.setText("Paused");
      return;
    }

    const dt = delta / 1000;
    this.activeLadder = this.findActiveLadder();

    this.handleMovement(dt);
    this.moveHazards(dt);
    this.collectTreasureOverlaps();
    this.checkHazardOverlaps();
    this.checkRoomTransitions();
    this.checkFallDeath();

    this.runState = this.controller.tick(delta);
    this.refreshHud();

    if (this.runState.status === "lost") {
      this.scene.start("end", { outcome: "lost", runState: this.runState });
    }
  }

  private drawBackdrop(): void {
    const { width, height } = this.scale;
    this.add.rectangle(width / 2, height / 2, width, height, 0x15304c);

    for (let i = 0; i < 6; i += 1) {
      this.add
        .rectangle(120 + i * 160, 90 + (i % 2) * 16, 180, 90, 0x2d6a4f, 0.65)
        .setAngle(i % 2 === 0 ? -5 : 5);
    }
  }

  private createUi(): void {
    this.roomLabel = this.add.text(24, 20, "", {
      fontFamily: "Courier New",
      fontSize: "18px",
      color: GAME_CONFIG.palette.text,
    });

    this.hudLabel = this.add.text(24, 48, "", {
      fontFamily: "Courier New",
      fontSize: "20px",
      color: "#f4d35e",
    });

    this.statusLabel = this.add.text(GAME_CONFIG.world.width - 24, 20, "", {
      fontFamily: "Courier New",
      fontSize: "18px",
      color: GAME_CONFIG.palette.text,
    });
    this.statusLabel.setOrigin(1, 0);
  }

  private createInput(): void {
    this.cursors = this.input.keyboard?.createCursorKeys() as Phaser.Types.Input.Keyboard.CursorKeys;
    this.keys = {
      left: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.A),
      right: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.D),
      up: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.W),
      down: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.S),
      jump: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.SPACE),
      pause: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.P),
      restart: this.input.keyboard!.addKey(Phaser.Input.Keyboard.KeyCodes.R),
    };

    this.keys.pause.on("down", () => {
      this.isPaused = !this.isPaused;
      this.refreshHud();
    });

    this.keys.restart.on("down", () => {
      this.scene.restart();
    });
  }

  private handleMovement(dt: number): void {
    const moveLeft = this.cursors.left.isDown || this.keys.left.isDown;
    const moveRight = this.cursors.right.isDown || this.keys.right.isDown;
    const moveUp = this.cursors.up.isDown || this.keys.up.isDown;
    const moveDown = this.cursors.down.isDown || this.keys.down.isDown;
    const wantsJump =
      Phaser.Input.Keyboard.JustDown(this.cursors.space) || Phaser.Input.Keyboard.JustDown(this.keys.jump);

    if (this.activeLadder && (moveUp || moveDown || this.isClimbing)) {
      this.isClimbing = true;
      this.isGrounded = false;
      this.vx = 0;
      this.vy = 0;
      this.player.x = Phaser.Math.Linear(this.player.x, this.activeLadder.x, 0.7);

      if (moveUp) {
        this.player.y -= GAME_CONFIG.physics.climbSpeed * dt;
      } else if (moveDown) {
        this.player.y += GAME_CONFIG.physics.climbSpeed * dt;
      }
    } else {
      if (this.isClimbing) {
        this.isClimbing = false;
      }

      if (moveLeft) {
        this.vx = -GAME_CONFIG.physics.runSpeed;
        this.facing = -1;
      } else if (moveRight) {
        this.vx = GAME_CONFIG.physics.runSpeed;
        this.facing = 1;
      } else {
        this.vx = 0;
      }

      if (wantsJump && this.isGrounded) {
        this.vy = GAME_CONFIG.physics.jumpVelocity;
        this.isGrounded = false;
        this.sfx.jump();
      }

      this.player.x += this.vx * dt;
      this.vy += GAME_CONFIG.physics.gravityY * dt;
      this.player.y += this.vy * dt;

      this.resolvePlatformLanding();
    }

    if (wantsJump && this.isClimbing) {
      this.isClimbing = false;
      this.vy = GAME_CONFIG.physics.jumpVelocity * 0.8;
      this.isGrounded = false;
      this.sfx.jump();
    }

    this.updatePlayerPose();
  }

  private resolvePlatformLanding(): void {
    const support = this.findSupportingPlatform();
    if (!support || this.vy < 0) {
      this.isGrounded = false;
      return;
    }

    const platformTop = support.y - support.height / 2;
    this.player.y = platformTop - GAME_CONFIG.player.height / 2;
    this.vy = 0;
    this.isGrounded = true;
  }

  private findSupportingPlatform(): PlatformDefinition | null {
    const feetX = this.player.x;
    const feetY = this.player.y + GAME_CONFIG.player.height / 2;

    for (const platform of this.currentRoom.platforms) {
      const left = platform.x - platform.width / 2;
      const right = platform.x + platform.width / 2;
      const top = platform.y - platform.height / 2;
      const withinX = feetX >= left + 6 && feetX <= right - 6;
      const closeToTop = feetY >= top - 4 && feetY <= top + 14;

      if (withinX && closeToTop) {
        return platform;
      }
    }

    return null;
  }

  private findActiveLadder(): LadderDefinition | null {
    const playerBounds = this.getPlayerBounds();

    for (const ladder of this.currentRoom.ladders) {
      const ladderRect = new Phaser.Geom.Rectangle(
        ladder.x - ladder.width / 2 - 8,
        ladder.y - ladder.height / 2 - 20,
        ladder.width + 16,
        ladder.height + 32,
      );

      if (Phaser.Geom.Intersects.RectangleToRectangle(playerBounds, ladderRect)) {
        return ladder;
      }
    }

    return null;
  }

  private moveHazards(dt: number): void {
    this.hazardViews.forEach((hazard) => {
      const speed = hazard.config.speed ?? GAME_CONFIG.physics.hazardSpeed;
      hazard.sprite.x += speed * hazard.direction * dt;

      if (hazard.sprite.x >= hazard.config.maxX) {
        hazard.sprite.x = hazard.config.maxX;
        hazard.direction = -1;
        hazard.sprite.setFlipX(true);
      } else if (hazard.sprite.x <= hazard.config.minX) {
        hazard.sprite.x = hazard.config.minX;
        hazard.direction = 1;
        hazard.sprite.setFlipX(false);
      }
    });
  }

  private collectTreasureOverlaps(): void {
    const playerBounds = this.getPlayerBounds();

    for (const treasure of [...this.treasureViews]) {
      if (!Phaser.Geom.Intersects.RectangleToRectangle(playerBounds, treasure.sprite.getBounds())) {
        continue;
      }

      this.runState = this.controller.collectTreasure(treasure.config.id);
      treasure.sprite.destroy();
      this.treasureViews = this.treasureViews.filter((entry) => entry !== treasure);
      this.sfx.pickup();
      this.refreshHud();

      if (this.runState.status === "won") {
        this.sfx.win();
        this.scene.start("end", { outcome: "won", runState: this.runState });
      }
    }
  }

  private checkHazardOverlaps(): void {
    if (this.isRespawning || this.runState.status !== "playing") {
      return;
    }

    const playerBounds = this.getPlayerBounds();
    for (const hazard of this.hazardViews) {
      if (Phaser.Geom.Intersects.RectangleToRectangle(playerBounds, hazard.sprite.getBounds())) {
        void this.handleDeath();
        return;
      }
    }
  }

  private loadRoom(roomId: RoomId, spawnKey: "default" | "fromLeft" | "fromRight"): void {
    this.controller.moveToRoom(roomId);
    this.runState = this.controller.snapshot;
    this.currentRoom = getRoomDefinition(roomId);

    this.platformViews.forEach((view) => view.destroy());
    this.ladderViews.forEach((view) => view.destroy());
    this.hazardViews.forEach((view) => view.sprite.destroy());
    this.treasureViews.forEach((view) => view.sprite.destroy());
    this.platformViews = [];
    this.ladderViews = [];
    this.hazardViews = [];
    this.treasureViews = [];

    this.currentRoom.platforms.forEach((platform) => {
      const block = this.add.rectangle(platform.x, platform.y, platform.width, platform.height, 0x5c3d2e);
      this.platformViews.push(block);
    });

    this.currentRoom.ladders.forEach((ladder) => {
      const view = this.add.graphics();
      const left = ladder.x - ladder.width / 2;
      const top = ladder.y - ladder.height / 2 - 10;
      const height = ladder.height + 10;
      const railWidth = 4;
      const rungInset = 4;
      const rungHeight = 2;
      const rungSpacing = 10;

      view.fillStyle(0x8d5524, 1);
      view.fillRect(left, top, railWidth, height);
      view.fillRect(left + ladder.width - railWidth, top, railWidth, height);

      view.fillStyle(0xf6bd60, 1);
      for (let y = top + 6; y < top + height - 2; y += rungSpacing) {
        view.fillRect(left + rungInset, y, ladder.width - rungInset * 2, rungHeight);
      }

      view.lineStyle(1, 0x432818, 1);
      view.strokeRect(left, top, railWidth, height);
      view.strokeRect(left + ladder.width - railWidth, top, railWidth, height);
      this.ladderViews.push(view);
    });

    this.currentRoom.hazards.forEach((hazardConfig) => {
      const texture = hazardConfig.type === "log" ? "log" : "scorpion";
      const sprite = this.add.image(hazardConfig.x, hazardConfig.y, texture);
      sprite.setDisplaySize(hazardConfig.width, hazardConfig.height);
      this.hazardViews.push({
        config: hazardConfig,
        sprite,
        direction: 1,
      });
    });

    this.currentRoom.treasures
      .filter((treasure) => !this.controller.hasCollected(treasure.id))
      .forEach((treasureConfig) => {
        const sprite = this.add.image(treasureConfig.x, treasureConfig.y, "treasure");
        sprite.setDisplaySize(20, 20);
        this.treasureViews.push({ config: treasureConfig, sprite });
      });

    const spawn = this.currentRoom.spawn[spawnKey] ?? this.currentRoom.spawn.default;
    this.player.setPosition(spawn.x, spawn.y);
    this.player.setTint(0xffffff);
    this.vx = 0;
    this.vy = 0;
    this.isClimbing = false;
    this.isGrounded = false;
    this.activeLadder = null;
    this.facing = 1;
    this.updatePlayerPose();
    this.roomLabel.setText(this.currentRoom.title);
    this.refreshHud();
  }

  private checkRoomTransitions(): void {
    if (this.player.x < -20 && this.currentRoom.exits.left) {
      this.loadRoom(this.currentRoom.exits.left, "fromRight");
    } else if (this.player.x > GAME_CONFIG.world.width + 20 && this.currentRoom.exits.right) {
      this.loadRoom(this.currentRoom.exits.right, "fromLeft");
    }
  }

  private checkFallDeath(): void {
    if (this.player.y > GAME_CONFIG.world.height + 32) {
      void this.handleDeath();
    }
  }

  private async handleDeath(): Promise<void> {
    if (this.isRespawning || this.runState.status !== "playing") {
      return;
    }

    this.isRespawning = true;
    this.sfx.hurt();
    this.player.setTint(0xff0000);
    this.vx = 0;
    this.vy = 0;
    this.runState = this.controller.loseLife();
    this.refreshHud();

    await new Promise<void>((resolve) => {
      this.time.delayedCall(500, () => resolve());
    });

    if (this.runState.status === "lost") {
      this.scene.start("end", { outcome: "lost", runState: this.runState });
      return;
    }

    this.loadRoom(this.runState.currentRoomId, "default");
    this.isRespawning = false;
  }

  private refreshHud(): void {
    const seconds = Math.ceil(this.runState.timeRemainingMs / 1000);
    this.hudLabel.setText(
      `Score ${this.runState.score}   Lives ${this.runState.lives}   Relics ${this.runState.collectedTreasureIds.length}/5   Time ${seconds}`,
    );
    this.statusLabel.setText(this.isPaused ? "Paused" : "P pause  R restart");
  }

  private updatePlayerPose(): void {
    const texture = !this.isGrounded && !this.isClimbing ? "player-jump" : "player-stand";
    this.player.setTexture(texture);
    this.player.setFlipX(this.facing < 0);
  }

  private getPlayerBounds(): Phaser.Geom.Rectangle {
    return new Phaser.Geom.Rectangle(
      this.player.x - GAME_CONFIG.player.width / 2,
      this.player.y - GAME_CONFIG.player.height / 2,
      GAME_CONFIG.player.width,
      GAME_CONFIG.player.height,
    );
  }
}
