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
  originY: number;
  modeTimer: number;
  safeMinX: number;
  safeMaxX: number;
  isCharging: boolean;
};

type TreasureView = {
  config: TreasureDefinition;
  sprite: Phaser.GameObjects.Image;
  sparkles: Phaser.GameObjects.Image[];
};

type HorizontalSpan = {
  minX: number;
  maxX: number;
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
  private platformViews: Phaser.GameObjects.Container[] = [];
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
      const span = this.getPlatformSpan(platform, 6);
      const top = this.getPlatformTop(platform);
      const withinX = feetX >= span.minX && feetX <= span.maxX;
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
      switch (hazard.config.type) {
        case "paul_crab":
          this.movePaul(hazard, dt);
          break;
        case "dave_goat":
          this.moveDave(hazard, dt);
          break;
        case "mark_wasp":
          this.moveMark(hazard, dt);
          break;
      }
    });
  }

  private movePaul(hazard: HazardView, dt: number): void {
    const speed = hazard.config.speed ?? 120;
    hazard.sprite.x += speed * hazard.direction * dt;

    if (hazard.sprite.x >= hazard.safeMaxX) {
      hazard.sprite.x = hazard.safeMaxX;
      hazard.direction = -1;
    } else if (hazard.sprite.x <= hazard.safeMinX) {
      hazard.sprite.x = hazard.safeMinX;
      hazard.direction = 1;
    }

    hazard.sprite.setFlipX(hazard.direction < 0);
  }

  private moveDave(hazard: HazardView, dt: number): void {
    const pauseMs = hazard.config.chargePauseMs ?? 1000;
    hazard.modeTimer += dt * 1000;

    if (hazard.modeTimer < pauseMs) {
      hazard.isCharging = false;
      hazard.sprite.setTexture("dave-goat");
      return;
    }

    if (!hazard.isCharging) {
      hazard.isCharging = true;
    }

    hazard.sprite.setTexture("dave-goat-scream");
    const speed = (hazard.config.speed ?? 90) * 1.65;
    hazard.sprite.x += speed * hazard.direction * dt;

    if (hazard.sprite.x >= hazard.safeMaxX) {
      hazard.sprite.x = hazard.safeMaxX;
      hazard.direction = -1;
      hazard.modeTimer = 0;
      hazard.isCharging = false;
      hazard.sprite.setTexture("dave-goat");
    } else if (hazard.sprite.x <= hazard.safeMinX) {
      hazard.sprite.x = hazard.safeMinX;
      hazard.direction = 1;
      hazard.modeTimer = 0;
      hazard.isCharging = false;
      hazard.sprite.setTexture("dave-goat");
    }

    hazard.sprite.setFlipX(hazard.direction < 0);
  }

  private moveMark(hazard: HazardView, dt: number): void {
    const speed = hazard.config.speed ?? 90;
    hazard.modeTimer += dt;
    hazard.sprite.x += speed * hazard.direction * dt;

    if (hazard.sprite.x >= hazard.safeMaxX) {
      hazard.sprite.x = hazard.safeMaxX;
      hazard.direction = -1;
    } else if (hazard.sprite.x <= hazard.safeMinX) {
      hazard.sprite.x = hazard.safeMinX;
      hazard.direction = 1;
    }

    const swoopDepth = hazard.config.swoopDepth ?? 22;
    const swoopRate = hazard.config.swoopRate ?? 2.5;
    hazard.sprite.y = hazard.originY + Math.sin(hazard.modeTimer * swoopRate * Math.PI) * swoopDepth;
    hazard.sprite.setFlipX(hazard.direction < 0);
  }

  private collectTreasureOverlaps(): void {
    const playerBounds = this.getPlayerBounds();

    for (const treasure of [...this.treasureViews]) {
      if (!Phaser.Geom.Intersects.RectangleToRectangle(playerBounds, treasure.sprite.getBounds())) {
        continue;
      }

      this.runState = this.controller.collectTreasure(treasure.config.id);
      treasure.sprite.destroy();
      treasure.sparkles.forEach((sparkle) => {
        this.tweens.killTweensOf(sparkle);
        sparkle.destroy();
      });
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
        void this.handleDeath(hazard.config.type);
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
    this.treasureViews.forEach((view) => {
      view.sprite.destroy();
      view.sparkles.forEach((sparkle) => {
        this.tweens.killTweensOf(sparkle);
        sparkle.destroy();
      });
    });
    this.platformViews = [];
    this.ladderViews = [];
    this.hazardViews = [];
    this.treasureViews = [];

    this.currentRoom.platforms.forEach((platform) => {
      this.platformViews.push(this.createPlatformView(platform));
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
      const safeBounds = this.resolveHazardTravelBounds(hazardConfig);
      const texture =
        hazardConfig.type === "paul_crab"
          ? "paul-crab"
          : hazardConfig.type === "dave_goat"
            ? "dave-goat"
            : "mark-wasp";
      const sprite = this.add.image(hazardConfig.x, hazardConfig.y, texture);
      sprite.setDisplaySize(hazardConfig.width, hazardConfig.height);
      this.hazardViews.push({
        config: hazardConfig,
        sprite,
        direction: 1,
        originY: hazardConfig.y,
        modeTimer: hazardConfig.type === "dave_goat" ? Math.random() * (hazardConfig.chargePauseMs ?? 1000) : 0,
        safeMinX: safeBounds.minX,
        safeMaxX: safeBounds.maxX,
        isCharging: false,
      });
    });

    this.currentRoom.treasures
      .filter((treasure) => !this.controller.hasCollected(treasure.id))
      .forEach((treasureConfig) => {
        const sprite = this.add.image(treasureConfig.x, treasureConfig.y, "treasure");
        sprite.setDisplaySize(40, 40);
        sprite.setDepth(4);
        sprite.setAlpha(0.98);
        const sparkles = this.createTreasureSparkles(sprite);
        this.treasureViews.push({ config: treasureConfig, sprite, sparkles });
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

  private async handleDeath(cause?: HazardDefinition["type"]): Promise<void> {
    if (this.isRespawning || this.runState.status !== "playing") {
      return;
    }

    this.isRespawning = true;
    if (cause === "dave_goat") {
      this.sfx.humanScream();
    } else {
      this.sfx.hurt();
    }
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
      `Score ${this.runState.score}   Lives ${this.runState.lives}   Clams ${this.runState.collectedTreasureIds.length}/5   Time ${seconds}`,
    );
    this.statusLabel.setText(this.isPaused ? "Paused" : "P pause  R restart");
  }

  private updatePlayerPose(): void {
    const texture = !this.isGrounded && !this.isClimbing ? "player-jump" : "player-stand";
    this.player.setTexture(texture);
    this.player.setFlipX(this.facing < 0);
  }

  private resolveHazardTravelBounds(hazard: HazardDefinition): { minX: number; maxX: number } {
    const supportingPlatform = this.findHazardPlatform(hazard);
    if (!supportingPlatform) {
      return { minX: hazard.minX, maxX: hazard.maxX };
    }

    const platformSpan = this.getPlatformSpan(supportingPlatform, this.getHazardEdgeInset(hazard));
    const minX = Math.max(hazard.minX, platformSpan.minX);
    const maxX = Math.min(hazard.maxX, platformSpan.maxX);

    if (minX <= maxX) {
      return { minX, maxX };
    }

    const fallbackX = Phaser.Math.Clamp(hazard.x, platformSpan.minX, platformSpan.maxX);
    return { minX: fallbackX, maxX: fallbackX };
  }

  private findHazardPlatform(hazard: HazardDefinition): PlatformDefinition | null {
    const hazardBottom = hazard.y + hazard.height / 2;
    const laneCenter = (hazard.minX + hazard.maxX) / 2;
    const laneWidth = Math.max(hazard.maxX - hazard.minX, hazard.width);
    let bestPlatform: PlatformDefinition | null = null;
    let bestScore = Number.POSITIVE_INFINITY;

    for (const platform of this.currentRoom.platforms) {
      const span = this.getPlatformSpan(platform);
      const top = this.getPlatformTop(platform);
      const overlapsLane = hazard.maxX >= span.minX && hazard.minX <= span.maxX;
      const nearTop = Math.abs(hazardBottom - top) <= 18;

      if (!overlapsLane || !nearTop) {
        continue;
      }

      const platformCenter = (span.minX + span.maxX) / 2;
      const score = Math.abs(top - hazardBottom) * 8 + Math.abs(platformCenter - laneCenter) + Math.abs(platform.width - laneWidth) * 0.05;
      if (score < bestScore) {
        bestPlatform = platform;
        bestScore = score;
      }
    }

    return bestPlatform;
  }

  private getPlatformSpan(platform: PlatformDefinition, inset = 0): HorizontalSpan {
    const left = platform.x - platform.width / 2;
    const right = platform.x + platform.width / 2;
    const clampedInset = Math.min(inset, platform.width / 2 - 1);

    return {
      minX: left + clampedInset,
      maxX: right - clampedInset,
    };
  }

  private getPlatformTop(platform: PlatformDefinition): number {
    return platform.y - platform.height / 2;
  }

  private getHazardEdgeInset(hazard: HazardDefinition): number {
    return Math.max(8, hazard.width / 2 - 4);
  }

  private getPlayerBounds(): Phaser.Geom.Rectangle {
    return new Phaser.Geom.Rectangle(
      this.player.x - GAME_CONFIG.player.width / 2,
      this.player.y - GAME_CONFIG.player.height / 2,
      GAME_CONFIG.player.width,
      GAME_CONFIG.player.height,
    );
  }

  private createPlatformView(platform: PlatformDefinition): Phaser.GameObjects.Container {
    const textureKey = Phaser.Utils.Array.GetRandom(["wood-plank-a", "wood-plank-b", "wood-plank-c"]);
    const left = platform.x - platform.width / 2;
    const top = platform.y - platform.height / 2;
    const container = this.add.container(left, top);
    const plank = this.add.tileSprite(platform.width / 2, platform.height / 2, platform.width, platform.height, textureKey);
    plank.tilePositionX = Phaser.Math.Between(0, 31);
    const topLip = this.add.rectangle(platform.width / 2, 2, platform.width, 3, 0xd4a373);
    topLip.setAlpha(0.45);
    const bottomShadow = this.add.rectangle(platform.width / 2, platform.height - 2, platform.width, 4, 0x3b2418);
    bottomShadow.setAlpha(0.65);

    container.add([plank, topLip, bottomShadow]);
    return container;
  }

  private createTreasureSparkles(sprite: Phaser.GameObjects.Image): Phaser.GameObjects.Image[] {
    const sparkleConfigs = [
      { key: "glint-8", offsetX: -18, offsetY: -16, baseScale: 0.9, duration: 260, repeatDelay: 900, angle: 0 },
      { key: "glint-4", offsetX: 20, offsetY: -6, baseScale: 0.62, duration: 220, repeatDelay: 760, angle: 14 },
      { key: "glint-4", offsetX: 6, offsetY: 18, baseScale: 0.52, duration: 210, repeatDelay: 1040, angle: -12 },
    ] as const;

    return sparkleConfigs.map((config, index) => {
      const sparkle = this.add.image(sprite.x + config.offsetX, sprite.y + config.offsetY, config.key);
      sparkle.setDepth(6 + index);
      sparkle.setScale(0.15);
      sparkle.setAlpha(0);
      sparkle.setAngle(config.angle);

      this.tweens.add({
        targets: sparkle,
        alpha: { from: 0, to: 1 },
        scaleX: { from: 0.15, to: config.baseScale },
        scaleY: { from: 0.15, to: config.baseScale },
        angle: config.angle + (index % 2 === 0 ? 12 : -12),
        ease: "Sine.easeOut",
        yoyo: true,
        repeat: -1,
        duration: config.duration,
        repeatDelay: config.repeatDelay,
        delay: 180 + index * 170,
      });

      return sparkle;
    });
  }
}
