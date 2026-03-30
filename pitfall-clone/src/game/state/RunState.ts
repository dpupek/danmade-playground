import { GAME_CONFIG } from "../config";
import { ROOM_ORDER, getRoomDefinition } from "../data/rooms";
import type { RoomId, RunState } from "../types";

export class RunStateController {
  private state: RunState;

  constructor(initialState?: RunState) {
    this.state = initialState ?? RunStateController.createInitialState();
  }

  static createInitialState(): RunState {
    return {
      currentRoomId: ROOM_ORDER[0],
      lives: GAME_CONFIG.run.lives,
      score: 0,
      timeRemainingMs: GAME_CONFIG.run.timerSeconds * 1000,
      collectedTreasureIds: [],
      status: "playing",
    };
  }

  get snapshot(): RunState {
    return {
      ...this.state,
      collectedTreasureIds: [...this.state.collectedTreasureIds],
    };
  }

  moveToRoom(roomId: RoomId): RunState {
    getRoomDefinition(roomId);
    this.state.currentRoomId = roomId;
    return this.snapshot;
  }

  collectTreasure(treasureId: string): RunState {
    if (!this.state.collectedTreasureIds.includes(treasureId)) {
      this.state.collectedTreasureIds = [...this.state.collectedTreasureIds, treasureId];
      this.state.score += GAME_CONFIG.run.treasureScore;
    }

    if (this.state.collectedTreasureIds.length === this.totalTreasureCount) {
      this.state.status = "won";
    }

    return this.snapshot;
  }

  tick(deltaMs: number): RunState {
    if (this.state.status !== "playing") {
      return this.snapshot;
    }

    this.state.timeRemainingMs = Math.max(0, this.state.timeRemainingMs - deltaMs);
    if (this.state.timeRemainingMs === 0) {
      this.state.status = "lost";
    }

    return this.snapshot;
  }

  loseLife(): RunState {
    this.state.lives -= 1;
    if (this.state.lives <= 0) {
      this.state.lives = 0;
      this.state.status = "lost";
    }
    return this.snapshot;
  }

  hasCollected(treasureId: string): boolean {
    return this.state.collectedTreasureIds.includes(treasureId);
  }

  get totalTreasureCount(): number {
    return ROOM_ORDER.reduce((count, roomId) => count + getRoomDefinition(roomId).treasures.length, 0);
  }
}
