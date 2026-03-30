export type RoomId = "canopy-gate" | "ember-bridge" | "monkey-step" | "sunken-vault" | "idol-hall";

export interface PlatformDefinition {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface LadderDefinition {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TreasureDefinition {
  id: string;
  x: number;
  y: number;
  label: string;
}

export type HazardType = "paul_crab" | "dave_goat" | "mark_wasp";

export interface HazardDefinition {
  id: string;
  type: HazardType;
  x: number;
  y: number;
  width: number;
  height: number;
  minX: number;
  maxX: number;
  speed?: number;
  swoopDepth?: number;
  swoopRate?: number;
  chargePauseMs?: number;
}

export interface RoomDefinition {
  id: RoomId;
  title: string;
  platforms: PlatformDefinition[];
  ladders: LadderDefinition[];
  hazards: HazardDefinition[];
  treasures: TreasureDefinition[];
  exits: {
    left?: RoomId;
    right?: RoomId;
  };
  spawn: {
    default: { x: number; y: number };
    fromLeft?: { x: number; y: number };
    fromRight?: { x: number; y: number };
  };
}

export interface RunState {
  currentRoomId: RoomId;
  lives: number;
  score: number;
  timeRemainingMs: number;
  collectedTreasureIds: string[];
  status: "title" | "playing" | "won" | "lost";
}

export interface ScenePayload {
  runState?: RunState;
}

export interface EndScenePayload {
  outcome: "won" | "lost";
  runState: RunState;
}
