import { GAME_CONFIG } from "../config";
import type { RoomDefinition, RoomId } from "../types";

const floorY = GAME_CONFIG.world.floorY;

export const ROOM_ORDER: RoomId[] = [
  "canopy-gate",
  "ember-bridge",
  "monkey-step",
  "sunken-vault",
  "idol-hall",
];

export const ROOM_DEFINITIONS: Record<RoomId, RoomDefinition> = {
  "canopy-gate": {
    id: "canopy-gate",
    title: "Canopy Gate",
    platforms: [
      { x: 120, y: floorY, width: 240, height: 28 },
      { x: 530, y: floorY, width: 520, height: 28 },
      { x: 740, y: 340, width: 180, height: 24 },
    ],
    ladders: [{ x: 740, y: 404, width: 28, height: 128 }],
    hazards: [
      {
        id: "gate-log",
        type: "log",
        x: 640,
        y: floorY - 14,
        width: 40,
        height: 24,
        minX: 560,
        maxX: 860,
      },
    ],
    treasures: [{ id: "jade-mask", x: 810, y: 298, label: "Jade Mask" }],
    exits: { right: "ember-bridge" },
    spawn: {
      default: { x: 120, y: floorY - 42 },
      fromLeft: { x: 120, y: floorY - 42 },
      fromRight: { x: 700, y: floorY - 42 },
    },
  },
  "ember-bridge": {
    id: "ember-bridge",
    title: "Ember Bridge",
    platforms: [
      { x: 140, y: floorY, width: 280, height: 28 },
      { x: 520, y: floorY, width: 220, height: 28 },
      { x: 840, y: floorY, width: 200, height: 28 },
      { x: 470, y: 320, width: 160, height: 24 },
    ],
    ladders: [{ x: 470, y: 394, width: 28, height: 148 }],
    hazards: [
      {
        id: "ember-scorpion",
        type: "scorpion",
        x: 470,
        y: floorY - 16,
        width: 34,
        height: 20,
        minX: 390,
        maxX: 610,
        speed: 100,
      },
      {
        id: "ember-log",
        type: "log",
        x: 510,
        y: 304,
        width: 40,
        height: 24,
        minX: 410,
        maxX: 530,
      },
    ],
    treasures: [{ id: "sun-disc", x: 560, y: 278, label: "Sun Disc" }],
    exits: { left: "canopy-gate", right: "monkey-step" },
    spawn: {
      default: { x: 140, y: floorY - 42 },
      fromLeft: { x: 140, y: floorY - 42 },
      fromRight: { x: 840, y: floorY - 42 },
    },
  },
  "monkey-step": {
    id: "monkey-step",
    title: "Monkey Step",
    platforms: [
      { x: 120, y: floorY, width: 200, height: 28 },
      { x: 500, y: floorY, width: 320, height: 28 },
      { x: 270, y: 360, width: 180, height: 24 },
      { x: 690, y: 280, width: 170, height: 24 },
    ],
    ladders: [
      { x: 270, y: 404, width: 28, height: 120 },
      { x: 690, y: 322, width: 28, height: 122 },
    ],
    hazards: [
      {
        id: "monkey-log",
        type: "log",
        x: 580,
        y: floorY - 14,
        width: 40,
        height: 24,
        minX: 510,
        maxX: 760,
      },
      {
        id: "upper-scorpion",
        type: "scorpion",
        x: 290,
        y: 344,
        width: 34,
        height: 20,
        minX: 210,
        maxX: 350,
        speed: 80,
      },
    ],
    treasures: [{ id: "amber-idol", x: 760, y: 238, label: "Amber Idol" }],
    exits: { left: "ember-bridge", right: "sunken-vault" },
    spawn: {
      default: { x: 120, y: floorY - 42 },
      fromLeft: { x: 120, y: floorY - 42 },
      fromRight: { x: 600, y: floorY - 42 },
    },
  },
  "sunken-vault": {
    id: "sunken-vault",
    title: "Sunken Vault",
    platforms: [
      { x: 120, y: floorY, width: 200, height: 28 },
      { x: 460, y: floorY, width: 180, height: 28 },
      { x: 820, y: floorY, width: 220, height: 28 },
      { x: 310, y: 330, width: 160, height: 24 },
      { x: 640, y: 250, width: 180, height: 24 },
    ],
    ladders: [
      { x: 310, y: 394, width: 28, height: 146 },
      { x: 640, y: 292, width: 28, height: 124 },
    ],
    hazards: [
      {
        id: "vault-scorpion",
        type: "scorpion",
        x: 500,
        y: floorY - 16,
        width: 34,
        height: 20,
        minX: 430,
        maxX: 560,
        speed: 90,
      },
      {
        id: "vault-log",
        type: "log",
        x: 720,
        y: 234,
        width: 40,
        height: 24,
        minX: 580,
        maxX: 760,
      },
    ],
    treasures: [{ id: "moon-gem", x: 720, y: 208, label: "Moon Gem" }],
    exits: { left: "monkey-step", right: "idol-hall" },
    spawn: {
      default: { x: 120, y: floorY - 42 },
      fromLeft: { x: 120, y: floorY - 42 },
      fromRight: { x: 820, y: floorY - 42 },
    },
  },
  "idol-hall": {
    id: "idol-hall",
    title: "Idol Hall",
    platforms: [
      { x: 120, y: floorY, width: 240, height: 28 },
      { x: 530, y: floorY, width: 520, height: 28 },
      { x: 510, y: 330, width: 180, height: 24 },
      { x: 820, y: 250, width: 160, height: 24 },
    ],
    ladders: [
      { x: 510, y: 394, width: 28, height: 146 },
      { x: 820, y: 292, width: 28, height: 124 },
    ],
    hazards: [
      {
        id: "idol-log",
        type: "log",
        x: 620,
        y: floorY - 14,
        width: 40,
        height: 24,
        minX: 540,
        maxX: 850,
      },
      {
        id: "idol-scorpion",
        type: "scorpion",
        x: 540,
        y: 314,
        width: 34,
        height: 20,
        minX: 430,
        maxX: 590,
        speed: 85,
      },
    ],
    treasures: [{ id: "sky-crown", x: 860, y: 208, label: "Sky Crown" }],
    exits: { left: "sunken-vault" },
    spawn: {
      default: { x: 120, y: floorY - 42 },
      fromLeft: { x: 120, y: floorY - 42 },
      fromRight: { x: 700, y: floorY - 42 },
    },
  },
};

export function getRoomDefinition(roomId: RoomId): RoomDefinition {
  return ROOM_DEFINITIONS[roomId];
}
