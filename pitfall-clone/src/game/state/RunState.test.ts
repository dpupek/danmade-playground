import { describe, expect, it } from "vitest";
import { ROOM_ORDER } from "../data/rooms";
import { RunStateController } from "./RunState";

describe("RunStateController", () => {
  it("starts in the first room with fresh run stats", () => {
    const controller = new RunStateController();
    const state = controller.snapshot;

    expect(state.currentRoomId).toBe(ROOM_ORDER[0]);
    expect(state.lives).toBe(3);
    expect(state.score).toBe(0);
    expect(state.status).toBe("playing");
  });

  it("collects treasure once and wins after all treasures are found", () => {
    const controller = new RunStateController();
    const ids = ["jade-mask", "sun-disc", "amber-idol", "moon-gem", "sky-crown"];

    ids.forEach((id) => controller.collectTreasure(id));

    const state = controller.snapshot;
    expect(state.collectedTreasureIds).toHaveLength(ids.length);
    expect(state.score).toBe(250 * ids.length);
    expect(state.status).toBe("won");
  });

  it("counts down time and loses when it reaches zero", () => {
    const controller = new RunStateController();

    controller.tick(140_000);

    expect(controller.snapshot.timeRemainingMs).toBe(0);
    expect(controller.snapshot.status).toBe("lost");
  });

  it("drops lives and loses the run at zero", () => {
    const controller = new RunStateController();

    controller.loseLife();
    controller.loseLife();
    controller.loseLife();

    expect(controller.snapshot.lives).toBe(0);
    expect(controller.snapshot.status).toBe("lost");
  });

  it("moves between rooms without resetting collected state", () => {
    const controller = new RunStateController();

    controller.collectTreasure("jade-mask");
    controller.moveToRoom("monkey-step");

    expect(controller.snapshot.currentRoomId).toBe("monkey-step");
    expect(controller.snapshot.collectedTreasureIds).toEqual(["jade-mask"]);
  });

  it("ignores duplicate treasure pickups", () => {
    const controller = new RunStateController();

    controller.collectTreasure("jade-mask");
    controller.collectTreasure("jade-mask");

    expect(controller.snapshot.collectedTreasureIds).toEqual(["jade-mask"]);
    expect(controller.snapshot.score).toBe(250);
  });
});
