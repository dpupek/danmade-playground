import Phaser from "phaser";

interface ToneStep {
  frequency: number;
  duration: number;
  type?: OscillatorType;
  volume?: number;
}

export class RetroSfx {
  constructor(private readonly scene: Phaser.Scene) {}

  jump(): void {
    this.play([
      { frequency: 420, duration: 0.05, type: "square", volume: 0.04 },
      { frequency: 620, duration: 0.08, type: "square", volume: 0.03 },
    ]);
  }

  pickup(): void {
    this.play([
      { frequency: 660, duration: 0.05, type: "triangle", volume: 0.04 },
      { frequency: 880, duration: 0.08, type: "triangle", volume: 0.03 },
    ]);
  }

  hurt(): void {
    this.play([
      { frequency: 180, duration: 0.08, type: "sawtooth", volume: 0.05 },
      { frequency: 120, duration: 0.12, type: "sawtooth", volume: 0.04 },
    ]);
  }

  humanScream(): void {
    this.play([
      { frequency: 540, duration: 0.05, type: "triangle", volume: 0.04 },
      { frequency: 820, duration: 0.07, type: "sawtooth", volume: 0.045 },
      { frequency: 680, duration: 0.08, type: "triangle", volume: 0.04 },
      { frequency: 420, duration: 0.12, type: "sawtooth", volume: 0.03 },
    ]);
  }

  win(): void {
    this.play([
      { frequency: 520, duration: 0.06, type: "triangle", volume: 0.04 },
      { frequency: 660, duration: 0.06, type: "triangle", volume: 0.04 },
      { frequency: 880, duration: 0.12, type: "triangle", volume: 0.03 },
    ]);
  }

  private play(steps: ToneStep[]): void {
    if (!("context" in this.scene.sound)) {
      return;
    }

    const context = this.scene.sound.context;
    if (!context || context.state === "closed") {
      return;
    }

    if (context.state === "suspended") {
      void context.resume();
    }

    let startTime = context.currentTime;
    for (const step of steps) {
      const oscillator = context.createOscillator();
      const gain = context.createGain();
      oscillator.type = step.type ?? "square";
      oscillator.frequency.value = step.frequency;
      gain.gain.value = step.volume ?? 0.03;
      oscillator.connect(gain);
      gain.connect(context.destination);
      oscillator.start(startTime);
      oscillator.stop(startTime + step.duration);
      startTime += step.duration;
    }
  }
}
