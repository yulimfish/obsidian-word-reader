export interface MetafilePoint {
  readonly x: number;
  readonly y: number;
}

export interface MetafileViewBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface MetafileTransform {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export interface MetafileColor {
  readonly red: number;
  readonly green: number;
  readonly blue: number;
  readonly alpha: number;
}

export type MetafilePathCommand =
  | { readonly type: "moveTo"; readonly point: MetafilePoint }
  | { readonly type: "lineTo"; readonly point: MetafilePoint }
  | {
      readonly type: "cubicTo";
      readonly control1: MetafilePoint;
      readonly control2: MetafilePoint;
      readonly point: MetafilePoint;
    }
  | { readonly type: "closePath" };

export interface MetafileStroke {
  readonly color: MetafileColor;
  readonly width: number;
  readonly lineCap: "butt" | "round" | "square";
  readonly lineJoin: "bevel" | "miter" | "round";
  readonly miterLimit: number;
}

export interface MetafilePathOperation {
  readonly kind: "path";
  readonly commands: readonly MetafilePathCommand[];
  readonly fill: MetafileColor | null;
  readonly stroke: MetafileStroke | null;
  readonly transform?: MetafileTransform;
}

export interface MetafileLogicalFont {
  readonly height: number;
  readonly width: number;
  readonly escapement: number;
  readonly orientation: number;
  readonly weight: number;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strikeOut: boolean;
  readonly charset: number;
  readonly outPrecision: number;
  readonly clipPrecision: number;
  readonly quality: number;
  readonly pitchAndFamily: number;
}

interface MetafileTextStyle {
  readonly kind: "text";
  readonly x: number;
  readonly y: number;
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly color: MetafileColor;
  readonly direction: "ltr" | "rtl";
  readonly advances?: readonly number[];
  readonly transform?: MetafileTransform;
  readonly logicalFont?: MetafileLogicalFont;
  readonly codePage?: number;
  readonly sourceBytes?: readonly number[];
  readonly sourceAdvances?: readonly number[];
  readonly textAlign?: number;
  readonly textOptions?: number;
  readonly backgroundMode?: "transparent" | "opaque";
  readonly backgroundColor?: MetafileColor;
  readonly clipRect?: MetafileViewBox;
  readonly opaqueRect?: MetafileViewBox;
}

// A font-specific byte run is NOT Unicode text. Consumers must resolve the
// exact font cmap or report its glyph gap; it cannot be painted as Latin-1.
export type MetafileTextOperation = MetafileTextStyle & (
  | { readonly text: string; readonly encoding?: "unicode" }
  | { readonly text: null; readonly encoding: "font-specific"; readonly sourceBytes: readonly number[]; readonly codePage: 42 }
);

export type MetafileClipOperation =
  | {
      readonly kind: "clip";
      readonly action: "set" | "intersect" | "exclude";
      readonly commands: readonly MetafilePathCommand[];
      readonly fillRule: "evenOdd" | "nonZero";
    }
  | { readonly kind: "clip"; readonly action: "reset" };

export type MetafileSceneOperation =
  | MetafilePathOperation
  | MetafileTextOperation
  | MetafileClipOperation;

export type MetafileSceneWarning =
  | { readonly kind: "missing-font"; readonly fontName: string }
  | { readonly kind: "unconfirmed-font"; readonly fontName: string }
  | { readonly kind: "font-metrics-required"; readonly fontName: string; readonly recordOffset: number; readonly recordType: number }
  | { readonly kind: "font-glyph-gap"; readonly fontName: string; readonly charset: number; readonly byteCodes: readonly number[]; readonly recordOffset: number; readonly recordType: number }
  | { readonly kind: "opaque-comment"; readonly byteLength: number; readonly recordOffset: number; readonly recordType: number }
  | {
      readonly kind: "unsupported-record";
      readonly recordOffset: number;
      readonly recordType: number;
    };

export interface MetafileScene {
  readonly viewBox: MetafileViewBox;
  readonly operations: readonly MetafileSceneOperation[];
  readonly requiredFonts: readonly string[];
  readonly warnings: readonly MetafileSceneWarning[];
}

export class MetafileDecodeError extends Error {
  constructor(
    readonly code: string,
    readonly recordOffset: number,
    readonly recordType: number,
    message: string,
  ) {
    super(`${message} (record 0x${recordType.toString(16)} at byte ${recordOffset}).`);
    this.name = "MetafileDecodeError";
  }
}
