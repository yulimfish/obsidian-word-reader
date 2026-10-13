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

export interface MetafileTextOperation {
  readonly kind: "text";
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly color: MetafileColor;
  readonly direction: "ltr" | "rtl";
  readonly advances?: readonly number[];
  readonly transform?: MetafileTransform;
}

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
