declare module "jszip/lib/defaults" {
  const defaults: Record<string, unknown>;
  export default defaults;
}

declare module "jszip/lib/load" {
  import type JSZip from "jszip";

  const loadAsync: <Host>(
    this: Host,
    data: ArrayBuffer,
    options: JSZip.JSZipLoadOptions,
  ) => Promise<Host>;
  export default loadAsync;
}

declare module "jszip/lib/zipObject" {
  import type JSZip from "jszip";

  const ZipObject: new (
    name: string,
    data: unknown,
    options: Record<string, unknown>,
  ) => JSZip.JSZipObject;
  export default ZipObject;
}
