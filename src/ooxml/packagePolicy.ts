import JSZip from "jszip";
import { inspectEquationPreviews } from "./equationPreviews";
import type { PreviewInventory } from "./previewTypes";

export type OoxmlFormat = "docx" | "pptx" | "xlsx";

export class OoxmlPolicyError extends Error {
  constructor(
    readonly kind: "active-content" | "ole-object" | "script-media",
    message: string,
  ) {
    super(message);
    this.name = "OoxmlPolicyError";
  }
}

const SCRIPT_MEDIA_EXTENSIONS = new Set([
  "bat",
  "cmd",
  "com",
  "dll",
  "exe",
  "hta",
  "htm",
  "html",
  "jar",
  "js",
  "mht",
  "mhtml",
  "msi",
  "ps1",
  "svg",
  "swf",
  "vbe",
  "vbs",
  "wsf",
]);

export async function enforceOoxmlPackagePolicy(
  zip: JSZip,
  format: OoxmlFormat,
  options?: { legacyEquationPreviews?: boolean },
): Promise<PreviewInventory | null> {
  const allowPreviews = format !== "xlsx" && options?.legacyEquationPreviews === true;
  const names = Object.keys(zip.files);
  for (const name of names) {
    const lowerName = name.toLowerCase();
    if (
      lowerName.endsWith("/vbaproject.bin") ||
      lowerName.includes("/macrosheets/") ||
      lowerName.includes("/dialogsheets/") ||
      lowerName.includes("/activex/")
    ) {
      throw new OoxmlPolicyError(
        "active-content",
        `Macro and ActiveX content is not supported in ${format.toUpperCase()} previews.`,
      );
    }
    if (
      !allowPreviews && (lowerName.includes("/embeddings/") ||
      lowerName.includes("/oleobjects/"))
    ) {
      throw new OoxmlPolicyError(
        "ole-object",
        `Embedded OLE objects are not supported in ${format.toUpperCase()} previews.`,
      );
    }
    const extension = lowerName.split(".").at(-1);
    if (
      lowerName.includes("/media/") &&
      extension &&
      SCRIPT_MEDIA_EXTENSIONS.has(extension)
    ) {
      throw new OoxmlPolicyError(
        "script-media",
        `Script-capable media is not supported in ${format.toUpperCase()} previews.`,
      );
    }
  }

  const contentTypesEntry = zip.file("[Content_Types].xml");
  const contentTypes = contentTypesEntry ? (await contentTypesEntry.async("string")).toLowerCase() : "";
  if (
    contentTypes.includes("macroenabled") ||
    contentTypes.includes("vbaproject") ||
    contentTypes.includes("activex")
  ) {
    throw new OoxmlPolicyError(
      "active-content",
      `Macro and ActiveX content is not supported in ${format.toUpperCase()} previews.`,
    );
  }
  if (
    !allowPreviews && (contentTypes.includes("oleobject") ||
    contentTypes.includes("ole-object"))
  ) {
    throw new OoxmlPolicyError(
      "ole-object",
      `Embedded OLE objects are not supported in ${format.toUpperCase()} previews.`,
    );
  }
  if (format !== "xlsx") {
    try {
      const inventory = await inspectEquationPreviews(zip, format);
      if (allowPreviews) return inventory;
      if (inventory.occurrences.length) throw new Error("Embedded OLE objects are not supported.");
    } catch (error) {
      throw new OoxmlPolicyError("ole-object", error instanceof Error ? error.message : "Invalid equation preview metadata.");
    }
  }
  return null;
}
