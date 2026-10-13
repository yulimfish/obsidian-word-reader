import JSZip from "jszip";

const PACKAGE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MARKUP_COMPATIBILITY_NAMESPACE =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const ALTERNATIVE_NAMESPACE =
  "http://schemas.microsoft.com/office/drawing/2010/main";
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1, 0, 0, 0));

export interface EquationFixtureOptions {
  progId: "Equation.3" | "Equation.DSMT4";
  previewRelationship: {
    id: string;
    target: string;
  };
  previewBytes: Uint8Array;
  opaqueEmbeddingBytes: Uint8Array;
  alternative?: boolean;
  duplicate?: boolean;
  missingPreview?: boolean;
  unsafeTarget?: boolean;
}

type FixtureFormat = "docx" | "pptx";

export async function createEquationFixture(
  format: FixtureFormat,
  options: EquationFixtureOptions,
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const isDocx = format === "docx";
  const documentPath = isDocx ? "word/document.xml" : "ppt/slides/slide1.xml";
  const relationshipPath = isDocx
    ? "word/_rels/document.xml.rels"
    : "ppt/slides/_rels/slide1.xml.rels";
  const mediaPath = isDocx
    ? "word/media/preview.png"
    : "ppt/media/preview.png";
  const objectDirectory = isDocx ? "word/embeddings" : "ppt/embeddings";
  const previewRelationshipTarget = options.unsafeTarget
    ? "../../unsafe-preview.png"
    : options.previewRelationship.target;
  const occurrenceCount = options.duplicate ? 2 : 1;

  addFile(zip, "[Content_Types].xml", contentTypesXml(format, occurrenceCount));
  addFile(zip, "_rels/.rels", rootRelationshipsXml(format));
  if (isDocx) {
    addFile(zip, documentPath, docxDocumentXml(options, occurrenceCount));
  } else {
    addFile(zip, "ppt/presentation.xml", presentationXml());
    addFile(zip, "ppt/_rels/presentation.xml.rels", presentationRelationshipsXml());
    addFile(zip, documentPath, pptxSlideXml(options, occurrenceCount));
  }
  addFile(
    zip,
    relationshipPath,
    documentRelationshipsXml(
      format,
      options,
      occurrenceCount,
      previewRelationshipTarget,
    ),
  );

  if (!options.missingPreview) {
    addFile(zip, mediaPath, options.previewBytes);
  }
  for (let index = 1; index <= occurrenceCount; index += 1) {
    const embeddingBytes =
      index === 1
        ? options.opaqueEmbeddingBytes
        : Uint8Array.from([...options.opaqueEmbeddingBytes, index]);
    addFile(
      zip,
      `${objectDirectory}/oleObject${index}.bin`,
      embeddingBytes,
    );
  }

  return zip.generateAsync({
    type: "arraybuffer",
    compression: "STORE",
    platform: "DOS",
    streamFiles: false,
  });
}

function addFile(
  zip: JSZip,
  path: string,
  content: string | Uint8Array,
): void {
  zip.file(path, content, {
    createFolders: false,
    date: new Date(FIXED_ZIP_DATE.getTime()),
    compression: "STORE",
  });
}

function contentTypesXml(format: FixtureFormat, occurrenceCount: number): string {
  const overrides =
    format === "docx"
      ? '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      : [
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
          '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
        ].join("\n  ");
  const embeddingOverrides = Array.from(
    { length: occurrenceCount },
    (_, index) =>
      `<Override PartName="/${format === "docx" ? "word" : "ppt"}/embeddings/oleObject${index + 1}.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>`,
  ).join("\n  ");

  return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  ${overrides}
  ${embeddingOverrides}
</Types>`;
}

function rootRelationshipsXml(format: FixtureFormat): string {
  const target = format === "docx" ? "word/document.xml" : "ppt/presentation.xml";
  return relationshipsXml([
    {
      id: "rId1",
      type: `${OFFICE_RELATIONSHIP_NAMESPACE}/officeDocument`,
      target,
    },
  ]);
}

function presentationXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:r="${OFFICE_RELATIONSHIP_NAMESPACE}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
</p:presentation>`;
}

function presentationRelationshipsXml(): string {
  return relationshipsXml([
    {
      id: "rId1",
      type: `${OFFICE_RELATIONSHIP_NAMESPACE}/slide`,
      target: "slides/slide1.xml",
    },
  ]);
}

function documentRelationshipsXml(
  format: FixtureFormat,
  options: EquationFixtureOptions,
  occurrenceCount: number,
  previewTarget: string,
): string {
  const relationships: Array<{ id: string; type: string; target: string }> = [];
  if (!options.missingPreview) {
    relationships.push({
      id: options.previewRelationship.id,
      type: `${OFFICE_RELATIONSHIP_NAMESPACE}/image`,
      target: previewTarget,
    });
  }

  for (let index = 1; index <= occurrenceCount; index += 1) {
    relationships.push({
      id: `rIdOle${index}`,
      type: `${OFFICE_RELATIONSHIP_NAMESPACE}/oleObject`,
      target: `${format === "docx" ? "embeddings" : "../embeddings"}/oleObject${index}.bin`,
    });
  }
  return relationshipsXml(relationships);
}

function relationshipsXml(
  entries: Array<{ id: string; type: string; target: string }>,
): string {
  const children = entries
    .map(
      ({ id, type, target }) =>
        `  <Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`,
    )
    .join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">
${children}
</Relationships>`;
}

function docxDocumentXml(
  options: EquationFixtureOptions,
  occurrenceCount: number,
): string {
  const objects = Array.from({ length: occurrenceCount }, (_, index) => {
    const shapeId = `Shape${index + 1}`;
    const object = `<w:object>
      <v:shape id="${shapeId}" style="width:36pt;height:18pt"><v:imagedata r:id="${escapeXml(options.previewRelationship.id)}" o:title=""/></v:shape>
      <o:OLEObject Type="Embed" ProgID="${escapeXml(options.progId)}" ShapeID="${shapeId}" ObjectID="_Equation${index + 1}" r:id="rIdOle${index + 1}"/>
    </w:object>`;
    return wrapAlternative(object, options.alternative);
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${OFFICE_RELATIONSHIP_NAMESPACE}" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" xmlns:a14="${ALTERNATIVE_NAMESPACE}">
  <w:body><w:p><w:r>${objects}</w:r></w:p><w:sectPr/></w:body>
</w:document>`;
}

function pptxSlideXml(
  options: EquationFixtureOptions,
  occurrenceCount: number,
): string {
  const shapes = Array.from({ length: occurrenceCount }, (_, index) => {
    const id = index + 2;
    const graphicFrame = `<p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Equation${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
        <p:oleObj progId="${escapeXml(options.progId)}" r:id="rIdOle${index + 1}"><p:embed/><p:pic>
          <p:nvPicPr><p:cNvPr id="${id + 20}" name="EquationPreview${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="${escapeXml(options.previewRelationship.id)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        </p:pic></p:oleObj>
      </a:graphicData></a:graphic>
    </p:graphicFrame>`;
    return wrapAlternative(graphicFrame, options.alternative);
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${OFFICE_RELATIONSHIP_NAMESPACE}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" xmlns:a14="${ALTERNATIVE_NAMESPACE}">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld>
</p:sld>`;
}

function wrapAlternative(markup: string, enabled = false): string {
  if (!enabled) {
    return markup;
  }
  return `<mc:AlternateContent>
    <mc:Choice Requires="a14">${markup}</mc:Choice>
    <mc:Fallback>${markup}</mc:Fallback>
  </mc:AlternateContent>`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
