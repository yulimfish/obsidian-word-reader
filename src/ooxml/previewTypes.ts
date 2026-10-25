export interface PreviewIssue {
  partPath: string;
  occurrenceId: string;
  code: string;
  detail: string;
}

export interface PreviewOccurrence {
  occurrenceId: string;
  partPath: string;
  shapeId: string;
  embeddingPath: string;
  previewPath: string | null;
  previewRelationshipId: string | null;
  progId: "Equation.3" | "Equation.DSMT4";
  sourceElementPath: string;
}

export interface PreviewInventory {
  occurrences: readonly PreviewOccurrence[];
  issues: readonly PreviewIssue[];
}
