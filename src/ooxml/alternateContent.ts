import { childElements } from "../pptx/xml";

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_DEPTH = 128;
const MAX_ALTERNATIVES = 10_000;

export interface AlternativeSelection {
  sourceElementPath: string;
  branch: "choice" | "fallback";
  selectedElementPath: string;
}

export interface AlternativeIssue {
  sourceElementPath: string;
  code: string;
}

export function normalizeAlternateContent(
  document: Document,
  supportedNamespaces: ReadonlySet<string>,
): {
  document: Document;
  selections: readonly AlternativeSelection[];
  issues: readonly AlternativeIssue[];
} {
  const clone = document.cloneNode(true) as Document;
  const selections: AlternativeSelection[] = [];
  const issues: AlternativeIssue[] = [];
  const chosenBranches = new Map<Element, Element>();
  let count = 0;
  const root = clone.documentElement;
  if (!root) return { document: clone, selections, issues };
  const stack: Array<{ element: Element; path: string; depth: number }> = [
    { element: root, path: `/${root.nodeName}[1]`, depth: 0 },
  ];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) break;
    const { element, path, depth } = current;
    const children = childElements(element);
    const counters = new Map<string, number>();
    const paths = new Map<Element, string>();
    for (const child of children) {
      const index = (counters.get(child.nodeName) ?? 0) + 1;
      counters.set(child.nodeName, index);
      paths.set(child, `${path}/${child.nodeName}[${index}]`);
    }
    if (element.namespaceURI !== MC || element.localName !== "AlternateContent") {
      for (const child of children.reverse()) {
        stack.push({ element: child, path: paths.get(child) ?? path, depth });
      }
      continue;
    }
    count += 1;
    const issue = (code: string): void => { issues.push({ sourceElementPath: path, code }); };
    if (depth >= MAX_DEPTH || count > MAX_ALTERNATIVES) {
      issue(depth >= MAX_DEPTH ? "depth-limit" : "count-limit");
      element.parentNode?.removeChild(element);
      continue;
    }
    let selected: Element | undefined;
    for (const child of children) {
      if (child.namespaceURI !== MC || child.localName !== "Choice") continue;
      const requires = (child.getAttribute("Requires") ?? "").trim().split(/\s+/).filter(Boolean);
      const uris = requires.map((prefix) => child.lookupNamespaceURI(prefix));
      if (requires.length === 0 || uris.some((uri) => !uri)) {
        issue("unresolved-requires");
      } else if (uris.every((uri) => uri !== null && supportedNamespaces.has(uri))) {
        selected = child;
        break;
      } else {
        issue("unsupported-requires");
      }
    }
    selected ??= children.find((child) => child.namespaceURI === MC && child.localName === "Fallback");
    if (!selected) {
      issue("missing-fallback");
      element.parentNode?.removeChild(element);
      continue;
    }
    const selectedPath = paths.get(selected) ?? path;
    chosenBranches.set(element, selected);
    selections.push({ sourceElementPath: path, branch: selected.localName === "Choice" ? "choice" : "fallback", selectedElementPath: selectedPath });
    // Visit before unwrapping so diagnostics retain the original branch paths.
    stack.push({ element: selected, path: selectedPath, depth: depth + 1 });
    // The branch wrapper remains temporarily attached to the detached alternative.
    // Moving its children now would lose the nested traversal; unwrap after traversal.
  }
  // Unwrap in reverse order, including detached ancestors, after all selections.
  for (const element of [...chosenBranches.keys()].reverse()) {
    const parent = element.parentNode;
    if (!parent) continue;
    // A selection marker is tracked independently of namespace prefix spelling.
    const chosen = chosenBranches.get(element);
    if (chosen) {
      while (chosen.firstChild) parent.insertBefore(chosen.firstChild, element);
    }
    parent.removeChild(element);
  }
  return { document: clone, selections, issues };
}
