import { FEATURES, isFeature, type Feature } from "./features";

// The slice of an MDX syntax tree this plugin reads; mdast types are not a
// direct dependency of the docs app.
interface TreeNode {
  type: string;
  name?: string | null;
  attributes?: readonly { type: string; name?: string; value?: unknown }[];
  children?: TreeNode[];
}

/**
 * Removes `<Feature name="…">` blocks whose feature is disabled and unwraps the
 * rest, before the page renders or its processed markdown feeds llms.txt. An
 * unknown or non-literal name fails the build rather than guessing.
 */
export function remarkFeatureFlags() {
  return (tree: TreeNode, file: { path?: string }) => {
    applyFlags(tree, file.path ?? "an MDX file");
  };
}

function applyFlags(node: TreeNode, location: string): void {
  if (!node.children) return;
  node.children = node.children.flatMap((child) => {
    applyFlags(child, location);
    if (!isFeatureElement(child)) return [child];
    return FEATURES[featureOf(child, location)] ? (child.children ?? []) : [];
  });
}

function isFeatureElement(node: TreeNode): boolean {
  return (
    (node.type === "mdxJsxFlowElement" || node.type === "mdxJsxTextElement") &&
    node.name === "Feature"
  );
}

function featureOf(node: TreeNode, location: string): Feature {
  const attribute = node.attributes?.find(
    (candidate) =>
      candidate.type === "mdxJsxAttribute" && candidate.name === "name",
  );
  const name = attribute?.value;
  if (typeof name !== "string" || !isFeature(name)) {
    throw new Error(
      `${location}: <Feature> needs name set to one of ${Object.keys(FEATURES).join(", ")} as a string literal`,
    );
  }
  return name;
}
