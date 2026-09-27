import { describe, expect, it } from "vitest";

import { FEATURES, isHiddenPage } from "./features";
import { remarkFeatureFlags } from "./remark-feature-flags";
import { SITE_TOOLS, TOOLS } from "./tools";

function featureBlock(name: unknown, text: string) {
  return {
    type: "mdxJsxFlowElement",
    name: "Feature",
    attributes: [{ type: "mdxJsxAttribute", name: "name", value: name }],
    children: [
      { type: "paragraph", children: [{ type: "text", value: text }] },
    ],
  };
}

function run(children: object[]) {
  const tree = { type: "root", children };
  remarkFeatureFlags()(tree as never, { path: "test.mdx" });
  return JSON.stringify(tree);
}

describe("site feature flags", () => {
  it("keeps the database feature off while it is paused", () => {
    expect(FEATURES.database).toBe(false);
  });

  it("removes a disabled feature's MDX sections, including nested ones", () => {
    const output = run([
      { type: "paragraph", children: [{ type: "text", value: "kept" }] },
      featureBlock("database", "hidden"),
      {
        type: "mdxJsxFlowElement",
        name: "Cards",
        attributes: [],
        children: [featureBlock("database", "nested")],
      },
    ]);
    expect(output).toContain("kept");
    expect(output).not.toContain("hidden");
    expect(output).not.toContain("nested");
    expect(output).not.toContain('"Feature"');
  });

  it("fails on an unknown or missing feature name", () => {
    expect(() => run([featureBlock("databse", "x")])).toThrow(
      /test\.mdx: <Feature> needs name/u,
    );
    expect(() => run([featureBlock(undefined, "x")])).toThrow(/needs name/u);
  });

  it("hides the database guide page and registry entry", () => {
    expect(isHiddenPage("tools/data-workspace.mdx")).toBe(true);
    expect(isHiddenPage("tools\\data-workspace.mdx")).toBe(true);
    expect(isHiddenPage("tools/pdf-split.mdx")).toBe(false);
    expect(TOOLS.some((tool) => tool.slug === "data-workspace")).toBe(true);
    expect(SITE_TOOLS.some((tool) => tool.slug === "data-workspace")).toBe(
      false,
    );
  });
});
