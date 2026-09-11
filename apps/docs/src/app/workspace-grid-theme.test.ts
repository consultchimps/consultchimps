import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { describe, expect, it } from "vitest";

/**
 * The record grid's theme layer, checked against the stylesheet it overrides.
 *
 * Three promises are kept here, and the fourth (contrast, and the mode switch
 * in a real browser) is kept in `e2e/workspace-grid-theme.spec.ts`, because
 * only a browser can report what a `color-mix` over a token actually paints.
 *
 * - No colour literal appears in the layer, so every colour the grid shows is a
 *   site token and a later palette overrides a role rather than a selector.
 * - Every colour Tabulator declares for a part this grid can render is bound,
 *   and every colour it declares for a part this grid cannot render is listed
 *   below with the option that would be needed to reach it. A Tabulator upgrade
 *   that adds or moves a colour fails this test until somebody classifies it.
 * - Every selector in the layer repeats its first class, so the layer outranks
 *   Tabulator's own rule by exactly one class and cannot be undone by a change
 *   in which stylesheet a bundler emits first.
 *
 * The stylesheet parsed here is the minified one the grid component imports,
 * not its unminified twin, so what is checked is what ships.
 */

const require_ = createRequire(import.meta.url);

const LAYER_PATH = path.join(import.meta.dirname, "workspace-grid.css");
const TABULATOR_PATH = require_.resolve(
  "tabulator-tables/dist/css/tabulator.min.css",
);

interface StyleRule {
  /** The at-rule prelude chain around the rule, whitespace removed, or "". */
  context: string;
  selector: string;
  property: string;
  value: string;
}

/**
 * Normalize a selector so the layer's Prettier-wrapped, quoted form and
 * Tabulator's minified form compare equal: one space between compounds, no
 * quotes in attribute selectors, no padding inside `:not(...)`, and the single
 * colon a minifier leaves on a pseudo-element.
 */
function normalizeSelector(selector: string): string {
  return selector
    .replace(/\s+/g, " ")
    .replace(/["']/g, "")
    .replace(/\s*([(),])\s*/g, "$1")
    .replace(/::/g, ":")
    .trim();
}

/**
 * A small CSS reader. It is deliberately not a full parser: the two files it
 * reads are plain rule sets inside at-rules, with no nesting beyond that, and a
 * dependency for this would be a dependency in the shipped app's tree.
 */
function readRules(css: string): StyleRule[] {
  const source = css.replace(/\/\*[\s\S]*?\*\//g, "");
  const rules: StyleRule[] = [];
  const open: Array<{ prelude: string; isAtRule: boolean }> = [];
  let buffer = "";

  for (const character of source) {
    if (character === "{") {
      const prelude = buffer.replace(/\s+/g, " ").trim();
      open.push({ prelude, isAtRule: prelude.startsWith("@") });
      buffer = "";
      continue;
    }
    if (character === "}") {
      const block = open.pop();
      if (block !== undefined && !block.isAtRule) {
        const context = open
          .filter((entry) => entry.isAtRule)
          .map((entry) => entry.prelude.replace(/\s+/g, ""))
          .join("&&");
        const selectors = block.prelude.split(",").map(normalizeSelector);
        for (const declaration of buffer.split(";")) {
          const colon = declaration.indexOf(":");
          if (colon < 0) {
            continue;
          }
          const property = declaration.slice(0, colon).trim();
          const value = declaration.slice(colon + 1).trim();
          for (const selector of selectors) {
            rules.push({ context, selector, property, value });
          }
        }
      }
      buffer = "";
      continue;
    }
    buffer += character;
  }

  return rules;
}

const HEX = /#[0-9a-f]{3,8}\b/i;
const COLOR_FUNCTION = /\b(?:rgba?|hsla?|hwb|lab|lch|oklab|oklch|color)\(/i;
/**
 * The CSS named colours Tabulator uses, plus the ones most likely to be reached
 * for by hand. `transparent`, `currentColor`, and `inherit` are not colours in
 * this sense: they carry no value of their own, so the layer may use them.
 */
const NAMED_COLOR =
  /\b(?:aqua|black|blue|brown|coral|crimson|cyan|fuchsia|gold|gray|grey|green|indigo|ivory|khaki|lime|magenta|maroon|navy|olive|orange|orchid|pink|plum|purple|red|salmon|silver|tan|teal|tomato|turquoise|violet|wheat|white|yellow)\b/i;

function hasColorLiteral(value: string): boolean {
  return (
    HEX.test(value) || COLOR_FUNCTION.test(value) || NAMED_COLOR.test(value)
  );
}

/**
 * Parts of Tabulator this grid's configuration cannot render, and the option
 * each one waits on. Nothing here is bound, because binding a colour nobody can
 * see is a claim the grid does not make; the reasons are the record of why.
 */
const UNREACHABLE: ReadonlyArray<readonly [string, string]> = [
  [
    "tabulator-footer",
    "the grid renders no footer: there is no pagination, no page-size control, and no spreadsheet tabs",
  ],
  ["tabulator-calcs", "no column calculations are configured"],
  [
    "tabulator-alert",
    "the loading overlay is created and removed inside one microtask when rows are handed over directly, so it never paints",
  ],
  ["tabulator-menu", "no header menu and no cell menu are configured"],
  ["tabulator-toggle", "the toggle formatter is not used by any column"],
  ["tabulator-print-table", "nothing prints the grid"],
  ["tabulator-col-group", "columns are not grouped under a shared header"],
  ["tabulator-group", "rows are not grouped"],
  ["tabulator-data-tree", "rows are flat records, not a tree"],
  [
    "tabulator-responsive-collapse",
    "the responsive-collapse layout is not used",
  ],
  ["tabulator-row-handle", "rows carry no drag handle"],
  ["tabulator-row-moving", "rows are not movable"],
  ["tabulator-moving", "columns are not movable"],
  ["tabulator-title-editor", "column titles are not editable"],
  ["tabulator-frozen", "no column is frozen"],
  ["tabulator-row-resize-guide", "row heights are not resizable"],
];

/**
 * What the layer must declare to have bound one of Tabulator's colours. A side
 * is covered by that side's longhand or by `border-color`; a shorthand
 * Tabulator writes whole is covered by the longhand that carries its colour.
 */
function bindingsFor(property: string): string[] {
  switch (property) {
    case "color":
      return ["color"];
    case "background":
    case "background-color":
      return ["background-color"];
    case "border":
    case "border-color":
      return ["border-color"];
    case "border-top":
    case "border-right":
    case "border-bottom":
    case "border-left":
      return [`${property}-color`, "border-color"];
    case "outline":
    case "outline-color":
      return ["outline-color"];
    case "box-shadow":
      return ["box-shadow"];
    case "fill":
      return ["fill"];
    case "stroke":
      return ["stroke"];
    default:
      return [property];
  }
}

/** The layer's own selector, with the repeated first class taken back off. */
function undouble(selector: string): string {
  return selector.replace(/^\.([A-Za-z0-9_-]+)\.\1(?![A-Za-z0-9_-])/, ".$1");
}

function keyOf(context: string, selector: string): string {
  return `${context}|${selector}`;
}

const layerCss = readFileSync(LAYER_PATH, "utf8");
const layerRules = readRules(layerCss);
const tabulatorRules = readRules(readFileSync(TABULATOR_PATH, "utf8"));

/** Tabulator's colour roles: one entry per declaration that carries a colour. */
const tabulatorColorRules = tabulatorRules.filter((rule) =>
  hasColorLiteral(rule.value),
);

const layerSelectors = new Set(
  layerRules
    .filter((rule) => rule.selector !== ":root")
    .map((rule) => keyOf(rule.context, undouble(rule.selector))),
);

const layerProperties = new Map<string, Set<string>>();
for (const rule of layerRules) {
  if (rule.selector === ":root") {
    continue;
  }
  const key = keyOf(rule.context, undouble(rule.selector));
  const properties = layerProperties.get(key) ?? new Set<string>();
  properties.add(rule.property);
  layerProperties.set(key, properties);
}

function unreachableReason(selector: string): string | undefined {
  return UNREACHABLE.find(([pattern]) => selector.includes(pattern))?.[1];
}

describe("the record grid theme layer", () => {
  it("declares no colour of its own", () => {
    const literals = layerRules
      .filter((rule) => hasColorLiteral(rule.value))
      .map((rule) => `${rule.selector} { ${rule.property}: ${rule.value} }`);

    expect(literals).toEqual([]);
  });

  it("spends every declaration on a site token", () => {
    // A rule that set a colour without a var() would be a colour by another
    // name: a keyword, or a value smuggled through a non-colour property.
    const colorProperties = /(?:^|-)(?:color|background|fill|stroke|shadow)$/;
    const offenders = layerRules
      .filter(
        (rule) =>
          rule.selector !== ":root" && colorProperties.test(rule.property),
      )
      .filter(
        (rule) =>
          !rule.value.includes("var(--workspace-grid-") &&
          // Neither keyword carries a colour of its own: one takes what is
          // already there, the other takes nothing at all.
          rule.value !== "inherit" &&
          rule.value !== "none",
      )
      .map((rule) => `${rule.selector} { ${rule.property}: ${rule.value} }`);

    expect(offenders).toEqual([]);
  });

  it("names every role in terms of the site's own tokens", () => {
    const roles = layerRules.filter((rule) => rule.selector === ":root");
    expect(roles.length).toBeGreaterThan(0);

    const notATokenAlias = roles
      .filter((rule) => rule.property.startsWith("--workspace-grid-"))
      .filter(
        (rule) =>
          // A length or a size is a plain value; a colour must come from a token.
          !/^-?[\d.]+(?:rem|px|em)$/.test(rule.value) &&
          !rule.value.includes("var(--color-fd-"),
      )
      .map((rule) => `${rule.property}: ${rule.value}`);

    expect(notATokenAlias).toEqual([]);
  });

  it("outranks Tabulator by repeating the first class of every selector", () => {
    const flat = layerRules
      .filter((rule) => rule.selector !== ":root")
      .map((rule) => rule.selector);
    const notDoubled = [...new Set(flat)].filter(
      (selector) => undouble(selector) === selector,
    );

    expect(notDoubled).toEqual([]);
  });

  it("overrides only selectors Tabulator actually declares", () => {
    const tabulatorSelectors = new Set(
      tabulatorRules.map((rule) => keyOf(rule.context, rule.selector)),
    );
    const invented = [...layerSelectors].filter(
      (key) => !tabulatorSelectors.has(key),
    );

    expect(invented).toEqual([]);
  });

  it("binds every colour Tabulator applies to a part this grid renders", () => {
    const unbound: string[] = [];

    for (const rule of tabulatorColorRules) {
      if (unreachableReason(rule.selector) !== undefined) {
        continue;
      }
      const key = keyOf(rule.context, rule.selector);
      const bound = layerProperties.get(key);
      const accepted = bindingsFor(rule.property);
      if (
        bound === undefined ||
        !accepted.some((property) => bound.has(property))
      ) {
        unbound.push(
          `${rule.context === "" ? "" : `${rule.context} `}${rule.selector} { ${rule.property}: ${rule.value} }`,
        );
      }
    }

    expect(unbound).toEqual([]);
  });

  it("leaves nothing unclassified, and keeps no dead exemption", () => {
    // Every exemption has to earn its place, so an option this grid later turns
    // on cannot leave a stale reason behind claiming the colour is unreachable.
    const unused = UNREACHABLE.filter(
      ([pattern]) =>
        !tabulatorColorRules.some((rule) => rule.selector.includes(pattern)),
    ).map(([pattern]) => pattern);

    expect(unused).toEqual([]);

    // And an exemption may not cover something the layer binds, which would
    // leave two contradictory claims about the same colour.
    const contradictions = [...layerSelectors].filter((key) => {
      const selector = key.slice(key.indexOf("|") + 1);
      return unreachableReason(selector) !== undefined;
    });

    expect(contradictions).toEqual([]);
  });
});
