import eslint from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    ignores: [
      "**/.next/**",
      "**/.source/**",
      "**/dist/**",
      "**/dist-bundle/**",
      "**/coverage/**",
      "**/node_modules/**",
      "apps/docs/**",
      // Local agent configuration and scratch space, never committed.
      ".claude/**",
      "tmp/**",
      // Generated Emscripten glue, committed for reproducibility.
      "packages/pbi/wasm/xpress9.mjs",
    ],
  },
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      "@typescript-eslint/consistent-type-imports": "error",
      "@typescript-eslint/no-explicit-any": "error",
      "no-console": ["error", { allow: ["error", "warn"] }],
    },
  },
  {
    // The operation libraries read rows and cells, so their arrays grow with
    // the workbook.
    files: ["packages/{core,files,tabular,xlsx,pptx,pdf}/src/**/*.ts"],
    rules: {
      // A spread passes every element as an argument, and a large array
      // overflows the call stack: a split of 150,000 rows crashed this way.
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "CallExpression[callee.property.name=/^(max|min|push|unshift|fromCharCode|fromCodePoint)$/] > SpreadElement",
          message:
            "Spreading an array into arguments overflows the stack when it is large. Use a loop.",
        },
      ],
    },
  },
);
