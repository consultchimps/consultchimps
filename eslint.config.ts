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
);
