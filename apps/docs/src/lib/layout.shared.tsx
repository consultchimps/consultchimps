import { BrandMark } from "@/components/brand-mark";
import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";

import { gitConfig } from "./shared";

export const homeLinks: NonNullable<BaseLayoutProps["links"]> = [
  {
    text: "Tools",
    url: "/docs/tools/spreadsheets",
  },
  {
    text: "Libraries",
    url: "/docs/libraries",
  },
];

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: <BrandMark />,
      transparentMode: "top",
    },
    githubUrl: `https://github.com/${gitConfig.user}/${gitConfig.repo}`,
  };
}
