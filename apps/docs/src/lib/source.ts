import {
  basePath,
  docsContentRoute,
  docsImageRoute,
  docsRoute,
} from "@/lib/shared";
import { isHiddenPage } from "@/lib/features";
import { docs } from "collections/server";
import { lucideIconsPlugin } from "fumadocs-core/source/lucide-icons";
import { loader } from "fumadocs-core/source";

const docsSource = docs.toFumadocsSource();

// Every docs route (pages, sidebar, search, llms output, OG images) reads this
// loader, so dropping a disabled feature's pages here removes them everywhere.
export const source = loader({
  baseUrl: docsRoute,
  source: {
    files: docsSource.files.filter((file) => !isHiddenPage(file.path)),
  },
  plugins: [lucideIconsPlugin()],
});

export function getPageImageUrl(page: (typeof source)["$inferPage"]) {
  const segments = [...page.slugs, "image.png"];

  return {
    segments,
    url:
      basePath +
      "/" +
      [page.locale, ...docsImageRoute.split("/"), ...segments]
        .filter(Boolean)
        .join("/"),
  };
}

export function getPageMarkdownUrl(page: (typeof source)["$inferPage"]) {
  const segments = [...page.slugs, "content.md"];

  return {
    segments,
    url:
      basePath +
      "/" +
      [page.locale, ...docsContentRoute.split("/"), ...segments]
        .filter(Boolean)
        .join("/"),
  };
}

export async function getLLMText(page: (typeof source)["$inferPage"]) {
  const processed = await page.data.getText("processed");

  return `# ${page.data.title} (${page.url})

${processed}`;
}
