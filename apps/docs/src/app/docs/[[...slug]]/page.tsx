import { getMDXComponents } from "@/components/mdx";
import { gitConfig } from "@/lib/shared";
import { getPageImageUrl, getPageMarkdownUrl, source } from "@/lib/source";
import { findToolByDocUrl } from "@/lib/tools";
import { createRelativeLink } from "fumadocs-ui/mdx";
import { ArrowRight } from "lucide-react";
import Link from "next/link";
import {
  DocsBody,
  DocsDescription,
  DocsPage,
  DocsTitle,
  MarkdownCopyButton,
  ViewOptionsPopover,
} from "fumadocs-ui/layouts/docs/page";
import type { Metadata } from "next";
import { notFound } from "next/navigation";

export default async function Page(props: PageProps<"/docs/[[...slug]]">) {
  const params = await props.params;
  const page = source.getPage(params.slug);
  if (!page) notFound();

  const MDX = page.data.body;
  const markdownUrl = getPageMarkdownUrl(page).url;
  const browserTool = findToolByDocUrl(page.url);

  return (
    <DocsPage
      toc={page.data.toc}
      full={page.data.full}
      tableOfContent={{ style: "clerk" }}
    >
      <DocsTitle>{page.data.title}</DocsTitle>
      <DocsDescription className="mb-0">
        {page.data.description}
      </DocsDescription>
      <div className="flex flex-row items-center gap-2 border-b pb-6">
        {browserTool?.browserHref ? (
          <Link
            className="inline-flex items-center gap-1.5 rounded-lg bg-fd-primary px-3 py-1.5 text-sm font-semibold text-fd-primary-foreground no-underline transition-opacity hover:opacity-90"
            href={browserTool.browserHref}
          >
            Try it online
            <ArrowRight className="size-3.5" aria-hidden="true" />
          </Link>
        ) : null}
        <MarkdownCopyButton markdownUrl={markdownUrl} />
        <ViewOptionsPopover
          markdownUrl={markdownUrl}
          githubUrl={`https://github.com/${gitConfig.user}/${gitConfig.repo}/blob/${gitConfig.branch}/apps/docs/content/docs/${page.path}`}
        />
      </div>
      <DocsBody>
        <MDX
          components={getMDXComponents({
            a: createRelativeLink(source, page),
          })}
        />
      </DocsBody>
    </DocsPage>
  );
}

export function generateStaticParams() {
  return source.generateParams();
}

export async function generateMetadata(
  props: PageProps<"/docs/[[...slug]]">,
): Promise<Metadata> {
  const params = await props.params;
  const page = source.getPage(params.slug);
  if (!page) notFound();

  return {
    title: page.data.title,
    description: page.data.description,
    openGraph: {
      images: getPageImageUrl(page).url,
    },
  };
}
