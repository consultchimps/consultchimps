import { BROWSER_TOOLS, isBrowserTool, TOOLS } from "@/lib/tools";
import { ArrowRight, BookOpen } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";

export const metadata: Metadata = {
  title: "Online tools",
  description:
    "Run ConsultChimps document tools directly in your browser. Files never leave your machine — everything happens in this tab.",
};

const guideOnlyTools = TOOLS.filter((tool) => !isBrowserTool(tool));

export default function Page() {
  return (
    <main className="flex-1">
      <div className="mx-auto w-full max-w-[1100px] px-6 pb-24 pt-14 lg:px-8">
        <div className="manual-kicker">Online tools · in-browser</div>
        <h1 className="mt-6 text-4xl font-bold tracking-[-0.04em] md:text-5xl">
          Run a tool right here.
        </h1>
        <p className="mt-5 max-w-2xl text-lg leading-8 text-fd-muted-foreground">
          These tools run entirely in your browser tab. Your files are never
          uploaded, and the results match the ConsultChimps command line
          byte-for-byte.
        </p>

        <div className="mt-10 grid gap-5 sm:grid-cols-2 lg:grid-cols-3">
          {BROWSER_TOOLS.map(
            ({ slug, title, description, surfaces, icon: Icon }) => (
              <Link
                className="tool-card"
                href={surfaces.browser.href}
                key={slug}
              >
                <div>
                  <span className="tool-card__icon">
                    <Icon className="size-5" />
                  </span>
                  <h2>{title}</h2>
                  <p>{description}</p>
                  <span className="tool-card__link">
                    Open the tool
                    <ArrowRight className="size-3.5" />
                  </span>
                </div>
              </Link>
            ),
          )}
        </div>

        <h2 className="mt-16 text-2xl font-bold tracking-[-0.03em]">
          Not in the browser yet
        </h2>
        <p className="mt-3 max-w-2xl leading-7 text-fd-muted-foreground">
          The rest of the kit currently runs through the command line or the
          TypeScript libraries. Each guide covers installation and a worked
          example.
        </p>
        <ul className="mt-6 grid gap-3 sm:grid-cols-2">
          {guideOnlyTools.map(({ slug, title, docHref, icon: Icon }) => (
            <li key={slug}>
              <Link
                className="inline-flex items-center gap-2.5 rounded-lg border bg-fd-card px-4 py-3 text-sm font-medium transition-colors hover:bg-fd-accent"
                href={docHref}
              >
                <Icon className="size-4 text-fd-primary" aria-hidden="true" />
                {title}
                <BookOpen
                  className="ml-auto size-4 text-fd-muted-foreground"
                  aria-hidden="true"
                />
              </Link>
            </li>
          ))}
        </ul>
      </div>
    </main>
  );
}
