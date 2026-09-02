"use client";

import { BROWSER_TOOL_GROUPS } from "@/lib/tools";
import { LayoutGrid } from "lucide-react";
import Link from "next/link";
import { usePathname } from "next/navigation";

const baseTabClass =
  "inline-flex shrink-0 items-center gap-1.5 border-b-2 px-3 py-2.5 text-sm font-medium transition-colors";

/**
 * Sub-bar shown under the header on every /tools page. One tab per
 * in-browser tool, driven by the shared registry, so switching tools never
 * requires hunting for an inline "try the other tool" link.
 */
export function ToolsTabs() {
  const pathname = usePathname();

  return (
    <nav aria-label="Online tools" className="border-b bg-fd-card/60">
      <div className="mx-auto flex w-full max-w-[1320px] items-center gap-1 overflow-x-auto px-4 lg:px-8">
        <Link
          className={
            baseTabClass +
            (pathname === "/tools"
              ? " border-fd-primary text-fd-foreground"
              : " border-transparent text-fd-muted-foreground hover:text-fd-foreground")
          }
          href="/tools"
        >
          <LayoutGrid className="size-4" aria-hidden="true" />
          All tools
        </Link>
        {BROWSER_TOOL_GROUPS.map(({ category, tools }) => (
          <div
            aria-label={`${category} tools`}
            className="flex shrink-0 items-center gap-1 border-l pl-2 first:border-l-0 first:pl-0"
            key={category}
            role="group"
          >
            <span className="px-2 text-[0.65rem] font-semibold uppercase tracking-[0.14em] text-fd-muted-foreground">
              {category}
            </span>
            {tools.map(({ slug, tabLabel, surfaces, icon: Icon }) => (
              <Link
                className={
                  baseTabClass +
                  (pathname === surfaces.browser.href
                    ? " border-fd-primary text-fd-foreground"
                    : " border-transparent text-fd-muted-foreground hover:text-fd-foreground")
                }
                href={surfaces.browser.href}
                key={slug}
              >
                <Icon className="size-4" aria-hidden="true" />
                {tabLabel}
              </Link>
            ))}
          </div>
        ))}
      </div>
    </nav>
  );
}
