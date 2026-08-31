import { packageReleases } from "@/lib/releases";
import { gitConfig } from "@/lib/shared";
import { ArrowUpRight } from "lucide-react";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Release history",
  description:
    "What shipped in each ConsultChimps package, generated from the changelogs at deploy time.",
};

export default function ReleasesPage() {
  const releases = packageReleases();

  return (
    <main className="flex-1">
      <div className="mx-auto w-full max-w-[900px] px-6 pb-24 pt-14 lg:px-8">
        <div className="manual-kicker">Release history · auto-generated</div>
        <h1 className="mt-6 text-4xl font-bold tracking-[-0.04em] md:text-5xl">
          What shipped, and when.
        </h1>
        <p className="mt-5 max-w-2xl text-lg leading-8 text-fd-muted-foreground">
          This page is generated from the package changelogs on every site
          deploy. Full histories live in each package&apos;s CHANGELOG on
          GitHub.
        </p>

        <div className="mt-12 space-y-12">
          {releases.map(({ name, folder, version, description, entries }) => (
            <section key={name}>
              <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                <h2 className="font-mono text-xl font-bold tracking-tight">
                  {name}
                </h2>
                <span className="rounded-full border bg-fd-card px-2.5 py-0.5 font-mono text-xs font-semibold text-fd-primary">
                  v{version}
                </span>
                <a
                  className="ml-auto inline-flex items-center gap-1 text-sm text-fd-muted-foreground hover:text-fd-foreground"
                  href={`https://github.com/${gitConfig.user}/${gitConfig.repo}/blob/${gitConfig.branch}/packages/${encodeURIComponent(folder)}/CHANGELOG.md`}
                  rel="noreferrer"
                  target="_blank"
                >
                  Full changelog
                  <ArrowUpRight className="size-3.5" aria-hidden="true" />
                </a>
              </div>
              {description ? (
                <p className="mt-2 text-sm leading-6 text-fd-muted-foreground">
                  {description}
                </p>
              ) : null}
              <div className="mt-4 space-y-4 border-l-2 border-fd-border pl-5">
                {entries.length === 0 ? (
                  <p className="text-sm text-fd-muted-foreground">
                    No released changes recorded yet.
                  </p>
                ) : (
                  entries.map(({ version: entryVersion, notes }) => (
                    <div key={entryVersion}>
                      <h3 className="font-mono text-sm font-semibold">
                        {entryVersion}
                      </h3>
                      {notes.length > 0 ? (
                        <ul className="mt-1.5 list-disc space-y-1 pl-5 text-sm leading-6 text-fd-muted-foreground">
                          {notes.map((note) => (
                            <li key={note}>{note}</li>
                          ))}
                        </ul>
                      ) : (
                        <p className="mt-1.5 text-sm text-fd-muted-foreground">
                          Dependency updates only.
                        </p>
                      )}
                    </div>
                  ))
                )}
              </div>
            </section>
          ))}
        </div>
      </div>
    </main>
  );
}
