import { ShortcutFinder } from "@/components/shortcut-finder";
import { EXCEL_SHORTCUTS } from "@/lib/excel-shortcuts-data";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Excel shortcuts",
  description:
    "A searchable database of Excel keyboard shortcuts for Windows. Press the keys you remember and the list narrows to the shortcuts that start with them.",
};

export default function ShortcutsPage() {
  return (
    <main className="flex-1">
      <div className="mx-auto w-full max-w-[1100px] px-6 pb-24 pt-14 lg:px-8">
        <div className="manual-kicker">
          Excel shortcuts · {EXCEL_SHORTCUTS.length} for Windows
        </div>
        <h1 className="mt-6 text-4xl font-bold tracking-[-0.04em] md:text-5xl">
          Press what you remember
        </h1>
        <p className="mt-5 max-w-2xl text-lg leading-8 text-fd-muted-foreground">
          Most shortcut lists ask you to already know the answer. This one
          searches by the keys themselves: hold Ctrl and Shift and the list
          narrows to the shortcuts that start that way, then add the letter you
          think it was. Ribbon routes such as Alt, H, O, I work the same way,
          one step at a time
        </p>
        <p className="mt-4 max-w-2xl leading-7 text-fd-muted-foreground">
          The keys are the ones on a Windows keyboard, and each entry names the
          physical key rather than the character a shifted key types, so the
          combination published as Ctrl+Shift+* appears here as Ctrl+Shift+8.
          Everything runs in this tab, and nothing you press is recorded
        </p>

        <div className="mt-10">
          <ShortcutFinder />
        </div>
      </div>
    </main>
  );
}
