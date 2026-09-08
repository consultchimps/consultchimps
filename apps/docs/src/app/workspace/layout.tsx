import { baseOptions, homeLinks } from "@/lib/layout.shared";
import { HomeLayout } from "fumadocs-ui/layouts/home";

/**
 * The workspace page sits at the site root rather than inside the (home) route
 * group, like the reference page, so it declares the shared header itself. Same
 * layout, same links, no sub-bar: it is one page, not a family of them.
 */
export default function Layout({ children }: LayoutProps<"/workspace">) {
  return (
    <HomeLayout {...baseOptions()} links={homeLinks}>
      {children}
    </HomeLayout>
  );
}
