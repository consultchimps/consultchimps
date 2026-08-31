import { baseOptions, homeLinks } from "@/lib/layout.shared";
import { HomeLayout } from "fumadocs-ui/layouts/home";

/**
 * The reference page sits at the site root rather than inside the (home)
 * route group, so it declares the shared header itself. Same layout, same
 * links, no sub-bar: there is one page here, not a family of them.
 */
export default function Layout({ children }: LayoutProps<"/shortcuts">) {
  return (
    <HomeLayout {...baseOptions()} links={homeLinks}>
      {children}
    </HomeLayout>
  );
}
