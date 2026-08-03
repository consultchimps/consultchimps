import { baseOptions, homeLinks } from "@/lib/layout.shared";
import { HomeLayout } from "fumadocs-ui/layouts/home";

export default function Layout({ children }: LayoutProps<"/tools">) {
  return (
    <HomeLayout {...baseOptions()} links={homeLinks}>
      {children}
    </HomeLayout>
  );
}
