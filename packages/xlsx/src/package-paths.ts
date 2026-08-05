/**
 * OOXML package part paths are always POSIX-style, regardless of the host
 * platform. These helpers replace node:path.posix so every module that walks a
 * workbook package stays usable in a browser.
 */

/** The part path without its final segment; "" for a top-level part. */
export function packagePartDirectory(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  return separator < 0 ? "" : partPath.slice(0, separator);
}

/** The final segment of a part path. */
export function packagePartName(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  return separator < 0 ? partPath : partPath.slice(separator + 1);
}

/**
 * Resolve "." and ".." segments. A relative path that climbs above its root
 * keeps its leading ".." segments so callers can reject it.
 */
export function normalizePackagePath(partPath: string): string {
  const absolute = partPath.startsWith("/");
  const segments: string[] = [];

  for (const segment of partPath.split("/")) {
    if (segment === "" || segment === ".") {
      continue;
    }
    if (segment === "..") {
      const previous = segments[segments.length - 1];
      if (previous !== undefined && previous !== "..") {
        segments.pop();
      } else if (!absolute) {
        segments.push("..");
      }
      continue;
    }
    segments.push(segment);
  }

  const joined = segments.join("/");
  if (absolute) {
    return `/${joined}`;
  }
  return joined || ".";
}

/** Join and normalize package path segments, ignoring empty ones. */
export function joinPackagePath(...segments: string[]): string {
  return normalizePackagePath(
    segments.filter((segment) => segment !== "").join("/"),
  );
}
