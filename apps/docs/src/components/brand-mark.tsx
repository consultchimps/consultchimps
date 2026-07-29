import { basePath } from "@/lib/shared";
import Image from "next/image";

interface BrandMarkProps {
  compact?: boolean;
}

export function BrandMark({ compact = false }: BrandMarkProps) {
  return (
    <span className="brand-mark">
      <span className="brand-mark__image">
        <Image
          src={`${basePath}/favicon.png`}
          alt=""
          width={34}
          height={34}
          priority
        />
      </span>
      {!compact && (
        <span className="brand-mark__word">
          Consult<span>Chimps</span>
        </span>
      )}
    </span>
  );
}
