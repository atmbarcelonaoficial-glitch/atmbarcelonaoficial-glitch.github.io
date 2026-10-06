import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Emma — Generador d’horaris de bus",
  description: "Crea, adapta i exporta horaris de bus modulars.",
  icons: { icon: "/favicon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ca"><body>{children}</body></html>;
}
