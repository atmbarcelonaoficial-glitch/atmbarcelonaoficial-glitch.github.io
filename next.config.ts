import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Emma runs entirely in the browser: GTFS parsing, editing and PDF output do
  // not require a Node server. A static export can therefore be hosted on the
  // free GitHub Pages tier.
  output: "export",
  trailingSlash: true,
};

export default nextConfig;
