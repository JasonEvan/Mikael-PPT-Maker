import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  serverExternalPackages: ["skia-canvas"],
  outputFileTracingIncludes: {
    "/api/generate-pptx": [
      "./app/template/*.pptx",
      "./app/fonts/english-mass/**/*",
    ],
  },
};

export default nextConfig;
