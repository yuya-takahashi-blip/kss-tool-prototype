/** @type {import('next').NextConfig} */
const nextConfig = {
  eslint: {
    ignoreDuringBuilds: true,
  },
  images: { unoptimized: true },
  webpack: (config, { isServer }) => {
    if (!isServer) {
      // pptxgenjs references Node built-ins via node: scheme; stub them for the browser
      config.resolve.fallback = {
        ...config.resolve.fallback,
        fs: false,
        path: false,
        https: false,
        http: false,
        stream: false,
        zlib: false,
        crypto: false,
      };
      // Handle node: URI scheme (e.g. node:fs, node:https)
      config.plugins.push(
        new (require("webpack").NormalModuleReplacementPlugin)(
          /^node:/,
          (resource) => {
            resource.request = resource.request.replace(/^node:/, "");
          }
        )
      );
    }
    return config;
  },
};

module.exports = nextConfig;
