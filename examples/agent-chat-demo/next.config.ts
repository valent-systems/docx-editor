import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  transpilePackages: ['@sqren/docx-editor-react', '@sqren/docx-editor-core'],
};

export default nextConfig;
