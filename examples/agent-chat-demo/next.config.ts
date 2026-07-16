import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  transpilePackages: ['@valent/docx-editor-react', '@valent/docx-editor-core'],
};

export default nextConfig;
