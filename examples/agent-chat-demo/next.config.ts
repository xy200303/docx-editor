import type { NextConfig } from 'next';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(dirname, '../..');

const nextConfig: NextConfig = {
  transpilePackages: [
    '@eigenpal/docx-editor-react',
    '@eigenpal/docx-editor-core',
    '@eigenpal/docx-editor-i18n',
  ],
  turbopack: {
    root: repoRoot,
  },
};

export default nextConfig;
