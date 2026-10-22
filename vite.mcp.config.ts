import { defineConfig } from 'vite';

// Keep the Node-only MCP transport outside the browser component bundle.
export default defineConfig({
  build: {
    outDir: 'lib',
    emptyOutDir: false,
    target: 'node22',
    lib: { entry: 'src/mcp/server.ts', formats: ['es'], fileName: () => 'mcp.js' },
    rollupOptions: { external: id => !id.startsWith('.') && !id.startsWith('/') },
  },
});
