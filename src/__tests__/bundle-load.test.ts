/**
 * Smoke test for the bundled dist/index.js — proves the CJS/ESM dual export
 * in src/index.ts (`export default mod; module.exports = mod;`) against
 * actual esbuild output. The build + require + activate() plumbing is the
 * SDK kit's `bundleLoadSmoke`, shared by every connector repo.
 */
import { join } from 'node:path';
import { bundleLoadSmoke } from '@kiagent/connector-sdk/testing';

describe('dist bundle loads standalone', () => {
  it('require()s dist/index.js and activate() returns the ms365 source', async () => {
    await bundleLoadSmoke({
      root: join(__dirname, '..', '..'),
      selfId: 'kia.ms365',
      sourceIds: ['ms365'],
    });
  }, 30_000);
});
