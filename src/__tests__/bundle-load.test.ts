/**
 * Smoke test for the bundled dist/index.js — proves the CJS/ESM dual export
 * in src/index.ts (`export default mod; module.exports = mod;`) against
 * actual esbuild output. The build + require + activate() plumbing is the
 * SDK kit's `bundleLoadSmoke`, shared by every connector repo.
 *
 * The second case carries what the shared kit does not express, restored
 * from this file at 2bdad48 (pre-SDK): a BARE-NODE child-process require —
 * the kit's own `require` runs inside jest's registry, so only a fresh
 * `node -e` proves the bundle loads the way the extension-host child loads
 * it — and the `descriptor.auth` field, which the kit has no option for.
 * Both would otherwise have been lost to the migration.
 */
import { execSync } from 'node:child_process';
import { join } from 'node:path';
import { bundleLoadSmoke } from '@kiagent/connector-sdk/testing';
import type { HostFor } from '@kiagent/connector-sdk';

describe('dist bundle loads standalone', () => {
  it('require()s dist/index.js and activate() returns the ms365 source', async () => {
    await bundleLoadSmoke({
      root: join(__dirname, '..', '..'),
      selfId: 'kia.ms365',
      sourceIds: ['ms365'],
    });
  }, 30_000);

  it('loads in a bare Node process and contributes an oauth source', async () => {
    const root = join(__dirname, '..', '..');
    // Kept from the original: this case builds for itself, so it still holds
    // when run alone (`-t`), without leaning on the smoke above.
    execSync('npm run build', { cwd: root });

    // Bare Node child-process require — proves the bundle loads outside
    // jest's module registry, the way the extension host child loads it.
    const out = execSync(
      'node -e "const m=require(\'./dist/index.js\');const e=m.default??m;' +
        "if(typeof e.activate!=='function')throw new Error('no activate');" +
        'console.log(\'activate:\'+typeof e.activate)"',
      { cwd: root },
    ).toString();
    expect(out).toContain('activate:function');

    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const mod = require(join(root, 'dist', 'index.js'));
    const entry = mod.default ?? mod;
    expect(typeof entry.activate).toBe('function');

    const unused = () => {
      throw new Error('unused in this smoke test');
    };
    const host: HostFor<'net'> = {
      self: { id: 'kia.ms365', dataDir: '/tmp' },
      log: () => {},
      net: { fetch: unused },
    };
    const result = await entry.activate(host);

    // The kit's `sourceIds` covers length + id; `auth` is the field it has no
    // option for, so it is asserted here.
    expect(result.sources?.[0]?.descriptor.auth).toBe('oauth');
  }, 30_000);
});
