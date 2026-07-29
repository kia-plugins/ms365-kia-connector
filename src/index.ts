import type { ExtensionModule } from '@kiagent/connector-sdk';
import { createMs365Source } from './source';

const mod = {
  async activate(host) {
    return { sources: [createMs365Source(host)] };
  },
} satisfies ExtensionModule<'net'>;

export default mod;
module.exports = mod; // dual export — the host child require()s CJS
