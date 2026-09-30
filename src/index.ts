import type { ExtensionModule } from '@kiagent/connector-sdk';
import { createMs365Sender } from './sender';
import { createMs365Source } from './source';

const mod = {
  async activate(host) {
    return {
      sources: [createMs365Source(host)],
      senders: { ms365: createMs365Sender(host) },
    };
  },
} satisfies ExtensionModule<'net' | 'send'>;

export default mod;
module.exports = mod; // dual export — the host child require()s CJS
