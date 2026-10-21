import type { ExtensionModule } from '@kiagent/connector-sdk';
import { createMs365CalendarWriter } from './calendar/writer';
import { createMs365Sender } from './sender';
import { createMs365Source } from './source';

const mod = {
  async activate(host) {
    return {
      sources: [createMs365Source(host)],
      senders: { ms365: createMs365Sender(host) },
      calendarWriters: { ms365: createMs365CalendarWriter(host) },
    };
  },
} satisfies ExtensionModule<'net' | 'send' | 'calendar.write'>;

export default mod;
module.exports = mod; // dual export — the host child require()s CJS
