// Shared test bootstrap: portable Chromium launch + endpoint config.
//
// Chromium resolution order:
//   1. CHROMIUM_PATH env var
//   2. /opt/pw-browsers/chromium (the Claude Code sandbox's pre-installed build)
//   3. channel: 'chrome' (a locally installed Chrome)
// Proxy: honored automatically when HTTPS_PROXY is set (sandbox egress).

import { chromium } from 'playwright-core';
import { existsSync } from 'fs';

export const APP_URL = process.env.DCSPAD_URL || 'http://localhost:8642/index.html';
export const FIXTURES_URL = process.env.DCSPAD_FIXTURES || 'http://localhost:8643';

export async function launchBrowser() {
  const candidate = process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium';
  const usePath = existsSync(candidate);
  return chromium.launch({
    executablePath: usePath ? candidate : undefined,
    channel: usePath ? undefined : 'chrome',
    proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: 'localhost' } : undefined,
    args: ['--ignore-certificate-errors'],
  });
}

let failures = 0;
export async function check(name, fn) {
  try {
    const ok = typeof fn === 'function' ? await fn() : fn;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
    if (!ok) failures++;
    return ok;
  } catch (e) {
    console.log(`FAIL  ${name} — ${e.message.split('\n')[0]}`);
    failures++;
    return false;
  }
}

export function exitWithResult() {
  process.exit(failures ? 1 : 0);
}
