/** Test-only public stdio client. No OfficeMCP implementation or app imports. */
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { once } from 'node:events';

export function stdioClient({ cli, config, cwd, version = '2026-07-28', env = process.env, timeoutMs = 120000 }) {
  const child = spawn(process.execPath, [cli, 'serve', '--config', config], {
    cwd, env, stdio: ['pipe', 'pipe', 'pipe'],
  });
  const lines = createInterface({ input: child.stdout }), pending = new Map();
  let serial = 0, diagnostic = '';
  child.stderr.on('data', chunk => { diagnostic = (diagnostic + chunk).slice(-8192); });
  function rejectAll(error) {
    for (const request of pending.values()) { clearTimeout(request.timer); request.reject(error); }
    pending.clear();
  }
  child.on('error', rejectAll);
  child.on('exit', () => rejectAll(new Error('MCP exited: ' + diagnostic)));
  lines.on('line', line => {
    try {
      const message = JSON.parse(line), request = pending.get(message.id);
      if (!request) return;
      pending.delete(message.id); clearTimeout(request.timer);
      message.error ? request.reject(new Error(JSON.stringify(message.error))) : request.resolve(message.result);
    } catch (error) { rejectAll(error); }
  });
  function request(method, params = {}) {
    if (child.exitCode !== null) return Promise.reject(new Error('MCP already exited'));
    return new Promise((resolve, reject) => {
      const id = ++serial, timer = setTimeout(() => {
        pending.delete(id); reject(new Error('MCP timeout: ' + method));
      }, timeoutMs);
      pending.set(id, { resolve, reject, timer });
      child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params: {
        ...params, ...(version === '2026-07-28' ? { _meta: {
          'io.modelcontextprotocol/protocolVersion': version,
          'io.modelcontextprotocol/clientInfo': { name: 'independent-product-check', version: '1' },
          'io.modelcontextprotocol/clientCapabilities': {},
        } } : {}),
      } }) + '\n');
    });
  }
  return {
    async connect() {
      if (version === '2026-07-28') await request('server/discover');
      else {
        await request('initialize', { protocolVersion: version, capabilities: {}, clientInfo: { name: 'independent-product-check', version: '1' } });
        child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
      }
      return (await request('tools/list')).tools;
    },
    request,
    callTool: (name, args) => request('tools/call', { name, arguments: args }),
    async close() {
      lines.close();
      if (child.exitCode === null && child.signalCode === null) {
        const exited = once(child, 'exit');
        child.kill('SIGTERM');
        const timer = setTimeout(() => child.kill('SIGKILL'), 5000);
        try { await exited; } finally { clearTimeout(timer); }
      }
      rejectAll(new Error('MCP connection closed'));
    },
  };
}

export function toolPayload(result) {
  return result.structuredContent ?? JSON.parse(result.content.find(item => item.type === 'text').text);
}
