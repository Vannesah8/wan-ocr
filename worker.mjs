// Static pages and the Python API share the dashboard origin and its host-only cookie.
const methods = new Map([
  ...['/session', '/files', '/state', '/report', '/download'].map(path => [path, 'GET']),
  ...['/batch', '/upload', '/run', '/convert', '/compile'].map(path => [path, 'POST']),
]);

function failure(message, status) {
  return Response.json({error: message}, {status, headers: {'Cache-Control': 'no-store'}});
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const method = methods.get(url.pathname);
    if (!method) return env.ASSETS.fetch(request);
    if (method !== request.method) return failure('Unsupported method.', 405);
    if (method === 'POST' && (request.headers.get('Origin') !== url.origin ||
        request.headers.get('X-Factory') !== '1')) {
      return failure('Open the upload form on the dashboard.', 403);
    }
    if (!env.PYTHON_ORIGIN || !env.WAN_PROXY_SECRET) {
      return failure('The hosted workspace needs its Python backend connected. You can run factory.py locally for this batch.', 503);
    }
    try {
      const origin = new URL(env.PYTHON_ORIGIN);
      if (origin.protocol !== 'https:' || origin.origin === url.origin) {
        return failure('The configured Python backend address is invalid.', 503);
      }
      const upstream = new URL(url.pathname + url.search, origin.origin);
      const headers = new Headers({'X-Wan-Proxy': env.WAN_PROXY_SECRET});
      for (const key of ['Content-Type', 'Content-Length', 'X-Factory', 'Origin']) {
        if (request.headers.has(key)) headers.set(key, request.headers.get(key));
      }
      const cookies = (request.headers.get('Cookie') || '').split(';')
        .map(part => part.trim()).filter(part => /^(wan_browser|wan_session)=/.test(part));
      if (cookies.length) headers.set('Cookie', cookies.join('; '));
      const response = await fetch(upstream, {
        method, headers, body: method === 'POST' ? request.body : undefined, redirect: 'manual',
      });
      if (response.status >= 300 && response.status < 400) {
        return failure('The Python backend redirected the request. Check its configured address.', 502);
      }
      return response;
    } catch {
      return failure('The Python backend could not be reached. Check that it is running.', 502);
    }
  },
};
