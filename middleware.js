// Bloqueo real de acceso: corre en el edge de Vercel antes de servir
// cualquier página. Si no hay una sesión válida (cookie kw_session, firmada
// por api/login.py), redirige a /login.html sin llegar a entregar index.html
// ni ninguno de los datos financieros embebidos en él.
import { next } from '@vercel/functions';

function base64urlToBytes(b64url) {
  const b64 = b64url.replace(/-/g, '+').replace(/_/g, '/');
  const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function getCookie(request, name) {
  const header = request.headers.get('cookie') || '';
  for (const part of header.split(';')) {
    const [k, ...rest] = part.trim().split('=');
    if (k === name) return rest.join('=');
  }
  return null;
}

async function verifySession(token, secret) {
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [headerB64, payloadB64, sigB64] = parts;

  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['verify']
  );
  const signingInput = new TextEncoder().encode(`${headerB64}.${payloadB64}`);
  const signature = base64urlToBytes(sigB64);
  const valid = await crypto.subtle.verify('HMAC', key, signature, signingInput);
  if (!valid) return null;

  const payload = JSON.parse(new TextDecoder().decode(base64urlToBytes(payloadB64)));
  if (!payload.exp || payload.exp < Math.floor(Date.now() / 1000)) return null;
  return payload;
}

export default async function middleware(request) {
  const secret = process.env.AUTH_JWT_SECRET;
  const token = getCookie(request, 'kw_session');
  const session = token && secret ? await verifySession(token, secret) : null;

  if (!session) {
    const url = new URL('/login.html', request.url);
    return new Response(null, { status: 302, headers: { Location: url.toString() } });
  }

  return next(); // sesión válida: continuar normal hacia index.html
}

export const config = {
  matcher: ['/((?!api|login\\.html|favicon\\.ico|CNAME).*)'],
};
