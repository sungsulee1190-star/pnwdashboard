// data.js → data.enc.js 암호화 (Node 18+, 추가 설치 없음)
// 사용: node tools/encrypt-data.mjs [비밀번호]   (비밀번호를 생략하면 환경변수 PNW_PASSWORD)
// 형식: gzip 압축 → PBKDF2-SHA256(600,000회)로 키 생성 → AES-256-GCM. index.html의 로그인 화면이 같은 방식으로 복호화.
import { readFileSync, writeFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { webcrypto as crypto } from 'node:crypto';

const password = process.argv[2] || process.env.PNW_PASSWORD;
if (!password) { console.error('비밀번호를 인자나 PNW_PASSWORD로 주세요.'); process.exit(1); }

const src = new URL('../data.js', import.meta.url), out = new URL('../data.enc.js', import.meta.url);
const plain = gzipSync(readFileSync(src), { level: 9 });
const iter = 600000;
const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, plain));
const b64 = u => Buffer.from(u).toString('base64');
writeFileSync(out, `// 암호화된 대시보드 데이터 (tools/encrypt-data.mjs로 생성, ${new Date().toISOString().slice(0, 10)})\nwindow.DATA_ENC = ${JSON.stringify({ v: 1, iter, salt: b64(salt), iv: b64(iv), ct: b64(ct) })};\n`);
console.log(`data.enc.js 생성: 원본 ${(readFileSync(src).length / 1024).toFixed(0)}KB → 암호문 ${(ct.length * 4 / 3 / 1024).toFixed(0)}KB`);
