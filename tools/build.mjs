// 사용법: node tools/build.mjs <비밀번호>
// places.js(로컬 원본, git에 올리지 않음)를 암호화해 places.enc.js(게시용)를 만든다.
// 방식: JSON → gzip 압축 → AES-256-GCM 암호화(PBKDF2, 아래 ITER 회) → base64
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { gzipSync } from "node:zlib";
import { randomBytes, pbkdf2Sync, createCipheriv } from "node:crypto";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ITER = 200000;
const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const pw = process.argv[2];
if (!pw) { console.error("사용법: node tools/build.mjs <비밀번호>"); process.exit(1); }
if (!existsSync(join(root, "places.js"))) {
  console.error("places.js 로컬 원본이 없습니다. 이 파일은 git에 올라가지 않으므로 백업본에서 복원하거나, 관리자 기기에서 실행해야 합니다.");
  process.exit(1);
}

const src = readFileSync(join(root, "places.js"), "utf8");
const PLACES = new Function(src + "\n;return PLACES;")();
const json = JSON.stringify(PLACES);
const gz = gzipSync(Buffer.from(json));

const salt = randomBytes(16);
const iv = randomBytes(12);
const key = pbkdf2Sync(pw, salt, ITER, 32, "sha256");
const cipher = createCipheriv("aes-256-gcm", key, iv);
// WebCrypto의 AES-GCM은 암호문 뒤에 인증태그가 붙은 형태를 기대함
const ct = Buffer.concat([cipher.update(gz), cipher.final(), cipher.getAuthTag()]);

const out = `const PLACES_ENC={it:${ITER},salt:"${salt.toString("base64")}",iv:"${iv.toString("base64")}",data:"${ct.toString("base64")}"};\n`;
writeFileSync(join(root, "places.enc.js"), out);
console.log(`places.enc.js 생성: 장소 ${PLACES.length}곳, JSON ${(json.length/1024).toFixed(1)}KB → 암호문 ${(ct.length/1024).toFixed(1)}KB`);
