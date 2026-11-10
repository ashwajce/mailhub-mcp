import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export interface AppConfig {
  google?: { clientId?: string; clientSecret?: string };
  microsoft?: { clientId?: string; tenant?: string };
  downloadDir?: string;
}

/** Data directory. Holds config.json, accounts.json and downloads. */
export function homeDir(): string {
  return process.env.MAILHUB_HOME || path.join(os.homedir(), ".mailhub-mcp");
}

export function configPath(): string {
  return path.join(homeDir(), "config.json");
}

export function accountsPath(): string {
  return path.join(homeDir(), "accounts.json");
}

export function ensureHome(): void {
  fs.mkdirSync(homeDir(), { recursive: true, mode: 0o700 });
}

export function readJson<T>(file: string, fallback: T): T {
  let raw: string;
  try {
    raw = fs.readFileSync(file, "utf8");
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") return fallback;
    throw err;
  }
  try {
    return JSON.parse(raw) as T;
  } catch {
    throw new Error(`${file} is not valid JSON. Fix or delete it and try again.`);
  }
}

/** Write atomically with owner-only permissions: these files hold credentials. */
export function writeJsonPrivate(file: string, value: unknown): void {
  ensureHome();
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2) + "\n", { mode: 0o600 });
  fs.renameSync(tmp, file);
}

/** config.json merged with environment overrides (env wins). */
export function loadConfig(): AppConfig {
  const file = readJson<AppConfig>(configPath(), {});
  const env = process.env;
  return {
    google: {
      clientId: env.MAILHUB_GOOGLE_CLIENT_ID || file.google?.clientId,
      clientSecret: env.MAILHUB_GOOGLE_CLIENT_SECRET || file.google?.clientSecret,
    },
    microsoft: {
      clientId: env.MAILHUB_MS_CLIENT_ID || file.microsoft?.clientId,
      tenant: env.MAILHUB_MS_TENANT || file.microsoft?.tenant || "common",
    },
    downloadDir: env.MAILHUB_DOWNLOAD_DIR || file.downloadDir || path.join(homeDir(), "downloads"),
  };
}

export function saveConfigPatch(patch: AppConfig): void {
  const current = readJson<AppConfig>(configPath(), {});
  const next: AppConfig = {
    ...current,
    ...patch,
    google: { ...current.google, ...patch.google },
    microsoft: { ...current.microsoft, ...patch.microsoft },
  };
  writeJsonPrivate(configPath(), next);
}
