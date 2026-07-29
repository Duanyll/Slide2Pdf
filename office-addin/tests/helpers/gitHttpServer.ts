import { createServer, type Server } from "node:http";
import { execFileSync, spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

export function seedBareGitProject(root: string, seed: string): string {
  const remote = path.join(root, "project.git");
  execFileSync("git", ["init", "--bare", "--initial-branch=master", remote]);
  execFileSync("git", ["-C", remote, "config", "http.receivepack", "true"]);
  execFileSync("git", ["init", "--initial-branch=master", seed]);
  execFileSync("git", ["-C", seed, "config", "user.name", "Test Author"]);
  execFileSync("git", ["-C", seed, "config", "user.email", "test@example.com"]);
  fs.writeFileSync(path.join(seed, "main.tex"), "original project\n");
  execFileSync("git", ["-C", seed, "add", "main.tex"]);
  execFileSync("git", ["-C", seed, "commit", "-m", "Initial project"]);
  execFileSync("git", ["-C", seed, "remote", "add", "origin", remote]);
  execFileSync("git", ["-C", seed, "push", "origin", "master"]);
  return remote;
}

export async function startGitHttpServer(
  projectRoot: string,
  options: {
    beforeRequest?: (method: string, url: string) => void;
  } = {},
): Promise<{
  origin: string;
  close: () => Promise<void>;
}> {
  const server = createServer((request, response) => {
    options.beforeRequest?.(request.method ?? "GET", request.url ?? "/");
    const requestUrl = new URL(request.url ?? "/", "http://localhost");
    const backend = spawn("git", ["http-backend"], {
      env: {
        ...process.env,
        CONTENT_LENGTH: request.headers["content-length"] ?? "",
        CONTENT_TYPE: request.headers["content-type"] ?? "",
        GIT_HTTP_EXPORT_ALL: "1",
        GIT_PROJECT_ROOT: projectRoot,
        PATH_INFO: decodeURIComponent(requestUrl.pathname),
        QUERY_STRING: requestUrl.search.slice(1),
        REQUEST_METHOD: request.method ?? "GET",
      },
      stdio: ["pipe", "pipe", "pipe"],
    });

    request.pipe(backend.stdin);

    const headerChunks: Buffer[] = [];
    let headersSent = false;
    backend.stdout.on("data", (chunk: Buffer) => {
      if (headersSent) {
        response.write(chunk);
        return;
      }

      headerChunks.push(chunk);
      const buffered = Buffer.concat(headerChunks);
      const headerEnd = buffered.indexOf("\r\n\r\n");
      if (headerEnd === -1) return;

      const headerText = buffered.subarray(0, headerEnd).toString("utf8");
      for (const line of headerText.split("\r\n")) {
        const separator = line.indexOf(":");
        if (separator === -1) continue;
        const name = line.slice(0, separator);
        const value = line.slice(separator + 1).trim();
        if (name.toLowerCase() === "status") {
          response.statusCode = Number.parseInt(value, 10);
        } else {
          response.setHeader(name, value);
        }
      }

      headersSent = true;
      response.write(buffered.subarray(headerEnd + 4));
    });

    backend.stderr.on("data", () => undefined);
    backend.on("close", (code) => {
      if (!headersSent) {
        response.statusCode = 500;
      }
      response.end(code === 0 ? undefined : "Git HTTP backend failed.");
    });
  });

  await listen(server);
  const address = server.address();
  if (!address || typeof address === "string") {
    throw new Error("Git test server did not expose a TCP port.");
  }

  return {
    origin: `http://127.0.0.1:${address.port}`,
    close: () => close(server),
  };
}

function listen(server: Server): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => resolve());
  });
}

function close(server: Server): Promise<void> {
  return new Promise((resolve) => {
    server.closeAllConnections();
    server.close();
    resolve();
  });
}
