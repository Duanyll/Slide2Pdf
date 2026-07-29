import "../polyfills/browserBuffer";

import LightningFS from "@isomorphic-git/lightning-fs";
import http from "isomorphic-git/http/web";

import {
  OverleafGitClient,
  type PushPdfOptions,
  type PushPdfResult,
} from "./overleafGitClient";

const clients = new Map<string, OverleafGitClient>();
const pendingPushes = new Map<string, Promise<unknown>>();

export function pushPdfFromBrowser(
  options: PushPdfOptions,
): Promise<PushPdfResult> {
  const previousPush = pendingPushes.get(options.remoteUrl) ?? Promise.resolve();
  const push = previousPush
    .catch(() => undefined)
    .then(() => getClient(options.remoteUrl).pushPdf(options));

  pendingPushes.set(options.remoteUrl, push);
  const clearPendingPush = () => {
    if (pendingPushes.get(options.remoteUrl) === push) {
      pendingPushes.delete(options.remoteUrl);
    }
  };
  void push.then(clearPendingPush, clearPendingPush);
  return push;
}

function getClient(remoteUrl: string): OverleafGitClient {
  const existing = clients.get(remoteUrl);
  if (existing) return existing;

  const filesystemName = `slide2pdf-overleaf-v1:${remoteUrl}`;
  const client = new OverleafGitClient({
    dir: "/repo",
    fs: new LightningFS(filesystemName),
    http,
  });
  clients.set(remoteUrl, client);
  return client;
}
