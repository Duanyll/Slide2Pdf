import { Buffer as BrowserBuffer } from "buffer";

if (typeof globalThis.Buffer === "undefined") {
  globalThis.Buffer = BrowserBuffer;
}
