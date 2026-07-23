"use strict";

importScripts("./word-extractor-browser.js");

self.onmessage = async (event) => {
  try {
    const arrayBuffer = event.data?.arrayBuffer;
    if (!(arrayBuffer instanceof ArrayBuffer)) {
      throw new Error("Hiányzik a DOC tartalma.");
    }
    const text = String(await self.LegacyDocReader.extract(arrayBuffer)).slice(
      0,
      1024 * 1024,
    );
    self.postMessage({ ok: true, text });
  } catch (error) {
    self.postMessage({
      ok: false,
      error:
        error instanceof Error
          ? error.message
          : "A régi Word-dokumentum nem olvasható.",
    });
  }
};
