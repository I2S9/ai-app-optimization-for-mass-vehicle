/** Yield the main thread so clicks / menu paint before heavy JS. */
export function yieldToMain() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  });
}

const CHUNK_YIELD_EVERY = 2500;

/** Process a large array in slices, yielding so clicks/menu stay responsive. */
export async function runInChunks(items, processChunk, chunkSize = CHUNK_YIELD_EVERY) {
  const out = [];
  for (let i = 0; i < items.length; i += chunkSize) {
    if (i > 0) await yieldToMain();
    const slice = items.slice(i, i + chunkSize);
    const part = processChunk(slice, i);
    if (Array.isArray(part)) out.push(...part);
    else if (part != null) out.push(part);
  }
  return out;
}

export function deferHeavy(fn) {
  return new Promise((resolve, reject) => {
    const run = () => {
      void yieldToMain().then(() => {
        Promise.resolve()
          .then(fn)
          .then(resolve, reject);
      });
    };
    if (typeof scheduler !== 'undefined' && typeof scheduler.postTask === 'function') {
      scheduler.postTask(run, { priority: 'background' });
    } else if (typeof requestIdleCallback === 'function') {
      requestIdleCallback(run, { timeout: 32 });
    } else {
      setTimeout(run, 0);
    }
  });
}
