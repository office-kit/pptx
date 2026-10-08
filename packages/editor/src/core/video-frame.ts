const minimumVideoReadyState = 2;
const frameDecodeTimeoutMs = 15000;
// Chromium (seen on Linux x86-64) can fire `loadeddata` on a detached video
// before its first decoded frame reaches the element's paint path; drawing it
// then paints nothing. The frame follows within a few milliseconds, so redraw
// until it does. A video whose first frame really is fully transparent (alpha
// VP8/VP9) keeps the blank draw once the wait runs out.
const firstFramePaintWaitMs = 1000;
const firstFramePaintPollMs = 16;

function drawVideoFrame(video: HTMLVideoElement): {
  canvas: HTMLCanvasElement;
  context: CanvasRenderingContext2D;
} {
  if (video.readyState < minimumVideoReadyState) {
    throw new Error('The video has no decoded frame available.');
  }
  if (video.videoWidth <= 0 || video.videoHeight <= 0) {
    throw new Error('The video has no decoded dimensions.');
  }
  const canvas = document.createElement('canvas');
  canvas.width = video.videoWidth;
  canvas.height = video.videoHeight;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('The browser could not create a 2D canvas context.');
  context.drawImage(video, 0, 0, canvas.width, canvas.height);
  return { canvas, context };
}

function paintedAnything(context: CanvasRenderingContext2D, canvas: HTMLCanvasElement): boolean {
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
  for (let alpha = 3; alpha < pixels.length; alpha += 4) if (pixels[alpha] !== 0) return true;
  return false;
}

export async function encodeVideoFrame(video: HTMLVideoElement): Promise<Uint8Array> {
  return encodeCanvas(drawVideoFrame(video).canvas);
}

async function encodeCanvas(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('The browser could not encode the video frame as PNG.'));
        return;
      }
      resolve(blob);
    }, 'image/png');
  });
  return new Uint8Array(await blob.arrayBuffer());
}

export async function decodeFirstVideoFrame(
  bytes: Uint8Array,
  contentType: string,
): Promise<Uint8Array> {
  // The reference desktop app's (Mac) Reset replaces the poster with the first decoded frame.
  // Use a separate decoder so the selected player's position is not changed.
  const video = document.createElement('video');
  const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: contentType }));
  video.preload = 'auto';
  video.muted = true;
  video.playsInline = true;
  let cleanup = () => {};
  try {
    await new Promise<void>((resolve, reject) => {
      const loaded = () => resolve();
      const failed = () =>
        reject(new Error(video.error?.message || 'The browser could not decode this video.'));
      const timer = setTimeout(
        () => reject(new Error('Timed out decoding the first video frame.')),
        frameDecodeTimeoutMs,
      );
      cleanup = () => {
        clearTimeout(timer);
        video.removeEventListener('loadeddata', loaded);
        video.removeEventListener('error', failed);
      };
      video.addEventListener('loadeddata', loaded);
      video.addEventListener('error', failed);
      video.src = url;
      video.load();
    });
    const deadline = Date.now() + firstFramePaintWaitMs;
    let drawn = drawVideoFrame(video);
    while (!paintedAnything(drawn.context, drawn.canvas) && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, firstFramePaintPollMs));
      drawn = drawVideoFrame(video);
    }
    return await encodeCanvas(drawn.canvas);
  } finally {
    cleanup();
    video.removeAttribute('src');
    video.load();
    URL.revokeObjectURL(url);
  }
}
