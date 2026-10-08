const minimumVideoReadyState = 2;
const frameDecodeTimeoutMs = 15000;

export async function encodeVideoFrame(video: HTMLVideoElement): Promise<Uint8Array> {
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
    return await encodeVideoFrame(video);
  } finally {
    cleanup();
    video.removeAttribute('src');
    video.load();
    URL.revokeObjectURL(url);
  }
}
