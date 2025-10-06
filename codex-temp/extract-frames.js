const fs = require('fs');
const path = require('path');
const { fileURLToPath } = require('url');

const { createFFmpeg, fetchFile } = require('@ffmpeg/ffmpeg');

const originalFetch = global.fetch;

const readFileAsResponse = async (targetPath) => {
  const data = await fs.promises.readFile(targetPath);
  return new Response(data);
};

global.fetch = async (resource, init) => {
  if (typeof resource === 'string') {
    if (/^[a-zA-Z]:\\\\/.test(resource)) {
      return readFileAsResponse(resource);
    }
    if (resource.startsWith('file:///')) {
      return readFileAsResponse(fileURLToPath(resource));
    }
    if (!resource.startsWith('http://') && !resource.startsWith('https://')) {
      const absolutePath = path.resolve(resource);
      return readFileAsResponse(absolutePath);
    }
  } else if (resource instanceof URL && resource.protocol === 'file:') {
    return readFileAsResponse(fileURLToPath(resource));
  }
  if (originalFetch) {
    return originalFetch(resource, init);
  }
  throw new Error('Fetch not supported for resource ' + resource);
};

(async () => {
  const ffmpeg = createFFmpeg({
    log: false,
    corePath: require.resolve('@ffmpeg/core/dist/ffmpeg-core.js'),
  });
  await ffmpeg.load();
  const data = await fetchFile('../Video Project 3.mp4');
  ffmpeg.FS('writeFile', 'video.mp4', data);
  const timestamps = Array.from({ length: 13 }, (_, i) => String(i * 5));
  for (const [index, ts] of timestamps.entries()) {
    const output = `frame-${index + 1}.png`;
    try {
      await ffmpeg.run('-ss', ts, '-i', 'video.mp4', '-frames:v', '1', '-q:v', '2', output);
      const file = ffmpeg.FS('readFile', output);
      fs.writeFileSync(path.join('..', `codex-temp-${output}`), file);
    } catch (error) {
      console.error('Failed to extract frame at', ts, error.message);
    }
  }
  process.exit(0);
})();
