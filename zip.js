// @ts-check

/**
 * Reading a .zip, from SQL for Excel's sql/zip.js. Installed JOLT apps can't import libraries, and a zip is
 * simple: each file's header and data, then a directory of them. Compressed files are unpacked with the
 * browser's own inflate (DecompressionStream). No Office calls.
 */

/** Run bytes through a DecompressionStream. @param {Uint8Array} data @param {TransformStream} stream */
async function through(data, stream) {
  return new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(stream)).arrayBuffer());
}

/**
 * The files in a .zip, from its directory at the end. Folders are skipped. Throws an Error with a message
 * for the user if it isn't a zip, or uses compression this browser can't undo.
 * @param {ArrayBuffer} buffer
 * @returns {Promise<{ name: string, data: Uint8Array }[]>}
 */
export async function readZip(buffer) {
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer);
  let end = -1;
  for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 22 - 65535); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error("The dataset isn't a .zip, or it's damaged.");
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const dec = new TextDecoder();
  const out = [];
  for (let n = 0; n < count; n++) {
    if (view.getUint32(at, true) !== 0x02014b50) throw new Error("The dataset's list of files is damaged.");
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 20, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const name = dec.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    at += 46 + nameLength + extraLength + commentLength;
    if (name.endsWith("/")) continue;
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const body = bytes.subarray(start, start + size);
    if (method === 0) out.push({ name, data: body });
    else if (method === 8) {
      if (typeof DecompressionStream === "undefined") throw new Error("This version of Excel can't unpack compressed .zip files.");
      out.push({ name, data: await through(body, new DecompressionStream("deflate-raw")) });
    } else throw new Error(`${name} in the dataset is compressed in a way that can't be read here.`);
  }
  return out;
}
