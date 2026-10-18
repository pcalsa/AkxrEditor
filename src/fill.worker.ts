/// <reference lib="webworker" />
const workerScope = self as unknown as DedicatedWorkerGlobalScope
workerScope.onmessage = (event: MessageEvent<{ pixels: ArrayBuffer; width: number; height: number; x: number; y: number; fill: number[] }>) => {
  const { pixels, width, height, x, y, fill } = event.data
  const data = new Uint8ClampedArray(pixels)
  const start = (y * width + x) * 4
  const target = Array.from(data.slice(start,start + 4))
  const count = width * height
  const visited = new Uint8Array(count), stack = new Int32Array(count)
  let length = 1
  stack[0] = y * width + x; visited[stack[0]] = 1
  while (length) {
    const pixel = stack[--length], offset = pixel * 4
    if (target.some((value, i) => Math.abs(value - data[offset + i]) > 28)) continue
    data.set(fill, offset)
    const px = pixel % width
    const neighbors = [px > 0 ? pixel - 1 : -1, px < width - 1 ? pixel + 1 : -1, pixel >= width ? pixel - width : -1, pixel < count - width ? pixel + width : -1]
    for (const next of neighbors) if (next >= 0 && !visited[next]) { visited[next] = 1; stack[length++] = next }
  }
  workerScope.postMessage(pixels, [pixels])
}
