import { Path, PencilBrush, type Point, type TEvent, type TSimplePathData } from 'fabric'

/** A filled stroke outline preserves real pen pressure in saved vector data. */
export class PressureBrush extends PencilBrush {
  private samples: { x: number; y: number; radius: number }[] = []
  private sample(point: Point, event: TEvent['e']) {
    const pressure = event instanceof PointerEvent && event.pointerType === 'pen' ? Math.max(0.08, event.pressure) : 0.5
    this.samples.push({ x: point.x, y: point.y, radius: Math.max(0.5, this.width * pressure) })
  }
  override onMouseDown(point: Point, options: { e: TEvent['e'] }) {
    this.samples = []
    this.sample(point, options.e)
    super.onMouseDown(point, options)
  }
  override onMouseMove(point: Point, options: { e: TEvent['e'] }) {
    this.sample(point, options.e)
    super.onMouseMove(point, options)
  }
  override createPath(pathData: TSimplePathData): Path {
    if (this.samples.length < 2) return super.createPath(pathData)
    const left: [number, number][] = [], right: [number, number][] = []
    this.samples.forEach((sample, index, all) => {
      const before = all[Math.max(0, index - 1)], after = all[Math.min(all.length - 1, index + 1)]
      const angle = Math.atan2(after.y - before.y, after.x - before.x) + Math.PI / 2
      const dx = Math.cos(angle) * sample.radius, dy = Math.sin(angle) * sample.radius
      left.push([sample.x + dx, sample.y + dy]); right.push([sample.x - dx, sample.y - dy])
    })
    const outline = [...left, ...right.reverse()]
    return new Path(outline.map(([x,y], index) => `${index ? 'L' : 'M'} ${x} ${y}`).join(' ') + ' Z', {
      fill: this.color, stroke: undefined, strokeWidth: 0, objectCaching: false,
    })
  }
}
