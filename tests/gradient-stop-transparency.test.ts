/**
 * Fork feature: per-stop transparency in gradient fills.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

async function gradFillXml(fill: PptxGenJS.ShapeFillProps): Promise<string> {
	const pres = new PptxGenJS()
	const slide = pres.addSlide()
	slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 1, h: 1, fill })
	const xml = await getSlideXml(pres)
	return xml.split('<a:gradFill>')[1]?.split('</a:gradFill>')[0] ?? ''
}

describe('gradient stop transparency', () => {
	it('renders an alpha for each stop with a transparency', async () => {
		const xml = await gradFillXml({
			type: 'solid',
			gradient: { angle: 10, stops: { 0: { color: '000000', transparency: 19 }, 100: { color: '000000', transparency: 100 } } },
		})
		expect(xml).toContain('<a:gs pos="0"><a:srgbClr val="000000"><a:alpha val="81000"/></a:srgbClr></a:gs>')
		expect(xml).toContain('<a:gs pos="100000"><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr></a:gs>')
	})

	it('uses the fill transparency for stops without their own', async () => {
		const xml = await gradFillXml({
			type: 'solid',
			transparency: 50,
			gradient: { angle: 0, stops: { 0: 'FF0000', 100: { color: '00FF00', transparency: 25 } } },
		})
		expect(xml).toContain('<a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>')
		expect(xml).toContain('<a:gs pos="100000"><a:srgbClr val="00FF00"><a:alpha val="75000"/></a:srgbClr></a:gs>')
	})
})
