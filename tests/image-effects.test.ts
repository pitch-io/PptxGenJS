/**
 * Fork features on images: blur, fillOverlay, solidFill, shadow (+ inner/none shadow fix).
 * Commits: f0649255 (blur), c833913a/754d4e22/e5e8bfa7 (overlay/solidFill/validation),
 *          28e91ff3/2339e1bb (image shadow), 0eac3d65 (inner/none shadow fix)
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml, TEST_PNG } from './helpers'

const IMG = { data: TEST_PNG, x: 1, y: 1, w: 1, h: 1 }

describe('image blur', () => {
	it('writes <a:blur> inside the blip', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, blur: { radius: 40000 } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:blur rad="40000"/>')
		expect(xml).not.toContain('grow="1"')
	})

	it('writes a growing blur into the effect list when `grow` is set', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, blur: { radius: 80, grow: true } })

		const xml = await getSlideXml(pres)
		// blip-level blur keeps the raw radius; the effectLst blur is converted to points (80/10 pt -> EMU)
		expect(xml).toContain('<a:blur rad="80"/>')
		expect(xml).toContain(`<a:blur rad="${Math.round(8 * 12700)}" grow="1"/>`)
	})
})

describe('image fillOverlay', () => {
	afterEach(() => {
		vi.restoreAllMocks()
	})

	it('writes <a:fillOverlay> with blend mode, color and alpha', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, fillOverlay: { color: '00FF00', transparency: 25, blend: 'mult' } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:fillOverlay blend="mult">')
		const overlay = xml.split('<a:fillOverlay')[1].split('</a:fillOverlay>')[0]
		expect(overlay).toContain('val="00FF00"')
		expect(overlay).toContain('<a:alpha val="75000"/>')
	})

	it('defaults the blend mode to "over" and transparency to 0', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, fillOverlay: { color: 'FF0000' } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:fillOverlay blend="over">')
		expect(xml.split('<a:fillOverlay')[1]).toContain('<a:alpha val="100000"/>')
	})

	it('falls back to "over" for an invalid blend mode (with a warning)', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, fillOverlay: { color: 'FF0000', blend: 'nope' as any } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:fillOverlay blend="over">')
		expect(warn).toHaveBeenCalled()
	})

	it('strips a leading hash from the color (with a warning)', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, fillOverlay: { color: '#00FF00' } })

		const xml = await getSlideXml(pres)
		expect(xml.split('<a:fillOverlay')[1]).toContain('val="00FF00"')
		expect(warn).toHaveBeenCalled()
	})
})

describe('image solidFill', () => {
	it('writes a solid fill with alpha into the picture shape properties', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, solidFill: { color: '0000FF', transparency: 40 } })

		const xml = await getSlideXml(pres)
		const spPr = xml.split('<p:spPr>')[1].split('</p:spPr>')[0]
		expect(spPr).toContain('<a:solidFill>')
		expect(spPr).toContain('val="0000FF"')
		expect(spPr).toContain('<a:alpha val="60000"/>')
	})

	it('is ignored when no color is given', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, solidFill: {} as any })

		const xml = await getSlideXml(pres)
		const spPr = xml.split('<p:spPr>')[1].split('</p:spPr>')[0]
		expect(spPr).not.toContain('<a:solidFill>')
	})
})

describe('image shadow', () => {
	it('writes an outer shadow with scale attributes', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, shadow: { type: 'outer', blur: 8, offset: 4, angle: 45, opacity: 0.5, color: '123456' } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0"')
		expect(xml).toContain('dir="2700000"')
		expect(xml).toContain('val="123456"')
		expect(xml).toContain('<a:alpha val="50000"/>')
		expect(xml).toContain('</a:outerShdw>')
	})

	it('writes an inner shadow without outer-only attributes and with a matching closing tag', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, shadow: { type: 'inner', blur: 8, offset: 4, angle: 45, opacity: 0.5, color: '123456' } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:innerShdw')
		expect(xml).toContain('</a:innerShdw>')
		expect(xml).not.toContain('outerShdw')
		expect(xml).not.toContain('sx="100000"')
	})

	it('writes no shadow at all for type "none"', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...IMG, shadow: { type: 'none' } })

		const xml = await getSlideXml(pres)
		expect(xml).not.toContain('Shdw')
	})
})

describe('text/shape shadow (inner/none fix)', () => {
	it('closes an inner text shadow with </a:innerShdw> (not </a:outerShdw>)', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('shadowed', { x: 1, y: 1, shadow: { type: 'inner', blur: 3, offset: 2, angle: 45, opacity: 0.5, color: '000000' } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:innerShdw')
		expect(xml).toContain('</a:innerShdw>')
		expect(xml).not.toContain('outerShdw')
	})

	it('writes no shadow for text shadow type "none"', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('no shadow', { x: 1, y: 1, shadow: { type: 'none' } })

		const xml = await getSlideXml(pres)
		expect(xml).not.toContain('Shdw')
	})
})
