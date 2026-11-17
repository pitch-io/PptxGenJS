/**
 * Fork feature: gradient slide backgrounds.
 * Commit: 1b743fb0
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

describe('slide background gradient', () => {
	it('renders a gradient background', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.background = { gradient: { angle: 90, stops: { 0: 'FF0000', 100: '00FF00' } } }
		slide.addText('content', { x: 1, y: 1 })

		const xml = await getSlideXml(pres)
		const bg = xml.split('<p:bg>')[1]?.split('</p:bg>')[0]
		expect(bg, 'no <p:bg> element was rendered').toBeTruthy()
		expect(bg).toContain('<a:gradFill>')
		expect(bg).toContain('val="FF0000"')
		expect(bg).toContain('val="00FF00"')
		expect(bg).toContain('<a:lin ang="5400000"/>')
	})

	it('still renders plain color backgrounds', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.background = { color: 'ABCDEF' }

		const xml = await getSlideXml(pres)
		const bg = xml.split('<p:bg>')[1]?.split('</p:bg>')[0]
		expect(bg).toBeTruthy()
		expect(bg).toContain('val="ABCDEF"')
	})
})
