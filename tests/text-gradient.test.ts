/**
 * Fork feature: text gradient rendering (PITCH-1261)
 * Commits: b17b8e01, 18b5b8bb
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

describe('text gradient', () => {
	afterEach(() => {
		vi.restoreAllMocks()
	})

	it('renders a gradient fill on a text run', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('gradient text', { x: 1, y: 1, gradient: { angle: 45, stops: { 0: 'FF0000', 100: '0000FF' } } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:gradFill>')
		expect(xml).toContain('<a:gs pos="0">')
		expect(xml).toContain('<a:gs pos="100000">')
		expect(xml).toContain('val="FF0000"')
		expect(xml).toContain('val="0000FF"')
		expect(xml).toContain('<a:lin ang="2700000"/>')
	})

	it('defaults the gradient angle to 0', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('gradient text', { x: 1, y: 1, gradient: { stops: { 0: 'FF0000', 100: '0000FF' } } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:lin ang="0"/>')
	})

	it('supports theme colors as gradient stops', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('theme gradient', { x: 1, y: 1, gradient: { stops: { 0: pres.SchemeColor.accent1, 100: pres.SchemeColor.accent2 } } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:gradFill>')
		expect(xml).toContain('<a:schemeClr val="accent1"')
		expect(xml).toContain('<a:schemeClr val="accent2"')
	})

	it('gradient wins over a solid color when both are given', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('both', { x: 1, y: 1, color: 'EEEEEE', gradient: { stops: { 0: 'FF0000', 100: '0000FF' } } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:gradFill>')
		expect(xml).not.toContain('val="EEEEEE"')
	})

	it('falls back to a default background1/background2 gradient when stops are invalid', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
		const pres = new PptxGenJS()
		pres.addSlide().addText('bad gradient', { x: 1, y: 1, gradient: { stops: {} as any } })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:gradFill>')
		expect(xml).toContain('<a:schemeClr val="bg1"')
		expect(xml).toContain('<a:schemeClr val="bg2"')
		expect(warn).toHaveBeenCalled()
	})
})
