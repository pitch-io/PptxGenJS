/**
 * Fork fix: fractional font sizes keep their value in run properties.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

describe('fractional font size', () => {
	it('writes the run size in hundredths of a point', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addText('text', { x: 0, y: 0, w: 1, h: 1, fontSize: 7.5 })

		const xml = await getSlideXml(pres)
		expect(xml).toContain('<a:rPr lang="en-US" sz="750"')
		expect(xml).toContain('<a:endParaRPr lang="en-US" sz="750"')
	})

	it('rounds sizes finer than a hundredth of a point', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addText('text', { x: 0, y: 0, w: 1, h: 1, fontSize: 10.0049 })

		expect(await getSlideXml(pres)).toContain('<a:rPr lang="en-US" sz="1000"')
	})
})
