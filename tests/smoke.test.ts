import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml, presToZip } from './helpers'

describe('smoke', () => {
	it('generates a valid pptx zip with a slide', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addText('hello world', { x: 1, y: 1 })

		const zip = await presToZip(pres)
		expect(zip.file('ppt/presentation.xml')).toBeTruthy()

		const slideXml = await getSlideXml(pres)
		expect(slideXml).toContain('<a:t>hello world</a:t>')
	})
})
