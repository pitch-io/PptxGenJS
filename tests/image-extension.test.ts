/**
 * Fork features: explicit `extn` option for images, and regex-based extension
 * detection that copes with query strings.
 * Commits: 7134a2aa, 098c5b3e
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideRelsXml } from './helpers'

const POS = { x: 1, y: 1, w: 1, h: 1 }

describe('image file extension', () => {
	it('uses the explicit `extn` option when the path has no extension', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...POS, path: '/nonexistent/remote-image', extn: 'jpg' })

		const rels = await getSlideRelsXml(pres)
		expect(rels).toContain('Target="../media/image-1-1.jpg"')
	})

	it('detects the extension from a path with a query string', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...POS, path: '/nonexistent/photo.jpeg?width=540' })

		const rels = await getSlideRelsXml(pres)
		expect(rels).toContain('Target="../media/image-1-1.jpeg"')
	})

	it('defaults to png when no extension can be determined', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ ...POS, path: '/nonexistent/no-extension-image' })

		const rels = await getSlideRelsXml(pres)
		expect(rels).toContain('Target="../media/image-1-1.png"')
	})

	it('uses `extn` for slide background images too', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.background = { path: '/nonexistent/background-image', extn: 'jpg' }

		const rels = await getSlideRelsXml(pres)
		// note: background images correct 'jpg' to 'jpeg' to avoid PPT content warnings
		expect(rels).toMatch(/Target="\.\.\/media\/[^"]+\.jpeg"/)
	})
})
