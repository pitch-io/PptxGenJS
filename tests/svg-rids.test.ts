/**
 * Fork fix: rIds must be incremented correctly for SVG images (each SVG consumes
 * two rIds: png preview + svg), including when combined with hyperlinks.
 * Commit: 60c57ca8
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideRelsXml, getSlideXml, TEST_PNG, TEST_SVG } from './helpers'

describe('SVG image rIds', () => {
	it('assigns unique sequential rIds across svg previews, images and hyperlinks', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addImage({ data: TEST_SVG, x: 0, y: 0, w: 1, h: 1 })
		slide.addImage({ data: TEST_PNG, x: 1, y: 1, w: 1, h: 1 })
		slide.addImage({ data: TEST_SVG, x: 2, y: 2, w: 1, h: 1, hyperlink: { url: 'https://pitch.com' } })

		const rels = await getSlideRelsXml(pres)
		const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1])).sort((a, b) => a - b)

		// svg1(png+svg) + png + svg2(png+svg) + hyperlink = 6 rels, plus the standard
		// slideLayout + notesSlide rels = 8 total — no duplicates, no gaps
		expect(ids).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
		expect(rels).toContain('Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image-1-1.png"')
		expect(rels).toContain('Target="../media/image-1-2.svg"')
		expect(rels).toContain('Target="../media/image-1-3.png"')
		expect(rels).toContain('Target="../media/image-1-4.png"')
		expect(rels).toContain('Target="../media/image-1-5.svg"')

		// the hyperlink points at its own rel, not one of the image rels
		const hyperlinkRel = rels.match(/<Relationship Id="(rId\d+)"[^>]*Target="https:\/\/pitch\.com"/)
		expect(hyperlinkRel).toBeTruthy()

		const slideXml = await getSlideXml(pres)
		expect(slideXml).toContain(`r:id="${hyperlinkRel![1]}"`)
	})

	it('references the png preview and svg via consecutive rIds in the slide XML', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addImage({ data: TEST_SVG, x: 0, y: 0, w: 1, h: 1 })

		const slideXml = await getSlideXml(pres)
		expect(slideXml).toContain('<a:blip r:embed="rId1">')
		expect(slideXml).toContain('r:embed="rId2"')

		const rels = await getSlideRelsXml(pres)
		expect(rels).toContain('Target="../media/image-1-1.png"')
		expect(rels).toContain('Target="../media/image-1-2.svg"')
	})
})
