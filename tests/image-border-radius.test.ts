/**
 * Fork feature: image outline and rounded corners.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml, TEST_PNG } from './helpers'

async function picSpPr(options: PptxGenJS.ImageProps): Promise<string> {
	const pres = new PptxGenJS()
	const slide = pres.addSlide()
	slide.addImage({ data: TEST_PNG, x: 0, y: 0, w: 2, h: 1, ...options })
	const xml = await getSlideXml(pres)
	return xml.split('<p:pic>')[1].split('<p:spPr>')[1].split('</p:spPr>')[0]
}

describe('image outline and rounded corners', () => {
	it('renders rounded corners relative to the shorter side', async () => {
		const spPr = await picSpPr({ rectRadius: 0.25 })
		expect(spPr).toContain('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>')
	})

	it('caps rounded corners at half the shorter side', async () => {
		const spPr = await picSpPr({ rectRadius: 3 })
		expect(spPr).toContain('<a:gd name="adj" fmla="val 50000"/>')
	})

	it('measures rounded corners against the cropped size', async () => {
		const spPr = await picSpPr({ rectRadius: 0.25, sizing: { type: 'crop', w: 1, h: 0.5 } })
		expect(spPr).toContain('<a:gd name="adj" fmla="val 50000"/>')
	})

	it('keeps rectangles and full rounding as before', async () => {
		expect(await picSpPr({})).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>')
		expect(await picSpPr({ rounding: true, rectRadius: 0.25 })).toContain('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>')
	})

	it('renders an outline after the geometry', async () => {
		const spPr = await picSpPr({ line: { color: 'D92661', width: 2, dashType: 'dash' } })
		expect(spPr).toContain('</a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="D92661"/></a:solidFill><a:prstDash val="dash"/></a:ln>')
	})
})
