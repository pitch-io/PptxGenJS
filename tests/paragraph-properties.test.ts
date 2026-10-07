/**
 * Fork fix: one `a:pPr` per paragraph, before its first run.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

async function paragraphs(texts: PptxGenJS.TextProps[]): Promise<string[]> {
	const pres = new PptxGenJS()
	pres.addSlide().addText(texts, { x: 0, y: 0, w: 4, h: 1, fontSize: 7.5, lineSpacing: 9.75, paraSpaceAfter: 4.5 })
	const xml = await getSlideXml(pres)
	return xml.split('<a:p>').slice(1).map(p => p.split('</a:p>')[0])
}

describe('paragraph properties', () => {
	it('writes paragraph properties once, before the first run', async () => {
		const [p] = await paragraphs([{ text: 'Marta', options: { bold: true } }, { text: 'Lead', options: { softBreakBefore: true } }, { text: 'Email' }])
		expect(p.match(/<a:pPr/g)).toHaveLength(1)
		expect(p.startsWith('<a:pPr')).toBe(true)
		expect(p).toContain('<a:lnSpc><a:spcPts val="975"/></a:lnSpc>')
		expect(p).toContain('</a:r><a:br/><a:r>')
	})

	it('still writes paragraph properties for every paragraph', async () => {
		const ps = await paragraphs([
			{ text: 'One', options: { breakLine: true } },
			{ text: 'Two', options: { align: 'center' } },
		])
		expect(ps).toHaveLength(2)
		expect(ps[1]).toContain('<a:pPr algn="ctr"')
	})
})
