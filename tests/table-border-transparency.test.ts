/**
 * Fork feature: `transparency` on table cell borders.
 * Commit: 6359dd21
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

describe('table border transparency', () => {
	it('writes an alpha element for a border with transparency', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addTable([[{ text: 'cell', options: { border: { type: 'solid', color: 'FF0000', pt: 1, transparency: 50 } } }]], { x: 1, y: 1, w: 4 })

		const xml = await getSlideXml(pres)
		// all four borders (lnL/lnR/lnT/lnB) carry the 50% alpha
		for (const side of ['lnL', 'lnR', 'lnT', 'lnB']) {
			const border = xml.split(`<a:${side} `)[1]
			expect(border, `border ${side} missing`).toBeTruthy()
			expect(border.split(`</a:${side}>`)[0]).toContain('<a:alpha val="50000"/>')
		}
	})

	it('defaults border transparency to 0 (fully opaque)', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addTable([[{ text: 'cell', options: { border: { type: 'solid', color: 'FF0000', pt: 1 } } }]], { x: 1, y: 1, w: 4 })

		const xml = await getSlideXml(pres)
		const border = xml.split('<a:lnL ')[1].split('</a:lnL>')[0]
		expect(border).toContain('<a:alpha val="100000"/>')
	})
})
