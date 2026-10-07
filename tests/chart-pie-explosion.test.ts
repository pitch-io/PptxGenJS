/**
 * Fork feature: `chartPieExplosion` option for pie charts.
 * Commits: 86f443c5, ea844bf8
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getChartXml } from './helpers'

const PIE_DATA = [{ name: 'Status', labels: ['Red', 'Amber', 'Green'], values: [8, 20, 30] }]

describe('pie chart explosion', () => {
	it('writes <c:explosion> when chartPieExplosion is an integer', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addChart('pie', PIE_DATA, { x: 1, y: 1, w: 6, h: 4, chartPieExplosion: 20 })

		const xml = await getChartXml(pres)
		expect(xml).toContain('<c:explosion val="20"/>')
	})

	it('omits <c:explosion> for a non-integer value', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addChart('pie', PIE_DATA, { x: 1, y: 1, w: 6, h: 4, chartPieExplosion: 12.5 })

		const xml = await getChartXml(pres)
		expect(xml).not.toContain('<c:explosion')
	})

	it('omits <c:explosion> when the option is not set', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addChart('pie', PIE_DATA, { x: 1, y: 1, w: 6, h: 4 })

		const xml = await getChartXml(pres)
		expect(xml).not.toContain('<c:explosion')
	})
})
