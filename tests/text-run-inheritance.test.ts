/**
 * Fork fix: block-level text options must only fill in run-level options that are
 * `undefined` — an explicit falsey value (e.g. `bold: false`) must be kept.
 * Commit: 99a8b659
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getSlideXml } from './helpers'

describe('text run option inheritance', () => {
	it('keeps an explicit `bold: false` on a run when the block sets `bold: true`', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText(
			[
				{ text: 'not-bold-run', options: { bold: false } },
				{ text: 'inherits-bold-run' },
			],
			{ x: 1, y: 1, bold: true }
		)

		const xml = await getSlideXml(pres)
		const runs = xml.split('<a:r>').slice(1)
		const notBoldRun = runs.find(r => r.includes('not-bold-run'))
		const inheritingRun = runs.find(r => r.includes('inherits-bold-run'))

		expect(notBoldRun).toBeTruthy()
		expect(inheritingRun).toBeTruthy()
		expect(notBoldRun).not.toContain('b="1"')
		expect(inheritingRun).toContain('b="1"')
	})

	it('still inherits block-level options a run does not define', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText([{ text: 'run-a' }, { text: 'run-b', options: { italic: true } }], { x: 1, y: 1, bold: true })

		const xml = await getSlideXml(pres)
		const runs = xml.split('<a:r>').slice(1)
		expect(runs.find(r => r.includes('run-a'))).toContain('b="1"')
		const runB = runs.find(r => r.includes('run-b'))
		expect(runB).toContain('b="1"')
		expect(runB).toContain('i="1"')
	})
})
