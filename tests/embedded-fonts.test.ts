/**
 * Fork feature: `addFonts()` — embed TrueType fonts in the presentation.
 * Commit: cffe1172
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getXml, presToZip } from './helpers'

const FONT_BYTES = 'fake-font-bytes'
const FONT_DATA = 'data:application/x-fontdata;base64,' + Buffer.from(FONT_BYTES).toString('base64')

describe('addFonts', () => {
	afterEach(() => {
		vi.restoreAllMocks()
	})

	it('embeds fonts: rels, presentation.xml, content types and the font file itself', async () => {
		const pres = new PptxGenJS()
		pres.addFonts([{ name: 'My Font', styles: [{ name: 'regular', data: FONT_DATA }, { name: 'bold', data: FONT_DATA }] }])
		pres.addSlide().addText('embedded fonts', { x: 1, y: 1, fontFace: 'My Font' })

		const zip = await presToZip(pres)

		// 1. presentation.xml.rels: one font relationship per style, target under fonts/
		const rels = await getXml(zip, 'ppt/_rels/presentation.xml.rels')
		const fontRels = [...rels.matchAll(/<Relationship Id="(rId\d+)" Type="[^"]*\/font" Target="(fonts\/[^"]+)"\/>/g)]
		expect(fontRels).toHaveLength(2)
		expect(fontRels.map(m => m[2])).toEqual(['fonts/MyFont-regular.fntdata', 'fonts/MyFont-bold.fntdata'])

		// 2. presentation.xml: embeddedFontLst referencing exactly those rel ids
		const presXml = await getXml(zip, 'ppt/presentation.xml')
		expect(presXml).toContain('<p:embeddedFontLst><p:embeddedFont><p:font typeface="My Font"/>')
		expect(presXml).toContain(`<p:regular r:id="${fontRels[0][1]}"/>`)
		expect(presXml).toContain(`<p:bold r:id="${fontRels[1][1]}"/>`)

		// 3. [Content_Types].xml declares the fntdata extension
		const contentTypes = await getXml(zip, '[Content_Types].xml')
		expect(contentTypes).toContain('<Default Extension="fntdata" ContentType="application/x-fontdata"/>')

		// 4. the decoded font bytes are stored in the package
		expect(await zip.file('ppt/fonts/MyFont-regular.fntdata')!.async('string')).toBe(FONT_BYTES)
		expect(await zip.file('ppt/fonts/MyFont-bold.fntdata')!.async('string')).toBe(FONT_BYTES)

		// 5. font rel ids must not collide with the standard presentation rels
		const allIds = [...rels.matchAll(/Id="(rId\d+)"/g)].map(m => m[1])
		expect(new Set(allIds).size).toBe(allIds.length)
	})

	it('rejects invalid font definitions with a warning and embeds nothing', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
		const pres = new PptxGenJS()
		pres.addFonts([{ name: 'Broken', styles: [{ name: 'oblique' as any, data: FONT_DATA }] }])
		pres.addSlide().addText('no fonts', { x: 1, y: 1 })

		expect(warn).toHaveBeenCalled()
		expect(pres.fonts).toEqual([])

		const zip = await presToZip(pres)
		expect(await getXml(zip, 'ppt/presentation.xml')).not.toContain('embeddedFontLst')
	})

	it('rejects styles that have neither data nor path', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
		const pres = new PptxGenJS()
		pres.addFonts([{ name: 'NoSource', styles: [{ name: 'regular' }] }])

		expect(warn).toHaveBeenCalled()
		expect(pres.fonts).toEqual([])
	})
})
