/**
 * Fork feature: `pptx.theme.colors`/`name` fill the theme's color scheme, and
 * `pptx.textStyles` fills the master text styles and `defaultTextStyle`.
 * Omitted values keep the hard-coded Office output.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getXml, presToZip } from './helpers'

function makePres(configure: (pres: PptxGenJS) => void = () => {}): PptxGenJS {
	const pres = new PptxGenJS()
	configure(pres)
	pres.addSlide().addText('theme test', { x: 1, y: 1 })
	return pres
}

async function parts(pres: PptxGenJS) {
	const zip = await presToZip(pres)
	return {
		theme: await getXml(zip, 'ppt/theme/theme1.xml'),
		master: await getXml(zip, 'ppt/slideMasters/slideMaster1.xml'),
		presentation: await getXml(zip, 'ppt/presentation.xml'),
	}
}

function between(xml: string, open: string, close: string): string {
	const start = xml.indexOf(open)
	return xml.slice(start, xml.indexOf(close, start) + close.length)
}

describe('theme.colors', () => {
	it('writes the given colors into the color scheme', async () => {
		const { theme } = await parts(
			makePres(pres => {
				pres.theme = { name: 'Pitch Orange', colors: { dk1: '1A1A1A', lt1: 'FFFFFF', accent1: 'FF5A00', hlink: 'FF5A00' } }
			})
		)

		expect(theme).toContain('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Pitch Orange">')
		expect(theme).toContain('<a:clrScheme name="Pitch Orange">')
		expect(theme).toContain('<a:dk1><a:srgbClr val="1A1A1A"/></a:dk1>')
		expect(theme).toContain('<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>')
		expect(theme).toContain('<a:accent1><a:srgbClr val="FF5A00"/></a:accent1>')
		expect(theme).toContain('<a:hlink><a:srgbClr val="FF5A00"/></a:hlink>')
	})

	it('keeps the Office value of every omitted color', async () => {
		const { theme } = await parts(makePres(pres => (pres.theme = { colors: { accent1: 'FF5A00' } })))

		expect(theme).toContain('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
		expect(theme).toContain('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>')
		expect(theme).toContain('<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>')
	})

	it('writes the Office theme when no theme is set', async () => {
		const { theme } = await parts(makePres())

		expect(theme).toContain('name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
		expect(theme).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>')
	})
})

describe('textStyles', () => {
	it('writes title, body and other styles into the master', async () => {
		const { master } = await parts(
			makePres(pres => {
				pres.textStyles = {
					title: { fontFace: '+mj-lt', fontSize: 45, color: 'accent1', bold: true },
					body: { fontFace: '+mn-lt', fontSize: 15, color: 'tx1', bulletColor: 'accent1' },
					other: { fontSize: 15, color: '545465' },
				}
			})
		)

		const title = between(master, '<p:titleStyle>', '</p:titleStyle>')
		expect(title).toContain('<a:defRPr sz="4500" b="1" kern="1200"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr>')

		const body = between(master, '<p:bodyStyle>', '</p:bodyStyle>')
		expect(body).toContain('<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1500" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/>')
		expect(body.match(/sz="1500"/g)).toHaveLength(9)

		const other = between(master, '<p:otherStyle>', '</p:otherStyle>')
		expect(other).toContain('<a:defRPr sz="1500" kern="1200"><a:solidFill><a:srgbClr val="545465"/></a:solidFill><a:latin typeface="+mn-lt"/>')
	})

	it('writes a literal font face with matching ea/cs faces', async () => {
		const { master } = await parts(makePres(pres => (pres.textStyles = { title: { fontFace: 'Inter' } })))

		expect(between(master, '<p:titleStyle>', '</p:titleStyle>')).toContain('<a:latin typeface="Inter"/><a:ea typeface="Inter"/><a:cs typeface="Inter"/>')
	})

	it('writes the default style into defaultTextStyle', async () => {
		const { presentation } = await parts(makePres(pres => (pres.textStyles = { default: { fontFace: '+mn-lt', fontSize: 15, color: 'tx1' } })))

		const def = between(presentation, '<p:defaultTextStyle>', '</p:defaultTextStyle>')
		expect(def.match(/<a:defRPr sz="1500" kern="1200"><a:solidFill><a:schemeClr val="tx1"\/><\/a:solidFill><a:latin typeface="\+mn-lt"\/>/g)).toHaveLength(9)
	})

	it('writes the current styles when textStyles is not set', async () => {
		const plain = await parts(makePres())
		const empty = await parts(makePres(pres => (pres.textStyles = {})))

		expect(between(plain.master, '<p:txStyles>', '</p:txStyles>')).toContain('<a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/>')
		expect(empty.master).toBe(plain.master)
		expect(empty.presentation).toBe(plain.presentation)
	})
})

describe('theme font references on runs', () => {
	it('writes +mj-lt / +mn-lt without pitchFamily and charset', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText([
			{ text: 'major', options: { fontFace: '+mj-lt' } },
			{ text: 'minor', options: { fontFace: '+mn-lt' } },
		], { x: 1, y: 1 })
		const slide = await getXml(await presToZip(pres), 'ppt/slides/slide1.xml')

		expect(slide).toContain('<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/>')
		expect(slide).toContain('<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>')
	})

	it('keeps pitchFamily and charset on literal faces', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addText('literal', { x: 1, y: 1, fontFace: 'Inter' })
		const slide = await getXml(await presToZip(pres), 'ppt/slides/slide1.xml')

		expect(slide).toContain('<a:latin typeface="Inter" pitchFamily="34" charset="0"/>')
	})
})
