/**
 * Shared helpers: build a presentation, unzip it with JSZip, and return the XML parts
 * so tests can assert on the generated Open XML.
 */
import JSZip from 'jszip'
import PptxGenJS from '../src/pptxgen'

/** 1x1 red PNG */
export const TEST_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** 10x10 red square SVG */
export const TEST_SVG =
	'data:image/svg+xml;base64,' +
	Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="#FF0000"/></svg>').toString('base64')

export async function presToZip(pres: PptxGenJS): Promise<JSZip> {
	const buf = (await pres.write({ outputType: 'nodebuffer' })) as Buffer
	return JSZip.loadAsync(buf)
}

export async function getXml(zip: JSZip, path: string): Promise<string> {
	const file = zip.file(path)
	if (!file) throw new Error(`"${path}" not found in generated pptx. Files: ${Object.keys(zip.files).join(', ')}`)
	return file.async('string')
}

export async function getSlideXml(pres: PptxGenJS, slideNum = 1): Promise<string> {
	return getXml(await presToZip(pres), `ppt/slides/slide${slideNum}.xml`)
}

export async function getSlideRelsXml(pres: PptxGenJS, slideNum = 1): Promise<string> {
	return getXml(await presToZip(pres), `ppt/slides/_rels/slide${slideNum}.xml.rels`)
}

export async function getChartXml(pres: PptxGenJS): Promise<string> {
	const zip = await presToZip(pres)
	const chartFile = Object.keys(zip.files).find(name => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	if (!chartFile) throw new Error(`No chart XML found in generated pptx. Files: ${Object.keys(zip.files).join(', ')}`)
	return getXml(zip, chartFile)
}
