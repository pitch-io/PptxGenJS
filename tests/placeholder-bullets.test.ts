/**
 * Fork fix: layout placeholders without a `bullet` option inherit the master's
 * bullets instead of being forced to `<a:buNone/>`.
 */
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { getXml, presToZip } from './helpers'

async function layoutXml(bullet?: boolean): Promise<string> {
	const pres = new PptxGenJS()
	pres.defineSlideMaster({
		title: 'Title and Content',
		objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 8, h: 3, ...(bullet === undefined ? {} : { bullet }) }, text: '' } }],
	})
	pres.addSlide()
	return getXml(await presToZip(pres), 'ppt/slideLayouts/slideLayout2.xml')
}

describe('layout placeholder bullets', () => {
	it('inherits the master bullets when no bullet option is given', async () => {
		expect(await layoutXml()).not.toContain('<a:buNone/>')
	})

	it('writes buNone for bullet: false', async () => {
		expect(await layoutXml(false)).toContain('<a:buNone/>')
	})
})
