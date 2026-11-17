/**
 * Fork fix: the default online-video poster stays the 3.11 image (light grey play button);
 * upstream 4.0 swapped in a larger dark one.
 */
import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import PptxGenJS from '../src/pptxgen'
import { presToZip } from './helpers'

// SHA-1 of ppt/media/image-1-2.png as written by the 3.11 fork (cd424856)
const POSTER_3_11_SHA1 = '6dec658ed9da1211a6dae0b12af873349d4f4943'

describe('online video poster', () => {
	it('writes the 3.11 default poster', async () => {
		const pres = new PptxGenJS()
		pres.addSlide().addMedia({ type: 'online', link: 'https://www.youtube.com/embed/abc', x: 1, y: 1, w: 4, h: 3 })
		const poster = await (await presToZip(pres)).file('ppt/media/image-1-2.png')!.async('nodebuffer')

		expect(createHash('sha1').update(poster).digest('hex')).toBe(POSTER_3_11_SHA1)
	})
})
