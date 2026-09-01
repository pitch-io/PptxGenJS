/**
 * Fork fix: a single failing media asset must not break the whole export —
 * `Promise.allSettled` is used instead of `Promise.all`, and the broken asset is
 * replaced by a 1x1 transparent placeholder (IMG_BROKEN).
 * Commits: 69613a5a, 3a054ac0
 */
import { describe, expect, it } from 'vitest'
import { IMG_BROKEN } from '../src/core-enums'
import PptxGenJS from '../src/pptxgen'
import { presToZip, TEST_PNG } from './helpers'

describe('media loading resilience', () => {
	it('still exports when an image path cannot be read, embedding the placeholder image', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		slide.addImage({ path: '/definitely/missing/image.png', x: 0, y: 0, w: 1, h: 1 })
		slide.addImage({ data: TEST_PNG, x: 1, y: 1, w: 1, h: 1 })

		// With upstream's Promise.all this rejects; the fork must resolve.
		const zip = await presToZip(pres)

		const broken = zip.file('ppt/media/image-1-1.png')
		expect(broken).toBeTruthy()
		const brokenB64 = Buffer.from(await broken!.async('nodebuffer')).toString('base64')
		expect(brokenB64).toBe(IMG_BROKEN.split(',').pop())

		// the healthy image is untouched
		const healthy = zip.file('ppt/media/image-1-2.png')
		expect(healthy).toBeTruthy()
		const healthyB64 = Buffer.from(await healthy!.async('nodebuffer')).toString('base64')
		expect(healthyB64).toBe(TEST_PNG.split(',').pop())
	})

	it('still exports when an image URL cannot be fetched (request-level https error)', async () => {
		const pres = new PptxGenJS()
		const slide = pres.addSlide()
		// .invalid is a reserved TLD - DNS resolution always fails, which emits an
		// error on the https request (not the response)
		slide.addImage({ path: 'https://broken.invalid/missing.png', x: 0, y: 0, w: 1, h: 1 })

		const zip = await presToZip(pres)

		const broken = zip.file('ppt/media/image-1-1.png')
		expect(broken).toBeTruthy()
		const brokenB64 = Buffer.from(await broken!.async('nodebuffer')).toString('base64')
		expect(brokenB64).toBe(IMG_BROKEN.split(',').pop())
	}, 15000)
})
