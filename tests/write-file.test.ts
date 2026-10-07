/**
 * Fork feature: browser `writeFile()` resolves with `{ name, type, size }` instead
 * of just the file name (node keeps resolving the file name).
 * Commit: 74129d26
 */
import fs from 'fs'
import os from 'os'
import path from 'path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import PptxGenJS from '../src/pptxgen'

function makePres(): PptxGenJS {
	const pres = new PptxGenJS()
	pres.addSlide().addText('writeFile test', { x: 1, y: 1 })
	return pres
}

describe('writeFile in node', () => {
	it('writes the file to disk and resolves with the file name', async () => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptxgenjs-test-'))
		const fileName = path.join(dir, 'out.pptx')

		const result = await makePres().writeFile({ fileName })

		expect(result).toBe(fileName)
		expect(fs.statSync(fileName).size).toBeGreaterThan(0)
		fs.rmSync(dir, { recursive: true, force: true })
	})
})

describe('writeFile in the browser', () => {
	afterEach(async () => {
		// let the deferred revokeObjectURL/removeChild cleanup (setTimeout 100ms) run
		// against the stubs before removing them
		await new Promise(resolve => setTimeout(resolve, 150))
		vi.unstubAllGlobals()
	})

	it('resolves with name, mime type and size of the generated blob', async () => {
		const anchor = { setAttribute: vi.fn(), click: vi.fn(), href: '', download: '', dataset: {} as Record<string, string> }
		vi.stubGlobal('document', {
			createElement: vi.fn(() => anchor),
			body: { appendChild: vi.fn(), removeChild: vi.fn() },
		})
		vi.stubGlobal('window', {
			URL: { createObjectURL: vi.fn(() => 'blob:fake-url'), revokeObjectURL: vi.fn() },
		})

		const result = await makePres().writeFile({ fileName: 'browser-test.pptx' })

		expect(result).toMatchObject({
			name: 'browser-test.pptx',
			type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		})
		expect((result as { size: number }).size).toBeGreaterThan(0)
		expect(anchor.click).toHaveBeenCalled()
	})
})
